"""
Парсинг результатов API в строки для Excel.
"""
from __future__ import annotations

from datetime import datetime, timezone, timedelta

from .constants import STATUS_MAP, EMISSION_TYPE_MAP


def iso_to_datetime_str(iso_str: str | None) -> str:
    """ISO-строка → 'DD.MM.YYYY HH:MM:SS' (МСК)."""
    if not iso_str:
        return ""
    try:
        s = iso_str.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        dt_msk = dt.astimezone(timezone(timedelta(hours=3)))
        return dt_msk.strftime("%d.%m.%Y %H:%M:%S")
    except (ValueError, TypeError):
        return iso_str


def ts_to_datetime_str(ts: int | str | None) -> str:
    """Timestamp (ms) → строка даты (МСК)."""
    if ts is None:
        return ""
    try:
        ts = int(ts)
        if ts > 1_000_000_000_000:
            ts = ts // 1000
        dt = datetime.fromtimestamp(ts, tz=timezone.utc) + timedelta(hours=3)
        return dt.strftime("%d.%m.%Y %H:%M:%S")
    except (ValueError, OSError):
        return str(ts) if ts else ""


def translate_status(status: str) -> str:
    """Переводит статус на русский."""
    return STATUS_MAP.get(status.upper(), status)


def translate_emission_type(emission_type: str) -> str:
    """Переводит способ ввода в оборот."""
    return EMISSION_TYPE_MAP.get(emission_type.upper(), emission_type)


def extract_group_data(data: dict) -> dict | None:
    """Извлекает данные товарной группы (lpData, milkData...)."""
    for key in [
        "lpData", "milkData", "tobaccoData", "waterData", "beerData",
        "shoesData", "perfumeData", "tiresData", "cameraData",
        "bicycleData", "fursData", "wheelchairData",
    ]:
        if key in data:
            return data[key]
    return None


def parse_result(code: str, data: dict, mode: str) -> list[str]:
    """Преобразует ответ API в строку для Excel (10 колонок)."""
    if "error" in data:
        return [code, "", "", "", f"ОШИБКА: {data['error']}", "", "", "", "", ""]

    row = [""] * 10
    row[0] = code

    if mode == "true" and "_public_data" in data:
        return parse_public_row(code, data["_public_data"])

    if mode == "true":
        cis_info = data.get("cisInfo", data)
        cis_status = cis_info.get("status", cis_info.get("cisStatus", ""))
        owner_name = cis_info.get("ownerName", "")
        owner_inn = cis_info.get("ownerInn", "")
        producer = cis_info.get("producerName", "")
        introduced = cis_info.get("introducedDate", "")
        emission = cis_info.get("emissionType", "")
        gtin = cis_info.get("gtin", "")
        brand = cis_info.get("brand", "")

        row[1] = gtin
        row[2] = brand
        row[4] = translate_status(cis_status)
        row[5] = "1"

        parts = []
        if owner_name:
            parts.append(owner_name)
        if owner_inn:
            parts.append(f"ИНН {owner_inn}")
        row[6] = ", ".join(parts) if parts else ""

        row[7] = producer
        row[8] = iso_to_datetime_str(introduced)
        row[9] = translate_emission_type(emission)

        for field in ["imageIndex", "image_index", "pictureIndex"]:
            val = cis_info.get(field)
            if val is not None:
                row[3] = str(val)
                break
    else:
        return parse_public_row(code, data)

    return row


def parse_public_row(code: str, data: dict) -> list[str]:
    """Парсит строку из публичного API (10 колонок)."""
    row = [""] * 10
    row[0] = code

    pub_status = data.get("status", "")
    outer = data.get("outerStatus", "")
    status_str = outer or pub_status

    crd = data.get("codeResolveData") or {}
    if crd:
        row[1] = crd.get("gtin", "")

    catalog = (data.get("catalogData") or [None])[0]
    if catalog:
        row[2] = catalog.get("brand_name", "")

    if pub_status == "wrong":
        row[4] = "Невалидный код"
    elif data.get("codeFounded") is False:
        row[4] = "Не найден"
    else:
        row[4] = translate_status(status_str)

    row[5] = "1"

    group = extract_group_data(data)
    if group:
        if not row[2]:
            row[2] = group.get("brand", "")
        if not row[1]:
            cd = group.get("codeData") or {}
            if cd:
                row[1] = cd.get("gtin", "")
        row[7] = group.get("producerName", "")
        row[8] = ts_to_datetime_str(group.get("introducedDate", ""))
        row[9] = translate_emission_type(
            (group.get("productProperty") or {}).get("emissionType", "")
        )

    if catalog and not row[7]:
        row[7] = catalog.get("producer_name", "")

    return row
