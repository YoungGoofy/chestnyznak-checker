"""
Константы, URL-ы, маппинги для CISChecker.
"""
from __future__ import annotations

# ── API endpoints ──────────────────────────────────────────────────────
PUBLIC_API = "https://mobile.api.crpt.ru/mobile/check"
TRUE_API = "https://markirovka.crpt.ru/api/v3/true-api/cises/info"
AUTH_KEY_URL = "https://markirovka.crpt.ru/api/v3/auth/key"
AUTH_SIGN_URL = "https://markirovka.crpt.ru/api/v3/auth/simpleSignIn"

# ── HTTP настройки ─────────────────────────────────────────────────────
TIMEOUT = 20
RETRY = 3
RETRY_DELAY = 2
BATCH_SIZE = 100

# ── Маппинг category → pg для True API ────────────────────────────────
PG_ALIASES: dict[str, list[str]] = {
    "lp":         ["lp", "light_industry", "lightIndustry"],
    "milk":       ["milk", "dairy"],
    "tobacco":    ["tobacco"],
    "water":      ["water", "packed_water", "packedWater"],
    "beer":       ["beer", "brewery"],
    "shoes":      ["shoes", "footwear"],
    "perfume":    ["perfume", "perfumery"],
    "tires":      ["tires", "tyres"],
    "camera":     ["camera", "cameras"],
    "bicycle":    ["bicycle", "bicycles"],
    "furs":       ["furs"],
    "medicine":   ["medicine", "medicines", "drugs", "pharma"],
    "bio":        ["bio", "supplements", "dietary_supplements"],
    "antiseptic": ["antiseptic", "antiseptics"],
    "wheelchair": ["wheelchair", "wheelchairs"],
}

# ── Товарные группы (GUI) ──────────────────────────────────────────────
PRODUCT_GROUPS: dict[str, str] = {
    "Лёгкая промышленность (одежда, бельё, текстиль)": "lp",
    "Молочная продукция": "milk",
    "Табак": "tobacco",
    "Упакованная вода": "water",
    "Пиво и пивные напитки": "beer",
    "Обувь": "shoes",
    "Парфюмерия": "perfume",
    "Шины": "tires",
    "Фото- и видеокамеры": "camera",
    "Велосипеды": "bicycle",
    "Меховые изделия": "furs",
    "Лекарственные средства": "medicine",
    "БАД (биологически активные добавки)": "bio",
    "Дезинфицирующие средства (антисептики)": "antiseptic",
    "Кресла-коляски": "wheelchair",
}
PRODUCT_GROUPS_DEFAULT = "Лёгкая промышленность (одежда, бельё, текстиль)"

# ── Маппинг статусов ───────────────────────────────────────────────────
STATUS_MAP: dict[str, str] = {
    "EMITTED":             "Эмитирован",
    "APPLIED":             "Нанесён",
    "INTRODUCED":          "В обороте",
    "WRITTEN_OFF":         "Списан",
    "RETIRED":             "Выбыл",
    "DISAGGREGATED":       "Разагрегирован",
    "DESTROYED":           "Уничтожен",
    "SOLD":                "Продан",
    "WITHHELD":            "Приостановлен",
    "SHIPPED":             "Отгружен",
    "IN_CIRCULATION":      "В обороте",
    "IN_CIRCULATION_SOLD": "Продан",
    "WITHDRAWN":           "Выведен из оборота",
}

# ── Маппинг способа ввода в оборот ─────────────────────────────────────
EMISSION_TYPE_MAP: dict[str, str] = {
    "PRODUCTION":    "Производство в РФ",
    "IMPORT":        "Ввезён в РФ",
    "REMAINS":       "Маркировка остатков",
    "REMARK":        "Перемаркировка",
    "CROSSBORDER":   "Трансграничная торговля",
    "COMMISSIONING": "Ввод в оборот",
}

# ── Колонки Excel ──────────────────────────────────────────────────────
EXCEL_HEADERS = [
    "Штрихкод",
    "GTIN",
    "Бренд",
    "Индекс картинки вида продукции",
    "Статус",
    "Количество кодов маркировки",
    "Владелец",
    "Производитель",
    "Дата ввода в оборот",
    "Способ ввода в оборот",
]
