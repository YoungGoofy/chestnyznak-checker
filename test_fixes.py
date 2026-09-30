"""
Мини-проверки логики без сети: 404-элементы True API и разбор ответа авторизации.

Запуск: python3 test_fixes.py
"""
import json
import subprocess as _sp
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from cischecker.core import checker
from cischecker.core.parser import parse_result
from cischecker.auth import auth_flow
from cischecker.updater import github as gh

CODE = "010460405660001921WXAXAKZ4zjsOZ"


def _fake_http(status, body):
    def _fake(url, payload, headers, debug=False):
        return status, body
    return _fake


def test_404_with_list_body_returns_data():
    """HTTP 404 с телом-массивом — это по-кодовые ошибки, а не провал батча."""
    checker.http_post = _fake_http(404, json.dumps([
        {"code": CODE, "errorCode": "404", "errorDescription": "КИ не найден"},
    ]))
    status, data = checker.true_check_batch([CODE], "lp", "tok")
    assert status == 404, status
    assert isinstance(data, list) and data[0]["errorCode"] == "404", data


def test_404_with_dict_body_still_fails():
    checker.http_post = _fake_http(404, json.dumps({"error_message": "метод не найден"}))
    status, data = checker.true_check_batch([CODE], "lp", "tok")
    assert status == 404 and data is None, (status, data)


def test_parse_result_error_item():
    row = parse_result(CODE, {"code": CODE, "errorCode": "404",
                              "errorDescription": "КИ не найден"}, "true")
    assert len(row) == 10, row
    assert row[0] == CODE and row[4] == "ОШИБКА: КИ не найден", row


def test_extract_uuid_token():
    token, exp = auth_flow._extract_token(
        {"uuidToken": "abc-123", "expireDate": "2026-10-10T00:00:00.123Z"})
    assert token == "abc-123", token
    expected = datetime(2026, 10, 10, tzinfo=timezone.utc).timestamp()
    assert exp is not None and abs(exp - expected) < 1, exp


def test_extract_legacy_jwt_token():
    token, exp = auth_flow._extract_token({"token": "eyJ.x.y"})
    assert token == "eyJ.x.y" and exp is None, (token, exp)


def test_perform_update_swaps_and_relaunches():
    """Обновление: старый exe → .old, новый на его место, новый процесс запущен."""
    tmp = Path(tempfile.mkdtemp())
    exe = tmp / "CISChecker.exe"
    exe.write_bytes(b"OLD")
    launched = []

    def fake_download(url, dest, progress_fn=None):
        Path(dest).write_bytes(b"NEW" * 500_000)  # > 1 МБ, проходит проверку размера
        return True

    import os
    saved = (gh.download_exe, gh.is_frozen, gh.get_exe_path, _sp.Popen)
    launch_kwargs = {}
    try:
        gh.download_exe = fake_download
        gh.is_frozen = lambda: True
        gh.get_exe_path = lambda: exe

        def fake_popen(args, **kw):
            launch_kwargs.update(kw)
            launched.append(args)
        _sp.Popen = fake_popen

        # Симулируем окружение старого PyInstaller-процесса
        os.environ["_MEIPASS2"] = "/tmp/_MEIold"
        ok, msg = gh.perform_update("http://x")
        assert ok, msg
        assert exe.read_bytes()[:3] == b"NEW", "новый exe не на месте"
        assert (tmp / "CISChecker.old").read_bytes() == b"OLD", "старый exe не сохранён в .old"
        assert launched and launched[0][0] == str(exe), launched
        assert not (tmp / "_update_tmp").exists(), "temp не удалён"
        # Новый exe не должен наследовать распаковку старого процесса
        assert launch_kwargs.get("env", {}).get("_MEIPASS2") is None, launch_kwargs.get("env")

        # cleanup_after_update подтирает .old
        gh.cleanup_after_update()
        assert not (tmp / "CISChecker.old").exists()
    finally:
        gh.download_exe, gh.is_frozen, gh.get_exe_path, _sp.Popen = saved
        os.environ.pop("_MEIPASS2", None)


def test_perform_update_rollback_when_replace_fails():
    """Сбой подмены после переименования → exe откатывается из .old."""
    tmp = Path(tempfile.mkdtemp())
    exe = tmp / "CISChecker.exe"
    exe.write_bytes(b"OLD")

    def fake_download(url, dest, progress_fn=None):
        Path(dest).write_bytes(b"NEW" * 500_000)
        return True

    real_replace = Path.replace

    def failing_replace(self, target):
        raise OSError("antivirus")

    saved = (gh.download_exe, gh.is_frozen, gh.get_exe_path)
    try:
        gh.download_exe = fake_download
        gh.is_frozen = lambda: True
        gh.get_exe_path = lambda: exe
        Path.replace = failing_replace
        ok, msg = gh.perform_update("http://x")
        assert not ok and "Не удалось заменить" in msg, (ok, msg)
        assert exe.read_bytes() == b"OLD", "exe не откатился из .old"
        assert not (tmp / "CISChecker.old").exists(), "бэкап не вернулся на место"
    finally:
        Path.replace = real_replace
        gh.download_exe, gh.is_frozen, gh.get_exe_path = saved


def test_save_token_to_env_roundtrip():
    """Токен перезаписывается, thumbprint последнего сертификата сохраняется."""
    import os
    from cischecker.core.env import save_token_to_env
    tmp = Path(tempfile.mkdtemp())

    save_token_to_env(tmp, "tok1", "1234567890", "ABC")
    text = (tmp / ".env").read_text("utf-8")
    assert "CHESTNYZNAK_TOKEN=tok1" in text
    assert "CHESTNYZNAK_INN=1234567890" in text
    assert "CHESTNYZNAK_THUMBPRINT=ABC" in text

    # Обновили только токен — thumbprint и ИНН остаются на месте
    save_token_to_env(tmp, "tok2", "1234567890", "")
    text = (tmp / ".env").read_text("utf-8")
    assert "CHESTNYZNAK_TOKEN=tok2" in text
    assert "CHESTNYZNAK_THUMBPRINT=ABC" in text
    assert "CHESTNYZNAK_INN=1234567890" in text

    assert os.environ.get("CHESTNYZNAK_TOKEN") == "tok2"
    assert os.environ.get("CHESTNYZNAK_THUMBPRINT") == "ABC"


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
            print(f"  ✓ {name}")
    print("Все проверки пройдены.")
