#!/usr/bin/env python3
"""
CLI для проверки кодов маркировки «Честный Знак».
Обратная совместимость — реэкспортирует модули из новой структуры.

Usage:
  python check_codes.py --true codes.txt -o result.xlsx
  python check_codes.py --public codes.txt -o result.xlsx
"""
import argparse
import json
import os
import sys
import textwrap
import time
from pathlib import Path

# Re-exports для обратной совместимости
from cischecker.core.api import http_post
from cischecker.core.checker import (
    public_check,
    true_check_batch,
    true_check_with_retry_pg,
    true_check_auto,
    get_pg_from_public,
    resolve_pg_aliases,
    explain_http_status,
)
from cischecker.core.parser import parse_result, parse_public_row
from cischecker.core.excel import save_excel
from cischecker.core.env import load_env
from cischecker.core.constants import (
    PUBLIC_API, TRUE_API, EXCEL_HEADERS,
    PG_ALIASES, STATUS_MAP, EMISSION_TYPE_MAP, BATCH_SIZE,
)


def main():
    script_dir = Path(__file__).resolve().parent
    load_env(script_dir)

    parser = argparse.ArgumentParser(
        description="Проверка кодов маркировки «Честный Знак» с выгрузкой в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Примеры:
              python check_codes.py --true codes.txt -o result.xlsx
              python check_codes.py --true --pg lp codes.txt -o result.xlsx
              python check_codes.py --public codes.txt -o result.xlsx

            Токен: в .env → CHESTNYZNAK_TOKEN=...
              или export CHESTNYZNAK_TOKEN="..."
              или --token "..."
        """),
    )
    parser.add_argument("codes", nargs="*", metavar="CODE")
    parser.add_argument("--true", action="store_true")
    parser.add_argument("--public", action="store_true")
    parser.add_argument("-f", "--file", help="Файл со списком кодов")
    parser.add_argument("--stdin", action="store_true")
    parser.add_argument("-o", "--output", default="result.xlsx")
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--token", help="Токен True API")
    parser.add_argument("--pg", help="Код товарной группы (lp, milk, ...)")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--debug", action="store_true")

    args = parser.parse_args()
    use_true = args.true or not args.public

    # Сбор кодов
    codes = list(args.codes)
    if args.file:
        p = Path(args.file)
        if not p.exists():
            print(f"❌ Файл не найден: {args.file}", file=sys.stderr)
            sys.exit(1)
        codes.extend(l.strip() for l in p.read_text("utf-8").splitlines() if l.strip())
    if args.stdin:
        codes.extend(l.strip() for l in sys.stdin.read().splitlines() if l.strip())
    if not codes:
        parser.print_help()
        sys.exit(1)

    # Дедупликация
    seen = set()
    codes = [c for c in codes if not (c in seen or seen.add(c))]

    mode_label = "True API" if use_true else "Публичный API"
    print(f"Режим: {mode_label}")
    print(f"Кодов: {len(codes)}")
    print()

    rows: list[list[str]] = []

    if use_true:
        token = args.token or os.environ.get("CHESTNYZNAK_TOKEN", "")
        if not token:
            print("❌ Токен не задан!", file=sys.stderr)
            sys.exit(1)

        if args.pg:
            print(f"Запрос True API (pg={args.pg}, {len(codes)} кодов)...")
            status, results = true_check_batch(codes, args.pg, token, debug=args.debug)
            if results:
                result_map = {}
                for item in results:
                    cis = item.get("cis", item.get("code", item.get("requestedCis", "")))
                    if cis:
                        result_map[cis] = item
                for code in codes:
                    item = result_map.get(code)
                    if item:
                        if args.verbose:
                            print(json.dumps(item, ensure_ascii=False, indent=2))
                        rows.append(parse_result(code, item, "true"))
                    else:
                        rows.append([code, "", "", "", "Нет данных в ответе", "", "", "", "", ""])
                print(f"  Получено: {len(results)}/{len(codes)}")
            else:
                err_msg = explain_http_status(status)
                print(f"  ❌ {err_msg}", file=sys.stderr)
                for code in codes:
                    rows.append([code, "", "", "", f"ОШИБКА: {err_msg}", "", "", "", "", ""])
        else:
            results = true_check_auto(codes, token, debug=args.debug)
            for code in codes:
                item = results.get(code, {"error": "Нет данных"})
                if args.verbose and "error" not in item:
                    print(json.dumps(item, ensure_ascii=False, indent=2))
                rows.append(parse_result(code, item, "true"))
    else:
        for i, code in enumerate(codes, 1):
            print(f"[{i}/{len(codes)}] {code[:40]}...", end=" ", flush=True)
            data = public_check(code, debug=args.debug)
            if data:
                if args.verbose:
                    print()
                    print(json.dumps(data, ensure_ascii=False, indent=2))
                else:
                    status = data.get("outerStatus") or data.get("status", "?")
                    print(f"→ {status}")
                rows.append(parse_result(code, data, "public"))
            else:
                print("→ ❌ нет данных")
                rows.append([code, "", "", "", "ОШИБКА: нет ответа от API", "", "", "", "", ""])
            if i < len(codes):
                time.sleep(args.delay)

    output_path = args.output
    if not output_path.startswith("/"):
        output_path = str(script_dir / output_path)
    save_excel(rows, output_path)

    ok = sum(1 for r in rows if not r[4].startswith("ОШИБКА"))
    errors = sum(1 for r in rows if r[4].startswith("ОШИБКА"))
    print(f"\n{'=' * 50}")
    print(f"СВОДКА: проверено {len(codes)} кодов")
    print(f"  Успешно: {ok}")
    print(f"  Ошибки: {errors}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    main()
