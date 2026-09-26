#!/usr/bin/env python3
"""Загрузка сводок ПРС из xlsx в бэкенд без почты.

    python upload_summaries.py ~/Documents/prs                      # все xlsx из папки
    python upload_summaries.py "Сводка ПРС за сентябрь.xlsx" --dry-run --out ./parsed
    SUMMARIES_API_URL=http://localhost:8023/prs-analytics/api/repairs/v1/summaries/parsed \\
        python upload_summaries.py ~/Documents/prs --move-to ./processed

Адрес API — переменная окружения SUMMARIES_API_URL (или --api-url); по умолчанию
локальный бэкенд на сервере. Код возврата 1, если хоть один файл не загрузился.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

import requests
from dotenv import load_dotenv

from summary_parser import ParseResult, parse_workbook

DEFAULT_API_URL = "http://localhost:8023/prs-analytics/api/repairs/v1/summaries/parsed"


def collect_files(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in paths:
        p = Path(raw).expanduser()
        if p.is_dir():
            files.extend(sorted(x for x in p.iterdir() if x.suffix.lower() in (".xlsx", ".xlsm") and not x.name.startswith("~$")))
        elif p.is_file():
            files.append(p)
        else:
            print(f"[!] не найдено: {p}", file=sys.stderr)
    return files


def send_records(records: list[dict], api_url: str, *, timeout: int = 600) -> tuple[bool, dict | str]:
    """POST {"summaries": [...]} -> (успех, разобранный ответ или текст ошибки)."""
    try:
        resp = requests.post(api_url, json={"summaries": records}, timeout=timeout)
    except requests.RequestException as exc:
        return False, f"{type(exc).__name__}: {exc}"
    try:
        body = resp.json()
    except ValueError:
        body = resp.text[:500]
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {body}"
    return True, body


def format_result(body: dict | str) -> str:
    if not isinstance(body, dict):
        return str(body)
    data = body.get("data", body)
    keys = ("received", "created", "updated", "deduplicated", "skipped_unknown_well", "linked_brigades")
    parts = [f"{k}={data.get(k)}" for k in keys if k in data]
    unknown = data.get("unknown_wells") or []
    if unknown:
        parts.append(f"unknown_wells({len(unknown)})={unknown[:15]}{'…' if len(unknown) > 15 else ''}")
    return ", ".join(parts)


def process_file(path: Path, *, api_url: str, dry_run: bool, out_dir: Path | None, move_to: Path | None) -> bool:
    print(f"\n[ПАРСЕР] {path.name}")
    try:
        result: ParseResult = parse_workbook(path)
    except Exception as exc:  # noqa: BLE001
        print(f"  [!] не разобран: {type(exc).__name__}: {exc}")
        return False
    print(result.report.summary())

    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)
        target = out_dir / f"{path.stem}.json"
        target.write_text(json.dumps(result.records, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"  json: {target}")

    if not result.records:
        print("  [API] нет записей — пропуск")
        return True
    if dry_run:
        return True

    ok, body = send_records(result.records, api_url)
    print(f"  [API] {'OK' if ok else 'ОШИБКА'}: {format_result(body)}")
    if ok and move_to is not None:
        move_to.mkdir(parents=True, exist_ok=True)
        shutil.move(str(path), str(move_to / path.name))
        print(f"  перемещён в {move_to}")
    return ok


def main(argv: list[str] | None = None) -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Загрузка сводок ПРС (xlsx) в бэкенд.")
    parser.add_argument("paths", nargs="+", help="xlsx-файлы или папки с ними")
    parser.add_argument("--api-url", default=os.getenv("SUMMARIES_API_URL", DEFAULT_API_URL))
    parser.add_argument("--dry-run", action="store_true", help="только разобрать, не отправлять")
    parser.add_argument("--out", type=Path, default=None, help="папка для JSON с разобранными записями")
    parser.add_argument("--move-to", type=Path, default=None, help="куда переносить успешно загруженные файлы")
    args = parser.parse_args(argv)

    files = collect_files(args.paths)
    if not files:
        print("Нет xlsx-файлов для обработки.", file=sys.stderr)
        return 1
    print(f"API: {args.api_url}{'  (dry-run)' if args.dry_run else ''}; файлов: {len(files)}")
    failures = 0
    for path in files:
        if not process_file(path, api_url=args.api_url, dry_run=args.dry_run, out_dir=args.out, move_to=args.move_to):
            failures += 1
    print(f"\nГотово: {len(files) - failures} из {len(files)} файлов" + (" (dry-run)" if args.dry_run else ""))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
