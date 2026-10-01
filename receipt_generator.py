#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор чеков и печатных документов из HTML-шаблонов и CSV-данных.

Скрипт работает в Windows и macOS (а также в Linux как бонус).
Всё построено на стандартной библиотеке, кроме WeasyPrint, который нужен
только для создания PDF. Если WeasyPrint недоступен (например, на Windows
без нативных библиотек GTK/Pango), используйте --pdf-engine auto:
программа создаст PDF через headless-браузер (Chrome/Edge/Chromium/Brave).
"""

from __future__ import annotations

import argparse
import calendar
import contextlib
import csv
import datetime
import html as html_module
import io
import logging
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

APP_TITLE = "ГЕНЕРАТОР ЧЕКОВ"
LINE = "=" * 32

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SEARCH_DIRS: List[Path] = [
    BASE_DIR / "templates",
    BASE_DIR / "data",
    BASE_DIR / "receipts",
]
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"
DEFAULT_FONTS_DIR = BASE_DIR / "fonts"

HTML_SUFFIXES = {".html", ".htm"}
CSV_SUFFIXES = {".csv"}
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "koi8-r", "latin-1")

PLACEHOLDER_RE = re.compile(r"{{\s*([A-Za-z0-9_.\-]+)\s*}}")
INVALID_FILENAME_CHARS = '<>:"/\\|?*'

PREFERRED_ID_FIELDS = (
    "id", "номер", "номер_чека", "number", "номердокумента", "doc", "code", "код",
)
PREFERRED_PREVIEW_FIELDS = (
    "id", "number", "номер", "name", "клиент", "client", "total", "итого", "сумма",
)

FONT_STACK = '"Roboto", "Liberation Sans", Arial, "Helvetica Neue", sans-serif'
FONT_FILE_HINTS = (
    "roboto-regular", "roboto-bold", "roboto-italic", "roboto-bolditalic",
    "roboto-black", "roboto-light", "roboto-medium",
    "liberationsans-regular", "liberationsans-bold",
    "liberationsans-italic", "liberationsans-bolditalic",
    "arial-regular", "arialbd",
)

DATE_FORMATS = (
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
    "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
)

DEADLINE_HEADER_RE = re.compile(
    r"(due|deadline|expir|срок|оплат|платеж|крайн|рассрочк)", re.IGNORECASE
)

DATE_TOKEN_RE = re.compile(
    r"(?<![\d.])(\d{4}-\d{2}-\d{2}|\d{1,2}[./]\d{1,2}[./]\d{2,4})(?![\d])"
)

TEXT_LETTER_RE = re.compile(r"[A-Za-z\u0400-\u04FF]")

MONTH_END_RE = re.compile(
    r"до\s+конца\s+месяца|к\s+концу\s+месяца|в\s+конце\s+месяца|концу?\s+месяца",
    re.IGNORECASE,
)
MONTH_MID_RE = re.compile(
    r"в\s+середине\s+месяца|середин[ау]\s+месяца", re.IGNORECASE
)
MONTH_START_RE = re.compile(
    r"начал[аеоу]\s+месяца|до\s+начала\s+месяца|в\s+начале\s+месяца", re.IGNORECASE
)
PLUS_DAYS_RE = re.compile(
    r"в\s+течение\s+(\d{1,3})\s*(?:календарных\s+)?(?:дн|день|дней|рабочих\s+дн)",
    re.IGNORECASE,
)
NOT_LATER_RE = re.compile(
    r"не\s+позднее\s+(\d{1,2})\s*(?:го\s+)?числа", re.IGNORECASE
)


class UserError(Exception):
    """Ожидаемая пользовательская ошибка: показывается без traceback."""


class UserCancelled(Exception):
    """Пользователь отменил текущую операцию."""


def decode_text(raw: bytes) -> Tuple[str, str]:
    """Определяет кодировку файла и возвращает пару (текст, кодировка)."""
    for encoding in CSV_ENCODINGS:
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8 (с заменой нечитаемых символов)"


def find_files(directories: Sequence[Path], suffixes: set) -> List[Path]:
    """Рекурсивно ищет файлы с нужными расширениями во всех директориях."""
    found: Dict[str, Path] = {}
    for directory in directories:
        directory = Path(directory).expanduser()
        if not directory.is_dir():
            continue
        for path in directory.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.lower() not in suffixes:
                continue
            if path.name.startswith("~$") or path.name.startswith("."):
                continue
            try:
                resolved = path.resolve()
            except OSError:
                resolved = path
            found.setdefault(str(resolved), resolved)
    return sorted(found.values(), key=lambda item: str(item).lower())


def scan_all(directories: Sequence[Path]) -> Tuple[List[Path], List[Path]]:
    """Возвращает пары (шаблоны, csv-файлы) по всем заданным директориям."""
    return (
        find_files(directories, HTML_SUFFIXES),
        find_files(directories, CSV_SUFFIXES),
    )


def display_path(path: Path) -> str:
    """Показывает путь относительно каталога проекта, если это возможно."""
    parent = path.parent
    name = path.name
    if parent == BASE_DIR:
        return name
    try:
        return str(parent.relative_to(BASE_DIR) / name)
    except ValueError:
        return str(path)


def has_duplicate_names(files: Sequence[Path]) -> bool:
    names = [item.name.lower() for item in files]
    return len(set(names)) != len(names)


def prompt(message: str, allowed: Optional[set] = None) -> str:
    """Запрашивает строку у пользователя с проверкой допустимых ответов."""
    while True:
        try:
            value = input(message).strip()
        except (EOFError, KeyboardInterrupt):
            print()
            raise UserCancelled() from None
        if not value:
            print("Пустой ввод, попробуйте ещё раз.")
            continue
        if allowed is None or value in allowed:
            return value
        hint = ", ".join(sorted(allowed))
        print(f"Некорректный ввод: «{value}». Допустимо: {hint}.")


def prompt_yes_no(message: str, default: bool = True) -> bool:
    suffix = "[Y/n]" if default else "[y/N]"
    while True:
        value = prompt(f"{message} {suffix}: ").lower()
        if value in {"y", "yes", "д", "да"}:
            return True
        if value in {"n", "no", "н", "нет"}:
            return False


def select_file(title: str, files: Sequence[Path], show_dirs: Optional[bool] = None) -> Path:
    """Выводит список файлов и позволяет выбрать один по номеру."""
    if not files:
        raise UserError(f"{title}: файлы не найдены.")

    if show_dirs is None:
        show_dirs = has_duplicate_names(files) or True

    print()
    print(title)
    print()
    for index, path in enumerate(files, start=1):
        if show_dirs:
            print(f"{index:>3}. {display_path(path)}")
        else:
            print(f"{index:>3}. {path.name}")
    print()

    allowed = {str(index) for index in range(1, len(files) + 1)} | {"0"}
    while True:
        answer = prompt("Выберите файл (0 — назад): ", allowed)
        if answer == "0":
            raise UserCancelled()
        return files[int(answer) - 1]


def detect_delimiter(text: str) -> str:
    """Определяет разделитель CSV: запятая, точка с запятой, таб или вертикальная черта."""
    sample = "\n".join(text.splitlines()[:5])
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        pass

    first_line = sample.splitlines()[0] if sample.splitlines() else ""
    counts = {candidate: first_line.count(candidate) for candidate in (",", ";", "\t", "|")}
    best = max(counts, key=lambda key: counts[key])
    return best if counts[best] else ","


def read_csv(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    """Читает CSV с заголовками и возвращает (заголовки, записи)."""
    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise UserError(f"Не удалось прочитать файл «{display_path(path)}»: {exc}") from None

    if not raw.strip():
        raise UserError(f"CSV-файл «{display_path(path)}» пуст.")

    text, encoding = decode_text(raw)
    if encoding not in ("utf-8", "utf-8-sig"):
        print(f"[i] Файл прочитан в кодировке {encoding}.")

    reader = csv.reader(text.splitlines(), delimiter=detect_delimiter(text))
    try:
        rows = [row for row in reader if any(cell.strip() for cell in row)]
    except csv.Error as exc:
        raise UserError(f"Ошибка разбора CSV «{display_path(path)}»: {exc}") from None

    if not rows:
        raise UserError(f"В CSV «{display_path(path)}» нет строки заголовков.")

    headers = [cell.strip().lstrip("\ufeff") for cell in rows[0]]
    if not any(headers):
        raise UserError(f"В CSV «{display_path(path)}» пустая строка заголовков.")

    seen: Dict[str, int] = {}
    unique_headers: List[str] = []
    for position, header in enumerate(headers):
        if not header:
            header = f"column_{position + 1}"
        base = header
        counter = seen.get(base.lower(), 0)
        seen[base.lower()] = counter + 1
        if counter:
            header = f"{base}_{counter + 1}"
        unique_headers.append(header)

    records: List[Dict[str, str]] = []
    for row in rows[1:]:
        record: Dict[str, str] = {}
        for index, header in enumerate(unique_headers):
            record[header] = row[index].strip() if index < len(row) else ""
        for index in range(len(unique_headers), len(row)):
            record[f"column_{index + 1}"] = row[index].strip()
        records.append(record)

    if not records:
        raise UserError(
            f"В CSV «{display_path(path)}» нет данных: найдена только строка заголовков."
        )

    return unique_headers, records


def format_record_preview(headers: Sequence[str], record: Dict[str, str], limit: int = 4) -> str:
    """Формирует короткое описание записи без привязки к конкретным колонкам."""
    filled = [
        (header, (record.get(header) or "").strip())
        for header in headers
        if (record.get(header) or "").strip()
    ]
    if not filled:
        return "(запись без данных)"

    priority = [
        item for item in filled
        if item[0].strip().lower() in PREFERRED_PREVIEW_FIELDS
    ]
    rest = [item for item in filled if item not in priority]
    ordered = (priority + rest)[:limit]
    if not ordered:
        ordered = filled[:limit]

    head_header, head_value = ordered[0]
    parts = [f"{head_header.upper()}: {head_value}"]
    parts.extend(value for _, value in ordered[1:])
    suffix = " | ..." if len(filled) > limit else ""
    return " | ".join(parts) + suffix


def parse_date_value(value: str) -> Optional[datetime.date]:
    """Пытается разобрать строку как дату. Возвращает None, если это не дата."""
    text = (value or "").strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def format_date_ru(value: datetime.date) -> str:
    """Формат даты по-русски для сообщений пользователю."""
    return f"{value.day} {MONTH_NAMES_RU[value.month - 1]} {value.year}"


MONTH_NAMES_RU = (
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
)


def collect_deadline_claims(
    headers: Sequence[str],
    record: Dict[str, str],
) -> Tuple[List[Dict[str, object]], Optional[datetime.date]]:
    """Собирает все утверждения записи о сроках и датах.

    Утверждения бывают трёх видов: значение колонки целиком является датой,
    дата упомянута внутри свободного текста, либо срок задан фразой
    («до конца месяца», «в течение 14 дней»).
    """
    whole_dates: Dict[str, datetime.date] = {}
    for header in headers:
        parsed = parse_date_value(record.get(header, ""))
        if parsed is not None:
            whole_dates[header] = parsed

    if not whole_dates:
        return [], None

    reference = min(whole_dates.values())
    days_in_month = calendar.monthrange(reference.year, reference.month)[1]
    month_first = datetime.date(reference.year, reference.month, 1)
    month_last = datetime.date(reference.year, reference.month, days_in_month)

    claims: List[Dict[str, object]] = []
    for header in headers:
        value = (record.get(header) or "").strip()
        if not value:
            continue

        if header in whole_dates:
            claims.append({
                "source": header,
                "kind": "value",
                "label": f"значение поля «{header}»",
                "date": whole_dates[header],
                "text": value,
            })
            continue

        if not TEXT_LETTER_RE.search(value):
            continue

        for match in DATE_TOKEN_RE.finditer(value):
            parsed = parse_date_value(match.group(1))
            if parsed is not None:
                claims.append({
                    "source": header,
                    "kind": "token",
                    "label": f"дата «{match.group(1)}» в тексте",
                    "date": parsed,
                    "text": value,
                })

        phrase_date: Optional[datetime.date] = None
        phrase_label = ""
        if MONTH_END_RE.search(value):
            phrase_date, phrase_label = month_last, "«до конца месяца»"
        elif MONTH_MID_RE.search(value):
            phrase_date = datetime.date(reference.year, reference.month, min(15, days_in_month))
            phrase_label = "«в середине месяца»"
        elif MONTH_START_RE.search(value):
            phrase_date, phrase_label = month_first, "«в начале месяца»"
        else:
            match = PLUS_DAYS_RE.search(value)
            if match:
                phrase_date = reference + datetime.timedelta(days=int(match.group(1)))
                phrase_label = f"«в течение {match.group(1)} дн.»"
            else:
                match = NOT_LATER_RE.search(value)
                if match:
                    day = min(int(match.group(1)), days_in_month)
                    phrase_date = datetime.date(reference.year, reference.month, day)
                    phrase_label = f"«не позднее {match.group(1)} числа»"

        if phrase_date is not None:
            claims.append({
                "source": header,
                "kind": "phrase",
                "label": f"{phrase_label} от {format_date_ru(reference)}",
                "date": phrase_date,
                "text": value,
            })

    return claims, reference


def validate_record(
    headers: Sequence[str],
    record: Dict[str, str],
) -> List[Dict[str, object]]:
    """Ищет расхождения между сроком в колонке-дате и текстом записи.

    Проверяются только два случая, где срок действительно продублирован:

    1. в свободном тексте есть дата или фраза о сроке, не совпадающая
       с колонкой, название которой выглядит как срок оплаты;
    2. две разные колонки-срока содержат разные даты.

    Обычные колонки-даты (дата выставления, дата поставки) не сравниваются
    между собой: разные даты у разных фактов — это норма.
    """
    claims, _reference = collect_deadline_claims(headers, record)
    if not claims:
        return []

    deadlines: Dict[str, datetime.date] = {}
    for header in headers:
        parsed = parse_date_value(record.get(header, ""))
        if parsed is not None and DEADLINE_HEADER_RE.search(header):
            deadlines[header] = parsed

    if not deadlines:
        return []

    issues: List[Dict[str, object]] = []
    seen: set = set()

    def add_issue(
        source: str,
        label: str,
        text: str,
        kind: str,
        actual: datetime.date,
        deadline: str,
        expected: datetime.date,
    ) -> None:
        key = (source, label, deadline)
        if key in seen:
            return
        seen.add(key)
        issues.append({
            "source": source,
            "label": label,
            "text": text,
            "kind": kind,
            "actual": actual,
            "deadline": deadline,
            "expected": expected,
            "days": abs((actual - expected).days),
        })

    for claim in claims:
        if claim["kind"] == "value":
            continue
        source = str(claim["source"])
        if source in deadlines:
            continue
        for header, expected in deadlines.items():
            actual = claim["date"]
            if actual == expected:
                continue
            add_issue(
                source,
                str(claim["label"]),
                str(claim["text"]),
                str(claim["kind"]),
                actual,
                header,
                expected,
            )

    names = list(deadlines)
    for index, first in enumerate(names):
        for second in names[index + 1:]:
            if deadlines[first] == deadlines[second]:
                continue
            earlier, later = sorted((first, second), key=lambda key: deadlines[key])
            add_issue(
                earlier,
                f"значение поля «{earlier}»",
                deadlines[earlier].isoformat(),
                "value",
                deadlines[earlier],
                later,
                deadlines[later],
            )

    return issues


def format_consistency_issues(
    issues: Sequence[Dict[str, object]],
    record_label: str,
) -> str:
    """Готовит читаемый блок предупреждения о расхождениях."""
    lines = [f"[!] Противоречие в данных записи {record_label}:"]
    for issue in issues:
        title = "явное расхождение" if issue["kind"] == "phrase" else "возможное расхождение"
        actual = issue["actual"]
        expected = issue["expected"]
        lines.append(
            f"    {title}: поле «{issue['source']}» — {issue['label']}"
            f" = {format_date_ru(actual)} ({actual.isoformat()})"
        )
        lines.append(
            f"        поле «{issue['deadline']}» = {format_date_ru(expected)}"
            f" ({expected.isoformat()}), расхождение {issue['days']} дн."
        )
        lines.append(f"        текст: {issue['text']}")
    lines.append(
        "    Один и тот же срок задан дважды по-разному. "
        "Уберите срок из свободного текста или исправьте колонку."
    )
    return "\n".join(lines)


def confirm_single_generation(
    headers: Sequence[str],
    record: Dict[str, str],
    label: str,
) -> bool:
    """Спрашивает подтверждение, если в записи есть противоречие о сроке."""
    issues = validate_record(headers, record)
    if not issues:
        return True
    print()
    print(format_consistency_issues(issues, label))
    return prompt_yes_no("Создать документ с такими данными?", default=False)


def confirm_batch_generation(
    headers: Sequence[str],
    records: Sequence[Dict[str, str]],
) -> bool:
    """Проверяет все записи разом и спрашивает подтверждение один раз."""
    problems: List[Tuple[str, List[Dict[str, object]]]] = []
    for position, record in enumerate(records, start=1):
        issues = validate_record(headers, record)
        if issues:
            problems.append((format_record_identity(headers, record, position), issues))

    if not problems:
        return True

    print()
    print(f"[!] Противоречия найдены в {len(problems)} из {len(records)} записей:")
    for label, issues in problems[:20]:
        print(format_consistency_issues(issues, label))
        print()
    if len(problems) > 20:
        print(f"    ... и ещё {len(problems) - 20} записей с расхождениями.")
        print()

    return prompt_yes_no(
        f"Создать документы для всех записей, несмотря на расхождения ({len(records)} шт.)?",
        default=False,
    )


def format_record_identity(
    headers: Sequence[str],
    record: Dict[str, str],
    position: int,
) -> str:
    """Короткая подпись записи для сообщений о проверке."""
    identifier = build_record_identifier(headers, record, position)
    for header in headers:
        if header.strip().lower() in PREFERRED_ID_FIELDS:
            return f"{header}={identifier}"
    return f"строка {position} ({identifier})"


def select_record(headers: Sequence[str], records: Sequence[Dict[str, str]]) -> Dict[str, str]:
    """Показывает список записей CSV и возвращает выбранную."""
    if not records:
        raise UserError("В выбранном CSV нет ни одной записи.")

    print()
    print(f"Записей в файле: {len(records)}")
    print()
    for index, record in enumerate(records, start=1):
        print(f"{index:>3}. {format_record_preview(headers, record)}")
    print()

    allowed = {str(index) for index in range(1, len(records) + 1)} | {"0"}
    while True:
        answer = prompt("Введите номер записи (0 — назад): ", allowed)
        if answer == "0":
            raise UserCancelled()
        return records[int(answer) - 1]


def load_template(path: Path) -> str:
    """Загружает HTML-шаблон, определяя кодировку автоматически."""
    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise UserError(f"Не удалось прочитать шаблон «{display_path(path)}»: {exc}") from None

    if not raw.strip():
        raise UserError(f"Шаблон «{display_path(path)}» пуст.")

    text, _ = decode_text(raw)
    return text


def collect_font_faces(fonts_dir: Path) -> str:
    """Ищет файлы шрифтов и строит блок @font-face (кроссплатформенно)."""
    faces: List[str] = []
    seen: set = set()

    def register(font_path: Path) -> None:
        try:
            key = font_path.stat().st_size
        except OSError:
            key = 0
        if (str(font_path), key) in seen:
            return
        seen.add((str(font_path), key))
        try:
            uri = font_path.resolve().as_uri()
        except (OSError, ValueError):
            return
        weight = "700" if "bold" in font_path.stem.lower() else "400"
        style = "italic" if "italic" in font_path.stem.lower() else "normal"
        faces.append(
            "@font-face {\n"
            f'    font-family: "ReceiptFont";\n'
            f"    src: url({uri}) format(\"truetype\");\n"
            f"    font-weight: {weight};\n"
            f"    font-style: {style};\n"
            "}"
        )

    search_dirs: List[Path] = [Path(fonts_dir).expanduser()]
    if sys.platform == "darwin":
        search_dirs += [
            Path("/Library/Fonts"),
            Path("/System/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ]
    elif os.name == "nt":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        search_dirs += [
            Path(windir) / "Fonts",
            Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts",
        ]
    else:
        search_dirs += [
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path.home() / ".fonts",
            Path.home() / ".local" / "share" / "fonts",
        ]

    for directory in search_dirs:
        if not directory.is_dir():
            continue
        candidates = directory.rglob("*.ttf") if directory == search_dirs[0] else directory.glob("*.ttf")
        for font_path in sorted(candidates):
            if not font_path.is_file():
                continue
            if not any(font_path.stem.lower().startswith(hint) for hint in FONT_FILE_HINTS):
                continue
            register(font_path)

    if not faces:
        return ""
    return "\n".join(faces) + "\n"


def build_base_css(fonts_dir: Path) -> str:
    """Базовая таблица стилей для аккуратного PDF и кириллицы."""
    faces = collect_font_faces(fonts_dir)
    family = f'"ReceiptFont", {FONT_STACK}' if faces else FONT_STACK
    return f"""
{faces}
    @page {{
        size: A4;
        margin: 14mm 12mm;
    }}

    html, body {{
        font-family: {family};
        font-size: 11pt;
        line-height: 1.4;
        color: #111;
        background: #fff;
    }}

    body {{
        margin: 0;
        padding: 0;
        word-wrap: break-word;
        overflow-wrap: anywhere;
    }}

    h1, h2, h3, h4, p, ul, ol, blockquote {{
        page-break-inside: avoid;
        margin: 0 0 6pt 0;
    }}

    img, svg {{
        max-width: 100%;
        height: auto;
    }}

    table {{
        width: 100%;
        max-width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        margin: 8pt 0;
        page-break-inside: auto;
    }}

    thead {{
        display: table-header-group;
    }}

    tr {{
        page-break-inside: avoid;
        page-break-after: auto;
    }}

    th, td {{
        border: 1px solid #444;
        padding: 6px;
        text-align: center;
        vertical-align: middle;
        overflow-wrap: anywhere;
        word-wrap: break-word;
        white-space: normal;
    }}

    th {{
        background: #ececec;
        font-weight: 700;
    }}

    .kv {{
        width: 100%;
        border-collapse: collapse;
    }}

    .kv th {{
        width: 35%;
        text-align: left;
        background: #f6f6f6;
    }}

    .total {{
        font-weight: 700;
        font-size: 13pt;
    }}

    .right {{
        text-align: right;
    }}

    .left {{
        text-align: left;
    }}

    .nowrap {{
        white-space: nowrap;
    }}
"""


def inject_css(document: str, css: str) -> str:
    """Вставляет базовый CSS в документ, не ломая разметку шаблона."""
    if not css.strip():
        return document
    style_block = f"<style>\n{css}\n</style>"

    match = re.search(r"</head\s*>", document, flags=re.IGNORECASE)
    if match:
        return document[: match.start()] + style_block + "\n" + document[match.start():]

    match = re.search(r"<body\b[^>]*>", document, flags=re.IGNORECASE)
    if match:
        return document[: match.start()] + style_block + "\n" + document[match.start():]

    return style_block + "\n" + document


def render_template(
    template: str,
    record: Dict[str, str],
    base_css: str,
) -> Tuple[str, List[str], List[str]]:
    """Заменяет плейсхолдеры {{column}} значениями записи.

    Возвращает (html, неизвестные_плейсхолдеры, использованные_колонки).
    """
    used: List[str] = []
    missing: List[str] = []

    def substitute(match: "re.Match[str]") -> str:
        column = match.group(1)
        if column in record:
            if column not in used:
                used.append(column)
            return html_module.escape(record[column], quote=True)
        if column not in missing:
            missing.append(column)
        return match.group(0)

    rendered = PLACEHOLDER_RE.sub(substitute, template)
    rendered = inject_css(rendered, base_css)
    return rendered, missing, used


def report_template_issues(missing: Sequence[str], headers: Sequence[str]) -> None:
    """Сообщает о плейсхолдерах, для которых нет данных в CSV."""
    if missing:
        listed = ", ".join(f"{{{{{name}}}}}" for name in missing)
        available = ", ".join(headers) or "нет"
        print(f"[!] В шаблоне есть плейсхолдеры без данных: {listed}")
        print(f"    Доступные колонки CSV: {available}")
        print("    Такие места оставлены в исходном виде с плейсхолдерами.")


def report_unused_columns(used: Sequence[str], headers: Sequence[str]) -> None:
    """Сообщает о колонках CSV, которые не попали в шаблон."""
    unused = [header for header in headers if header not in used]
    if unused:
        print(f"[i] Колонки CSV, не использованные в шаблоне: {', '.join(unused)}")


def sanitize_filename(name: str, fallback: str = "document", max_length: int = 120) -> str:
    """Приводит строку к безопасному имени файла для Windows/macOS."""
    cleaned = unicodedata.normalize("NFC", str(name))
    cleaned = "".join(
        "_" if character in INVALID_FILENAME_CHARS else character
        for character in cleaned
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    cleaned = re.sub(r"_{3,}", "__", cleaned)
    if not cleaned:
        cleaned = fallback
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length].strip(" .")
    return cleaned or fallback


def build_record_identifier(headers: Sequence[str], record: Dict[str, str], position: int) -> str:
    """Определяет идентификатор записи для имени выходного файла."""
    lowered = {header.strip().lower(): header for header in headers}
    for candidate in PREFERRED_ID_FIELDS:
        header = lowered.get(candidate)
        if header:
            value = (record.get(header) or "").strip()
            if value:
                return value
    for header in headers:
        value = (record.get(header) or "").strip()
        if value:
            return value
    return f"row_{position:03d}"


def save_html(rendered_html: str, output_dir: Path, file_name: str) -> Path:
    """Сохраняет итоговый HTML в папку output (создаёт её при необходимости)."""
    try:
        output_dir = Path(output_dir).expanduser()
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise UserError(f"Не удалось подготовить папку «{output_dir}»: {exc}") from None

    target = output_dir / file_name
    try:
        target.write_text(rendered_html, encoding="utf-8")
    except OSError as exc:
        raise UserError(f"Не удалось сохранить HTML «{target}»: {exc}") from None
    return target


@contextlib.contextmanager
def quiet_weasyprint_import():
    """Заглушает баннер WeasyPrint о нативных библиотеках во время импорта."""
    logger = logging.getLogger("weasyprint")
    previous_level = logger.level
    previous_propagate = logger.propagate
    logger.setLevel(logging.CRITICAL + 10)
    logger.propagate = False
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            yield
    finally:
        logger.setLevel(previous_level)
        logger.propagate = previous_propagate


def weasyprint_available() -> bool:
    """Проверяет, что WeasyPrint установлен и его нативные библиотеки доступны."""
    try:
        with quiet_weasyprint_import():
            import weasyprint  # noqa: F401
    except Exception:
        return False
    return True


def find_headless_browser() -> Optional[Path]:
    """Ищет Chrome/Edge/Chromium для резервной печати в PDF."""
    candidates: List[Path] = []

    from_env = os.environ.get("CHROME_PATH") or os.environ.get("BROWSER_PATH")
    if from_env:
        candidates.append(Path(from_env).expanduser())

    if sys.platform == "win32":
        relative_paths = (
            r"Google\Chrome\Application\chrome.exe",
            r"Microsoft\Edge\Application\msedge.exe",
            r"Chromium\Application\chrome.exe",
            r"BraveSoftware\Brave-Browser\Application\brave.exe",
        )
        roots = [
            os.environ.get("PROGRAMFILES"),
            os.environ.get("PROGRAMFILES(X86)"),
            os.environ.get("LOCALAPPDATA"),
        ]
        for root in roots:
            if not root:
                continue
            for relative in relative_paths:
                candidates.append(Path(root) / relative)
    elif sys.platform == "darwin":
        names = (
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
        )
        candidates.extend(Path(name) for name in names)
        candidates.append(
            Path.home() / "Applications" / "Google Chrome.app" / "Contents" / "MacOS" / "Google Chrome"
        )
    else:
        for binary in ("google-chrome", "chromium", "chromium-browser", "microsoft-edge", "brave-browser"):
            found = shutil.which(binary)
            if found:
                candidates.append(Path(found))

    for candidate in candidates:
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


def engine_weasyprint(html_path: Path, pdf_path: Path) -> Path:
    """Создаёт PDF через WeasyPrint (основной движок)."""
    try:
        with quiet_weasyprint_import():
            from weasyprint import HTML
    except Exception as exc:
        raise UserError(
            "Не удалось импортировать WeasyPrint.\n"
            "    Установка: pip install -r requirements.txt\n"
            "    На Windows WeasyPrint требует нативных библиотек GTK/Pango "
            "(см. https://weasyprint.org/).\n"
            "    Либо запустите программу с ключом --pdf-engine auto."
        ) from exc

    try:
        from weasyprint.text.fonts import FontConfiguration
    except Exception:
        FontConfiguration = None  # type: ignore[assignment]

    options = {"font_config": FontConfiguration()} if FontConfiguration is not None else {}
    try:
        HTML(filename=str(html_path), base_url=str(html_path.parent)).write_pdf(
            str(pdf_path), **options
        )
    except Exception as exc:
        raise UserError(f"WeasyPrint не смог создать PDF «{pdf_path.name}»:\n    {exc}") from exc

    return pdf_path


def engine_browser(html_path: Path, pdf_path: Path) -> Path:
    """Создаёт PDF через headless Chrome/Edge (резервный движок)."""
    browser = find_headless_browser()
    if browser is None:
        raise UserError(
            "Headless-браузер не найден (Chrome, Edge, Chromium или Brave).\n"
            "    Укажите путь вручную через переменную окружения CHROME_PATH."
        )

    try:
        url = html_path.resolve().as_uri()
    except ValueError as exc:
        raise UserError(f"Некорректный путь к HTML: {exc}") from None

    command = [
        str(browser),
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-pdf-header-footer",
        "--virtual-time-budget=8000",
        f"--print-to-pdf={pdf_path.resolve()}",
        url,
    ]

    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except subprocess.TimeoutExpired:
        raise UserError(f"{browser.name} не ответил за 120 секунд.") from None
    except OSError as exc:
        raise UserError(f"Не удалось запустить {browser.name}: {exc}") from None

    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        detail = (completed.stderr or completed.stdout or "").strip().splitlines()
        hint = detail[-1] if detail else "код возврата " + str(completed.returncode)
        raise UserError(f"{browser.name} не создал PDF: {hint}")

    return pdf_path


PDF_ENGINES = {
    "weasyprint": [("WeasyPrint", engine_weasyprint)],
    "browser": [("браузер", engine_browser)],
    "auto": [
        ("WeasyPrint", engine_weasyprint),
        ("браузер", engine_browser),
    ],
}


def create_pdf(html_path: Path, pdf_path: Path, engine: str = "weasyprint") -> Path:
    """Конвертирует HTML в PDF выбранным движком."""
    engines = PDF_ENGINES.get(engine)
    if engines is None:
        raise UserError(f"Неизвестный движок PDF: {engine}")

    try:
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise UserError(f"Не удалось подготовить папку «{pdf_path.parent}»: {exc}") from None

    if not html_path.is_file():
        raise UserError(f"Исходный HTML не найден: {html_path}")

    failures: List[str] = []
    for label, func in engines:
        try:
            path = func(html_path, pdf_path)
        except UserError as exc:
            failures.append(f"{label}: {exc}")
            continue
        if not path.exists() or path.stat().st_size == 0:
            failures.append(f"{label}: PDF создан пустым или не создан.")
            continue
        return path

    details = "\n".join(f"    - {item}" for item in failures)
    raise UserError(f"Не удалось создать PDF «{pdf_path.name}»:\n{details}")


def open_pdf(pdf_path: Path) -> bool:
    """Открывает PDF системной программой просмотра."""
    path = str(pdf_path)
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        if sys.platform == "darwin":
            subprocess.run(["open", path], check=True)
            return True
        result = subprocess.run(["xdg-open", path], check=False)
        return result.returncode == 0
    except Exception as exc:
        print(f"[!] Не удалось открыть PDF автоматически: {exc}")
        return False


def ask_output_format() -> str:
    """Спрашивает, что нужно создать: HTML, PDF или оба файла."""
    print()
    print("Что создать?")
    print()
    print("  1. HTML и PDF")
    print("  2. Только HTML")
    print("  3. Только PDF")
    print()
    answer = prompt("Выберите пункт: ", {"1", "2", "3"})
    return {"1": "both", "2": "html", "3": "pdf"}[answer]


def generate_single(
    template_text: str,
    headers: Sequence[str],
    record: Dict[str, str],
    base_name: str,
    output_dir: Path,
    base_css: str,
    output_format: str,
    open_after: bool = True,
    pdf_engine: str = "weasyprint",
) -> Tuple[Optional[Path], Optional[Path]]:
    """Создаёт HTML и/или PDF для одной записи."""
    identifier = build_record_identifier(headers, record, 1)
    stem = sanitize_filename(f"{base_name}_{identifier}", fallback="document")

    rendered, missing, used = render_template(template_text, record, base_css)
    report_template_issues(missing, headers)
    report_unused_columns(used, headers)

    html_path: Optional[Path] = None
    if output_format in ("html", "both"):
        html_path = save_html(rendered, output_dir, f"{stem}.html")
        print(f"[+] HTML создан: {html_path}")

    pdf_path: Optional[Path] = None
    if output_format in ("pdf", "both"):
        source = html_path
        if source is None:
            source = save_html(rendered, output_dir, f"{stem}.html")
            print(f"[+] Временный HTML создан: {source}")
        pdf_path = create_pdf(source, output_dir / f"{stem}.pdf", engine=pdf_engine)
        print(f"[+] PDF создан:  {pdf_path}")
        if output_format == "pdf":
            try:
                source.unlink()
                print("[i] Временный HTML удалён.")
            except OSError:
                pass
            html_path = None

    if pdf_path is not None and open_after:
        if prompt_yes_no("Открыть PDF в системной программе просмотра?"):
            if not open_pdf(pdf_path):
                print(f"    Откройте файл вручную: {pdf_path}")

    return html_path, pdf_path


def generate_all(
    template_text: str,
    headers: Sequence[str],
    records: Sequence[Dict[str, str]],
    base_name: str,
    output_dir: Path,
    base_css: str,
    output_format: str,
    pdf_engine: str = "weasyprint",
) -> List[Tuple[Optional[Path], Optional[Path]]]:
    """Создаёт отдельный документ для каждой записи CSV."""
    results: List[Tuple[Optional[Path], Optional[Path]]] = []
    unknown_columns: set = set()
    used_columns: set = set()

    for position, record in enumerate(records, start=1):
        identifier = build_record_identifier(headers, record, position)
        stem = sanitize_filename(f"{base_name}_{identifier}", fallback=f"document_{position:03d}")

        rendered, missing, used = render_template(template_text, record, base_css)
        unknown_columns.update(missing)
        used_columns.update(used)

        html_path: Optional[Path] = None
        if output_format in ("html", "both"):
            html_path = save_html(rendered, output_dir, f"{stem}.html")

        pdf_path: Optional[Path] = None
        if output_format in ("pdf", "both"):
            source = html_path or save_html(rendered, output_dir, f"{stem}.html")
            pdf_path = create_pdf(source, output_dir / f"{stem}.pdf", engine=pdf_engine)

        label = pdf_path.name if pdf_path is not None else (html_path.name if html_path else stem)
        print(f"[{position}/{len(records)}] {label}")

        if html_path is not None and output_format == "pdf":
            try:
                html_path.unlink()
            except OSError:
                pass
            html_path = None

        results.append((html_path, pdf_path))

    if unknown_columns:
        listed = ", ".join(sorted(f"{{{{{name}}}}}" for name in unknown_columns))
        print(f"[!] Плейсхолдеры без данных (оставлены как есть): {listed}")

    report_unused_columns(sorted(used_columns), headers)

    return results


def print_menu(templates: Sequence[Path], csv_files: Sequence[Path]) -> None:
    """Показывает главное меню программы."""
    print()
    print(LINE)
    print(f"{APP_TITLE:>{len(LINE)}}")
    print(LINE)
    print()
    print(f"Найдено шаблонов: {len(templates)}")
    print(f"Найдено CSV-файлов: {len(csv_files)}")
    print()
    print("1. Создать один чек")
    print("2. Создать документы для всех записей CSV")
    print("3. Обновить список файлов")
    print("0. Выход")


def choose_template(templates: Sequence[Path]) -> Tuple[Path, str]:
    """Выбирает шаблон и возвращает пару (путь, имя для выходных файлов)."""
    if not templates:
        raise UserError(
            "HTML-шаблоны не найдены.\n"
            f"    Положите файлы .html в папку «{BASE_DIR / 'templates'}»\n"
            "    или запустите программу с ключом --dirs <путь>."
        )
    path = select_file("Доступные шаблоны:", templates, show_dirs=True)
    return path, sanitize_filename(path.stem, fallback="document")


def choose_csv(csv_files: Sequence[Path]) -> Tuple[Path, List[str], List[Dict[str, str]]]:
    """Выбирает CSV, читает его и возвращает (путь, заголовки, записи)."""
    if not csv_files:
        raise UserError(
            "CSV-файлы не найдены.\n"
            f"    Положите файлы .csv в папку «{BASE_DIR / 'data'}»\n"
            "    или запустите программу с ключом --dirs <путь>."
        )
    path = select_file("Доступные файлы данных:", csv_files, show_dirs=True)
    headers, records = read_csv(path)
    return path, headers, records


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Разбирает необязательные аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Генератор чеков и документов из HTML-шаблонов и CSV-данных.",
    )
    parser.add_argument(
        "--dirs",
        nargs="*",
        default=None,
        metavar="ПАПКА",
        help="Директории для поиска .html и .csv (по умолчанию templates, data, receipts).",
    )
    parser.add_argument(
        "--output",
        default=None,
        metavar="ПАПКА",
        help="Папка для готовых файлов (по умолчанию ./output).",
    )
    parser.add_argument(
        "--fonts",
        default=None,
        metavar="ПАПКА",
        help="Папка с файлами шрифтов для @font-face (по умолчанию ./fonts).",
    )
    parser.add_argument(
        "--pdf-engine",
        choices=sorted(PDF_ENGINES),
        default="weasyprint",
        metavar="ДВИЖОК",
        help=(
            "Движок PDF: weasyprint (по умолчанию), auto (WeasyPrint, затем "
            "headless-браузер Chrome/Edge), browser."
        ),
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Только проверить все CSV на противоречия в датах и ничего не создавать.",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Не открывать созданный PDF автоматически.",
    )
    return parser.parse_args(argv)


def run_consistency_check(csv_files: Sequence[Path]) -> int:
    """Проверяет все записи всех CSV и печатает отчёт. Возвращает код выхода."""
    if not csv_files:
        print("[!] CSV-файлы не найдены — проверять нечего.")
        return 1

    total_records = 0
    total_problems = 0

    for path in csv_files:
        try:
            headers, records = read_csv(path)
        except UserError as exc:
            print(f"\n[!] {exc}")
            total_problems += 1
            continue

        print()
        print(f"--- {display_path(path)} ({len(records)} записей) ---")
        file_problems = 0
        for position, record in enumerate(records, start=1):
            total_records += 1
            issues = validate_record(headers, record)
            if not issues:
                continue
            file_problems += 1
            print(format_consistency_issues(issues, format_record_identity(headers, record, position)))
            print()

        total_problems += file_problems
        if not file_problems:
            print("Противоречий не найдено.")

    print(LINE)
    print(f"Проверено записей: {total_records}")
    print(f"Записей с противоречиями: {total_problems}")
    return 1 if total_problems else 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Точка входа: сканирование файлов и работа с консольным меню."""
    args = parse_args(argv)

    search_dirs = (
        [Path(item).expanduser() for item in args.dirs] if args.dirs else list(DEFAULT_SEARCH_DIRS)
    )
    output_dir = Path(args.output).expanduser() if args.output else DEFAULT_OUTPUT_DIR
    fonts_dir = Path(args.fonts).expanduser() if args.fonts else DEFAULT_FONTS_DIR
    pdf_engine = args.pdf_engine
    auto_open = not args.no_open

    print(LINE)
    print(f"{APP_TITLE:>{len(LINE)}}")
    print(LINE)
    print("Папки поиска:")
    for directory in search_dirs:
        mark = "ok" if directory.is_dir() else "нет"
        print(f"  - {directory}  [{mark}]")
    print(f"Папка результатов: {output_dir}")

    if not weasyprint_available():
        print("[!] WeasyPrint недоступен — PDF не удастся создать им.")
        if pdf_engine == "weasyprint":
            print("    Установка: pip install -r requirements.txt")
            print("    Либо используйте --pdf-engine auto, чтобы задействовать резервный движок.")
    else:
        print(f"[+] PDF-движок: {pdf_engine}")

    base_css = build_base_css(fonts_dir)
    if "ReceiptFont" in base_css:
        print("[+] Подключены локальные шрифты через @font-face.")
    else:
        print("[i] Локальные шрифты не найдены, используется системный fallback.")

    templates, csv_files = scan_all(search_dirs)

    if args.check:
        return run_consistency_check(csv_files)

    while True:
        print_menu(templates, csv_files)
        choice = prompt("Выберите пункт меню: ", {"0", "1", "2", "3"})

        if choice == "0":
            print("Работа завершена.")
            return 0

        if choice == "3":
            templates, csv_files = scan_all(search_dirs)
            print("\nСписок файлов обновлён.")
            continue

        try:
            if choice == "1":
                template_path, base_name = choose_template(templates)
                csv_path, headers, records = choose_csv(csv_files)
                template_text = load_template(template_path)
                record = select_record(headers, records)
                print()
                print("Выбрана запись:")
                print(f"  {format_record_preview(headers, record, limit=6)}")
                print(f"  Шаблон: {display_path(template_path)}")
                print(f"  Данные:  {display_path(csv_path)}")
                if not confirm_single_generation(
                    headers, record, format_record_identity(headers, record, 1)
                ):
                    print("[*] Создание отменено из-за противоречий в данных.")
                    continue
                output_format = ask_output_format()
                generate_single(
                    template_text, headers, record, base_name,
                    output_dir, base_css, output_format,
                    open_after=auto_open, pdf_engine=pdf_engine,
                )
            else:
                template_path, base_name = choose_template(templates)
                csv_path, headers, records = choose_csv(csv_files)
                template_text = load_template(template_path)
                print()
                print(f"Будет обработано записей: {len(records)}")
                if not confirm_batch_generation(headers, records):
                    print("[*] Создание отменено из-за противоречий в данных.")
                    continue
                output_format = ask_output_format()
                results = generate_all(
                    template_text, headers, records, base_name,
                    output_dir, base_css, output_format, pdf_engine=pdf_engine,
                )
                pdf_paths = [pdf for _, pdf in results if pdf is not None]
                print()
                print(f"[+] Готово документов: {len(results)}")
                if pdf_paths and auto_open and prompt_yes_no(
                    f"Открыть первый PDF ({len(pdf_paths)} шт.)?"
                ):
                    if not open_pdf(pdf_paths[0]):
                        print(f"    Откройте файл вручную: {pdf_paths[0]}")
        except UserCancelled:
            print("\n[*] Операция отменена.")
        except UserError as exc:
            print(f"\n[!] {exc}")
        except KeyboardInterrupt:
            print("\n[*] Прервано пользователем.")
            return 130


if __name__ == "__main__":
    try:
        sys.exit(main())
    except UserCancelled:
        print("\nРабота завершена.")
        sys.exit(0)
    except KeyboardInterrupt:
        print("\nПрервано пользователем.")
        sys.exit(130)
