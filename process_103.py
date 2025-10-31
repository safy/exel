#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
process_103.py
Обработка исходной таблицы "данные.xlsx" в формат "Таблица1" с правилами из переписки.

Правила:
- famo: оставить ТОЛЬКО ФИО (3 слова кириллицей), без "Паспорт ..." и пр.
- adro: убрать переносы и лишние пробелы; склеить разорванные слова; нормализовать "Г. ", "УЛ. ", "Б-Р ".
  Исправления: 
    "Г.НЯЗЕПЕ ТРОВСК" -> "Г. НЯЗЕПЕТРОВСК"
    "УЛ.ФРУН ЗЕ"      -> "УЛ. ФРУНЗЕ"
    "УЛ.ЛОМО НОСОВА"  -> "УЛ. ЛОМОНОСОВА"
    "Б-Р ГАЙДАР А"    -> "Б-Р ГАЙДАРА"
    "БЕРЕГОВ ОЙ"      -> "БЕРЕГОВОЙ"
- oblo: "Челябинская область" (если не задано иное через аргумент)
- indo (почтовый индекс) по адресу:
    Озерск + (Гайдара/Гайдара) -> 456785, а для домов {3,4,6,8,10,12} -> 456789
    Нязепетровск + Фрунзе -> 456971
    Касли + Ломоносова -> дом 39: 456835, иначе: 456830
  Для прочих адресов индекс оставить пустым.
- tipo: по умолчанию 0; если в "Платежные реквизиты" есть фраза "счёт получателя платежа" (счет/счёт) -> 2
- inni (ИНН отправителя): задаётся аргументом --inn ИЛИ берётся из шаблона (--template) из колонки "inni".
- Сумма выплаты берётся из колонки вида "Сумма к выплате"/"Сумма выплаты" (автопоиск).

Выходной Excel содержит колонки:
['inni','tipo','indo','oblo','adro','famo','telo','inno','biko','scho','kscho','nampo','nambo','msg','sumi']

Установка зависимостей:
    pip install pandas openpyxl

Примеры запуска:
    python process_103.py --input данные.xlsx --output 103_result.xlsx --inn 7409002190
    python process_103.py --input данные.xlsx --output 103_result.xlsx --template таблица1.xlsx
    python process_103.py --input данные.xlsx --output 103_result.xlsx --inn 7409002190 --region "Челябинская область"
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from typing import Iterable, List, Optional

import pandas as pd

COLUMNS_OUT: List[str] = [
    "inni",
    "tipo",
    "indo",
    "oblo",
    "adro",
    "famo",
    "telo",
    "inno",
    "biko",
    "scho",
    "kscho",
    "nampo",
    "nambo",
    "msg",
    "sumi",
]


@dataclass
class ProcessResult:
    """Информация об обработке файла."""

    total_rows: int
    tipo2_rows: int
    output_path: str


def find_col(df: pd.DataFrame, contains_list: Iterable[str]) -> Optional[str]:
    for column in df.columns:
        name = str(column).lower()
        if all(part.lower() in name for part in contains_list):
            return column
    return None


def normalize_spaces(value: str) -> str:
    text = str(value).replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def fix_known_breaks(value: str) -> str:
    """Нормализация известных разрывов и сокращений."""

    upper = value.upper()
    upper = re.sub(r"НЯЗЕПЕ\s*ТРОВСК", "НЯЗЕПЕТРОВСК", upper)
    upper = re.sub(r"ФРУН\s*ЗЕ", "ФРУНЗЕ", upper)
    upper = re.sub(r"ЛОМО\s*НОСОВА", "ЛОМОНОСОВА", upper)
    upper = re.sub(r"ГАЙДАР\s*А", "ГАЙДАРА", upper)
    upper = re.sub(r"БЕРЕГОВ\s*ОЙ", "БЕРЕГОВОЙ", upper)
    upper = re.sub(r"Г\.\s*", "Г. ", upper)
    upper = re.sub(r"УЛ\.\s*", "УЛ. ", upper)
    upper = re.sub(r"Б-Р\s*", "Б-Р ", upper)
    upper = re.sub(r"\s+", " ", upper).strip()
    return upper


def extract_fio_only(text: str) -> Optional[str]:
    if not isinstance(text, str):
        return None
    upper = text.upper()
    upper = upper.split(" ПАСПОРТ")[0]
    parts = re.findall(r"[А-ЯЁ\-]+", upper)
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1]} {parts[2]}"
    match = re.search(r"([А-ЯЁ][А-ЯЁ\-]+)\s+([А-ЯЁ][А-ЯЁ\-]+)\s+([А-ЯЁ][А-ЯЁ\-]+)", upper)
    return " ".join(match.groups()) if match else normalize_spaces(upper) if upper else None


def extract_house(addr: str) -> Optional[int]:
    if not isinstance(addr, str):
        return None
    match = re.search(r"(\d+)(?:\s*-\s*\d+)?\s*$", addr.strip())
    return int(match.group(1)) if match else None


def resolve_index(addr: str) -> Optional[int]:
    if not isinstance(addr, str):
        return None
    address_upper = addr.upper()
    house = extract_house(addr)
    if "ОЗЕРСК" in address_upper and ("ГАЙДАР" in address_upper or "ГАЙДАРА" in address_upper):
        return 456789 if house in {3, 4, 6, 8, 10, 12} else 456785
    if "НЯЗЕПЕТРОВСК" in address_upper and "ФРУНЗЕ" in address_upper:
        return 456971
    if "КАСЛИ" in address_upper and "ЛОМОНОСОВА" in address_upper:
        return 456835 if house == 39 else 456830
    return None


def has_schet_poluchatelya(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return re.search(r"сч[её]т\s+получателя\s+платежа", text.lower()) is not None


def load_inn(template_path: Optional[str]) -> Optional[str]:
    if not template_path:
        return None
    try:
        template_df = pd.read_excel(template_path)
    except Exception as exc:  # pragma: no cover - отчёт пользователю
        print(f"Не удалось получить INN из шаблона: {exc}", file=sys.stderr)
        return None
    if "inni" not in template_df.columns:
        return None
    series = template_df["inni"].dropna()
    return str(series.iloc[0]) if not series.empty else None


def detect_sum_column(data_df: pd.DataFrame) -> Optional[str]:
    return (
        find_col(data_df, ["Сумма", "выплат"])  # type: ignore[list-item]
        or find_col(data_df, ["Сумма к выплате"])
        or find_col(data_df, ["Сумма дохода"])
        or find_col(data_df, ["Сумма"])
    )


def process_excel(
    input_path: str,
    output_path: str,
    *,
    inn: Optional[str] = None,
    template_path: Optional[str] = None,
    region: str = "Челябинская область",
) -> ProcessResult:
    """Обработка входного Excel-файла согласно правилам."""

    try:
        data_df = pd.read_excel(input_path)
    except Exception as exc:  # pragma: no cover - отчёт пользователю
        print(f"Ошибка чтения файла {input_path}: {exc}", file=sys.stderr)
        raise

    inn_value = inn or load_inn(template_path)
    if inn_value is None:
        raise ValueError("ИНН не задан. Передайте --inn или --template с колонкой 'inni'.")

    fio_col = next((c for c in data_df.columns if str(c).startswith("Фамилия")), None)
    if fio_col is None:
        raise ValueError("Не найден столбец с ФИО (начинается с 'Фамилия').")

    addr_col = find_col(data_df, ["Адрес", "почтовый"]) or find_col(data_df, ["Адрес"])
    if addr_col is None:
        raise ValueError("Не найден столбец с адресом (содержит 'Адрес').")

    pay_col = find_col(data_df, ["Платежные", "реквизиты"]) or find_col(data_df, ["реквизиты"])
    sum_col = detect_sum_column(data_df)

    rows = []
    for _, row in data_df.iterrows():
        fio_raw = row.get(fio_col)
        addr_raw = row.get(addr_col)
        if pd.isna(fio_raw) or pd.isna(addr_raw):
            continue

        addr_fix = fix_known_breaks(normalize_spaces(addr_raw))
        fio_fix = extract_fio_only(fio_raw)
        sumi = row.get(sum_col) if sum_col else None
        tipo = 2 if has_schet_poluchatelya(str(row.get(pay_col, ""))) else 0

        rows.append(
            {
                "inni": inn_value,
                "tipo": tipo,
                "indo": resolve_index(addr_fix),
                "oblo": region,
                "adro": addr_fix,
                "famo": fio_fix,
                "telo": None,
                "inno": None,
                "biko": None,
                "scho": None,
                "kscho": None,
                "nampo": None,
                "nambo": None,
                "msg": None,
                "sumi": sumi,
            }
        )

    out_df = pd.DataFrame(rows, columns=COLUMNS_OUT)
    try:
        out_df.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - отчёт пользователю
        print(f"Ошибка записи файла {output_path}: {exc}", file=sys.stderr)
        raise

    total_rows = len(out_df)
    tipo2_rows = int((out_df["tipo"] == 2).sum()) if total_rows else 0
    return ProcessResult(total_rows=total_rows, tipo2_rows=tipo2_rows, output_path=output_path)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Преобразование данных в формат Таблица1 (Ф.103)."
    )
    parser.add_argument("--input", "-i", required=True, help="Путь к исходному Excel")
    parser.add_argument("--output", "-o", required=True, help="Путь к выходному Excel")
    parser.add_argument("--inn", type=str, default=None, help="Общий ИНН отправителя")
    parser.add_argument(
        "--template",
        type=str,
        default=None,
        help="Путь к шаблону Таблица1.xlsx, чтобы взять inni",
    )
    parser.add_argument(
        "--region", type=str, default="Челябинская область", help="Название региона для oblo"
    )
    return parser


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    try:
        result = process_excel(
            args.input,
            args.output,
            inn=args.inn,
            template_path=args.template,
            region=args.region,
        )
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"Готово: {result.total_rows} записей. tipo=2 для {result.tipo2_rows} записей.")
    print(f"Файл сохранён: {result.output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
