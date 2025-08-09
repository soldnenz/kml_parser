#!/usr/bin/env python3
# parse_to_json.py

import pandas as pd
import json
import sys

def read_table(filename):
    """
    Попытка прочитать таблицу сначала как Excel (.xls),
    а при неудаче — как HTML с тегами <table>.
    """
    try:
        # Читаем как настоящий Excel
        df = pd.read_excel(filename, engine='xlrd')
    except Exception:
        # Парсим HTML-таблицы
        tables = pd.read_html(filename, encoding='utf-8')
        if not tables:
            raise ValueError(f"No tables found in {filename}")
        df = tables[0]

    # Объединяем MultiIndex заголовков (двухстрочные) в один уровень
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [' '.join(filter(None, map(str, col))).strip() for col in df.columns.values]

    # Очищаем пробелы в названиях колонок
    df.columns = [str(col).strip() for col in df.columns]
    return df

def main():
    filenames = ['Unknown.xls', 'Unknown-2.xls']
    merged = []

    for fname in filenames:
        try:
            df = read_table(fname)
        except Exception as e:
            print(f"Error reading {fname}: {e}", file=sys.stderr)
            continue

        # Добавляем источник и конвертируем строки в словари
        for _, row in df.iterrows():
            record = {col: (row[col] if pd.notnull(row[col]) else None) for col in df.columns}
            record['source_file'] = fname
            merged.append(record)

    # Запись объединённого JSON
    out_path = 'zones.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)

    print(f"Merged JSON saved to {out_path}")

if __name__ == '__main__':
    main()

