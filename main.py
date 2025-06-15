#!/usr/bin/env python3
# transform_1803.py

import pandas as pd
import argparse
import re
import sys
import os

# Фиксированное имя файла-справочника в той же папке
MAPPING_FILE = "Type.xlsx"

def detect_block_ranges(path, valid_types):
    df = pd.read_excel(path, header=None, dtype=str)
    valid_types_set = set(valid_types)
    block_starts = df[df.apply(lambda row: row.astype(str).str.strip().isin(valid_types_set).any(), axis=1)].index.tolist()
    block_ranges = [(start, block_starts[i + 1] - 1 if i + 1 < len(block_starts) else df.shape[0] - 1)
                    for i, start in enumerate(block_starts)]
    return block_ranges

def load_type_map():
    if not os.path.exists(MAPPING_FILE):
        raise FileNotFoundError(f"Файл справочника не найден: {MAPPING_FILE}")
    dfm = pd.read_excel(MAPPING_FILE, dtype={'Type': str})
    required = {'Type', 'Width', 'Depth'}
    if not required.issubset(dfm.columns):
        raise ValueError(f"Файл {MAPPING_FILE} должен содержать колонки: {required}")
    return dict(zip(dfm['Type'].astype(str), zip(dfm['Width'], dfm['Depth'])))

def transform_file(input_path, output_path, verbose=False):
    type_map = load_type_map()
    valid_types = list(type_map.keys())
    if verbose:
        print(f"Загружено {len(type_map)} типов из {MAPPING_FILE}")

    block_ranges = detect_block_ranges(input_path, valid_types)
    if verbose:
        print(f"Обнаружено блоков: {len(block_ranges)} → {block_ranges}")

    rows = []
    for block_start, block_end in block_ranges:
        df = pd.read_excel(input_path, header=[block_start, block_start + 1, block_start + 2], nrows=block_end - block_start - 2, dtype=str)
        lvl2 = df.columns.get_level_values(2)
        mask = lvl2.str.contains('№', na=False) | lvl2.str.contains('размер', na=False)
        df_f = df.loc[:, mask]

        lvl1 = df.columns.get_level_values(1)
        groups = []
        for g in lvl1[mask]:
            if pd.notna(g) and g not in groups and re.search(r'\d+', str(g)):
                groups.append(g)

        for g in groups:
            mo = re.search(r'(\d+)', str(g))
            if not mo:
                print(f"⚠ Пропускаем группу без номера: '{g}'", file=sys.stderr)
                continue
            nst = int(mo.group(1))

            types = df_f.loc[:, (slice(None), g, slice(None))].columns.get_level_values(0).unique()
            for t in types:
                try:
                    s_num  = df_f[(t, g, '№')]
                    s_size = df_f[(t, g, 'размер')]
                except KeyError:
                    continue

                size_mm = s_size.astype(str).str.extract(r'(\d+)').astype(float)[0]
                size_cm = ((size_mm - 3) / 10).round(1)

                tmp = pd.DataFrame({
                    'nst':    nst,
                    'nsafe':  pd.to_numeric(s_num, errors='coerce'),
                    'height': size_cm,
                    'Type':   t,
                })
                tmp = tmp[tmp['nsafe'].notna() & tmp['height'].notna()]
                tmp = tmp[tmp['nsafe'] != 0]
                if tmp.empty:
                    continue
                tmp['nsafe'] = tmp['nsafe'].astype(int)

                if t in type_map:
                    w, d = type_map[t]
                else:
                    print(f"⚠ Нет справочника для Type='{t}', ставим Width=26, Depth=39", file=sys.stderr)
                    w, d = 26, 39
                tmp['Width'] = w
                tmp['Depth'] = d

                tmp = tmp[['nst','nsafe','height','Width','Depth','Type']]
                rows.append(tmp)

    if not rows:
        print("❌ Ошибка: нет строк для обработки.", file=sys.stderr)
        sys.exit(1)

    df_out = pd.concat(rows, ignore_index=True)
    df_out.to_excel(output_path, index=False)
    print(f"✅ Готово: {len(df_out)} строк, сохранено в '{output_path}'")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Transform Excel → длинный формат с Width/Depth (Type.xlsx фиксирован, автоматическое разделение блоков)"
    )
    parser.add_argument('input',  help='Путь к исходному файлу')
    parser.add_argument('output', nargs='?', help='(опц.) Выходной файл')
    parser.add_argument('--verbose', action='store_true', help='Подробный вывод')
    args = parser.parse_args()

    out_path = args.output if args.output else f"{os.path.splitext(args.input)[0]}_transformed{os.path.splitext(args.input)[1]}"

    try:
        transform_file(args.input, out_path, verbose=args.verbose)
    except Exception as e:
        print("❌ Ошибка:", e, file=sys.stderr)
        sys.exit(1)
