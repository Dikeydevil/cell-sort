#!/usr/bin/env python3
# transform_1803.py

import pandas as pd
import argparse
import re
import sys
import os

# Фиксированное имя файла-справочника в той же папке
MAPPING_FILE = "Type.xlsx"

def detect_header_rows(path, max_scan_rows=50):
    df0 = pd.read_excel(path, header=None, nrows=max_scan_rows, dtype=str)
    header_row = None
    for i, row in df0.iterrows():
        vals = row.dropna().astype(str).str.strip().tolist()
        if vals.count('№') >= 3 and vals.count('размер') >= 3:
            header_row = i
            break
    if header_row is None or header_row < 2:
        raise RuntimeError("Не удалось автоматически найти строку заголовков с '№' и 'размер'.")
    return header_row-2, header_row-1, header_row

def load_type_map():
    """
    Читает фиксированный файл Type.xlsx рядом со скриптом,
    содержащий колонки ['Type','Width','Depth'].
    Возвращает словарь: type_map[Type] = (Width, Depth)
    """
    if not os.path.exists(MAPPING_FILE):
        raise FileNotFoundError(f"Файл справочника не найден: {MAPPING_FILE}")
    dfm = pd.read_excel(MAPPING_FILE, dtype={'Type': str})
    required = {'Type', 'Width', 'Depth'}
    if not required.issubset(dfm.columns):
        raise ValueError(f"Файл {MAPPING_FILE} должен содержать колонки: {required}")
    return dict(zip(dfm['Type'], zip(dfm['Width'], dfm['Depth'])))

def transform_file(input_path, output_path, headers=None, verbose=False):
    # 1) загрузка маппинга
    type_map = load_type_map()
    if verbose:
        print(f"Загружено {len(type_map)} типов из {MAPPING_FILE}")

    # 2) определение строк заголовков
    if headers:
        try:
            h0, h1, h2 = map(int, headers.split(','))
        except:
            print("ERROR: --headers должен быть вида H0,H1,H2", file=sys.stderr)
            sys.exit(1)
    else:
        h0, h1, h2 = detect_header_rows(input_path)
    if verbose:
        print(f"Используем строки заголовков: {h0},{h1},{h2}")

    # 3) чтение исходного файла
    df = pd.read_excel(input_path, header=[h0, h1, h2], dtype=str)

    # 4) фильтрация по 3-му уровню '№' или 'размер'
    lvl2 = df.columns.get_level_values(2)
    mask = lvl2.str.contains('№', na=False) | lvl2.str.contains('размер', na=False)
    df_f = df.loc[:, mask]

    # 5) определяем группы (lvl1) в порядке колонок
    lvl1 = df.columns.get_level_values(1)
    groups = []
    for g in lvl1[mask]:
        if pd.notna(g) and g not in groups and re.search(r'\d+', str(g)):
            groups.append(g)
    if verbose:
        print(f"Найдено групп: {len(groups)} → {groups}")

    # 6) сбор строк
    rows = []
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

            # Преобразуем значения размера: вычесть 3 мм, перевести в см и округлить до 1 знака
            size_mm = s_size.astype(str).str.extract(r'(\d+)').astype(float)[0]
            size_cm = ((size_mm - 3) / 10).round(1)

            tmp = pd.DataFrame({
                'nst':    nst,
                'nsafe':  pd.to_numeric(s_num, errors='coerce'),
                'height': size_cm,
                'Type':   t,
            })
            tmp = tmp[tmp['nsafe'].notna() & tmp['height'].notna()]
            tmp = tmp[tmp['nsafe'] != 0]  # исключаем ячейки со значением 0
            if tmp.empty:
                continue
            tmp['nsafe'] = tmp['nsafe'].astype(int)

            # подстановка Width/Depth
            if t in type_map:
                w, d = type_map[t]
            else:
                print(f"⚠ Нет справочника для Type='{t}', ставим Width=26, Depth=39", file=sys.stderr)
                w, d = 26, 39
            tmp['Width'] = w
            tmp['Depth'] = d

            # порядок колонок
            tmp = tmp[['nst','nsafe','height','Width','Depth','Type']]
            rows.append(tmp)

    if not rows:
        print("❌ Ошибка: нет строк для обработки.", file=sys.stderr)
        sys.exit(1)

    # 7) объединение и сохранение
    df_out = pd.concat(rows, ignore_index=True)
    df_out.to_excel(output_path, index=False)
    print(f"✅ Готово: {len(df_out)} строк, сохранено в '{output_path}'")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Transform 1803.xlsx → длинный формат с Width/Depth (Type.xlsx фиксирован)"
    )
    parser.add_argument('input',  help='Путь к исходному файлу')
    parser.add_argument('output', nargs='?', help='(опц.) Выходной файл')
    parser.add_argument('--headers', help='(опц.) Индексы строк заголовков H0,H1,H2')
    parser.add_argument('--verbose', action='store_true', help='Подробный вывод')
    args = parser.parse_args()

    out_path = args.output if args.output else f"{os.path.splitext(args.input)[0]}_transformed{os.path.splitext(args.input)[1]}"

    try:
        transform_file(args.input, out_path, headers=args.headers, verbose=args.verbose)
    except Exception as e:
        print("❌ Ошибка:", e, file=sys.stderr)
        sys.exit(1)
