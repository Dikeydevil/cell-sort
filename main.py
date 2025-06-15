# Updated transform_1803.py with automatic output naming

import pandas as pd
import argparse
import re
import sys
import os

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

def transform_file(input_path: str, output_path: str, headers=None, verbose=False):
    if headers:
        try:
            h0, h1, h2 = map(int, headers.split(','))
        except:
            print("ERROR: --headers должен быть вида H0,H1,H2", file=sys.stderr)
            sys.exit(1)
    else:
        h0, h1, h2 = detect_header_rows(input_path)
    if verbose:
        print(f"→ Используем строки заголовков: {h0}, {h1}, {h2}")

    df = pd.read_excel(input_path, header=[h0, h1, h2], dtype=str)
    lvl2 = df.columns.get_level_values(2)
    mask = lvl2.str.contains('№', na=False) | lvl2.str.contains('размер', na=False)
    df_f = df.loc[:, mask]

    groups = []
    lvl1 = df.columns.get_level_values(1)
    for g in lvl1[mask]:
        if pd.notna(g) and g not in groups and re.search(r'\d+', str(g)):
            groups.append(g)
    if verbose:
        print(f"Найдено групп: {len(groups)} → {groups}")

    frames = []
    for g in groups:
        mo = re.search(r'(\d+)', str(g))
        if not mo:
            print(f"⚠️ Пропускаем группу без числа: '{g}'", file=sys.stderr)
            continue
        nst = int(mo.group(1))
        types = df_f.loc[:, (slice(None), g, slice(None))].columns.get_level_values(0).unique()
        for t in types:
            try:
                s_num  = df_f[(t, g, '№')]
                s_size = df_f[(t, g, 'размер')]
            except KeyError:
                continue
            tmp = pd.DataFrame({
                'nst':   nst,
                'Type':  t,
                'nsafe': pd.to_numeric(s_num, errors='coerce'),
                'height': s_size,
            })
            tmp = tmp[tmp['nsafe'].notna() & tmp['height'].notna()]
            if not tmp.empty:
                tmp['nsafe'] = tmp['nsafe'].astype(int)
                frames.append(tmp[['nst','nsafe','height','Type']])

    if not frames:
        print("❌ Ошибка: не найдено ни одной группы.", file=sys.stderr)
        sys.exit(1)

    df_out = pd.concat(frames, ignore_index=True)
    df_out.to_excel(output_path, index=False)
    print(f"✅ Обработано {len(df_out)} строк, сохранено в '{output_path}'")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Превращает 1803.xlsx в длинный формат Example2"
    )
    parser.add_argument('input',  help='Исходный Excel-файл')
    parser.add_argument('output', nargs='?', help='(Необязательно) Файл для результата')
    parser.add_argument('--headers', help='(Необязательно) H0,H1,H2')
    parser.add_argument('--verbose', action='store_true', help='Показывать детали')
    args = parser.parse_args()

    out_path = args.output
    if not out_path:
        base, ext = os.path.splitext(args.input)
        out_path = f"{base}_transformed{ext}"

    try:
        transform_file(args.input, out_path, headers=args.headers, verbose=args.verbose)
    except Exception as e:
        print("❌ Ошибка:", e, file=sys.stderr)
        sys.exit(1)
