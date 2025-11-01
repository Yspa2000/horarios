#!/usr/bin/env python3
from pathlib import Path
import argparse
import re
from datetime import datetime
import pandas as pd

MONTHS_ES = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
    'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
    'septiembre': 9, 'setiembre': 9, 'octubre': 10,
    'noviembre': 11, 'diciembre': 12
}

DATE_CELL_REGEX = re.compile(r'^\\s*(\\d{1,2})\\s*[-/]\\s*([A-Za-záéíóúÁÉÍÓÚ]+)\\s*$', re.IGNORECASE)

def parse_args():
    p = argparse.ArgumentParser(description="Parsea horarios tipo Excel/CSV y exporta a .xlsx con hojas por departamento.")
    p.add_argument('--input', '-i', required=True, help='Ruta al .csv o .xlsx con el horario base')
    p.add_argument('--year', '-y', type=int, default=None, help='Año de las fechas (si no se pasa se infiere)')
    p.add_argument('--output', '-o', default='normalized_schedule.xlsx', help='Archivo .xlsx de salida')
    p.add_argument('--csv-output', default=None, help='Genera además un CSV con este nombre (opcional)')
    p.add_argument('--include-empty', action='store_true', help='Incluir celdas vacías como filas en la salida')
    p.add_argument('--no-infer-year', action='store_true', help='No inferir año automáticamente; usar año actual si no se pasa --year')
    return p.parse_args()

def read_table(path: Path):
    if path.suffix.lower() in ('.xls', '.xlsx'):
        df = pd.read_excel(path, header=None, dtype=str)
    else:
        try:
            df = pd.read_csv(path, header=None, dtype=str, engine='python', keep_default_na=False)
        except Exception:
            df = pd.read_csv(path, header=None, dtype=str, sep=';', engine='python', keep_default_na=False)
    df = df.fillna('')
    df = df.astype(str)
    return df

def find_date_header_row(row):
    for cell in row:
        if DATE_CELL_REGEX.match(cell.strip()):
            return True
    return False

def parse_date_cell(cell_text, year):
    m = DATE_CELL_REGEX.match(cell_text.strip())
    if not m:
        return None
    day = int(m.group(1))
    month_str = m.group(2).lower()
    month = MONTHS_ES.get(month_str)
    if not month:
        month = MONTHS_ES.get(month_str.replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u'))
    if not month:
        return None
    try:
        return datetime(year, month, day).date()
    except Exception:
        return None

def normalize_time_part(s):
    s = s.strip()
    if not s:
        return ''
    s = s.replace(',', ':').replace('.', ':')
    s = re.sub(r'[^0-9:]', '', s)
    parts = s.split(':')
    try:
        if len(parts) == 1:
            hh = int(parts[0]) if parts[0] else 0
            mm = 0
        else:
            hh = int(parts[0]) if parts[0] else 0
            mm = int(parts[1]) if parts[1] else 0
        return f"{hh:02d}:{mm:02d}"
    except Exception:
        return ''

def split_shift_to_times(shift_raw):
    if not shift_raw or shift_raw.strip() == '':
        return ('', '', '')
    s = shift_raw.strip()
    if '-' in s:
        left, right = s.split('-', 1)
        start = normalize_time_part(left)
        end = normalize_time_part(right)
        return (s, start, end)
    m = re.search(r'(\\d{1,2}[:.,]\\d{1,2})', s)
    if m:
        t = normalize_time_part(m.group(1))
        return (s, t, '')
    return (s, '', '')

def extract_sections(df):
    nrows, ncols = df.shape
    i = 0
    while i < nrows:
        row = df.iloc[i].tolist()
        if find_date_header_row(row):
            dept = ''
            j = i - 1
            while j >= 0:
                cand = str(df.iat[j, 0]).strip()
                if cand:
                    dept = cand
                    break
                j -= 1
            date_map = {}
            for col in range(ncols):
                cell = df.iat[i, col].strip()
                if cell:
                    m = DATE_CELL_REGEX.match(cell)
                    if m:
                        date_map[col] = cell.strip()
            emp_rows = []
            k = i + 1
            while k < nrows:
                row_k = df.iloc[k].tolist()
                if find_date_header_row(row_k):
                    break
                if all((str(x).strip() == '' for x in row_k)):
                    break
                emp_rows.append((k, row_k))
                k += 1
            yield (dept, date_map, emp_rows)
            i = k
            continue
        i += 1

def infer_year_from_date_cells(df, provided_year=None, no_infer=False):
    if provided_year:
        return provided_year
    if no_infer:
        return datetime.now().year
    detected_months = set()
    for dept, date_map, _ in extract_sections(df):
        for col, cell_text in date_map.items():
            m = DATE_CELL_REGEX.match(cell_text)
            if m:
                month_str = m.group(2).lower()
                month = MONTHS_ES.get(month_str) or MONTHS_ES.get(month_str.replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u'))
                if month:
                    detected_months.add(month)
    now = datetime.now()
    if not detected_months:
        return now.year
    min_month = min(detected_months)
    if now.month >= 10 and min_month <= 3:
        return now.year + 1
    return now.year

def build_normalized(df, year, include_empty=False):
    rows_out = []
    for dept, date_map, emp_rows in extract_sections(df):
        if not date_map:
            continue
        date_cols = {}
        for col, cell_text in date_map.items():
            dt = parse_date_cell(cell_text, year)
            if dt:
                date_cols[col] = dt
        if not date_cols:
            continue
        for (rindex, row_vals) in emp_rows:
            name = str(row_vals[0]).strip()
            if not name:
                continue
            for col_idx, date in date_cols.items():
                raw_shift = str(row_vals[col_idx]).strip() if col_idx < len(row_vals) else ''
                if raw_shift == '' and not include_empty:
                    continue
                raw_shift = raw_shift.replace('\xa0', ' ').strip()
                raw, start, end = split_shift_to_times(raw_shift)
                rows_out.append({
                    'department': dept,
                    'employee': name,
                    'date': date.isoformat(),
                    'raw_shift': raw_shift,
                    'start_time': start,
                    'end_time': end
                })
    return rows_out

def sanitize_sheet_name(name):
    name = re.sub(r'[:\\/?*\[\]]+', '_', str(name))
    return name[:31]

def export_to_excel(rows, excel_path, csv_out=None):
    df = pd.DataFrame(rows, columns=['department','employee','date','raw_shift','start_time','end_time'])
    df['date_dt'] = pd.to_datetime(df['date'])
    excel_path = Path(excel_path)
    engine = 'xlsxwriter'
    with pd.ExcelWriter(excel_path, engine=engine) as writer:
        depts = sorted(df['department'].fillna('UNKNOWN').unique().tolist())
        for dept in depts:
            dfn = df[df['department'] == dept].copy()
            if dfn.empty:
                continue
            dfn = dfn.sort_values(['employee','date_dt'])
            sheet_name = sanitize_sheet_name(dept or 'SIN_NOMBRE')
            dfn_out = dfn[['date','employee','raw_shift','start_time','end_time']].copy()
            dfn_out.to_excel(writer, sheet_name=sheet_name, index=False)
        pivot = df.pivot_table(index='employee', columns='date', values='raw_shift', aggfunc=lambda x: '; '.join(x.astype(str)))
        pivot_cols = sorted(pivot.columns, key=lambda d: pd.to_datetime(d))
        pivot = pivot[pivot_cols]
        pivot.to_excel(writer, sheet_name='by_employee')
        summary_rows = []
        for dept in depts:
            cnt = df[df['department'] == dept].shape[0]
            uniq_emp = df[df['department'] == dept]['employee'].nunique()
            summary_rows.append({'department': dept, 'rows': cnt, 'unique_employees': uniq_emp})
        summary_df = pd.DataFrame(summary_rows)
        totals = {'department': 'TOTAL', 'rows': df.shape[0], 'unique_employees': df['employee'].nunique()}
        summary_df = pd.concat([summary_df, pd.DataFrame([totals])], ignore_index=True)
        summary_df.to_excel(writer, sheet_name='summary', index=False)
    if csv_out:
        df_out = df[['department','employee','date','raw_shift','start_time','end_time']].copy()
        df_out.to_csv(csv_out, index=False, encoding='utf-8')

def main():
    args = parse_args()
    p = Path(args.input)
    if not p.exists():
        print(f"Archivo no encontrado: {p}")
        return
    print(f"Leyendo {p} ...")
    df = read_table(p)
    print("Detectando cabeceras de fechas y secciones...")
    inferred_year = infer_year_from_date_cells(df, provided_year=args.year, no_infer=args.no_infer_year)
    print(f"Año usado para parseo: {inferred_year} (pasado por --year: {args.year is not None})")
    rows = build_normalized(df, inferred_year, include_empty=args.include_empty)
    if not rows:
        print("No se extrajeron filas. Revisa que el archivo tenga celdas tipo '1-enero' en las cabeceras.")
        return
    excel_out = args.output
    csv_out = args.csv_output
    print(f"Generando Excel: {excel_out} ...")
    export_to_excel(rows, excel_out, csv_out)
    print(f"Excel guardado en: {excel_out}")
    if csv_out:
        print(f"CSV normalizado guardado en: {csv_out}")
    df_rows = pd.DataFrame(rows)
    depts = sorted(df_rows['department'].dropna().unique().tolist())
    print(f"Se detectaron {len(depts)} secciones/departamentos: {', '.join(depts[:20])}{'...' if len(depts)>20 else ''}")
    print(f"Filas normalizadas: {len(df_rows)}  |  Empleados únicos: {df_rows['employee'].nunique()}")

if __name__ == '__main__':
    main()