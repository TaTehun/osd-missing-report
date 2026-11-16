"""
main_example.py

Reference implementation of the core processing pipeline.
Sensitive values (plant codes, WH mappings, paths) are replaced with generic placeholders.

Flow per plant: IBD load -> OSD D rows -> IBD vs OSD match -> OVG check -> master resolve -> save
"""

import re
import pandas as pd
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from utils.date_utils import previous_week, current_week, two_weeks_ago
from utils.log_utils import setup_logger
from email_pipeline import email_setup as cfg

OSD_DIR     = cfg.OSD_DIR
IBD_DIR     = cfg.IBD_DIR
MASTER_DIR  = cfg.MASTER_DIR
MISSING_DIR = Path.home() / 'Desktop' / 'your-project' / 'report'
LOG_DIR     = Path.home() / 'Desktop' / 'your-project' / 'logs'

logger, _log_file = setup_logger(LOG_DIR, 'main')

# Warehouse code -> Plant code mapping
WH_TO_PLANT = {
    'WH1': 'PLANT1',
    'WH2': 'PLANT2',
    'WH3': 'PLANT3',
}
PLANT_CODES = list(WH_TO_PLANT.values())

MISSING_FILENAME  = "{week} {code} Missing items from OSD Report.xlsx"
KEEP_MASTER_WEEKS = 15
KEEP_OVG_WEEKS    = 15

IBD_GROUP_COLS = ['WH N.', 'Week', 'CTN #', 'Serial number', 'Material', 'QC G.']
MASTER_COLS    = ['Plant', 'WH N.', 'Week', 'CTN #', 'SN', 'Material', 'Total', 'QC G.', 'Serial number']

# Fixed 13-column schema for all OSD report files
OSD_COLS = ['Plant', 'Week', 'Date', 'Container', 'SN', 'Live/Drop',
            'Seal in Tact', 'O/S/D', 'Model', 'SN QTY', 'Qty',
            'Serial Number(s)', 'WH Note']


@dataclass
class ReportResult:
    success:      list[tuple[str, bool]] = field(default_factory=list)  # (code, has_report)
    file_missing: list[str]              = field(default_factory=list)
    fail:         list[tuple[str, str]]  = field(default_factory=list)  # (code, reason)
    sn_corrected: dict[str, list[str]]   = field(default_factory=dict)  # code -> [osd_sn, ...]


# ============================================================
# 1. Load IBD Origin
# ============================================================

def load_ibd(week: str) -> pd.DataFrame:
    wk = week.replace('W', 'WK')
    matches = list(IBD_DIR.glob(f"*IBD Origin*{wk}*"))
    if not matches:
        raise FileNotFoundError(f"IBD Origin file not found for {wk}")
    path = matches[0]
    df = pd.read_excel(path, engine='pyxlsb') if path.suffix == '.xlsb' else pd.read_excel(path)
    logger.info(f"IBD loaded: {len(df)} rows from {path.name}")
    return df


# ============================================================
# 2. Extract Damage (D) rows from OSD reports
# ============================================================

def load_osd_d_rows(week: str, result: ReportResult) -> dict[str, pd.DataFrame]:
    osd_by_plant = {}
    for code in PLANT_CODES:
        matches = list(OSD_DIR.glob(f"*{week}*{code}*"))
        if not matches:
            logger.warning(f"OSD file not found: {code} {week}")
            result.file_missing.append(code)
            continue

        frames = []
        fail_reason = None
        for path in matches:
            df, reason = _read_osd_file(path)
            if df is None:
                fail_reason = reason
                logger.warning(reason)
                continue
            d_rows = df[df['O/S/D'].astype(str).str.strip().str.upper().isin(['D', 'DAMAGE'])].copy()
            frames.append(d_rows)

        if not frames:
            result.fail.append((code, fail_reason or 'Unknown error'))
            continue

        combined = pd.concat(frames, ignore_index=True)
        logger.info(f"{code}: {len(combined)} D rows from {len(frames)} file(s)")
        osd_by_plant[code] = combined

    return osd_by_plant


def _read_osd_file(path: Path) -> tuple[pd.DataFrame | None, str | None]:
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        sheet_name = next((s for s in wb.sheetnames if 'summary' in s.strip().lower()), None)
        wb.close()
        if sheet_name is None:
            return None, f"No 'Summary' sheet: {path.name}"

        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=10)
        header_row = next((i for i, row in raw.iterrows() if 'O/S/D' in row.values), None)
        if header_row is None:
            return None, f"'O/S/D' header not found: {path.name}"

        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    except Exception as e:
        return None, f"Failed to read {path.name}: {e}"

    # Normalize columns to the fixed OSD_COLS schema regardless of case/variant names
    col_lower = {str(c).strip().lower(): c for c in df.columns}
    rename_map = {}
    for target in OSD_COLS:
        tl = target.lower()
        if tl in col_lower:
            rename_map[col_lower[tl]] = target
        elif target == 'Serial Number(s)':
            found = next((c for cl, c in col_lower.items() if 'serial number' in cl), None)
            if found:
                rename_map[found] = target
        elif target == 'Qty':
            found = next((c for cl, c in col_lower.items() if cl == 'damage qty'), None)
            if found:
                rename_map[found] = target

    df = df.rename(columns=rename_map)
    df = df[[c for c in OSD_COLS if c in df.columns]]
    return df, None


# ============================================================
# 3. IBD vs OSD matching -> missing items
# ============================================================

def find_missing(ibd: pd.DataFrame, osd: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Consume OSD D rows with each IBD row. IBD rows left over are missing.
    OSD rows left over become new OVG (overage).
    Returns (missing_df, remaining_osd, corrected_sns).
    """
    if ibd.empty:
        return pd.DataFrame(), osd, []

    sn_col_ibd  = _col_serial(ibd)
    mat_col_ibd = _col_material(ibd)
    qty_col_ibd = _col_qty_ibd(ibd)
    sn_col_osd  = _col_serial(osd)
    mod_col_osd = _col_model(osd)
    qty_col_osd = _col_qty_osd(osd)

    missing_rows = []
    corrected_sns = []
    for _, row in ibd.iterrows():
        sn       = _str(row[sn_col_ibd]) if sn_col_ibd else ''
        material = _str(row[mat_col_ibd]) if mat_col_ibd else ''
        qty      = _to_int(row[qty_col_ibd]) if qty_col_ibd else 1

        remaining, osd, corrected = _match_and_consume(sn, material, qty, osd, sn_col_osd, mod_col_osd, qty_col_osd)
        corrected_sns.extend(corrected)
        if remaining > 0:
            r = row.copy()
            if qty_col_ibd:
                r[qty_col_ibd] = remaining
            missing_rows.append(r)

    if not missing_rows:
        return pd.DataFrame(), osd, corrected_sns

    df      = pd.DataFrame(missing_rows)
    col_map = {str(c).strip().lower(): c for c in df.columns}
    qty_col = col_map.get('quantity')
    grp     = [col_map[c.lower()] for c in IBD_GROUP_COLS if c.lower() in col_map]
    inbd    = col_map.get('inb d/o')

    if qty_col and grp:
        agg = {qty_col: 'sum'}
        if inbd:
            agg[inbd] = 'first'
        df = df.groupby(grp, dropna=False).agg(agg).reset_index()
        rn = {qty_col: 'Total'}
        if inbd:
            rn[inbd] = 'SN'
        df = df.rename(columns=rn)
    else:
        df = df[grp]

    ordered = [c for c in ['WH N.', 'Week', 'CTN #', 'SN', 'Material', 'Total', 'QC G.', 'Serial number'] if c in df.columns]
    return df[ordered], osd, corrected_sns


# ============================================================
# 4. OVG check: deduct items found in the OVG from two weeks ago
# ============================================================

def _ovg_path(week: str) -> Path:
    return cfg.OVG_DIR / f'OSD_OVG_{week}.xlsx'


def load_ovg(plant: str, week: str) -> pd.DataFrame:
    path = _ovg_path(week)
    if not path.exists():
        return pd.DataFrame(columns=OSD_COLS)
    df = pd.read_excel(path)
    col = next((c for c in df.columns if c.strip().lower() == 'plant'), None)
    if not col:
        return pd.DataFrame(columns=OSD_COLS)
    return df[df[col].astype(str).str.strip() == plant].reset_index(drop=True)


def check_ovg(missing: pd.DataFrame, ovg: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Deduct OVG-matched items from the missing list. Returns (still_missing, remaining_ovg)."""
    if missing.empty or ovg.empty:
        return missing, ovg

    sn_col_miss  = _col_serial(missing)
    mat_col_miss = _col_material(missing)
    qty_col_miss = next((c for c in missing.columns if str(c).strip().lower() == 'total'), None)
    sn_col_ovg   = _col_serial(ovg)
    mod_col_ovg  = _col_model(ovg)
    qty_col_ovg  = _col_qty_osd(ovg)

    still_missing = []
    for _, row in missing.iterrows():
        sn       = _str(row[sn_col_miss]) if sn_col_miss else ''
        material = _str(row[mat_col_miss]) if mat_col_miss else ''
        qty      = _to_int(row[qty_col_miss]) if qty_col_miss else 1

        remaining, ovg, _ = _match_and_consume(sn, material, qty, ovg, sn_col_ovg, mod_col_ovg, qty_col_ovg)
        if remaining > 0:
            r = row.copy()
            if qty_col_miss:
                r[qty_col_miss] = remaining
            still_missing.append(r)

    final = pd.DataFrame(still_missing) if still_missing else pd.DataFrame(columns=missing.columns)
    return final, ovg


def save_ovg(remaining_ovg: pd.DataFrame, ovg_new: pd.DataFrame, plant: str, week: str) -> None:
    """Write this week's OVG file: unmatched old OVG + new OVG rows, older than KEEP_OVG_WEEKS dropped."""
    new_rows     = ovg_new[OSD_COLS].copy() if not ovg_new.empty else pd.DataFrame(columns=OSD_COLS)
    plant_update = pd.concat([remaining_ovg, new_rows], ignore_index=True)

    cfg.OVG_DIR.mkdir(parents=True, exist_ok=True)
    path = _ovg_path(week)

    if path.exists():
        existing  = pd.read_excel(path)
        plant_col = next((c for c in existing.columns if c.strip().lower() == 'plant'), None)
        other     = existing[existing[plant_col].astype(str).str.strip() != plant] if plant_col else existing
        combined  = pd.concat([other, plant_update], ignore_index=True)
    else:
        combined = plant_update

    if not combined.empty and 'Week' in combined.columns:
        wk_num   = int(week.replace('W', ''))
        combined = combined[combined['Week'].apply(
            lambda w: pd.notna(w) and (wk_num - int(w)) % 52 <= KEEP_OVG_WEEKS
        )]

    combined.to_excel(path, index=False)
    logger.info(f"OVG saved: {len(new_rows)} new, {len(remaining_ovg)} remaining for {plant} {week}")


# ============================================================
# 5. Master resolve: drop master items whose serial shows up in this week's OSD
# ============================================================

def load_master(week: str) -> pd.DataFrame:
    path = MASTER_DIR / f'{week}_missing_master.xlsx'
    if not path.exists():
        logger.warning(f"Master file not found: {path.name}")
        return pd.DataFrame(columns=MASTER_COLS)
    return pd.read_excel(path, header=1)


def resolve_master(master: pd.DataFrame, osd: pd.DataFrame, plant: str) -> pd.DataFrame:
    if master.empty or osd.empty or 'Plant' not in master.columns:
        return master

    plant_rows = master[master['Plant'] == plant]
    other_rows = master[master['Plant'] != plant]
    if plant_rows.empty:
        return master

    sn_col      = _col_serial(osd)
    osd_serials = set(osd[sn_col].astype(str).str.strip()) if sn_col else set()

    def not_resolved(s):
        s = str(s).strip()
        return s.lower() in ('nan', '') or s not in osd_serials

    remaining    = plant_rows[plant_rows['Serial number'].apply(not_resolved)]
    resolved_cnt = len(plant_rows) - len(remaining)
    logger.info(f"{plant}: {resolved_cnt} master items resolved by OSD")
    return pd.concat([other_rows, remaining], ignore_index=True)


# ============================================================
# 6. Save missing report + accumulate master
# ============================================================

def save_missing(df: pd.DataFrame, code: str, week: str) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HEADER_FONT = Font(name='Calibri', bold=True, size=10)
    ROW_FONT    = Font(name='Calibri', size=10)
    THIN        = Side(style='thin')
    BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    COL_WIDTHS  = {'WH N.': 8, 'Week': 7, 'CTN #': 16, 'SN': 14,
                   'Material': 22, 'Total': 7, 'QC G.': 8, 'Serial number': 18}

    week_dir = MISSING_DIR / week
    week_dir.mkdir(parents=True, exist_ok=True)
    path = week_dir / MISSING_FILENAME.format(week=week, code=code)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    cols = df.columns.tolist()

    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for ri, (_, row_data) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(cols, 1):
            val = row_data[col]
            if pd.isna(val):
                val = ''
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = ROW_FONT
            cell.border = BORDER

    for ci, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col, 14)
    ws.freeze_panes = 'A2'

    wb.save(path)
    logger.info(f"Missing saved: {path.name}")


def add_to_master(master: pd.DataFrame, new_missing: pd.DataFrame, plant: str) -> pd.DataFrame:
    if new_missing.empty:
        return master
    rows = new_missing.copy()
    rows['Plant'] = plant
    for col in MASTER_COLS:
        if col not in rows.columns:
            rows[col] = ''
    rows = rows[MASTER_COLS]
    logger.info(f"{plant}: {len(rows)} items added to master")
    return pd.concat([master, rows], ignore_index=True)


def clean_master(master: pd.DataFrame, week: str) -> pd.DataFrame:
    if master.empty or 'Week' not in master.columns:
        return master
    wk_num = int(week.replace('W', ''))

    def is_recent(w):
        try:
            return (wk_num - int(w)) % 52 <= KEEP_MASTER_WEEKS
        except Exception:
            return True

    cleaned = master[master['Week'].apply(is_recent)].reset_index(drop=True)
    removed = len(master) - len(cleaned)
    if removed:
        logger.info(f"Master cleanup: {removed} expired items (>{KEEP_MASTER_WEEKS} weeks)")
    return cleaned


def save_master(df: pd.DataFrame, week: str) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    HEADER_FONT = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    ROW_FONT    = Font(name='Calibri', size=10)
    ALT_FILL    = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    BORDER      = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    COL_WIDTHS  = {'Plant': 8, 'WH N.': 8, 'Week': 7, 'CTN #': 16, 'SN': 14,
                   'Material': 22, 'Total': 7, 'QC G.': 8, 'Serial number': 18}

    MASTER_DIR.mkdir(parents=True, exist_ok=True)
    out_path = MASTER_DIR / f'{week}_missing_master.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Master Missing Report'
    cols = df.columns.tolist()

    ws.merge_cells(f'A1:{get_column_letter(len(cols))}1')
    title           = ws['A1']
    title.value     = 'Missing Report Master'
    title.font      = Font(name='Calibri', bold=True, size=13, color='1F4E79')
    title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22

    for ci, col in enumerate(cols, 1):
        cell           = ws.cell(row=2, column=ci, value=col)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = BORDER
    ws.row_dimensions[2].height = 18

    for ri, (_, row_data) in enumerate(df.iterrows(), 3):
        fill = ALT_FILL if ri % 2 == 0 else None
        for ci, col in enumerate(cols, 1):
            val = row_data[col]
            if pd.isna(val):
                val = ''
            cell           = ws.cell(row=ri, column=ci, value=val)
            cell.font      = ROW_FONT
            cell.border    = BORDER
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if fill:
                cell.fill = fill

    for ci, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col, 14)
    ws.freeze_panes = 'A3'

    wb.save(out_path)
    logger.info(f"Master saved: {len(df)} total items -> {out_path.name}")


def _process_plant(code: str, osd: pd.DataFrame, ibd: pd.DataFrame,
                   wh_col: str, ovg_week: str, week: str,
                   master: pd.DataFrame, result: ReportResult) -> pd.DataFrame:

    wh        = next((k for k, v in WH_TO_PLANT.items() if v == code), None)
    ibd_plant = ibd[ibd[wh_col].astype(str).str.strip() == wh] if wh_col and wh else ibd

    # QC Grade A / LS items are not damage candidates
    qc_col = next((c for c in ibd_plant.columns if 'qc g' in str(c).strip().lower()), None)
    if qc_col:
        ibd_plant = ibd_plant[~ibd_plant[qc_col].astype(str).str.strip().str.upper().isin(['A', 'LS'])]

    # 3. IBD vs OSD matching
    missing, ovg_new, corrected = find_missing(ibd_plant, osd.copy())
    if corrected:
        result.sn_corrected[code] = corrected
        logger.info(f"{code}: {len(corrected)} SN(s) matched after O->0 correction: {corrected}")
    logger.info(f"{code}: {len(missing)} missing candidates, {len(ovg_new)} OVG new")

    # 4. OVG check
    existing_ovg           = load_ovg(code, ovg_week)
    missing, remaining_ovg = check_ovg(missing, existing_ovg)
    logger.info(f"{code}: {len(missing)} final missing after OVG check")

    # 5. Master resolve
    master = resolve_master(master, osd, code)

    # 6. Save missing + accumulate master
    if not missing.empty:
        save_missing(missing, code, week)
        master = add_to_master(master, missing, code)
        result.success.append((code, True))
    else:
        logger.info(f"{code}: no missing items")
        result.success.append((code, False))

    save_ovg(remaining_ovg, ovg_new, code, week)

    return master


# ============================================================
# Helpers
# ============================================================

def _str(val) -> str:
    s = str(val).strip()
    return '' if s.lower() == 'nan' else s

def _to_int(val, default=1) -> int:
    try:
        return int(val)
    except (ValueError, TypeError):
        return default

def _is_serial(val: str) -> bool:
    return bool(re.match(r'^[A-Za-z0-9]{14,15}$', val))

def _col_serial(df):
    return next((c for c in df.columns if 'serial' in str(c).lower()), None)

def _col_material(df):
    return next((c for c in df.columns if 'material' in str(c).lower()), None)

def _col_model(df):
    return next((c for c in df.columns if 'model' in str(c).lower()), None)

def _col_qty_osd(df):
    return next((c for c in df.columns if str(c).strip().lower() in ('qty', 'damage qty')), None)

def _col_qty_ibd(df):
    return next((c for c in df.columns if 'quantity' in str(c).lower()), None)


def _normalize_sn(sn) -> str:
    """Turn a leading letter O into digit 0 (common OSD data-entry error)."""
    if not isinstance(sn, str) or not sn:
        return str(sn)
    if sn[0].upper() == 'O':
        return '0' + sn[1:]
    return sn


def _match_and_consume(sn: str, material: str, src_qty: int,
                       osd: pd.DataFrame,
                       sn_col, mod_col, qty_col) -> tuple[int, pd.DataFrame, list[str]]:
    """
    Match one source row against the OSD work set and consume matched rows.
    Valid serial -> 1:1 match (retried with O->0 correction). No serial -> Model + Qty deduction.
    Returns (remaining_qty, updated_osd, corrected_sns).
    """
    has_serial = sn and _is_serial(sn)

    if has_serial:
        if sn_col:
            osd_serials = osd[sn_col].astype(str).str.strip()
            idx = osd[osd_serials == sn].index.tolist()
            if idx:
                osd = osd.drop(index=idx).reset_index(drop=True)
                return 0, osd, []
            # Retry after correcting a leading O -> 0 in the OSD serial
            normalized = osd_serials.apply(_normalize_sn)
            idx = osd[normalized == sn].index.tolist()
            if idx:
                corrected = osd_serials.iloc[idx].tolist()
                osd = osd.drop(index=idx).reset_index(drop=True)
                return 0, osd, corrected
        return src_qty, osd, []

    if material and mod_col and qty_col:
        idx = osd[osd[mod_col].astype(str).str.strip() == material].index.tolist()
        if not idx:
            return src_qty, osd, []
        osd_qty = sum(_to_int(osd.loc[i, qty_col]) for i in idx)
        osd     = osd.drop(index=idx).reset_index(drop=True)
        return max(src_qty - osd_qty, 0), osd, []

    return 0, osd, []


# ============================================================
# Run
# ============================================================

def run() -> tuple[ReportResult, str]:
    from datetime import datetime
    import pytz
    _tz = pytz.timezone('America/Chicago')
    is_friday = datetime.now(_tz).weekday() == 4

    result = ReportResult()
    # Friday runs cover the current week; Mon-Thu runs cover the previous week
    week   = current_week() if is_friday else previous_week()
    logger.info(f"Target week: {week}")

    # 1. Load IBD
    try:
        ibd = load_ibd(week)
    except FileNotFoundError:
        logger.error("IBD Origin file not found")
        result.file_missing.append('IBD')
        return result, week

    # 2. Extract OSD D rows
    osd_by_plant = load_osd_d_rows(week, result)
    if not osd_by_plant:
        logger.error("No OSD files loaded, exiting")
        return result, week

    wh_col   = next((c for c in ibd.columns if str(c).strip().lower() == 'wh n.'), None)
    ovg_week = two_weeks_ago()
    master   = load_master(two_weeks_ago())

    for code, osd in osd_by_plant.items():
        try:
            master = _process_plant(code, osd, ibd, wh_col, ovg_week, week, master, result)
        except Exception:
            logger.exception(f"{code}: unexpected error")
            result.fail.append((code, 'Unexpected error'))

    master = clean_master(master, week)
    save_master(master, week)

    accounted = {c for c, _ in result.success} | set(result.file_missing) | {c for c, _ in result.fail}
    result.fail += [(c, 'Not processed') for c in PLANT_CODES if c not in accounted]

    logger.info(f"Success: {[c for c, _ in result.success]}")
    logger.info(f"File Missing: {result.file_missing}")
    logger.info(f"Fail: {result.fail}")
    return result, week


if __name__ == '__main__':
    from email_pipeline.sender import load_sender, send_status_report, send_plant_report, send_main_failure

    sender = load_sender()
    week   = None
    try:
        result, week = run()

        if sender:
            send_status_report(sender, result, week)
            for code, has_report in result.success:
                if has_report:
                    report_path = MISSING_DIR / week / MISSING_FILENAME.format(week=week, code=code)
                    send_plant_report(sender, code, report_path, week)

    except Exception:
        logger.exception("Unexpected error")
        if sender:
            send_main_failure(sender, _log_file, week or '')
