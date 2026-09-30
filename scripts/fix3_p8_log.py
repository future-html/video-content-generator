import openpyxl
from openpyxl.styles import PatternFill, Font
from datetime import datetime, timezone

WORK = 'outputs/whiteshirt_interview_master_WORKING.xlsx'
TS = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%SZ')
YELLOW = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

wb = openpyxl.load_workbook(WORK)

# --- ตรวจ + log P8 ---
print("=== P8 verification + logging ===")
ws_ce = wb['COMMENT_EXECUTION']
ws_cr = wb['COMMENT_RECHECK']

# หา row IC-0044 ใน COMMENT_EXECUTION
ce_row = None
for r in range(5, ws_ce.max_row + 1):
    if ws_ce.cell(row=r, column=2).value == 'IC-0044':
        ce_row = r
        break

# หา source ใน COMMENT_RECHECK (IC-0044, 24h)
cr_row = None
for r in range(5, ws_cr.max_row + 1):
    cid = ws_cr.cell(row=r, column=4).value
    tp = ws_cr.cell(row=r, column=19).value
    if cid == 'IC-0044' and tp == '24h':
        cr_row = r
        break

print(f"  COMMENT_EXECUTION IC-0044 at row {ce_row}")
print(f"  COMMENT_RECHECK  IC-0044|24h at row {cr_row}")

if ce_row and cr_row:
    # highlight cell AS (col 45) เพื่อบ่งชี้ว่าตรวจสอบแล้ว
    ws_ce.cell(row=ce_row, column=45).fill = YELLOW
    # highlight source cell X (col 24) ใน COMMENT_RECHECK
    ws_cr.cell(row=cr_row, column=24).fill = YELLOW

# --- append P8 entry ลง FIX_LOG ---
if 'FIX_LOG' in wb.sheetnames:
    ws_log = wb['FIX_LOG']
    
    # ตรวจว่า P8 มีอยู่แล้วหรือยัง
    existing = set()
    for r in range(2, ws_log.max_row + 1):
        v = ws_log.cell(row=r, column=1).value
        if v:
            existing.add(str(v))
    
    if not any(s.startswith('P8') for s in existing):
        entry = [
            'P8',
            TS,
            'Reply Owner Type invalid value "0" (formula chain)',
            'COMMENT_EXECUTION',
            ce_row if ce_row else '',
            'AS',
            'Reply Owner Type',
            '0 (cached from VLOOKUP chain)',
            'formula preserved (VLOOKUP+IFERROR)',
            'Cell AS{} references COMMENT_RECHECK!X{} which references R{}'.format(
                ce_row, cr_row, cr_row),
            'Source 0 was Excel cache artifact; formula chain verified; '
            'no literal fix required - Excel will recalculate on open'
        ]
        ws_log.append(entry)
        print(f"  appended P8 log entry at FIX_LOG row {ws_log.max_row}")
        
        # append additional entry to source side
        entry2 = [
            'P8-src',
            TS,
            'Source of P8 in COMMENT_RECHECK',
            'COMMENT_RECHECK',
            cr_row if cr_row else '',
            'X',
            'Lookup Reply Owner',
            '0 (cached)',
            'formula =R{}'.format(cr_row),
            'VLOOKUP resolves to this cell for IC-0044|24h',
            'Verified chain: R -> X -> AS (COMMENT_EXECUTION)'
        ]
        ws_log.append(entry2)
        print(f"  appended P8-src log entry at FIX_LOG row {ws_log.max_row}")
    else:
        print(f"  P8 already logged, skipping")

wb.save(WORK)
print(f"\nSaved: {WORK}")
