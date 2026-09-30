import pandas as pd
import os

f = 'whiteshirt_interview_master_SANITIZED.xlsx'
assert os.path.exists(f), f"ไม่เจอ: {f} (cwd={os.getcwd()})"

def load(sheet, hr=3):
    df = pd.read_excel(f, sheet_name=sheet, header=hr)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how='all')

da = load('DAILY_ACCOUNT')
db = load('DAILY_BRAND')
ce = load('COMMENT_EXECUTION')
cr = load('COMMENT_RECHECK')
ei = load('EVIDENCE_INDEX')
cc = load('COMPLETENESS_CHECK')

print("="*100); print("S1 DAILY_ACCOUNT"); print("="*100)
print(f"rows={len(da)}")
print(da['Date'].astype(str).value_counts().sort_index().to_string())

print("\n--- P1: rows 2026-07-23 ---")
d23 = da[da['Date'].astype(str).str.startswith('2026-07-23')]
print(d23[['Date','Comment Account','Account Group','Test Group',
           'Comment Count Today','Visible Comment Count Today']].to_string(index=False))
for i, r in d23.iterrows():
    print(f"  pandas idx={i} | Excel row={i+5} | {r['Comment Account']}")

print("\n" + "="*100); print("S2 DAILY_BRAND"); print("="*100)
print(db[['Date','Product Clicks','Add to Cart','Orders','GMV']].to_string(index=False))

print("\n" + "="*100); print("S3 COMMENT_EXECUTION"); print("="*100)
print(f"rows={len(ce)} | unique ID={ce['Comment ID'].nunique()}")
print("Dup IDs:", ce['Comment ID'][ce['Comment ID'].duplicated(keep=False)].tolist())
print("\nQC Flag:", ce['QC Exception Flag'].value_counts(dropna=False).to_dict())
print("Image Fit:", ce['Image Fit To Context'].value_counts(dropna=False).to_dict())
print("Reply Owner:", ce['Reply Owner Type'].value_counts(dropna=False).to_dict())
print("Cart Status:", ce['Cart Status'].value_counts(dropna=False).to_dict())

print("\n--- P2: IC-0012 rows ---")
for i, r in ce.iterrows():
    if r['Comment ID'] == 'IC-0012':
        print(f"  pandas idx={i} | Excel row={i+5} | {r['Comment Account']} | {str(r['Comment Text'])[:60]}")

used = set(ce['Comment ID'].dropna())
missing = [f"IC-{i:04d}" for i in range(1,101) if f"IC-{i:04d}" not in used]
print("\nMissing IDs in 1..100:", missing)

print("\n--- P4: Direct-sell keyword + QC flag ---")
kw = ['ซื้อเลย','ลิงก์','โปรแรง','ทักด่วน','ทักแชต','ลดราคา','รีบซื้อ','ส่งฟรี','คลิก']
ce['hit'] = ce['Comment Text'].astype(str).str.lower().apply(lambda x: [k for k in kw if k in x])
sel = ce[ce['hit'].apply(len) > 0][['Comment ID','Comment Account','Comment Text','QC Exception Flag','hit']]
print(sel.to_string(index=False))

print("\n--- P4b: QC Flag = Review ---")
print(ce[ce['QC Exception Flag']=='Review'][['Comment ID','Comment Account','Comment Text']].to_string(index=False))

print("\n" + "="*100); print("S4 COMMENT_RECHECK"); print("="*100)
print(f"rows={len(cr)}")
piv = cr.pivot_table(index='Comment ID', columns='Verification Timepoint',
                     values='Recheck Result', aggfunc='first')
piv['change'] = piv['30m'].astype(str) + " -> " + piv['24h'].astype(str)
print(piv['change'].value_counts().to_string())

print("\n" + "="*100); print("S5 EVIDENCE_INDEX"); print("="*100)
print(f"rows={len(ei)}")
print(ei['Capture Type'].value_counts().to_string())
print("Drive URL non-null:", ei['Drive URL'].notna().sum())

ei['linked_row'] = pd.to_numeric(ei['Linked Row'], errors='coerce')
def expected_row(r):
    if r['Linked Sheet'] == 'COMMENT_EXECUTION' and str(r['Entity Reference']).startswith('IC-'):
        return int(str(r['Entity Reference']).split('-')[1]) + 4
    return None
ei['expected'] = ei.apply(expected_row, axis=1)
bad = ei[ei['expected'].notna() & (ei['linked_row'] != ei['expected'])]
print("\n--- P11: EVIDENCE linked_row mismatch ---")
print(bad[['Evidence ID','Entity Reference','Linked Sheet','Linked Row','expected']].to_string(index=False))

print("\n" + "="*100); print("S6 COMPLETENESS_CHECK"); print("="*100)
print(cc.to_string(index=False))
