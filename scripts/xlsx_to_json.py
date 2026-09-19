#!/usr/bin/env python3
"""Convert the PES field-mapping workbook into data.json for the static viewer.
Usage: python xlsx_to_json.py <workbook.xlsx> <out.json>
"""
import sys, json
from openpyxl import load_workbook

MODS = ['Application','Contract','Monitoring Visit','Remediation','Payment','Beneficiary','Project & Activity']
ROLES = ['Agg FO','Agg M&E','Agg Manger','Agg Data Admin','Agg Super admin',
         'Impl FO','Impl M&E','Impl Manger','Impl Data Admin','Impl Super admin']
# column order on each module sheet (1-indexed): matches build3.py COLS
# 1 #,2 Field,3 Section,4 Sub,5 Activity,6 Status,7 Chain,8-17 roles,18 kf,19 kq,20 hide,21 cov

def clean(v):
    return '' if v is None else str(v).strip()

def main(src, out):
    wb = load_workbook(src, data_only=True)
    rows = []
    for m in MODS:
        name = m[:31]
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or not clean(r[1]):
                continue
            roles = [clean(r[7+i]) for i in range(10)]
            rows.append({
                'mod': m, 'f': clean(r[1]), 'sec': clean(r[2]), 'sub': clean(r[3]),
                'act': clean(r[4]) or 'Any', 'st': clean(r[5]) or 'All', 'ch': clean(r[6]) or 'both',
                'roles': roles, 'kf': clean(r[17]), 'kq': clean(r[18]), 'hide': clean(r[19]),
                'cov': clean(r[20]),
            })
    files = []
    if 'Files' in wb.sheetnames:
        ws = wb['Files']
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or not clean(r[2]):
                continue
            files.append({'mod': clean(r[1]), 'name': clean(r[2]), 'act': clean(r[3]),
                           'req': clean(r[4]), 'multi': clean(r[5]), 'key': clean(r[6])})
    json.dump({'roles': ROLES, 'modules': MODS, 'rows': rows, 'files': files},
               open(out, 'w'), ensure_ascii=False)
    print(f'wrote {out}: {len(rows)} rows, {len(files)} files')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
