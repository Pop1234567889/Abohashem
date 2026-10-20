# -*- coding: utf-8 -*-
"""Parse the 9 per-salesperson customer-account reports."""
import fitz, glob, os, json, re

DIR = os.path.join(os.path.dirname(__file__), '..')
FILES = [f for f in glob.glob(os.path.join(DIR, '*.pdf'))
         if 'فواتير' not in os.path.basename(f)]

COLS = {
    'phone':       (15, 110),
    'area':        (110, 150),
    'debit':       (152, 208),
    'credit':      (220, 262),
    'salesperson': (282, 348),
    'ctype':       (360, 398),
    'custname':    (400, 492),
    'code':        (502, 548),
    'rownum':      (553, 572),
}
def col_of(x):
    for c,(a,b) in COLS.items():
        if a<=x<=b: return c
    return None
def num(s):
    try: return float(s.replace(',',''))
    except: return None

def parse_file(path):
    doc=fitz.open(path)
    custs=[]
    for page in doc:
        words=page.get_text('words')
        hdr=[w[3] for w in words if w[4] in ('المندوب','كود') and w[1]<170]
        header_bottom=max(hdr) if hdr else 18
        foot=[w[1] for w in words if w[4]=='اجمالى']
        footer_top=min(foot) if foot else 1e9
        data=[w for w in words if header_bottom+1<w[1]<footer_top-1]
        anchors=sorted((w[1]+w[3])/2 for w in data
                       if COLS['rownum'][0]<=(w[0]+w[2])/2<=COLS['rownum'][1])
        rows={a:[] for a in anchors}
        for w in data:
            yc=(w[1]+w[3])/2
            if not anchors: continue
            a=min(anchors,key=lambda an:abs(an-yc)); rows[a].append(w)
        for a in anchors:
            cells={}
            for w in rows[a]:
                c=col_of((w[0]+w[2])/2)
                if c: cells.setdefault(c,[]).append((w[0],w[4]))
            def val(c):
                if c not in cells: return None
                return ' '.join(p[1] for p in sorted(cells[c]))
            code=val('code')
            if not code: continue
            nm=val('custname'); nm=' '.join(reversed(nm.split(' '))) if nm else None
            sp=val('salesperson')
            custs.append({
                'code':code.strip(),
                'customer':nm,
                'salesperson':sp.strip() if sp else None,
                'debit':num(val('debit') or '0') or 0.0,
                'credit':num(val('credit') or '0') or 0.0,
                'phone':val('phone'),
                'area':val('area'),
                'report_file':os.path.basename(path),
            })
    return custs

if __name__=='__main__':
    allc=[]
    for f in sorted(FILES):
        c=parse_file(f)
        deb=sum(x['debit'] for x in c); cr=sum(x['credit'] for x in c)
        print(f'{os.path.basename(f):35s} customers={len(c):3d} debit={deb:14,.2f} credit={cr:10,.2f} net={deb-cr:14,.2f}')
        allc+=c
    # collision check
    seen={}
    for c in allc:
        seen.setdefault(c['code'],set()).add(c['salesperson'])
    coll={k:v for k,v in seen.items() if len(v)>1}
    print('codes mapped to >1 salesperson:',len(coll))
    print('total customer rows:',len(allc))
    json.dump(allc,open(os.path.join(os.path.dirname(__file__),'salespeople.json'),'w'),
              ensure_ascii=False,indent=1)
    print('wrote salespeople.json')
