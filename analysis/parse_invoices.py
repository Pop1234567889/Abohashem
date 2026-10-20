# -*- coding: utf-8 -*-
"""Parse the Abu Hashem sales-invoices PDF into structured records (coordinate based)."""
import fitz, re, json, os

PDF = os.path.join(os.path.dirname(__file__), '..',
                   'فواتير مبيعات من 1-6 -2026  الى فواتير بيع 26-6-2027.pdf')

# column x-center ranges derived from the table header
COLS = {
    'line_total':  (25, 62),
    'discount':    (63, 92),
    'tax':         (93, 122),
    'unit_price':  (123, 175),
    'qty':         (185, 275),
    'unit':        (280, 345),
    'name':        (346, 480),
    'item_code':   (485, 525),
    'rownum':      (545, 570),
}

def col_of(x):
    for c, (a, b) in COLS.items():
        if a <= x <= b:
            return c
    return None

def num(s):
    try:
        return float(s.replace(',', ''))
    except ValueError:
        return None

def parse():
    doc = fitz.open(PDF)
    invoices = []
    for page in doc:
        t = page.get_text()
        inv = {'page': page.number + 1}
        m = re.search(r'([bB]\d+)فاتورة مبيعات رقم', t)
        inv['invoice_no'] = m.group(1) if m else None
        m = re.search(r'/ الكود(\d{4}/\d{1,2}/\d{1,2})', t)
        inv['date'] = m.group(1) if m else None
        lines = t.split('\n')
        name = code = None
        for i, ln in enumerate(lines):
            if ln.strip() == 'اسم العميل /':
                name = lines[i-1].strip(); code = lines[i+1].strip(); break
        inv['customer'] = name
        inv['customer_code'] = code
        m = re.search(r'([\d.]+)عدد كميات الفاتورة', t)
        inv['total_qty'] = num(m.group(1)) if m else None
        m = re.search(r'([\d.]+) / إجمالى الفاتورة', t)
        inv['invoice_total'] = num(m.group(1)) if m else None
        m = re.search(r'/ المدفوع([\d.]+)', t)
        inv['paid'] = num(m.group(1)) if m else None
        m = re.search(r'/ الباقي([\d.]+)', t)
        inv['remaining'] = num(m.group(1)) if m else None

        words = page.get_text('words')  # x0,y0,x1,y1,text,block,line,word
        # header bottom: lowest y among header label words
        hdr_y = [w[3] for w in words if w[4] in ('الكمية', 'الصنف', 'الكود', 'الجمالي')]
        header_bottom = max(hdr_y) if hdr_y else 250
        foot_y = [w[1] for w in words if w[4] == 'كميات']
        footer_top = min(foot_y) if foot_y else 1e9
        data = [w for w in words if header_bottom + 1 < w[1] < footer_top - 1]
        # anchor each row on its row-number word (one per item), then assign
        # every data word to the nearest anchor by vertical position.
        anchors = sorted((w[1] + w[3]) / 2 for w in data
                         if COLS['rownum'][0] <= (w[0] + w[2]) / 2 <= COLS['rownum'][1])
        if not anchors:
            inv['items'] = []
            invoices.append(inv)
            continue
        rows = {a: [] for a in anchors}
        for w in data:
            yc = (w[1] + w[3]) / 2
            a = min(anchors, key=lambda an: abs(an - yc))
            rows[a].append(w)
        items = []
        for a in anchors:
            cells = {}
            for w in rows[a]:
                xc = (w[0] + w[2]) / 2
                c = col_of(xc)
                if not c:
                    continue
                cells.setdefault(c, []).append((w[0], w[4]))
            if 'rownum' not in cells and 'line_total' not in cells:
                continue
            def val(c):
                if c not in cells:
                    return None
                parts = [p[1] for p in sorted(cells[c])]
                return ' '.join(parts)
            lt = num(val('line_total') or '')
            qty = num(val('qty') or '')
            up = num(val('unit_price') or '')
            ic = val('item_code')
            if lt is None and qty is None and up is None:
                continue
            nm = val('name')
            if nm:  # reverse word order for natural RTL reading
                nm = ' '.join(reversed(nm.split(' ')))
            items.append({
                'item_code': ic.strip() if ic else None,
                'name': nm,
                'unit': val('unit'),
                'unit_price': up,
                'qty': qty,
                'line_total': lt,
                'discount_pct': num(val('discount') or '0') or 0.0,
                'tax_pct': num(val('tax') or '0') or 0.0,
            })
        inv['items'] = items
        invoices.append(inv)
    return invoices

if __name__ == '__main__':
    invs = parse()
    bad = 0; tot = 0; noit = 0; mism = 0
    for inv in invs:
        if not inv['items']:
            noit += 1
        for it in inv['items']:
            tot += 1
            if None in (it['unit_price'], it['qty'], it['line_total']):
                bad += 1; continue
            exp = it['unit_price'] * it['qty'] * (1 - it['discount_pct']/100)
            if abs(exp - it['line_total']) > 0.6:
                bad += 1
        if inv['items'] and inv['invoice_total'] is not None:
            s = sum((it['line_total'] or 0) for it in inv['items'])
            if abs(s - inv['invoice_total']) > 1.0:
                mism += 1
    print('invoices:', len(invs))
    print('invoices with no parsed items:', noit)
    print('line items:', tot)
    print('line items failing price*qty*(1-disc)=total:', bad)
    print('invoices where sum(lines)!=invoice_total:', mism)
    json.dump(invs, open(os.path.join(os.path.dirname(__file__), 'invoices.json'), 'w'),
              ensure_ascii=False, indent=1)
    print('wrote invoices.json')
