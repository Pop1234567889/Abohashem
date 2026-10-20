# -*- coding: utf-8 -*-
"""Accounts-receivable AGING report per customer.
Unpaid invoices are bucketed by days outstanding (invoice date -> reference
date) into 0-30 / 31-60 / 61-90 / 90+. Each customer's buckets are reconciled
to the standing debt balance from the reports; any balance not explained by
dated period invoices is shown as an older 'opening/carryover (undated)' bucket."""
import json, os, csv
from collections import defaultdict
from datetime import date

HERE = os.path.dirname(__file__)
OUT = os.path.join(HERE, 'output')
invoices = json.load(open(os.path.join(HERE, 'invoices.json'), encoding='utf-8'))
custrows = json.load(open(os.path.join(HERE, 'salespeople.json'), encoding='utf-8'))

FILE_TO_SP = {
    'ايمن فارس 26-6.pdf': 'ايمن فارس', 'حسام حسن 26-6.pdf': 'حسام حسن',
    'شعبان 26-6.pdf': 'شعبان', 'محمد امام 26-.pdf': 'محمد امام',
    'محمد امام الصعيد 26-6.pdf': 'محمد امام الصعيد', 'محمد بشرى 26-6.pdf': 'محمد بشرى',
    'محمد خليل 26-6.pdf': 'محمد خليل', 'محمود السيد.pdf': 'محمود السيد', 'هانى 26-6.pdf': 'هانى',
}
def norm(c): return (c or '').replace(',', '').replace(' ', '').strip()
def to_date(s):
    y, m, d = map(int, s.split('/')); return date(y, m, d)

REF = date(2026, 6, 27)   # reference (today)

# report debt + salesperson per customer
debt = {}; sp_of = {}; name_of = {}
for c in custrows:
    code = norm(c['code'])
    debt[code] = c['debit'] - c['credit']
    sp_of[code] = FILE_TO_SP.get(c['report_file'], c['salesperson'])
    name_of[code] = c['customer']

# unpaid invoice amounts bucketed by age, per customer
B = lambda: [0.0, 0.0, 0.0, 0.0]   # 0-30, 31-60, 61-90, 90+
buck = defaultdict(B); inv_name = {}
for inv in invoices:
    amt = inv['remaining'] or 0.0
    if amt <= 0:
        continue
    code = norm(inv['customer_code']); inv_name[code] = inv['customer']
    age = (REF - to_date(inv['date'])).days
    i = 0 if age <= 30 else 1 if age <= 60 else 2 if age <= 90 else 3
    buck[code][i] += amt

rows = []
codes = set(debt) | set(buck)
for code in codes:
    d = debt.get(code)                      # report balance (may be None if not in report)
    b = list(buck.get(code, B()))
    dated = sum(b)
    in_report = code in debt
    eff = d if in_report else dated         # the receivable to reconcile to
    opening = 0.0
    if eff is None:
        eff = dated
    if eff >= dated:
        opening = round(eff - dated, 2)     # older balance not covered by dated invoices
    elif dated > 0:                         # invoices exceed balance -> scale down
        f = eff / dated
        b = [x * f for x in b]
        opening = 0.0
    total = sum(b) + opening
    if total <= 0.005:
        continue
    rows.append({
        'customer_code': code,
        'customer': name_of.get(code) or inv_name.get(code, ''),
        'salesperson': sp_of.get(code, '(غير مرتبط/unmapped)'),
        'days_0_30_EGP': round(b[0], 2),
        'days_31_60_EGP': round(b[1], 2),
        'days_61_90_EGP': round(b[2], 2),
        'days_over_90_EGP': round(b[3], 2),
        'opening_carryover_undated_EGP': opening,
        'total_debt_EGP': round(total, 2),
    })
rows.sort(key=lambda r: r['total_debt_EGP'], reverse=True)

# totals row
tot = {k: round(sum(r[k] for r in rows), 2) for k in
       ('days_0_30_EGP', 'days_31_60_EGP', 'days_61_90_EGP', 'days_over_90_EGP',
        'opening_carryover_undated_EGP', 'total_debt_EGP')}
totrow = {'customer_code': '', 'customer': 'الإجمالى / TOTAL', 'salesperson': '', **tot}
rows.append(totrow)

path = os.path.join(OUT, 'aging_by_customer.csv')
with open(path, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

print(f'reference date: {REF}')
print(f'customers: {len(rows)-1}')
for k, v in tot.items():
    print(f'  {k:32s} {v:14,.2f}')
print('wrote', path)
