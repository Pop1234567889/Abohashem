# -*- coding: utf-8 -*-
"""Amounts due per customer under a ONE-MONTH credit term (maturity = invoice
date + 1 month). Amount = the unpaid balance of each invoice (الباقي)."""
import json, os, csv
from collections import defaultdict
from datetime import date, timedelta

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
code_sp = {norm(c['code']): FILE_TO_SP.get(c['report_file'], c['salesperson']) for c in custrows}

CREDIT_TERM = timedelta(days=30)            # one-month maturity
TODAY = date(2026, 6, 27)                    # reference (today)
WINDOW_END = TODAY + CREDIT_TERM            # next one month from today

def to_date(s):
    y, m, d = map(int, s.split('/')); return date(y, m, d)
def add_month(d):
    m = d.month + 1; y = d.year + (m - 1) // 12; m = (m - 1) % 12 + 1
    import calendar
    day = min(d.day, calendar.monthrange(y, m)[1])
    return date(y, m, day)

# per-customer aggregation + per-invoice detail
agg = defaultdict(lambda: {'name': '', 'sp': '', 'n': 0, 'amount': 0.0,
                           'due_within': 0.0, 'overdue': 0.0, 'later': 0.0,
                           'first_due': None, 'last_due': None})
detail = []
for inv in invoices:
    amt = inv['remaining'] or 0.0          # unpaid balance of this invoice
    if amt <= 0:
        continue
    code = norm(inv['customer_code'])
    inv_d = to_date(inv['date'])
    due = add_month(inv_d)                  # maturity date = invoice date + 1 month
    a = agg[code]
    a['name'] = inv['customer']; a['sp'] = code_sp.get(code, '(غير مرتبط/unmapped)')
    a['n'] += 1; a['amount'] += amt
    if due < TODAY:
        a['overdue'] += amt
    elif due <= WINDOW_END:
        a['due_within'] += amt
    else:
        a['later'] += amt
    a['first_due'] = due if not a['first_due'] else min(a['first_due'], due)
    a['last_due'] = due if not a['last_due'] else max(a['last_due'], due)
    detail.append({
        'invoice_no': inv['invoice_no'], 'invoice_date': inv['date'],
        'due_date': str(due), 'customer_code': code, 'customer': inv['customer'],
        'salesperson': code_sp.get(code, '(غير مرتبط/unmapped)'),
        'amount_due_EGP': round(amt, 2),
    })

rows = []
for code, a in agg.items():
    rows.append({
        'customer_code': code, 'customer': a['name'], 'salesperson': a['sp'],
        'invoices': a['n'],
        'total_due_EGP': round(a['amount'], 2),
        'overdue_EGP': round(a['overdue'], 2),
        'due_within_1_month_EGP': round(a['due_within'], 2),
        'due_later_EGP': round(a['later'], 2),
        'first_due_date': str(a['first_due']), 'last_due_date': str(a['last_due']),
    })
rows.sort(key=lambda r: r['due_within_1_month_EGP'] + r['overdue_EGP'], reverse=True)
detail.sort(key=lambda r: (r['due_date'], -r['amount_due_EGP']))

def write(name, rows):
    p = os.path.join(OUT, name)
    with open(p, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    return p
write('due_within_month_by_customer.csv', rows)
write('maturity_schedule_detail.csv', detail)

tot = sum(r['total_due_EGP'] for r in rows)
within = sum(r['due_within_1_month_EGP'] for r in rows)
over = sum(r['overdue_EGP'] for r in rows)
print(f'credit term: 30 days | reference today: {TODAY} | window end: {WINDOW_END}')
print(f'customers with unpaid invoices: {len(rows)}')
print(f'total unpaid (credit) amount: {tot:,.2f} EGP')
print(f'  overdue (due_date<today): {over:,.2f}')
print(f'  due within 1 month:       {within:,.2f}')
print(f'  due later:                {tot-within-over:,.2f}')
print('Top 10 by amount due within the month:')
for r in rows[:10]:
    print(f"  {r['customer'][:24]:24s} within={r['due_within_1_month_EGP']:12,.2f} "
          f"due {r['first_due_date']}..{r['last_due_date']} [{r['salesperson']}]")
