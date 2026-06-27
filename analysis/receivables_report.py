# -*- coding: utf-8 -*-
"""Per-customer due-indebtedness (accounts receivable) report with monthly
credit exposure (debt expressed in months of sales). No interest applied."""
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
    'محمد خليل 26-6.pdf': 'محمد خليل', 'محمود السيد.pdf': 'محمود السيد',
    'هانى 26-6.pdf': 'هانى',
}
def norm(c): return (c or '').replace(',', '').replace(' ', '').strip()

def to_date(s):
    y, m, d = map(int, s.split('/')); return date(y, m, d)
dates = [to_date(i['date']) for i in invoices if i['date']]
period_days = (max(dates) - min(dates)).days + 1
period_months = period_days / 30.0

# sales per customer code
sales = defaultdict(float); inv_count = defaultdict(int); names = {}
for inv in invoices:
    c = norm(inv['customer_code'])
    sales[c] += inv['invoice_total'] or 0
    inv_count[c] += 1
    names[c] = inv['customer']

# debt per customer from reports
rows = []
seen = set()
for c in custrows:
    code = norm(c['code']); seen.add(code)
    sp = FILE_TO_SP.get(c['report_file'], c['salesperson'])
    debit, credit = c['debit'], c['credit']
    due = debit - credit
    msales = sales.get(code, 0.0) / period_months if period_months else 0
    exposure = round(due / msales, 2) if msales > 0 and due > 0 else ''
    rows.append({
        'customer_code': code,
        'customer': c['customer'] or names.get(code, ''),
        'salesperson': sp,
        'due_debt_EGP': round(due, 2),
        'debit_EGP': round(debit, 2),
        'credit_EGP': round(credit, 2),
        'period_sales_EGP': round(sales.get(code, 0.0), 2),
        'monthly_sales_EGP': round(msales, 2),
        'credit_exposure_months': exposure,
    })
# invoiced customers not present in any report (cash / new accounts)
for code, s in sales.items():
    if code in seen:
        continue
    rows.append({
        'customer_code': code, 'customer': names.get(code, ''),
        'salesperson': '(غير مرتبط/unmapped)',
        'due_debt_EGP': 0.0, 'debit_EGP': 0.0, 'credit_EGP': 0.0,
        'period_sales_EGP': round(s, 2),
        'monthly_sales_EGP': round(s / period_months, 2),
        'credit_exposure_months': '',
    })

rows.sort(key=lambda r: r['due_debt_EGP'], reverse=True)

path = os.path.join(OUT, 'receivables_by_customer.csv')
with open(path, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader(); w.writerows(rows)

tot_due = sum(r['due_debt_EGP'] for r in rows)
with_debt = [r for r in rows if r['due_debt_EGP'] > 0]
print(f'customers total: {len(rows)} | with due debt: {len(with_debt)} | total due: {tot_due:,.2f} EGP')
print(f'period: {min(dates)} -> {max(dates)} ({period_days}d, {period_months:.2f} months)')
print('Top 10 by due debt:')
for r in rows[:10]:
    print(f"  {r['customer'][:26]:26s} {r['due_debt_EGP']:13,.2f}  exp={r['credit_exposure_months']} mo  [{r['salesperson']}]")
print('wrote', path)
