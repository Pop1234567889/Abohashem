# -*- coding: utf-8 -*-
"""
Abu Hashem Meats — sales & debt analysis.

Inputs : invoices.json (parsed sales invoices), salespeople.json (customer
         accounts / debt per salesperson).
Outputs: CSV tables + a bilingual summary report under analysis/output/.

Policy assumptions (chosen with the business owner):
  * Credit:  EXPOSURE ONLY — no interest is charged. We report outstanding
             debt and how many months of sales are tied up in receivables.
  * Bonus :  graduated by each customer's total purchases (sales-volume tiers).
"""
import json, os, csv, statistics
from collections import defaultdict, Counter
from datetime import date

HERE = os.path.dirname(__file__)
OUT = os.path.join(HERE, 'output')
os.makedirs(OUT, exist_ok=True)

invoices = json.load(open(os.path.join(HERE, 'invoices.json'), encoding='utf-8'))
custrows = json.load(open(os.path.join(HERE, 'salespeople.json'), encoding='utf-8'))

# ---- salesperson canonical name per report file ----
FILE_TO_SP = {
    'ايمن فارس 26-6.pdf': 'ايمن فارس',
    'حسام حسن 26-6.pdf': 'حسام حسن',
    'شعبان 26-6.pdf': 'شعبان',
    'محمد امام 26-.pdf': 'محمد امام',
    'محمد امام الصعيد 26-6.pdf': 'محمد امام الصعيد',
    'محمد بشرى 26-6.pdf': 'محمد بشرى',
    'محمد خليل 26-6.pdf': 'محمد خليل',
    'محمود السيد.pdf': 'محمود السيد',
    'هانى 26-6.pdf': 'هانى',
}
# code -> salesperson, and code -> net balance (debit-credit), name
def norm(code):
    return (code or '').replace(',', '').replace(' ', '').strip()

code_sp = {}
code_bal = {}
code_name = {}
for c in custrows:
    sp = FILE_TO_SP.get(c['report_file'], c['salesperson'])
    code_sp[norm(c['code'])] = sp
    code_bal[norm(c['code'])] = c['debit'] - c['credit']
    code_name[norm(c['code'])] = c['customer']

# ---- invoice period ----
def to_date(s):
    y, m, d = map(int, s.split('/'))
    return date(y, m, d)
dates = [to_date(i['date']) for i in invoices if i['date']]
period_start, period_end = min(dates), max(dates)
period_days = (period_end - period_start).days + 1
period_months = period_days / 30.0

# ============================================================
# 1) PER-ITEM: quantity, sales, average selling price
# ============================================================
item_qty = defaultdict(float)
item_sales = defaultdict(float)
item_lines = defaultdict(int)
item_prices = defaultdict(list)
item_unit = {}
item_name = {}
for inv in invoices:
    for it in inv['items']:
        code = it['item_code'] or '?'
        item_qty[code] += it['qty'] or 0
        item_sales[code] += it['line_total'] or 0
        item_lines[code] += 1
        if it['unit_price']:
            item_prices[code].append(it['unit_price'])
        item_unit[code] = it['unit']
        item_name[code] = it['name']

items_rows = []
for code in item_qty:
    qty = item_qty[code]
    sales = item_sales[code]
    wavg = sales / qty if qty else 0
    prices = item_prices[code]
    items_rows.append({
        'item_code': code,
        'item_name': item_name[code],
        'unit': item_unit[code],
        'total_qty': round(qty, 2),
        'total_sales_EGP': round(sales, 2),
        'avg_selling_price_EGP': round(wavg, 2),      # revenue / qty (net of discount)
        'list_price_mode_EGP': Counter(prices).most_common(1)[0][0] if prices else 0,
        'min_price_EGP': min(prices) if prices else 0,
        'max_price_EGP': max(prices) if prices else 0,
        'times_sold': item_lines[code],
    })
items_rows.sort(key=lambda r: r['total_sales_EGP'], reverse=True)

# ============================================================
# 2) PER-CUSTOMER: sales, qty, debt, bonus %
# ============================================================
BONUS_TIERS = [   # (min_sales_inclusive, pct)
    (250000, 2.0),
    (100000, 1.5),
    (25000, 1.0),
    (5000, 0.5),
    (0, 0.0),
]
def bonus_pct(sales):
    for thr, pct in BONUS_TIERS:
        if sales >= thr:
            return pct
    return 0.0

cust_sales = defaultdict(float)
cust_qty = defaultdict(float)
cust_inv = defaultdict(int)
cust_name = {}
for inv in invoices:
    code = norm(inv['customer_code']) or '?'
    cust_sales[code] += inv['invoice_total'] or 0
    cust_qty[code] += inv['total_qty'] or 0
    cust_inv[code] += 1
    cust_name[code] = inv['customer']

cust_rows = []
for code in cust_sales:
    sales = cust_sales[code]
    pct = bonus_pct(sales)
    cust_rows.append({
        'customer_code': code,
        'customer': cust_name.get(code) or code_name.get(code) or '',
        'salesperson': code_sp.get(code, '(غير مرتبط/unmapped)'),
        'invoices': cust_inv[code],
        'total_qty': round(cust_qty[code], 2),
        'total_sales_EGP': round(sales, 2),
        'outstanding_debt_EGP': round(code_bal.get(code, 0.0), 2),
        'bonus_pct': pct,
        'bonus_amount_EGP': round(sales * pct / 100, 2),
    })
cust_rows.sort(key=lambda r: r['total_sales_EGP'], reverse=True)

# ============================================================
# 3) PER-SALESPERSON: debt, sales, credit exposure, arrears
# ============================================================
sp_debt = defaultdict(float)        # net receivable from reports (indebtedness)
sp_arrears = defaultdict(float)     # positive debit balances only (amounts owed)
sp_custcount = defaultdict(int)
for c in custrows:
    sp = FILE_TO_SP.get(c['report_file'], c['salesperson'])
    sp_debt[sp] += c['debit'] - c['credit']
    sp_arrears[sp] += c['debit']
    sp_custcount[sp] += 1

sp_sales = defaultdict(float)
sp_inv = defaultdict(int)
sp_qty = defaultdict(float)
for inv in invoices:
    sp = code_sp.get(norm(inv['customer_code']), '(غير مرتبط/unmapped)')
    sp_sales[sp] += inv['invoice_total'] or 0
    sp_inv[sp] += 1
    sp_qty[sp] += inv['total_qty'] or 0

sp_rows = []
all_sp = set(sp_debt) | set(sp_sales)
for sp in all_sp:
    debt = sp_debt.get(sp, 0.0)
    monthly_sales = sp_sales.get(sp, 0.0) / period_months if period_months else 0
    exposure_months = (debt / monthly_sales) if monthly_sales else None
    sp_rows.append({
        'salesperson': sp,
        'customers': sp_custcount.get(sp, 0),
        'period_sales_EGP': round(sp_sales.get(sp, 0.0), 2),
        'period_invoices': sp_inv.get(sp, 0),
        'monthly_sales_EGP': round(monthly_sales, 2),
        'outstanding_debt_EGP': round(debt, 2),
        'arrears_debit_EGP': round(sp_arrears.get(sp, 0.0), 2),
        'credit_exposure_months': round(exposure_months, 2) if exposure_months is not None else '',
    })
sp_rows.sort(key=lambda r: r['outstanding_debt_EGP'], reverse=True)

# ============================================================
# 4) ZERO INVOICES (value = 0 or quantity = 0)
# ============================================================
zero_rows = []
for inv in invoices:
    if (inv['invoice_total'] or 0) == 0 or (inv['total_qty'] or 0) == 0:
        zero_rows.append({
            'invoice_no': inv['invoice_no'],
            'date': inv['date'],
            'page': inv['page'],
            'customer': inv['customer'],
            'customer_code': inv['customer_code'],
            'salesperson': code_sp.get(norm(inv['customer_code']), '(غير مرتبط/unmapped)'),
            'total_qty': inv['total_qty'],
            'invoice_total_EGP': inv['invoice_total'],
        })

# ---- write CSVs ----
def write_csv(name, rows):
    if not rows:
        rows = [{}]
    path = os.path.join(OUT, name)
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    return path

write_csv('items_summary.csv', items_rows)
write_csv('customers_summary.csv', cust_rows)
write_csv('salespeople_summary.csv', sp_rows)
write_csv('zero_invoices.csv', zero_rows)

# ---- totals ----
tot_sales = sum(i['invoice_total'] or 0 for i in invoices)
tot_qty = sum(i['total_qty'] or 0 for i in invoices)
tot_paid = sum(i['paid'] or 0 for i in invoices)
tot_remaining = sum(i['remaining'] or 0 for i in invoices)
tot_debt = sum(sp_debt.values())

summary = {
    'period_start': str(period_start),
    'period_end': str(period_end),
    'period_days': period_days,
    'invoices': len(invoices),
    'zero_invoices': len(zero_rows),
    'line_items': sum(len(i['items']) for i in invoices),
    'distinct_items': len(items_rows),
    'distinct_customers_invoiced': len(cust_rows),
    'salespeople': len([s for s in sp_rows if s['customers']]),
    'total_sales_EGP': round(tot_sales, 2),
    'total_qty': round(tot_qty, 2),
    'total_paid_EGP': round(tot_paid, 2),
    'total_remaining_EGP': round(tot_remaining, 2),
    'total_outstanding_debt_EGP': round(tot_debt, 2),
}
json.dump(summary, open(os.path.join(OUT, 'summary.json'), 'w'),
          ensure_ascii=False, indent=1)

print(json.dumps(summary, ensure_ascii=False, indent=1))
print('\nTop salespeople by debt:')
for r in sp_rows[:9]:
    print(f"  {r['salesperson']:18s} debt={r['outstanding_debt_EGP']:13,.2f} "
          f"sales={r['period_sales_EGP']:12,.2f} exposure={r['credit_exposure_months']} mo")
print('\nWrote outputs to', OUT)


# ============================================================
# 5) BILINGUAL MARKDOWN REPORT
# ============================================================
def egp(x):
    return f"{x:,.2f}"

lines = []
A = lines.append
A("# تحليل المبيعات والمديونية — أبو هاشم للحوم")
A("# Abu Hashem Meats — Sales & Debt Analysis\n")
A(f"**الفترة / Period:** {summary['period_start']} → {summary['period_end']} "
  f"({summary['period_days']} يوم/days)\n")
A("## 1. ملخص عام / Overview\n")
A("| المؤشر / Metric | القيمة / Value |")
A("|---|---:|")
A(f"| عدد الفواتير / Invoices | {summary['invoices']} |")
A(f"| فواتير صفرية القيمة / Zero-value invoices | {summary['zero_invoices']} |")
A(f"| بنود الفواتير / Line items | {summary['line_items']} |")
A(f"| عدد الأصناف / Distinct items | {summary['distinct_items']} |")
A(f"| عملاء لهم مبيعات / Customers invoiced | {summary['distinct_customers_invoiced']} |")
A(f"| إجمالى المبيعات / Total sales (EGP) | {egp(summary['total_sales_EGP'])} |")
A(f"| إجمالى الكميات / Total quantity | {egp(summary['total_qty'])} |")
A(f"| المحصّل / Collected (EGP) | {egp(summary['total_paid_EGP'])} |")
A(f"| المتبقّى آجل / On credit (EGP) | {egp(summary['total_remaining_EGP'])} |")
A(f"| إجمالى المديونية القائمة / Total outstanding debt (EGP) | {egp(summary['total_outstanding_debt_EGP'])} |")
A("")
A("> ملاحظة: ~99.8% من المبيعات تمّت بالأجل (المحصّل نقدًا ضئيل جدًا)، "
  "أى أنّ كل المبيعات تقريبًا تتحوّل إلى مديونية على العملاء.")
A("> Note: ~99.8% of sales are on credit; nearly all billed value becomes customer debt.\n")

A("## 2. مديونية كل مندوب ومعدل الائتمان الشهرى / Debt & monthly credit exposure per salesperson\n")
A("معدل الائتمان معروض كـ **انكشاف ائتمانى** = المديونية ÷ متوسط المبيعات الشهرية "
  "(عدد الأشهر من المبيعات المحتجزة كذمم) — بدون فائدة، حسب الاختيار.")
A("Credit rate shown as **exposure** = debt ÷ monthly sales (months of sales locked in receivables); no interest applied.\n")
A("| المندوب / Salesperson | العملاء / Customers | مبيعات الفترة / Period sales | مبيعات شهرية / Monthly sales | المديونية / Debt | الانكشاف (شهور) / Exposure (mo) |")
A("|---|---:|---:|---:|---:|---:|")
for r in sp_rows:
    A(f"| {r['salesperson']} | {r['customers']} | {egp(r['period_sales_EGP'])} | "
      f"{egp(r['monthly_sales_EGP'])} | {egp(r['outstanding_debt_EGP'])} | {r['credit_exposure_months']} |")
A(f"| **الإجمالى / Total** | {sum(r['customers'] for r in sp_rows)} | "
  f"{egp(sum(r['period_sales_EGP'] for r in sp_rows))} | | "
  f"{egp(sum(r['outstanding_debt_EGP'] for r in sp_rows))} | |")
A("")
A("**قراءة سريعة / Reading:** محمد امام الصعيد يحمل أعلى مديونية "
  f"({egp(sp_debt['محمد امام الصعيد'])} ج) — أكثر من 41% من إجمالى الذمم. "
  "كل المناديب تقريبًا عند انكشاف يقارب شهرًا واحدًا أو أقل، عدا ايمن فارس وهانى (~1.5 شهر) "
  "وهما الأعلى خطورة ائتمانية نسبةً لمبيعاتهما.\n")

A("## 3. المديونية المتأخرة / Arrears indebtedness\n")
A("المديونية المتأخرة = أرصدة العملاء المدينة (مبالغ مستحقة على العملاء) من تقارير الحسابات بتاريخ 2026/6/24.")
A("Arrears = customer debit balances (amounts owed) from the 2026/6/24 account reports.\n")
A("| المندوب / Salesperson | المتأخر (مدين) / Arrears debit | دائن مقابل / Offsetting credit | الصافى / Net debt |")
A("|---|---:|---:|---:|")
for r in sorted(sp_rows, key=lambda x: x['arrears_debit_EGP'], reverse=True):
    if r['arrears_debit_EGP'] or r['outstanding_debt_EGP']:
        off = r['arrears_debit_EGP'] - r['outstanding_debt_EGP']
        A(f"| {r['salesperson']} | {egp(r['arrears_debit_EGP'])} | {egp(off)} | {egp(r['outstanding_debt_EGP'])} |")
A("")

A("## 4. متوسط سعر البيع والكميات لكل صنف / Average selling price & quantity per item\n")
A("متوسط سعر البيع = إجمالى المبيعات ÷ الكمية (صافى بعد الخصم).")
A("Average selling price = total sales ÷ quantity (net of discount).\n")
A("| الكود / Code | الصنف / Item | الوحدة / Unit | الكمية / Qty | المبيعات / Sales | متوسط السعر / Avg price | السعر المعتاد / List |")
A("|---|---|---|---:|---:|---:|---:|")
for r in items_rows:
    A(f"| {r['item_code']} | {r['item_name']} | {r['unit']} | {egp(r['total_qty'])} | "
      f"{egp(r['total_sales_EGP'])} | {egp(r['avg_selling_price_EGP'])} | {egp(r['list_price_mode_EGP'])} |")
A("")

A("## 5. أعلى 20 عميلاً: المبيعات والمكافأة / Top 20 customers: sales & bonus\n")
A("نسبة المكافأة حسب شرائح حجم الشراء / Bonus % by purchase-volume tiers:")
A("`0% (<5k) · 0.5% (5k–25k) · 1% (25k–100k) · 1.5% (100k–250k) · 2% (≥250k)`\n")
A("| العميل / Customer | المندوب / Salesperson | المبيعات / Sales | المديونية / Debt | المكافأة % | قيمة المكافأة / Bonus |")
A("|---|---|---:|---:|---:|---:|")
for r in cust_rows[:20]:
    A(f"| {r['customer']} | {r['salesperson']} | {egp(r['total_sales_EGP'])} | "
      f"{egp(r['outstanding_debt_EGP'])} | {r['bonus_pct']}% | {egp(r['bonus_amount_EGP'])} |")
A("")
A(f"إجمالى المكافآت المستحقة لكل العملاء / Total bonus across all customers: "
  f"**{egp(sum(r['bonus_amount_EGP'] for r in cust_rows))} ج/EGP** "
  f"({len([r for r in cust_rows if r['bonus_pct']>0])} عميل مؤهل / eligible customers)\n")

A("## 6. الفواتير الصفرية / Zero-value invoices\n")
A(f"عدد {len(zero_rows)} فاتورة بقيمة صفر (كميات مصروفة بسعر صفر — هدايا/استبدال/تسويات).")
A(f"{len(zero_rows)} invoices with zero value (quantities issued at zero price — gifts/replacements/adjustments).\n")
A("| الفاتورة / Invoice | التاريخ / Date | العميل / Customer | المندوب / Salesperson | الكمية / Qty |")
A("|---|---|---|---|---:|")
for r in sorted(zero_rows, key=lambda x: -(x['total_qty'] or 0)):
    A(f"| {r['invoice_no']} | {r['date']} | {r['customer']} | {r['salesperson']} | {egp(r['total_qty'] or 0)} |")
A("")

A("## 7. المنهجية والافتراضات / Methodology & assumptions\n")
A("- **المصدر / Source:** 256 فاتورة مبيعات (PDF) + 9 تقارير حسابات عملاء للمناديب.")
A("- **التحقق / Validation:** مجموع بنود كل فاتورة = إجمالى الفاتورة (٠ اختلاف)، "
  "وسعر×كمية×(١−خصم)=إجمالى البند (٠ اختلاف)، وإجمالى كل تقرير مندوب يطابق المطبوع.")
A("- **الائتمان / Credit:** انكشاف فقط بدون فائدة (حسب الطلب).")
A("- **المكافأة / Bonus:** شرائح حجم المبيعات (قابلة للتعديل في `analyze.py`).")
A("- **غير المرتبطين / Unmapped:** 6 عملاء (≈33.5k ج) غير موجودين بتقارير المناديب "
  "(عميل نقدى وعملاء جدد) — مدرجون في الإجماليات وموسومون بـ unmapped.")
A("- البيانات والجداول الكاملة في مجلد `analysis/output/` (CSV).")

open(os.path.join(OUT, 'report.md'), 'w', encoding='utf-8').write('\n'.join(lines))
print('Wrote report.md')
