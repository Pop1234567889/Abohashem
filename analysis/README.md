# تحليل مبيعات ومديونية أبو هاشم للحوم / Abu Hashem Meats — Sales & Debt Analysis

تحليل مبني على ملفات الـ PDF الموجودة في جذر المستودع:
Analysis built from the PDF files in the repository root:

- `فواتير مبيعات ...pdf` — 256 فاتورة مبيعات (الفترة 2026/6/2 → 2026/6/27).
  256 sales invoices (period 2026/6/2 → 2026/6/27).
- 9 تقارير "حسابات العملاء" لكل مندوب (الأرصدة المدينة/الدائنة بتاريخ 2026/6/24).
  9 per-salesperson "customer accounts" reports (debit/credit balances as of 2026/6/24).

## المخرجات / Outputs (`analysis/output/`)

| الملف / File | المحتوى / Contents |
|---|---|
| `report.md` | التقرير الكامل ثنائي اللغة / full bilingual report |
| `items_summary.csv` | لكل صنف: الكمية، المبيعات، **متوسط سعر البيع** / per item: qty, sales, **avg selling price** |
| `customers_summary.csv` | لكل عميل: المبيعات، المديونية، **نسبة وقيمة المكافأة** / per customer: sales, debt, **bonus % & amount** |
| `salespeople_summary.csv` | لكل مندوب: المديونية، المبيعات، **الانكشاف الائتماني** / per salesperson: debt, sales, **credit exposure** |
| `zero_invoices.csv` | الفواتير صفرية القيمة / zero-value invoices |
| `summary.json` | المؤشرات الإجمالية / headline KPIs |

## إعادة التشغيل / Reproduce

```bash
pip install pymupdf
python analysis/parse_invoices.py      # PDF الفواتير  -> invoices.json
python analysis/parse_salespeople.py   # تقارير المناديب -> salespeople.json
python analysis/analyze.py             # -> analysis/output/
```

## ما الذي يطلبه التحليل ويغطّيه / What the analysis covers

1. **استخراج البيانات من الـ dataset** — قراءة كل الفواتير والتقارير برمجيًا.
   Data extraction from the dataset — all invoices/reports parsed programmatically.
2. **تحليل المبيعات** — إجماليات، أصناف، عملاء، مناديب.
   Sales analysis — totals, items, customers, salespeople.
3. **مديونية كل مندوب + معدل ائتمان شهري** — معروض كـ *انكشاف ائتماني* (شهور المبيعات المحتجزة كذمم)، **بدون فائدة** حسب الاختيار.
   Debt per salesperson at a monthly credit rate — shown as *credit exposure* (months of sales tied in receivables); **no interest**, per choice.
4. **المديونية المتأخرة** — أرصدة العملاء المدينة من تقارير الحسابات.
   Arrears indebtedness — customer debit balances from the account reports.
5. **متوسط سعر بيع كل صنف** = إجمالي المبيعات ÷ الكمية (صافي بعد الخصم).
   Average selling price per item = sales ÷ qty (net of discount).
6. **الفواتير الصفرية** — الفواتير بقيمة صفر.
   Zero invoices — zero-value invoices.
7. **نسبة مكافأة لكل عميل** — شرائح حسب حجم الشراء.
   Bonus % per customer — purchase-volume tiers.
8. **مبيعات كل عميل وكميات الأصناف.**
   Sales per customer and item quantities.

## الافتراضات (قابلة للتعديل في `analyze.py`) / Assumptions (configurable in `analyze.py`)

- **الائتمان / Credit:** انكشاف فقط بدون فائدة. Exposure only, no interest.
- **شرائح المكافأة / Bonus tiers:** `0% (<5k) · 0.5% (5k–25k) · 1% (25k–100k) · 1.5% (100k–250k) · 2% (≥250k)` ج/EGP.

## جودة الاستخراج / Extraction quality

تم التحقق آليًا: مجموع بنود كل فاتورة = إجمالي الفاتورة (٠ اختلاف)؛ سعر×كمية×(١−خصم)=إجمالي البند (٠ اختلاف)؛ وكل تقرير مندوب يطابق إجمالياته المطبوعة.
Validated automatically: per-invoice line sum = invoice total (0 mismatches); price×qty×(1−disc)=line total (0 mismatches); every salesperson report ties to its printed totals.
