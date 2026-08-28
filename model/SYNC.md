# مزامنة بيانات التكلفة

`model_rows.json` المولَّد هنا تستهلكه أيضًا منصة تحليل المبيعات، حيث يُربط
بفواتير المبيعات لاحتساب الربحية لكل عميل ومندوب وعلامة تجارية وصنف.

| | |
|---|---|
| الوجهة | [`hannopo1/ABOHASHEM-sales-`](https://github.com/hannopo1/ABOHASHEM-sales-) |
| المسار هناك | `data/cost/model_rows.json` |
| المستهلِك | `analysis/13_join_cost_margin.py` |

## بعد أي تعديل على نموذج التكاليف

```bash
python3 model/costing_pricing_model.py     # يعيد توليد model_rows.json
# ثم انسخه إلى مستودع المبيعات وحدّث الالتزام في data/cost/PROVENANCE.md
```

هذا المستودع هو **المصدر الوحيد** لبيانات التكلفة. النسخة في مستودع المبيعات
مشتقّة ولا تُحرَّر مباشرة.
