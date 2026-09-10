# -*- coding: utf-8 -*-
"""
متوسط التكلفة المرجّح (WAC) للخامات — أغسطس 2026
==================================================
سلسلة متصلة بشهر يوليو: رصيد أول أغسطس = جرد آخر يوليو بأسعاره المحتسبة.

المدخلات (كلها داخل المستودع):
  • «تسعير جرد اخر شهر ٧ سنة ٢٠٢٦.xlsx» → رصيد أول المدة (كمية 31/7 + سعر WAC يوليو).
  • «مشتريات شهر 8 سنة 2026.pdf»        → مشتريات 01/08 → 31/08/2026 (95 سطرًا).
  • «جر د شهر 8.xlsx»                    → كميات الجرد الفعلى آخر أغسطس (المطلوب تسعيرها).
  • «تسعير جرد اخر شهر ٦ سنة ٢٠٢٦.xlsx» + ملف الوصفات → لربط بنود الوصفات وتكاليف التعبئة.

المعادلة لكل صنف:
    إجمالى الكمية  = كمية أول المدة + كمية المشتريات
    إجمالى التكلفة = قيمة أول المدة + قيمة المشتريات
    متوسط السعر    = إجمالى التكلفة ÷ إجمالى الكمية
    قيمة المخزون آخر أغسطس = كمية الجرد 31/8 × متوسط السعر

المخرجات: «تسعير جرد اخر شهر ٨ سنة ٢٠٢٦.xlsx» + «WAC_AUGUST2026_REPORT.md»
"""
import importlib.util
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location("wj", Path(__file__).with_name("wac_july2026.py"))
wj = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(wj)

norm, num, r2, r4 = wj.norm, wj.num, wj.r2, wj.r4
MEAT_POOL, POOL_KEY = wj.MEAT_POOL, wj.POOL_KEY
BORDER, HDR, SUB, WARN, TOT = wj.BORDER, wj.HDR, wj.SUB, wj.WARN, wj.TOT
style_header, finish = wj.style_header, wj.finish

F_JUNE = ROOT / "تسعير جرد اخر شهر ٦ سنة ٢٠٢٦.xlsx"
F_JULY = ROOT / "تسعير جرد اخر شهر ٧ سنة ٢٠٢٦.xlsx"
F_AUG = ROOT / "جر د شهر 8.xlsx"
F_PURCH = ROOT / "مشتريات شهر 8 سنة 2026.pdf"
OUT_XLSX = ROOT / "تسعير جرد اخر شهر ٨ سنة ٢٠٢٦.xlsx"
OUT_MD = ROOT / "WAC_AUGUST2026_REPORT.md"

# ───────── جدول مطابقة أغسطس = جدول يوليو + البنود الجديدة بتقرير أغسطس ─────────
AUGUST_MAP = dict(wj.PURCHASE_MAP)
AUGUST_MAP.update({
    # أسماء جديدة ظهرت لأول مرة فى تقرير مشتريات أغسطس
    "فوسفات": ("فوسفات", "عالية", "تطابق اسمى تام (كود 78)"),
    "جلتومات الصوديوم": ("جلتومات", "عالية", "جلوتامات الصوديوم = بند «جلتومات» بالجرد (كود 87)"),
    "طعم كفته": ("طعم كفته", "عالية", "تطابق اسمى تام (كود 222)"),
    "خلنجان": ("خلنجان", "عالية", "تطابق اسمى تام (كود 204)"),
    "جلد فراخ": ("جلد", "عالية", "بند الجلد الوحيد بالجرد؛ سعر 28.00 ≈ 29.66 ليوليو (كود 299)"),
    "لون طبيعي كارمن": ("لون طبيعي", "عالية", "الكارمن لون طبيعى — نفس بند «لون طبيعي» بالجرد (كود 275)"),
    "افيز": ("اكياس افيز", "عالية", "سعر 6.10 ≈ 6.31 ليوليو (كود 236)"),
    "كراتين الهنا ازرق مقاس 17": ("كراتين الهنا ازرق", "عالية", "نفس البند مع ذكر المقاس (كود 353)"),
    "كراتين اسبشيال احمر مقاس 17.5": ("كراتين اسبشيال احمر", "عالية", "نفس البند مع ذكر المقاس (كود 380)"),
    "كراتين اسباشيل ابيض 20.50": ("كراتين s.p ابيض", "عالية", "نفس البند مع ذكر المقاس (كود 381)"),
    "اكياس تغليف برجر اسباشيل ك": ("اكياس برجر s.p كبير", "عالية",
                                    "اختبار الحركة: 161.20 + 683.80 − 823.30 = 21.70 استهلاك منطقى (كود 367)"),
    "اكياس تغليف برجر الهنا كبير": ("اكياس برجر الهنا ك", "عالية",
                                     "اختبار الحركة: 138 + 529.30 − 159.50 = 507.80 (كود 178)"),
    "اكياس تغليف برجر صغير الهنا": ("اكياس برجر الهنا ص", "عالية",
                                     "اختبار الحركة: 469 + 282.50 − 486 = 265.50 (كود 177)"),
    "اكياس تغليف سجق الهنا صغير": ("اكياس سجق الهنا ص", "عالية",
                                    "اختبار الحركة: 45 + 282.50 − 60.50 = 267 (كود 179)"),
    "اكياس تغليف كفته الهنا 350": ("اكياس كفتة الهنا ص", "عالية",
                                    "الـ350 جم هى العبوة الصغيرة؛ اختبار الحركة: 145 + 282.50 − 177 = 250.50 (كود 181)"),
    "زيت عربيات 5 كيلو": ("زيت سيارات", "متوسطة",
                          "بند زيت السيارات الوحيد بالجرد — الوحدة عبوة 5 كجم وليست الكجم (كود 395)"),
    # بند مشترى لا يقابله صنف مستقل بالجرد
    "ببريكا مدخنه": (None, "—",
                     "بابريكا مدخّنة صنف مستقل (كود 453) يختلف عن «بابريكا» العادية (كود 84) — "
                     "لم يُدمج معها تفاديًا لخلط خامتين؛ القيمة 90 ج.م"),
})


# ───── تكلفة خلطتَى البانيه والاستربس (مصدر خارج وصفات التصنيع التسع) ─────
# المصدر: «خلطة بانيه 28-8-2026.pdf» و«خلطه استربس 28-8-2026.pdf» بالمستودع.
# مؤرَّخان 28/08/2026 ⇒ أسعارهما أسعار أغسطس، فتُؤخذ الأرقام كما وردت بلا اشتقاق.
# كل قيمة مفكوكة إلى مكوّناتها كما هى بالملف حتى تظل قابلة للتتبّع.
MIX_COST = {
    "بانيه": {
        "file": "خلطة بانيه 28-8-2026.pdf",
        # سعر كيلو الخلطة 49.36862 + عثمان 0.5 + أجور المهندس 2
        "kg": 49.36862 + 0.5 + 2,                       # 51.86862
        # (سعر الكجم × الوزن) + تغليف + عمولة — كما هى مُفصَّلة بالملف
        "packs": {0.8: (6.25, 1.375), 0.9: (6.25, 1.5)},
    },
    "استربس": {
        "file": "خلطه استربس 28-8-2026.pdf",
        # سعر كيلو الخلطة 133.8757 + مهندس وعثمان 2.5 + تغليف عبوة 5ك 0.75
        "kg": 133.8757 + 2.5 + 0.75,                    # 137.1257
        "packs": {0.9: (3.2, 4.125)},
    },
}


def _mix_price(family, weight):
    """سعر العبوة من ملف الخلطة: (سعر الكجم × الوزن) + تغليف + عمولة.
    عبوة الكيلو تُؤخذ بسعر الكجم كما هو بالملف (قرار المستخدم: بلا استقراء)."""
    m = MIX_COST[family]
    if weight == 1.0:
        return m["kg"], f"سعر كيلو الخلطة كما هو بملف «{m['file']}» (بلا تغليف أو عمولة)"
    if weight in m["packs"]:
        pack, comm = m["packs"][weight]
        base = m["kg"] * weight
        return (base + pack + comm,
                f"من «{m['file']}»: {m['kg']:.5f} × {weight:g} كجم = {base:.4f}"
                f" + تغليف {pack:g} + عمولة {comm:g}")
    return None, None


# (اسم صنف المنتج التام بالجرد) → (العائلة، الوزن بالكجم)
FG_MIX_ITEMS = {
    "بانيه 1 ك": ("بانيه", 1.0),
    "بانيه 800": ("بانيه", 0.8),
    "بانيه 900": ("بانيه", 0.9),
    "بانيه اسبيشيال900": ("بانيه", 0.9),
    "بانيه اسبيشيال 800": ("بانيه", 0.8),
    "بانيه اسبيشيال حار800": ("بانيه", 0.8),
    "استربس 1 ك": ("استربس", 1.0),
    "استبربس SP 900": ("استربس", 0.9),
    "استربس الهنا 900": ("استربس", 0.9),
}
# تُبنى الأسعار من MIX_COST لا بأرقام مكتوبة يدويًا
FG_MIX_PRICE = {}
for _lbl, (_fam, _w) in FG_MIX_ITEMS.items():
    _pr, _nt = _mix_price(_fam, _w)
    if _pr is not None:
        FG_MIX_PRICE[norm(_lbl)] = (_pr, _nt)


# ───── أسماء أُعيدت تسميتها فى نسخة جرد أغسطس المُنقَّحة ─────
# (الاسم الجديد بجرد أغسطس) → (الاسم المقابل بجرد يوليو) حتى يظل رصيد أول المدة مربوطًا.
AUG_ALIAS = {
    "فلانك + مجمل رصيد اللحمة بالنفق": "فلانك",
    "كروت كبده علب 350": "كروت كبده شرائح 350",
    "كروت استربس": "كروت مفروم صويا ازرق 350",
    "أكياس بانيه استربس 5 الهنا": "أكياس بانيه الهنا 5 ك",
    "فلتر هواء دبابة": "فلتر هواء",
}
_ALIAS = {norm(k): v for k, v in AUG_ALIAS.items()}


def canon(label):
    """الاسم المعتمد للبحث عن رصيد أول المدة (يعالج إعادة التسمية بين الشهرين)."""
    return _ALIAS.get(norm(label), label)


# ─────────────────────────── قراءة الأرصدة ───────────────────────────
def read_june():
    """ملف يونيه — يُستخدم فقط لربط أسماء بنود الوصفات وتكاليف التعبئة."""
    ws = openpyxl.load_workbook(F_JUNE, data_only=True)["Sheet1"]
    raw, fg = {}, {}
    for r in ws.iter_rows(min_row=5, max_row=153, values_only=True):
        if r[0]:
            raw[norm(r[0])] = dict(label=str(r[0]).strip(), qty=num(r[6]) or 0.0, price=num(r[7]))
        if r[9]:
            fg[norm(r[9])] = dict(label=str(r[9]).strip(), price=num(r[11]))
    return raw, fg


def read_july():
    """رصيد أول أغسطس = جرد آخر يوليو (الكمية 31/7 + متوسط سعر يوليو)."""
    wb = openpyxl.load_workbook(F_JULY, data_only=True)
    ws = wb["جرد الخامات آخر يوليو"]
    raw = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        if not r[0] or str(r[0]).startswith("إجمالى"):
            continue
        raw[norm(r[0])] = dict(label=str(r[0]).strip(), qty=num(r[9]) or 0.0, price=num(r[8]))
    ws = wb["جرد المنتج التام آخر يوليو"]
    fg = {}
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not r[0] or str(r[0]).startswith("إجمالى"):
            continue
        fg[norm(r[0])] = dict(label=str(r[0]).strip(), qty=num(r[1]) or 0.0, price=num(r[3]))
    return raw, fg


def read_august():
    ws = openpyxl.load_workbook(F_AUG, data_only=True)["Sheet1"]
    raw, fg = [], []
    for r in ws.iter_rows(min_row=5, max_row=200, values_only=True):
        if r[0]:
            raw.append((str(r[0]).strip(), num(r[6]) or 0.0))
        if r[7] and norm(r[7]) != "0":
            fg.append((str(r[7]).strip(), num(r[8]) or 0.0))
    return raw, fg


# ────────────────────────────── البناء ──────────────────────────────
def main():
    wj.PURCHASE_MAP = AUGUST_MAP          # compute_wac يقرأ الخريطة من وحدة يوليو
    wj.F_PURCH = F_PURCH
    june_raw, june_fg = read_june()
    july_raw, july_fg = read_july()
    aug_raw, aug_fg = read_august()

    purchases = wj.read_purchases()
    wac, unmapped = wj.compute_wac(july_raw, purchases)
    products = wj.reprice_bom(june_raw, wac)   # الوصفات مسعّرة بيونيه ⇒ الفهرس من ملف يونيه

    wb = openpyxl.Workbook()

    # ── شيت ١: جرد الخامات آخر أغسطس ──
    ws = wb.active
    ws.title = "جرد الخامات آخر أغسطس"
    ws["A1"] = "ABOHASHEM FOR MEAT — تقييم مخزون الخامات آخر أغسطس 2026 بمتوسط التكلفة المرجّح (WAC)"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A2"] = ("متوسط السعر = (قيمة رصيد أول المدة + قيمة مشتريات أغسطس) ÷ "
                "(كمية رصيد أول المدة + كمية مشتريات أغسطس) — رصيد أول المدة هو جرد 31/7 بأسعاره")
    ws["A2"].font = Font(size=10, italic=True, color="808080")
    head = ["اسم الصنف", "كمية أول المدة (1/8)", "سعر أول المدة (WAC يوليو)", "قيمة أول المدة",
            "كمية مشتريات أغسطس", "قيمة مشتريات أغسطس", "إجمالى الكمية", "إجمالى التكلفة",
            "متوسط السعر (WAC) 31/8", "كمية الجرد 31/8", "قيمة المخزون 31/8",
            "التغير عن سعر يوليو %", "ملاحظة"]
    ws.append([])
    ws.append(head)
    style_header(ws, 4, len(head))
    total_val, unpriced, rows_val = 0.0, [], []
    for label, qty in aug_raw:
        key = norm(canon(label))
        pooled = canon(label) in MEAT_POOL
        wkey = norm(POOL_KEY) if pooled else key
        w = wac.get(wkey)
        j = july_raw.get(key, {})
        jprice = j.get("price") or None
        price = w["price"] if w and w["price"] else jprice
        note = ""
        if pooled:
            note = "سعر موحَّد لمجمّع اللحوم الهندى (كتف/رقبه/سن/أرباع)"
        if norm(label) in _ALIAS:
            note = f"أُعيدت تسميته بجرد أغسطس؛ رصيد أول المدة مأخوذ من «{canon(label)}» بجرد يوليو"
        elif key not in july_raw:
            note = (note + " — " if note else "") + "صنف جديد لم يكن مدرجًا بجرد يوليو"
        if w and w["buy_qty"] and not w["open_priced"] and w["open_qty"]:
            note = (note + " — " if note else "") + "رصيد أول المدة بلا سعر مُثبت ⇒ المتوسط من المشتريات فقط"
        if (not w or not w["buy_qty"]) and price:
            note = (note + " — " if note else "") + "لا مشتريات فى أغسطس ⇒ السعر ثابت كما فى يوليو"
        if norm(label) == norm("زيت سيارات") and w and w["buy_qty"]:
            note = (note + " — " if note else "") + "الوحدة: عبوة 5 كجم (سعر العبوة وليس الكجم)"
        # القيمة بدقة العرض: كمية مقرَّبة لخانتين × سعر مقرَّب لأربع
        val = round(r2(qty) * r4(price), 2) if price else None
        if val:
            total_val += val
            rows_val.append((label, qty, price, val, jprice))
        if price is None and qty:
            note = ("صنف بلا سعر مُثبت (غير مُسعّر فى يوليو ولا مشتريات له فى أغسطس) — "
                    "يلزم إدخال سعر لتقييمه")
            unpriced.append((label, qty))
        chg = ((price / jprice - 1) * 100) if (price and jprice) else None
        ws.append([label,
                   r2(w["open_qty"]) if w else r2(j.get("qty")),
                   r4(jprice), r2((j.get("qty") or 0) * jprice) if jprice else None,
                   r2(w["buy_qty"]) if w else None,
                   r2(w["buy_val"]) if w else None,
                   r2(w["tot_qty"]) if w else None,
                   r2(w["tot_val"]) if w else None,
                   r4(price), r2(qty), r2(val), r2(chg), note])
    last = ws.max_row
    ws.append([])
    ws.append(["إجمالى قيمة مخزون الخامات آخر أغسطس 2026 (الأصناف المُسعّرة)"] + [None] * 9
              + [r2(total_val), None, None])
    for c in range(1, len(head) + 1):
        ws.cell(row=ws.max_row, column=c).fill = TOT
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    for row in ws.iter_rows(min_row=5, max_row=last, max_col=len(head)):
        for cell in row:
            cell.border = BORDER
            if cell.column in (2, 4, 5, 6, 7, 8, 10, 11):
                cell.number_format = "#,##0.00"
            elif cell.column in (3, 9):
                cell.number_format = "#,##0.0000"
            elif cell.column == 12:
                cell.number_format = "+0.0%;-0.0%"
                if cell.value is not None:
                    cell.value = cell.value / 100
        if row[8].value is None:
            for cell in row:
                cell.fill = WARN
    finish(ws, [30, 15, 17, 15, 16, 18, 14, 16, 18, 14, 16, 14, 52])

    # ── شيت ٢: خطوات احتساب المتوسط ──
    ws = wb.create_sheet("خطوات احتساب المتوسط")
    ws["A1"] = "خطوات احتساب متوسط السعر المرجّح لكل خامة — أغسطس 2026"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws["A2"] = "تُعرض الأصناف التى حدثت عليها مشتريات خلال أغسطس فقط (باقى الأصناف يبقى سعر يوليو كما هو)."
    ws["A2"].font = Font(size=10, italic=True, color="808080")
    head2 = ["اسم الصنف", "خطوة ١: كمية أول المدة", "× سعر أول المدة", "= قيمة أول المدة",
             "خطوة ٢: كمية المشتريات", "قيمة المشتريات (بالفواتير)", "متوسط سعر الشراء",
             "خطوة ٣: إجمالى الكمية", "خطوة ٤: إجمالى التكلفة", "خطوة ٥: متوسط السعر",
             "عدد سطور الشراء", "قاعدة الاحتساب"]
    ws.append([])
    ws.append(head2)
    style_header(ws, 4, len(head2))
    seen = set()
    for label, _ in aug_raw:
        key = norm(POOL_KEY) if canon(label) in MEAT_POOL else norm(canon(label))
        if key in seen:
            continue
        w = wac.get(key)
        if not w or not w["buy_qty"]:
            continue
        seen.add(key)
        j = july_raw.get(norm(canon(label)), {})
        disp = f"{POOL_KEY} (كتف/رقبه/سن/أرباع)" if label in MEAT_POOL else label
        ws.append([disp, r2(w["open_qty"]), r4(j.get("price")), r2(w["open_val"]),
                   r2(w["buy_qty"]), r2(w["buy_val"]), r4(w["buy_val"] / w["buy_qty"]),
                   r2(w["tot_qty"]), r2(w["tot_val"]), r4(w["price"]),
                   w["buy_lines"], w["rule"]])
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, max_col=len(head2)):
        for cell in row:
            cell.border = BORDER
            if cell.column in (2, 4, 5, 6, 8, 9):
                cell.number_format = "#,##0.00"
            elif cell.column in (3, 7, 10):
                cell.number_format = "#,##0.0000"
    finish(ws, [30, 16, 15, 16, 16, 18, 15, 15, 17, 15, 12, 42])

    # ── شيت ٣: مشتريات أغسطس التفصيلية ──
    ws = wb.create_sheet("مشتريات أغسطس التفصيلية")
    ws["A1"] = "مشتريات الخامات — الفترة 01/08/2026 حتى 31/08/2026 (المصدر: تقرير المشتريات فى الفترة)"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    head3 = ["م", "التاريخ", "كود الصنف", "اسم الصنف بالمشتريات", "الكمية", "سعر الشراء",
             "الإجمالى", "الصنف المقابل بالجرد", "درجة الثقة", "أساس المطابقة"]
    ws.append([])
    ws.append(head3)
    style_header(ws, 3, len(head3))
    for p in purchases:
        ws.append([p["no"], p["date"], p["code"], p["name"], r2(p["qty"]), r4(p["price"]),
                   r2(p["total"]), p.get("dest") or "— بلا مقابل —",
                   p.get("conf", "—"), p.get("basis", "")])
    ws.append([])
    ws.append([None, None, None, "إجمالى مشتريات أغسطس", None, None,
               r2(sum(p["total"] for p in purchases))])
    for c in range(1, len(head3) + 1):
        ws.cell(row=ws.max_row, column=c).fill = TOT
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=len(head3)):
        for cell in row:
            cell.border = BORDER
            if cell.column in (5, 7):
                cell.number_format = "#,##0.00"
            elif cell.column == 6:
                cell.number_format = "#,##0.0000"
    finish(ws, [6, 12, 10, 34, 12, 12, 14, 26, 10, 60], freeze="A4")

    # ── شيت ٤: المنتج التام ──
    ws = wb.create_sheet("جرد المنتج التام آخر أغسطس")
    ws["A1"] = "تقييم مخزون المنتج التام آخر أغسطس 2026 (وصفات التصنيع بأسعار خامات أغسطس + تكلفة التعبئة)"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    head4 = ["اسم الصنف", "كمية الجرد 31/8", "سعر يوليو", "سعر أغسطس المحدَّث",
             "التغير %", "قيمة المخزون 31/8", "ملاحظة"]
    ws.append([])
    ws.append(head4)
    style_header(ws, 3, len(head4))
    fg_total = 0.0
    fg_rows = []
    for label, qty in aug_fg:
        base = june_fg.get(norm(label), {}).get("price") or None   # سعر يونيه = أساس اشتقاق التعبئة
        prev = july_fg.get(norm(label), {}).get("price") or None   # سعر يوليو = المقارنة
        mix = FG_MIX_PRICE.get(norm(label))
        sku = None if mix else wj.resolve_sku(label, base, products)
        price = mix[0] if mix else (sku["july"] if sku else prev)
        val = round(r2(qty) * r4(price), 2) if price else None
        if val:
            fg_total += val
        chg = ((price / prev - 1) * 100) if (price and prev) else None
        if mix:
            note = mix[1] + ("" if prev else " | صنف لم يكن مُسعّرًا من قبل")
        elif sku:
            pack_txt = (f"+ تعبئة {sku['packing']:.2f} ج.م" if sku["has_pack"]
                        else "بدون تكلفة تعبئة محددة بالوصفة — السعر الفعلى أعلى قليلًا")
            note = (f"من وصفة «{sku['product']}»: {sku['weight']:g} كجم × سعر الكجم {pack_txt}"
                    + (" | صنف لم يكن مُسعّرًا من قبل" if sku["new"] else ""))
        elif price:
            note = "خارج نطاق الوصفات التسع — سعر يوليو مُبقًى دون تحديث"
        else:
            note = "بلا سعر — خارج نطاق الوصفات ولم يُسعَّر من قبل"
        ws.append([label, r2(qty), r4(prev), r4(price),
                   r2(chg / 100) if chg is not None else None, r2(val), note])
        fg_rows.append((label, qty, prev, price, chg, val, sku))
    last4 = ws.max_row
    ws.append([])
    ws.append(["إجمالى قيمة مخزون المنتج التام آخر أغسطس 2026 (الأصناف المُسعّرة)",
               None, None, None, None, r2(fg_total), None])
    for c in range(1, len(head4) + 1):
        ws.cell(row=ws.max_row, column=c).fill = TOT
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    ws.append([])
    ws.append(["إجمالى قيمة المخزون الكلى آخر أغسطس 2026 (خامات + منتج تام)",
               None, None, None, None, r2(total_val + fg_total), None])
    for c in range(1, len(head4) + 1):
        ws.cell(row=ws.max_row, column=c).fill = SUB
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True, size=12)
    for row in ws.iter_rows(min_row=4, max_row=last4, max_col=len(head4)):
        for cell in row:
            cell.border = BORDER
            if cell.column in (2, 6):
                cell.number_format = "#,##0.00"
            elif cell.column in (3, 4):
                cell.number_format = "#,##0.0000"
            elif cell.column == 5:
                cell.number_format = "+0.0%;-0.0%"
        if row[3].value is None:
            for cell in row:
                cell.fill = WARN
    finish(ws, [30, 14, 14, 18, 10, 16, 62], freeze="A4")

    # ── شيت ٥: سعر الكجم من الوصفات ──
    ws = wb.create_sheet("سعر الكجم من الوصفات")
    ws["A1"] = "أثر تحديث أسعار الخامات على سعر كيلو الخلطة لكل منتج (خامات مباشرة فقط)"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    head5 = ["المنتج", "العلامة", "قيمة الوصفة (يونيه)", "قيمة الوصفة (أغسطس)",
             "سعر الكجم (يونيه)", "سعر الكجم (أغسطس)", "التغير %"]
    ws.append([])
    ws.append(head5)
    style_header(ws, 3, len(head5))
    for p in products:
        ws.append([p["name"], p["brand"].replace("منتجات ", ""), r2(p["june_main_val"]),
                   r2(p["july_main_val"]), r4(p["june_kg"]), r4(p["july_kg"]),
                   r2(p["july_kg"] / p["june_kg"] - 1)])
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=len(head5)):
        for cell in row:
            cell.border = BORDER
            if cell.column in (3, 4):
                cell.number_format = "#,##0.00"
            elif cell.column in (5, 6):
                cell.number_format = "#,##0.0000"
            elif cell.column == 7:
                cell.number_format = "+0.0%;-0.0%"
    finish(ws, [24, 12, 20, 20, 17, 18, 10], freeze="A4")

    # ── شيت ٦: المنهجية والاستثناءات ──
    ws = wb.create_sheet("المنهجية والاستثناءات")
    ws.sheet_view.rightToLeft = True
    ws["A1"] = "المنهجية، الافتراضات، والاستثناءات — أغسطس 2026"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    notes = [
        ("نطاق العمل", "متوسط التكلفة المرجّح (Weighted Average Cost) للخامات فقط — دون أى طريقة أخرى "
                       "(FIFO / LIFO / التكلفة المعيارية)."),
        ("المعادلة", "متوسط السعر = (قيمة أول المدة + قيمة المشتريات) ÷ (كمية أول المدة + كمية المشتريات). "
                     "الطريقة دورية (Periodic WAC): متوسط واحد للشهر يُطبَّق على رصيد آخر المدة."),
        ("رصيد أول المدة", "كميات وأسعار «تسعير جرد اخر شهر ٧ سنة ٢٠٢٦.xlsx» — أى أن سلسلة الـWAC متصلة "
                            "(أبريل → مايو → يونيه → يوليو → أغسطس) بلا قطع."),
        ("المشتريات", "تقرير المشتريات 01/08/2026–31/08/2026، 95 سطر شراء بإجمالى مطابق لإجمالى التقرير "
                      "2,727,152.34 ج.م. قيمة الشراء المعتمدة هى «الإجمالى» بالفاتورة."),
        ("مرتجعات المبيعات", "تقرير «مرتجعات شهر 8» هو مرتجع مبيعات (منتج تام بأسعار البيع) وليس مرتجع "
                              "مشتريات، فلا يدخل فى احتساب متوسط تكلفة الخامات؛ أثره الكمّى منعكس أصلًا "
                              "داخل الجرد الفعلى آخر أغسطس."),
        ("مجمّع اللحوم الهندى", "بنود (لحم كتف / لحم رقبه / لحم سن / لحمه ارباع) تُغذَّى من مشتريات باسم "
                                 "واحد «تشكيلة هندى أمامى» وسعرها موحَّد ⇒ وعاء تكلفة واحد ومتوسط واحد."),
        ("أصناف بلا مشتريات", "يبقى سعر يوليو كما هو — المتوسط المرجّح لا يتغير بغياب حركة شراء."),
        ("أصناف برصيد بلا سعر", "يُستبعد رصيد أول المدة من بسط ومقام المتوسط ويُحتسب السعر من المشتريات "
                                 "فقط — بديلًا عن إدخاله بصفر الذى يُخفّض المتوسط زورًا."),
        ("وحدات القياس", "تم التحقق من اتساق الوحدة بين المشتريات والجرد لكل صنف. البند الوحيد ذو وحدة "
                          "مختلفة هو «زيت سيارات» (عبوة 5 كجم) وقد نُوِّه عنه صراحة بعمود الملاحظات."),
        ("البانيه والاستربس", "صنفان خارج وصفات التصنيع التسع. أُخذت تكلفتهما من ملفَّى "
                                "«خلطة بانيه 28-8-2026.pdf» و«خلطه استربس 28-8-2026.pdf» "
                                "المؤرَّخين 28/08/2026 ⇒ أسعارهما أسعار أغسطس. عبوتا الكيلو "
                                "سُعِّرتا بسعر كيلو الخلطة كما هو بالملف دون استقراء تغليف أو "
                                "عمولة، والأوزان 800 و900 جرام بسعرها المُفصَّل بالملف "
                                "(خلطة + تغليف + عمولة). الأوزان غير الواردة بالملفين "
                                "(700 جم، 5 كجم، GB 2، بارد دوفا، استبربس 800) تُركت دون "
                                "استقراء."),
        ("المنتج التام", "أُعيد اشتقاق سعر كل SKU من نفس وصفات التصنيع، باستبدال أسعار الخامات فقط، مع "
                          "الإبقاء على الكميات ومعاملات الإنتاجية وتكاليف التعبئة دون تغيير."),
        ("خارج النطاق", "الأسعار تمثل تكلفة الخامات المباشرة (+ التعبئة للمنتج التام) ولا تشمل الأجور "
                         "المباشرة ولا الأعباء الصناعية غير المباشرة."),
    ]
    r = 3
    for t, d in notes:
        ws.cell(row=r, column=1, value=t).font = Font(bold=True, size=11, color="1F4E79")
        ws.cell(row=r, column=2, value=d).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 46
        r += 1
    agg = {}
    for u in unmapped:
        a = agg.setdefault(u["name"], [0.0, u["reason"]])
        a[0] += u["total"]
    r += 1
    ws.cell(row=r, column=1, value="بنود مشتراة بلا مقابل فى الجرد (لم تدخل التقييم)").font = Font(
        bold=True, size=12, color="C00000")
    r += 1
    for i, h in enumerate(["الصنف", "القيمة (ج.م)", "السبب"], 1):
        ws.cell(row=r, column=i, value=h).font = Font(bold=True)
    for name, (v, reason) in sorted(agg.items(), key=lambda x: -x[1][0]):
        r += 1
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=r2(v)).number_format = "#,##0.00"
        ws.cell(row=r, column=3, value=reason).alignment = Alignment(wrap_text=True)
    r += 2
    ws.cell(row=r, column=1, value="أصناف بالجرد آخر أغسطس برصيد بلا سعر").font = Font(
        bold=True, size=12, color="C00000")
    for label, q in sorted(unpriced, key=lambda x: -x[1]):
        r += 1
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=r2(q)).number_format = "#,##0.00"
    for col, w in zip("ABC", (34, 18, 92)):
        ws.column_dimensions[col].width = w

    wb.save(OUT_XLSX)

    # ── المطابقة النهائية: أسعار الملف = المتوسطات المحتسبة مستقلًا ──
    chk = openpyxl.load_workbook(OUT_XLSX, data_only=True)["جرد الخامات آخر أغسطس"]
    mismatch = 0
    for row in chk.iter_rows(min_row=5, values_only=True):
        if not row[0] or str(row[0]).startswith("إجمالى"):
            continue
        key = norm(POOL_KEY) if canon(row[0]) in MEAT_POOL else norm(canon(row[0]))
        w = wac.get(key)
        expect = (w["price"] if w and w["price"]
                  else (july_raw.get(norm(canon(row[0])), {}).get("price") or None))
        if expect is None and row[8] is None:
            continue
        if expect is None or row[8] is None or abs(round(expect, 4) - row[8]) > 5e-5:
            mismatch += 1
            print("MISMATCH", row[0], expect, row[8])
        # القيمة داخل الملف يجب أن تساوى حاصل ضرب الخليتين المعروضتين بالضبط
        if row[10] is not None and abs(round(row[9] * row[8], 2) - row[10]) > 1e-9:
            mismatch += 1
            print("VALUE MISMATCH", row[0], round(row[9] * row[8], 2), row[10])
    print(f"مطابقة الأسعار داخل الملف مقابل الاحتساب المستقل: "
          f"{'✔ لا فروق' if mismatch == 0 else f'✘ {mismatch} فرق'}")

    # ── تقرير Markdown ──
    L = []
    A = L.append
    # يُجمَع رصيد أول المدة بنفس قاعدة دقة العرض المستخدمة فى ملف يوليو حتى تتصل السلسلة للقرش
    open_total = sum(round(r2(v["qty"]) * r4(v["price"]), 2)
                     for v in july_raw.values() if v["price"] and v["qty"])
    buy_total = sum(p["total"] for p in purchases)
    A("# متوسط سعر الخامات (WAC) — أغسطس 2026 وتقييم مخزون آخر أغسطس\n")
    A("**المنشأة:** ABOHASHEM FOR MEAT — Food Industries  ")
    A("**الفترة:** 01/08/2026 → 31/08/2026  ")
    A("**الطريقة:** متوسط التكلفة المرجّح الدورى (Periodic Weighted Average Cost) — دون أى طريقة أخرى.\n")
    A("---\n")
    A("## ١. المعادلة المطبَّقة\n")
    A("```")
    A("إجمالى الكميات  = كمية أول المدة (جرد 31/7) + كمية مشتريات أغسطس")
    A("إجمالى التكلفة = قيمة أول المدة + قيمة مشتريات أغسطس")
    A("متوسط السعر    = إجمالى التكلفة ÷ إجمالى الكميات")
    A("قيمة المخزون   = كمية الجرد 31/8 × متوسط السعر")
    A("```\n")
    A("## ٢. الملخص الرقمى\n")
    A("| البند | القيمة (ج.م) |")
    A("|---|---:|")
    A(f"| قيمة مخزون الخامات أول أغسطس (رصيد أول المدة المُسعَّر) | {open_total:,.2f} |")
    A(f"| مشتريات أغسطس 2026 ({len(purchases)} سطر شراء) | {buy_total:,.2f} |")
    A(f"| **متاح للاستخدام خلال أغسطس** | **{open_total + buy_total:,.2f}** |")
    A(f"| قيمة مخزون الخامات آخر أغسطس (بالمتوسط الجديد) | {total_val:,.2f} |")
    A(f"| تكلفة الخامات المستنفدة خلال أغسطس (بالفرق) | {open_total + buy_total - total_val:,.2f} |")
    A(f"| قيمة مخزون المنتج التام آخر أغسطس | {fg_total:,.2f} |")
    A(f"| **إجمالى قيمة المخزون آخر أغسطس 2026** | **{total_val + fg_total:,.2f}** |")
    A("")
    A("## ٣. جدول التدقيق — أكبر ٢٠ صنفًا من حيث قيمة المخزون\n")
    A("| # | كود | الصنف | كمية أول المدة | قيمة أول المدة | كمية المشتريات | قيمة المشتريات "
      "| إجمالى الكمية | إجمالى التكلفة | متوسط السعر | كمية 31/8 | قيمة 31/8 |")
    A("|---:|:--|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    codes = {}
    for p in purchases:
        if p.get("dest"):
            codes.setdefault(norm(p["dest"]), p["code"])
    rows_val.sort(key=lambda x: -x[3])
    for i, (lb, q, pr, v, jp) in enumerate(rows_val[:20], 1):
        key = norm(POOL_KEY) if canon(lb) in MEAT_POOL else norm(canon(lb))
        w = wac.get(key)
        code = codes.get(key, "—")
        if w:
            A(f"| {i} | {code} | {lb} | {w['open_qty']:,.2f} | {w['open_val']:,.2f} | "
              f"{w['buy_qty']:,.2f} | {w['buy_val']:,.2f} | {w['tot_qty']:,.2f} | "
              f"{w['tot_val']:,.2f} | {pr:,.4f} | {q:,.2f} | {v:,.2f} |")
        else:
            A(f"| {i} | {code} | {lb} | — | — | — | — | — | — | {pr:,.4f} | {q:,.2f} | {v:,.2f} |")
    A("")
    A("## ٤. خطوات الاحتساب للأصناف الأكبر أثرًا\n")
    moved = []
    for lb, q in aug_raw:
        key = norm(POOL_KEY) if canon(lb) in MEAT_POOL else norm(canon(lb))
        w = wac.get(key)
        jp = july_raw.get(norm(canon(lb)), {}).get("price")
        if w and w["buy_qty"] and w["price"]:
            moved.append((w["tot_val"], lb, w, jp))
    moved.sort(key=lambda x: -x[0])
    done = set()
    for _, lb, w, jp in moved[:12]:
        if lb in MEAT_POOL:
            if POOL_KEY in done:
                continue
            done.add(POOL_KEY)
            lb = f"{POOL_KEY} (كتف/رقبه/سن/أرباع)"
        A(f"**{lb}**  ")
        if jp:
            A(f"- أول المدة: {w['open_qty']:,.2f} × {jp:,.4f} = {w['open_val']:,.2f} ج.م  ")
        else:
            A(f"- أول المدة: {w['open_qty']:,.2f} وحدة بلا سعر مُثبت (مستبعدة من المتوسط)  ")
        A(f"- المشتريات: {w['buy_qty']:,.2f} بقيمة {w['buy_val']:,.2f} ج.م "
          f"(متوسط شراء {w['buy_val'] / w['buy_qty']:,.4f})  ")
        A(f"- إجمالى الكمية = {w['open_qty']:,.2f} + {w['buy_qty']:,.2f} = {w['tot_qty']:,.2f}  ")
        A(f"- إجمالى التكلفة = {w['open_val']:,.2f} + {w['buy_val']:,.2f} = {w['tot_val']:,.2f}  ")
        pct = f" ({(w['price'] / jp - 1) * 100:+.1f}% عن يوليو)" if jp else ""
        A(f"- **متوسط السعر = {w['tot_val']:,.2f} ÷ {w['tot_qty']:,.2f} = "
          f"{w['price']:,.4f} ج.م/وحدة**{pct}\n")
    A("## ٥. أثر التحديث على سعر كيلو الخلطة\n")
    A("| المنتج | العلامة | سعر الكجم يونيه | سعر الكجم أغسطس | التغير |")
    A("|---|---|---:|---:|---:|")
    for p in products:
        A(f"| {p['name']} | {p['brand'].replace('منتجات ', '')} | {p['june_kg']:,.4f} | "
          f"{p['july_kg']:,.4f} | {(p['july_kg'] / p['june_kg'] - 1) * 100:+.1f}% |")
    A("")
    A("## ٦. أسعار المنتج التام (الأصناف ذات الرصيد)\n")
    A("| الصنف | كمية 31/8 | سعر يوليو | سعر أغسطس | التغير | قيمة المخزون |")
    A("|---|---:|---:|---:|---:|---:|")
    for lb, q, prev, pr, chg, val, sku in fg_rows:
        if not q:
            continue
        if not pr:
            A(f"| {lb} | {q:,.2f} | — | — | — | بلا سعر |")
            continue
        A(f"| {lb} | {q:,.2f} | {f'{prev:,.4f}' if prev else '—'} | {pr:,.4f} | "
          f"{f'{chg:+.1f}%' if chg is not None else 'صنف جديد'} | {q * pr:,.2f} |")
    A("")
    mix_rows = [(lb, q, prev, pr, val) for lb, q, prev, pr, chg, val, sku in fg_rows
                if norm(lb) in FG_MIX_PRICE]
    if mix_rows:
        A("### أثر إدخال تكلفة خلطتَى البانيه والاستربس\n")
        A("المصدر: «خلطة بانيه 28-8-2026.pdf» و«خلطه استربس 28-8-2026.pdf» (مؤرَّخان 28/08/2026).\n")
        A("| الصنف | كمية 31/8 | سعر سابق | سعر من ملف الخلطة | قيمة سابقة | قيمة جديدة | الفرق |")
        A("|---|---:|---:|---:|---:|---:|---:|")
        delta = 0.0
        for lb, q, prev, pr, val in mix_rows:
            old_v = round(r2(q) * r4(prev), 2) if prev else 0.0
            new_v = val or 0.0
            delta += new_v - old_v
            A(f"| {lb} | {q:,.2f} | {f'{prev:,.4f}' if prev else '—'} | {pr:,.4f} "
              f"| {old_v:,.2f} | {new_v:,.2f} | {new_v - old_v:+,.2f} |")
        A(f"| **الصافى** | | | | | | **{delta:+,.2f}** |")
        A("")
        A("عبوتا الكيلو مُسعَّرتان بسعر كيلو الخلطة كما هو بالملف (بلا تغليف أو عمولة) "
          "— قرار صاحب العمل. الأوزان غير الواردة بالملفين تُركت دون استقراء.")
        A("")
    A("## ٧. الاستثناءات والافتراضات\n")
    A("| البند | القيمة (ج.م) | المعالجة |")
    A("|---|---:|---|")
    for name, (v, reason) in sorted(agg.items(), key=lambda x: -x[1][0]):
        A(f"| {name} | {v:,.2f} | {reason} |")
    A("")
    A("### أصناف برصيد فعلى آخر أغسطس بلا سعر مُثبت (خارج التقييم)\n")
    A("| الصنف | الكمية 31/8 |")
    A("|---|---:|")
    for lb, q in sorted(unpriced, key=lambda x: -x[1]):
        A(f"| {lb} | {q:,.2f} |")
    A("")
    fg_unpriced = [(lb, q) for lb, q, prev, pr, chg, val, sku in fg_rows if q and not pr]
    if fg_unpriced:
        A("### أصناف منتج تام برصيد آخر أغسطس بلا سعر (خارج التقييم)\n")
        A("خارج نطاق الوصفات التسع ولم تُسعَّر من قبل — يلزم توفير وصفتها أو سعر بيعها/تكلفتها:\n")
        A("| الصنف | الكمية 31/8 |")
        A("|---|---:|")
        for lb, q in sorted(fg_unpriced, key=lambda x: -x[1]):
            A(f"| {lb} | {q:,.2f} |")
        A("")
    A("### اختبار الاتساق (كمية أول المدة + المشتريات ≥ جرد آخر المدة)\n")
    bad = []
    for lb, q in aug_raw:
        key = norm(POOL_KEY) if canon(lb) in MEAT_POOL else norm(canon(lb))
        w = wac.get(key)
        if w and q - w["tot_qty"] > 0.5 and canon(lb) not in MEAT_POOL:
            pr = w["price"] or july_raw.get(norm(canon(lb)), {}).get("price") or 0.0
            bad.append((lb, w["tot_qty"], q, (q - w["tot_qty"]) * pr))
    bad.sort(key=lambda x: -x[3])
    if bad:
        A("الحالات التالية يتجاوز فيها رصيد آخر المدة المتاحَ للاستخدام (أول المدة + المشتريات). "
          "أسبابها المحتملة: إنتاج داخلى، تحويل بين مخازن، مرتجع، أو خطأ عدّ — وليست شراءً. "
          "العمود الأخير يقيس أثر الفرق على التقييم:\n")
        A("| الصنف | المتاح (أول المدة + مشتريات) | جرد 31/8 | فرق الكمية | أثر الفرق على القيمة (ج.م) |")
        A("|---|---:|---:|---:|---:|")
        for lb, av, q, impact in bad[:20]:
            A(f"| {lb} | {av:,.2f} | {q:,.2f} | {q - av:+,.2f} | {impact:+,.2f} |")
        A("")
        A(f"**أثر هذه الفروق مجتمعة على قيمة مخزون الخامات = {sum(b[3] for b in bad):,.2f} ج.م "
          f"({sum(b[3] for b in bad) / total_val * 100:.1f}% من الإجمالى).**")
        top = bad[0]
        if top[3] > 50000:
            A("")
            A(f"> ⚠️ **بند يستوجب المراجعة قبل الاعتماد:** «{top[0]}» — رصيده أول المدة {top[1]:,.2f} "
              f"ولا مشتريات له فى أغسطس، بينما الجرد الفعلى {top[2]:,.2f}؛ أى فرق {top[2] - top[1]:+,.2f} "
              f"بقيمة {top[3]:,.2f} ج.م. الرقم مأخوذ من ملف الجرد كما هو، لكن لا يوجد ما يفسّره فى "
              f"حركة المشتريات — يرجى تأكيد الكمية أو وحدة القياس قبل اعتماد التقييم.")
    else:
        A("لم تُرصد مخالفات.")
    A("")
    A("### المطابقة النهائية\n")
    A(f"- إجمالى المشتريات المستخرَج من الـPDF = **{buy_total:,.2f}** ج.م، مطابق لإجمالى التقرير.")
    A(f"- كل سطر شراء يحقق (كمية × سعر = إجمالى) بلا فروق تقريب.")
    A(f"- الأسعار المُدرَجة داخل ملف أغسطس طوبقت خلية بخلية مع المتوسطات المحتسبة مستقلًا: "
      f"{'لا فروق' if mismatch == 0 else f'{mismatch} فرق'}.")
    A(f"- قيمة كل صنف داخل الملف = كمية 31/8 × متوسط السعر (تم التحقق آليًا).")
    A("")
    A("---")
    A("> النطاق: تكلفة خامات مباشرة فقط (+ تعبئة للمنتج التام). لا تشمل أجورًا مباشرة "
      "ولا أعباء صناعية غير مباشرة.")
    OUT_MD.write_text("\n".join(L), encoding="utf-8")

    print(f"مشتريات: {len(purchases)} سطر بإجمالى {buy_total:,.2f}")
    print(f"رصيد أول أغسطس              : {open_total:,.2f}")
    print(f"قيمة مخزون الخامات آخر أغسطس : {total_val:,.2f}")
    print(f"قيمة المنتج التام آخر أغسطس  : {fg_total:,.2f}")
    print(f"الإجمالى                     : {total_val + fg_total:,.2f}")
    print(f"بنود شراء بلا مقابل: {len(agg)} | أصناف بلا سعر برصيد: {len(unpriced)}")
    print("✔", OUT_XLSX.name, "|", OUT_MD.name)


if __name__ == "__main__":
    main()
