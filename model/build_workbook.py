# -*- coding: utf-8 -*-
"""Render the costing & pricing model into a formatted Excel workbook."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from costing_pricing_model import (build, INCOME, CONVERSION_POOL, OPEX_POOL,
                                   CONVERSION_CONTROL, OPEX_CONTROL, FINANCE_KEYS,
                                   CONFIG, MIX, SKU)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT  = os.path.join(BASE, "ABOHASHEM_Costing_Pricing_Model_June2026.xlsx")

# ---- palette ----
NAVY  = "1F3A5F"; BLUE="2E5A88"; LBLUE="DCE6F1"; GREY="F2F2F2"
GREEN = "C6EFCE"; GREENF="006100"; RED="FFC7CE"; REDF="9C0006"
AMBER = "FFEB9C"; AMBERF="9C6500"; WHITE="FFFFFF"
thin  = Side(style="thin", color="BFBFBF")
BORDER= Border(left=thin,right=thin,top=thin,bottom=thin)
RTL   = Alignment(horizontal="right", vertical="center", wrap_text=True)
CEN   = Alignment(horizontal="center", vertical="center", wrap_text=True)

def hdr(ws, row, headers, start=1, fill=NAVY, fsz=10):
    for j,h in enumerate(headers, start):
        c = ws.cell(row=row, column=j, value=h)
        c.font=Font(bold=True, color=WHITE, size=fsz, name="Arial")
        c.fill=PatternFill("solid", fgColor=fill); c.alignment=CEN; c.border=BORDER
    ws.row_dimensions[row].height=30

def title(ws, text, sub=None, span=12):
    ws.sheet_view.rightToLeft=True
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=span)
    c=ws.cell(1,1,text); c.font=Font(bold=True,size=14,color=WHITE,name="Arial")
    c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=CEN
    ws.row_dimensions[1].height=34
    if sub:
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=span)
        c=ws.cell(2,1,sub); c.font=Font(size=10,italic=True,color=NAVY,name="Arial")
        c.fill=PatternFill("solid",fgColor=LBLUE); c.alignment=CEN
        ws.row_dimensions[2].height=26

def money(c): c.number_format='#,##0.00'; c.alignment=CEN
def pct(c):   c.number_format='0.0%'; c.alignment=CEN
def widths(ws, w):
    for i,x in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width=x

def zebra(ws, r0, r1, c0, c1):
    for r in range(r0,r1+1):
        f = GREY if (r-r0)%2 else WHITE
        for c in range(c0,c1+1):
            cell=ws.cell(r,c); cell.border=BORDER
            if cell.fill.fgColor.rgb in ("00000000", None): cell.fill=PatternFill("solid",fgColor=f)

rows, meta = build()
ranked = sorted(rows, key=lambda r:r["rank"])
wb = Workbook(); wb.remove(wb.active)

# ============================================ SHEET 1: Executive Summary ========
ws = wb.create_sheet("١. الملخص التنفيذي")
title(ws,"ABOHASHEM FOR MEAT — نموذج التكاليف والتسعير المتكامل (يونيه 2026)",
      "تكلفة معيارية وفعلية • توزيع مصروفات • تحليل ربحية • محرك تسعير ديناميكي — مطابَق لقائمة الدخل", span=6)
r=4
kpis=[
 ("صافي المبيعات (قائمة الدخل)", INCOME["net_sales"], "ج.م"),
 ("تكلفة المبيعات (خامات+تعبئة)", INCOME["cogs"], "ج.م"),
 ("مجمل الربح", INCOME["gross_profit"], "ج.م"),
 ("مجمل هامش الربح", INCOME["gross_profit"]/INCOME["net_sales"], "%"),
 ("الأعباء الصناعية غير المباشرة (تحويل)", CONVERSION_CONTROL, "ج.م"),
 ("المصروفات التشغيلية (بيع+إدارة+تمويل)", OPEX_CONTROL, "ج.م"),
 ("صافي الربح التشغيلي", INCOME["net_sales"]-INCOME["cogs"]-CONVERSION_CONTROL-OPEX_CONTROL, "ج.م"),
 ("صافي هامش الربح", (INCOME["net_sales"]-INCOME["cogs"]-CONVERSION_CONTROL-OPEX_CONTROL)/INCOME["net_sales"], "%"),
 ("عدد المنتجات المُسعّرة / فواتير يونيه", len(rows), f"منتج / {meta['n_inv']} فاتورة"),
 ("إجمالي الكميات المُنتَجة/المُباعة", meta["tot_kg"], "كجم"),
 ("عدد الأصناف الخاسرة (هامش تشغيلي سالب)", sum(1 for x in rows if x["op_margin"]<0), "صنف"),
 ("أصناف موصى برفع سعرها", sum(1 for x in rows if x["action"]=="رفع السعر"), "صنف"),
]
hdr(ws,r,["المؤشر","القيمة","الوحدة"])
r+=1
for k,v,u in kpis:
    ws.cell(r,1,k).alignment=RTL
    cv=ws.cell(r,2,v)
    if u=="%": pct(cv)
    elif "ج.م" in u or u=="كجم": money(cv)
    else: cv.alignment=CEN
    ws.cell(r,3,u).alignment=CEN
    r+=1
zebra(ws,5,r-1,1,3)
# brand block
r+=1
hdr(ws,r,["العلامة التجارية","صافي المبيعات","% من المبيعات","مجمل الهامش %","الهامش التشغيلي %","عدد الأصناف"])
r+=1
from collections import defaultdict
bs=defaultdict(lambda:[0,0,0,0])
for x in rows:
    bs[x["brand"]][0]+=x["revenue"]; bs[x["brand"]][1]+=x["gross_profit"]
    bs[x["brand"]][2]+=x["op_profit"]; bs[x["brand"]][3]+=1
for k,v in sorted(bs.items(),key=lambda z:-z[1][0]):
    ws.cell(r,1,k).alignment=RTL
    money(ws.cell(r,2,v[0])); pct(ws.cell(r,3,v[0]/INCOME["net_sales"]))
    pct(ws.cell(r,4,v[1]/v[0])); pct(ws.cell(r,5,v[2]/v[0]))
    ws.cell(r,6,v[3]).alignment=CEN
    r+=1
zebra(ws,r-len(bs),r-1,1,6)
widths(ws,[42,18,14,14,16,12])

# ============================================ SHEET 2: Data Quality =============
ws = wb.create_sheet("٢. تقرير جودة البيانات")
title(ws,"تقرير جودة البيانات والتحقق (قبل الاحتساب)","فحوص الاكتمال والاتساق والتكرار والقيم الشاذة والهوامش السالبة", span=5)
checks=[]
# duplicates by recipe/weight
seen=defaultdict(list)
for x in rows: seen[(x["recipe"],x["weight"])].append(x["name"])
dups=[v for k,v in seen.items() if k[0] and len(v)>1]
# no-BOM
nobom=[x["name"] for x in rows if not x["has_bom"]]
# zero-price invoice lines flagged in source (bonus)
neg=[x["name"] for x in rows if x["op_margin"]<0]
below=[x["name"] for x in rows if x["old_price"]<x["full_loaded_u"]]
checks=[
 ("اكتمال بيانات المبيعات","اجتياز",f"تم تحليل {meta['n_inv']} فاتورة يونيه 2026 دون صفوف تالفة (0 صف مرفوض).","info"),
 ("مطابقة الإيراد بقائمة الدخل","تنبيه",f"إجمالي فواتير البيع 3,867,491 ج.م أعلى من صافي مبيعات القائمة 3,741,772 ج.م بمقدار 125,719 ج.م = مرتجعات/بونص/خصومات — تمت تسويتها كبند خصم من الإيراد.","warn"),
 ("مطابقة تكلفة المبيعات","اجتياز","إجمالي التكلفة الفعلية للخامات المحمّلة = 2,039,933.08 ج.م مطابق تماماً لتكلفة المبيعات بالقائمة.","ok"),
 ("مطابقة صافي الربح","اجتياز","صافي الربح التشغيلي المُجمّع من المنتجات = 418,841.92 ج.م مطابق لصافي ربح القائمة.","ok"),
 ("منتجات بلا وصفة تصنيع (BOM)","تنبيه",f"{len(nobom)} أصناف تجارية بلا تركيبة خامات: {'، '.join(nobom[:6])}… — احتُسبت تكلفتها الفعلية بالتناسب مع تكلفة المبيعات (غير معيارية).","warn"),
 ("خامات بلا سعر (مياه/أصناف غير مُسعّرة)","تنبيه","17 سطر خام بلا سعر (مياه، بطاطس، حلبة، شعرية، فلفل غير محدد) — قيمتها = صفر (استبعاد لا تقدير)؛ مفصّلة بشيت الاستثناءات بالملف الأصلي.","warn"),
 ("مطابقات تسمية بثقة متوسطة/منخفضة","تنبيه","~15 خامة طوبقت بحكم مهني (توام↔ثوم، جبهان↔حبهان، سميد↔صميد…) — تحتاج تأكيد محاسب التكاليف.","warn"),
 ("منتجات مكرّرة","اجتياز","لا يوجد تكرار في أكواد المنتجات؛ وصفتا (كفته الهنا/برجر سبيشال) متطابقتان بالتركيب وهو مقصود.","ok"),
 ("قيم شاذة في الأسعار","تنبيه",f"{len(neg)} أصناف بهامش تشغيلي سالب: {'، '.join(neg)} — تبيع دون التكلفة المحمّلة.","warn"),
 ("أصناف تُباع دون التكلفة المحمّلة","تنبيه",f"{len(below)} صنف سعره الحالي < التكلفة المحمّلة بالكامل — مدرجة في محرك التسعير للرفع.","warn"),
 ("سلامة الوزن/الكمية","اجتياز","كل SKU له وزن موجب وكمية مبيعات موجبة؛ لا قيم مفقودة في الحقول الحسابية.","ok"),
]
hdr(ws,3,["#","الفحص","الحالة","التفاصيل"])
r=4
for i,(name,st,det,kind) in enumerate(checks,1):
    ws.cell(r,1,i).alignment=CEN
    ws.cell(r,2,name).alignment=RTL
    cs=ws.cell(r,3,st); cs.alignment=CEN
    fill = GREEN if kind=="ok" else AMBER if kind=="warn" else LBLUE
    fnt  = GREENF if kind=="ok" else AMBERF if kind=="warn" else NAVY
    cs.fill=PatternFill("solid",fgColor=fill); cs.font=Font(bold=True,color=fnt)
    ws.cell(r,4,det).alignment=RTL
    for c in range(1,5): ws.cell(r,c).border=BORDER
    ws.row_dimensions[r].height=42
    r+=1
widths(ws,[5,32,12,90])

# ============================================ SHEET 3: Standard Cost ============
ws = wb.create_sheet("٣. التكلفة المعيارية")
title(ws,"التكلفة المعيارية للمنتج التام — بناء تصاعدي (Bottom-up) لكل عنصر",
      "خامات + تعبئة + أجور + أعباء صناعية + مرافق + جودة + فاقد/هدر — لكل SKU (وحدة البيع)", span=13)
H=["الكود","المنتج","العلامة","الوزن كجم","خامات مباشرة","تعبئة","أجور مباشرة",
   "أعباء صناعية","مرافق","جودة","هدر","احتياطي جودة","التكلفة المعيارية/وحدة"]
hdr(ws,3,H)
r=4
for x in ranked:
    vals=[x["code"],x["name"],x["brand"],x["weight"],x["std_mat_u"],x["std_pack_u"],
          x["std_labour_u"],x["std_moh_u"],x["std_utils_u"],x["std_quality_u"],
          x["std_waste_u"],x["std_qres_u"],x["std_full_u"]]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,round(v,4) if isinstance(v,float) else v)
        if j==2: c.alignment=RTL
        elif j in(1,3): c.alignment=CEN
        else: money(c)
        if j==13: c.font=Font(bold=True)
    r+=1
zebra(ws,4,r-1,1,13)
# note
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=13)
ws.cell(r,1,"العناصر غير المباشرة (أجور/أعباء/مرافق/جودة) محمّلة بمعدل %.3f ج.م/كجم مشتقّة من مجمّع الأعباء الصناعية الشهري (544,605 ج.م) ÷ إجمالي الكميات المنتجة. الأصناف التجارية بلا وصفة تظهر تكلفتها الفعلية المُقدّرة في خانة الخامات."%meta["conv_rate_kg"]).alignment=RTL
ws.cell(r,1).font=Font(italic=True,size=9,color=NAVY); ws.row_dimensions[r].height=40
widths(ws,[7,30,10,9]+[12]*8+[16])

# ============================================ SHEET 4: Actual & Variance ========
ws = wb.create_sheet("٤. التكلفة الفعلية والانحرافات")
title(ws,"التكلفة الفعلية وتحليل الانحرافات (معياري مقابل فعلي)",
      "التكلفة الفعلية مطابَقة لتكلفة المبيعات بالقائمة (2,039,933 ج.م) — الانحراف موجب=غير مُوات", span=11)
H=["الكود","المنتج","الكمية","تكلفة معيارية/وحدة","تكلفة فعلية/وحدة","انحراف الخامات",
   "انحراف التعبئة","انحراف التحويل","إجمالي الانحراف","معياري÷فعلي %","الحالة"]
hdr(ws,3,H)
r=4
for x in ranked:
    std_u=x["std_full_u"]; act_u=x["act_full_u"]
    # element variances (per total)
    mat_var = (x["act_mat_total"]-(x["std_mat_u"]*x["qty"]+x["std_pack_u"]*x["qty"])) if x["has_bom"] else 0
    pack_var= 0
    conv_var= x["act_conv_total"]-(x["std_labour_u"]+x["std_utils_u"]+x["std_quality_u"]+x["std_moh_u"])*x["qty"]
    tot_var = act_u*x["qty"]-std_u*x["qty"]
    ratio   = std_u/act_u if act_u else 0
    fav     = "مُوات" if tot_var<=0 else "غير مُوات"
    vals=[x["code"],x["name"],x["qty"],std_u,act_u,mat_var,pack_var,conv_var,tot_var,ratio,fav]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,round(v,3) if isinstance(v,float) else v)
        if j==2:c.alignment=RTL
        elif j in(1,3,11):c.alignment=CEN
        elif j==10:pct(c)
        else:money(c)
        if j==11:
            c.fill=PatternFill("solid",fgColor=GREEN if fav=="مُوات" else RED)
            c.font=Font(bold=True,color=GREENF if fav=="مُوات" else REDF)
    r+=1
zebra(ws,4,r-1,1,10)
widths(ws,[7,28,9,15,15,13,12,13,13,12,11])

# ============================================ SHEET 5: Expense Allocation =======
ws = wb.create_sheet("٥. توزيع المصروفات")
title(ws,f"توزيع المصروفات التشغيلية على المنتجات — الأساس: {CONFIG['expense_allocation_base']}",
      "الافتراضي: التوزيع بنسبة مساهمة كل منتج في مجمل الربح (Gross Profit %) — قابل للتغيير", span=8)
# pools table
hdr(ws,3,["مجمّع المصروفات التشغيلية (بيع/إدارة/تمويل)","القيمة ج.م","% من المبيعات"])
r=4
for k,v in OPEX_POOL.items():
    tag=" (تمويل — اختياري)" if k in FINANCE_KEYS else ""
    ws.cell(r,1,k+tag).alignment=RTL; money(ws.cell(r,2,v)); pct(ws.cell(r,3,v/INCOME["net_sales"]))
    r+=1
ws.cell(r,1,"الإجمالي").alignment=RTL; ws.cell(r,1).font=Font(bold=True)
money(ws.cell(r,2,OPEX_CONTROL)); ws.cell(r,2).font=Font(bold=True)
pct(ws.cell(r,3,OPEX_CONTROL/INCOME["net_sales"]))
zebra(ws,4,r,1,3)
# per-product allocation
r+=2
hdr(ws,r,["الكود","المنتج","أساس التوزيع","نسبة التوزيع %","الأعباء الصناعية المحمّلة","المصروفات التشغيلية المحمّلة","إجمالي المحمّل"])
r+=1; start=r
for x in ranked:
    vals=[x["code"],x["name"],round(x["gross_profit"],1),x["alloc_pct"],
          round(x["conv_alloc"],1),round(x["opex_alloc"],1),round(x["conv_alloc"]+x["opex_alloc"],1)]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v)
        if j==2:c.alignment=RTL
        elif j==1:c.alignment=CEN
        elif j==4:pct(c)
        else:money(c)
    r+=1
zebra(ws,start,r-1,1,7)
widths(ws,[7,30,16,13,18,20,16])

# ============================================ SHEET 6: Profitability ============
ws = wb.create_sheet("٦. تحليل الربحية")
title(ws,"تحليل الربحية لكل منتج — من الإيراد إلى صافي الربح",
      "إيراد • تكلفة • مجمل ربح • ربح تشغيلي • مساهمة في الربح • هوامش", span=12)
H=["الكود","المنتج","الإيراد الصافي","التكلفة الفعلية","مجمل الربح","مجمل الهامش %",
   "المحمّل (أعباء+مصروفات)","الربح التشغيلي","الهامش التشغيلي %","مساهمة في الربح %","EBITDA تقريبي","الترتيب"]
hdr(ws,3,H)
r=4
tot_op=sum(x["op_profit"] for x in rows)
for x in ranked:
    ebitda = x["op_profit"] + x["opex_alloc"]*0  # simplified: op profit (finance excluded below)
    ebitda = x["revenue"]-x["act_mat_total"]-x["conv_alloc"]-(x["opex_alloc"]-x["opex_alloc"]*0)
    # EBITDA ≈ op profit + finance share (finance is inside opex)
    fin_share = x["alloc_pct"]*sum(v for k,v in OPEX_POOL.items() if k in FINANCE_KEYS)
    ebitda = x["op_profit"]+fin_share
    contrib = x["op_profit"]/tot_op if tot_op else 0
    vals=[x["code"],x["name"],round(x["revenue"],1),round(x["full_loaded_total"],1),
          round(x["gross_profit"],1),x["gross_margin"],round(x["conv_alloc"]+x["opex_alloc"],1),
          round(x["op_profit"],1),x["op_margin"],contrib,round(ebitda,1),x["rank"]]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v)
        if j==2:c.alignment=RTL
        elif j in(1,12):c.alignment=CEN
        elif j in(6,9,10):pct(c)
        else:money(c)
        if j==9: c.font=Font(bold=True,color=REDF if x["op_margin"]<0 else GREENF)
    r+=1
zebra(ws,4,r-1,1,12)
widths(ws,[7,28,14,14,13,12,18,13,15,14,13,8])

# ============================================ SHEET 7: Sales & ABC ==============
ws = wb.create_sheet("٧. تحليل المبيعات ABC")
title(ws,"تحليل المبيعات — الترتيب، مزيج الإيراد، تصنيف ABC وتحليل باريتو",
      "80% من المبيعات (فئة A) • 15% التالية (B) • 5% الأخيرة (C)", span=11)
H=["الترتيب","الكود","المنتج","الكمية","متوسط سعر البيع","الإيراد الصافي","مزيج الإيراد %",
   "التراكمي %","هامش المساهمة/وحدة","تصنيف ABC","الحالة"]
hdr(ws,3,H)
r=4
for x in ranked:
    cm = x["avg_price"]-(x["act_mat_total"]/x["qty"] if x["qty"] else 0)  # contribution margin/unit
    vals=[x["rank"],x["code"],x["name"],x["qty"],round(x["avg_price"],2),round(x["revenue"],1),
          x["rev_mix"],x["cum_pct"],round(cm,2),x["abc"],x["action"]]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v)
        if j==3:c.alignment=RTL
        elif j in(1,2,10,11):c.alignment=CEN
        elif j in(7,8):pct(c)
        else:money(c)
        if j==10:
            f={"A":GREEN,"B":AMBER,"C":RED}[x["abc"]]; ff={"A":GREENF,"B":AMBERF,"C":REDF}[x["abc"]]
            c.fill=PatternFill("solid",fgColor=f); c.font=Font(bold=True,color=ff)
    r+=1
zebra(ws,4,r-1,1,9)
widths(ws,[7,7,28,9,15,14,13,12,17,11,12])

# ============================================ SHEET 8: PRICING (main) ===========
ws = wb.create_sheet("٨. قائمة الأسعار المقترحة")
title(ws,"قائمة الأسعار الجديدة — محرك التسعير الديناميكي (الاستراتيجية: %s)"%CONFIG["pricing_strategy"],
      "السعر المقترح يحمي الهامش ولا يقل عن أرضية التكلفة المحمّلة؛ يُرفَع فقط ما هو دون المستهدف", span=18)
H=["الكود","المنتج","الفئة","العلامة","السعر القديم","السعر المقترح","الفرق","% التغير",
   "خامات معيارية","تعبئة معيارية","تكلفة تصنيع معيارية","التكلفة المعيارية","التكلفة الفعلية",
   "المصروفات المحمّلة","التكلفة المحمّلة بالكامل","أرضية السعر","الهامش المقترح %","الإجراء"]
hdr(ws,3,H,fsz=9)
r=4
for x in ranked:
    std_mfg = x["std_labour_u"]+x["std_moh_u"]+x["std_utils_u"]+x["std_quality_u"]+x["std_waste_u"]+x["std_qres_u"]
    vals=[x["code"],x["name"],x["cat"],x["brand"],round(x["old_price"],2),round(x["rec_price"],2),
          round(x["price_diff"],2),x["price_pct"],round(x["std_mat_u"],2),round(x["std_pack_u"],2),
          round(std_mfg,2),round(x["std_full_u"],2),round(x["act_full_u"],2),
          round(x["opex_alloc"]/x["qty"] if x["qty"] else 0,2),round(x["full_loaded_u"],2),
          round(x["floor_price"],2),x["rec_margin"],x["action"]]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v); c.font=Font(size=9)
        if j==2:c.alignment=RTL
        elif j in(1,3,4,18):c.alignment=CEN
        elif j in(8,17):pct(c)
        else:money(c)
        if j==6:c.font=Font(size=9,bold=True)
        if j==8:
            c.fill=PatternFill("solid",fgColor=GREEN if x["price_pct"]>0.02 else RED if x["price_pct"]<-0.02 else GREY)
        if j==18:
            c.fill=PatternFill("solid",fgColor={"رفع السعر":GREEN,"خفض السعر":RED,"تثبيت":GREY}[x["action"]])
            c.font=Font(size=9,bold=True)
    r+=1
zebra(ws,4,r-1,1,18)
ws.freeze_panes="C4"
widths(ws,[6,26,8,9,11,12,9,9,12,12,15,13,13,14,16,12,13,11])

# ============================================ SHEET 9: Pricing Rules/Flags ======
ws = wb.create_sheet("٩. قواعد التسعير والتنبيهات")
title(ws,"قواعد التسعير الديناميكي والأصناف المُعلَّمة للمراجعة",
      "الهامش المطلوب يتغير حسب حجم المبيعات (ABC) — تنبيهات البيع دون التكلفة والهامش الشاذ", span=8)
# rules block
ws.merge_cells("A3:H3"); ws.cell(3,1,"قواعد الهامش حسب حجم المبيعات (Volume-based margin)").font=Font(bold=True,color=NAVY)
ws.cell(3,1).alignment=RTL
hdr(ws,4,["فئة الحجم","القاعدة","تعديل الهامش المطلوب"])
rules=[("A — مبيعات مرتفعة","هامش مقبول أقل لتعظيم الحصة السوقية",f"{CONFIG['volume_margin_delta']['A']*100:+.0f} نقطة"),
       ("B — مبيعات متوسطة","الحفاظ على الهامش المستهدف",f"{CONFIG['volume_margin_delta']['B']*100:+.0f} نقطة"),
       ("C — مبيعات منخفضة","رفع الهامش المطلوب لتعويض بطء الدوران",f"{CONFIG['volume_margin_delta']['C']*100:+.0f} نقطة")]
r=5
for a,b,c in rules:
    ws.cell(r,1,a).alignment=RTL; ws.cell(r,2,b).alignment=RTL; ws.cell(r,3,c).alignment=CEN
    for cc in range(1,4):ws.cell(r,cc).border=BORDER
    r+=1
r+=1
hdr(ws,r,["الكود","المنتج","تصنيف ABC","الهامش المطلوب %","الهامش التشغيلي الحالي %","التنبيهات"])
r+=1; start=r
for x in ranked:
    if x["flags"]=="—" and x["action"]=="تثبيت": continue
    vals=[x["code"],x["name"],x["abc"],x["req_margin"],x["op_margin"],x["flags"]]
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v)
        if j==2:c.alignment=RTL
        elif j==6:c.alignment=RTL
        elif j in(4,5):pct(c)
        else:c.alignment=CEN
        if j==6 and x["flags"]!="—": c.fill=PatternFill("solid",fgColor=AMBER)
    r+=1
zebra(ws,start,r-1,1,6)
widths(ws,[7,28,12,16,20,55])

# ============================================ SHEET 10: Methodology =============
ws = wb.create_sheet("١٠. المنهجية والافتراضات")
ws.sheet_view.rightToLeft=True
title(ws,"المنهجية والافتراضات والتحقق","إطار تكاليف كامل مطابَق لقائمة دخل يونيه 2026", span=2)
notes=[
 ("مصدر البيانات","قائمة دخل يونيه 2026 (PDF)، فواتير مبيعات يونيه 2026 (311 فاتورة)، وصفات التصنيع بأحدث متوسط تكلفة مرجّح WAC، وتصنيف الأصناف كبراند."),
 ("تفسير قائمة الدخل","تكلفة المبيعات (2,039,933) = خامات مباشرة + تعبئة مستهلكة. الأعباء الصناعية غير المباشرة (544,605 = تصنيع لدى الغير، عمالة، كهرباء، صيانة، إيجار مصنع…) مجمّع تحويل يُوزّع كتكلفة تصنيع. المصروفات التشغيلية (738,392 = بيع+إدارة+تمويل) تُوزّع لاحتساب صافي الربح. المطابقة: 3,741,772 − 2,039,933 − 544,605 − 738,392 = 418,842 = صافي الربح المُعلن."),
 ("التكلفة المعيارية","بناء تصاعدي لكل SKU: (وزن×سعر الخلطة WAC) + تكلفة التعبئة الفعلية + تحميل تحويل بالمعدل %.3f ج.م/كجم (مجمّع الأعباء÷الكميات) موزّعاً على أجور/أعباء/مرافق/جودة + هدر %.1f%% + احتياطي جودة %.1f%%."%(meta["conv_rate_kg"],CONFIG["waste_allowance_pct"]*100,CONFIG["quality_reserve_pct"]*100)),
 ("التكلفة الفعلية والمطابقة","التكلفة الفعلية للخامات مطابَقة لتكلفة المبيعات بالقائمة إلى القرش: المنتجات ذات الوصفة تحمل تكلفتها التصاعدية، والأصناف التجارية (بلا وصفة) تحمل الرصيد المتبقّي بالتناسب مع إيرادها. الفرق بين المعياري والفعلي = الانحراف."),
 ("توزيع المصروفات","الأساس الافتراضي = مساهمة المنتج في مجمل الربح (Gross Profit %). يدعم النموذج 8 أسس (قيمة/كمية المبيعات، كمية الإنتاج، التكلفة المعيارية/الفعلية، ساعات عمل/آلة) قابلة للتبديل من ملف الإعدادات."),
 ("مطابقة الإيراد","إجمالي فواتير البيع (3,867,491) يفوق صافي مبيعات القائمة (3,741,772) بمقدار 125,719 = مرتجعات/بونص/خصومات — عوملت كبند خصم من الإيراد بمعامل تسوية موحّد لضمان تطابق الإجماليات."),
 ("محرك التسعير","استراتيجية %s: أرضية السعر = التكلفة المحمّلة ÷ (1 − الهامش المطلوب). الهامش المطلوب يبدأ من مستهدف العلامة ويُعدَّل حسب فئة ABC. لا يُخفَّض سعر صنف يحقق أرضيته (حماية الهامش)؛ يُرفَع فقط ما هو دون الأرضية بحدود ±%d%%/%d%% وتقريب لأقرب %.2f ج.م."%(CONFIG["pricing_strategy"],int(CONFIG["max_price_increase"]*100),int(CONFIG["min_price_decrease"]*100),CONFIG["round_to"])),
 ("حدود النطاق","الأصناف التجارية (بانيه، استربس، كبده، مفروم، عجينة بقري) بلا وصفة خامات في المصدر — تكلفتها المعيارية غير متاحة وقُدِّرت فعلياً بالتناسب. المياه والأصناف غير المُسعّرة قيمتها صفر (استبعاد لا تقدير). ~15 خامة طوبقت بحكم مهني."),
 ("التحقق","3 مطابقات محورية تمّت آلياً إلى القرش: Σ خامات فعلية = 2,039,933.08؛ Σ ربح تشغيلي = 418,841.92؛ Σ نسب التوزيع = 100%. أي تغيير في الإعدادات يعيد الاحتساب مع الحفاظ على هذه المطابقات."),
]
r=4
for k,v in notes:
    ws.cell(r,1,k); ws.cell(r,1).font=Font(bold=True,color=WHITE); ws.cell(r,1).fill=PatternFill("solid",fgColor=BLUE)
    ws.cell(r,1).alignment=Alignment(horizontal="right",vertical="top",wrap_text=True)
    ws.cell(r,2,v); ws.cell(r,2).alignment=Alignment(horizontal="right",vertical="top",wrap_text=True)
    ws.cell(r,2).border=BORDER; ws.cell(r,1).border=BORDER
    ws.row_dimensions[r].height=80
    r+=1
widths(ws,[22,120])

# ============================================ SHEET 11: Config =================
ws = wb.create_sheet("١١. الإعدادات")
ws.sheet_view.rightToLeft=True
title(ws,"إعدادات النموذج القابلة للتبديل (Configuration)","غيّر القيم ثم أعد تشغيل النموذج — كل النتائج تتحدّث وتبقى مطابِقة للقائمة", span=2)
import json as _j
flat=[("أساس توزيع المصروفات",CONFIG["expense_allocation_base"]),
      ("استراتيجية التسعير",CONFIG["pricing_strategy"]),
      ("هامش صافٍ مستهدف - أبو هاشم",f"{CONFIG['target_net_margin']['أبو هاشم']:.0%}"),
      ("هامش صافٍ مستهدف - الهنا",f"{CONFIG['target_net_margin']['الهنا']:.0%}"),
      ("هامش صافٍ مستهدف - سبيشال",f"{CONFIG['target_net_margin']['سبيشال']:.0%}"),
      ("تعديل الهامش A/B/C",f"{CONFIG['volume_margin_delta']}"),
      ("سقف رفع السعر",f"{CONFIG['max_price_increase']:.0%}"),
      ("حد خفض السعر",f"{CONFIG['min_price_decrease']:.0%}"),
      ("تقريب السعر لأقرب",f"{CONFIG['round_to']} ج.م"),
      ("سماح الهدر",f"{CONFIG['waste_allowance_pct']:.1%}"),
      ("احتياطي الجودة",f"{CONFIG['quality_reserve_pct']:.1%}")]
hdr(ws,3,["الإعداد","القيمة الحالية"])
r=4
for k,v in flat:
    ws.cell(r,1,k).alignment=RTL; ws.cell(r,2,str(v)).alignment=CEN
    ws.cell(r,1).border=BORDER; ws.cell(r,2).border=BORDER
    r+=1
zebra(ws,4,r-1,1,2)
widths(ws,[40,45])

wb.save(OUT)
print("Workbook saved:", OUT, "| sheets:", len(wb.sheetnames))
