# -*- coding: utf-8 -*-
"""
ABOHASHEM FOR MEAT — Product Costing & Pricing Model (June 2026)
=================================================================
A full standard/actual costing, expense-allocation, profitability, sales
and dynamic-pricing engine, reconciled to the June-2026 income statement.

Author: Cost Accounting / FP&A build
Inputs (all in-repo):
  - WAC recipe costs .......... ABOHASHEM_...يونيه2026.xlsx  (3 brand sheets)
  - Line-item sales ........... فواتير_المبيعات_يونيو_2026.md (311 invoices)
  - Income statement .......... قائمة دخل تفصيلى شهر ٦ سنة ٢٠٢٦.pdf
  - Brand classification ...... تصنيف الاصناف كبراند.pdf

Output:
  - ABOHASHEM_Costing_Pricing_Model_June2026.xlsx  (12 sheets)
  - COSTING_PRICING_REPORT.md                      (narrative report)

Everything below the CONFIG block is deterministic and reconciles to the
statement to the cent, so any assumption is auditable and swappable.
"""
import os, re, json
from collections import defaultdict, OrderedDict

# ============================================================= CONFIG ==========
CONFIG = {
    # Default allocation base for OPERATING expenses (req. 5).
    # options: gross_profit | sales_value | sales_qty | production_qty |
    #          standard_cost | actual_cost | labor_hours | machine_hours
    "expense_allocation_base": "gross_profit",

    # Pricing strategy (req. 8):
    # cost_plus | target_gross_margin | target_net_margin |
    # contribution_optimization | competitive | hybrid
    "pricing_strategy": "hybrid",

    # Target margins by brand (on fully-loaded cost) — Chief Cost Accountant set
    "target_net_margin": {"أبو هاشم": 0.14, "الهنا": 0.12, "سبيشال": 0.10, "أخرى (تجاري)": 0.11},

    # Volume-based margin rules (req. 9): ABC class -> margin delta (pp)
    "volume_margin_delta": {"A": -0.02, "B": 0.00, "C": +0.03},

    # Guard rails
    "max_price_increase": 0.15,   # cap a single-cycle increase at +15%
    "min_price_decrease": -0.10,  # floor a single-cycle decrease at -10%
    "round_to": 0.25,             # round recommended prices to nearest 0.25 EGP

    # Standard production-loss / waste / quality allowances (% of prime cost)
    "waste_allowance_pct": 0.010,   # 1.0% standard waste
    "quality_reserve_pct": 0.005,   # 0.5% standard quality reserve
}

CUR = "ج.م"  # EGP

# ===================================================== 1. INCOME STATEMENT =====
# From: قائمة دخل تفصيلى شهر ٦ سنة ٢٠٢٦.pdf  (June-2026, fully reconciled)
INCOME = {
    "net_sales": 3_741_772.00,
    "cogs": 2_039_933.08,                # = direct materials + packaging consumed
    "gross_profit": 1_701_838.92,
}
# Manufacturing conversion overhead (period pool, block 1 = 544,605 exactly) —
# req.2/5. NB: the "مرتبات الإنتاج 45,716" line in the PDF is a memo already
# absorbed in COGS/labour; excluded here so the pool ties to the 544,605 control.
CONVERSION_POOL = OrderedDict([
    ("تصنيع لدى الغير (Outsourced mfg)",   295_530.00),
    ("عمالة خارجية (External labour)",       97_460.00),
    ("ايجارات المصنع (Factory rent)",        66_600.00),
    ("فواتير الكهرباء (Factory utilities)",  40_000.00),
    ("صيانات وقطع غيار (Maintenance)",       17_765.00),
    ("بدلات ومكافآت (Allowances)",           17_255.00),
    ("منظفات ومكافحة (Cleaning / QC)",        4_230.00),
    ("أخرى إنتاجية (Other)",                   3_765.00),
    ("غفرة رمضان (Ramadan cover)",             2_000.00),
])
CONVERSION_CONTROL = 544_605.00  # block-1 subtotal from the income statement
# Operating-expense pools (blocks 2-4) to allocate for NET profit — req.5
OPEX_POOL = OrderedDict([
    # --- Administrative (block 2 = 162,928) ---
    ("مرتبات إدارية (Admin salaries)",        125_800.00),
    ("مصاريف المكاتب (Office)",                  8_585.00),
    ("هدايا (Gifts / relations)",                6_300.00),
    ("محاماة (Legal)",                           5_000.00),
    ("مكتب المحاسبة (Accounting)",               5_050.00),
    ("كهرباء ونت إداري (Admin utilities)",       4_563.00),
    ("ايجارات إدارية (Admin rent)",              4_000.00),
    ("تأمينات وحكومي (Insurance/Gov)",           2_000.00),
    ("ضيافات (Hospitality)",                     1_630.00),
    # --- Selling & Distribution (block 3 = 383,445) ---
    ("عمولة البيع + بدلات (Sales commission)",  76_846.00),
    ("عربية جامبو (Jumbo truck)",               88_207.00),
    ("مرتبات المبيعات (Sales salaries)",        61_400.00),
    ("مركبات توزيع أخرى (Other vehicles)",      93_057.00),  # 24430+12510+17450+18117+20450
    ("زيت وفلاتر (Oil & filters)",              21_875.00),
    ("ايجار عربية (Vehicle rent)",              15_000.00),
    ("فنادق (Hotels)",                           7_200.00),
    ("مكتب المبيعات (Sales office)",             4_000.00),
    ("دخان المبيعات (Sales misc.)",              1_860.00),
    ("كباس تير (Tyres)",                        14_100.00),  # part of vehicle block
    # --- Finance & other (block 4 = 192,019) ---
    ("تأمين المصنع والعربيات (Insurance)",      96_424.00),
    ("قرض المبادرة (Initiative loan)",          95_595.00),  # optional finance cost
    ("قرض السيارة (Car loan)",                  14_085.00),  # optional finance cost
    ("مصاريف الإشهار (Advertising)",             3_450.00),
])
FINANCE_KEYS = {"قرض المبادرة (Initiative loan)", "قرض السيارة (Car loan)"}
# Reported operating-expense control = admin(162,928)+selling(383,445)+finance(192,019)
OPEX_CONTROL = 162_928.00 + 383_445.00 + 192_019.00   # = 738,392
# add a reclassification line so the detailed pool ties exactly to the control
OPEX_POOL["تسوية فروق تصنيف (Reclass. adj.)"] = OPEX_CONTROL - sum(OPEX_POOL.values())

# ============================================== 2. WAC RECIPE COST (per kg) =====
# Direct-material WAC cost per kg of blended mix — from the 3 brand sheets.
MIX = {
 "سجق ابوهاشم":91.0834, "كفته ابوهاشم":88.3292, "برجر ابوهاشم":78.1394,
 "برجر اسبشيال":24.2895, "سجق اسبشيال":27.7778, "كفته اسبشيال":30.8696,
 "سجق الهنا":23.8061,    "برجر الهنا":24.6792,   "كفته الهنا":24.2895,
}
BRAND_OF_RECIPE = {r: ("أبو هاشم" if "ابوهاشم" in r else "الهنا" if "الهنا" in r else "سبيشال")
                   for r in MIX}

# Single-material recipes (management 26/10/2026): «مفروم صافي 350» = 0.32 kg Indian meat
# + 6.25 packaging, «كبده اسكندراني 400» = 0.34 kg liver + 6.25, «كيلو مفروم هندي» = 1 kg
# Indian meat + 2.00. The standard is the recipe at the month-end WAC, so June uses
# «تسعير جرد اخر شهر ٦ سنة ٢٠٢٦»: Indian-meat pool («تشكيله هندى امامى») 253.00, «كبده» 169.4903.
MIX["لحم هندي"] = 253.0
MIX["كبده"] = 169.4903
# Material kilos per unit where they differ from the pack weight in SKU.
MAT_KG = {1: 0.32, 2: 0.34}

# Casing (هانك) on sausage packs — management 26/10/2026: every kilo of sausage of any brand
# carries 1/15 casing, and the standard is the recipe at the month-end average, so June prices
# it at June's «هانك» average (250.00 in «تسعير جرد اخر شهر ٦»). MIX is paste only for all
# three sausages: سجق ابوهاشم was 91.0834 = June materials ÷ 98.96 kg, a divisor back-solved
# from the standard file's 90.28 (paste 70.28 + casing 20 over 127.12 kg); it is now the
# materials over the recipe's own 127.12 kg, like الهنا and اسبشيال, with the casing apart.
MIX["سجق ابوهاشم"] = round(91.0834 * 98.9600314296844 / 127.12, 4)   # 70.9064
HANK_JUNE = 250.0
CASING_PER_KG = {r: HANK_JUNE / 15 for r in ("سجق الهنا", "سجق اسبشيال", "سجق ابوهاشم")}
# Kilos of sausage the standard costs the casing on, where that differs from the paste
# weight in SKU: «سجق الهنا 300/700» in the standard are the 350/800 packs sold.
CASING_KG = {108: 0.3, 99: 0.7}


def casing_u(code, recipe, w, cat):
    """Casing cost per unit; paste («عجينة») shares the sausage recipe but carries none."""
    if cat != "سجق" or recipe not in CASING_PER_KG:
        return 0.0
    return CASING_PER_KG[recipe] * CASING_KG.get(code, w)

# ============================================ 3. SOLD-SKU MASTER (June 2026) ====
# code : (name, recipe|None, weight_kg, packaging_cost, category)
# recipe=None => traded/other product with no BOM in source data (documented gap)
SKU = OrderedDict([
 (13 ,("سجق شرقي 3 ك ابو هاشم","سجق ابوهاشم",3.0 ,9.23 ,"سجق")),
 (101,("كفته الهنا 800 جرام","كفته الهنا",0.8 ,6.08 ,"كفته")),
 (120,("برجر الهنا 900 جرام","برجر الهنا",0.9 ,6.63 ,"برجر")),
 (402,("كفته اسبيشيال عائلي","كفته اسبشيال",0.9 ,7.2 ,"كفته")),
 (404,("برجر اسبيشيال عائلي","برجر اسبشيال",1.0 ,6.7 ,"برجر")),
 (346,("بانيه 800 جرام",None,0.8 ,0.0 ,"بانيه")),
 (102,("كفته الهنا 350","كفته الهنا",0.35,5.175,"كفته")),
 (398,("بانيه 900 جرام اسباشيل",None,0.9 ,0.0 ,"بانيه")),
 (439,("استربس الهنا 800 جرام",None,0.8 ,0.0 ,"استربس")),
 (104,("برجر الهنا 400 جرام","برجر الهنا",0.4 ,4.22 ,"برجر")),
 (257,("سجق الهنا 5 كيلو","سجق الهنا",5.0 ,16.06,"سجق")),
 (108,("سجق الهنا 350","سجق الهنا",0.35,4.63 ,"سجق")),
 (403,("برجر اسبيشيال صغير","برجر اسبشيال",0.4 ,5.5 ,"برجر")),
 (399,("كفته اسبيشيال صغير","كفته اسبشيال",0.35,5.5 ,"كفته")),
 (435,("عجينة بقري 5 ك",None,5.0 ,0.0 ,"عجينة")),
 (1  ,("مفروم صافي 400 جم (ابو هاشم)","لحم هندي",0.4 ,6.25 ,"مفروم")),
 (441,("استريس اسبشيال 900 جرام",None,0.9 ,0.0 ,"استربس")),
 (2  ,("كبده اسكندراني 400 (ابو هاشم)","كبده",0.4 ,6.25 ,"كبده")),
 (80 ,("برجر 1200 جم","برجر ابوهاشم",1.2 ,6.18 ,"برجر")),
 (99 ,("سجق الهنا 800 جرام","سجق الهنا",0.8 ,5.99 ,"سجق")),
 (255,("برجر ابو هاشم جامبو 1.5 كيلو","برجر ابوهاشم",1.5 ,6.07 ,"برجر")),
 (107,("سجق 800 ابو هاشم","سجق ابوهاشم",0.8 ,6.27 ,"سجق")),
 (405,("سجق اسبيشيال صغير","سجق اسبشيال",0.3 ,5.2 ,"سجق")),
 (406,("سجق اسبيشيال عائلى","سجق اسبشيال",0.7 ,7.2 ,"سجق")),
 (420,("برجر الهنا 800 جرام","برجر الهنا",0.8 ,6.63 ,"برجر")),
 (421,("برجر ابو هاشم 1 ك","برجر ابوهاشم",1.0 ,5.6 ,"برجر")),
 (114,("كفته ابو هاشم 800","كفته ابوهاشم",0.8 ,5.35 ,"كفته")),
 (292,("عجينه سجق 5 كيلو الهنا","سجق الهنا",5.0 ,16.06,"عجينة")),
 (7  ,("سجق 350 جم ابو هاشم","سجق ابوهاشم",0.35,4.67 ,"سجق")),
 (33 ,("كيلو مفروم هندي","لحم هندي",1.0 ,2.0 ,"مفروم")),
 (8  ,("كفته 350 جم ابو هاشم","كفته ابوهاشم",0.35,4.5 ,"كفته")),
 (447,("استربس الهنا 900 جرام",None,0.9 ,0.0 ,"استربس")),
 (440,("استربس الهنا 1 كيلو",None,1.0 ,0.0 ,"استربس")),
])

def brand_of(recipe, name):
    if recipe in BRAND_OF_RECIPE: return BRAND_OF_RECIPE[recipe]
    if "اسبشيال" in name or "اسباشيل" in name or "اسبيشيال" in name or "اسبشيل" in name: return "سبيشال"
    if "الهنا" in name: return "الهنا"
    if "ابو هاشم" in name or "ابوهاشم" in name: return "أبو هاشم"
    return "أخرى (تجاري)"

# ================================================ 4. PARSE SALES INVOICES =======
def parse_sales(path):
    agg = defaultdict(lambda: {"qty":0.0,"val":0.0,"lines":0,"prices":[]})
    row_re = re.compile(r'^\|\s*\d+\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|'
                        r'\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|')
    inv = 0
    for ln in open(path, encoding="utf-8"):
        if ln.startswith("## فاتورة"): inv += 1
        m = row_re.match(ln)
        if not m: continue
        code,name,unit,qty,up,tax,disc,total = [x.strip() for x in m.groups()]
        if not re.match(r'^\d', code): continue
        def num(s):
            try: return float(s.replace(",","").strip())
            except: return None
        q,t,p = num(qty),num(total),num(up)
        if q is None or t is None: continue
        a = agg[int(code)]
        a["qty"]+=q; a["val"]+=t; a["lines"]+=1
        if p: a["prices"].append(p)
    return agg, inv

# =================================================== 5. BUILD PRODUCT MODEL ======
def build():
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sales, n_inv = parse_sales(os.path.join(base,"فواتير_المبيعات_يونيو_2026.md"))

    rows = []
    # ---- standard prime cost (materials + packaging) per SKU ----
    for code,(name,recipe,w,pkg,cat) in SKU.items():
        s = sales.get(code, {"qty":0,"val":0,"lines":0,"prices":[]})
        qty, rev = s["qty"], s["val"]
        avg_price = rev/qty if qty else 0.0        # qty-weighted realized price
        list_price = max(s["prices"]) if s["prices"] else avg_price
        old_price = avg_price                      # basis = realized avg (consistent w/ margin)
        std_cas  = casing_u(code, recipe, w, cat)
        std_mat  = (MAT_KG.get(code, w)*MIX[recipe] + std_cas) if recipe else None      # per unit, casing included
        std_pack = pkg if recipe else 0.0
        rows.append(dict(code=code, name=name, recipe=recipe, cat=cat,
                         brand=brand_of(recipe,name), weight=w, qty=qty, kg=w*qty,
                         revenue=rev, avg_price=avg_price, old_price=old_price, list_price=list_price,
                         std_mat_u=std_mat, std_casing_u=std_cas, std_pack_u=std_pack,
                         has_bom=recipe is not None))

    # ---- reconcile invoice gross sales to statement NET sales ----
    # Invoice gross sales exceed IS net sales by returns/bonus/allowances.
    invoice_gross = sum(r["revenue"] for r in rows) or 1
    ret_factor = INCOME["net_sales"] / invoice_gross      # <1 : contra-revenue
    for r in rows:
        r["gross_revenue"] = r["revenue"]
        r["revenue"]       = r["revenue"] * ret_factor    # net revenue (ties to IS)
        r["returns"]       = r["gross_revenue"] - r["revenue"]

    # ---- reconcile ACTUAL material COGS (2,039,933) to products ----
    # recipe products carry their bottom-up std material COGS; the residual COGS
    # is spread over TRADED products by revenue (they have no BOM). Guarantees
    # Σ actual material = reported COGS exactly.
    std_cogs_recipe = sum(r["std_mat_u"]*r["qty"] + r["std_pack_u"]*r["qty"]
                          for r in rows if r["has_bom"])
    traded_rev = sum(r["revenue"] for r in rows if not r["has_bom"])
    traded_cogs = INCOME["cogs"] - std_cogs_recipe
    for r in rows:
        if r["has_bom"]:
            r["act_mat_total"] = r["std_mat_u"]*r["qty"] + r["std_pack_u"]*r["qty"]
        else:  # imputed material cost for traded goods
            share = r["revenue"]/traded_rev if traded_rev else 0
            r["act_mat_total"] = traded_cogs*share
            u = r["act_mat_total"]/r["qty"] if r["qty"] else 0
            r["std_mat_u"], r["std_pack_u"] = u, 0.0  # single imputed line
    tot_kg = sum(r["kg"] for r in rows) or 1

    # ---- standard CONVERSION add-on per kg (block-1 pool ÷ kg produced) ----
    conv_total = sum(CONVERSION_POOL.values())
    conv_rate_kg = conv_total / tot_kg
    # split conversion into the requested standard elements (per kg)
    def pool(keyfrag): return sum(v for k,v in CONVERSION_POOL.items() if keyfrag in k)
    elem = {
        "labour":  (pool("عمالة")+pool("مرتبات الإنتاج"))/tot_kg,
        "utils":   pool("الكهرباء")/tot_kg,
        "quality": pool("منظفات")/tot_kg,
        "moh":     (conv_total-pool("عمالة")-pool("مرتبات الإنتاج")
                    -pool("الكهرباء")-pool("منظفات"))/tot_kg,
    }

    # ---- per-unit standard full manufacturing cost ----
    for r in rows:
        r["std_labour_u"]  = elem["labour"]*r["weight"]
        r["std_utils_u"]   = elem["utils"]*r["weight"]
        r["std_quality_u"] = elem["quality"]*r["weight"]
        r["std_moh_u"]     = elem["moh"]*r["weight"]
        prime = (r["std_mat_u"] or 0)+r["std_pack_u"]
        r["std_waste_u"]   = prime*CONFIG["waste_allowance_pct"]
        r["std_qres_u"]    = prime*CONFIG["quality_reserve_pct"]
        r["std_full_u"]    = (prime + r["std_labour_u"]+r["std_utils_u"]
                              + r["std_quality_u"]+r["std_moh_u"]
                              + r["std_waste_u"]+r["std_qres_u"])
        r["std_full_total"]= r["std_full_u"]*r["qty"]
        # actual full = actual material + actual conversion (by kg share)
        r["act_conv_total"]= conv_rate_kg*r["kg"]
        r["act_full_total"]= r["act_mat_total"]+r["act_conv_total"]
        r["act_full_u"]    = r["act_full_total"]/r["qty"] if r["qty"] else 0
        # gross profit (statement basis: revenue - material COGS)
        r["gross_profit"]  = r["revenue"] - r["act_mat_total"]
        r["gross_margin"]  = r["gross_profit"]/r["revenue"] if r["revenue"] else 0

    # ---- OPERATING-expense allocation (req.5) ----
    bases = {
        "gross_profit": {r["code"]: max(r["gross_profit"],0) for r in rows},
        "sales_value":  {r["code"]: r["revenue"] for r in rows},
        "sales_qty":    {r["code"]: r["qty"] for r in rows},
        "production_qty":{r["code"]: r["kg"] for r in rows},
        "standard_cost":{r["code"]: r["std_full_total"] for r in rows},
        "actual_cost":  {r["code"]: r["act_full_total"] for r in rows},
        "labor_hours":  {r["code"]: r["kg"] for r in rows},   # proxy: kg
        "machine_hours":{r["code"]: r["kg"] for r in rows},   # proxy: kg
    }
    base_map = bases[CONFIG["expense_allocation_base"]]
    base_sum = sum(base_map.values()) or 1
    opex_total = sum(OPEX_POOL.values())
    for r in rows:
        r["alloc_pct"]  = base_map[r["code"]]/base_sum
        r["opex_alloc"] = opex_total*r["alloc_pct"]
        r["conv_alloc"] = r["act_conv_total"]   # manufacturing conversion share
        r["full_loaded_total"] = r["act_mat_total"]+r["conv_alloc"]+r["opex_alloc"]
        r["full_loaded_u"] = r["full_loaded_total"]/r["qty"] if r["qty"] else 0
        r["op_profit"]  = r["revenue"] - r["full_loaded_total"]
        r["op_margin"]  = r["op_profit"]/r["revenue"] if r["revenue"] else 0

    # ---- SALES analysis: ABC / Pareto (req.7) ----
    ranked = sorted(rows, key=lambda r:-r["revenue"])
    cum=0
    for i,r in enumerate(ranked,1):
        cum += r["revenue"]; r["rank"]=i
        r["cum_pct"]=cum/INCOME["net_sales"] if INCOME["net_sales"] else 0
        r["rev_mix"]=r["revenue"]/INCOME["net_sales"] if INCOME["net_sales"] else 0
        r["abc"]="A" if r["cum_pct"]<=0.80 else ("B" if r["cum_pct"]<=0.95 else "C")

    # ---- PRICING engine (req.8/9) ----
    for r in rows:
        recommend_price(r)
        r["action"] = ("رفع السعر" if r["price_pct"] > 0.02 else
                       "خفض السعر" if r["price_pct"] < -0.02 else "تثبيت")

    meta = dict(n_inv=n_inv, tot_kg=tot_kg, conv_rate_kg=conv_rate_kg,
                conv_total=conv_total, opex_total=opex_total, elem=elem,
                std_cogs_recipe=std_cogs_recipe, traded_cogs=traded_cogs)
    return rows, meta

# =================================================== 6. PRICING ENGINE ==========
def recommend_price(r):
    """Margin-protecting dynamic pricing.

    The required net margin is a FLOOR, adjusted by sales volume (ABC):
      A (high volume) -> accept a lower floor (more competitive)
      C (low volume)  -> demand a higher floor
    We never cut the price of a product that already clears its floor
    (that would destroy earned margin); we only lift under-priced / loss SKUs
    up to their floor, and flag excessively-high margins for review.
    """
    brand = r["brand"]
    req = CONFIG["target_net_margin"].get(brand, 0.12)
    req += CONFIG["volume_margin_delta"].get(r.get("abc","B"), 0.0)  # volume rule
    req = min(max(req, 0.05), 0.40)
    r["req_margin"] = req
    cost = r["full_loaded_u"]                 # fully-loaded unit cost
    old  = r["old_price"] or r["avg_price"]
    floor = cost/(1-req) if req < 1 else cost # min price to hit required margin
    strat = CONFIG["pricing_strategy"]

    if strat == "cost_plus":
        price = cost*(1+req)
    elif strat == "target_gross_margin":
        mat_u = (r["act_mat_total"]/r["qty"]) if r["qty"] else cost
        price = mat_u/(1-min(req+0.30, 0.70))
    elif strat == "target_net_margin":
        price = floor
    elif strat == "competitive":
        price = max(old, cost*1.02)           # hold market, never below cost
    elif strat == "contribution_optimization":
        # push A-items on volume (hold), lift weak margins toward floor
        price = old if old >= floor else 0.5*(old+floor)
    else:  # hybrid (default): protect margin, lift only below-floor SKUs
        if old >= floor:
            price = old                       # already clears target -> hold
        else:
            price = floor                     # under-priced -> lift to floor

    # guard rails relative to current price
    if old:
        hi = old*(1+CONFIG["max_price_increase"])
        lo = old*(1+CONFIG["min_price_decrease"])
        # allow lifting a below-cost SKU beyond the +cap up to its floor
        if price > hi and old < cost:
            price = min(price, max(hi, floor))
        else:
            price = min(price, hi)
        price = max(price, lo)
    price = max(price, cost*1.01)             # never below fully-loaded cost
    step = CONFIG["round_to"]
    price = round(price/step)*step
    r["floor_price"] = floor
    r["rec_price"]  = price
    r["price_diff"] = price-(old or 0)
    r["price_pct"]  = (price/old-1) if old else 0
    # recommended margin at the new price
    r["rec_margin"] = (price-cost)/price if price else 0
    # flags (req.9)
    flags=[]
    if old and old < r["act_full_u"]:        flags.append("بيع < تكلفة فعلية")
    if old and old < r["full_loaded_u"]:     flags.append("بيع < تكلفة محمّلة")
    if old and r["has_bom"] and old < r["std_full_u"]: flags.append("بيع < تكلفة معيارية")
    if r["op_margin"] < req-0.03:             flags.append("هامش دون المستهدف")
    if r["op_margin"] > 0.40:                 flags.append("هامش مرتفع جداً — مراجعة")
    r["flags"]="؛ ".join(flags) if flags else "—"

if __name__ == "__main__":
    rows, meta = build()
    print("Invoices parsed:", meta["n_inv"], "| SKUs:", len(rows))
    print("Total kg:", round(meta["tot_kg"],1),
          "| conversion/kg:", round(meta["conv_rate_kg"],3))
    print("Std material COGS (recipe):", round(meta["std_cogs_recipe"],2),
          "| imputed traded COGS:", round(meta["traded_cogs"],2))
    # reconciliation
    act_mat = sum(r["act_mat_total"] for r in rows)
    op = sum(r["op_profit"] for r in rows)
    print("Σ actual material  = %.2f  (COGS %.2f)"%(act_mat, INCOME["cogs"]))
    print("Σ operating profit = %.2f  (IS net profit 418,841.92)"%op)
    json.dump({str(r["code"]):{k:(round(v,4) if isinstance(v,float) else v)
              for k,v in r.items() if k!="prices"} for r in rows},
              open(os.path.join(os.path.dirname(__file__),"model_rows.json"),"w"),
              ensure_ascii=False, indent=1)
