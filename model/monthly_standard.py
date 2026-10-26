#!/usr/bin/env python3
"""المعيار الشهرى: خلطة التصنيع بمتوسط سعر الخامات آخر الشهر (تأكيد الإدارة 26/10/2026).

لكل وصفة من الوصفات التسع (monthly/reference/recipes.xlsx — نفس كميات أوراق الوصفات فى «سعر انتاج تام»):
    سعر كجم العجينة = قيمة الوصفة بمتوسط الشهر ÷ إجمالى كميات الوصفة (بالمياه)، كما تحسبه أوراق الوصفات
    السجق: + الهانك = 0.0666666667 هانك للكيلو × متوسط «هانك» فى الشهر
ووصفات الخامة الواحدة: مفروم صافي 350 = 0.32 كجم لحم هندى + 6.25، كبده اسكندراني = 0.34 كجم كبده + 6.25،
كيلو مفروم هندي = 1 كجم لحم هندى + 2.00.

ملف يونيه يقسم سجق أبو هاشم على 98.96 كجم (مستنتج من 90.28 بالملف = عجينة 70.28 + هانك 20) لا على كميات
الوصفة 127.12؛ هنا القسمة على كميات الوصفة والهانك بند مستقل، كبقية العلامات. wac.py (تقييم المخزون) لم يتغير.

    python3 model/monthly_standard.py 2026-09     يكتب monthly/2026-09/standard.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import wac  # noqa: E402
from wac_august_2026 import POOL, key, recipe_cost  # noqa: E402

HANK_PER_KG = 0.0666666667
HANK_ITEM = "هانك"
SINGLE = {  # اسم الوصفة: (الخامة، كجم للعبوة، تعبئة)
    "مفروم صافي 350": (POOL, 0.32, 6.25),
    "كبده اسكندراني 400": ("كبده", 0.34, 6.25),
    "كيلو مفروم هندي": (POOL, 1.0, 2.0),
}


def build(month: str) -> dict:
    R = wac.run(month)
    if R.get("bootstrap"):
        raise SystemExit(f"{month}: شهر البداية بلا متوسط محسوب — المعيار الشهرى يبدأ من الشهر التالى له")
    prices = {k: (x["open_p"] if x["unresolved"] else x["wac"]) for k, x in R["items"].items()}
    pool = R["pool"]["wac"]

    def price_of(item):
        return pool if item == POOL else prices.get(key(item))

    hank = price_of(HANK_ITEM)
    assert hank, f"{month}: «{HANK_ITEM}» بلا متوسط"
    recipes = {}
    for name, rc in wac.recipes().items():
        value, _, _ = recipe_cost(rc, price_of)
        batch = sum(i["qty"] for i in rc["main"])
        paste = value / batch
        casing = HANK_PER_KG * hank if name.startswith("سجق") else 0.0
        unpriced = [i["ing"] for i in rc["main"] + rc["sub"]
                    if i["qty"] and i["item"] not in (None, "@SUB") and not isinstance(i["item"], tuple)
                    and not price_of(i["item"])]
        recipes[name] = {"paste_kg": round(paste, 6), "casing_kg": round(casing, 6), "kg": round(paste + casing, 6),
                         "batch_kg": round(batch, 4), "value": round(value, 4), "unpriced": unpriced}
    single = {}
    for name, (item, kg, pack) in SINGLE.items():
        p = price_of(item)
        single[name] = {"item": item.lstrip("@"), "kg": kg, "price": p, "pack": pack,
                        "std": None if p is None else round(kg * p + pack, 4)}
    return {"month": month, "basis": "خلطة التصنيع بمتوسط سعر الخامات آخر الشهر (تأكيد الإدارة 26/10/2026)",
            "hank_price": hank, "hank_per_kg": HANK_PER_KG, "pool": pool, "recipes": recipes, "single": single}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("month", nargs="+")
    for m in ap.parse_args(argv).month:
        data = build(m)
        p = wac.month_dir(m) / "standard.json"
        p.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        print(f"{m}: هانك {data['hank_price']:.4f} | لحم هندى {data['pool']:.4f} → {p.relative_to(wac.ROOT)}")
        for n, r in data["recipes"].items():
            print(f"  {n:14} عجينة {r['paste_kg']:9.4f} + هانك {r['casing_kg']:7.4f} = {r['kg']:9.4f}/كجم")
        for n, r in data["single"].items():
            print(f"  {n:18} {r['std']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
