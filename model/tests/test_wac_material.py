"""وصفات الخامة الواحدة فى تقييم المنتج التام (تأكيد الإدارة 26/10/2026)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import wac  # noqa: E402


def test_rules_name_real_finished_goods():
    names = {wac.key(f["name"]) for f in wac.run("2026-09")["fg"]}
    for n in wac.MATERIAL_RULES:
        assert wac.key(n) in names, n


def test_issued_months_keep_their_prices():
    for m in ("2026-08", "2026-09"):
        assert wac.material_rule(m, "مفروم ابو هاشم 350") is None
        fg = {wac.key(f["name"]): f for f in wac.run(m)["fg"]}
        assert fg[wac.key("مفروم ابو هاشم 350")]["price"] == 84.61
        assert fg[wac.key("كبده اسكندراني 370")]["price"] == 64.1


def test_recipe_price_from_the_month_end_average():
    R = wac.run("2026-09")
    prices = {k: x["wac"] for k, x in R["items"].items()}
    meat = wac.material_rule("2026-10", "مفروم ابو هاشم  350")
    liver = wac.material_rule("2026-10", "كبده اسكندرانى 370")
    assert wac.material_price(meat, prices, R["pool"]["wac"]) == pytest.approx(0.32 * 260.2402 + 6.25, abs=1e-4)
    assert wac.material_price(liver, prices, R["pool"]["wac"]) == pytest.approx(0.34 * 168.6305 + 6.25, abs=1e-4)
    assert wac.material_price(liver, {}, None) is None
