"""وصفات المفروم والكبدة (تأكيد الإدارة 26/10/2026).

«مفروم صافي 350» (= «مفروم صافي 400 جم» فى المبيعات) = 0.32 كجم لحم هندى + 6.25 تغليف،
«كبده اسكندراني 400» = 0.34 كجم كبده + 6.25، «كيلو مفروم هندي» = 1 كجم لحم هندى + 2.00.
المعيار = الوصفة بمتوسط سعر آخر الشهر؛ يونيه من «تسعير جرد اخر شهر ٦ سنة ٢٠٢٦».
"""
import contextlib
import io
import sys
from pathlib import Path

import openpyxl
import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
with contextlib.redirect_stdout(io.StringIO()):
    import costing_pricing_model as M  # noqa: E402

JUNE = HERE.parents[1] / "تسعير جرد اخر شهر ٦ سنة ٢٠٢٦.xlsx"


@pytest.fixture(scope="module")
def rows():
    return {r["code"]: r for r in M.build()[0]}


def test_prices_are_the_june_month_end_average():
    if not JUNE.exists():
        pytest.skip("تسعير جرد اخر شهر ٦ غير موجود")
    ws = openpyxl.load_workbook(JUNE, data_only=True).active
    price = {str(r[0].value).strip(): r[7].value for r in ws.iter_rows() if r[0].value}
    assert M.MIX["لحم هندي"] == pytest.approx(price["لحم كتف"])   # one price for the Indian-meat pool
    assert M.MIX["كبده"] == pytest.approx(price["كبده"])


def test_recipe_cost_per_unit(rows):
    for code, kg, recipe, pack in ((1, 0.32, "لحم هندي", 6.25), (2, 0.34, "كبده", 6.25),
                                   (33, 1.0, "لحم هندي", 2.0)):
        r = rows[code]
        assert r["has_bom"] and r["recipe"] == recipe
        assert r["std_mat_u"] == pytest.approx(kg * M.MIX[recipe])
        assert r["std_pack_u"] == pack
        assert r["act_mat_total"] == pytest.approx((kg * M.MIX[recipe] + pack) * r["qty"])
    assert rows[1]["std_mat_u"] + rows[1]["std_pack_u"] == pytest.approx(87.21)
    assert rows[2]["std_mat_u"] + rows[2]["std_pack_u"] == pytest.approx(63.877, abs=1e-3)


def test_brands_unchanged(rows):
    assert rows[1]["brand"] == rows[2]["brand"] == "أبو هاشم"
    assert rows[33]["brand"] == "أخرى (تجاري)"
