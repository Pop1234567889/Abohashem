"""الهانك فى نموذج التكاليف (تأكيد الإدارة 26/10/2026).

كل كيلو سجق من أى علامة عليه 1/15 هانك، والمعيار = الوصفة بمتوسط آخر الشهر، فيونيه بمتوسط «هانك»
فى «تسعير جرد اخر شهر ٦» (250). MIX سعر عجينة فقط للسجق الثلاثة (أبو هاشم على كميات وصفته 127.12 كجم)،
فيُضاف الهانك لعبوات السجق وحدها، لا للعجينة التى تشترك فى الوصفة.
"""
import contextlib
import io
import json
import sys
from pathlib import Path

import openpyxl
import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
with contextlib.redirect_stdout(io.StringIO()):
    import costing_pricing_model as M  # noqa: E402

STANDARD = HERE.parents[1] / "سعر انتاج تام-1.xlsx"
JUNE = HERE.parents[1] / "تسعير جرد اخر شهر ٦ سنة ٢٠٢٦.xlsx"


@pytest.fixture(scope="module")
def rows():
    return {r["code"]: r for r in M.build()[0]}


def _casing_cell(ws, label, header=None):
    """The number right of `label` (under `header` when given) in a recipe sheet."""
    cells = [c for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    col = None
    if header:
        col = next(c.column for c in cells if c.value.strip() == header)
    lab = next(c for c in cells if c.value.strip() == label and (col is None or c.column == col))
    return next(float(ws.cell(lab.row, c).value) for c in range(lab.column + 1, lab.column + 4)
                if isinstance(ws.cell(lab.row, c).value, (int, float)))


def test_casing_is_june_average_for_every_brand():
    if not JUNE.exists():
        pytest.skip("تسعير جرد اخر شهر ٦ غير موجود")
    ws = openpyxl.load_workbook(JUNE, data_only=True).active
    june = next(r[7].value for r in ws.iter_rows() if r[0].value and str(r[0].value).strip() == "هانك")
    assert M.HANK_JUNE == june
    assert M.CASING_PER_KG == {r: june / 15 for r in ("سجق الهنا", "سجق اسبشيال", "سجق ابوهاشم")}


def test_standard_workbook_casing_quantity():
    """The workbook's casing per kilo is 1/15 of its casing price (الهنا 20 = 300/15, اسبشيال 250/15)."""
    if not STANDARD.exists():
        pytest.skip("سعر انتاج تام-1.xlsx غير موجود")
    wb = openpyxl.load_workbook(STANDARD, data_only=True)
    assert _casing_cell(wb["الهنا"], "هانك") == pytest.approx(300 / 15)
    assert _casing_cell(wb["اسبشيال 2026"], "هانك") == pytest.approx(250 / 15)
    assert _casing_cell(wb["الهاشم"], "هنكات للكيلو", "تكلفه سجق ابوهاشم") == 20


def test_abu_hashem_paste_over_its_recipe_kilos():
    # June materials were 91.0834 × 98.96 kg; the recipe's own kilos are 127.12
    assert M.MIX["سجق ابوهاشم"] == pytest.approx(91.0834 * 98.9600314296844 / 127.12, abs=1e-4)


def test_casing_goes_on_sausage_packs_only(rows):
    c = 250 / 15
    expected = {257: 5 * c, 108: 0.3 * c, 99: 0.7 * c, 405: 0.3 * c, 406: 0.7 * c, 13: 3 * c, 107: 0.8 * c, 7: 0.35 * c}
    for code, r in rows.items():
        assert r["std_casing_u"] == pytest.approx(expected.get(code, 0.0)), (code, r["name"])
    assert rows[292]["cat"] == "عجينة" and rows[292]["std_casing_u"] == 0      # paste shares the recipe
    for code in expected:
        r = rows[code]
        assert r["std_mat_u"] == pytest.approx(r["weight"] * M.MIX[r["recipe"]] + r["std_casing_u"])


def test_statement_still_ties_to_the_cent(rows):
    assert sum(r["act_mat_total"] for r in rows.values()) == pytest.approx(M.INCOME["cogs"], abs=0.01)
    assert sum(r["gross_profit"] for r in rows.values()) == pytest.approx(1_701_838.92, abs=0.01)
    assert sum(r["op_profit"] for r in rows.values()) == pytest.approx(418_841.92, abs=0.01)


def test_committed_model_rows_are_current(rows):
    committed = json.loads((HERE.parent / "model_rows.json").read_text(encoding="utf-8"))
    for code, r in rows.items():
        c = committed[str(code)]
        for k in ("std_casing_u", "std_mat_u", "act_mat_total", "gross_profit", "op_profit"):
            assert c[k] == pytest.approx(round(r[k], 4), abs=1e-4), (code, k)
