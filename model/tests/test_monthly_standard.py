"""المعيار الشهرى = الوصفة بمتوسط آخر الشهر (تأكيد الإدارة 26/10/2026)."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import monthly_standard as S  # noqa: E402
import wac  # noqa: E402


@pytest.mark.parametrize("month", ["2026-08", "2026-09"])
def test_standard(month):
    s = S.build(month)
    kg = wac.run(month)["kg"]
    hank = s["hank_price"] * S.HANK_PER_KG
    for name, r in s["recipes"].items():
        assert r["casing_kg"] == pytest.approx(hank if name.startswith("سجق") else 0.0)
        assert r["kg"] == pytest.approx(r["paste_kg"] + r["casing_kg"])
        if name != "سجق ابوهاشم":        # same divisor as the valuation: paste = its per-kg price
            assert r["paste_kg"] == pytest.approx(kg[name], abs=1e-4)  # سجق اسبشيال: 228.65888 vs Σ 228.659
    ah = s["recipes"]["سجق ابوهاشم"]
    assert ah["batch_kg"] == pytest.approx(127.12)
    assert ah["paste_kg"] == pytest.approx(kg["سجق ابوهاشم"] * 98.9600314296844 / 127.12, abs=1e-5)
    pool, liver = s["pool"], next(x["wac"] for x in wac.run(month)["items"].values() if x["name"] == "كبده")
    assert s["single"]["مفروم صافي 350"]["std"] == pytest.approx(0.32 * pool + 6.25, abs=1e-4)
    assert s["single"]["كبده اسكندراني 400"]["std"] == pytest.approx(0.34 * liver + 6.25, abs=1e-4)
    assert s["single"]["كيلو مفروم هندي"]["std"] == pytest.approx(pool + 2.0, abs=1e-4)


@pytest.mark.parametrize("month", ["2026-08", "2026-09"])
def test_committed_file_is_current(month):
    committed = json.loads((wac.month_dir(month) / "standard.json").read_text(encoding="utf-8"))
    assert committed == json.loads(json.dumps(S.build(month), ensure_ascii=False))
