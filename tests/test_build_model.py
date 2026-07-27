"""Tests for the cost-per-km math behind scripts/build_model.py.

build_model.py emits an Excel workbook in which every computed value is a *formula
string* (e.g. "=F5/D5"), evaluated later by LibreOffice — so there are no Python
functions returning numbers to assert on directly. What IS in Python, and what this
suite pins, are the model's authoritative input constants (`cities`, `veh_data`) and
the exact per-km formula the sheets apply to them:

    cost/km = fuel_price / (mileage * (1 - mileage_drop))      # petrol, E20 drop
    cost/km = fuel_price / mileage                             # diesel / CNG / EV
    blended EV price = elec_home * home_share + elec_public * (1 - home_share)

The default lever values (E20 drop 4%, home-charge share 80%, public tariff Rs 21/kWh)
are the ones build_model writes into its Assumptions sheet. Expected per-km figures
below were hand-computed from the Delhi row and cross-checked against the recalculated
workbook's cached values.
"""
import os

import pytest

# Model default levers set in build_model.py's Assumptions sheet (inp() calls).
E20_DROP = 0.04
HOME_SHARE = 0.80
ELEC_PUBLIC = 21.0


@pytest.fixture(scope="module")
def bm(tmp_path_factory):
    """Import build_model with cwd redirected so its `wb.save('outputs/..')` writes
    into a throwaway tmp dir instead of mutating the repo's real output file."""
    work = tmp_path_factory.mktemp("bm_run")
    (work / "outputs").mkdir()
    prev = os.getcwd()
    os.chdir(work)
    try:
        import build_model  # noqa: E402  (import triggers the workbook build + save)
        return build_model
    finally:
        os.chdir(prev)


@pytest.fixture(scope="module")
def delhi(bm):
    # ("Delhi", petrol Rs/L, diesel Rs/L, CNG Rs/kg, elec home, petrol VAT, diesel VAT)
    return next(c for c in bm.cities if c[0] == "Delhi")


def _veh(bm, name):
    # ("name", petrol km/L, diesel km/L, CNG km/kg, EV km/kWh, note)
    return next(v for v in bm.veh_data if v[0] == name)


def test_data_shapes_are_sane(bm, delhi):
    assert delhi[1] > 0 and delhi[3] > 0            # petrol & CNG prices positive
    hatch = _veh(bm, "Hatchback")
    assert hatch[1] == 16 and hatch[3] == 28 and hatch[4] == 7.5
    assert hatch[2] is None                          # hatchback has no diesel variant


def test_e20_petrol_cost_per_km_delhi_hatchback(bm, delhi):
    hatch = _veh(bm, "Hatchback")
    price, mileage = delhi[1], hatch[1]              # 102.12, 16
    cpk = price / (mileage * (1 - E20_DROP))
    assert cpk == pytest.approx(6.6484, abs=1e-3)    # matches cached CostPerKm value


def test_diesel_cost_per_km_delhi_compact_suv(bm, delhi):
    suv = _veh(bm, "Compact SUV")
    price, mileage = delhi[2], suv[2]                # 95.20, 19
    cpk = price / mileage
    assert cpk == pytest.approx(5.0105, abs=1e-3)


def test_cng_cost_per_km_delhi_hatchback(bm, delhi):
    hatch = _veh(bm, "Hatchback")
    price, mileage = delhi[3], hatch[3]              # 83.09, 28
    cpk = price / mileage
    assert cpk == pytest.approx(2.9675, abs=1e-3)


def test_ev_blended_price_and_cost_per_km_delhi_hatchback(bm, delhi):
    hatch = _veh(bm, "Hatchback")
    elec_home = delhi[4]                             # 8.0
    blended = elec_home * HOME_SHARE + ELEC_PUBLIC * (1 - HOME_SHARE)
    assert blended == pytest.approx(10.6, abs=1e-9)  # 8*0.8 + 21*0.2
    cpk = blended / hatch[4]                          # / 7.5
    assert cpk == pytest.approx(1.4133, abs=1e-3)


def test_running_cost_ranking_ev_cheapest_then_cng_then_petrol(bm, delhi):
    """README headline: EV cheapest, CNG next, E20 petrol dearest (hatchback)."""
    hatch = _veh(bm, "Hatchback")
    petrol = delhi[1] / (hatch[1] * (1 - E20_DROP))
    cng = delhi[3] / hatch[3]
    ev = (delhi[4] * HOME_SHARE + ELEC_PUBLIC * (1 - HOME_SHARE)) / hatch[4]
    assert ev < cng < petrol


def test_diesel_beats_petrol_for_compact_suv(bm, delhi):
    suv = _veh(bm, "Compact SUV")
    petrol = delhi[1] / (suv[1] * (1 - E20_DROP))    # 102.12 / (13*0.96)
    diesel = delhi[2] / suv[2]                        # 95.20 / 19
    assert diesel < petrol
