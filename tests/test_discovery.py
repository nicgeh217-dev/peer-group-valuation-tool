"""Tests for auto-discovery — constituent screening by GICS sector."""
import discovery


def test_constituents_loaded():
    cons = discovery.load_constituents()
    assert len(cons) > 100          # S&P 500 should have ~500
    assert all("ticker" in c and "sector" in c for c in cons)


def test_sectors_are_gics():
    secs = discovery.list_gics_sectors()
    # Core GICS sectors must be present
    assert "Information Technology" in secs
    assert "Health Care" in secs
    assert "Financials" in secs


def test_screen_filters_by_sector():
    it = discovery.screen(sector="Information Technology")
    assert len(it) > 20             # IT sector is large
    assert "AAPL" in it             # Apple is IT
    # A health-care name must NOT be in the IT screen
    hc = discovery.screen(sector="Health Care")
    assert set(it).isdisjoint(set(hc))  # no overlap between distinct sectors


def test_screen_limit():
    capped = discovery.screen(sector="Financials", limit=5)
    assert len(capped) == 5


def test_screen_unknown_sector_empty():
    assert discovery.screen(sector="Nonexistent Sector") == []
