from agent_preflight_guard import verify_scale_mathematics, calculate_margin_bands, preflight_dimension_chain

def test_scale():
    ok, _ = verify_scale_mathematics(100.0, 250.0)
    assert ok

def test_bands():
    bands = calculate_margin_bands(1000.0, 100.0, 1.0, 'left')
    assert bands['band_1_local'] == 0.0

def test_chain():
    ok, _ = preflight_dimension_chain([0.0, 100.0, 500.0], 0.0, 100.0, 1)
    assert ok
