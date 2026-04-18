from fsg_bom.matcher import AssemblyMatcher

def test_resolve_label_exact():
    matcher = AssemblyMatcher()
    options = ["Calipers", "Brake Discs"]
    assert matcher.resolve_label("Calipers", options) == "Calipers"

def test_resolve_label_remap():
    matcher = AssemblyMatcher()
    options = ["Calipers", "Brake Master Cylinder"]
    assert matcher.resolve_label("brake caliper", options) == "Calipers"
    assert matcher.resolve_label("reservoir", options) == "Brake Master Cylinder"

def test_resolve_label_fuzzy():
    matcher = AssemblyMatcher()
    options = ["Gearbox", "Wheels"]
    assert matcher.resolve_label("gear box", options) == "Gearbox"

def test_resolve_label_substring():
    matcher = AssemblyMatcher()
    options = ["Front Uprights", "Rear Uprights"]
    assert matcher.resolve_label("Uprights", options) == "Front Uprights" # First match

def test_system_label():
    matcher = AssemblyMatcher()
    assert matcher.get_system_label("BR") == "BR - Brake System"
    assert matcher.get_system_label("UNKNOWN") == "UNKNOWN"
