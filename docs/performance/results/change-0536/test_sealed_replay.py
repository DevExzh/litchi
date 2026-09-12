"""Check sealed report output boundaries using only temporary external files."""
import json
from pathlib import Path
import tempfile
import analyze_profiles
import instruction_analysis

results = []
with tempfile.TemporaryDirectory(prefix="litchi-0536-replay-test-", dir="/home/zhuhe") as folder:
    root = Path(folder)
    bundle = root / "bundle"
    bundle.mkdir()
    (bundle / "SHA256SUMS").write_text("")
    for module in (analyze_profiles, instruction_analysis):
        previous = module.HERE
        module.HERE = bundle
        try:
            outside = root / (module.__name__ + ".json")
            module.write_report({"probe": True}, outside)
            module.write_report({"probe": True}, outside)
            assert json.loads(outside.read_text()) == {"probe": True}
            for value, path in [({"probe": False}, outside), ({"probe": True}, bundle / "missing.json")]:
                try:
                    module.write_report(value, path)
                except module.EvidenceError:
                    pass
                else:
                    raise AssertionError("unsafe report write admitted")
            assert not (bundle / "missing.json").exists()
            results.append({"module": module.__name__, "external_exact_replay": "pass", "nonidentical_overwrite_refused": True, "sealed_internal_creation_refused": True})
        finally:
            module.HERE = previous
print(json.dumps({"status": "pass", "checks": results}, indent=2))
