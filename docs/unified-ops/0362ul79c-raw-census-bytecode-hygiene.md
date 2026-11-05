# 0362ul79c — raw-census bytecode hygiene

## Review repair

The validator's own local imports could create ignored `tools/__pycache__`
files before it inspected source overlays. That would make a clean frozen
capture fail because of the checker itself. `raw_api_coverage.py` now sets
`sys.dont_write_bytecode = True` before importing local glossary/resolver
modules. The f5 capture policy and every per-profile receipt also require
`PYTHONDONTWRITEBYTECODE=1`; this is an identity/control field, not an
allowlist exception for ignored pycache trees.

## Source-only check

The regression copies the validator and its local modules into a fresh source
directory, launches the direct CLI with the parent `PYTHONDONTWRITEBYTECODE`
setting removed (bytecode enabled), and verifies that the CLI succeeds without
creating a `tools/__pycache__`. The receipt fixture carries the external
target, trace, tools, and this expected environment field.

```text
python3 -m py_compile tools/raw_api_coverage.py tools/test_raw_api_coverage.py
exit 0
python3 -m unittest tools.test_raw_api_coverage
Ran 17 tests in 0.928s ... OK
python3 tools/raw_api_coverage.py --root . \
  --manifest tools/raw_api_coverage.json \
  --glossary tools/api_glossary.json --json
exit 0; complete=false; source_failures=[]; current_source_failures=[];
missing_receipts=15; missing_inputs=15; missing glossary profiles=53;
missing_workspace_packages=0
```

No Cargo, rustdoc, native build, or performance measurement was run. The
manifest remains incomplete; final profile and external capture identities are
rebuilt after source freeze.
