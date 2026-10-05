# 0917 postcapture reduction repair

2026-10-05 — frozen capture SHA-256 `fde5219539cf538d51b8bf92ec5527fd9959ac0f39eb2ea5412b8ddf8cd7a894`; one in-memory correction in the postcapture adapter: `for base_item, (_, item) in zip(baseline_rows, rows, strict=True):` → `for (_, base_item), (_, item) in zip(baseline_rows, rows, strict=True):`.

`analysis.py` asserts that SHA, substitutes exactly one occurrence without changing the frozen runner, and invokes the unchanged reduction/replay algorithms. No recapture, protocol, manifest, raw evidence, or statistic changed.
