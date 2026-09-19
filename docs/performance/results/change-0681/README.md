# Change 0681 evidence audit

This packet audits retained evidence and current selector availability. It
contains no new benchmark run and registers no performance claim.

- `0672-recomputed.json` independently recomputes p50 reductions from all eight
  retained 30-sample files, carrying their SHA-256 hashes and the audit revision.
  Positive percentages mean lower latency, as in the original `stats.py`.
- `selector-audit.json` records twelve exact selector names found in the harness
  registry and its source hash. This proves presence only; it does not replace
  the CRUD index or establish measured coverage.
- `xls-duplicate-refusal-review.json` preserves the source control-flow excerpt
  showing why a latest-occurrence-only XLS cache would change target errors.
- `gates/` records the applicable evidence, coverage and boundary checks.

Reproduce the timing check from the repository root:

```sh
sha256sum -c docs/performance/results/change-0672/timing/runs-sha256.txt
python3 docs/performance/results/change-0672/stats.py docs/performance/results/change-0672/timing/runs floor-poi abba-poi
```

The median uses Python's `statistics.median`; each pair's reduction is
`100 * (before - after) / before`. Original samples and result packets were not
modified. Neither this recomputation nor a positive warm median establishes a
tail, cold-cache, RSS, cross-platform or general throughput improvement.

Recheck the selector inventory without rebuilding:

```sh
python3 - <<'PY'
from pathlib import Path
import hashlib
import json
p = Path('docs/performance/results/change-0681/selector-audit.json')
audit = json.loads(p.read_text())
source = Path(audit['source']).read_bytes()
assert hashlib.sha256(source).hexdigest() == audit['source_sha256']
lines = source.decode().splitlines()
for row in audit['selectors']:
    observed = [i for i, line in enumerate(lines, 1)
                if '"' + row['selector'] + '"' in line]
    assert observed == row['name_occurrence_lines'], row['selector']
print('12 selector names and source digest match')
PY
```

Final review distinguished XLS cell-error values from operation errors:
out-of-range SST indices return `Ok(CellValue::Error(...))` and may be
overwritten by later duplicates; locator/read/decode failures still abort.
The design now requires both witnesses. It also distinguishes worksheet-only
replay work from dependency work and logical clean-index weight from live
managed reservations. The DOCX pressure/trim seam was already present in the
final author revision; no additional lifecycle claim was inferred.
