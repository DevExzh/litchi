# XLSX retained-history paired adapter

The retained-history allocator route is `tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs`.
It is one source file for both revisions. The default build selects the W4
deferred owner; the `legacy-xlsx-source-history` feature selects the pre-W4
`workbook::source_merge` owner. Both branches use the same generated fixtures,
alternating merge/unmerge steps, replay and inverse byte checks, patch count,
and allocator snapshot interval.

The canonical source is recorded with SHA-256
`8a188086459202a5014f25ce1e8d15bb296c3f2f902e31d1f7a398d96988647c` after the
final source edit and before the release build. Regenerate the temporary
baseline adapter from that source as follows; the archived baseline checkout
stays outside the repository:

```sh
cp tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs \
  /tmp/litchi-w4-baseline/tools/perf-baseline/src/bin/
python3 - /tmp/litchi-w4-baseline/tools/perf-baseline/Cargo.toml <<'PY'
from pathlib import Path
import sys

path = Path(sys.argv[1])
text = path.read_text()
text = text.replace(
    'xlsb-crud = ["litchi/xlsb"]\n',
    'xlsb-crud = ["litchi/xlsb"]\nlegacy-xlsx-source-history = []\n',
    1,
)
text = text.replace(
    '[[bin]]\nname = "litchi-perf-baseline-alloc"\n',
    '[[bin]]\nname = "xlsx_merge_history_alloc"\n'
    'path = "src/bin/xlsx_merge_history_alloc.rs"\n'
    'required-features = ["allocator-metrics", "legacy-xlsx-source-history"]\n\n'
    '[[bin]]\nname = "litchi-perf-baseline-alloc"\n',
    1,
)
path.write_text(text)
PY
sha256sum /tmp/litchi-w4-baseline/tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs
```

The baseline command enables both `allocator-metrics` and
`legacy-xlsx-source-history`; the candidate enables only `allocator-metrics`.
The measured counters are live bytes, high-water live bytes, allocation and
deallocation calls, reallocation calls, and allocated/deallocated bytes over
the interval that retains every patch. The W4 deferred publisher shares the
change record's `Arc` with the retained patch, so each retained commit has no
second closure allocation for its output endpoint. Exact replay and inverse
checks still stream through the recipient owner and its limits. For an
unobserved deferred after endpoint, the patch binds the recipient to the exact
before artifact and retains an immutable publisher whose catalog/order,
metadata, timestamps, save options and replacement bytes are all captured by
the change record. That deterministic publisher's single recipient-policy
stream is the authenticated after artifact; hashing the same archive in a
second prepass would duplicate the splice. Opened or independently observed
after endpoints retain their artifact fingerprint and use the exact restore
proof. Inverse replay authenticates its own before endpoint by the same rule.
The recipient source freshness check remains in the streaming write path, so a
source change during output is reported rather than silently accepted. RSS is
supplemental only.

The report has separate retained-history, verification, common replay, and
owner-admission allocator intervals. `verification_allocator` is the
complete caller workflow: every step opens an independent equal source-backed
recipient, replays forward and inverse, and serializes both outputs. Its W4
call chain is `SourceBackedWorkbook::from_read_at` →
`open_source_backed_package` → catalog and sheet-graph admission → deferred
replay/publication. The legacy adapter uses `OpcPackage::from_bytes` and its
materialized patch API, so that interval is retained as an unmatched workflow
result; it does not construct the new read-only facade. The report also records
`source_owner_admission_allocator`, which repeats source owner construction
and selected worksheet admission inside the measured interval for each
revision, and `common_replay_allocator`, which constructs both recipients
before the interval and measures the shared forward/inverse replay and
publication work. The latter keeps the exact replay and inverse byte oracles
while removing the unequal owner-admission cost from the common route.
`common_replay_apply_forward_allocator` and
`common_replay_apply_inverse_allocator` split the same recipient-only apply
work before serialization, so deferred authentication allocations in inverse
replay are attributable separately from publication and output buffering.

An earlier probe was rejected: it independently opened a source owner for each
retained step and materialized every commit output while measuring history.
Its candidate live and peak bytes were therefore contaminated by the probe's
per-step owners. The final adapter reuses the two fixture owners for commit
history and retains only the patch records; its replay oracles remain outside
the retained commit output. The rejected JSON remains outside the repository
at `/tmp/litchi-w4-history-abba-20261021.json` for audit continuity.
