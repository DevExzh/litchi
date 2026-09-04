# Change 0408 evidence

See [the change record](../../changes/0408-opc-materialization-evidence.md).
`performance_claim: none`; `claim_authorized: false`.

`capture.py` records exact commands in `commands.json`. The capture inventory
is preserved as `capture-artifacts.json`. `artifact-manifest.json` binds the
published files, including checks and summary. The original 22 MiB symbolized
`perf-script.stdout` is retained only as lossless `.zst`; its uncompressed
size/hash remains recorded in both manifests. Run `python3 verify-artifacts.py`
from this directory to verify published files and compressed payloads.

Normal, allocator and accounted reports have distinct instrumentation scopes.
The L1 aliases returned zeroes of unvalidated viability; LLC aliases reported
not supported. Required commands passed; command success does not establish
counter viability. Read `summary.json` and the change record before interpreting
the raw counter output. Captured binaries were not archived; their identities,
build settings and IDs are retained with fully symbolized CPU artifacts.

`checks/` preserves initial failures as well as successful final gates. Initial
harness selector count and allocator failure-test requests were corrected; the
broader dependency Clippy warnings remain open. No full-workspace gate or broad
goal completion is claimed.
