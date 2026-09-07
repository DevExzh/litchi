# Integration notes

The original Rust source epoch is 36df468bd. The two normal comparison binaries
use identical release settings; one source line differs. The candidate test
suite passes 368 cases, including attribute-cache/fresh-scan differential checks
and native-file preservation. The source manifest delta is exactly
`crates/litchi-odp/src/codec/parser/codec/xml/validation.rs`; `experiment.patch`
retains the tested change. The predeclared matrix is A1/B1/B2/A2, on CPU 2 with
one workload at a time. Baseline R2 executes its retained baseline binary while
the worktree still contains candidate source; the binary binding is authoritative
for measured code, while the check receipt separately records worktree custody.

The original diagnostic classifier used Counter insertion order to break equal
sample-period ties. Set iteration could make this unstable across processes.
Before bundle verification, the derived helper was corrected to sort ties by
symbol. `classify-initial.py`, `diagnostic-summary-initial.json`, and the first
passing classification receipt remain retained. `classify-r1` supplies the final
summary. Frozen binary/protocol/profile inputs and raw measurements are unchanged.

All 24 formal lanes and both counter captures pass their report oracles. The
candidate fails the latency keep gate and has no allocation reduction. Six
adverse timing flags are retained, including R2 allocator-large p50 +43.039% and
normal-medium p50 +13.327%. The data does not establish the cause of this
variation; it is not discarded or explained away. No additional confirmation
run is needed to reject the experiment. The single changed production line is
restored exactly to the baseline source, with its digest in `decision.json`.

The final source is unchanged production code. Post-restoration formatting,
strict owner lint, docs and crate boundaries validate that state. The existing
ODP suite covers the tested experiment and native fixtures; no Office GUI was
launched. The ODF fuzz crate currently targets detection and ODT parsing, not
this ODP lookup. No unrelated fuzz run is claimed as coverage for it.

Raw profiler/assembly/log whitespace is authenticated evidence and is retained.
Finalization verifies retained executable custody before inventory-based cleanup
of only `/tmp/litchi-goal-0459`, including its separate diagnostic build tree.
Shared Cargo targets and user-owned `docs/GOAL.md` are preserved.

The broad `cargo fmt --all -- --check` attempt failed on an existing formatting
difference at `crates/litchi-keynote/src/document.rs:592`. Its receipt and log
are retained. iWork is explicitly outside this task; that file is unchanged.
The scoped `cargo fmt -p litchi-odp -- --check` replacement passes as
`format-r1`. This is an owner-format result, not a whole-workspace format pass.

Final integration found the adapted frozen capture driver retained numeric
`change: 458` and omitted a separate variant field in the 24 capture receipts.
Their schema is 0459, and directories, argv, build/source and binary-binding
hashes identify each variant exactly. `capture-metadata-amendment.json` binds
all 24 original receipt hashes plus the unchanged driver/protocol. The derived
verifier admits only this disclosed metadata discrepancy under that amendment.
No frozen input, binary or raw measurement/receipt is rewritten.

The first full derived-verifier run failed because default gate patterns were
tuples while the validator accepted lists. `verify-initial.py` and
`verification-attempt-initial.json` retain that helper/result. Converting
default tuples to lists fixes this integration error without changing gates
or measurement artifacts.

Final source-delta verification also authenticates complete retained baseline
and candidate source text against their compiled manifests, then proves the
exact predicate substitution produces the candidate bytes. The earlier passing
negative-summary probe remains in `negative-verification-initial.json`; the
final probe additionally binds the strengthened verifier hash.
