# 0096 — Immutable XLSX edit-entry promotion

The XLSX edit seam previously consumed the handle and selected a source or
eager editor based on reference counts. Its shared-handle conversion used
ordinary OPC materialization, losing exact source framing. That violates the
immutable edit contract and the exact-source promotion required by records
0009 D4 and 0015 D7.

The replacement borrows the handle and returns the native workbook snapshot.
Source promotion uses the existing preserving-source OPC conversion, so it
reuses the prepared index and retains the source archive rather than
serializing/reopening a reconstructed package. A successful promoted native
owner is shared across clones of that open. Source checks still run when the
promoted owner is already available. Failed promotions remain retryable.
Signature authorization and retained-source policy are admitted before edit
entry. Signature removal is explicit and cannot contaminate signed aliases.

The same preserving conversion fixes eager adoption of an already prepared
source. A public regression checks exact ZIP bytes and retained path origin
through `open_prepared_with` and publication.

This is the native owner foundation for XLSX cell writes, not a claim that
`File::set` or batches are integrated. Managed materialization remains a typed
leaf refusal; the broader managed-open gap from 0095 remains open. Other
format writes, Python, CLI, performance and acceptance requirements remain
part of the active goal. No timing improvement is claimed.

The full native workbook is an owner promotion under M2 D7, not an ADR 0032
small derived-value cache. Ordinary and removal-authorized owners have separate
success-only slots; both retain the original immutable snapshot. A mutex
serializes each first promotion, and failures do not initialize the slot.

An editor snapshot is detached from the positional input. The future unified
write adapter must fence the complete admission/edit/commit operation, not just
editor construction. That requirement is explicit at the internal edit entry;
no new public mutation support is advertised by this record.

Validation includes a transient `ReadAt` I/O failure followed by a successful
retry with unchanged bytes and source identity; eight concurrent first
promotions sharing the same native snapshot; clone/selected-sheet reuse;
cached promotion refusing a changed source; explicit unsigned owner reuse
while signed aliases still refuse; and prepared eager exact publication.
The eager leaf intentionally defers worksheet semantic decoding; malformed
unselected worksheet XML is not an appropriate promotion-failure oracle.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx,xlsx,encryption --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Results: 561 XLSX/encryption tests/doctests and 817 combined DOCX/XLSX/encryption
tests/doctests passed. The final concurrent test with a start barrier also
passed. All 14 capability-delta tests and the generated-delta check passed.
Read-only review found no production blocker. Previously promoted snapshots
may remain retained after the physical source becomes stale; every edit-entry
and publication path still fences that source and refuses its reuse.

Warning-denied Clippy, formatting, diff hygiene and rewritten-base ancestry
checks passed. The changed-file artifact audit found no forbidden large
evidence or executables; both required staged checks printed nothing. Git
staging remains blocked by the read-only common Git directory (`index.lock`).
Temporary logs, generated root lockfile and build outputs are removed after
verification. No temporary worktree was created.
