# 0550 source-backed XLSX commit attribution

The preceding turn made progress: 0549 was committed, its restored baseline
passed final checks and strict replay, and its owned target was removed.
This campaign selects the next implementation from fresh OOXML attribution.
It changes no production or harness Rust. OLE2/OOXML remain the priority; ODF
is deferred until that goal completes and iWork is excluded.

The frozen plan measures ordinary source-backed one-cell and one-percent
edit/save on medium, dense-sparse, noncompact, and vendor-extension corpora.
Every child checks the existing source/output, untouched-member, lifecycle,
publication and semantic oracles. Source manifests include tracked fixture
bytes and the harness. Prior accepted ADR hashes are revalidated unchanged.
All root-owned build/capture/check children run serially. Shared-host activity
is not controlled and is not a basis for a host-isolation claim.

Two native repeats each use 3 warmups and 30 samples. Their phase vectors are
reported descriptively, including every absolute repeat change above 5%.
These sample counts cannot support a new registered latency claim. The native
commit region includes edit staging; the workflow adds open, planning, commit,
and publication while the final verification/reopen phase remains separate.
The two allocator repeats use the existing separate allocator binary and the
same matrix/sample counts. Allocation regions are not exact function profiles.
Instrumented elapsed is excluded from native analysis. Missing copied bytes,
route counters, hardware events and cache/scaling data remain unavailable.

Eight profiles use zero warmups and one measured iteration each. Collection
starts disabled, toggles only the exact `MultiSourceEdit::commit` owner, resets
before each invocation, and dumps after each invocation. The analyzer must
classify lifecycle calls separately using positive incoming caller evidence;
only the unique measured runner-to-owner invocation is attributed. The final
termination dump must be checked. Inclusive descendants overlap and cannot be
summed into an opportunity size. Commit-internal temporary destruction remains
included; external staging, publication, reopen and returned-commit drop are
excluded from exact-owner Ir.

The single-sheet `SourceEdit::commit` API, managed budget execution, malicious
output and no-op performance are outside this matrix. Existing lifecycle gates
exercise some of these contracts, but do not establish separate performance
baselines. Any future optimization still requires matched before/after native
and memory gates, no-op/resource/error-order tests and relevant native Office
checks. This diagnostic cannot authorize deleting validation or substituting
staged values for independent output readback.

No Rust change means no new broad test-suite claim. Fresh formatting,
crate-boundary and registered-claim checks run; prior source-bound quality is
identified separately. All attempts are retained, deterministic analyzers are
replayed read-only, raw evidence is sealed, and only the owned target is
removed after verifying both binary hashes and process references.

The first profile child failed before producing a report because Valgrind's
unused gdbserver could not initialize its shared-memory file in `/tmp`.
Its logs, receipt, artifact hashes, and original paths are preserved under
`failed-attempts`. The separately frozen capture amendment disables gdbserver
and points debugger paths inside the owned target; collection toggles,
reset/dump boundaries, binary, cases, and sample counts are unchanged. Passed
native and preflight captures were not rerun. The exact failed-child debugger
temporary file was removed with a retained cleanup receipt.
