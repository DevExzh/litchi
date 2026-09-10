# Development attempts

All validation attempts retain their source manifest, command, environment,
stdout/stderr and terminal state. Only source-stable final receipts certify the
accepted source. Intermediate compiler failures are retained for review.

* `dependencies1`: release DOCX dependency build passed. This prewarms the
  batch-owned Cargo target; it is not the final harness build.
* `harness-check1`: development compile ran while the new module was still
  being edited. It rejected a Debug derive over `dyn ReadAt`, a foreign-source
  oracle using the read facade instead of a transaction snapshot, and an
  incomplete duplicate report initializer. Its source changed during the
  command, so it cannot certify the final implementation even if a subset
  compiled.

The shared filesystem filled during setup while this batch target occupied
about 700 MiB. The coordinator removed only stale compiler intermediate files
older than six hours from an inactive external Cargo target, while holding its
Cargo lock and checking process references. The receipt is
`external-cache-cleanup.json`; the protected source worktree was untouched.
Further build commands use the batch-owned target and explicit temporary root.

Review requires preallocating trace storage outside the timer, streaming the
executable hash instead of loading a debug binary into one buffer, and explicit
separation of preflight patch checks from per-sample output/semantic checks.
Cold-cache status, unmanaged budget unavailability, and copied source ownership
must remain explicit rather than inferred from another benchmark's counters.

The development-only executable/cache directory was absent when this batch
resumed after the disk cleanup. Its earlier smoke observations are historical
notes, not accepted measurements. `dependencies-recovery1` rebuilt dependencies
while the harness source was being corrected; its changing source manifest is
explicitly excluded from final validation.

Independent cold review found that matching preflight/measured counts did not
by itself prove exactly one materialization. Both paths now require one; an
eligible row also requires positive logical source reads. Early cold verifier
ineligibility is serialized before optional identity fields are required.
`final2-clippy` caught a missing raw-range field in the new test fixture;
`final3-clippy` caught placement of that fixture before a public item. Both
failed attempts are retained. The corrected source is checked by the final
command manifest, with fresh builds and captures required.

The cold driver independently rejects matching counts greater than one,
retains private-directory cleanup receipts, and verifies its exact invocation
and temporary path. The outer capture driver owns the process-group timeout;
direct Rust CLI invocations have no local child deadline. Report flags for
semantic/media checks record successful completion of the fallible shared
output verifier, rather than separate semantic algorithms.

The first pilot coordinator used the same attempt directory (`pilot1`) for
both warm and cold lanes. All 36 warm and six cold operations completed, but
the cold analyzer correctly refused that mixed directory because it contained
extra warm cells. These artifacts remain as a rejected integration attempt.
Accepted pilots use separate `warm-pilot1` and `cold-pilot1` directories; the
formal lanes likewise use distinct attempt names. No timing result was used
to select which pilot to retain.

The separated `warm-pilot1` and `cold-pilot1` pilots both passed. Final review
then found an overly restrictive validator equation: the Rust realloc callback
does not increment standalone deallocation calls. The unsupported inequality
was removed and covered by a regression test. Warm scratch ownership is now
bound to the exact expected path as well. The earlier frozen protocols and
their exact capture-helper sources are retained in `development-protocol1`,
with hashes checked against those protocols. Their pilots are historical;
accepted evidence uses new `warm-pilot2` and `cold-pilot2` captures. The Rust
source and retained binaries did not change.

The `profile-r1` hardware-counter and syscall workloads passed. Its `perf record`
workload also passed and retained 1,382 samples, but the subsequent stack export
failed because it requested a CPU attribute absent from that recording. The
original failed export and exact profiling helpers are retained. Recovery uses
only the existing `perf.data`, omits the unavailable CPU field, and writes
separate process receipts and derived stacks; it does not rerun the workload.

Postprocessing also exposed three summary defects: the perf CSV parser used
runtime nanoseconds as its running percentage, the strace parser omitted rows
whose error column was blank, and qualified-only stack markers failed to match
bare DWARF inline names. Corrected summaries must be derived from the retained
raw outputs, conserve syscall/sample totals, and keep partial-callchain
attribution separate from whole-process cost. Intermediate exports remain
historical evidence. Changes in unrelated ODG source files during recovery are
recorded but do not alter the authenticated original executable or recording.
