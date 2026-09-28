This review is carried forward from 0828 after checking that the 0829 probe
changes only packet names and the pinned base revision. Runtime operation and
oracle code are identical. Fresh 0829 gates and live-symbol checks remain required.

# 0829 PPTX edit profile probe review

Status: static review pass. This review covers the packet-local probe and the
0829 profile/source design documents. No Cargo command, formatter, build,
workload, or profiler was run while preparing it. The review is against the
packet's stated base `90b9466b99639d411f9b0a89208aad21525e0006`; production
files were not changed.

The probe measures the same public edit sequence as the PPTX branch of
`tools/perf-baseline/src/ordinary_save.rs:758-777`:

```text
Package::open(input)                                      outside the clock
  opened_presentation_transaction()
  set_shape_text(0, 0, "litchi-perf-0638-ordinary-save")
  commit()
  Commit::is_changed()
  apply_opened_presentation_commit(commit)                 inside the clock
Package::to_bytes(), owner drop, hashing, reopen/readback outside the clock
```

`edit_helper_0829` contains that direct sequence at
`probe-src/src/main.rs:272-290`. Wrapped mode performs the same calls through
the capture, set-text, and commit/apply phase functions at lines 296-343. The
phase functions only add no-inline call boundaries and `black_box` observation;
they do not discover another target, add an edit, or replace a public operation.
The fixed `(0, 0)` selector and marker match the pinned `shapes.pptx` contract.

The timing boundary is correctly placed. `execute_one` opens a fresh
`Package` at lines 353-356, starts `Instant::now()` immediately before the
direct or wrapped region at line 357, and reads the elapsed time after the
region returns at lines 358-363. Error propagation with `edited?` is after the
clock stops. `to_bytes`, package destruction, output hashing, and semantic
verification are after the clock at lines 365-370. In direct mode the returned
publication snapshot is the temporary at the `apply` statement (line 288), so
it drops inside the helper. Wrapped mode names the returned snapshot and calls
`drop(snapshot)` at lines 340-343. The transaction is consumed by `commit`, the
commit is consumed by `apply`, and the package owner is dropped at line 368;
these ownership points match the ordinary edit boundary.

There is no sample-package pre-capture or memo warm. `run` builds the oracle
from the separately opened reference bytes before the loop (lines 174-192),
but every sample and warmup calls `Package::open` on a new owner immediately
before its timer. It does not call `opened_presentation` or any semantic view
on that owner before timing. Thus the public
`opened_presentation_transaction` call retains the production capture,
complete-source revision, and publication validation path. In particular,
`Package::opened_presentation_transaction` delegates to
`opened_presentation` (`crates/litchi-pptx/src/package/model.rs:271-278`), whose
graph check and capture/fingerprint path remain intact (`:247-268`), while
`apply_opened_presentation_commit` still enters the shared graph and mutation
checks and `apply_committed` path (`:530-597`). The probe introduces no memo or
fingerprint shortcut.

The oracle is sufficiently strict for release capture. `read_pinned_file`
checks regular-file status, size stability, the 32 MiB bound, and exact input
and reference identities (`main.rs:488-535`), with the admitted identities
recorded at lines 36-41. `build_oracle` reopens the separately pinned output
and derives complete presentation text, slide count, and `(slide 0, shape 0)`
text (`:379-419`). Each warmup and timed sample runs `verify_output`
(`:422-486`), which requires the exact output byte vector, output SHA-256 and
size, a successful public reopen, exact full text and full-text digest, exact
slide count, and exact target text equal to the marker. The run loop rejects
any failed verification before retaining a sample (`:194-237`). The focused
parity test also compares direct and wrapped bytes and semantics against the
pinned real-file reference (`:756-807`); CLI and wrong-identity tests are at
`:726-754` and `:809-831`.

The wrapper ownership and call-boundary shape is suitable for the planned
selectors. `edit_region_0829` is `#[inline(never)]`, retains its result across
`black_box(&result)`, and returns it only afterward (`:292-302`), which gives
the outer owner a real call boundary rather than a tail-call-shaped wrapper.
Each phase is also `#[inline(never)]` (`:304-344`). Capture returns the owned
`Transaction`; set-text borrows it; publish consumes it, consumes the commit,
and explicitly drops the publication snapshot. The live-binary assembly gate
must still confirm the emitted call/frame-pointer shape, but the source has
the required no-tailcall structure and the configured selectors match the
function names in `plan.json` and `analysis.py`.

One documentation detail is worth retaining for reviewers: the comment at
`main.rs:272-274` calls this “the same five public methods,” while the actual
sequence is four public calls plus the read-only `Commit::is_changed()` check.
This is not a timing or semantic issue; the calls themselves match the
ordinary-save sequence exactly. No probe-source blocker was found.
