# DOC batched stream publication

Date: 2026-08-11

Production base: `4c4e81a873909884e62df355ab461c7f8fbf1bd8`

Scope: native binary DOC body-text publication and the shared OLE2 object
editor only. iWork/IWA crates were explicitly excluded.

## Hypothesis

A changed DOC body transaction publishes both `WordDocument` and the selected
table stream. `RevisionEditor::commit` previously sent those replacements to
the CFB object editor separately. Each call cloned the editor, rendered the
entire compound file, reopened it, checked protection, and recaptured every
stream before the next replacement began. A normal one-paragraph edit
therefore performed this full package-publication sequence twice before the
mandatory final snapshot validation.

Applying the two replacements to one isolated candidate and publishing that
candidate once should remove one complete intermediate CFB render/reopen while
preserving the final owner and public-reader validation boundary.

## Change

`litchi-ole-common::object::Editor` now exposes
`put_streams_shared`, a bounded failure-atomic batch replacement primitive. It
applies replacements in order to one candidate, retains the last value for a
repeated path, reuses supplied `Arc<[u8]>` allocations, and renders/reopens the
CFB only after every replacement succeeds. An all-no-op batch does not clone or
publish a candidate. A failed replacement leaves the editor at its exact
source state.

`RevisionEditor::commit` collects the ordinary `WordDocument` and table-stream
replacements in an inline `SmallVec` and publishes them through that primitive.
An existing changed `Data` stream joins the same batch. Creation of a missing
`Data` stream keeps its existing separately validated add path; that uncommon
structural case is outside the measured paragraph-only workload.

The following publication gates remain mandatory and unchanged:

- every individual stream path and size is checked against explicit limits;
- the batch is applied to an isolated candidate and failures are atomic;
- the resulting CFB is rendered, reopened, protection-checked, and completely
  recaptured before the object editor accepts it;
- the final DOC snapshot still performs the strict `RevisionEditor` reopen and
  the independent complete `Package::document` reopen;
- changed commits still produce exact-source forward patches, exact inverse
  restoration, and complete diagnostic/readback state;
- encrypted, signed, DRM, and protected mutation refusals are unchanged;
- exact no-op commits continue to return the original shared byte allocation.

No parser trust boundary, public DOC transaction contract, dependency edge,
executor, durability mechanism, or format capability changed.

## Matched latency measurement

Both binaries use the identical unchanged harness at the production base. The
before binary SHA-256 is
`edb3a8100249e7e17ec761e149b5bd77e8ba6061d6168b4e34ebaa041db45b05`;
the after binary SHA-256 is
`3aa09bc02601c76a6cfb93a0bad10b762ba562364f67e545bc84823e84356f3c`.

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large DOC contains 512 paragraphs,
is 97,792 bytes, and has archive SHA-256
`3d96764fe48e213b972ff5921df183dab9e8bfc8c8e751bcf3bf20190de4fec6`.
Its 81,920-byte `WordDocument` stream has SHA-256
`33e6cd70a45181c28d4a3e7bfa4e7817bd82d7b2e89e39437a589243abdc38eb`.

The primary ABBA run used 50 warmups and 500 samples per leg. Pooling the two
legs gives 1,000 raw samples per state; pooled statistics are recomputed from
the raw samples rather than from leg medians.

| Large DOC one paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 1,506.076 us | 1,347.672 us | **-10.52%** |
| p95 | 1,631.869 us | 1,487.580 us | **-8.84%** |
| mean | 1,525.904 us | 1,365.960 us | **-10.48%** |

The approximate independent-sample 95% interval for the mean delta is
`[-11.00%, -9.96%]`. Matched A and B p50 comparisons improved by 10.10% and
10.71%, respectively. Within-state p50 drift was 0.02% before and 0.70% after.

Raw primary reports and their SHA-256 digests:

- `abba-doc-stream-batch-one-edit-before-a.json`:
  `b4de7af5973c26d6c81d7716650305e27d4b8c5dad965b9a3ce37814dfa39512`
- `abba-doc-stream-batch-one-edit-before-b.json`:
  `0d33494a89e40e72963b3068a84537edaa81cfdb23addcd28ed02d8328e221a9`
- `abba-doc-stream-batch-one-edit-after-a.json`:
  `5227f9622ecc821e60ac640a69250413931ad09d342f3c0ebf1bc1ac78206fbe`
- `abba-doc-stream-batch-one-edit-after-b.json`:
  `2eb7950471eaf24cda83b6f7ca52111735a6c3a0f10bb26a1d42f107f75b0333`

## Guardrails

The independent large-input open ABBA run used 30 warmups and 250 samples per
leg. The shorter exact no-op operation used 50 warmups and 1,000 samples per
leg.

| Guardrail | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| Open | 910.010 us | 783.255 us | -13.93% | -13.64% | -12.46% |
| Exact no-op edit/save | 228.700 us | 227.053 us | -0.72% | -0.38% | -0.53% |

Neither guardrail can execute the new batch method. The open improvement is
therefore treated as an unrelated binary-layout or system effect and is not
attributed to this change. The exact no-op path is effectively flat and its
allocation-sharing contract is covered directly by tests. Raw reports are the
`abba-doc-stream-batch-{open,noop}-*.json` files beside the primary reports.

Before and after binaries also passed one-sample tiny runs of all three cases.
The harness verified exact no-op bytes, changed bytes, forward patching,
inverse restoration, diagnostics, and full snapshot reopen. Both binaries
reported identical archive and target-stream hashes.

## Allocations, RSS, and hardware counters

Matched Heaptrack processes used two warmups and 20 samples of the same large
one-edit case. These are whole-process totals and include the exhaustive
post-timing verifier in both states:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 1,173,837 | 1,169,630 | -0.36% |
| Temporary allocations | 668,178 | 667,536 | -0.10% |
| Leading `commit_candidate` RawVec growth calls | 3,565 | 1,794 | -49.68% |
| Peak heap | 4.38 MiB | 4.38 MiB | unchanged |
| Heaptrack RSS | 15.90 MiB | 16.43 MiB | +3.33% |
| Leaked bytes | 544 B | 544 B | unchanged |

The path-filtered allocation stacks nearly halve at the removed duplicate CFB
publication boundary, directly supporting the proposed mechanism. The small
whole-process Heaptrack RSS increase remains below the 5% review threshold.

Uninstrumented GNU Time ABBA runs used ten warmups and 200 samples per leg.
Maximum RSS was 30,848/30,976 KiB before and 30,848/30,848 KiB after, so there
is no measured uninstrumented RSS regression.

Matched `perf stat` ABBA runs over the same 210 iterations per leg reported:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 2,604 ms | 2,274 ms | -12.69% |
| cycles | 12,625,171,563 | 11,124,260,306 | -11.89% |
| instructions | 48,298,555,525 | 47,726,732,797 | -1.18% |
| branches | 10,611,888,480 | 10,510,277,730 | -0.96% |
| branch misses | 21,640,553 | 20,778,635 | -3.98% |
| cache references | 704,954,419 | 653,856,149 | -7.25% |
| cache misses | 51,947,397 | 47,867,184 | -7.85% |
| page faults | 96,376 | 21,206 | -78.00% |
| context switches | 65 | 78 | +20.00% |
| CPU migrations | 0 | 0 | unchanged |

The greater-than-5% movements were reviewed. Lower task time, cycles, cache
traffic, misses, and page faults agree with removal of an intermediate package
render/reopen. Cache-miss ratio also improved slightly, from 7.37% to 7.32%.
Context switches rose by 13 events across the two complete processes, but CPU
migrations stayed at zero and the direct latency, cycle, allocation, heap, and
uninstrumented RSS evidence all improved or remained stable. It is retained as
a process-level guardrail rather than omitted.

Raw evidence is in `perf-doc-stream-batch-*.csv`,
`time-doc-stream-batch-*.txt`, and
`heaptrack-doc-stream-batch-{before,after}.txt`.

## Correctness verification

- the new common-layer differential test proves exact byte equality between
  sequential and batched two-stream publication, shared allocation retention,
  and exact failure atomicity when a later path is invalid;
- new DOC tests prove exact no-op allocation sharing through commit, patch
  apply, and inverse apply, and compare every unmodeled CFB stream path and
  payload across a changed paragraph commit;
- complete `litchi-ole-common --all-features` and `litchi-doc --all-features`
  test and doctest suites passed;
- warning-denied all-target, all-feature clippy passed for both affected crates;
- the unchanged benchmark harness's 23 tests and warning-denied clippy passed;
- warning-denied rustdoc passed for the new public common-layer API;
- all benchmark JSON parsed successfully, and `git diff --check` and formatting
  checks passed.

The combined warning-denied DOC rustdoc command remains blocked by unrelated
pre-existing broken links in `section/columns`, `shape`, `mtef_extractor`,
`document/model/semantic`, and `parts/text`. None of those files changed in this
batch.

The final strict revision-owner and full public-reader reopens remain the DOC
publication boundary. A later DOC optimization should target a different
source of duplicate whole-document work rather than remove either validation
layer.
