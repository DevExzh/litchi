# 0898 — bounded emitted-name output spans in legacy MCE

Change 0898 evaluates a narrow allocation candidate in the legacy markup
compatibility codec. An emitted element used to retain an owned qualified name
so that its closing tag could be written later. The candidate retains the
qualified name's byte range in the append-only output buffer instead. It also
represents emitted elements below inactive branches explicitly, so those
frames retain no closing-name allocation.

The candidate has passed the source, quality, semantic differential, capture,
reduction, and verification gates. Its native primary p50 ratio is
`0.9850658963`, a `1.4934%` improvement that misses the required `2%`
improvement threshold even though its bootstrap upper endpoint is below `1.0`.
The candidate is rejected for adoption. Independent raw replay and
post-cleanup verification pass. Production adoption is false and the two
production files are restored. iWork is outside this batch's scope.

## Motivation and bounded hypothesis

Change 0897 attributed allocation calls in the restored XLSX full-cell
fallback. Its scan-closure partition had 518,660 calls over 20 scans, of which
343,320 were nearest to the legacy MCE `start` owner. Two identical exported
stack rows carried 166,800 and 420 calls, a 167,220-call string-copy
signature, or 32.24% of scan-closure calls. Those are stack signatures, not
source-line counts. They do not prove which `Mode::Emit(q.to_owned())`
expression allocated or how many allocations a name-span change would remove.
The 0897 source audit explicitly leaves that call-site attribution
unresolved.

The source-supported mechanism is smaller and more specific: both start-tag
writer paths already append `<` and the emitted qualified name to the bounded
output. A closing tag can refer to those bytes by checked offsets and copy
them through the existing fallible output reservation. This candidate does not
bundle the rejected QName, metadata, or stream candidates from 0894, 0896, or
earlier changes. Allocation attribution alone cannot authorize adoption.

## Candidate implementation

The overlay contains only:

- `crates/litchi-ooxml-common/src/mce/codec.rs`
- `crates/litchi-ooxml-common/src/mce/tests.rs`

`Mode::Emit(String)` is replaced by an output `Range<usize>`, and an explicit
`Inactive` mode represents an element that would otherwise have retained an
owned name below a dropped or inactive branch. `write_start` records the
offset immediately after `<` in both its source-copy and rebuilt paths. It
returns the checked span only after the complete start tag, including the
final `>`, has succeeded. A failed start therefore cannot publish a frame
with an invalid closing span.

The output helper validates the range before copying, checks the subtraction
and source end, calls the existing fallible bounded reservation, and then uses
safe `Vec::extend_from_within`. The output remains append-only, so later
`Vec` growth may move the allocation without changing a retained offset.
Closing writes keep the existing order and limit boundaries as three
operations: `</`, the stored name range, and `>`. Empty elements use the same
path. Opaque elements, inactive opaque descendants, skipped and unwrapped
elements, alternate-content branches, namespace hoisting, report counters,
and validation order retain their existing control behavior. No unsafe code,
public API, dependency, reader path, or unchecked slice was introduced.

The focused tests include exact bytes for a preserved, unwrapped, skipped,
opaque, and alternate-content document. A second test keeps the root name
span live while 64 child tags normalize `&#34;` to `&quot;`, forcing output
reallocation. It checks the complete output, the exact output limit, and the
one-byte-under limit refusal.

## Correctness and quality evidence

The independent source review approves the overlay for qualification and
explicitly does not approve production adoption or a performance claim. Its
final source identities are retained in
[`source_review.md`](results/change-0898/source_review.md) and
[`candidate-source.json`](results/change-0898/candidate-source.json). The
quality receipt reports ten passing gates in
[`quality.json`](results/change-0898/quality.json):

| Gate | Result |
| --- | --- |
| `litchi-ooxml-common` tests | 445 passed |
| `litchi-ooxml-common` Clippy with warnings denied | pass |
| `litchi-ooxml-common` documentation with warnings denied | pass |
| XLSX library tests | 1,390 passed |
| DOCX library tests | 1,230 passed |
| PPTX library tests | 851 passed, 1 ignored |
| diagnostic tests | 24 passed |
| producer-shape tests | 22 passed |
| scoped rustfmt | pass |
| crate-boundary check | pass |

The first common-crate test attempt failed to compile because the new focused
test used `Name` without importing it. That retained failed receipt is part
of the packet history; the import correction was followed by the accepted
445-test pass. The first PPTX attempt was separately invalidated by shared
filesystem `ENOSPC` while writing its log and receipt. The owned rustdoc
output was removed, and `qualify-v4` resumed the PPTX gate under a fresh label;
the accepted run passed with 851 tests and one ignored test. These attempts
are not hidden or folded into the final pass count.

The three-package differential corpus is byte, record, and semantic equal
between baseline and candidate. It used 20,000 generated inputs and 20,000
aliasing inputs over three packages, 67 members, and 200,201 oracle records;
the retained legacy result count is 200,405. The outcome totals are 48,804
codec refusals, 51,330 codec successes, 48,133 stream refusals, and 51,934
stream successes. The compressed baseline and candidate reports are equal at
226,196,057 uncompressed bytes and round-trip successfully. The full receipt
is [`corpus-comparison.json`](results/change-0898/corpus-comparison.json).

These gates establish behavior and build quality. They do not establish a
latency, throughput, allocation-byte, peak-memory, RSS, or whole-producer
improvement.

## Matched performance protocol and result

The frozen protocol uses the retained producer and format controls under the
same release settings, warm filesystem cache, CPU 12, and fixed AB/BA paired
ordering. It has 72 fresh sequential children and 3,320 measured samples:

| Lane | Children | Warmup per child | Samples per child | Measured samples |
| --- | ---: | ---: | ---: | ---: |
| Native | 64 | 3 | 50 | 3,200 |
| Allocator | 8 | 3 | 15 | 120 |
| **Total** | **72** |  |  | **3,320** |

The native cases are the producer XLSX primary case, a marker-free XLSX
control, a DOCX control, and a PPTX control. The primary admission gate is a
candidate/baseline native primary p50 paired median ratio at most `0.98` with
the bootstrap 95% upper endpoint below `1.0`. Bootstrap seed `898`, 10,000
draws, and endpoints 249 and 9,749 are fixed before capture. Pair-ratio and
native p50 spread thresholds are review flags; RSS is descriptive and has no
causal interpretation.

The capture contains all 72 children and 3,320 measured samples. Its reduction
uses whole paired child ratios with seed `898`, 10,000 draws, and bootstrap
indices 249 and 9,749. The table reports the median of the eight per-child
aggregate p50 values for each side; the ratio and interval are the reducer's
median paired ratio and its bootstrap interval.

| Case | Baseline p50 median (ns) | Candidate p50 median (ns) | Paired p50 ratio [95% interval] |
| --- | ---: | ---: | ---: |
| XLSX producer medium full-cell scan (primary) | 3,430,518.5 | 3,376,939.5 | 0.9850658963 [0.9772547023, 0.9877410001] |
| XLSX marker-free full-cell control | 133,290.5 | 132,760.0 | 1.0003987089 [0.9890346975, 1.0069365476] |
| DOCX selected-paragraph control | 107,490.0 | 106,185.5 | 0.9877077533 [0.9576820262, 1.0118348233] |
| PPTX selected-slide control | 178,746.0 | 174,511.0 | 0.9804438585 [0.9749882626, 0.9843865583] |

The primary gate requires a median ratio at most `0.98` and a 95% upper
endpoint below `1.0`. The observed primary ratio is above `0.98`, so the gate
is false despite the interval's upper endpoint of `0.9877410001` being below
`1.0`. The primary p95 and p99 ratios are `0.9780905393` and `0.9761591086`,
respectively; they do not override the p50 admission rule.

The separate allocator lane records the following primary-route medians:

| Metric | Baseline | Candidate | Difference | Relative change |
| --- | ---: | ---: | ---: | ---: |
| Allocation calls | 25,933 | 17,565 | −8,368 | −32.2678% |
| Allocated bytes | 37,618,900 | 37,609,931 | −8,969 | −0.023842% |
| Region peak-live bytes | 9,252,430 | 9,252,427 | −3 | −0.000032% |

The allocator observations are useful diagnostic evidence, but they cannot
replace the failed native p50 gate. No native p50-spread flags fired. Two
control p99 pair flags remain visible: XLSX control pair `native-b01-c01` at
`1.1892610831`, and DOCX control pair `native-b02-c02` at `1.0520929826`.
All peak-RSS flags are descriptive observations with the known fork/exec and
executable-residency bias; they do not establish a heap or causal RSS result.

The sealed reduction and verification receipts are
[`reduction.json`](results/change-0898/reduction.json) and
[`verification.json`](results/change-0898/verification.json). Independent
replay passes in [`native-review.json`](results/change-0898/native-review.json):
all 72 children, 3,320 raw samples, guard vectors, artifact hashes, paired
statistics, allocator counters, and bootstrap results match. The original v1
helper failed on an undefined source-path variable; v2 rejected the valid
`approved_for_measurement` status; v3 and v4 expected flat binary descriptors
where child receipts and commands use declared/identity wrappers. The
versioned v5 helper verifies both wrapper components and passes. Every failed
helper and receipt is retained; none launched a workload or changed capture.

The [`decision.json`](results/change-0898/decision.json) rejects production
adoption. The two owned production files are restored, with exact hashes in
[`restoration.json`](results/change-0898/restoration.json). Candidate overlays,
source/lock receipts, release and diagnostic binaries, compressed differential
reports, raw samples, and logs remain retained. Executables are locally
retained and ignored by Git; their exact hashes and build commands are tracked.

Cleanup removed 3,302,605,032 bytes from the marked isolated source/build roots,
after earlier removal of 3,005,435,288 bytes of redundant raw reports/debug
outputs and 15,694,823 bytes of qualified rustdoc output. Fixture relocations
map deleted source paths to identical repository fixtures. No worktree was
created; unrelated worktrees and the untracked benchmark file were preserved.
[`owned-cleanup.json`](results/change-0898/owned-cleanup.json) and
[`verification-after-cleanup.json`](results/change-0898/verification-after-cleanup.json)
record passing cleanup and retained-evidence replay.

## ADR and contract matrix

| Constraint | 0898 evidence |
| --- | --- |
| ADR 0001 priorities and API layers | The change is a private codec allocation-shape candidate; no public API or semantic shortcut was added. |
| ADR 0002 topology and dependency direction | The implementation remains in `litchi-ooxml-common` and adds no dependency or cross-crate ownership. |
| ADR 0005 I/O, memory, and measured performance | Output growth remains bounded and fallible; the fixed paired native gate is required before any performance claim or adoption. |
| ADR 0006 validation, security, and compatibility | Name validation, namespace handling, branch behavior, typed limit errors, and preservation paths remain in place; the candidate does not repair or normalize input. |
| ADR 0008 migration and verification | Focused tests, common-format gates, differential corpus evidence, source review, and independent paired review precede any production decision. |
| ADR 0024 current topology | No workspace topology or format-owner boundary changes. |

## Limitations and next work

The 0897 allocation evidence motivates this candidate but cannot identify an
exact removable `q.to_owned` count. The 0898 focused reallocation test proves
offset stability under the exercised output growth, not a general allocation
or speed result. The differential corpus covers the recorded generated,
aliasing, and package corpus; it does not replace native Office compatibility
review or the matched producer controls.

The primary gate has now failed on the required 2% p50 benefit, so allocator
reductions do not authorize adoption. Independent replay and cleanup receipts
remain final evidence obligations; production restoration has passed. Future
work should profile the remaining fallback allocation owners and revisit this
area only with a stronger measured hypothesis that clears the fixed end-to-end
gate.

The qualification packet, including the methodology preflight and retained
failure receipts, is in [`results/change-0898/`](results/change-0898/).
