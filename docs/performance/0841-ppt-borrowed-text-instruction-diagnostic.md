# 0841 — current PPT borrowed-text instruction diagnosis

**Diagnostic complete; production restored.** The many-short-text and Unicode
regressions from rejected experiment 0782 do not reproduce in these current
binaries. The same archived `Option<&str>` candidate reduces scoped guest
instructions in all four profiled cases. Many-short-text native write latency
is approximately unchanged; ASCII payload and Unicode writes improve in this
comparison. This does not establish why 0782 regressed, and the frozen plan
explicitly excludes adoption. No production change or CRUD coverage promotion
is retained.

The result changes the next action: another representation redesign is not
justified by an assumed increase in text-encoding work. Any future adoption
trial must qualify the current ordinary public route, including a control
without this diagnostic wrapper and fresh allocation evidence. A separately
confirmed compression feature discrepancy is also a strong next target.

## Question and source boundary

Base: `307434b454`. Historical 0753 identified the conversion of public writer
text into private `UserShapeData` as remaining copy work. Experiments 0781 and
0782 removed those allocations but regressed some common workloads. Their
whole-process counters included preparation/readback, so they did not isolate
the actual write operation or establish a cause.

This batch reapplies the exact archived 0782 patch for diagnosis. All twelve
affected current files match its historical before hashes. Plain text becomes
a private borrowed slice valid through synchronous serialization; rich
paragraphs retain their owned mutation path. The public writer still owns its
strings. Both source states pass fresh owner quality gates, and every measured
output matches the sealed historical golden byte-for-byte. After capture,
`candidate.py restore` proves the complete production source inventory matches
the current baseline exactly.

No API, dependency, durable format, execution policy, filesystem policy or
unsafe production boundary changes. ADR 0001 correctness priorities,
0002/0024 ownership, 0003 immutable views, 0005 evidence and memory contracts,
0006 preservation and determinism, and 0008 verification remain binding.
OLE2/OOXML remain the priority; ODF stays deferred and iWork is excluded.

## Measurement boundary

The inherited fixtures cover tiny (2 slides × 3 short boxes), many (100 × 10
short boxes), payload (16 × 4 boxes of 40,000 ASCII bytes), Unicode (16 × 4
boxes of 40,000 UTF-16 units), and rich (2 × 2 formatted boxes). Every shape has
write and lifecycle cases. Write times publication after authoring; lifecycle
also includes writer creation and authoring from prepared fixture data.
Generation, semantic and raw text readback, hashing, and writer/output
object destruction remain outside the clocks.

Both legs use a shared non-inlined `ppt_profile_probe::profile_write` boundary
around `Writer::write_to`. The result and sink barriers keep a real call/return
boundary for Callgrind. The compiler includes writer work inside the wrapper;
conversion and validation are not separately emitted selected symbols. Their
absence is not treated as zero work. No wrapper-free timing control is in this
batch, so these are results for the stated diagnostic binaries.

Rust 1.95.0, root-aligned registry versions/checksums, opt3, fat LTO, one codegen
unit, CPU 12, equally long before/after executable paths, and the recorded
EPYC host are shared. This is not exclusive host isolation. The native schedule
has six alternating paired blocks, thirty samples after three warmups, all ten
cases. Each of four profile shapes has two independent one-operation processes
per leg, no warmup. Native and Callgrind processes are separate and serialized;
no formal workload overlaps a root Cargo invocation. Read-only boundary checking
can overlap qualification, which is not performance evidence.

All **156 reports / 3,636 samples** validate: 120 native reports / 3,600 samples,
twenty qualification samples, and sixteen Callgrind samples. Every workload
child exits successfully. No workload is discarded, replaced or rerun.
Qualification retains twenty PPT files, independently parses their CFB chains
and actual stream bytes, and compares exact archive bytes across both legs and
modes. The Rust oracle checks every expected slide and semantic/raw text value;
this is not a native Microsoft Office interoperability result.

## Current native observations

Absolute p50 columns are medians of six nearest-rank process p50s, in
microseconds. Ratios are medians of six matched block ratios, with 10,000
bootstrap resamples, seed 841084, sorted endpoints 249/9749. Ratios need not equal
the quotient of the displayed absolute medians.

| Case | Before p50 µs | Candidate p50 µs | Paired ratio [95% interval] |
| --- | ---: | ---: | --- |
| tiny/write | 12.395 | 12.695 | 1.024608 [1.015247, 1.036072] |
| tiny/lifecycle | 12.690 | 12.830 | 1.011401 [1.000425, 1.019419] |
| many/write | 264.987 | 264.286 | 0.999245 [0.980132, 1.012208] |
| many/lifecycle | 308.221 | 307.077 | 0.996078 [0.987301, 1.004213] |
| payload/write | 319.367 | 250.061 | 0.777416 [0.754436, 0.827423] |
| payload/lifecycle | 404.812 | 330.201 | 0.815250 [0.782656, 0.851678] |
| unicode/write | 2,213.759 | 1,536.812 | 0.693009 [0.684760, 0.712716] |
| unicode/lifecycle | 2,353.300 | 1,672.987 | 0.707957 [0.702395, 0.716485] |
| rich/write | 16.475 | 16.321 | 0.991153 [0.960139, 0.994530] |
| rich/lifecycle | 16.785 | 16.525 | 0.981524 [0.978171, 0.986866] |

There are zero predeclared native p50/RSS guard flags (ratio >1.05 with lower
interval >1) and zero block-p50 max/min >1.2 spread flags. Tiny/write still
increases **2.46%**, and tiny/lifecycle **1.14%**, with intervals above one.
Those observations are not dismissed as noise. Rich/write p99 has a paired
ratio of **1.115737 [0.665777, 1.942186]**, the only p95/p99 point estimate above
+5%; its wide interval crosses one and does not establish a tail improvement
or regression. Every tail, mean, RSS and fault row remains in the [complete table](results/change-0841/native.md) and
[CSV](results/change-0841/native.csv); no universal tail or memory claim follows.

Unicode whole-child RSS changes 53,252 → 48,476 KiB for write (paired ratio
0.909821 [0.909261, 0.910916]) and 53,270 → 48,446 KiB for lifecycle
(0.909675 [0.908290, 0.910532]). Rich lifecycle RSS has the largest increasing
point estimate, 1.020732 [0.988154, 1.033637]. Whole-child RSS includes fixture
setup, allocation history and readback. This batch collects no operation
allocation/retained-memory metrics and assigns no heap or kernel cause.

## Operation-scoped instruction evidence

Callgrind uses the same optimized native binaries, collecting only the named
wrapper. Each capture has exactly one positive publication and an empty
termination publication. Its incoming owner cost equals owner self plus
immediate-child inclusive cost; all recorded function self costs sum to the
same total. Nested inclusive costs overlap and are not added as phases.

| Shape | Before Ir, repeat 0 / 1 | Candidate Ir, repeat 0 / 1 | Ratios |
| --- | ---: | ---: | --- |
| many | 6,168,522 / 6,168,717 | 6,084,727 / 6,084,513 | 0.986416 / 0.986350 |
| Unicode | 54,554,799 / 54,555,031 | 51,564,359 / 51,564,549 | 0.945185 / 0.945184 |
| ASCII payload | 6,885,505 / 6,885,708 | 4,302,783 / 4,302,777 | 0.624904 / 0.624885 |
| rich | 236,296 / 236,300 | 235,970 / 235,967 | 0.998620 / 0.998591 |

The many case loses 33,000 recorded self Ir in `free` and 30,000 in `malloc`
in each repeat. Unicode loses 2,969,304 recorded self Ir in the emitted memcpy
routine; ASCII payload loses 2,560,460. These are guest instruction counts,
not physical copied-byte totals, native cycles or causal time shares.

The Unicode plain textbox encoder's self Ir is **49,778,569 in both legs and
both repeats**. Many and ASCII payload likewise keep this encoder's self Ir
at 199,000 and 531,712 respectively. Shape-container self cost increases by one
recorded instruction per plain shape (1,000 in many, 64 in Unicode/payload),
while rich decreases four. That local increase is much smaller than the
removed recorded allocation/copy work in these guest profiles. The selected
textbox symbol is 2,258 bytes in both binaries; the selected shape-container
symbol changes 11,729 → 11,697 bytes. These facts do not identify a native
layout, cache or branch-prediction cause for the historical regressions.

Complete function rows, direct edges and conserved self-cost differences are
in the [profile table](results/change-0841/profiles.md) and
[analysis](results/change-0841/analysis.json). Two repeats support a bounded
instruction diagnostic, not a profile confidence interval. Current source,
current dependencies, fat LTO and the shared wrapper differ from historical
0782; old timings are not pooled or used as a third leg.

## Independent compression finding

The retained package-scoped feature trees confirm a separate baseline issue:
0836 used the nested harness lock's flate2 1.1.9, which explicitly enabled
`zlib-rs/std`. Current root `litchi-opc` resolves flate2 1.1.10, whose zlib-rs
dependency only requests `rust-allocator`; `std` moved behind
`flate2/runtime_detection`. The current OPC-only feature graph has no `std`
edge. Both locks use zlib-rs 0.6.7. Other feature combinations can unify features
differently, so this is not a claim about every workspace build.

The next compression experiment should use the current root-aligned graph and
test explicit runtime detection as a possible restoration, keeping backend,
level, framing and preservation policy unchanged. It must prove exact compressed
and archive bytes before timing. The older 0836 profile already had runtime
dispatch and cannot establish the current graph's performance. See the
[dependency review](results/change-0841/compression-review.md) and retained
feature/source evidence. No compression feature is changed here.

## Verification, repairs and closure

Fresh PPT gates pass on both source states: formatting, all-feature/all-target
check, warning-denied library Clippy, warning-denied rustdoc, full boundary
checking, and tests (**1,284 before / 1,287 candidate**, eleven ignored each).
Both final standalone probe legs pass 31 default and 36 all-feature tests.
Independent arithmetic checks recompute thirty native quantile rows, all usable
paired interval endpoints, and eight instruction-difference partitions.

All initial failures remain retained. Shared unused counter APIs needed a
local lint annotation. All-feature synthetic counter tests initially observed
test-runner allocation events; the allocator wrapper is now installed globally
only outside `cfg(test)`, matching 0840's isolation. An initial dependency
resolution selected newer cached versions and was replaced before any release
build by a root-lock-preserving resolution; both locks and receipts remain.

After 120 native runs, a reader agent changed the frozen reader despite the
freeze instruction. The profile driver caught the hash mismatch before launching
any profile process. Root restored the original reader exactly and retained the
replacement as an additive version. The correction separates Callgrind's raw
child-command header from launcher flags and rejects an unfinished child edge
at a function transition or EOF. Original fifteen-mutation evidence remains;
the corrected twenty-one-check suite, all twenty qualification rows, and exact
120-report old/new reader parity pass before additive profile admission. No
workload, metric, schedule or threshold changed.

The first full reducer then rejected the extra syntax space in Valgrind's
`cmd:` header. An additive qualifier removes leading header whitespace before
exact child-command comparison; argv and all arithmetic stay unchanged. The
failed receipt, original code and successful version remain. The final reader
is locally adapted from the hash-checked historical Callgrind grammar; it is
not falsely described as byte-identical to that parser.

Production is restored, and cleanup verifies both binary hashes before removing
only the two marker-owned temporary roots. Unrelated work remains intact.
[Packet replay](results/change-0841/README.md) binds the corrections, original and
additive admissions, raw data and final source. The wider CRUD, cold/range,
concurrency, scaling and hardware-limit goals remain open.
