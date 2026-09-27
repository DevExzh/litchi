# 0794 — bounded inline duplicate-key experiment

**Rejected; exact baseline production restored.** Large PPTX capture allocation
calls fall 85.048%, but no primary case reaches the required timing benefit.
Tiny XLSX full-cell scanning regresses 5.511% and triggers the additional veto.
All six quality gates pass; 274 retained reports cover 8,487 measured samples
and 84 separate helper iterations. See the [report](../../0794-xml-inline-duplicate-keys.md).

This packet tests removal of small-tag duplicate-key heap allocations in the
shared fail-fast XML attribute iterator. Five dependency-isolated source copies
must remain aligned; existing external differential tests remain unchanged.
The similarly named formula helper is a separate lenient `first_wins` path and
remains unchanged. The base is `008bea923a`. The accepted 35 architecture/goal/taxonomy inputs are
hash-bound and unchanged. Production source changes are limited to the six
files allowed in the frozen plan, including new inline tests. That allowlist
was gathered by filename and includes the unchanged formula helper; it is an
upper bound, not a requirement to modify every listed path.

The exact 0793 evidence probe is reused, retaining its original 0785 schema,
fixture generator, and readback oracle. Native and counter binaries do not
enable the profile wrapper. Each leg additionally builds a profile binary
using the existing non-inlined `capture_region_0793` owner. All native work,
builds, tests, and profiler execution is root-only and serial.

The public matrix is identical in shape to 0792: fifteen capture, staged-commit,
and full-lifecycle cases across tiny, medium, large, ASCII-vendor, and
Unicode-vendor generated documents. Six alternating native blocks use thirty
samples after three warmups. Two counter blocks use three samples without
warmup. Fifteen baseline qualification children use one sample each. This is
255 reports and 5,595 measured samples before separate profile diagnostics.

The frozen adoption policy requires at least one capture/lifecycle improvement
of 3% or more with bootstrap upper bound below one, rejects any significant
p50 regression above 5%, and permits no increase in paired block median calls,
allocated bytes, net live bytes, or peak above entry. Allocation count alone
is insufficient for retention. Bootstrap seed 794079, 10,000 resamples, and
zero-based interval endpoints 250/9749 are fixed before baseline compilation.

Four separate heaptrack children profile large capture in alternating order,
then eight serial decodes retain whole-process and owner-filtered stacks and
histograms. Qualification requires exact owner subset equality, whole-process
stack/histogram/summary conservation, and owner cost equal to allocation_calls
alone. Reallocations are already included in that counter; 0793's failed sum
formula is not reused or retroactively amended. Profile elapsed time and
process heap peaks do not establish native latency, operation peak, or RSS.
Nested stack counts overlap. These four reports add four measured samples.

A separate count-only helper supplement measures 14 attribute counts, three
iterations each per leg, and records `size_of::<CheckedAttributes>()`. Its
84 iterations and two reports are separate from the public/profile totals.
It cannot establish latency benefit or substitute for the frozen adoption gates.

The cross-format supplement reuses the existing `tools/perf-baseline` native
harness for DOCX open/full-text (tiny and large) and XLSX open/full-cell-scan
(tiny and dense-wide). Its separate plan is frozen before its baseline build.
One qualification report has eight one-sample rows; six alternating paired
blocks yield twelve reports with eight 30-sample rows each. A significant
paired-p50 regression above 5% is an additional rejection gate, using seed
794080. These timings are never pooled with the differently built PPTX probe.
The supplement adds thirteen reports / 2,888 measured case samples. It does
not add a benefit requirement or weaken the original PPTX gates.

Six quality gates cover the fourteen substrate and OOXML/OLE2 consumer packages
listed in quality.py. No iWork or deferred ODF support claim follows. The
historical 0793 source review has one imprecise sentence suggesting 72,106
capture calls include fixture construction/readback; both are outside the
operation counter, as the probe and 0793 main report show. This experiment uses
the actual operation boundary.

The original before build and qualification preceded the count-only helper
and cross-format baseline build/qualification. An unintended baseline quality
run was retained separately. Candidate application and corrected quality gates
precede the after builds and helper capture. Final native PPTX and cross-format
measurements run before the allocation lane and heap profiles. See
`workflow-note.md` and `serial-execution.json` for the exact retained sequence. Frozen drivers refuse
existing output paths. Reproduction requires a separate packet and owned
target, never overwriting retained evidence. Offline replay after completion:

```sh
python3 -B docs/performance/results/change-0794/validate.py --require-final-seal
```

The rejected candidate and its evidence are retained for audit; production
source matches the exact baseline. Owned targets and caches are removed after
executable identities are recorded. No broad goal completion is claimed.
