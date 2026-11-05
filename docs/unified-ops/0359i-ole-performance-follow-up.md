# 0359i — W5 OLE performance follow-up

- Date: 2026-10-21.
- Base: the corrected W5 owner/property candidate after 0359g and 0359h.
- Status: required follow-up; no new performance exception is accepted.

The corrected W5 native run has exact-output preflight and 120 paired legs
for the owner, property, DOC, and XLS routes. The owner and history routes are
within the current comparison envelope. Three small residuals remain open:
the property read route has a median candidate/baseline ratio of `1.018868`,
the property edit route `1.013825`, and the XLS cell edit route `1.006251`.
These are retained as failed acceptance evidence rather than being called
noise. The read and edit samples are short enough that their per-operation
nanosecond values can be clock-quantized, while the XLS cell route combines
staging, commit, and serialization and therefore does not identify the costly
phase.

The paired v3 receipt is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/abba-final-851ca549ba-v3`.
Its SHA-256 manifest is
`6c15962e95a533acd277746d4d1d1b0b737df5b7b677aca5a314a0f33d14e2dd`.
The protected binaries and their source/provenance are recorded in
`final-probes-851ca549ba/MANIFEST.txt` in that same external directory. The
v3 run predates this follow-up probe change and remains diagnostic for the
three residuals.

The follow-up probe source currently hashes to
`3f03bd647f02d568a8491eae6b7e1ad9101eeb84a67360126a28dd2e07497b9c`; the
external driver `run-abba-w5-followup.sh` hashes to
`c89bb3e32dddfe8ea2b212eec4245b1767a3e16c31d692e984ec251d1542f800`.

## A — batched timing and XLS phase split

`tools/perf-baseline/examples/ole_w5_probe.rs` now accepts `--batch`. The
`read`, `edit`, and full `xls_cells` routes repeat their complete operation in
one timed sample and retain their semantic or exact-output checks. The new
`xls_cells_stage` route repeats edit and commit, stops the timer before the
package is serialized, and verifies the final package against the same exact
preflight output after the timer. This separates the cell staging/commit cost
from publication serialization without weakening the fidelity check.

The next pair must copy this exact probe source into the detached baseline
snapshot, rebuild both binaries in distinct absolute targets, and run three
ABBA blocks for `read`, `edit`, `xls_cells`, and `xls_cells_stage`. Calibrate
the batch per route so the faster endpoint's measured leg is at least 250 ms;
retain every leg and reject any shorter leg. Suggested starting batches are
`read=50000`, `edit=60000`, `xls_cells=800`, and `xls_cells_stage=800`; these
are calibration seeds, not results. The ordinary six owner/property routes
remain in the same source/binary fidelity preflight.

## B — allocation and retained-owner evidence

The follow-up must also capture paired allocation counts and peak memory for
ordinary `read`, `edit`, and `xls_cells` operations, plus the retained owner
and 128-step DOC/XLS history routes. The history capture must keep all
published packages and patches live for the full interval, and the ordinary
capture must exercise the same retained owner used by the timing probe.
Direct allocator/heaptrack invocation is required; an environment wrapper
that does not survive `taskset` is discarded. The receipt must include the
source and binary hashes, exact command, output/fidelity checks, live retained
cardinality, decimal heap units, and all discarded captures. No allocation or
memory result is claimed until this paired evidence is available.

The only acceptable closure is a fresh paired result that explains or removes
the three residuals and supplies the allocation, peak, ordinary-owner, and
history evidence. Until then, 0359g and 0359h remain implementation-complete
but performance-unaccepted.
