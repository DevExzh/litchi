# Change 0693 — capture-local notes root proof reuse

This packet retains the measured capture-only PPTX optimization.
`performance_claim: none` remains: results are scoped to the recorded host,
corpus and public-call phases, with no coverage promotion. Final native,
allocator, profile, refusal and trace receipts bind to the corrected source;
all integration and repository gates pass. Independent code review accepts
the measured late-refusal cost. The [change report](../../0693-pptx-capture-notes-proof.md)
records the final result and limitations. No 0692 latency is reused; its trace
is a source-bound call-count baseline only. iWork remains outside scope.

The candidate reuses work that the opened-PresentationML capture has already
performed. Capture currently processes a slide once to validate its root and
read its producer name, then drops the processed bytes. A later notes-graph
load scans the same slide again to establish its Transitional or Strict root
conformance. The candidate scans the already processed bytes immediately after
a successful name, retains only the exact second borrowed raw-slice witness
and an `Option<Conformance>`, and lets the later notes load use that proof only
after checking the current part. No processed XML, `Cow`, `Arc`, scanner
inventory, or XML tree survives the capture call.

The preprocessing helper has an intentional two-observation contract. Its
first `Part::blob()` observation is used for the existing 64 MiB limit check.
Its second observation is the exact borrowed slice passed to MCE and is the
slice whose pointer and length can become the proof witness. The implementation
must not infer the MCE input from the first observation or call a helper that
silently obtains a third, different part view.

The proof is a prefix in PresentationML slide order. A full notes scanner runs
over the processed bytes only after the slide root and name result has
succeeded. The first invalid scanner result is included as a proof entry with
`None` conformance and closes the prefix; a matching raw witness for that entry
reproduces the existing generic invalid-root error directly. A first name error
closes the prefix before that slide. The capture still performs every root/name
and identity check required by the existing error order. A private
`load_snapshot` path receives the expected catalog length and proof prefix. The
notes presentation scanner must have the same slide-inventory length before any
hint is enabled. For every presentation slide, including slides without speaker
notes, it validates a hint at the slide's original catalog position by obtaining
the current blob once and comparing its pointer and length. A missing proof, a
`None` entry with a mismatching witness, or an identity mismatch uses the
legacy `root_conformance` path. Graph traversal, relationship checks,
materialization, and publication ordering remain unchanged.

The separate proof vector is expected to cost about 24 bytes per slide on the
measurement target. Capture scratch entries are about 48 bytes per slide and
should be released before notes loading when the ownership flow permits. With
`N` equal to the actual capture-catalog bound, the two vectors account for
approximately `24N + 48N` bytes, before allocator overhead. For an ordinary
immutable input under the default `max_parts=4096`, that is the qualified
~294,912-byte two-vector bound; it is not a universal bound. The low-level
`MAX_SLIDES=100000` guard and caller-custom limits also apply, and a foreign
`Part` can expose a different catalog between the initial reference check and
the capture catalog. Existing catalog-length and identity mismatch handling is
unchanged. If its optional fallible reservation cannot be made, the root path
disables hints and continues with the existing capture behavior; it does not
introduce a new early allocation refusal.

Earlier full scans can make a late root or relationship refusal do more work,
because the scan now occurs after an earlier successful name and before the
later refusal. The candidate also removes an independent MCE allocation-failure
opportunity that existed when notes performed its later scan. The proof-only
scanner retains the existing infallible vector-growth behavior, however, so an
allocator exhaustion during that scanner can still abort before a later refusal.
Only the optional proof-vector reservation is treated as recoverable and
advisory. These are resource and refusal-behavior differences to measure and
review; neither is a document semantic shortcut.

The fresh 0693 native matrix covers five one-edit cases, five exact semantic
no-op cases, and three two-slide cases. The A/A baseline has 26 process
records; the comparison run adds 52 A/B/B/A records, for 78 process records
and 7,800 timed samples in all. Each process uses five warmups and 100
samples. Each timed sample separates capture, working clone, text edit, commit,
apply, and the enclosing total. Source materialization, package open, save,
reopen, and the semantic oracle are outside the phase timers. The no-op requires
the before and after serialized outputs to match each other, with unchanged
commit and revision; it does not compare the result to the original archive.
The two-slide workflow checks both edited markers and excludes exactly those
two slide payloads from its untouched-payload oracle. Metadata and relationship
inventories are compared order-insensitively, while content, types,
relationships, and non-part members remain covered.

The fresh post-fix native medians are below. Each pair is candidate versus its
paired baseline leg (`b0/a2` and `b1/a3`), and the values are total-phase p50
changes from [`tables.md`](tables.md).

| Workflow group | Paired p50 total change |
| --- | ---: |
| real one / no-op / two-slide | −29.42% / −29.11%; −34.22% / −34.68%; −28.83% / −28.93% |
| control one / no-op / two-slide | −9.57% / −10.78%; −16.71% / −16.03%; −8.77% / −11.15% |
| generated one / no-op / two-slide | −8.82% / −8.99%; −17.08% / −9.57%; −6.88% / −7.77% |
| notes POI one / no-op | −1.13% / −2.05%; −0.85% / −2.98% |
| notes LibreOffice one / no-op | −2.83% / −2.36%; +2.32% / −2.24% |

No total-phase p50 regressed by more than 5%. The generated no-op baseline
legs differ materially (`0.9082` versus `0.8301` ms), so the generated rows
remain paired rather than being reduced to one baseline number. The 28 review
triggers are phase-level tail or noise observations; the sole non-tail median
and mean trigger is the no-op LibreOffice apply phase (`+14.46%` and
`+14.71%`, about `+2.32` and `+2.39` µs), while its total remains within the
5% review boundary.

The counting-allocator baseline and candidate each use one separate process per
case (13 processes), with three samples and two warmups per case. Their counters
are diagnostics from separate binaries and are not native timing evidence. For
the one-real diagnostic, allocation calls fall from `270,242` to `190,338`,
reallocation calls from `6,950` to `4,870`, and requested bytes from `19,915,826`
to `13,815,905`; peak live bytes are nearly flat (`463,159` to `459,170`) and
net live bytes are unchanged (`195,730`). The baseline and candidate prefix
profiles contain sampling, hardware counters, and whole-child RSS for the real
deck; they are diagnostics, not isolated phase measurements. Per open-plus-
capture profile counters change from `39.738` to `28.782` million cycles,
`159.173` to `108.711` million instructions, and `211.32` to `175.075` page
faults. Whole-child peak RSS changes from `5,540` to `5,632` KiB (`+1.66%`, one
run each). Build inputs, source hashes, probe hashes, leg metadata, raw TSVs,
and summaries are retained in this directory, beginning with
[`baseline.json`](baseline.json),
[`builds-baseline.json`](builds-baseline.json), and
[`native-runs-baseline.json`](native-runs-baseline.json).

The supplemental refusal matrix has ten bounded fixtures and exercises
duplicate names, late slide-root and relationship refusals, notes tails, mixed
conformance, valid controls, and the notes-size boundary. The exact-baseline
rebuild and source restoration are recorded in
[`expanded-refusal-baseline/manifest.json`](expanded-refusal-baseline/manifest.json);
the earlier eight-case receipts and source archive remain under
[`refusal-eight-case/`](refusal-eight-case/). An initial attempt serialized a
deliberately malformed graph and was rejected by OPC publication before timing;
its build and smoke artifacts remain under `refusal-probe-initial/` and
`refusal-initial/`. The corrected probe hashes the in-memory graph directly;
its ten-case, two-leg, 100-sample baseline and retained bindings are recorded in
`refusal-runs-baseline.json` and `refusal-summary.json`. The notes-invalid-tail
case intentionally reports the direct XML parser refusal from the notes
resource scan. No timing or refusal result is based on the rejected attempt.

The fresh ten-case refusal comparison is tabulated in
[`refusal-tables.md`](refusal-tables.md). The late-root and late-missing-
relationship cases add `43.95–48.72 µs` at p50: marker-free cases rise by
`125–161%`, and MCE-bearing cases by `42–47%` across the paired legs. This is
the expected cost of scanning earlier successfully named slides before a late
refusal. The late relationship check still occurs before preprocessing its own
slide, so only earlier slides contribute those full scans. The removed repeated
MCE allocation-failure opportunity is a separate resource difference.

The marker control is a mechanism control with regenerated ZIP compression;
it is not semantic equivalence to the original deck and is never an untouched
payload oracle. Reproduction uses the retained 0693 probe, locked release
inputs, two Cargo jobs, CPU 12, warm OS caches, and the recorded source and
binary bindings. The candidate trace uses the same sibling target directory as
the other drivers and remains diagnostic only; it supplies no latency result.

The fresh native, allocator, profile, and ten-case refusal summaries are now
source-bound to the corrected candidate. All seven repository quality commands
record exit 0: default tests report 918 passed and 2 existing ignored tests,
all-feature tests report 932 passed and 2 existing ignored tests, and the
facade reports 45 passed. The diagnostic trace receipt now passes with no
skipped runs, and the evidence audit passes. Final cleanup and seal receipts
are retained as `cleanup.json` and `final-validation.json`.

Reproduction commands, run from the repository root, are:

```text
# Copy this packet into a fresh checkout at baseline_head first.
# Recreate the excluded generated marker control before measurement:
python3 docs/performance/results/change-0693/prepare-control.py
python3 docs/performance/results/change-0693/check-control.py
# At baseline_head, before applying the candidate source:
python3 docs/performance/results/change-0693/build.py baseline
python3 docs/performance/results/change-0693/measure.py baseline
python3 docs/performance/results/change-0693/build-refusal.py baseline
python3 docs/performance/results/change-0693/measure-refusal.py baseline
python3 docs/performance/results/change-0693/measure-allocations.py baseline
python3 docs/performance/results/change-0693/profile.py baseline

# With the candidate source restored:
python3 docs/performance/results/change-0693/build.py candidate
python3 docs/performance/results/change-0693/measure.py compare
python3 docs/performance/results/change-0693/build-refusal.py candidate
python3 docs/performance/results/change-0693/measure-refusal.py compare
python3 docs/performance/results/change-0693/measure-allocations.py candidate
python3 docs/performance/results/change-0693/profile.py candidate
python3 docs/performance/results/change-0693/report-metrics.py
python3 docs/performance/results/change-0693/trace.py --profile candidate --model-trace --with-generated --with-control --cpu 12 --target-dir ../litchi-target-0693
python3 docs/performance/results/change-0693/trace-summary.py
python3 docs/performance/results/change-0693/run-integration.py
python3 docs/performance/results/change-0693/quality-summary.py
python3 docs/performance/results/change-0693/audit.py
python3 docs/performance/results/change-0693/seal.py
```

The trace, audit, and seal commands are verification steps, not current final
retain claims. The trace driver archives and restores the patched source in
`finally`, and all drivers use `../litchi-target-0693` as their shared target.
The baseline build commands are source-bound to `baseline_head`; the expanded
refusal driver is the candidate-state bridge that temporarily restores the
baseline source, builds its ten-case refusal binary, and restores the candidate
bytes exactly.

`expanded-refusal-baseline.py` records the historical supplemental-baseline
build and exact candidate restoration. It is not needed in the ordinary
reproduction sequence above, which builds the final ten-case probe directly
at the baseline checkout. Retained output folders should be archived before
regenerating evidence; the historical helper refuses to overwrite its archive.

Independent reviews: [code review](code-review.md) and
[evidence review](evidence-review.md). The post-cleanup trace verifier accepts
a missing trace binary only beneath this batch’s exact owned target with its
recorded removal receipt; retained source archives and raw logs stay verified.
