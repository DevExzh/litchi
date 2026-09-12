# 0531: reject the shared OOXML MCE namespace-search candidate

The rejected candidate replaces one overlapping-window byte search with the existing
`memchr::memmem::find` dependency. It preserves the exact namespace-presence
predicate, input/output limit ordering, borrowed no-MCE result and complete
MCE parser. No dependency, public API, unsafe code or concurrency is added.
Nine public regressions exercise lexical triggers, search boundaries, every
single-byte near match, malformed inputs, resource errors and source sharing.

The original native pilot passes every primary shape/repeat. These initial candidate results did not establish retention. A later rebuilt binary differed, so a fresh complete 44-child native campaign repeated the frozen matrix. Its second dense-sparse repeat reduced total p50 by only **0.8937%** and mean by **0.6056%**, below both required 1% gates. The production change is reverted; no speedup is retained.

Historical initial-candidate results follow. The original pilot passes every primary shape/repeat. Source-backed XLSX
one-percent edit/save medians improve 1.44–1.92%, means improve 1.50–1.92%,
and planning medians improve 3.30–5.58%. Each primary child has 200 samples
after 20 warmups, in baseline–candidate–candidate–baseline order. Bootstrap
intervals describe within-child resampling, not between-machine uncertainty.
These are warm synthetic medium/dense-sparse results, not a general OOXML gain.

| Shape | Repeat | Total p50 reduction | Total mean reduction | Planning p50 reduction | Planning Ir reduction |
| --- | ---: | ---: | ---: | ---: | ---: |
| Medium | 1 | 1.9180% | 1.9232% | 3.5648% | 3.2744% |
| Dense-sparse | 1 | 1.4371% | 1.4998% | 5.5805% | 5.1073% |
| Medium | 2 | 1.7387% | 1.9154% | 3.3038% | 5.0412% |
| Dense-sparse | 2 | 1.8395% | 1.9047% | 4.9778% | 5.0470% |

Eight planning profiles pass the conditional instruction gate. The medium
MCE edge falls by the same 6,344,672 instructions in both repeats; unrelated
XML-validation variation reduces the first repeat's whole-planning gain.
Nested costs are not additive and instructions are not converted to latency.
Whole-child hardware counters move in mixed directions between repeats;
no isolated planning hardware gain is claimed.

The main comparison retains all 25 adverse metrics and 97 repeat-drift flags.
Second-repeat reopen medians rise by 7–13% in three rows. A separately frozen
12-child, 1,200-sample ABBA confirmation checks those scenarios without erasing
the original results: all six matched reopen/total p50 and mean guards pass,
with the largest reopen increase below 0.5%. Its eight adverse metrics and 57
drift flags remain individually reviewed. The evidence does not prove a cause
for the earlier regressions or establish general reopen stability.

DOCX/PPTX and existing XLSX guards preserve exact output identities. Eight
eager-read children have no matched timing/RSS regression above 5%; one
negative repeat-to-repeat RSS drift is retained. Eager dense repeat2 total
p50 rises 1.0656%, so this is a guard result, not a universal eager speedup.

Commit and publication allocation diagnostics separately preserve allocation
calls, reallocations, allocated bytes and incremental region peaks in all
paired samples. Absolute region peaks are three bytes higher in the candidate
because the pre-region live-byte baseline is higher. Planning allocations
are not isolated; no whole-process memory reduction is claimed.

The first correctness run exposed four mistaken expectations in the new tests.
Corrected expectations pass all nine tests against unchanged baseline
production. The final candidate rebuild retained the measured runtime source but produced a different binary hash and a different ELF text-section size. The cause is not established. Fresh measurements of that binary failed admission, and the checked-in release implementation is restored to the baseline. Final source also replaces `repeat().take()` with `repeat_n()` in
an existing XLSX `cfg(test)` helper to pass the strict test lint gate. Original failures, original/final test sources, the
correction patch and separate baseline reference receipt remain retained.
The restored source passes the same eight quality commands, including 4,757 successful test executions. Final quality and disposition are recorded in the evidence bundle.

OLE2/OOXML performance remains the active goal. ODF optimization is deferred
until it completes, and iWork is excluded. This batch does not establish
cold/range-source, native-producer, scaling or program-level completion.

See the [evidence bundle](../results/change-0531/README.md),
[native comparison](../results/change-0531/comparison.json),
[planning profiles](../results/change-0531/profile-comparison.md), and
[test correction](../results/change-0531/test-fix-review.md).

The [final native comparison](../results/change-0531/final-native-comparison.json) is the rejection authority. Its 57 adverse metrics and 101 same-build drift flags remain retained. Four additional final-binary planning profiles are diagnostic evidence; they cannot override the failed native gate. The next proposed work is [fresh OLE2 operation-local attribution](../results/change-0531/next-ole2-review.md).

All ten evidence-verifier components pass after removal of the owned temporary trees. The sealed bundle retains the raw measurements, rejection decision, and restored-source quality receipts.
