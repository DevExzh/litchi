# Verification disposition

The capture completed with 720 baseline rows and 3,450 candidate rows. The
candidate-wide preflight covered all 115 named cases before baseline timing;
the early preflight receipt and its build/output logs are retained under
`preflight-before-timing/`. Baseline, early-preflight, and candidate targets
all have `removed: true` cleanup receipts.

The frozen `performance/verify.py` command reached one read-bound assertion on
`cancellation-descriptive-avedev`: each timed child records one successful
resolver read before cancellation, then the shared cancellation remains set
for the remaining three internal repeats. The receipt therefore has
`reference_reads_p50 = 1` and `repeat = 4`, while integer floor normalization
records `reference_reads_per_repeat = 0`. The verifier's exact `(1, 1)` bound
does not represent this cross-repeat cancellation behavior. Both phases show
the same values; the typed `failure:cancelled` result, work receipt, and
allocator balance are valid.

The durable authoritative check is the batch-root `verify.py`:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-descriptive-statistics/verify.py
```

It requires exactly one total resolver read per timed child, four internal
repeats, integer-floor normalized reads/repeat `0`, and typed
`failure:cancelled`. Every other raw receipt, source-custody, numerical,
allocation-balance, and comparison check remains strict. All 4,170 accepted
measurement rows pass this verifier. The older `performance/verify.py` is
retained unchanged as a captured profile input; its normalization assertion
is a documented diagnostic defect, not the acceptance entry point.

Cancellation timing covers a mixed batch: cancellation is triggered mid-read
once, then the same execution context rejects the remaining three repeats.
It does not measure fresh mid-read cancellation on each repetition. No timed
row was rerun or discarded because of this normalization correction. The
frozen profile inputs and source manifests remain unchanged.
