# Change 0776 evidence

See the [integration report](../../0776-mce-expanded-attribute-duplicates.md).
The implementation candidate is `6ed76a881b`, based on `22f62ee328`.

The packet contains complete quality, paired release, differential, triage, and
allocation captures. Independent review accepted the candidate. `seal.json`
binds every other packet file; offline replay passes after target cleanup.

- `baseline-defect.json`: source-bound prior observation of the tree/stream gap.
- `first-source.json`, `quality-1/`: corrected first candidate and its passing
  nine-gate validation.
- `final-source.json`, `quality-2/`, `quality.json`, `review.md`: final source,
  nine passing serial gates, source binding, and review disposition.
- `quality-0/`: retained initial failure at the second gate; the test's
  unnecessary `super::stream` qualification violated `-D unused-qualifications`.
- `workspace-Cargo.lock`: archived workspace dependency resolution.
- `probe-src/`: identical standalone probe source used by both release legs.
- `measure-0.py`, `measure-0/`: first paired capture before the common-path
  revision, including the documented latency and RSS flags.
- `measure.py`, `measure-1/`: final paired capture for 18 cases, 108 timed
  processes and two differential runs at the final source.
- `analyze.py`, `analysis.json`: final input/outcome parity, statistics, and
  64,541-comparison differential.
- `triage.py`, `triage.json`, `changed-inputs-1/`, `generator-1.rs`: independent
  deterministic witnesses for all 248 generated aliasing changes across 86
  documents. The temporary generator binary was removed after capture.
- `allocations.py`, `allocations-0/`: one whole-process heaptrack sample per
  leg for worksheet, document, and the 32-name spill case. The allocation
  totals are reported in the integration report; instrumented elapsed times
  are excluded from native timing claims.
- `triage-0.json`, `changed-inputs/`, `generator.rs`: first successful witnesses.
  The initial failed classifier and its inputs remain under `triage-initial*`
  and `changed-inputs-initial/`; `analysis-corrections.md` explains corrections.
- `validate.py`: replay both analyses, quality/source custody, final duplicate
  witnesses, six allocation receipts, cleanup and the packet seal.
- `cleanup.py`, `cleanup.json`: three owned targets removed after source,
  binary and fixture hashes were checked.
- `next-0770.md`: read-only readiness notes for the separate 0770 follow-up.

The final quality run records 5,676 passing tests, 35 ignored tests, and 237
suites. The inherited `docx_styles_benign` row is a refusal control in both
legs (`style numPr is missing numId`), not an accepted DOCX success control.

The paired performance run uses three processes per case and leg on CPU 12 in
before/after/after/before/before/after order. Long URI cases use three samples
and one warmup; all other cases use nine samples and two warmups. Input
construction occurs before the `Instant` interval; parser and observer/digest
work are timed, while process startup and JSON report writing are outside that
interval. RSS is whole-process. These captures do not establish stable tail
quantiles, physical cold-cache behavior, range-source behavior, or CRUD
coverage.

The malformed duplicate control must be rejected by both processors on the
candidate. Generated differences require independent triage; valid-pair
failures are not silently accepted. The six-template alias-pair oracle is a
compatibility check rather than an exhaustive serialized-output oracle and does
not cover opaque extension output.

For offline replay, run:

```text
python3 -B docs/performance/results/change-0776/validate.py
```

Reproduction requires fresh worktrees at the recorded references, the archived
workspace and standalone locks, and local target paths adjusted in the capture
helpers. Run native workloads serially and preserve the recorded captures.
