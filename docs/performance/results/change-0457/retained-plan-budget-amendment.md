# Retained authored-fragment budget correction

Final review found that the initial insertion plan charged fragment length
during preparation but retained the caller's `Vec` capacity after that local
reservation ended. A tiny valid XML fragment could therefore carry a large
uncharged allocation in a long-lived plan. The original source epoch's captures
remain in `candidate/`; they do not prove this managed-budget lifetime property.

The correction adds a private capacity query to `AuthoredXmlFragment` and a
retained memory lease to `SourceContentInsertionPlan`. Preparation reserves the
actual capacity before reading the source, separately from transient scan
memory. The fragment drops before its lease, both in returned plan state and
on preparation failure. Publication independently reserves the retained
capacity under its supplied execution options, because those options may use a
different budget. Using the same budget for both phases conservatively charges
the overlapping leases twice; the public method documents this behavior.

The focused integration suite passes all 13 tests, including two regressions
using a valid fragment with 64 MiB spare capacity: an insufficient budget
refuses before source reads, and an admitted plan keeps the capacity charged
until drop releases it. See `checks/retained-plan-tests.json` and its log.
The full common suite subsequently passed 499 tests with one ignored. Clippy
flagged the deliberate drop-order shadow as a redundant local; naming the
parameter `input_fragment` preserves that move and drop order without a lint
allowance. `strict-retained-plan-r1` passed. The original failed lint receipt
and the distinct source manifests remain intact.

The final source epoch is captured separately in `candidate-final/` and
`native-final/`. Its build, output fixture identities, measurements, comparison,
and final validation are bound independently. Original reports, protocols,
failures and binaries are retained as historical evidence; the correction does
not silently relabel their source identity.
