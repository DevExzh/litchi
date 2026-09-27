# Baseline qualification and oracle correction

The initial probe stopped at Unicode write qualification: the fixture contains
one authored trailing space per text box, while the public reader trims each
text atom. The 64-box fixture therefore differed by 64 bytes in semantic
extraction. No production source or authored fixture was changed to accommodate
this result. The initial probe, binaries, build receipts, six successful
qualification reports, and failed qualification log remain archived under
`initial-probe`, `build-before-initial`, and `qualification-initial`, with binary
relocation identities in `qualification-correction.json`.

The corrected probe independently compares authored strings with untrimmed
decoded text atoms from the published ClientTextbox records and separately
compares reader-visible text with the documented trimming behavior. It uses
public OfficeArt/PPT record parsers; this is a direct raw text-atom check, not
an independent CFB or PPT parser. Header and payload bounds, UTF-16 validity,
ASCII byte-atom validity, record count, and traversal depth are checked.
Authored spaces and rich paragraph separators remain observable. Reopen and
both text checks run outside the measured interval.

Both corrected baseline binaries build from identical frozen probe sources
and the retained lockfile. All ten baseline cases pass. All six successful
initial reports have the same source and output identities as their corrected
counterparts. `authored-fixture-review.json` additionally records unchanged
source-level fixture generation functions. These observations qualify the
fixed comparison protocol before any production candidate is applied.

An extra probe unit-test run with all features exposed an invalid test setup:
the synthetic global counter tests assume only explicit counter calls, while
the installed global allocator also counts real allocations made by their
barriers, threads, and test harness. The first peak assertion failed and
poisoned the test mutex, causing six subsequent failures. The complete
29-pass/7-fail attempt remains under `probe-tests-initial`.

The corrected test driver runs all 31 default-feature tests with the synthetic
counter environment, then the 10 real allocator-wrapper and oracle tests with
all features. Only the 26 synthetic counter tests are filtered from the second
run; all have already run in the first. Both runs pass with one test thread.
No frozen probe or production source was changed for this test-driver repair.

Before candidate application, the payload/write Heaptrack trace records 192
allocation events and 7,680,000 requested bytes under
`convert_shape_to_escher_with_sound_mapping`, matching 64 strings × 40,000 bytes
× three writes. Trace totals cross-check against the allocation histogram and
print summary. This whole-process ancestry observation supports the proposed
transient-string target; it is neither timed phase attribution nor a measured
speedup. `baseline-attribution.json` binds the exact trace and totals.
