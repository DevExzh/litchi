# 0500: managed source-backed paragraph batches

0500 extends the format-owned `Edit::replace_body_paragraph_texts` seam to
managed source-backed document edits. The batch accepts a nonempty,
canonical, strictly increasing set of unique direct-body paragraph positions
and builds one final base-relative candidate instead of reconstructing the
document once per scalar replacement. Existing operation composition,
immutable source proofs, typed selectors and limits, semantic readback,
source-version checks, cancellation, and exact no-op restoration remain part
of the transaction contract. Paragraph run boundaries, formatting, drawings,
and unknown run XML remain under the scalar path's refusal and preservation
rules.

The managed batch retains its owners and reservations while planning the
candidate. Memory, object, and depth admission remains finite and releases
through the transaction owners; actual `Work` and `InputBytes` consumption is
monotonic. A failed selector, text validation, admission, source proof,
candidate construction, parse, or readback leaves the staged projection
unchanged. Publication and complete-artifact forward/inverse proofs remain
source checked. This is an explicit format API capability and does not add a
global scheduler, ambient capability, or package-ownership crossing.

## Baseline and measurement boundary

The isolated managed paragraph example compares the existing repeated scalar
route with the new batch route as two explicit API choices. Its deterministic
128- and 512-paragraph documents select 1, 8, or 32 paragraphs and contain a
256 KiB media member plus a 256 KiB opaque member. The providers are owned and
warm-file. The harness records full open/edit/commit/sequential-publication
latency and a separate edit interval; fixture setup and independent output,
semantic, preservation, and inverse verification are outside those clocks.
Warm files do not establish controlled-cold or native-producer behavior.

The fresh repeated-scalar baseline has 12 children, 720 measured samples, and
72 warmups. Its aggregate p50 lifecycle values are microseconds; owned and
warm-file are shown in that order. Edit share is the fraction of the timed
lifecycle spent in the edit phase:

| Paragraphs | Replacements | Owned p50 | Warm-file p50 | Edit share | Zero-edit arithmetic bound |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 128 | 1 | 650.082 | 638.722 | 42.4% / 42.9% | 1.74x / 1.75x |
| 128 | 8 | 1,518.316 | 1,530.175 | 75.2% / 76.1% | 4.04x / 4.18x |
| 128 | 32 | 5,516.121 | 5,919.412 | 93.2% / 93.3% | 14.71x / 14.89x |
| 512 | 1 | 2,214.499 | 2,256.748 | 45.9% / 45.3% | 1.85x / 1.83x |
| 512 | 8 | 5,268.439 | 5,367.890 | 77.2% / 76.7% | 4.38x / 4.29x |
| 512 | 32 | 16,696.953 | 17,268.215 | 92.7% / 92.9% | 13.77x / 14.12x |

The zero-edit column is an arithmetic upper bound under the hypothetical that
all edit time disappears; it is not a forecast or a measured speedup. The
baseline shows why repeated reconstruction is a candidate hotspot, especially
at 32 selected paragraphs. The completed after phase adds 24 children covering
repeated and batch routes; the combined evidence contains 36 children, 2,160
measured samples, and 216 warmups. Every matched output identity and source,
budget, cleanup, and preservation oracle passes. `Work` charges may differ as
the batch removes reconstruction passes; `InputBytes` remains equal in the
matched rows.

The primary comparison is batch versus repeated scalar in the same final
executable, an explicit API-choice comparison. Lifecycle p50 speedups for
K=8 are 2.346x, 2.324x, 2.341x, and 2.342x for p128 owned/file and p512
owned/file. K=32 lifecycle speedups are 6.997x, 6.947x, 7.088x, and 7.138x
in that same order; edit-phase speedups are 12.640x, 12.737x, 13.363x, and
13.456x. K=1 owned p128 is the small-input exception: lifecycle p50 rises
7.34% for batch, with the other K=1 rows within roughly one percent.

The batch-versus-final-scalar comparison retains five aggregate flags: p128
K=1 owned lifecycle p50 (+7.34%), p95 (+5.61%), mean (+6.33%), and output
throughput (−5.96%), plus p512 K=8 warm-file RSS (+5.77%). The p128 K=1
owned repeat flags are retained in full: repeat one p50/p95/mean rises
6.76%/5.13%/5.04%; repeat two p50/p95/p99/mean rises
7.50%/7.06%/6.04%/7.64% and throughput falls 7.09%. The cross-route
batch-after versus scalar-before review retains warm-file RSS flags at p512
K=8 (+7.89%) and K=32 (+5.83%). These are descriptive flags on a shared host;
larger-selection interpretation stays scoped.

The historical same-API scalar before/after review retains two p512 K=8 p99
flags: +6.72% owned and +15.31% warm-file, with the corresponding repeat tails
preserved. It is a separate control and does not turn the batch-versus-scalar
API-choice result into a historical same-method speedup claim.

The adverse phase review associates the p128 K=1 batch lifecycle flag with a
publication p50 increase of 15.23% while the edit p50 falls 0.35%; this is an
observed phase association, not a causal explanation. In the historical scalar
p512 K=8 tails, the worst owned sample moves publication from 1.196 ms to
1.546 ms while edit moves 4.102 ms to 4.107 ms; warm-file publication moves
1.264 ms to 2.112 ms while edit moves 4.144 ms to 4.130 ms. Aggregate edit p99
changes are +0.82% and +0.19%. All phase and tail flags remain retained.

Charged `Work` for the K=32 batch is 0.698x of repeated scalar at p128 and
0.946x at p512 (scalar/batch ratios 1.433x and 1.058x); `InputBytes` is
unchanged. These are accounting charges, not CPU measurements. Six
whole-child profiles include route-specific
preflights, setup, and output verification. The p512 K=32 batch versus
repeated profile changes cycles by −81.65%, instructions by −82.95%, and
branches by −83.26%; p512 K=1 batch cache misses rise 11.40%, while the two
scalar controls rise 11.32% and 31.30%. The profile rows are supplementary
shared-host evidence and do not provide edit-local CPU or allocation-count
attribution.

The production review confirms that decision-scan admission reservations are
dropped before reconstruction and that input reservations intentionally match
the scalar route's lifetime. The bounded selected-by-prior-operation metadata
matching cost is not separately metered in `Work`; charged `Work` must not be
read as exact CPU accounting.

XML forward/inverse preflight and complete-artifact inverse publication are
correctness gates outside the timer. Focused tests cover the complete-artifact
inverse path, but that proof is not asserted as a timed sample-level
measurement. The focused batch subset has 8 passing tests, the managed scalar
integration subset has 36, and the transaction subset has 17; these subsets
are separate and are not added to the full-suite totals. The all-target default
suite has 1,371 passing tests and no ignored tests; the all-features library
suite has 946 passing tests and no ignored tests; doctests have 74 passing and
31 existing ignored tests; and warnings-denied rustdoc passes. Final formatting
and all-target warnings-denied Clippy pass. The corrected downstream check,
`--no-default-features --features docx --lib`, passes. The initial DOCX
`--all-targets` attempt is retained as a feature-gating limitation: the
comprehensive DOCX example imports `litchi::ooxml_common`; enabling that path
also exposes `core_props_office` imports from `litchi::pptx` and `litchi::xlsx`,
whose features were disabled. These are example feature-gating limitations,
not a production defect inference. No manifest edits were made, and this
retained failure is not a candidate regression. The CRUD coverage validator
still reports contract-only mappings
without a full-run timing report, so this isolated warm synthetic scenario
does not close the one-percent, native-producer, cold-source, or broad CRUD
requirements.

Final cleanup removed 1,675,444,224 allocated bytes using unique device/inode
accounting and retained two frozen replay binaries. The shared workspace target
and the 16 protected unrelated files remain outside this change's cleanup
scope.

The full non-iWork `docs/GOAL.md` objective remains open.
