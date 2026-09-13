# 0545: exact-bound XLSX scanner diagnostic rejected

An isolated replacement for the rejected 0544 admission helper preserves its
exact event bound, but regresses sparse XML and early cap exits. Production is
unchanged. This batch avoids promoting dense-only microbenchmark gains into a
full-workflow claim.

Seven frozen inputs, one release binary with equal function call boundaries,
CPU 2 and A1/B1/B2/A2 yield 2,800 retained single-call samples. Dense diagnostic
96/128 and cap160/164 scans improve p50 by 37.2–44.9%; cap256 improves 10.2–11.0%.
Sparse 1 MiB text takes 42.98–44.18x baseline p50 time; the comment-prefix cap
case regresses 22.61–22.73%. All 16 adverse metric comparisons and two drift
flags are reviewed individually. The actual inner loop is scalar, so attempted
auto-vectorization is not evidence of SIMD execution. No allocation, RSS,
hardware-counter, cold-cache or total-workflow claim is made.

The safe 4,096-byte adjacent-pair reduction passes exact-bound differential and
pinned NsReader oracle checks, including cap/chunk boundaries and malformed
bytes. Four diagnostic quality checks pass (fmt, two tests, warning-denied
Clippy, release build); every native assertion passes. There are no production
source or workspace dependency changes requiring renewed workspace suites.

The full [evidence bundle](../results/change-0545/README.md) includes frozen
inputs, fixture generator/provenance, source/lock, all samples and command
receipts, measured disassembly, deterministic analysis, independent design
review and a strict replay verifier. The owned scratch/build tree is removed
after binding the executable hash; evidence remains reproducible.

Next investigate exact counting with sparse skipping and cheap early exits.
Non-short-circuit reductions are an unmeasured idea requiring assembly and the
same adversarial scan checks first. A later integrated candidate still needs
all original workflow, allocation, refusal and cap gates. The coarse one-pass
bound is unsuitable because it excludes the primary 128 shape. OLE2/OOXML
optimization remains active; ODF is deferred until that goal completes.
