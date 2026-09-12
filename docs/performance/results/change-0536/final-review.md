# 0536 final review

The measured collector cold-error-helper candidate is rejected. All eight
primary p50 rows improve, but only four meet the frozen 3% threshold. The
XLS constructor inclusive-Ir gate also fails in both repeats. The candidate
reduces collector self Ir by only 0.004642% in XLS and 0.001795% in CFB
few-large; shrinking the collector from 1,436 to 1,238 bytes does not establish
an end-to-end gain. The normal executable grows by 2,608 bytes.

The exact candidate patch is retained as evidence. Reversing that patch
restores the final source manifest byte-for-byte to baseline
9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75.
No production, harness or test changes remain from this experiment.

The native matrix retains 48,000 samples. The separate allocator lane retains
1,440 samples; allocation calls, reallocation calls, allocated bytes and
incremental peak vectors are identical in all 24 paired rows. All 32 matched
adverse and 60 same-build variation flags are individually reviewed, with no
sample removal or causal attribution. The 16 profiles retain 80 timed and
12 separately classified setup dumps. Four hardware captures are whole-child
diagnostics; they do not establish operation-local hardware behavior.

Candidate quality passed all 14 gates and 4,382 test executions. Exact-source
custody separately permits reuse of 14 prior baseline quality gates and
4,382 executions; those are not new executions. Restored final quality passed all 14 gates with 4,382 test executions,
for 8,764 new executions across candidate and final runs. Final evidence
verification passed and is recorded in [verification.json](verification.json). An early profile
replay refused the intentionally missing final baseline native repeat; after
ABBA completion it passed. An early verifier test refused the still-running
final quality inventory; no raw capture was retried or discarded.

OLE2 and OOXML remain the performance priority. ODF is deferred until their
optimization goal completes; iWork is excluded. This rejected experiment
changes the next optimization choice without completing the broader program.

All 96 receipt intervals passed serial and ABBA verification. Both owned
build paths and the coder draft directory were removed. Five tamper probes
passed. The final postcleanup receipt and SHA256SUMS bind the retained bundle.

Full postcleanup replay passed after correcting external report-output handling
in the two analyzers. Sealed internal writes and non-identical overwrites
remain refused; the focused replay regression probe verifies those boundaries.
Raw measurements and derived report contents remain unchanged.
