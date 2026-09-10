# Interpretation of the frozen 0501 protocol

The frozen protocol JSON is retained byte-for-byte. Its phrase "single-build
before/after" means one build per role: there are distinct before and after
executables, with the same harness source and build flags. The two reversed
workload orders within each role are not an interleaved ABBA experiment.

The supplementary profiles retain the originally frozen eight measured
iterations, one warmup, 99 Hz sampling, and DWARF callchains. Later suggestions
to increase sample frequency/count were not applied to that frozen protocol.
These are whole-child profiles, including expensive untimed corpus creation.
SHA symbols alone cannot attribute cost to the production touched digest.
The fresh captures have incomplete caller stacks; targeted 15.510–16.115%
attribution belongs to historical 0449 evidence, not these new captures.

The first owned profile export waited for the inherited Ubuntu debuginfod
server. The owned export subprocess was interrupted, its partial output and
failed receipt retained, and the same recorded perf.data was re-exported with
local symbols. That export reused the same recording. The driver subsequently repeated the
owned profiling workloads twice: one attempt failed the verifier’s profile
dimension check, and the final attempt passed after its correction. All failed
attempts and raw recordings are retained. Subsequent exports disable debuginfod.
See driver-transition.json for exact history; these supplementary attempts do
not replace or add to the formal 240-sample before matrix.

The original frozen verify.py checks current source against each role and
therefore cannot validate both historical and candidate sources after an edit.
It is retained as the initial helper, and verify-final.py supplies independent
historical-source and candidate-source validation. No source freeze or raw
measurement is rewritten to work around this validator limitation.
