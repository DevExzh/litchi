# 0499 evidence

The before revision is `b27718fe8`. `before-freeze.json` identifies the retained
0498 production executable and unchanged benchmark harness. `before/` contains
thirty fresh controls, not a relabeling of prior timings. Each child records two
repeats of three warmups and thirty measurements. Both phases use plain serial
reads or the production batch API with optional operation accounting disabled.

`capture.py before` and `capture.py after` run the fixed matrix. They require a
matching phase freeze receipt and refuse to overwrite existing child outputs.
`compare.py` checks raw hashes, sample identities, byte/source/work counters and
released budgets before computing matched before/after deltas. It retains all
adverse changes greater than five percent, with RSS counted once per child.
The few-large corpus contains four selected 1 MiB stored Parts; many-small
selects sixty-four 16 KiB stored Parts from a 512-Part archive. The unchanged
harness generates and hashes these fixtures deterministically.

`trace-threads.py PHASE` records clone/clone3 calls in separate traced children.
The before trace observes 64 successful thread creations for one many-small
batch at width four and four for one few-large batch. Traced elapsed times are
not accepted timing evidence. `profile.py PHASE` supplies separate whole-child
CPU/scheduling counters, including fixture setup and verification. The latency
matrix excludes those phases. Warm files and a fixed-delay short-read provider
do not prove cold-cache or real remote-service performance.

`run-gate.py NAME` executes commands from `gate-commands.json` with Rust 1.98.1,
release debug information disabled, incremental compilation disabled, two
build jobs and the dedicated 0499 target. Gate CPUs16–31 and measurement
CPUs0–7 remain on a shared host; memory/filesystem/host isolation is not claimed.
Initial failed checks are preserved as attempt files. `protected-work.json`
records unrelated work outside this batch's editing, staging and cleanup scope.

The RAII loader guard repairs cache-flight cleanup when provider code unwinds.
Normal results disarm it without an additional cache lock. Panic recovery is
only an unwind-build property; the workspace's ordinary release profile uses
panic=abort. No claim of recovery from an abort is made.

The completed matched matrix is in [comparison.md](comparison.md) and
`comparison.json`. [after-scaling.md](after-scaling.md) compares batch reads
against ordinary serial reads in the same final executable, including Amdahl
model limits and every adverse flag. Regenerate it with `analyze-scaling.py`,
which reuses the committed 0498 analyzer. `summarize-profiles.py` regenerates
`profile-comparison.json` from the supplementary raw counters.

The after trace records four thread creations for the many-small width-four
batch, down from 64; the four-Part single-wave case remains at four. This is
direct thread-creation evidence, not a decomposition of elapsed time. The
matched matrix retains five aggregate tail-latency flags and ten per-repeat
latency/throughput flags. Local small-Part batches remain slower than ordinary
serial reads despite materially improving over the prior batch implementation.

`candidate-source.json` seals the source and unchanged harness used to build
the final executable. `before-freeze.json` and `after-freeze.json` identify the
two retained replay binaries. `verify-evidence.py` independently checks the
captures and calculations. `cleanup.py` removes only this batch's dedicated
scratch after checks finish, preserving those executables; `cleanup.json`
records the result. The final inventory hashes retained evidence separately
from source and unrelated-work checks.

Final verification passes for all sixty children and all required gate
receipts. `verify-scaling.py` independently recomputes the 72 aggregate and
per-repeat serial/batch comparison scopes. Cleanup removed 1,060,274,176
allocated bytes and retained the two frozen executables. All sixteen protected
files remain unchanged. Run `inventory.py` to regenerate `inventory.json`
after intentionally changing evidence; it excludes only its own output.
