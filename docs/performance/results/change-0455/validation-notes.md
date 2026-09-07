# Validation and capture history

All ten release commands in `formal/run-checks.py` pass with the retained
candidate source manifest: 456 ZIP, 497 OPC, 854 PPTX and 381 harness tests,
2,188 total with 6 ignored. Strict lint, rustdoc with warnings denied, scoped
formatting, minimal workspace feature checks and crate boundaries pass. The
regression also passed with the 32 KiB baseline before candidate compilation.
The isolated sanitizer fuzz lock/build/smoke gates pass 1,000 runs with seed 455.

The six root preliminary baseline captures predate sink-summary instrumentation.
They remain untouched and are excluded from matched comparisons. Formal baseline
and candidate builds share the harness and new regression. Only the production
copy constant and the ZIP64 test-only bound differ. Initial formal baseline lanes
0–3 and 16–17 ran with the baseline checkout; later captures, including preserved
baseline binaries, ran with the candidate checkout. Build and capture source
identities are deliberately recorded separately.

The formal 24 lanes pass exact-output and work oracles (720 samples). Ordinary
confirmation adds 8 lanes (240 samples). Hardware-counter and instruction-profile
runs remain separate. All successful capture and release receipts report
unchanged sources. No measured Rust source changed after candidate compilation.

The first instruction-record command completed, but its `perf report` stalled
with remote symbol lookup configured. The coordinator stopped only that owned
report process. `formal/checks/stack-profiles.json` records the failed attempt;
its driver, log and partial reports are unchanged. The new
`profile-stacks-local.py` attempt disables remote lookup and succeeds for both
builds. Restricted kernel symbols and incomplete unwinding limit attribution.
The later fixed allocator-policy protocol was declared before its eight runs.
It is supplementary diagnosis and does not replace any original adverse result.

The final verifier replays command, binary, source, protocol, report and log
bindings; exact output oracles; the release/fuzz gate table; native proof;
profile artifacts; and JSON/Markdown derivations. It adds a largest-write versus
histogram consistency check without changing the frozen capture oracle. Negative
probes cover five report corruptions and two impossible maximum-write values.
Cleanup is authorized only after an actual precleanup pass. The separate portable
replay runs after the owned raw artifacts are inventoried and removed. Shared
build targets and user-owned `docs/GOAL.md` are retained.

Final precleanup and separate-copy portable verification pass. Cleanup removed
937 owned temporary files (1,032,052,420 bytes); shared targets were retained.
Run `python3 -B formal/verify.py --portable` from this bundle to additionally
validate the recorded portable pass, then `sha256sum -c SHA256SUMS` to check the
seal. No production sources changed after the measured candidate build.
