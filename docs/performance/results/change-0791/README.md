# Change 0791: current PPTX capture profile

Current-source supporting evidence, with no production optimization adopted.
See [the report](../../0791-pptx-current-capture-profile.md) and
[source review](source-review.md) for findings and the next candidate.

From the repository root, replay all retained evidence without rebuilding:

```sh
python3 -B docs/performance/results/change-0791/validate.py --require-final-seal --check-workspace
```

`--check-workspace` additionally checks the original unrelated files and other
worktree inventory; omit it when those unrelated workspace details have changed. Production source and the 35
architecture inputs must still match the captured manifests. The seal binds
all packet payloads and six report/index documents. Subsequent legitimate
changes to those documents require checking this packet at its commit.

The three analyzers replay 36 native reports, six scoped Callgrind reports,
and two native frame-pointer profiles, checking 1,286 measured outputs against
sealed fixture oracles from 0780/0784/0785. Historical timing is not pooled.
The two root cross-checks independently census raw function rows and stacks.
The Callgrind reader is reused from the hash-bound 0784 parser.

`origin.json`, `host.json`, `architecture-inputs.json`, `inheritance.json`,
`build/source.json`, and build receipts bind source, environment, and binaries.
`plan.json` and `perf-fp-plan.json` were frozen before capture. Build and capture
scripts retain their exact commands. Execution order was `build.py`,
`build_fp.py`, `capture.py`, `profile.py`, `perf_fp_capture.py`,
`perf_decode.py perf-fp`, then `perf_frames.py`. These are provenance scripts,
not replay commands: they deliberately refuse existing outputs and depend on
the recorded machine/tool setup. All native work ran serially.

Compressed perf data and both decodes retain original and compressed hashes.
The owned build target was removed only after checking all three executable
identities; `cleanup.json` supplies the exact post-cleanup witness. No native
capture was retried or excluded. `replay-notes.json` records offline integration
corrections. Fresh probe formatting and release builds passed; inherited probe
warnings remain, and no fresh production test-suite claim is made.

Native wrapper perturbation, unresolved interior frames, guest-versus-native
instruction differences, and whole-child RSS accounting limitations remain
explicit in the report. No native phase fraction, allocation improvement,
production speedup, or new CRUD coverage is established by this packet.
