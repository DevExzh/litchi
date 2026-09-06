# 0430 source and evidence review

This is a static review of the 0430 capture driver, receipts, build inputs,
portable verifiers, and scope documents. I did not run the capture, a verifier,
Cargo, tests, or a profiling command.

The scientific scope is stated carefully. `source-audit.md` separates the
whole-process `perf record` region from the four narrower API clocks and keeps
corpus construction, observers, checks, and drops outside those clocks. The
README and `goal-scope.md` also say that the three warmups cannot be separated
from retained stacks, that both providers are warm observations, that the
counts are unweighted `cycles:u` samples, and that no speed, latency, memory,
I/O, or scaling result follows. The build receipt binds the borrowed binary,
source revision, Rust 1.98.1 release build, `CARGO_PROFILE_RELEASE_DEBUG=1`,
and `-Cforce-frame-pointers=yes`; each profile receipt binds CPU 2,
`cycles:u`, 499 Hz, `--call-graph fp`, provider, corpus, warmup, and sample
arguments. The corrected stack parser uses exact `deflate_medium` ancestry,
accepts address zero, and keeps the historical comparison inside the bundle.
The bound analysis invocation in `verify.py` is the right custody boundary:
it reruns `profile-analysis.py --check` over the retained symbolized stacks,
while the report verifier checks each decoded lifecycle report.

## Conditions before sealing

The directory passed through intermediate sealing states while analysis and
portable replay were being added. The final evidence policy intentionally
retains the failed v1/v2 replay attempts as expected `failed` receipts, along
with the passing v3 replay and cleanup receipts. All of those receipts and
their logs must remain in the exact-check policy and the lossless inventory;
failed verifier attempts are historical evidence, not a production failure.
The four core source-custody receipts (`frame-pointer-profiles`,
`symbol-identity`, `profile-analysis`, and `analysis-portable`) must remain
required. The new
`replay-copy.py` has the correct ordering for the live-receipt race: it copies
the sealed bundle before asking `check.py` to create its receipt. Its generated
receipt and log must themselves be included in the final seal (or explicitly
kept outside the exported bundle); otherwise a later portable replay is
working against a stale inventory. The final post-cleanup receipt must be
added before the last inventory refresh.

## Build-input custody limitation

The retained source manifest is intentionally limited to files ending in
`.rs`, `.toml`, or `.lock` (`check.py:21-37`). That is enough to identify the
Rust source revision, but it is not a complete set of inputs for rebuilding
the borrowed executable. The directly built `tools/perf-baseline` crate
embeds `test-data/rtf/watermark.rtf` at `src/lib.rs:13660-13668` and
`test-data/poi/test-data/spreadsheet/54016.xls` at
`src/xls_numeric.rs:25-27`; these bytes are absent from
`input-source-manifest.json` and from `input-build-0429.json`'s listed
additional artifacts. Other linked crates also contain compile-time resource
templates. The corrected README explicitly scopes recapture to the preserved
external workspace/asset tree and says the bundle is not a hermetic rebuild
artifact. This does not weaken portable
replay, which deliberately needs neither the binary nor a source checkout.

The `additional_artifact_sha256` map in `input-build-0429.json` has the same
boundary: `verify.py` validates the top-level fields' shape and uses only the
copied `check.py` hash for the build receipt; it does not require or validate
the other named historical helper files. Those entries are therefore metadata
pins rather than independently retained custody. The corrected README labels
them as historical metadata and refers full 0429-history replay to that
original bundle. The
same applies to the `validation_amendment` path/hash: the field is validated
as metadata, but its referenced file is not present in the bundle.

## Remaining evidence limitations

The retained report logs visibly say `Total Lost Samples: 0`; the evidence
documents this as a raw-log observation, while `profile-analysis.py` checks the
record sample count and decoded block count rather than claiming an independent
lost-sample invariant. Likewise,
`machine-current.json` records platform, uname, affinity, and the hash of the
prior machine record, while the profile receipts bind the requested CPU and
perf arguments; it does not bind a current CPU/perf-version record into each
receipt. Since this batch makes no cross-machine performance claim, this is a
provenance limitation rather than a reason to reinterpret the samples.

Subject to the final post-cleanup receipt and inventory seal, the attribution
and non-optimization claims are supported. The compile-time assets,
historical helper hashes, and current-machine binding remain explicit
non-hermetic recapture/provenance limitations; they do not invalidate this
unchanged-binary attribution batch.
