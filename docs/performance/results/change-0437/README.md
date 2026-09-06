# Change 0437: bounded fresh ODP titled-slide publication

The opt-in streaming API is retained for its measured memory benefit. All 180
streaming allocator samples peak at 420,352 bytes above operation entry,
compared with 23,973,505 bytes for large Builder (98.247% lower). Large normal
p50 is 1.574 / 1.588 times Builder in the two repeats. All 37 review flags are
retained; RSS stays around 84.5–84.7 MB with no improvement claim. See the
[change record](../../changes/0437-odp-bounded-plain-slides.md), `summary.json`,
and `retention-decision.json` for individual results and limitations.

This bundle compares the established ODP Builder before and after the change,
then compares the candidate Builder with the new plain-slide sink provider.
The corpus contains 64, 4,096, and 8,192 slides. Each slide has an indexed title
and body cycling through plain ASCII, Unicode, escaped XML characters, and
mixed text. The formal corpus has no leading/trailing or adjacent control
whitespace. Production tests cover the supported control grammar and matched
audit refusals separately.

The original 32,768-slide proposal was refused by the Builder's default XML
attribute ceiling. It is not a completed measurement. The retained
`source-review/shape-boundary.md` records the corrected matched shapes; neither
the Builder limit nor its audit was relaxed.

## Measurement and identity scope

The normal operation timer includes fresh title/body String creation, authoring,
complete ZIP publication to a hashing discard sink, and release of operation
buffers. The streaming role creates source items lazily inside the operation.
Sink/context setup and finalization, corpus generation, archive/XML/semantic
readback, and binary hashing are outside that timer. Allocator vectors use the
same operation boundary; process RSS is recorded separately.

The forward/reverse order is A1/B1/C1/C2/B2/A2: before-buffered,
after-buffered, after-streaming, then their reverse. Each phase has six fresh
processes (normal/allocator × three shapes), each with three warmups and 30
retained samples. CPU 2 and one worker are fixed. This yields 36 reports and
1,080 samples; six large normal profiles follow the complete matrix. Eighteen
three-sample pilots and two baseline preparatory profiles are separate evidence.

Within each role, exact archive, content, semantic, and sink identities must
match across modes and repeats. The before/after Builder control retains exact
archive identity. Cross-API gates require decoded title/body, page order and
geometry, member topology, manifest bindings, and fixed styles/meta identity.
Streaming ZIP framing may differ. The verifier independently derives text
digests and checks report/corpus/sample bindings.

Profiles cover the complete fresh executable invocation, including setup,
corpus generation, warmups, operation samples, and hashing. The profile
receipt's historical scope string is overbroad: the independent copied Python
oracle runs afterward, outside perf and GNU time. Symbolization is also outside
the workload. Whole-process samples cannot establish an operation-only Amdahl
fraction. L1 zero readings do not establish zero misses; LLC is not captured.

The provider's 4,096-byte fragment window is a formal configuration, not a total
heap bound. Its execution Memory reservation models retained provider buffers
and excludes caller inputs and additional ZIP/auditor allocations. Work counts
content shell/slide XML and fixed styles/meta XML, excluding ZIP framing,
manifest, compression, and auditing. XML limits remain independent. Cooperative
cancellation can be observed after one bounded 256-byte ordinary-text span.

## Reproduction

Build both revisions with Rust 1.98.1, four Cargo jobs, incremental disabled,
`RUSTFLAGS=-Cforce-frame-pointers=yes`, and
`CARGO_PROFILE_RELEASE_DEBUG=1`. Use the exact `checks/before-build.json` and
`checks/after-build.json` commands and their retained source manifests.
The baseline implementation is commit `203c50a848204ed51013b0555f61a3c7e75644c9`;
the candidate is `4560c9478` (full identity is in the build receipt).

For a fresh capture destination, `formal-suite.py` runs the frozen phases and
profiles serially through `check.py`. The individual command receipts retain
every argument, executable hash, source manifest, environment, and artifact
digest. Drivers refuse to overwrite completed captures.

To verify retained data after cleanup, from any working directory:

```sh
python3 -B /absolute/path/to/change-0437/verify.py --portable-check --require-inventory --stage final
python3 -B /absolute/path/to/change-0437/lifecycle.py --stage final
python3 -B /absolute/path/to/change-0437/decision.py --check
python3 -B /absolute/path/to/change-0437/portable-probes.py --stage final
```

`summary.py --check` independently rederives all normal statistics, throughput,
allocator balances, RSS deltas, and review flags. `decision.py --check` rederives
the scoped retention decision. Failed checks remain retained with their
resolutions in `planned-checks.json`. `SHA256SUMS` covers every retained file;
`compression.json` binds stored gzip artifacts to original logical sizes and
hashes. Temporary measurement binaries are intentionally excluded from the
portable dependency set. Native Office/LibreOffice validation was unavailable;
this bundle makes no visual-rendering claim.

## Completed cleanup and portable proof

Copied replay and all eight corruption probes passed both before and after
cleanup. The twelve frozen task scratch directories contained 1,832,195,611
regular-file bytes and were removed. Both shared target directories retained
their identities, and the user goal retained SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`.
`checks/cleanup-inventory.json` binds the removed paths and preserved objects.

The initial RSS reader missed tab-indented GNU-time values. The corrected
reader requires one value and includes all RSS comparisons. Its original
source and derived results remain in `versions`; no workload was rerun or
measurement edited for this correction. The final lifecycle check rederives
the corrected summary and protocol-bound decision.
