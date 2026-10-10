# One publication per slide for PPTX selector fanout

PPTX `File::set` selector fanout now submits every same-slide text or combined
text/geometry selection through one native publication. The native chunk APIs
introduced in record 0158 already serve command coalescing; this record adopts
them for selectors such as `/slide[1]/*`. The facade retains one primary edit
and commits only after every selected slide succeeds.

Input chunks respect both the 256-target capacity and the native aggregate
text-byte limit. Uniform selector values let the facade calculate the byte
boundary before staging. A selection fitting one chunk retains the ordinary
native setter and avoids allocating a chunk-view vector. Empty selections,
semantic limits, source freshness, patch proofs and late-failure rollback
retain their existing behavior. No public vocabulary or capability entry changed.

## Verification

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,110 tests/doctests.
New coverage checks every resulting shape in a 10,000-shape mixed update for
both owner modes and verifies source immutability. Another regression saves and
reopens three text values whose aggregate exceeds one chunk's 8 MiB budget,
with and without geometry, on both owners. Existing transaction replay/inverse,
semantic-limit and late malformed-transform tests remain green.

The first large-text fixture used two strings of 4 MiB plus one byte. Staging
succeeded, but source serialization hit the independent OPC XML token limit
of 4 MiB. The final fixture uses three strings below that token limit, while
retaining aggregate text above 8 MiB. This record does not relax save validation
or claim arbitrary single-token size support. The returned facade code for
that save refusal was `corrupt_file` with a native `TokenBytes` limit source;
resource-error classification at that publication boundary was corrected in
[record 0160](0160-publication-resource-errors.md).

Facade and standalone harness all-target Clippy passed with warnings denied.
Facade rustdoc passed with warnings denied. Repository boundaries passed for 66 packages and 268 internal dependency
declarations. Scoped rustfmt and diff checks passed. The OfficeCLI capability-delta check passed against schemas in the
base checkout; this checkout has no local OfficeCLI help schemas. No schema
or capability delta is introduced. Native leaf code and Python bindings are
unchanged; this record does not claim rerunning the all-format/Python/fuzz gates.

## Measurement

The standalone `unified-pptx-fanout-publication` harness uses new routes ending
in `fanout_publication_v2`; the older repeated-call comparator is unchanged.
The native leg opens one primary edit, submits all 256-target borrowed chunk
views in one setter call, and commits once. The unified leg executes one
wildcard `File::set`. Payload/property preparation, opening, output serialization
and teardown are outside the timer; the native borrowed chunk-view vector is
constructed inside it.

All 20 scenarios ran on CPU 0 with three ABBA rounds, six samples per leg,
and no concurrent agent build/test/validation job. Warm-up outputs matched
exact bytes in every case. Native timed samples each changed the expected
shape count and used one setter call and one commit.

| Owner | Operation | Shapes | Native median (ms) | Unified median (ms) | Unified/native |
|---|---|---:|---:|---:|---:|
| eager | text | 1 | 0.302 | 0.321 | 1.065 |
| eager | text_x | 1 | 0.331 | 0.351 | 1.060 |
| source | text | 1 | 0.053 | 0.056 | 1.049 |
| source | text_x | 1 | 0.086 | 0.084 | 0.978 |
| eager | text | 10 | 0.412 | 0.455 | 1.104 |
| eager | text_x | 10 | 0.580 | 0.624 | 1.076 |
| source | text | 10 | 0.268 | 0.272 | 1.013 |
| source | text_x | 10 | 0.432 | 0.439 | 1.016 |
| eager | text | 100 | 1.437 | 1.730 | 1.204 |
| eager | text_x | 100 | 2.936 | 3.225 | 1.098 |
| source | text | 100 | 2.326 | 2.366 | 1.017 |
| source | text_x | 100 | 3.791 | 3.817 | 1.007 |
| eager | text | 1000 | 11.491 | 14.196 | 1.235 |
| eager | text_x | 1000 | 26.256 | 28.933 | 1.102 |
| source | text | 1000 | 22.985 | 23.211 | 1.010 |
| source | text_x | 1000 | 37.671 | 37.933 | 1.007 |
| eager | text | 10000 | 111.829 | 138.395 | 1.238 |
| eager | text_x | 10000 | 259.757 | 287.424 | 1.107 |
| source | text | 10000 | 229.745 | 231.257 | 1.007 |
| source | text_x | 10000 | 376.367 | 378.313 | 1.005 |

| Owner | Operation | Native 10k/1k | Unified 10k/1k |
|---|---|---:|---:|
| eager | text | 9.732 | 9.749 |
| eager | text_x | 9.893 | 9.934 |
| source | text | 9.996 | 9.963 |
| source | text_x | 9.991 | 9.973 |

At 1,000 and 10,000 shapes, every unified/native ratio is below 2x. Unified
10k/1k growth is 9.749–9.973x for 10x shapes. These results support linear
selector-fanout growth in this generated one-slide corpus. One-shape wildcard
ratios range from 0.978 to 1.065; ordinary direct-ordinal single-set overhead
remains the distinct record 0155 scenario. These observations do not establish
real-deck performance, allocation/RSS, managed 10k capacity, or full acceptance.

[Machine-readable evidence](0159-pptx-fanout-publication-results.json) retains
all raw samples, corpus sizes/SHA-256, the source revision and dirty-source
manifest digest, build command, and binary size/SHA-256. The source manifest
was unchanged between release build preparation and completed measurement.
The binary is not committed.

## Reproduction

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0159-target CARGO_NET_OFFLINE=true cargo test --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0159-target CARGO_NET_OFFLINE=true cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-fanout-publication/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0159-target/release/unified-pptx-fanout-publication-baseline --counts 1,10,100,1000,10000 --rounds 3
```

Set the provenance environment variables described in the harness README.


## Repository and cleanup

The rewritten base remains an ancestor; no rebase is required. The full goal
remains active. Remaining format coverage, broader CRUD, Python/CLI migration,
and complete performance acceptance are not claimed by this focused milestone.

Both required staged-file guards printed nothing. Scoped staging failed because
Git could not create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
on its read-only filesystem. No commit was created by this agent. The pending
record 0158 cleanup/provenance corrections are preserved alongside this work.

All verification and measurement jobs ended before cleanup. Turn-owned build
output, logs, original measurement JSON and temporary provenance files were
removed. The inherited root `Cargo.lock` was preserved. No fixture copies or
temporary worktrees were created. The base checkout was not modified.
