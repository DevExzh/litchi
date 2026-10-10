# Resource-specific PPTX slide-root admission

Native eager edits now admit the generated one-slide, 10,000-textbox corpus
that failed in record 0156. The slide-root classifier has an explicit
1,000,000-node ceiling aligned with the bounded slide readers. Actual notes
roots keep their 100,000-node ceiling. Byte, depth, attribute-count and
attribute-byte limits are unchanged.

Raw and already-processed root classifiers preserve typed `Limit` and
`Allocation` errors through Transitional/Strict retries. Malformed processed
roots retain their negative proof representation; only successful root
classifications enter the capture-local memo. Slide capture propagates typed
resource failures without replacing them with “invalid sld root or namespace”.
Relationship checks and deferred authored-name error precedence remain intact.

## Verification

The native PPTX/encryption suite passed 1,665 tests/doctests. New coverage
checks exact independent slide/notes node ceilings and a native eager edit of
the first and last of 10,000 textboxes, source immutability, serialization and
reopening. Raw/processed classification checks retain malformed-root behavior
and verify typed resource refusal. The existing 16 MiB root-vs-64 MiB part
limit regression now requires the preserved typed error.

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,107 tests/doctests.
Native and facade all-target Clippy, plus the standalone benchmark harness
Clippy, passed with warnings denied. The full PPTX API glossary profile passed
70 vocabulary bans using generated rustdoc JSON. Its first invocation could
not create the default scratch directory in the read-only worktree parent;
the explicit `/tmp` target-directory rerun passed.

## Measurement

The existing direct-command batch harness adds optional `--owners eager,source`
selection, defaulting to both owners. This permits a focused rerun of the two
eager 10,000-command scenarios refused by native admission in record 0156.
Route names, timing boundaries, native 256-target chunks, atomic facade command
batches and equivalence checks are unchanged. Three ABBA rounds produce six
samples per leg, pinned to CPU 0, with no concurrent build or test job.

Both cases completed successfully with exact-byte native/unified output
equivalence and one native commit per unified batch. Each native leg uses 40
bounded setter calls; each facade leg preserves 10,000 per-command results.

| Operation | Native median (ms) | Unified median (ms) | Unified/native |
|---|---:|---:|---:|
| text | 2639.203 | 2696.439 | 1.022 |
| text + x | 7481.092 | 7602.463 | 1.016 |

[Machine-readable evidence](0157-pptx-command-batch-results.json) retains all
raw samples, corpus identity, build command, source revision, dirty-source
manifest digest and binary size/SHA-256. No compiled binary is committed.
The generated 91,398-byte corpus has SHA-256
`3718c459492e44f5b68acc66b83a862e95b5998c172b8fd6350bc05535e1aacd`,
matching the previously refused corpus.

This closes the eager admission finding. It does not establish linear scaling;
record 0156's source 10,000/1,000 latency growth remains unresolved, and neither
staging algorithm changed here. Repeated full-slide work in bounded native
setter calls is the next performance target. Allocation/RSS, wider corpora,
other command families, remaining formats and full goal acceptance remain open.

## Reproduction

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0157-target cargo test --offline -p litchi-pptx --features encryption
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0157-target cargo test --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0157-target CARGO_NET_OFFLINE=true cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-command-batch/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0157-target/release/unified-pptx-command-batch-baseline --counts 10000 --owners eager --rounds 3
```

Set provenance environment variables as in the harness README. Native corpus
tests require LibreOffice fixtures `tdf163852.pptx`, `tdf164622.pptx`, and
`tdf169496_hidden_graphic.pptx` under `sd/qa/unit/data/pptx`; this turn copied
them from the base checkout only while testing, preserving that checkout.

## Repository and cleanup

The rewritten base remains an ancestor, so no rebase is required. The repository
boundary audit passed for 66 packages and 268 internal dependency declarations.
Scoped formatting and diff checks passed.

The artifact audit found no prohibited evidence or compiled binaries among
694 changed/untracked paths. Both required staged-file guards printed
nothing. Scoped staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on its
read-only filesystem; no commit was created.

All verification and measurement processes ended before cleanup. The three
turn-owned fixture copies were hash-checked and removed, along with the
generated root lockfile, temporary logs/raw JSON and build output. The small
evidence JSON retains every timing sample and binary provenance. No temporary
worktree was created; inherited files and the base checkout were preserved.
The full goal remains active.
