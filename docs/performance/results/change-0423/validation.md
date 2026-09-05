# Validation and remaining limits

The committed harness is `6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`.
Rust 1.98.1 was used for serialized Cargo checks. Source-hash receipts bind
each validation pass; raw failed invocations remain alongside the final checks.

| Check | Outcome | Evidence |
| --- | --- | --- |
| Public source-backed copy integration/adversarial tests | 59 passed | checks/source-api-tests.json and log |
| Existing owned phase/lifecycle and source-backed plain phase tests | 3 passed | validation-corrected-cross-copy-tests.log.gz |
| New source-backed plain/media lifecycle test | Passed | validation-ready-source-lifecycle.log.gz |
| Media oracle regression | Passed: four valid mutated archives plus expected part-count mismatch | validation-ready-media-oracle-regression.log.gz |
| CRUD coverage validator unit tests | 35 passed | checks/crud-tests.log.gz |
| CRUD index | 15 categories / 32 mapped selectors | checks/final-crud-coverage.log.gz |
| Registered performance claims | All nine strict replays passed | checks/claims.log.gz |
| Crate boundaries | Passed with pinned toolchain | checks/final-boundaries.log.gz |
| Warning-denied rustdoc | Passed | validation-lint-final-rustdoc.log.gz |
| Strict all-feature/all-target Clippy | Existing debt remains | validation-lint-final-clippy-strict.log.gz |
| Diagnostic Clippy | Passed; 27 existing library warnings | validation-lint-final-clippy-diagnostic.log.gz |

The diagnostic Clippy run allows the three established lints
`chunks_exact_to_as_chunks`, `clone_on_copy` and `needless_lifetimes`. It does
not satisfy warning-denied policy. The added archive mutation helper was
restricted to `cfg(test)` after the first lint pass identified it as unused in
normal builds; the helper's test body is unchanged. The scope receipt proves
that removing that one attribute reproduces the tested source. No source lint
allowance was added. The broader goal's strict lint requirement remains open.

The first compilation passes caught missing test imports, a stray owned-runner
variable and post-publication use of the consumed source editor. The first
runtime pass then rejected the new lifecycle's duplicated final gate. The
final implementation consumes the independent oracle's rich checks directly,
removes the duplicated order-sensitive media comparison and adds exact
presentation relationship cardinality. The existing owned runners and old
plain source-backed oracle were restored to their prior implementation.

Pinned formatting was applied to the changed scope. Selecting formatter hunks
initially removed an import block while skipping its relocated insertion;
compilation rejected the missing imports. A separately formatted crate import
block reconciles those imports and the new expectation/helper names. Both the
format receipt and import reconciliation are retained. The final lifecycle and
regression tests were run afterward.

The first boundary-check command selected the unavailable default 1.95 Cargo
component. Its pinned 1.98.1 rerun passed; this was an environment invocation
failure. Production source files remained unchanged through all 59 API tests.

The harness tests exercise actual API output. Independent output validation
binds each copied XML embed to the correct relationship, content type and
payload, checks all remaining slide XML bytes, exact added member/part sets,
relationship count, and untouched raw ZIP metadata/order. Separate hard gates
cover source stability, stale/foreign destination refusals and deterministic
publication. Mutation checks use valid rewritten archives, so malformed ZIP
framing is not their only rejection mechanism.

This batch changes benchmark and coverage code only. It adds no production
optimization, native Office fixture run, fuzz capture, hardware-counter trace
or general near-limit retention claim. Full non-iWork program completion is
not implied by these focused checks. Formal report, mutation and portable
replay results are recorded separately in the summary and checks directory.

The four fresh release functional checks (one retained sample, zero warmups)
passed against the same clean source and copied normal binary used for formal
capture. The first functional attempts exposed verifier integration errors:
the shared historical shape registry lacks the new source-backed selectors,
and its general configuration validator requires at least 15 samples. Those
failed reports and logs remain under `checks/functional-initial` and
`checks/functional-config-attempt`. The dedicated verifier now requires the
complete fixed configuration, with exact fields and values, for both contracts;
formal runs additionally use the shared configuration validator. Dedicated
source-backed row/corpus checks cover the new selectors. Existing shared tools
remain unchanged and hash-pinned. Both role reverifications and a fresh four-case
rerun passed. The media functional report also passed 12 rejection probes,
including configuration, corpus identity, phase ordering and logical read guards.

The authoritative behavioral test receipts are `validation-ready.json` and
`checks/source-api-tests.json`; final lint/rustdoc results are in
`validation-lint-final.json`. Despite their names, `validation-final.json` and
`validation-corrected.json` retain superseded failed attempts. The initial
`checks/script-syntax.json` is historical; `checks/script-syntax-final.json`
binds the final scripts.

All 16 formal runs passed, retaining 800 normal and 240 allocator samples.
Summary replay revalidated every report and its journal. Initial summary runs
failed on a premature transient journal-hash access and a renderer's use of
`count` instead of the shared `sample_count` field. `checks/summary.log.gz` and
`checks/summary-corrected.log.gz` retain those failures; `checks/summary-final.log.gz`
records the passing corrected summary. Neither correction changed captured
reports, journals, catalogs, protocol or the capture-time verifier.

The eight R1 report mutation suites rejected all 88 probes. This mutation
coverage spans both lanes, corpora and roles; it does not mutate the eight R2
reports. Portable replay independently revalidates all 16 reports and repeats
those eight R1 suites in an exported bundle with pinned tools. The portable
tool-binding check also requires a modified copied shared validator to fail
before replay. Hashes establish internal artifact consistency, not external
attestation. The final read-only review is retained in `checks/final-review.md`.

After removing the clean capture worktree and both copied binaries, portable
replay and the copied-tool tampering check passed again. All 35 raw `.log`
files were compressed losslessly with original and stored hashes in
`compression.json`; JSON receipts keep their original historical log paths,
which resolve through that compression map. Final replay uses the compressed
bundle and requires no original build artifacts. Shared targets and unrelated
worktrees were preserved.
