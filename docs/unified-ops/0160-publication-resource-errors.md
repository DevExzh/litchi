# Preserve resource errors from XML publication

Record 0159 found a PPTX edit that staged successfully but failed serialization
at the OPC XML token budget. Its native error retained the typed resource,
limit, observation and offset, while the unified facade returned `corrupt_file`.
The shared OPC classifier now returns `limit_exceeded` for XML publication
budget failures and `internal_error` for audit-buffer allocation failures.
Malformed or otherwise invalid XML remains `corrupt_file`.

The XML audit error exposes two borrowed predicates, `is_limit` and
`is_allocation`. These let wrappers classify the already-public typed error
without depending directly on the XML parser crate or converting diagnostics
to strings. The facade retains the original leaf error and standard source
chain. No publication policy, parser budget, supported capability, or success
path changed.

The end-to-end regression stages text of 4 MiB plus one byte, within the native
8 MiB edit budget, and then triggers the publication token limit. Both eager
and source-backed unified owners must return `limit_exceeded`, preserve the
native PPTX/OPC/XML error and leave the source serializable and unchanged.
XML audit tests distinguish real resource and malformed-input failures, plus
allocation classification. Other formats using the shared OPC classifier
inherit the correction, though this record's save regression is PPTX-specific.

## Verification

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,111
tests/doctests, and the XML audit crate passed 162 tests. Both crates passed
all-target Clippy with warnings denied. Clippy also exposed an existing
single-character `push_str` in the XML benchmark example; it now uses `push`
with identical output. Scoped formatting and diff checks passed. The generated
OfficeCLI capability delta matches the base checkout's help schemas.

Warning-denied rustdoc passed for both crates. Both new XML predicate
doctests passed. The facade also builds with no default features. Repository boundaries passed for 66 workspace packages and 268 internal
dependency declarations. No all-format Python or
fuzz campaign is claimed by these focused checks.

## Reproduction

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0160-target CARGO_NET_OFFLINE=true cargo test --offline -p xml-minifier -p litchi-ops --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/pptx,litchi-ops/encryption
CARGO_PROFILE_DEV_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0160-target CARGO_NET_OFFLINE=true cargo clippy --offline -p xml-minifier -p litchi-ops --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
```

## Scope and repository

The previous goal turn made verified progress on selector fanout. This turn
closes its observed error-taxonomy follow-up. No benchmark is claimed for an
error-only classification change; record 0159 retains the success-path timing
evidence. Broader format/CRUD coverage, Python/CLI work, remaining glossary
migration and full performance acceptance remain open.

The rewritten base remains an ancestor, so no rebase is required. Existing
pending work from records 0158–0159 is preserved. The full goal remains active.

Both required staged-file guards printed nothing. Scoped staging failed because
Git could not create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
on the read-only filesystem. This agent could not create the requested commit.

All verification jobs ended before cleanup. Turn-owned build output and logs
were removed. The inherited root lockfile and prior pending changes were
preserved. No fixture copies or temporary worktrees were created, and the base
checkout was not modified. The final artifact audit found no prohibited large
evidence or compiled binaries among the 16 pending paths.
