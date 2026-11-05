# 0362ugc — ODS source-backed construction-cost closure

- Date: 2026-10-23.
- Status: source implementation, review follow-up, and ODS validation complete
  on `uo2/ods-lazy-entry`; coordinator integration and the final workspace gate
  remain pending.
- Performance claim: construction-cost closure only. No native timing,
  baseline, ABBA, RSS, or allocation claim is made by this record.

## Source and review

The implementation is in commits `ce3c80bc74da7303dc103efe00ec30a172c457ff`
(`ods: defer source-backed edit construction`),
`2e2f902179d8ea107d05d3805542e89d9a42f27a` (`ods: validate lazy edit
admission`), and
`8507ec866e3434231f8b8307051999f9a5bbcdfb` (`ods: defer resource admission
to edit operations`), based on the clean ODS worktree at
`7559f44aa975254bd96f6e28eee17366b0ebeb6a`. It changes no public spelling,
alias, or facade contract.

Pinned Clippy found a redundant bare `#[must_use]` after `resource_count`
changed to return `Result<usize>`; `1fdfcdcdba11b2860c08037aa9939450770ecc84`
removes that attribute. The receipts below are for that corrected source
commit.

The source-backed DDE and sheet-metadata edit entries now retain the already
validated `Arc<str>` content, source version, limits, and execution policy.
`OnceLock` initializes the bounded parser only on the first read, staged
mutation, or commit. The ordinary read snapshots remain eager; only the edit
entry path is deferred. Admission still performs the O(1) limit validation,
cancellation/context check, source-version check, and source capture. The
edit keeps source freshness checks around parsing, staging, and publication;
signed changed commits, stale sources, parser failures, and output/operation
limits retain their prior behavior.

Source-cell protection is no longer parsed while constructing an edit. The
owner's protection result is cached on the first effective mutation, so a
dropped edit and a pristine no-op entry do not pay the protection scan. The
changed and signed publication paths still use the same owner checks and
deferred source-backed write path.

The package archive retains its validated index at open. Resource admission is
deferred to the first operation on an `Edit` (or to commit/replay validation)
and is then cached, so opening, shared handoff, and edit construction do not
allocate a file-name census. Staging validates both the source and candidate
before retaining changed edit state; an untouched edit validates at commit
before publishing even a no-op snapshot. Flat ODS edits stage only touched
sheets in a sparse `BTreeMap`; the full aligned staging vector is materialized
only for a changed commit. No eager output buffer is created by any of these
entries. The resource-admission boundary correction is recorded in
`0362ugca`.

The focused regressions cover both DDE and sheet metadata construction/drop,
invalid finite limits and canceled contexts before parsing, exact no-op
publication, source-cell protection sharing after the first mutation, and
the unchanged owner after dropping an edit. The named construction proofs are
`edit_entry_defers_scan_and_drop_keeps_the_owner_unchanged`,
`edit_entry_admits_limits_and_cancellation_before_parsing`, and
`no_op_commit_preserves_deferred_source_publication` in each source-backed
module, plus
`source_edit_protection_is_lazy_and_shared_across_sequential_edits` in the
facade tests. The resource proof
`resource_admission_stays_cold_until_edit_operation_or_commit` verifies an
empty cache after open and edit drop, first-operation population, no staged
state after an over-limit refusal, and first-commit admission.

## Fixture and sparse-checkout audit

The ODS test inputs excluded by the initial sparse checkout were tracked in
`HEAD`; none was copied from `3rdparty/` and none was added to Git. Thirty-four
compile-time ODS fixture paths total 4,626,860 bytes. The three runtime paths
needed by the full suite are:

```text
docs/report/spec-gap-validation-evidence/ods-formula-complex-functions/contract.md  7,969
docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle-vectors.json  45,915
docs/report/spec-gap-validation-evidence/ods-formula-lookups/lookup-goldens.json  56,281
```

The required 37 paths total 4,737,025 bytes and are all tracked by `HEAD`.
The sparse checkout was expanded only to materialize those tracked inputs and
the repository's root build configuration; no fixture or evidence artifact was
modified.

## Toolchain and configuration identity

The final validation used the tracked root configuration that had been absent
from the initial sparse checkout:

| file | `HEAD` blob | SHA-256 |
| --- | --- | --- |
| `clippy.toml` | `4c8560024c61794aaee1409ec56c9e007c45f1ef` | `f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c` |
| `rust-toolchain.toml` | `8934eb4a5fabfca2502492140bfb51e89f9e0394` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |
| `.cargo/config.toml` | `c0e59805cf480523635677e333124885d257dcdd` | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` |

The pinned tools report:

```text
rustc 1.95.0 (59807616e 2026-04-14)
cargo 1.95.0 (f2d3ce0bd 2026-03-21)
clippy 0.1.95 (59807616e1 2026-04-14)
```

All Cargo commands used the memory-capped environment in
`(deleted-coord-artifact)`, offline/locked resolution, one build job, and
the external target
`(deleted-coord-artifact)`. Its final
size was 3,024,349,798 bytes, below the 19.2-billion-byte preflight and
20-billion-byte hard ceiling.

## Validation receipts

The full ODS suite was run as:

```text
source (deleted-coord-artifact)
cargo test --offline --locked -j1 -p litchi-ods --all-targets
```

It exited `0` with 119 `test result: ok` blocks and 1,768 passed tests. The
raw log remains external:

| command | log bytes | SHA-256 |
| --- | ---: | --- |
| `(deleted-coord-artifact)/ods-full-tests-1fdfcdcdba.log` | 181,605 | `9067fb06658fbc3b37dabb7768b40d43fc520bfd68e4bb96f7416e229bb8a14a` |

Strict all-target/all-feature Clippy was run as:

```text
source (deleted-coord-artifact)
cargo clippy --offline --locked -j1 -p litchi-ods \
  --all-targets --all-features -- -D warnings
```

It exited `0` in 59.61 seconds with no diagnostics. The raw log is
`(deleted-coord-artifact)/ods-all-targets-clippy-1fdfcdcdba.log`,
165 bytes, SHA-256
`8a3273376dbbad60659de5d638bb73466f63177a29f171894a43a3f479e758c5`.

`git diff --check` passed for the completed source/record changes. The
repository-wide `cargo fmt --all -- --check` remains a coordinator concern: it
reports pre-existing formatting drift in unrelated workspace files and no
formatter rewrite was applied in this lane. No binaries or large evidence
files were committed.
