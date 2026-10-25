# 0362ul46 — XLS raw record limits and validation names

- Date: 2026-10-24.
- Source pin: `30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`.
- Scope: legacy XLS variable-payload record boundaries and the XLS validation
  vocabulary rename.
- Status: source and tracked-consumer preparation plus bounded Cargo
  validation through the final atomic candidate. The 0362ul46b/0362ul46c
  source-backed follow-ups below have passed their focused behavior, full XLS
  library, all-target check, and configured clippy rows after the exact tracked
  fixture corpus was materialized. No native, allocation, performance,
  full-workspace, or acceptance run is claimed here.
- Parent contract: 0351 §2.5.3 and the raw/record limit audit
  `.coord/raw-record-limit-audit-3431ba566c.md`.

## Canonical boundary

The two public XLS fragment codecs that retain variable payload data now take
one required caller-selected `&SourceLimits` argument. There is no default
wrapper or `_with_limits` twin.

| Capability | Canonical entry point | Limit checked before retention |
| --- | --- | --- |
| BIFF8 shared parsed formula (`ShrFmla`) | `litchi_xls::formula_metadata::shared::parse(data, &SourceLimits)` | `entry_bytes`, in addition to the BIFF8 8,224-byte format ceiling |
| BIFF8 worksheet hyperlink (`HLink`) | `litchi_xls::hyperlinks::parse_hlink_record(data, &SourceLimits)` | `entry_bytes` before display, location, frame, or moniker strings are decoded |
| BIFF8 worksheet hyperlink tooltip (`HLinkTooltip`) | collector-internal `parse_tooltip(data, &SourceLimits)` | `entry_bytes` before the UTF-16 tooltip is decoded |

The existing hard BIFF8 checks remain in place. A profile that is narrower
than the format ceiling is refused with the XLS `Error::Limit` and
`LimitKind::EntryBytes` before a token or string vector is copied. The parser
field order, malformed-record diagnostics, and successful decoded models are
otherwise unchanged.

## Selected-profile propagation

Every consumer in the source tree uses the policy already selected by its
owner:

- Eager workbook readers retain `Reader::source_limits()` and give that
  profile to `HyperlinkCollector` while parsing worksheet records.
- Source-backed `reader_ref()` materialization adapts the retained physical
  input ceiling into the eager `SourceLimits` profile before opening the raw
  reader; it no longer reopens under broad defaults. The logical worksheet
  scan budget remains on the source-backed owner and is not used as a physical
  CFB entry ceiling.
- `Package::sheet_features` adapts its eager `SourceLimits` directly. For a
  source-backed package it adapts the retained physical input ceiling through
  the checked crate-private `SourceLimits::from_physical_profile` helper;
  source-backed reads therefore do not reopen the codec under a broad default
  or reinterpret a logical worksheet budget as an entry limit.
- Cell-value structural certification adapts the snapshot's retained CFB and
  object limits before checking workbook dependency hyperlinks. The same
  profile reaches `parse_hlink_record` through the dependency walk.
- The collector stores its selected profile and has no `Default` or no-policy
  constructor. Both `HLink` and immediately following `HLinkTooltip` payloads
  use that selected profile before decoding. Tests and the BIFF writer use an
  explicit `SourceLimits` value, making the raw requirement visible at every
  direct call site.

The checked adapter preserves retained physical ceilings and rejects values
outside the finite XLS source contract instead of silently clamping them. It
does not change the owner's logical worksheet policy or add a second public
limit vocabulary.

## Adjacent XLS parser audit

The audit was limited to neighboring XLS record and codec surfaces, not every
semantic parser in the crate. The following classifications distinguish an
actual detached variable-payload boundary from a fixed or domain-owned helper:

| Surface | Classification | Decision |
| --- | --- | --- |
| `formula_metadata::array::parse_payload(data, Limits)` | variable formula tokens and ancillary bytes; crate-private module codec already receives `array::Limits` | already bounded; no duplicate `SourceLimits` bridge |
| `litchi_biff::Records::new(bytes, limits)` and `Encoder::new(limits)` | generic BIFF framing; explicit `litchi_biff::Limits` is already required | already bounded |
| `formula_metadata::codec::{parse_record, measure_record}` | crate-private Formula semantic path; the owning worksheet reader supplies the framed BIFF source and the format ceiling | domain-owned; no separate public raw entry point |
| `data_validation`, comments, external-link and worksheet collectors | crate-private semantic collectors reached through an owner; their record-specific constants are domain invariants | domain-owned; owner policy remains at the enclosing reader |
| `records::{BofRecord, DimensionsRecord, ExtendedFormat}` | fixed-width scalar/header fields with no retained variable collection | fixed/domain-only exclusion |
| `records::CellRecord` and `SharedStringTable` | worksheet/SST semantic models, not detached fragment APIs; current owners frame them through bounded BIFF records | domain-owned in this lane; no new public default or alias was introduced |

This audit did not silently treat a public variable fragment codec as fixed:
the shared-formula and hyperlink functions are the two confirmed detached
gaps and are the only public XLS APIs changed by this lane. Any later decision
to expose `CellRecord` or SST as an independent raw boundary must give that
boundary its own required limits in a separate owner-reviewed lane.

## Validation vocabulary migration

The public XLS validator now uses the focused module names:

- `litchi_xls::validation::Limits` replaces
  `litchi_xls::validation::XlsValidationLimits`.
- `litchi_xls::validation::Error` replaces
  `litchi_xls::validation::XlsValidationError`.

The validator implementation, XLS validation integration tests, unified
`litchi-ops` validation dispatch, and the performance-baseline consumer were
migrated together. No deprecated type, alias, or forwarding spelling remains.
The XLS glossary profile bans both retired names and points callers to the
focused replacements.

## Required-signature gate rows

The rows below use the required-signature shape coordinated with the PPT
owner and the DOCX `docx_encrypted_review` schema. They are source-review
rows, not a claim that the full gate has passed.

| Row | Required shape | Consumer proof | State |
| --- | --- | --- | --- |
| `xls.raw.formula_metadata.shared.parse` | `parse(data: &[u8], limits: &SourceLimits)` | shared-formula tests and BIFF writer call with explicit profile | ready for source/static review |
| `xls.raw.hyperlinks.parse_hlink_record` | `parse_hlink_record(data: &[u8], limits: &SourceLimits)` | hyperlink tests, collector, sheet features, and structural certification | ready for source/static review |
| `xls.raw.hyperlinks.parse_tooltip` | `parse_tooltip(data: &[u8], limits: &SourceLimits)` (collector-internal) | tooltip codec and collector tests reject over-limit payloads before UTF-16 decode | ready for source/static review |
| `xls.raw.hyperlinks.collector` | collector stores a selected `SourceLimits`; no no-arg constructor | eager reader and package feature projection pass retained/adapted profile to both hyperlink payload codecs | ready for source/static review |
| `xls.validation.names` | `validation::{Limits, Error}` only | XLS tests, `litchi-ops` dispatch, and perf baseline migrated | ready for source/static review |

## Source-only checks

The initial preparer review performed targeted source searches and
`git diff --check`. The old validation names and one-argument
`parse_hlink_record`/no-argument `HyperlinkCollector::new` call shapes are
absent from the migrated source files. A later bounded check also materialized
the tracked `clippy.toml` before running the lint rows; the distinction is
important because the first sparse checkout did not contain that file.

## 0362ul46b/0362ul46c source follow-up

The collector's private `HLinkTooltip` codec now requires the selected
`&SourceLimits`, checks `entry_bytes` before validating or decoding the UTF-16
payload, and receives the profile from `HyperlinkCollector::feed_record`.
Focused source tests cover both the codec boundary and the collector path with
an HLink payload that fits while the following tooltip exceeds the same
selected entry ceiling.

The source-backed eager materialization path now adapts only its retained
physical `max_input_bytes` into the eager `SourceLimits` profile before
constructing the compatibility `Reader`. The `SourceLimits` adapter is
checked, rejects out-of-range values rather than clamping them, and follows
the shared CFB hard input ceiling so a larger admitted source is not silently
narrowed. The logical `max_worksheet_scan_bytes` and record budget remain on
the positional source owner. A package-owner behavior test opens an XLS with
two real HLinks plus an unrelated 4 KiB CFB stream under a 64-byte worksheet
scan budget: the unrelated stream and HLink payloads remain readable while a
source-backed cell query returns the worksheet-scan resource limit. This
proves policy behavior rather than only checking stored profile fields.
The public-signature policy uses the c51a schema and registers the exact
`data, limits` inputs plus nominal `litchi_xls::SourceLimits` for the shared
formula and HLink functions. The c51a `checked_policy` pass is source/static
only; the prior bounded receipts below do not cover these post-ac5 changes.

## 0362ul46c bounded revalidation

The first fresh validation attempt was bound to `378a311c74654114172eea358b6f1ca0b7e377b7`.
All nine planned rows reached the runner, and every owned process exited, but
the source failed the configured `unused-qualifications` lint at
`cell_values::Snapshot::raw_record_limits` (`crate::Result` instead of the
imported `Result`). Its state and logs remain outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46c-validation-378a311c`;
the first diagnostic log SHA-256 is
`292b3631488d57b6c3bc41ca33fb6ac0d95a47b9639690c564305949fa8a4655`.

The one-line lint correction is `c6bad8ffeeb00eec85949197d3508d0c1c6373e5`.
The fresh rerun state is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46c-validation-c6bad8`;
it used the pinned 1.95.0 toolchain, `Cargo.lock` SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`,
`.cargo/config.toml` SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`,
`rust-toolchain.toml` SHA-256
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`,
and `clippy.toml` SHA-256
`f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c`.
The approved f5 runner/tool hashes are the ones recorded above. CPU affinity
was 30; the run used offline/locked Cargo, one job, one test thread,
incremental/debug info disabled, the 20G shim, and the decimal
18,000,000,000/19,000,000,000/20,000,000,000 target guards. The peak target
was 606,177,051 bytes.

The source-limit, shared-formula, hyperlink, and new package-owner behavior
rows passed, as did the XLS and ops-XLS all-target checks and both configured
clippy rows. Their individual receipts and empty/pass logs are recorded in
the fresh state. The complete `litchi-xls --lib` row is retained as a
diagnostic failure: 1,287 tests passed and 47 fixture/corpus tests failed
because the sparse checkout does not contain the broader XLS fixture corpus;
the row's log SHA-256 is
`0d7535600d0b3f0beb3e124b54a5ae63a748e88ed75c7caf742566b6f5dd675b`.
This is an environment-materialization gap, not an API or compile failure.

The exact tracked fixture directories named by those failures were then
materialized without hydrating `3rdparty/` or changing tracked source. A fresh
rerun at `857f4efe0d163dfa7e097096100fcf255b1a5406` used state
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46c-validation-857f4f-fixtures`.
It reached 1,333 passing tests and one SST corpus-digest failure because six
tracked `.xls`/`.xlt` files were still absent; that diagnostic log is
`de9a90ecd95da8575c9eef825203220904e57052ffe89ebb76bd37aff0e57f86`.
The six paths were materialized from the frozen tracked source, and the final
corpus-complete rerun used state
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46c-validation-857f4f-corpus-complete`.
All 1,334 `litchi-xls --all-features --lib` tests passed with exit 0; its log
SHA-256 is
`6233bd990bc9c3adf89e1d408f96ae594dea3423ec9e22ebd6668a3d3d7d3aaa`, and
the target ended at 320,659,975 bytes. The final input set contains all 126
tracked `.xls`/`.xlt` paths required by the corpus, with per-file hashes in
the external plan receipt. The fixture-only sparse materialization left the
worktree source tree clean. The final run retained the pinned 1.95.0,
offline/locked, single-job/single-test-thread, incremental/debug-disabled,
CPU-30, no-swap, and decimal 18/19/20 GB target guards. The final run is a
bounded correctness/library result; no native, allocation, performance,
full-workspace, or acceptance run is claimed.

## Bounded validation receipt

After the initial source-only review found one missed internal consumer, the
following follow-up commit was tested without changing the original API or
glossary work:

- Follow-up source commit: `62fd1507493187ddb77a2400990be853e5fd2a26`
  (`sheet_features::Edit::features` now forwards the package's retained raw
  record limits).
- The first `ad8713ede7` attempt is retained outside the checkout at
  `/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46-validation-ad8713e`.
  Its blocked workspace-parse log SHA-256 is
  `a67006dda034fb5da9bd8729f7818be8286385e0300eb2cbe66a748d5f0787ce`; its
  compile log SHA-256 is
  `951c7b6378de3b5422206e3900c81bfa2cba4b51c953bfcbaa10a60416ad71e7`.
- The completed follow-up receipts are outside the checkout at
  `/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46-validation-62fd1507`.
  The fresh target remained below 600 MiB. The f5 runner, helper, adjacent
  Cargo wrapper and underlying Cargo hashes were respectively
  `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`,
  `3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f`,
  `495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`, and
  `1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`.
- Passed under CPU 30, offline/locked, `JOBS=1`, test threads 1,
  incremental 0, debug info 0, swap 0 and the 18/19/20 GiB target guards:
  the shared-formula unit filter, hyperlink unit filter, ancillary formula
  integration target, validation integration target, raw unit filter,
  all-target XLS `cargo check`, and (after sparse-materializing the required
  `litchi-sheet` sources) all-target `litchi-ops --features xls cargo check`.
- The first clippy rows in this receipt are retained as diagnostic failures,
  but are not valid gate results: the sparse checkout was missing the tracked
  `clippy.toml`, whose hash is
  `f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c`.
  Without that file, the tool's default seven-argument threshold reported the
  unchanged CFB `emit_stream` and `publish_padded` functions as `8/7`; the
  same no-config run reported existing XLS `too_many_arguments` and
  `type_complexity` sites and existing ops `too_many_arguments` and
  `type_complexity` sites. The retained logs are
  `07-xls-clippy-all-targets-attempt-62fd1507-7.log`,
  `09-ops-xls-clippy-all-targets-attempt-62fd1507-9.log`,
  `11-xls-clippy-all-targets-no-deps-attempt-62fd1507-followup-11.log`, and
  `12-ops-xls-clippy-all-targets-no-deps-attempt-62fd1507-followup-12.log`
  under the validation directory above. The CFB file is byte-identical to the
  original 30aa source (`f9159593ab37b8b4a5804a201f3614190c62094e1ac28d865761bcdf29203aa9`),
  and the original 30aa checkout passed the same configured rows below. No
  lint was waived, suppressed, or relabeled as a pass.
- The required hyperlink fixtures were sparse-materialized from the frozen
  source with hashes `WithTwoHyperLinks.xls=
  139a3ace5945fd0a1e498dd6482365b4f7f2281c879b8dcd15e10b5fb0963aa9`,
  `HyperlinksOnManySheets.xls=
  e03ddd49d757af31b897ffcc394d733131dc28ab484f9cca484351810d0935af`, and
  `54016.xls=2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a`.

This receipt records bounded correctness/build diagnostics only; it is not an
acceptance, native, allocation, performance, or full-workspace result.

### Clippy gate correction and base comparison

The missing configuration was corrected before interpreting the lint gate. The
tracked `clippy.toml` is byte-identical in the current source, the 30aa source
checkout, and the frozen 9781 gate source. Its threshold is 12, so the newly
required raw-limit parameters do not create an unreviewed seven-argument
exception. No crate-level lint allowance, lint rename, or API reduction was
made.

The corrected current-source run was launched from the record-only tip
`b877fe72cde2b74b316affc728d5e86497e7e731`; its Rust/Cargo tree is unchanged
from the tested consumer-fix tip `62fd1507493187ddb77a2400990be853e5fd2a26`.
It used this exact pair of commands and passed both with exit 0 and empty logs:

```text
cargo clippy -q -p litchi-ops --no-default-features --features xls --all-targets -- -D warnings
cargo clippy -q -p litchi-xls --all-features --all-targets -- -D warnings
```

The receipts are in
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46-clippy-62fd1507`:
the plan SHA-256 is recorded there, both log SHA-256 values are
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`, and the
post-row target sizes are 159,742,896 and 187,853,315 bytes. The run used the
approved f5 runner/tool hashes already listed above, CPU 30, offline/locked,
one Cargo job, one test thread, incremental 0, debug info 0, swap 0, and the
18/19/20 GiB target guards.

For a direct base comparison, the same two commands, tool hashes,
`Cargo.lock` SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`, and
`clippy.toml` SHA-256 were run from the clean 30aa checkout. Both passed with
empty logs; their external receipts are under
`/home/zhuhe/code/litchi-worktrees/.coord/raw-limits-xls-0362ul46-base-clippy-30aa`
and their post-row target sizes are 159,748,823 and 187,861,703 bytes. The
frozen 9781 gate also records passed clippy rows with the same empty-log hash,
including `clippy-all-features` and the XLS-enabled ops rows, in
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781cf6d87/state.json`.
Thus every prior diagnostic is attributable to the absent preparation file,
not to the new required arguments: configured current, original-base, and
frozen gate rows all pass, while the unchanged CFB source remains unchanged.

## Coordinator composition review

Final atomic `2ffacb16e929cdad7b032f5e46f25ed72fbc4ef6` differs from
tested `c6bad8ffeeb00eec85949197d3508d0c1c6373e5` and the complete-corpus
source `857f4efe0d163dfa7e097096100fcf255b1a5406` only in this record.
The delta from independently reviewed `378a311c74` is a redundant `Result`
qualification cleanup, with no change to error type or behavior. Root verifies
all scoped command logs and Cargo traces and the final full-library receipt.
All 1,334 library tests pass; each terminal owned-process list is empty, the
approved lock is unchanged and the source stays clean.

The complete-corpus `state.json` is 28,851 bytes, SHA-256
`e9bb8b4769c8f42fc20ea4de9a8f58b6de715e449c72ebaaf79d2dc8d128a647`.
Its individual receipt is 27,553 bytes, SHA-256
`be68de8ccfe16d5f8d17f2d4861d952fb9623d42f716be4a35cd40b3d10b95b9`.
The previous 47-fixture and six-file corpus-digest failures remain retained.

The current performance harness uses the canonical validation type. The
DrawingML historical adapter migrates in the same composition, preserving
the exact historical full-file hash through three explicit RTF/XLS spelling
changes (record 0362ul57a). The evidence matrix classifies this directory as
source/validation evidence. The standalone performance-harness compiler
check on the composed source remains queued, alongside the final integrated
gate and actual native qualification; this record claims only the scoped
results above.
