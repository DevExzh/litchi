# iWork Progress Audit

> Audit baseline: 2026-09-01 at committed `d33f30f41` (`feat(numbers): own currency cell formats`). This update evaluates the focused Scientific-, Fraction-, Text-, and Date & Time-format owner work from that baseline; it remains scoped verification rather than release certification.

## Conclusion

The iWork implementation is substantial but is **not suite-complete or release-certified**.

- Pages, Keynote, and Numbers have credible bounded semantic readers.
- The focused owner crates have strong exact-source, fail-closed transactions for selected existing objects.
- Broad fresh creation, structural editing, rich formatting, charts, shapes, and media remain largely in the legacy `litchi-iwa` migration host.
- Three authoritative Pages, Keynote, and Numbers matrices now report focused-owner, legacy-host, preservation, and evidence boundaries; the existing implementation review still explicitly excludes iWork.
- Native fixtures establish basic parse/no-op fidelity, but there is no suite-wide automated proof that Litchi-modified files are accepted, saved, closed, and reopened by all three native applications.
- The focused Percentage owner and codec checks are green at the synthetic/source level. The focused Currency codec filter passes 4/4, the `litchi-numbers-wire` library passes 32/32, and the package target passes 22/22 after native BNC flag handling became shape-dependent. ADR 0008 records operation-specific native Currency E3/E4 evidence with exact candidate/native-resaved hashes, inverse restoration, and strict no-op reread. Percentage remains below formal E3/E4 promotion; Currency native evidence remains limited to its stated existing-cell operation.
- The focused Scientific owner now has an archive-free value, selector-first exact-source transaction, strict Buffa type-259 route, bounded sanitizer smokes, and operation-specific native E3/E4 evidence. The package target passed 19/19, strict codec filter 4/4, wire library 34/34, and legacy bridge filter 6/6. Numbers 14.4 opened, saved, closed, and reopened the precision-7 candidate without repair; strict reread was an exact semantic no-op. This evidence remains limited to the recorded existing-cell operation.
- The focused Fraction owner now has an archive-free value, selector-first exact-source transactions, strict native type-262 preflight followed by a lazy Buffa view, and a boundary-enforced retirement of the dedicated raw-ID `NumbersEditor` convenience API. Generic source-built `DataFormat::Fraction` mutation and Pages/Keynote compatibility helpers remain in the migration host. All nine `FractionAccuracy` strategies are covered. The frozen evidence records 22/22 package integration tests, 4/4 library codec tests, 3/3 direct codec tests, and 36/36 wire tests, plus 44 proto fuzz corpus seeds and 18 package fuzz corpus seeds; both fuzz targets completed 100-run AddressSanitizer smokes. Field 20 (`requires_fraction_replacement`) is preserved when absent or canonical `false`; `true` is rejected. Computer Use established operation-specific native E3/E4 evidence for a representative Eighths-to-Hundredths edit; exact source/candidate/native-resaved hashes are recorded in ADR 0008. This does not promote all nine native UI variants, arbitrary-producer parity, native byte parity, or package-wide performance.
- The focused Text owner now provides an archive-free marker and selector-first exact-source transactions for one existing cell. Strict source preflight precedes a private lazy Buffa view; explicit set/clear/reset, copy-on-write/refcount closure, candidate reopen/readback, scalar and unrelated-byte preservation, locality, inverse, and typed family refusal are covered. New explicit attachments use canonical marker `0x80`; unchanged admitted converted-Text sources with marker `0x81` and retained Number provenance remain byte-exact. The checked-in native fixture supplies E2 read/no-op evidence, while focused tests supply E1 evidence; no native E3/E4 acceptance claim is made and numeric-to-Text conversion is unsupported.
- The focused Date & Time owner now provides a selector-first, archive-free transaction for one existing cell's bounded native date/time pattern string. Type 261 requires explicit marker `0x0008`, kind `3`, and strict fields 1/14; Empty/type-5 Date and the evidenced type-9 numeric/formula shape are admitted, while plain type-2, marker-zero, wrong/reserved, and ambiguous shapes reject. Strict preflight precedes a lazy Buffa view, the native date/time pattern string is bounded to 4096 bytes, and exact COW/refcounts, unknown/unselected bytes, scalar value, inverse, and locality are retained. The owner checks the bounded string envelope and admitted fields only; pattern grammar is not validated. A real Numbers 14.4 (build 7043.0.93, macOS 26.5.2) candidate opened without repair/conversion, retained A1 marker text and B2 `2026/09/04 12:34:56` through native save/close/exact-path reopen, and preserved the DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` byte-for-byte across native normalization. Strict normalized reread then reported `changed=false`, `touched_components=0`, `full_reparse=false`, and `scalar_value=untouched`; its 138,725-byte output was byte-identical to the native-resaved artifact (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This is operation-specific E3/E4 evidence for this one existing type-9 cell/pattern only; no broad DateTime/package/native-byte-parity claim follows, and dedicated host raw-ID DateTime retirement is not claimed.
- Numbers persisted-sort and table-relocation compatibility remain narrower than the focused semantic owners. Exact package snapshots use the focused semantic transaction and any structural, family, lock, budget, stale-source, or locality refusal is terminal. Only historical source-built snapshots may use the respective crate-private physical bridge (field 44 for persisted sort or physical table relocation); these bridges are migration compatibility seams, not public owners or second implementations, and have no ADR 0028 debt/deletion-gate impact.
- The focused Keynote physical `Sort Now` owner now has a selector-first, source-bound transaction with strict wire admission, lazy Buffa inspection, bounded row-envelope/UID movement, exact-source inverse/locality checks, and atomic refusal of unproven row-affine state. The 26-case package suite and bounded fuzz targets are source-level evidence only. A disposable Computer Use run opened a candidate produced before the latest strict model-field admission, but the current owner rejects the app-authored source because model field 39 identifies an unowned conditional-style CalculationEngine dependency graph; that run is exploratory evidence, not current-owner E3/E4 promotion. Native acceptance remains pending.
- The focused Keynote `soundtrack::items` semantic/package owner now handles bounded read/add/insert/replace/remove operations while keeping native IDs, media topology, package bytes, and wire values private. Its focused lifecycle, shared-occurrence, malformed-reference, aggregate-only, exact apply/inverse, conflict, and fuzz gates pass, so the duplicate legacy item API, eager wire helper, example, and legacy-only tests are retired. One genuine replacement candidate passed Keynote save/close/reopen without repair; this is representative replacement evidence, not operation-wide native certification, soundtrack creation, broad media CRUD, or a monolith-exit claim.

The detailed provisional coverage assessment is in [IWORK_FEATURE_MATRIX_AUDIT.md](IWORK_FEATURE_MATRIX_AUDIT.md).

## Current maturity

| Layer | What is solid | What still lacks | Assessment |
|---|---|---|---|
| Root `litchi::iwork` facade | Immutable, bounded Pages/Keynote/Numbers projection; host-free dependency path | Read-only and text-oriented; no package/edit, metadata, media, chart, or formula editing/recalculation surface | 🟡 focused read facade |
| `litchi-pages` | ZIP/directory semantic read; exact retained-package output; selected text, layout, footnote, header/footer, and table-property transactions | Fresh package creation, section/table structural CRUD, table cells/formulas, rich formatting, drawings, charts, media, collaboration, export | 🟡 bounded reader/editor |
| `litchi-keynote` | Show/slide read; existing-slide state, text, notes, settings, background, transition, selected chart/movie/table properties, a source-level physical table-sort owner, and bounded soundtrack-item read/add/insert/replace/remove | Slide creation/duplication, public table row/value reads, native-certified physical sort or operation-wide soundtrack-item lifecycle, soundtrack creation, full table data, chart data/series/CRUD, general media bytes/CRUD, full build/animation editing, shapes/groups, masters/themes, collaboration, rendering/export | 🟡 bounded reader/editor |
| `litchi-numbers` | Rooted workbook read; selected scalar cells, formulas, controls, table settings, names/order, comments/replies, and existing-cell Number/Percentage/Currency/Scientific/Fraction/Text/Date & Time display formats | Sheet/table lifecycle, row/column topology, full formula engine, other generic formats/rich styles, package merge editing, charts/media/drawables, filters/pivots/categories, print/page setup and workbook export | 🟡 bounded reader/editor |
| `litchi-iwa` | Broad source-free authoring and native mutation across all three applications; extensive tests/examples | It is explicitly a compatibility/migration host, exposes native/raw seams, and still owns most rich/structural authoring | 🟡 broad but non-canonical |
| Shared IWA crates | Bounded Snappy/wire/archive parsing, package preservation, detection, focused codecs, exact artifacts, COW state, and archive-owned durable publication | Concrete-owner index adoption, aggregate graph/memory budgets, directory write parity, durable patch serialization/history, encryption/signatures | 🟡 mature substrate with open boundaries |

This assessment deliberately does not infer semantic support from generated protobuf availability or raw public declaration counts.

## Architecture and migration

ADR 0028 defines `litchi-iwa` as the sole legacy migration host and makes deletion conditional on closing every recorded edge and passing native reopen gates ([ADR 0028](adr/0028-iwa-monolith-exit.md#deletion-gate)). The authoritative current topology is **64 workspace packages, 238 internal dependency declarations, 227 canonical edges, 11 development-only edges, and 11 ordered migration debts**, with one migration host. The debt orders are `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]` ([boundary ledger](../tools/crate_boundaries.json)); these are ordered host-exit conditions, not a count of all IWA-related crates.

Positive progress:

- The root facade depends on the three concrete owners rather than `litchi-iwa` ([feature wiring](../crates/litchi/Cargo.toml#L60), [facade](../crates/litchi/src/iwork/mod.rs#L1)).
- Recent waves moved narrow Pages table/text owners, Numbers cells/settings/comments/display-format owners, and Keynote chart/movie/table owners into focused crates.
- `litchi-iwa-index` is a neutral leaf consistent with ADR 0029.
- No `TODO`, `FIXME`, `todo!`, or `unimplemented!` stubs were found across the iWork crates; incomplete work is represented by typed refusal paths and migration debt.

Open migration facts:

- The concrete app crates do not yet consume the neutral index directly; debt 008 remains.
- Generated protobuf modules remain a public compatibility surface, while production semantic coverage uses a much narrower set of focused codecs.
- The host still joins focused values to legacy objects through native IDs and still coordinates package reassembly and broad mutation.
- Focused owners couple exact-source authorization/reassembly to process-local logical COW patches. The logical patches are archive-neutral and do not contain full package bytes; the archive now provides durable owner `Package::save` publication, but neither layer provides durable, serializable, composable, mergeable patch history.

## Evidence and release posture

| Area | Evidence present | Gap before a strong completion claim |
|---|---|---|
| Native fixtures | One checked-in native fixture per app plus package-directory fixtures; native provenance and hashes are documented in [`test-data/iwork`](../test-data/iwork/README.md) | Fixtures are intentionally basic. Rich tables, formulas, charts, media, comments, builds, themes, and malformed/edge producer variants are not represented broadly |
| Package fidelity | Native parse, semantic parity, exact no-op output, inverse/locality tests, and substantial synthetic transaction suites | Self-roundtrip does not prove native acceptance of changed output; native verification for generated packages is explicitly external/opt-in |
| CI | Ubuntu workspace all-feature lib/integration/doc tests; all targets compile | iWork is absent from macOS/Windows test matrices; no leaf-feature isolation, native-app runner, fuzz/sanitizer, Miri, or coverage gate; iWork fixture paths do not trigger Rust CI |
| Ignored evidence | Four Numbers native-oracle tests exercise formula-rich cases | They depend on a private `/private/tmp` oracle and normal CI cannot reproduce them |
| Publish | Package manifests and protobuf provenance guards exist | Publish workflow does not install `protoc`; no clean unpacked-`.crate` test proves repository-external fixture references work |
| Documentation | Three authoritative app matrices, root-index links, extensive ADR wave log, and crate rustdocs | Root README/rustdocs understate focused writes while parts of `litchi-iwa` docs overstate canonical support; the implementation review still excludes iWork |

The prior [format implementation review](FORMAT_IMPLEMENTATION_REVIEW.md#scope-and-rubric) expressly excludes Pages, Keynote, Numbers, and shared IWA, so its pass result cannot certify this suite.

## Current-worktree verification

These results describe the focused pre-commit slice snapshot, not a release artifact. The
Scientific rows were rerun for this slice, and the Fraction and Date & Time rows use the frozen
source/build/test/fuzz/native evidence for this update. The Percentage and Currency rows retain
their previously recorded results. The Keynote physical-sort checks below are source-level and
exploratory-native diagnostics; they do not certify current-owner native interoperability:

| Check | Result |
|---|---|
| `python3 tools/check_crate_boundaries.py` | ❌ Three errors: untracked `crates/litchi-iwa/src/pages/editor/tables/lock.rs` restores the retired `body_table_lock_state` and `set_body_table_lock_state` host surface |
| `cargo test -p litchi-numbers --test table_cell_percentage_format --no-fail-fast` | ✅ 19/19 focused Percentage owner tests passed, including selector, family-boundary, malformed-input, locality, inverse, concurrency, and budget cases |
| `cargo test -p litchi-iwa-protos percentage_format --no-fail-fast` | ✅ 4/4 focused Percentage codec tests passed, covering native domains, canonical framing, unknown preservation, and Number/Percentage separation |
| `cargo test -p litchi-numbers --test table_cell_currency_format --no-fail-fast --locked --offline` | ✅ 22/22 focused Currency owner tests passed after native BNC flag handling became shape-dependent, including refcount and secondary-reference cases |
| `cargo test -p litchi-iwa-protos currency_format --lib --no-fail-fast` | ✅ 4/4 focused Currency codec tests passed, covering native domains, canonical framing, unknown preservation, and lazy-view separation |
| `cargo test -p litchi-numbers-wire --lib` | ✅ 32/32 focused Numbers wire library tests passed |
| Focused Currency fuzz targets | 🟡 Source-visible with bounded synthetic/adversarial corpus coverage; sanitizer smoke evidence remains separately scoped and is not used to broaden the native claim |
| Numbers Currency native validation | ✅ Operation-specific E3/E4 evidence: Numbers 14.4 opened the Rust candidate without repair/recovery/conversion, preserved B2/B3 text/scalar semantics and requested Currency settings through save/close/reopen, and strict Rust no-op reread plus exact inverse restoration matched the recorded hashes in ADR 0008 |
| Scientific focused owner/build/test/fuzz/native validation | ✅ Package 19/19, strict Buffa codec 4/4, wire 34/34, and filtered legacy bridge 6/6 passed. Targeted Scientific boundary audits reported zero violations. Sanitizer-backed codec/package smokes completed 100 executions from 33/10 corpus files. Numbers 14.4 opened the precision-7 Rust candidate without repair, retained text/scalar/format through save-close-reopen, and strict reread produced an exact no-op; ADR 0008 freezes source, candidate, native-resaved, inverse, and locality hashes. |
| Fraction package integration tests | ✅ 22/22 focused existing-cell Fraction owner tests passed, covering selector-first exact-source transactions, all nine accuracies, set/clear/reset, inverse/no-op, COW/refcounts, preservation/locality, budgets, conflicts, and fail-closed refusal paths |
| Fraction library codec tests | ✅ 4/4 strict native type-262 codec tests passed, covering canonical framing, all nine accuracies, lazy Buffa parity, unknown preservation, and strict field-20 handling |
| Fraction direct codec tests | ✅ 3/3 direct codec tests passed, including absent/canonical-`false` field-20 preservation and canonical-`true` rejection |
| Fraction wire tests | ✅ 36/36 focused Numbers wire tests passed |
| Fraction fuzz targets | ✅ Checked-in corpora contain 44 proto seeds and 18 package seeds; the protocol and package targets each completed a 100-run AddressSanitizer smoke from isolated corpus/build/artifact roots with no crash or invariant failure. These bounded runs are not exhaustive input coverage. |
| Numbers Fraction native validation | ✅ Operation-specific E3/E4 evidence: Computer Use created B3 Fraction/Eighths from a disposable copy, Numbers opened the Rust Eighths→Hundredths candidate without repair/recovery/conversion, preserved B2 text and B3 Actual `42`, then saved/closed/reopened a duplicate with the same Fraction/Hundredths readback and no prompt. Before native opening, exactly two uncompressed members differed while names/order and every other member payload matched. Rust no-op/inverse and native-resaved name-selector rereads were exact; source/candidate/native-resaved hashes and byte sizes are recorded in ADR 0008. This represents Eighths/Hundredths only, not all nine native UI variants or arbitrary-producer/native-byte parity after Numbers normalization. |
| Numbers Date & Time native validation | ✅ Operation-specific E3/E4 evidence for one existing type-9 cell/pattern: Numbers 14.4 (build 7043.0.93, macOS 26.5.2) opened the Litchi candidate without repair/recovery/conversion, retained A1 `Litchi DateTime producer — 東京😀`, B2 `2026/09/04 12:34:56`, Actual `9/4/2026 12:34:56 PM`, and inspector Date & Time/date `2026/01/05`/time `19:08:09` through save/close/exact-path reopen. Native normalization changed document, calculation, stylesheet, metadata, view-state, and preview entries, but the exact DateTime members in `Index/Tables/Tile.iwa` and `Index/Tables/DataList-904498-2.iwa` stayed byte-identical to the candidate. Strict normalized reread reported `changed=false`, `touched_components=0`, `full_reparse=false`, and `scalar_value=untouched`; its 138,725-byte output was `cmp=0` with the native-resaved artifact (SHA-256 `4e489da196bac7416c8c6d827afb2a6264892d4856a5a33dc0a18c6c2d2b1b0c`). This remains scoped to that one existing type-9 cell/pattern; no broad DateTime/package/native-byte-parity claim or dedicated host raw-ID DateTime retirement follows. |
| Keynote physical `Sort Now` package tests | ✅ 26/26 focused source-level tests passed, covering selector/range semantics, stable scalar ordering, row/header/UID movement, exact no-op/inverse/conflict, unknown/member preservation, aggregate-only metadata compatibility, optional format-route validation, noncanonical IWA framing refusal, strict dependency refusal, and budget paths. This does not provide a public row/value reader or native-app evidence. |
| Keynote physical `Sort Now` native probe | 🟡 A disposable pre-hardening candidate opened and round-tripped in Keynote 14.4, but the current strict owner rejects the app-authored source because model field 39 identifies an unowned conditional-style CalculationEngine dependency graph. The run and hashes are recorded in ADR 0008 as exploratory evidence only; no current-owner E3/E4 promotion follows, and the checked-in native fixture has no table. |
| Keynote soundtrack-item lifecycle | ✅ Focused owner tests cover bounded read/add/insert/replace/remove, shared occurrences, malformed references, aggregate-only preservation, atomic errors, exact apply/inverse, and stale-source conflicts; the focused fuzz target and corpus are present. The duplicate legacy item API is retired behind a boundary ratchet. One genuine replacement candidate passed Keynote save/close/reopen without repair; operation-wide native certification, soundtrack creation, general asset CRUD, and all whole-monolith deletion gates remain open. |
| Numbers persisted-sort/relocation physical compatibility | 🟡 Exact package snapshots stay on semantic focused transactions with terminal refusals; only historical source-built snapshots may use the private physical bridges for field 44 or table relocation. No public physical owner or ADR 0028 gate impact follows. |
| Broad Numbers baseline snapshot (2026-09-01; stale) | ❌ 405 unit tests passed/4 ignored; all integration targets passed except `table_data_list_reader_integration`, where 19 passed and 6 pre-existing comment/reply cases returned `InvalidSource` while validating synthetic comment metadata. This is a historical baseline snapshot, not a current-worktree result. |
| Numbers 14.4 Percentage probe | ✅ Rust candidate opened without repair; native save/close/reopen retained `4,200.000%`, three decimals, red-parentheses negatives, thousands separator, and Actual value `42`; strict Rust reread emitted an exact semantic no-op with zero touched components and the native-resaved SHA-256 unchanged |
| Historical broad Keynote baseline (2026-08-31) | ❌ 153 unit tests passed/1 failed; `soundtrack_order` integration was 3 passed/5 failed; rerun reproduced `selected_zip_suffix_and_central_records_allow_only_reassembly_fields` central-record failure |
| Documentation checks | ✅ `git diff --check` passed and every relative link in the audit and app-matrix documents resolves locally |
| Repository connectivity | ❌ `git fsck --connectivity-only --no-dangling` reports one missing commit and seven missing trees |

The four ignored Numbers native-oracle tests were not run because their private external fixture is unavailable.
The disposable Percentage probe establishes operation-specific native evidence, but formal E3/E4
promotion remains pending until its artifact and ledger are frozen and reproducible. The
Scientific record establishes operation-specific E3/E4 evidence for one existing-cell precision
change only: it does not certify arbitrary producers, other display families, package-wide
resource behavior, or a monolith deletion gate. The Fraction record establishes focused
source/build/test/fuzz plus operation-specific native E3/E4 evidence for the representative
Eighths-to-Hundredths scenario; field 20 true remains refused, and no all-variant native,
arbitrary-producer, native-byte-parity, package-wide performance, or monolith deletion-gate claim
is made. The Date & Time record establishes operation-specific E3/E4 evidence for the one
existing type-9 cell/pattern, including native save/close/exact-path reopen and a strict normalized
reread that was an exact byte-identical no-op; it does not generalize to other DateTime shapes,
producers, package bytes, or host-surface retirement. The Keynote physical-sort run was performed against a pre-hardening candidate; current
strict admission rejects the app-authored source's unproven model field 39. Its artifacts are
external exploratory evidence only, not current-owner E3/E4 evidence; native acceptance remains
pending until a current-admitted native source is rerun.

## Safety and integrity findings

Strengths include hard ceilings for Snappy, wire fields/depth/work, archive objects/messages, ZIP entries and sizes, exact source authorization, candidate reopen/readback, and inert handling of external resources.

Material follow-ups:

1. The legacy `litchi-iwa` directory and media paths use check-then-open flows without the pinned descriptor/no-follow design already available in `FrozenDirectoryBundle`; this leaves a symlink/TOCTOU window ([bundle path](../crates/litchi-iwa/src/bundle.rs#L745), [safe snapshot](../crates/litchi-iwa-archive/src/directory.rs#L443)).
2. `litchi-iwa-graph` insertion/traversal and `litchi-iwa-index` construction lack unified aggregate node/edge/work ceilings. Index buffers use fallible reservation, but graph materialization/traversal still grows collections infallibly.
3. Archive accounting has no single retained-memory or outer-plus-nested budget; the generic IWA path charges aggregate bytes after decoding, and replacement Deflate output is not checked against the configured compressed-member ceiling.
4. Encryption/password-protected packages are rejected. Cryptographic signature verification/re-signing is absent.
5. Focused/archive `write_to` APIs remain caller-owned streaming sinks and do not themselves flush, sync, rename, or provide durable publication. Pages, Numbers, and Keynote `Package::save` now delegate to the archive-owned atomic publication boundary; durable patch serialization/history remains absent and platform-specific durability limits stay explicit.
6. Public `Patch::is_noop()` semantics are not uniform: a few owners expose semantic equality while apply/commit restores exact-artifact authorization. The API contract should distinguish those predicates.
7. The shared `roundtrip_iwork` example can feed opaque compressed bytes into the Store-only logical writer. Unsupported compression must use preserve-mode reassembly or be rejected.

These are bounded-surface findings, not evidence of active content execution: no URL fetching, process execution, macro execution, or external-media retrieval was found.

## Worktree and repository constraints

At the first snapshot, 12 tracked paths were modified or deleted and additional files were untracked. The active work included Keynote table-title alias removal and sort-order hardening, a Pages table-lock seam, a Numbers table-info projection, protobuf/fuzz changes, boundary-checker changes, corpus seeds, and temporary artifacts. The workspace changed concurrently during the audit; the final pre-handoff check showed four tracked modifications (`protobuf.rs`, Keynote sort source/test, and a Numbers fuzz target) plus the untracked seams/artifacts. The verification results above are dirty-state diagnostics between those observations and must be rerun on a frozen revision. All WIP must be classified, wired, tested, and committed or removed before it enters a support claim.

Repository history is also incomplete locally. `git rev-list HEAD...@{upstream}` fails because commit `32bfe7da90d6d7133b4b4cfaa33bdbbc48971ece` references missing parent `1f61a7a6f478bc112e09a1e5ac8899ef2e665f97`; `git fsck --connectivity-only --no-dangling` reports that commit plus seven missing trees. Exact `HEAD` inspection works, but ancestry, ahead/behind, merge-base, and historical completeness claims are unreliable until the objects are restored.

## Recommended sequence

### P0 — make progress auditable

1. Restore the missing Git objects and establish a clean, reproducible audit revision.
2. Maintain the new authoritative app matrices against a frozen revision and keep focused-owner support, legacy-host-only support, opaque preservation, and native evidence distinct.
3. Keep the completed Percentage, Currency, Scientific, Fraction, Text, and Date & Time slices isolated from unresolved sort, lock, projection, and other untracked work before promoting any additional support claim.
4. Restore the boundary ratchet and fix the reproducible Numbers storage/comment and Keynote soundtrack/ZIP-preservation failures.

### P1 — close product and release gates

1. Migrate fresh-package, lifecycle, rich text/table, chart, shape, and media ownership from `litchi-iwa` into the concrete app crates, retiring debts by capability rather than by raw API count.
2. Add richer checked-in native fixtures and an operation-indexed native accept/save/close/reopen/readback ledger.
3. Add leaf-feature isolation and iWork macOS/Windows CI; include fixture and fuzz-corpus paths in triggers.
4. Install/pin `protoc` in publish CI and test each unpacked published crate.
5. Close the directory path race and add bounded graph/index plus aggregate archive-memory/work accounting.

### P2 — normalize the public story

1. Align README/rustdoc claims with the immutable root facade versus writable focused owners versus legacy host.
2. Add runnable examples for the focused table/chart/movie/comment/footnote APIs that currently have tests but no examples.
3. Document encryption, signatures, durable save, preview invalidation, formulas/external links, and Windows directory behavior as explicit unsupported or inert rows.
