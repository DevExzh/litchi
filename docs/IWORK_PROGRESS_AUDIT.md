# iWork Progress Audit

> Audit baseline: 2026-09-01 at committed `b6e6ada83` (`feat(numbers): own plain number cell formats`). This update evaluates the focused Percentage-format slice completed from that baseline; it remains scoped verification rather than release certification.

## Conclusion

The iWork implementation is substantial but is **not suite-complete or release-certified**.

- Pages, Keynote, and Numbers have credible bounded semantic readers.
- The focused owner crates have strong exact-source, fail-closed transactions for selected existing objects.
- Broad fresh creation, structural editing, rich formatting, charts, shapes, and media remain largely in the legacy `litchi-iwa` migration host.
- Three authoritative Pages, Keynote, and Numbers matrices now report focused-owner, legacy-host, preservation, and evidence boundaries; the existing implementation review still explicitly excludes iWork.
- Native fixtures establish basic parse/no-op fidelity, but there is no suite-wide automated proof that Litchi-modified files are accepted, saved, closed, and reopened by all three native applications.
- The focused Percentage owner and codec checks are green at the synthetic/source level. A disposable native Numbers open/save/close/reopen probe and strict Rust semantic no-op readback are recorded in ADR 0008, but formal E3/E4 promotion and frozen-fixture ledger evidence remain pending.

The detailed provisional coverage assessment is in [IWORK_FEATURE_MATRIX_AUDIT.md](IWORK_FEATURE_MATRIX_AUDIT.md).

## Current maturity

| Layer | What is solid | What still lacks | Assessment |
|---|---|---|---|
| Root `litchi::iwork` facade | Immutable, bounded Pages/Keynote/Numbers projection; host-free dependency path | Read-only and text-oriented; no package/edit, metadata, media, chart, or formula editing/recalculation surface | 🟡 focused read facade |
| `litchi-pages` | ZIP/directory semantic read; exact retained-package output; selected text, layout, footnote, header/footer, and table-property transactions | Fresh package creation, section/table structural CRUD, table cells/formulas, rich formatting, drawings, charts, media, collaboration, export | 🟡 bounded reader/editor |
| `litchi-keynote` | Show/slide read; existing-slide state, text, notes, settings, background, transition, selected chart/movie/table properties | Slide creation/duplication, full table data, chart data/series/CRUD, media bytes/CRUD, full build/animation editing, shapes/groups, masters/themes, collaboration, rendering/export | 🟡 bounded reader/editor |
| `litchi-numbers` | Rooted workbook read; selected scalar cells, formulas, controls, table settings, names/order, comments/replies, and existing-cell Number/Percentage display formats | Sheet/table lifecycle, row/column topology, full formula engine, other generic formats/rich styles, package merge editing, charts/media/drawables, filters/pivots/categories, print/page setup and workbook export | 🟡 bounded reader/editor |
| `litchi-iwa` | Broad source-free authoring and native mutation across all three applications; extensive tests/examples | It is explicitly a compatibility/migration host, exposes native/raw seams, and still owns most rich/structural authoring | 🟡 broad but non-canonical |
| Shared IWA crates | Bounded Snappy/wire/archive parsing, package preservation, detection, focused codecs, exact artifacts, COW state, and archive-owned durable publication | Concrete-owner index adoption, aggregate graph/memory budgets, directory write parity, durable patch serialization/history, encryption/signatures | 🟡 mature substrate with open boundaries |

This assessment deliberately does not infer semantic support from generated protobuf availability or raw public declaration counts.

## Architecture and migration

ADR 0028 defines `litchi-iwa` as the sole legacy migration host and makes deletion conditional on closing every recorded edge and passing native reopen gates ([ADR 0028](adr/0028-iwa-monolith-exit.md#deletion-gate)). The authoritative current topology is **64 workspace packages, 237 internal dependency declarations, 226 canonical edges, and 11 ordered migration debts**, with one migration host. The debt orders are `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]` ([boundary ledger](../tools/crate_boundaries.json)); these are ordered host-exit conditions, not a count of all IWA-related crates.

Positive progress:

- The root facade depends on the three concrete owners rather than `litchi-iwa` ([feature wiring](../crates/litchi/Cargo.toml#L60), [facade](../crates/litchi/src/iwork/mod.rs#L1)).
- Recent waves moved narrow Pages table/text owners, Numbers cells/settings/comments owners, and Keynote chart/movie/table owners into focused crates.
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
Percentage and current broad Numbers rows were rerun for this slice; the Keynote row remains a
historical dirty-worktree diagnostic and is not evidence against the focused Percentage
implementation:

| Check | Result |
|---|---|
| `python3 tools/check_crate_boundaries.py` | ❌ Three errors: untracked `crates/litchi-iwa/src/pages/editor/tables/lock.rs` restores the retired `body_table_lock_state` and `set_body_table_lock_state` host surface |
| `cargo test -p litchi-numbers --test table_cell_percentage_format --no-fail-fast` | ✅ 19/19 focused Percentage owner tests passed, including selector, family-boundary, malformed-input, locality, inverse, concurrency, and budget cases |
| `cargo test -p litchi-iwa-protos percentage_format --no-fail-fast` | ✅ 4/4 focused Percentage codec tests passed, covering native domains, canonical framing, unknown preservation, and Number/Percentage separation |
| Current broad Numbers gate | ❌ 405 unit tests passed/4 ignored; all integration targets passed except `table_data_list_reader_integration`, where 19 passed and 6 pre-existing comment/reply cases returned `InvalidSource` while validating synthetic comment metadata |
| Numbers 14.4 Percentage probe | ✅ Rust candidate opened without repair; native save/close/reopen retained `4,200.000%`, three decimals, red-parentheses negatives, thousands separator, and Actual value `42`; strict Rust reread emitted an exact semantic no-op with zero touched components and the native-resaved SHA-256 unchanged |
| Historical broad Keynote baseline (2026-08-31) | ❌ 153 unit tests passed/1 failed; `soundtrack_order` integration was 3 passed/5 failed; rerun reproduced `selected_zip_suffix_and_central_records_allow_only_reassembly_fields` central-record failure |
| Documentation checks | ✅ `git diff --check` passed and every relative link in the audit and app-matrix documents resolves locally |
| Repository connectivity | ❌ `git fsck --connectivity-only --no-dangling` reports one missing commit and seven missing trees |

The four ignored Numbers native-oracle tests were not run because their private external fixture is unavailable.
The disposable Percentage probe establishes operation-specific native evidence, but formal E3/E4
promotion remains pending until its artifact and ledger are frozen and reproducible.

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
3. Keep the completed Percentage slice isolated from unresolved sort, lock, projection, and other untracked drafts before promoting any additional support claim.
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
