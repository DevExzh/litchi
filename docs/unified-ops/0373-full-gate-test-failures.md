# 0373 — Full-gate test failures on trunk c594870566

The coordinator's full gate (`tools/gate/run_gate.sh`, run
`full-c594870566`) failed rows 30 (`tests-leaves`) and 33
(`tests-workspace-excluding-ods-ops`). Row 33 repeats row 30's failures and
adds two more (iwa-core, odf-formula). Every other row that had finished
(0–29, 31, 32, 34–39) passed. This record lists each failure, its cause, the
fix and the commit or record responsible. No test was deleted or ignored.

For every failure the responsible commit was found with `git log -S`/`-G`,
and for the two whose history did not settle it, with `git bisect run` on the
failing test (scripts kept outside the repository in
`.coord/fin-tests/scripts/`).

| # | Failure | Verdict | Cause (commit / record) | Fix |
|---|---|---|---|---|
| 1 | `litchi-drawingml` doctest `chart (line 31)`: E0277, `?` cannot convert into `std::io::Error` | test stale | 96ea764602 / 6637852e88 (records 0362dve, 0362dvea) changed `chart::writer::write` to return `litchi_drawingml::Result` but left the example's `Ok::<(), std::io::Error>` | example ends with `Ok::<(), litchi_drawingml::Error>(())` |
| 2 | `litchi-opc` `pkgwriter::tests::replacement_with_identical_edges_uses_its_own_relationship_xml` (`owned=true`) | **code regression** | c0d1c2499c (record 0362ul87c, "align VBA retained admission with writer") added `OpcPackage::source_relationships_retained`. Its open-time-capture shortcut was missing the check that the package still holds the owner's source relationships token. A replacement part that cloned the original's `Relationships` (capture included) therefore copied the source `.rels` member. `add_part` drops that token so that the replacement publishes its own relationships (f87b43cafd) | the capture shortcut now also requires `source_relationships_member_present(owner)`, matching the plan's `relationships_pristine` test. Admission metrics use the same function, so they agree with the writer |
| 3 | `litchi-opc` `xml_bytes_publication::authored_xml_add_checks_attribute_part_and_aggregate_limits_atomically`: `InvalidLimit { RelationshipTargetBytes, 4096, 3 }` | test wrong (never passed) | 9e2d43df11 wrote the test with `max_xml_attribute_bytes(3)` alone. `SourceLimitsBuilder::build` has required relationship targets ≤ the XML attribute ceiling since d1b775d443 (2026-08-09) | the test also sets `max_relationship_target_bytes(3)` and still asserts `XmlAttributeBytes` |
| 4 | `litchi-xlsx` `metadata::tests::no_op_edits_return_the_same_snapshot` and `no_op_patch_replay_still_requires_its_recorded_property_scope` | **code regression** | 78bd63c56f (record 0362ul87ac) made the shared `properties::Plan::set_metadata` force creation of an absent core part. That was for DOCX's `None`/`Some(default)` endpoints. XLSX, PPTX and XLSB read an absent part as `Metadata::default()`, so setting the value they read became a changed commit. This breaks the 0351 §2.3 *No-op* row and §2.6 fidelity ("tests pass unchanged") | `set_metadata` no longer forces the core part. `set_core_metadata` keeps forcing it. DOCX's bridge (`package/package/metadata.rs`) is the only caller that needs explicit creation, and it uses `set_core_metadata`, so DOCX keeps the 0362ul87ac behaviour (its tests pass) |
| 5 | `litchi-xlsx` `error_classification_tests::operation_boundary_join_preserves_both_typed_sources_through_wrapper`: "typed JoinError" | test wrong (never passed) | 95973e8c6a (records 0362ul94h/ha/hb, "no Cargo run") expected `source(Error::Join)` to be the `JoinError`. `Error::Join` has been `#[error(transparent)]` since f496d2d856, so its source is the `JoinError`'s own source, the retained leaf. 0362ul94hb says the production API does not change | the test matches `Error::Join(join)` directly, checks `join.normalization_source()`, and checks that the wrapper's standard source is the typed leaf |
| 6 | `litchi-xlsx` `xlsx_slicer_timeline_crud::raw_slicer_rejects_a_misdeclared_lazy_member_without_mutation`: `insert_cache` returned `Ok` | test asserts superseded behaviour | bisect: passes at c4e646f8ac, fails from the VBA series. At the good commit a backtrace showed the rejection came from `plan_source_relationships` calling `get_part(owner)`, which hydrated the unrelated lazy member as a side effect. d692fa25ac (record 0362ul87c family) made relationship planning metadata-only (`part_metadata`), as ADR 0030 and 0360adc require ("size-only admission … leaves deferred payloads cold"). The mismatch is now refused at `Edit::commit` with the same `OpcError::ZipError` | the test stages the insert (asserting the member stays cold), expects `commit()` to refuse with `ZipError`, keeps the no-charge counter assertion, and checks the source still has no `Misdeclared` cache |
| 7 | `litchi-iwa-core` `framing::retained_unknown_header_obeys_stricter_field_limits`: all four results `false` | **code regression** | bisect: first bad commit 846a0fea0b ("perf(iwa): cache retained header admissions"). The cached `HeaderAdmission` refuses with the whole-header total (`observed: 3`). Parse refuses with the streaming preflight's first failure (`observed: 2`), so publication and parse disagreed | `admit_retained_header` still accepts from the cached totals with no rewalk. On a refusal it re-runs the streaming preflight, so it reports the same diagnostic as parse (falling back to the cached error if the preflight admits). Only the refusal path changes |
| 8 | `litchi-odf-formula` `fifth_wave::deterministic_merge_transfer_and_history_properties_reopen_and_reverse`: `History::apply` on a stale origin was not `PackageError::PatchConflict` | **code regression** (never passed) | 8be62f1859 ("close public family error boundaries") moved `History` to the family error and changed the test to expect `litchi_odf_common::Error::PatchConflict`. `History::apply` still used `?` on the core `Error::PatchConflict`, which `From<litchi_core::Error>` wraps as `Error::Core` | `History::apply` checks each link with `is_applicable_to` and returns `Error::patch_conflict("odf_formula.history.apply", …)`, the same way `push` and `Package::apply` do |

## Verification

Each failing command was rerun and passes:

- `cargo test -p litchi-drawingml --doc`
- `cargo test -p litchi-opc --all-features --lib`
- `cargo test -p litchi-opc --all-features --test xml_bytes_publication`
- `cargo test -p litchi-xlsx --all-features --lib metadata`
- `cargo test -p litchi-xlsx --all-features --test xlsx_slicer_timeline_crud`
- `cargo test -p litchi-iwa-core --all-features`
- `cargo test -p litchi-odf-formula --all-features`

The full suites of the affected crates also pass (exit 0):

- `cargo test -p litchi-opc -p litchi-drawingml -p litchi-ooxml-common --all-features --no-fail-fast`
- `cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx -p litchi-xlsb --all-features --no-fail-fast`. The shared properties change reaches every OOXML metadata consumer.
- `cargo test -p litchi-ops -p litchi --all-features --no-fail-fast`

Static checks pass:

- `cargo fmt --all --check`
- `cargo clippy -p litchi-opc -p litchi-drawingml -p litchi-ooxml-common -p litchi-xlsx -p litchi-iwa-core -p litchi-odf-formula -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` on the changed crates

## Performance

Only refusal or slow paths changed:

- OPC: one extra map lookup in a branch that was already serializing relationships.
- iwa-core: a second walk only when a retained header is refused.
- odf-formula: one byte comparison per history link, which `Patch::apply` repeats anyway.
- properties: a flag no longer set.

No hot path was re-measured.

## Commits

- 25f045c746: OPC replacement relationships (2)
- ecde66248b: OPC limit test (3)
- a7d3a14efe: shared metadata no-op (4)
- 6ec5d40a2c: XLSX join wrapper and lazy slicer tests (5, 6)
- da5d33cdb7: DrawingML example (1)
- a56ef7b015: iwa-core refusal diagnostic (7)
- 04362a8333: odf-formula history conflict (8)
