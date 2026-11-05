//! Unified XLSX merged-range writes.
#![cfg(feature = "xlsx")]

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, OpenOptions, PropValue, Props, SourceMode, props};

const SHEET: &str = "/sheet[1]";
const CELL_A1: &str = "/sheet[1]/A1";
const CELL_B1: &str = "/sheet[1]/B1";
const CELL_D1: &str = "/sheet[1]/D1";
const CELL_F1: &str = "/sheet[1]/F1";
const CELL_H1: &str = "/sheet[1]/H1";
const RANGE_A1_B2: &str = "/sheet[1]/A1:B2";
const RANGE_D1_E2: &str = "/sheet[1]/D1:E2";
const RANGE_F1_G2: &str = "/sheet[1]/F1:G2";

fn fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("find Sheet1")
        .expect("Sheet1 worksheet");
    sheet.set("A1", "anchor").expect("seed anchor");
    sheet
        .set("D1", "second anchor")
        .expect("seed second anchor");
    sheet.set("F1", "third anchor").expect("seed third anchor");
    sheet
        .set("H1", "fourth anchor")
        .expect("seed fourth anchor");
    edit.commit()
        .expect("commit merge fixture")
        .into_package()
        .to_bytes()
        .expect("serialize merge fixture")
}

fn follower_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open merge fixture");
    let mut edit = workbook.edit().expect("start follower edit");
    edit.sheet("Sheet1")
        .expect("find Sheet1")
        .expect("Sheet1 worksheet")
        .set("B1", "follower")
        .expect("seed follower content");
    edit.commit()
        .expect("commit follower fixture")
        .into_package()
        .to_bytes()
        .expect("serialize follower fixture")
}

fn sparse_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create sparse workbook");
    let mut edit = workbook.edit().expect("start sparse edit");
    edit.sheet("Sheet1")
        .expect("find Sheet1")
        .expect("Sheet1 worksheet")
        .set("A1", "anchor")
        .expect("seed sparse anchor");
    edit.commit()
        .expect("commit sparse fixture")
        .into_package()
        .to_bytes()
        .expect("serialize sparse fixture")
}

fn two_sheet_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create two-sheet workbook");
    let mut edit = workbook.edit().expect("start two-sheet edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("find Sheet1")
            .expect("Sheet1 worksheet");
        sheet.set("A1", "first anchor").expect("seed first anchor");
        sheet
            .set("H1", "first payload")
            .expect("seed first payload");
    }
    {
        let mut sheet = edit.add("Sheet2").expect("add Sheet2");
        sheet
            .set("A1", "second anchor")
            .expect("seed second anchor");
        sheet
            .set("H1", "second payload")
            .expect("seed second payload");
    }
    edit.commit()
        .expect("commit two-sheet fixture")
        .into_package()
        .to_bytes()
        .expect("serialize two-sheet fixture")
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    let mode = if eager {
        SourceMode::Eager
    } else {
        SourceMode::Auto
    };
    litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode))
        .expect("open XLSX owner")
}

fn node_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve node")
        .view()
        .format()
        .expect("read node format")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn cell_property(file: &File, path: &str, key: &str) -> Option<String> {
    node_property(file, path, key)
}

fn root_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse merge selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

#[test]
fn range_merge_accepts_boolean_and_matching_a1_string_and_round_trips() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let merged = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("range merge=true");
        assert_eq!(
            cell_property(&merged, CELL_A1, "merge").as_deref(),
            Some("A1:B2")
        );
        assert_eq!(
            cell_property(&merged, CELL_A1, "mergeAnchor").as_deref(),
            Some("true")
        );
        assert_eq!(
            node_property(&merged, RANGE_A1_B2, "merge").as_deref(),
            Some("true")
        );

        let matching = original
            .set(RANGE_D1_E2, props! { "merge" => "D1:E2" })
            .expect("matching A1 merge string");
        assert_eq!(
            cell_property(&matching, CELL_D1, "merge").as_deref(),
            Some("D1:E2")
        );
        assert_eq!(
            cell_property(&matching, CELL_D1, "mergeAnchor").as_deref(),
            Some("true")
        );

        let published = merged.to_bytes().expect("publish range merge");
        let reopened = open_owner(&published, eager);
        assert_eq!(
            cell_property(&reopened, CELL_A1, "merge").as_deref(),
            Some("A1:B2")
        );
        let repeated = merged
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("repeat range merge")
            .to_bytes()
            .expect("publish repeated range merge");
        assert_eq!(repeated, published, "repeating a merge is an exact no-op");
        assert_ne!(published, original_bytes);
    }
}

#[test]
fn cell_and_sheet_range_lists_create_multiple_merges() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let cell = original
            .set(CELL_F1, props! { "merge" => "F1:G2,H1:I2" })
            .expect("cell range-list merge");
        assert_eq!(
            cell_property(&cell, CELL_F1, "merge").as_deref(),
            Some("F1:G2")
        );
        assert_eq!(
            cell_property(&cell, CELL_H1, "merge").as_deref(),
            Some("H1:I2")
        );

        let sheet = original
            .set(SHEET, props! { "merge" => "F1:G2,H1:I2" })
            .expect("sheet range-list merge");
        assert_eq!(
            cell_property(&sheet, CELL_F1, "merge").as_deref(),
            Some("F1:G2")
        );
        assert_eq!(
            cell_property(&sheet, CELL_H1, "merge").as_deref(),
            Some("H1:I2")
        );
    }
}

#[test]
fn explicit_merge_lists_selected_by_cells_are_planned_once_per_sheet() {
    let bytes = two_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(root_query("cell"), props! { "merge" => "B2:C3,D2:E3" })
            .expect("explicit merge list applies once for each selected sheet");

        for sheet in [1, 2] {
            assert_eq!(
                cell_property(&updated, &format!("/sheet[{sheet}]/B2"), "merge").as_deref(),
                Some("B2:C3")
            );
            assert_eq!(
                cell_property(&updated, &format!("/sheet[{sheet}]/D2"), "merge").as_deref(),
                Some("D2:E3")
            );
        }
    }
}

#[test]
fn explicit_merge_list_cache_does_not_hide_selected_node_limits() {
    let bytes = two_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let limits = SemanticLimits::builder()
            .max_expanded_targets(2)
            .expect("selected-node limit")
            .build()
            .expect("valid selected-node limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(root_query("cell"), props! { "merge" => "B2:C3" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("cached per-sheet planning must still count selected cells");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("serialize unchanged source"),
            bytes
        );
    }
}

#[test]
fn explicit_merge_geometry_can_share_one_edit_with_disjoint_cell_payloads() {
    let bytes = two_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|transaction| {
                transaction.set(
                    root_query("cell"),
                    Props::new()
                        .with("merge", "B2:C3")
                        .with("value", "rewritten"),
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("merge and disjoint selected-cell payloads share one edit");
        assert_eq!(commit.diagnostics().commits(), 1);

        for sheet in [1, 2] {
            let a1 = format!("/sheet[{sheet}]/A1");
            let h1 = format!("/sheet[{sheet}]/H1");
            let b2 = format!("/sheet[{sheet}]/B2");
            assert_eq!(
                commit
                    .file()
                    .get(&a1)
                    .expect("resolve selected A1")
                    .view()
                    .text()
                    .expect("read selected A1 text")
                    .as_deref(),
                Some("rewritten")
            );
            assert_eq!(
                commit
                    .file()
                    .get(&h1)
                    .expect("resolve selected H1")
                    .view()
                    .text()
                    .expect("read selected H1 text")
                    .as_deref(),
                Some("rewritten")
            );
            assert_eq!(
                cell_property(commit.file(), &b2, "merge").as_deref(),
                Some("B2:C3")
            );
        }
    }
}

#[test]
fn range_false_unmerges_exact_ranges_and_sweeps_contained_merges() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let no_existing = original
            .set(RANGE_A1_B2, props! { "merge" => false })
            .expect("unmerge absent range is a no-op");
        assert_eq!(
            no_existing.to_bytes().expect("publish absent unmerge"),
            original_bytes
        );

        let merged = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("create exact range");
        let unmerged = merged
            .set(RANGE_A1_B2, props! { "merge" => false })
            .expect("unmerge exact range");
        assert_eq!(cell_property(&unmerged, CELL_A1, "merge"), None);
        assert_eq!(
            node_property(&unmerged, RANGE_A1_B2, "merge").as_deref(),
            Some("false")
        );

        let two_merges = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("create first contained range")
            .set(RANGE_D1_E2, props! { "merge" => true })
            .expect("create second contained range");
        let swept = two_merges
            .set("/sheet[1]/A1:E2", props! { "merge" => "sweep" })
            .expect("sweep only contained ranges");
        assert_eq!(cell_property(&swept, CELL_A1, "merge"), None);
        assert_eq!(cell_property(&swept, CELL_D1, "merge"), None);

        let contained = merged
            .set("/sheet[1]/B1:B2", props! { "merge" => false })
            .expect_err("a non-exact contained unmerge must refuse");
        assert_eq!(contained.code(), Code::MergeNotExact);
        let overlap = merged
            .set("/sheet[1]/B1:C2", props! { "merge" => false })
            .expect_err("a partially overlapping unmerge must refuse");
        assert_eq!(overlap.code(), Code::MergeNotExact);
        let broad = merged
            .set("/sheet[1]/A1:E2", props! { "merge" => false })
            .expect_err("a broad non-exact unmerge must refuse");
        assert_eq!(broad.code(), Code::MergeNotExact);
    }
}

#[test]
fn merge_aliases_null_and_duplicate_properties_are_strictly_admitted() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for alias in ["true", "yes", "1"] {
            let merged = original
                .set(RANGE_A1_B2, props! { "merge" => alias })
                .expect("range truthy alias");
            assert_eq!(
                node_property(&merged, RANGE_A1_B2, "merge").as_deref(),
                Some("true")
            );
        }
        for alias in ["none", "empty", "unmerge", "false", "0", "no", ""] {
            let merged = original
                .set(RANGE_A1_B2, props! { "merge" => true })
                .expect("create merge for alias");
            let cleared = merged
                .set(RANGE_A1_B2, props! { "merge" => alias })
                .expect("merge alias clears exact range");
            assert_eq!(cell_property(&cleared, CELL_A1, "merge"), None, "{alias}");
            assert_eq!(
                node_property(&cleared, RANGE_A1_B2, "merge").as_deref(),
                Some("false"),
                "{alias} readback"
            );
        }

        let merged = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("merge before null reset");
        let cleared = merged
            .set(RANGE_A1_B2, Props::new().with("merge", PropValue::Null))
            .expect("null clears exact merge through registry policy");
        assert_eq!(cell_property(&cleared, CELL_A1, "merge"), None);

        let duplicate = original
            .set(
                RANGE_A1_B2,
                Props::new().with("merge", true).with("MERGE", false),
            )
            .expect_err("case-insensitive duplicate merge keys must refuse");
        assert_eq!(duplicate.code(), Code::InvalidArgument);
        assert_eq!(duplicate.property(), Some("MERGE"));
    }
}

#[test]
fn sheet_merge_range_list_shares_one_name_and_hidden_edit() {
    let bytes = two_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(
                "/sheet[1]",
                Props::new()
                    .with("name", "Data")
                    .with("hidden", true)
                    .with("merge", "A1:B2,C1:D2"),
            )
            .expect("sheet metadata and merge list share one native edit");

        assert_eq!(
            node_property(&updated, "/sheet[1]", "name").as_deref(),
            Some("Data")
        );
        assert_eq!(
            node_property(&updated, "/sheet[1]", "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_property(&updated, "/sheet[1]/A1", "merge").as_deref(),
            Some("A1:B2")
        );
        assert_eq!(
            cell_property(&updated, "/sheet[1]/C1", "merge").as_deref(),
            Some("C1:D2")
        );
        assert_eq!(
            node_property(&updated, "/sheet[2]", "name").as_deref(),
            Some("Sheet2")
        );
    }
}

#[test]
fn cell_false_unmerges_anchor_and_covered_follower_is_a_noop() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let merged = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("create merge");
        let unmerged = merged
            .set(CELL_A1, props! { "merge" => false })
            .expect("anchor cell false unmerge");
        assert_eq!(cell_property(&unmerged, CELL_A1, "merge"), None);

        let merged_again = unmerged
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("recreate merge");
        let follower = merged_again
            .set(CELL_B1, props! { "merge" => false })
            .expect("covered follower false is a no-op");
        assert_eq!(
            cell_property(&follower, CELL_A1, "merge").as_deref(),
            Some("A1:B2")
        );
        assert_eq!(
            follower.to_bytes().expect("publish follower no-op"),
            merged_again.to_bytes().expect("publish merged source")
        );

        let cell_enable = original
            .set(CELL_A1, props! { "merge" => true })
            .expect_err("cell merge=true is unsupported");
        assert_eq!(cell_enable.code(), Code::UnsupportedElement);

        let single = original
            .set("/sheet[1]/A1:A1", props! { "merge" => true })
            .expect_err("single-cell merge");
        assert_eq!(single.code(), Code::InvalidValue);
    }
}

#[test]
fn merge_creation_refuses_follower_content_and_overlaps_atomically() {
    let bytes = follower_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize follower source");
        let follower = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect_err("follower content must refuse merge creation");
        assert_eq!(follower.code(), Code::UnsupportedCapability);
        assert_eq!(
            original
                .to_bytes()
                .expect("serialize after follower refusal"),
            original_bytes
        );

        let empty = open_owner(&fixture(), eager);
        let existing = empty
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("create existing merge");
        let overlap = existing
            .set("/sheet[1]/B1:C2", props! { "merge" => true })
            .expect_err("overlapping merge must refuse");
        assert_eq!(overlap.code(), Code::UnsupportedCapability);
    }
}

#[test]
fn merge_transaction_replays_inverse_and_keeps_source_immutable() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let commit = original
            .edit(|transaction| {
                transaction.set(RANGE_A1_B2, props! { "merge" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("stage merge transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        let replayed = commit.patch().apply(&original).expect("replay merge patch");
        assert_eq!(
            replayed.to_bytes().expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse merge patch");
        assert_eq!(
            restored.to_bytes().expect("serialize inverse"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("serialize unchanged source"),
            original_bytes
        );
    }
}

#[test]
fn duplicate_merge_selectors_deduplicate_and_limits_are_checked_before_edit() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let merged = original
            .set(RANGE_A1_B2, props! { "merge" => true })
            .expect("create selector merge");
        let duplicate = merged
            .set(
                root_query("cell[mergeAnchor=true], cell[mergeAnchor=true]"),
                props! { "merge" => false },
            )
            .expect("duplicate anchor selector is deduplicated");
        assert_eq!(cell_property(&duplicate, CELL_A1, "merge"), None);

        let limits = SemanticLimits::builder()
            .max_staged_bytes(63)
            .expect("staged limit")
            .build()
            .expect("valid staged limit");
        let limited = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(RANGE_A1_B2, props! { "merge" => true })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("merge staging must honor the byte limit");
        assert_eq!(limited.code(), Code::LimitExceeded);
    }
}

#[test]
fn no_op_selected_cells_still_consume_the_expansion_limit() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("serialize no-op source");
        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("one selected cell")
            .build()
            .expect("valid target limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(root_query("cell"), props! { "merge" => false })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("selected no-op cells still count toward expansion");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("serialize after refusal"),
            before
        );
    }
}

#[test]
fn sweep_expansion_is_bounded_and_explicit_lists_deduplicate() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .unwrap()
            .build()
            .unwrap();
        let merged = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(CELL_A1, props! { "merge" => "A1:B2,A1:B2" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect("duplicate list entries consume one intent")
            .file()
            .set(RANGE_D1_E2, props! { "merge" => true })
            .expect("second merge");
        let before = merged.to_bytes().unwrap();
        let error = merged
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set("/sheet[1]/A1:E2", props! { "merge" => "sweep" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("one range selects two merge intents");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(merged.to_bytes().unwrap(), before);
        let swept = merged
            .set("/sheet[1]/A1:E2", props! { "merge" => "sweep" })
            .expect("bounded sweep");
        assert_eq!(cell_property(&swept, CELL_A1, "merge"), None);
        assert_eq!(cell_property(&swept, CELL_D1, "merge"), None);
    }
}

#[test]
fn huge_merge_is_one_structural_target_without_cell_expansion() {
    let bytes = sparse_fixture();
    let range = "/sheet[1]/A1:XFD1048576";
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("one structural target")
            .build()
            .expect("valid target limit");
        let merged = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(range, props! { "merge" => true })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect("large merge must not expand its cells");
        assert_eq!(
            cell_property(merged.file(), CELL_A1, "merge").as_deref(),
            Some("A1:XFD1048576")
        );
    }
}

#[test]
fn cell_payload_and_merge_share_one_native_edit_while_merge_is_a_barrier() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let mixed = original
            .edit(|transaction| {
                transaction.set(
                    CELL_D1,
                    props! { "value" => "new anchor", "merge" => "D1:E2" },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("cell payload and merge share one native edit");
        assert_eq!(mixed.diagnostics().operations(), 1);
        assert_eq!(mixed.diagnostics().commits(), 1);
        assert_eq!(
            cell_property(mixed.file(), CELL_D1, "merge").as_deref(),
            Some("D1:E2")
        );

        let barrier = mixed
            .file()
            .edit(|transaction| {
                transaction.set(CELL_A1, props! { "value" => "changed" })?;
                transaction.set(RANGE_F1_G2, props! { "merge" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("cell-to-merge transaction");
        assert_eq!(barrier.diagnostics().operations(), 2);
        assert_eq!(barrier.diagnostics().commits(), 2);
    }
}

#[test]
fn merge_ingress_and_multi_sheet_rename_are_bounded_before_publication() {
    for eager in [false, true] {
        let original = open_owner(&two_sheet_fixture(), eager);
        let before = original.to_bytes().unwrap();
        let limits = SemanticLimits::builder()
            .max_staged_bytes(128)
            .unwrap()
            .build()
            .unwrap();
        let repeated = std::iter::repeat_n("A1:B2", 30)
            .collect::<Vec<_>>()
            .join(",");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(CELL_A1, Props::new().with("merge", repeated.as_str()))?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("duplicate list text is bounded before deduplication");
        assert_eq!(error.code(), Code::LimitExceeded);
        let error = original
            .set(
                root_query("sheet"),
                props! { "name" => "Same", "merge" => "A1:B2" },
            )
            .expect_err("multi-sheet rename remains ambiguous with merge");
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(error.property(), Some("name"));
        assert_eq!(original.to_bytes().unwrap(), before);
    }
}

#[test]
fn unmerge_and_cell_payload_broadcast_share_one_atomic_native_commit() {
    for eager in [false, true] {
        let original = open_owner(&sparse_fixture(), eager)
            .set(RANGE_A1_B2, props! { "merge" => true })
            .unwrap();
        let commit = original
            .edit(|transaction| {
                transaction.set(
                    RANGE_A1_B2,
                    props! { "merge" => false, "value" => "filled" },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("unmerge precedes payload validation in one native edit");
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            node_property(commit.file(), RANGE_A1_B2, "merge").as_deref(),
            Some("false")
        );
        assert_eq!(
            commit
                .file()
                .get(CELL_B1)
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("filled")
        );
    }
}
