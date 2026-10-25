//! Unified XLSB cell writes through the public [`File`] and transaction APIs.
//!
//! The fixture is authored by the XLSB writer so every test is independent of
//! an external workbook.  Each behavior is run through both the retained
//! source owner and the eager owner; a cell write must leave the input handle
//! unchanged and publish a valid workbook that can be opened again.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::FileFormat;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::capabilities::Capabilities;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, PropValue, Props, SourceMode, props,
};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const A2: &str = "/sheet[1]/A2";
const B2: &str = "/sheet[1]/B2";
const SECOND_A1: &str = "/sheet[2]/A1";

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "seed");
    data.set_cell(0, 1, 2.5_f64);
    data.set_cell(0, 2, true);
    data.set_cell(1, 0, "second row");
    data.set_cell(1, 1, 4_i32);
    writer.add_worksheet(data);

    let mut archive = builder::Worksheet::new("Archive");
    archive.set_cell(0, 0, "untouched");
    archive.set_cell(0, 1, 9_i32);
    writer.add_worksheet(archive);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .expect("serialize XLSB fixture");
    output.into_inner()
}

fn named_fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("My Data");
    data.set_cell(0, 0, "seed");
    data.set_cell(0, 1, "second");
    writer.add_worksheet(data);
    writer.add_worksheet(builder::Worksheet::new("Archive"));

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build named XLSB fixture")
        .write_to(&mut output)
        .expect("serialize named XLSB fixture");
    output.into_inner()
}

fn chart_first_fixture() -> Vec<u8> {
    use litchi_xlsb::builder::chart::{Anchor, Chart};
    use litchi_xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut writer = Builder::new();
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("chart sheet");
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "seed");
    writer.add_worksheet(data);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build chart-first XLSB fixture")
        .write_to(&mut output)
        .expect("serialize chart-first XLSB fixture");
    output.into_inner()
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .expect("open XLSB owner")
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path} text: {error}"))
        .map(std::borrow::Cow::into_owned)
}

fn cell_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn cell_properties(file: &File, path: &str) -> Vec<(String, String)> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .map(|property| (property.key().to_owned(), property.value().to_string()))
        .collect()
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

fn workbook_cell_query() -> Query<'static> {
    workbook_query("cell")
}

fn workbook_query(selector: &'static str) -> Query<'static> {
    Query::parse(selector)
        .expect("parse cell selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

#[test]
fn direct_targets_honor_named_case_quoted_and_alias_paths() {
    let bytes = named_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish named fixture");
        let updated = original
            .set("'My Data'!a1", props! { "value" => "quoted" })
            .expect("quoted Excel sheet path")
            .set("/MY DATA/b1", props! { "value" => "root name" })
            .expect("case-insensitive root sheet path")
            .set("'my data'!C1", props! { "value" => "Excel name" })
            .expect("case-insensitive Excel sheet path")
            .set("'MY DATA'!D2:E2", props! { "value" => "range" })
            .expect("quoted named range path");

        assert_eq!(
            cell_text(&updated, "/sheet[1]/A1").as_deref(),
            Some("quoted")
        );
        assert_eq!(
            cell_text(&updated, "/sheet[1]/B1").as_deref(),
            Some("root name")
        );
        assert_eq!(
            cell_text(&updated, "/sheet[1]/C1").as_deref(),
            Some("Excel name")
        );
        for path in ["/sheet[1]/D2", "/sheet[1]/E2"] {
            assert_eq!(
                cell_text(&updated, path).as_deref(),
                Some("range"),
                "{path}"
            );
        }
        assert_eq!(
            original.to_bytes().expect("publish unchanged named base"),
            original_bytes
        );
    }
}

#[test]
fn direct_targets_guard_complete_catalog_tabs_and_missing_names() {
    let bytes = chart_first_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish chart-first fixture");
        let chart_error = original
            .set("/sheet[1]/A1", props! { "value" => "must refuse" })
            .expect_err("a chart tab cannot receive a cell write");
        assert_eq!(chart_error.code(), Code::UnsupportedType);

        let missing_error = original
            .set("Missing!A1", props! { "value" => "must refuse" })
            .expect_err("a missing semantic sheet must be refused");
        assert_eq!(missing_error.code(), Code::PathNotFound);

        let updated = original
            .set("data!a1", props! { "value" => "worksheet" })
            .expect("the worksheet after the chart tab remains addressable");
        assert_eq!(
            cell_text(&updated, "/sheet[2]/A1").as_deref(),
            Some("worksheet")
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged chart base"),
            original_bytes
        );
    }
}

#[test]
fn direct_targets_reject_invalid_coordinates_and_paths_atomically() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original fixture");
        for path in ["Data!A0", "Data!XFE1", "Data!A1048577", "Data!A1[0]"] {
            let error = original
                .set(path, props! { "value" => "must refuse" })
                .expect_err("invalid direct target must be refused");
            assert_eq!(error.code(), Code::InvalidPath, "{path}: {error:?}");
            assert_eq!(
                original.to_bytes().expect("publish after invalid path"),
                original_bytes,
                "invalid target mutated the {eager:?} owner: {path}"
            );
        }
    }
}

#[test]
fn semantic_cell_paths_survive_staged_reads_and_rename_barriers() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|transaction| {
                transaction.set("'Data'!A1", props! { "value" => "before rename" })?;
                {
                    let observed = transaction.get("/DATA/a1")?;
                    assert_eq!(observed.view().text()?.as_deref(), Some("before rename"));
                }
                transaction.set("/sheet[1]", props! { "name" => "Renamed" })?;
                transaction.set("'renamed'!b1", props! { "value" => "after rename" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("semantic paths across read and rename barriers");

        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 3);
        assert_eq!(
            sheet_property(commit.file(), "/sheet[1]", "name").as_deref(),
            Some("Renamed")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A1").as_deref(),
            Some("before rename")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/B1").as_deref(),
            Some("after rename")
        );
    }
}

#[test]
fn xlsb_cell_capabilities_advertise_scalar_and_clear_writes() {
    let cell = Capabilities::of(FileFormat::Xlsb)
        .element("cell")
        .expect("XLSB cell capability");
    let value = cell.property("value").expect("value capability");
    assert!(!value.get());
    assert!(value.set());
    let clear = cell.property("clear").expect("clear capability");
    assert!(!clear.get());
    assert!(clear.set());
    assert!(
        Capabilities::of(FileFormat::Xlsb)
            .element("range")
            .expect("range")
            .property("value")
            .expect("range value")
            .set()
    );
    let formula = cell.property("formula").expect("formula capability");
    assert!(formula.get());
    assert!(formula.set(), "formula text is authored (record 0261)");
}

#[test]
fn scalar_values_round_trip_for_source_and_eager_owners() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let before_format = cell_properties(&original, A1);

        let updated = original
            .set(A1, props! { "value" => "changed" })
            .expect("set text value")
            .set(B1, props! { "value" => 7.25 })
            .expect("set numeric value")
            .set(C1, props! { "value" => false })
            .expect("set boolean value");

        assert_eq!(cell_text(&updated, A1).as_deref(), Some("changed"));
        assert_eq!(cell_text(&updated, B1).as_deref(), Some("7.25"));
        assert_eq!(cell_text(&updated, C1).as_deref(), Some("FALSE"));
        assert_eq!(cell_properties(&updated, A1), before_format);
        assert_eq!(
            original.to_bytes().expect("publish detached original"),
            original_bytes,
            "the {eager:?} input owner was mutated by a cell edit"
        );
    }
}

#[test]
fn missing_cells_insert_and_clear_to_a_true_blank() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        let inserted = original
            .set("/sheet[1]/D4", props! { "value" => 11 })
            .expect("insert missing cell");
        assert_eq!(cell_text(&inserted, "/sheet[1]/D4").as_deref(), Some("11"));
        assert_eq!(
            cell_property(&inserted, "/sheet[1]/D4", "type").as_deref(),
            Some("Number")
        );

        let cleared = inserted
            .set("/sheet[1]/D4", props! { "clear" => true })
            .expect("clear inserted cell");
        assert_eq!(cell_text(&cleared, "/sheet[1]/D4").as_deref(), Some(""));
        assert_eq!(
            cell_property(&cleared, "/sheet[1]/D4", "empty").as_deref(),
            Some("true")
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged original"),
            original_bytes
        );
    }
}

#[test]
fn null_value_and_clear_true_are_explicit_blank_operations() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let cleared = original
            .set(A1, Props::new().with("value", PropValue::Null))
            .expect("null value clears a cell")
            .set(B1, props! { "clear" => true })
            .expect("clear true clears a cell");

        for path in [A1, B1] {
            assert_eq!(cell_text(&cleared, path).as_deref(), Some(""), "{path}");
            assert_eq!(
                cell_property(&cleared, path, "empty").as_deref(),
                Some("true"),
                "{path} blank property"
            );
        }
        assert_eq!(
            original.to_bytes().expect("publish detached original"),
            original_bytes
        );
    }
}

#[test]
fn exact_cell_noops_reuse_input_bytes_and_keep_the_base_immutable() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let same_text = original
            .set(A1, props! { "value" => "seed" })
            .expect("repeat text value");
        assert_eq!(
            same_text.to_bytes().expect("publish text no-op"),
            original_bytes
        );

        let same_number = original
            .set(B1, props! { "value" => 2.5 })
            .expect("repeat numeric value");
        assert_eq!(
            same_number.to_bytes().expect("publish number no-op"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged base"),
            original_bytes
        );
    }
}

#[test]
fn formulas_needing_workbook_context_are_a_typed_capability_gap() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .set(A1, props! { "formula" => "1+Archive!A1" })
            .expect_err("XLSB contextual formula authoring is unsupported");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.property(), Some("formula"));
        assert_eq!(
            original.to_bytes().expect("publish after formula refusal"),
            original_bytes
        );
    }
}

#[test]
fn ranges_cover_existing_and_missing_cells_and_preserve_the_base() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let updated = original
            .set("Data!A1:C2", props! { "value" => "range" })
            .expect("set a worksheet-qualified range");

        for path in [A1, B1, C1, A2, B2, "/sheet[1]/C2"] {
            assert_eq!(
                cell_text(&updated, path).as_deref(),
                Some("range"),
                "{path}"
            );
        }
        assert_eq!(cell_text(&updated, SECOND_A1).as_deref(), Some("untouched"));
        assert_eq!(
            original.to_bytes().expect("publish detached range base"),
            original_bytes
        );
    }
}

#[test]
fn workbook_scoped_query_targets_cells_on_both_worksheets() {
    let bytes = fixture();
    let selector = workbook_cell_query();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let matches = original
            .query(selector.clone())
            .expect("compile workbook cell selector")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("walk workbook cells");
        assert_eq!(matches.len(), 7, "fixture has seven stored cells");

        let commit = original
            .edit(|transaction| transaction.set(selector.clone(), props! { "value" => "query" }))
            .expect("set all query matches");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("query"));
        assert_eq!(
            cell_text(commit.file(), SECOND_A1).as_deref(),
            Some("query")
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged query base"),
            original_bytes
        );
    }
}

#[test]
fn mixed_worksheet_query_keeps_a_changed_prefix_when_the_last_cell_is_a_noop() {
    let original = open_owner(&fixture(), false);
    let original_bytes = original.to_bytes().expect("publish original union input");
    let selector = workbook_query("cell[value=seed or value=untouched]");
    let matches = original
        .query(selector.clone())
        .expect("compile cross-worksheet value query")
        .collect::<litchi_ops::Result<Vec<_>>>()
        .expect("walk cross-worksheet value query");
    assert_eq!(matches.len(), 2);
    let updated = original
        .set(selector, props! { "value" => "untouched" })
        .expect("set a cross-worksheet cell union");

    assert_eq!(cell_text(&updated, A1).as_deref(), Some("untouched"));
    assert_eq!(cell_text(&updated, SECOND_A1).as_deref(), Some("untouched"));
    assert_ne!(
        updated.to_bytes().expect("publish changed union"),
        original_bytes
    );
}

#[test]
fn distinct_same_sheet_cells_share_one_native_commit() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "coalesced" })?;
                transaction.set(B1, props! { "value" => 7.25 })?;
                transaction.set(C1, props! { "value" => false })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("coalesced same-sheet cell transaction");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(commit.diagnostics().leaf_operations(), 3);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("coalesced"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("7.25"));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some("FALSE"));

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay coalesced cells");
        assert_eq!(
            replayed.to_bytes().expect("publish replayed cells"),
            commit.file().to_bytes().expect("publish committed cells")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse coalesced cells");
        assert_eq!(
            restored.to_bytes().expect("publish restored cells"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged cell base"),
            original_bytes
        );
    }
}

#[test]
fn repeated_cell_is_a_barrier_and_last_value_wins() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "first" })?;
                transaction.set(B1, props! { "value" => 8 })?;
                transaction.set(A1, props! { "value" => "last" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeated-cell barrier transaction");

        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 2);
        assert_eq!(commit.diagnostics().leaf_operations(), 3);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("last"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("8"));
    }
}

#[test]
fn staged_reads_are_barriers_and_observe_prior_cell_writes() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before read" })?;
                {
                    let observed = transaction.get(A1)?;
                    assert_eq!(observed.view().text()?.as_deref(), Some("before read"));
                }
                transaction.set(B1, props! { "value" => 9 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("read barrier transaction");

        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 2);
        assert_eq!(commit.diagnostics().leaf_operations(), 2);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("before read"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("9"));
    }
}

#[test]
fn worksheet_name_changes_are_barriers_between_cell_runs() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before sheet" })?;
                transaction.set("/sheet[1]", props! { "name" => "Renamed" })?;
                transaction.set(B1, props! { "value" => "after sheet" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("worksheet-name barrier transaction");

        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 3);
        assert_eq!(commit.diagnostics().leaf_operations(), 3);
        assert_eq!(
            sheet_property(commit.file(), "/sheet[1]", "name").as_deref(),
            Some("Renamed")
        );
        assert_eq!(
            cell_text(commit.file(), A1).as_deref(),
            Some("before sheet")
        );
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("after sheet"));
    }
}

#[test]
fn empty_cell_props_are_exact_noops_like_xlsx() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let one_shot = original
            .set(A1, Props::new())
            .expect("empty cell properties are an admitted no-op");
        assert_eq!(
            one_shot.to_bytes().expect("publish empty-props no-op"),
            original_bytes
        );

        let commit = original
            .edit(|transaction| {
                transaction.set(A1, Props::new())?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("empty cell properties transaction");
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(commit.diagnostics().leaf_operations(), 0);
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit
                .file()
                .to_bytes()
                .expect("publish empty-props commit"),
            original_bytes
        );
        assert_eq!(
            original
                .to_bytes()
                .expect("publish unchanged empty-props base"),
            original_bytes
        );
    }
}

#[test]
fn cell_transactions_replay_inverse_and_accept_reopened_copies() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "staged" })?;
                transaction.set(SECOND_A1, props! { "value" => 13 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit cell transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 2);

        let replayed = commit.patch().apply(&original).expect("replay cell patch");
        assert_eq!(
            replayed.to_bytes().expect("publish replayed cell patch"),
            commit.file().to_bytes().expect("publish committed cells")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse cell patch");
        assert_eq!(
            restored.to_bytes().expect("publish restored cells"),
            original_bytes
        );

        let reopened = open_owner(&original_bytes, eager);
        assert_eq!(
            commit
                .patch()
                .apply(&reopened)
                .expect("a byte-identical copy takes the patch")
                .to_bytes()
                .expect("publish reopened replay"),
            commit.file().to_bytes().expect("publish committed cells")
        );
    }
}

#[test]
fn atomic_and_best_effort_batches_handle_mixed_cell_operations() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let atomic_error = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "atomic")),
                    batch_set(A2, Props::new().with("formula", "1+Archive!A1")),
                    batch_set(SECOND_A1, Props::new().with("value", "suffix")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must reject unsupported formula");
        assert_eq!(atomic_error.item(), Some(1));
        assert_eq!(atomic_error.code(), Code::UnsupportedCapability);
        assert_eq!(
            original.to_bytes().expect("publish atomic base"),
            original_bytes,
            "atomic failure published a partial cell batch"
        );

        let best_effort = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "best")),
                    batch_set(A2, Props::new().with("formula", "1+Archive!A1")),
                    batch_set(SECOND_A1, Props::new().with("value", "suffix")),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch keeps valid cell writes");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(result) if result.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error) if error.item() == Some(1)
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(result) if result.changed
        ));
        assert_eq!(cell_text(best_effort.file(), A1).as_deref(), Some("best"));
        assert_eq!(
            cell_text(best_effort.file(), SECOND_A1).as_deref(),
            Some("suffix")
        );
        let restored = best_effort
            .patch()
            .inverse()
            .apply(best_effort.file())
            .expect("inverse best-effort batch");
        assert_eq!(
            restored.to_bytes().expect("publish restored batch"),
            original_bytes
        );
    }
}

#[test]
fn invalid_properties_numbers_and_clear_false_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        for (path, properties, code) in [
            (
                A1,
                Props::new().with("unknownCellProperty", true),
                Code::UnsupportedProperty,
            ),
            (
                B1,
                Props::new().with("value", PropValue::Number(f64::NAN)),
                Code::InvalidValue,
            ),
            (B1, Props::new().with("value", -0.0), Code::InvalidValue),
            (
                B1,
                Props::new().with("value", f64::from_bits(1)),
                Code::InvalidValue,
            ),
            (C1, Props::new().with("clear", false), Code::InvalidValue),
            (
                A1,
                Props::new().with("value", "bad\u{1}text"),
                Code::InvalidValue,
            ),
        ] {
            let error = original
                .set(path, properties)
                .expect_err("invalid cell properties must refuse");
            assert_eq!(error.code(), code, "unexpected error for {path}: {error:?}");
            assert_eq!(
                original
                    .to_bytes()
                    .expect("publish unchanged invalid input"),
                original_bytes
            );
        }
    }
}

#[test]
fn target_and_snapshot_limits_fail_before_mutating_the_input() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let expanded = original
            .set("Data!A1:A65537", props! { "value" => "too-many" })
            .expect_err("oversized range must hit target limit");
        assert_eq!(expanded.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("publish after target refusal"),
            original_bytes
        );

        let strict_staged = SemanticLimits::builder()
            .max_staged_bytes(64)
            .expect("64-byte staged limit")
            .build()
            .expect("build staged limit");
        let staged = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(strict_staged),
                |transaction| transaction.set(A1, props! { "value" => "staged" }),
            )
            .expect_err("staged limit must reject before native edit");
        assert_eq!(staged.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("publish after staged refusal"),
            original_bytes
        );

        let probe = original
            .edit(|transaction| transaction.set(A1, props! { "value" => "snapshot" }))
            .expect("measure snapshot charge");
        let charge = probe.diagnostics().snapshot_bytes();
        assert!(charge > 0);
        let strict_snapshot = SemanticLimits::builder()
            .max_snapshot_bytes(charge - 1)
            .expect("snapshot bound below measured charge")
            .build()
            .expect("build snapshot limit");
        let snapshot = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(strict_snapshot),
                |transaction| transaction.set(A1, props! { "value" => "snapshot" }),
            )
            .expect_err("snapshot limit must reject before publication");
        assert_eq!(snapshot.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("publish after snapshot refusal"),
            original_bytes
        );
    }
}

#[test]
fn saved_cell_write_reopens_for_source_and_eager_routes() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(A1, props! { "value" => "saved" })
            .expect("set saved value");
        let direct = updated.to_bytes().expect("publish updated XLSB");
        let reopened = open_owner(&direct, eager);
        assert_eq!(cell_text(&reopened, A1).as_deref(), Some("saved"));

        let directory = tempfile::tempdir().expect("save directory");
        let destination = directory.path().join("cells.xlsb");
        updated.save_as(&destination).expect("save updated cells");
        let saved = litchi_ops::open(&destination).expect("reopen saved cells");
        assert_eq!(cell_text(&saved, A1).as_deref(), Some("saved"));
        assert_eq!(std::fs::read(destination).expect("read saved XLSB"), direct);
    }
}

#[test]
fn atomic_batch_reports_exact_noop_per_item() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let result = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "seed")),
                    batch_set(A2, Props::new().with("value", 77.0)),
                ],
                BatchMode::Atomic,
            )
            .expect("mixed batch");
        assert!(matches!(&result.value()[0], ItemResult::Applied(result) if !result.changed));
        assert!(matches!(&result.value()[1], ItemResult::Applied(result) if result.changed));
    }
}
