//! XLSB sheet-state transaction coalescing through the unified facade.
//!
//! A run of direct worksheet-tab state assignments is deliberately narrow:
//! direct paths, one exact `SheetState`, and distinct tabs.  This file keeps
//! the observable contract for that run separate from the broader visibility
//! and cell-write suites: native/source parity, view-state preservation,
//! barriers, deferred native errors, and transaction resource bounds.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_opc::{OpcPackage, PackURI};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, Props, SourceMode, props,
};
use litchi_xlsb::raw::{Encoder, Records, kind};

const FIRST: &str = "/sheet[1]";
const SECOND: &str = "/sheet[2]";
const THIRD: &str = "/sheet[3]";
const FOURTH: &str = "/sheet[4]";

fn fixture() -> Vec<u8> {
    fixture_with_count(4)
}

fn fixture_with_count(count: usize) -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let names = ["Data", "Archive", "Report", "Tail", "Spare", "Keep"];
    for index in 1..=count {
        let name = names
            .get(index - 1)
            .map_or_else(|| format!("Tab{index}"), ToString::to_string);
        let mut sheet = builder::Worksheet::new(name);
        sheet.set_cell(0, 0, format!("seed-{index}"));
        writer.add_worksheet(sheet);
    }

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB sheet-state fixture")
        .write_to(&mut output)
        .expect("serialize XLSB sheet-state fixture");
    output.into_inner()
}

fn worksheet_chart_worksheet_fixture() -> Vec<u8> {
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
    writer.add_worksheet(builder::Worksheet::new("Data"));
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("chart tab");
    writer.add_worksheet(builder::Worksheet::new("Tail"));

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build worksheet-chart-worksheet fixture")
        .write_to(&mut output)
        .expect("serialize worksheet-chart-worksheet fixture");
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

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path} text: {error}"))
        .map(std::borrow::Cow::into_owned)
}

fn scoped_sheet_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse sheet selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_set(path: impl Into<String>, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into().into());
    item.props = Some(properties);
    item
}

fn visibility_batch(paths: &[usize], state: &'static str) -> Vec<BatchItem<'static>> {
    paths
        .iter()
        .map(|index| {
            batch_set(
                format!("/sheet[{index}]"),
                Props::new().with("visibility", state),
            )
        })
        .collect()
}

fn without_chart_views(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse chart fixture");
    let chart_uri = package
        .iter_parts()
        .find(|part| part.content_type() == "application/vnd.ms-excel.chartsheet")
        .expect("chart part")
        .partname()
        .clone();
    let original = package
        .get_part(&chart_uri)
        .expect("chart part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    let mut in_views = false;
    let mut removed = false;
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = item.expect("chart record");
            if record.kind() == kind::BEGIN_CS_VIEWS {
                in_views = true;
                removed = true;
                continue;
            }
            if in_views {
                if record.kind() == kind::END_CS_VIEWS {
                    in_views = false;
                }
                continue;
            }
            writer
                .write_record(record.kind(), record.payload())
                .expect("chart record");
        }
    }
    assert!(removed && !in_views);
    package
        .get_part_mut(&chart_uri)
        .expect("chart part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("serialize chart fixture without views");
    output
}

fn worksheet_selected(bytes: &[u8], index: usize) -> bool {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let uri = PackURI::new(format!("/xl/worksheets/sheet{index}.bin")).expect("worksheet URI");
    let part = package.get_part(&uri).expect("worksheet part");
    let record = Records::new(part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .map(Result::unwrap)
        .find(|record| record.kind() == kind::BEGIN_WS_VIEW)
        .expect("worksheet view record");
    let flags = u16::from_le_bytes(record.payload()[..2].try_into().expect("view flags"));
    flags & 0x0040 != 0
}

fn assert_view_parity(left: &[u8], right: &[u8], worksheets: usize) {
    assert_eq!(
        litchi_xlsb::Workbook::new(Cursor::new(left))
            .expect("open left workbook")
            .active_catalog_position(),
        litchi_xlsb::Workbook::new(Cursor::new(right))
            .expect("open right workbook")
            .active_catalog_position(),
    );
    for index in 1..=worksheets {
        assert_eq!(
            worksheet_selected(left, index),
            worksheet_selected(right, index),
            "worksheet {index} selected flag differs",
        );
    }
}

fn assert_all_applied_changed(results: &[ItemResult]) {
    for result in results {
        match result {
            ItemResult::Applied(applied) => assert!(applied.changed),
            ItemResult::Failed(error) => panic!("unexpected batch failure: {error}"),
            _ => panic!("unexpected batch result"),
        }
    }
}

#[test]
fn distinct_direct_sheet_states_coalesce_for_closure_and_atomic_json() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let source_bytes = source.to_bytes().expect("publish source");

        let sequential = source
            .set(SECOND, props! { "visibility" => "hidden" })
            .expect("hide second worksheet")
            .set(THIRD, props! { "hidden" => true })
            .expect("hide third worksheet");
        let sequential_bytes = sequential.to_bytes().expect("publish sequential states");

        let closure = source
            .edit(|edit| {
                edit.set(SECOND, props! { "visibility" => "hidden" })?;
                edit.set(THIRD, props! { "hidden" => true })
            })
            .expect("coalesce direct worksheet states");
        assert!(closure.diagnostics().changed());
        assert_eq!(closure.diagnostics().operations(), 2);
        assert_eq!(closure.diagnostics().commits(), 1);
        assert_eq!(closure.diagnostics().leaf_operations(), 2);
        let closure_bytes = closure.file().to_bytes().expect("publish closure states");
        assert_eq!(closure_bytes, sequential_bytes);
        assert_view_parity(&closure_bytes, &sequential_bytes, 4);
        assert_eq!(
            source.to_bytes().expect("source remains unchanged"),
            source_bytes
        );
        assert_eq!(
            closure
                .patch()
                .inverse()
                .apply(closure.file())
                .expect("apply inverse")
                .to_bytes()
                .expect("publish inverse"),
            source_bytes
        );

        let batch = source
            .batch(
                [
                    batch_set(SECOND, Props::new().with("visibility", "hidden")),
                    batch_set(THIRD, Props::new().with("hidden", true)),
                ],
                BatchMode::Atomic,
            )
            .expect("coalesce atomic JSON states");
        assert_all_applied_changed(batch.value());
        assert_eq!(batch.diagnostics().operations(), 2);
        assert_eq!(batch.diagnostics().commits(), 1);
        assert_eq!(batch.diagnostics().leaf_operations(), 2);
        let batch_bytes = batch.file().to_bytes().expect("publish JSON states");
        assert_eq!(batch_bytes, sequential_bytes);
        assert_view_parity(&batch_bytes, &sequential_bytes, 4);
    }
}

#[test]
fn hiding_the_active_tab_forces_a_run_boundary_before_the_next_hide() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let sequential = source
            .set(FIRST, props! { "visibility" => "hidden" })
            .expect("hide active worksheet")
            .set(SECOND, props! { "visibility" => "hidden" })
            .expect("hide the relocated active worksheet");
        let sequential_bytes = sequential
            .to_bytes()
            .expect("publish sequential active hides");

        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "visibility" => "hidden" })?;
                edit.set(SECOND, props! { "visibility" => "hidden" })
            })
            .expect("active hide must flush before the next hide");
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 2);
        let bytes = commit.file().to_bytes().expect("publish active hides");
        assert_eq!(bytes, sequential_bytes);
        assert_view_parity(&bytes, &sequential_bytes, 4);
        assert_eq!(
            litchi_xlsb::Workbook::new(Cursor::new(&bytes))
                .expect("reopen active hides")
                .active_catalog_position(),
            Some(2)
        );

        let batch = source
            .batch(
                [
                    batch_set(FIRST, Props::new().with("visibility", "hidden")),
                    batch_set(SECOND, Props::new().with("visibility", "hidden")),
                ],
                BatchMode::Atomic,
            )
            .expect("atomic JSON active hides");
        assert_all_applied_changed(batch.value());
        assert_eq!(batch.diagnostics().commits(), 2);
        assert_eq!(
            batch.file().to_bytes().expect("publish JSON active hides"),
            bytes
        );
    }
}

#[test]
fn active_hide_does_not_bypass_a_missing_next_chart_view() {
    let bytes = without_chart_views(&worksheet_chart_worksheet_fixture());
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let error = source
            .batch(
                [
                    batch_set(FIRST, Props::new().with("hidden", true)),
                    batch_set(SECOND, Props::new().with("visibility", "hidden")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("missing next chart view must fail at active hide");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.item(), Some(0));
        assert_eq!(source.to_bytes().expect("publish unchanged source"), before);
    }
}

#[test]
fn exact_sheet_state_noops_report_false_changed_and_preserve_bytes() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let hidden = source
            .set(FIRST, props! { "visibility" => "hidden" })
            .expect("create hidden source");
        let before = hidden.to_bytes().expect("publish hidden source");
        let commit = hidden
            .batch(
                [
                    batch_set(FIRST, Props::new().with("visibility", "hidden")),
                    batch_set(SECOND, Props::new().with("hidden", false)),
                ],
                BatchMode::Atomic,
            )
            .expect("exact visibility no-op batch");
        assert_eq!(commit.value().len(), 2);
        assert!(
            commit.value().iter().all(|result| {
                matches!(result, ItemResult::Applied(applied) if !applied.changed)
            })
        );
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(commit.file().to_bytes().expect("publish no-op"), before);
    }
}

#[test]
fn mixed_hide_reveal_and_same_sheet_barriers_preserve_sequential_result() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let sequential = source
            .set(SECOND, props! { "visibility" => "hidden" })
            .expect("hide second")
            .set(THIRD, props! { "visibility" => "hidden" })
            .expect("hide third")
            .set(SECOND, props! { "visibility" => "visible" })
            .expect("reveal second");
        let sequential_bytes = sequential
            .to_bytes()
            .expect("publish sequential mixed states");

        let commit = source
            .edit(|edit| {
                edit.set(SECOND, props! { "visibility" => "hidden" })?;
                edit.set(THIRD, props! { "visibility" => "hidden" })?;
                edit.set(SECOND, props! { "visibility" => "visible" })
            })
            .expect("mixed hide/reveal transaction");
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 2);
        let bytes = commit.file().to_bytes().expect("publish mixed states");
        assert_eq!(bytes, sequential_bytes);
        assert_view_parity(&bytes, &sequential_bytes, 4);
        assert_eq!(
            sheet_property(commit.file(), SECOND, "visibility").as_deref(),
            None
        );
        assert_eq!(
            sheet_property(commit.file(), THIRD, "visibility").as_deref(),
            Some("hidden")
        );

        let batch = source
            .batch(
                [
                    batch_set(SECOND, Props::new().with("visibility", "hidden")),
                    batch_set(THIRD, Props::new().with("visibility", "hidden")),
                    batch_set(SECOND, Props::new().with("visibility", "visible")),
                ],
                BatchMode::Atomic,
            )
            .expect("mixed JSON hide/reveal transaction");
        assert_eq!(batch.diagnostics().commits(), 2);
        assert_eq!(
            batch.file().to_bytes().expect("publish mixed JSON states"),
            bytes
        );
    }
}

#[test]
fn name_combined_and_query_sets_are_flush_barriers_and_queries_see_staged_state() {
    for eager in [false, true] {
        let source = open_owner(&fixture_with_count(5), eager);
        let renamed = source
            .edit(|edit| {
                edit.set(FIRST, props! { "name" => "Input" })?;
                edit.set(
                    scoped_sheet_query("sheet[name=Input]"),
                    props! { "visibility" => "hidden" },
                )
            })
            .expect("query should resolve staged rename");
        assert_eq!(renamed.diagnostics().operations(), 2);
        assert_eq!(renamed.diagnostics().commits(), 2);
        assert_eq!(
            sheet_property(renamed.file(), FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(renamed.file(), FIRST, "visibility").as_deref(),
            Some("hidden")
        );

        let observed = source
            .edit(|edit| {
                edit.set(FIRST, props! { "visibility" => "hidden" })?;
                edit.set(
                    scoped_sheet_query("sheet[visibility=hidden]"),
                    props! { "name" => "Hidden" },
                )
            })
            .expect("query should resolve staged visibility");
        assert_eq!(observed.diagnostics().operations(), 2);
        assert_eq!(observed.diagnostics().commits(), 2);
        assert_eq!(
            sheet_property(observed.file(), FIRST, "name").as_deref(),
            Some("Hidden")
        );
        assert_eq!(
            sheet_property(observed.file(), FIRST, "visibility").as_deref(),
            Some("hidden")
        );

        let combined = source
            .edit(|edit| {
                edit.set(
                    FIRST,
                    props! {
                        "name" => "Combined",
                        "visibility" => "hidden",
                    },
                )?;
                edit.set(SECOND, props! { "visibility" => "hidden" })
            })
            .expect("combined sheet properties are a barrier");
        assert_eq!(combined.diagnostics().operations(), 2);
        assert_eq!(combined.diagnostics().commits(), 2);
        assert_eq!(
            sheet_property(combined.file(), FIRST, "name").as_deref(),
            Some("Combined")
        );
        assert_eq!(
            sheet_property(combined.file(), SECOND, "visibility").as_deref(),
            Some("hidden")
        );
    }
}

#[test]
fn empty_and_cell_commands_close_sheet_state_runs() {
    for eager in [false, true] {
        let source = open_owner(&fixture_with_count(6), eager);
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "visibility" => "hidden" })?;
                edit.set(SECOND, props! { "visibility" => "hidden" })?;
                edit.set(THIRD, Props::new())?;
                edit.set(FOURTH, props! { "visibility" => "hidden" })?;
                edit.set("/sheet[2]/A1", props! { "value" => "changed" })?;
                edit.set("/sheet[3]", props! { "visibility" => "hidden" })
            })
            .expect("empty and cell barriers");
        assert_eq!(commit.diagnostics().operations(), 6);
        assert_eq!(commit.diagnostics().commits(), 6);
        assert_eq!(
            sheet_property(commit.file(), FIRST, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(commit.file(), SECOND, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[2]/A1").as_deref(),
            Some("changed")
        );
        assert_eq!(
            sheet_property(commit.file(), THIRD, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(commit.file(), FOURTH, "visibility").as_deref(),
            Some("hidden")
        );
    }
}

#[test]
fn late_last_visible_failure_keeps_item_index_and_atomic_input() {
    for eager in [false, true] {
        let source = open_owner(&fixture_with_count(3), eager);
        let before = source.to_bytes().expect("publish source");
        let error = source
            .batch(visibility_batch(&[1, 2, 3], "hidden"), BatchMode::Atomic)
            .expect_err("hiding the final visible tab must fail");
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(error.item(), Some(2));
        assert_eq!(source.to_bytes().expect("publish unchanged source"), before);

        let mut malformed = visibility_batch(&[1, 2, 3], "hidden");
        malformed.push(batch_set(
            THIRD,
            Props::new().with("visibility", "not-a-state"),
        ));
        let error = source
            .batch(malformed, BatchMode::Atomic)
            .expect_err("deferred native failure precedes malformed following item");
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(error.item(), Some(2));
        assert_eq!(source.to_bytes().expect("publish unchanged source"), before);
    }
}

#[test]
fn transaction_quotas_distinguish_one_coalesced_patch_from_two_barrier_patches() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let before = source.to_bytes().expect("publish source");

        let operation_limits = SemanticLimits::builder()
            .max_operations(1)
            .expect("operation limit")
            .build()
            .expect("operation profile");
        let operation_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(operation_limits),
                |edit| {
                    edit.set(FIRST, props! { "visibility" => "hidden" })?;
                    edit.set(SECOND, props! { "visibility" => "hidden" })
                },
            )
            .expect_err("operation limit");
        assert_eq!(operation_error.code(), Code::LimitExceeded);
        assert_eq!(operation_error.item(), Some(1));
        assert_eq!(source.to_bytes().expect("source unchanged"), before);

        let target_limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("expanded-target limit")
            .build()
            .expect("expanded-target profile");
        let target_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(target_limits),
                |edit| {
                    edit.set(
                        scoped_sheet_query("sheet[name!=Tail]"),
                        props! { "visibility" => "hidden" },
                    )
                },
            )
            .expect_err("expanded target limit");
        assert_eq!(target_error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().expect("source unchanged"), before);

        let staged_limits = SemanticLimits::builder()
            .max_staged_bytes(1)
            .expect("staged-byte limit")
            .build()
            .expect("staged-byte profile");
        let staged_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(staged_limits),
                |edit| edit.set(FIRST, props! { "visibility" => "hidden" }),
            )
            .expect_err("staged-byte limit");
        assert_eq!(staged_error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().expect("source unchanged"), before);

        let coalesced = source
            .edit(|edit| {
                edit.set(SECOND, props! { "visibility" => "hidden" })?;
                edit.set(THIRD, props! { "visibility" => "hidden" })
            })
            .expect("measure coalesced snapshot");
        let budget = coalesced.diagnostics().snapshot_bytes();
        assert!(budget > 0);
        let one_patch_limits = SemanticLimits::builder()
            .max_snapshot_bytes(budget)
            .expect("one-patch snapshot limit")
            .build()
            .expect("one-patch profile");
        let accepted = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(one_patch_limits),
                |edit| {
                    edit.set(SECOND, props! { "visibility" => "hidden" })?;
                    edit.set(THIRD, props! { "visibility" => "hidden" })
                },
            )
            .expect("coalesced patch fits measured snapshot budget");
        assert_eq!(accepted.diagnostics().commits(), 1);

        let barrier_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(one_patch_limits),
                |edit| {
                    edit.set(SECOND, props! { "visibility" => "hidden" })?;
                    edit.set(THIRD, props! { "visibility" => "veryHidden" })
                },
            )
            .expect_err("two barrier patches exceed one-patch budget");
        assert_eq!(barrier_error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().expect("source unchanged"), before);
    }
}

#[test]
fn one_thousand_distinct_direct_state_commands_fit_default_limits() {
    let source = open_owner(&fixture_with_count(1_001), false);
    let commands = visibility_batch(&(2..=1_001).collect::<Vec<_>>(), "hidden");
    let commit = source
        .batch(commands, BatchMode::Atomic)
        .expect("one thousand non-active sheet commands");
    assert_eq!(commit.value().len(), 1_000);
    assert_all_applied_changed(commit.value());
    assert_eq!(commit.diagnostics().operations(), 1_000);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(commit.diagnostics().leaf_operations(), 1_000);
    assert_eq!(
        sheet_property(commit.file(), "/sheet[1]", "visibility"),
        None
    );
    assert_eq!(
        sheet_property(commit.file(), "/sheet[2]", "visibility").as_deref(),
        Some("hidden")
    );
    assert_eq!(
        sheet_property(commit.file(), "/sheet[1001]", "visibility").as_deref(),
        Some("hidden")
    );
}

#[test]
fn revealing_a_tab_replans_the_following_hide_after_the_state_barrier() {
    let bytes = fixture_with_count(2);
    for eager in [false, true] {
        let source = open_owner(&bytes, eager)
            .set(SECOND, props! { "hidden" => true })
            .expect("only first remains visible");
        let sequential = source
            .set(SECOND, props! { "hidden" => false })
            .expect("reveal second")
            .set(FIRST, props! { "hidden" => true })
            .expect("hide first after reveal");
        let commit = source
            .edit(|edit| {
                edit.set(SECOND, props! { "hidden" => false })?;
                edit.set(FIRST, props! { "hidden" => true })
            })
            .expect("replan against published reveal");
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            sequential.to_bytes().unwrap()
        );
        assert_eq!(commit.diagnostics().commits(), 2);
    }
}

#[test]
fn best_effort_failure_discards_its_sheet_run_and_later_commands_continue() {
    let bytes = fixture_with_count(3);
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let batch = source
            .batch(
                vec![
                    batch_set(SECOND, Props::new().with("hidden", true)),
                    batch_set(THIRD, Props::new().with("hidden", true)),
                    batch_set(FIRST, Props::new().with("hidden", true)),
                    batch_set(SECOND, Props::new().with("hidden", false)),
                ],
                BatchMode::BestEffort,
            )
            .expect("isolate last-visible failure");
        assert!(
            matches!(&batch.value()[2], ItemResult::Failed(error) if error.code() == Code::InvalidArgument && error.item() == Some(2))
        );
        assert!(matches!(&batch.value()[3], ItemResult::Applied(applied) if applied.changed));
        let sequential = source
            .set(SECOND, props! { "hidden" => true })
            .unwrap()
            .set(THIRD, props! { "hidden" => true })
            .unwrap()
            .set(SECOND, props! { "hidden" => false })
            .unwrap();
        assert_eq!(
            batch.file().to_bytes().unwrap(),
            sequential.to_bytes().unwrap()
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
    }
}
