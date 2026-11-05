//! XLS transaction and batch coalescing through the public operations API.
//!
//! Direct scalar cell commands are staged in one native XLS edit when their
//! targets are distinct.  This file keeps the source and eager routes side by
//! side because both routes must retain the same immutable and reversible
//! contract.
#![cfg(feature = "xls")]

use std::io::Cursor;
use std::ops::Deref;

use litchi_core::Position;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, PropValue, Props, SourceMode, props,
};
use litchi_xls::cell_values::{FormulaCache, Reference, Selector, Value};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const D1: &str = "/sheet[1]/D1";
const A2: &str = "/sheet[1]/A2";
const B2: &str = "/sheet[1]/B2";
const ARCHIVE_A1: &str = "/sheet[2]/A1";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").expect("Data worksheet");
    builder
        .write_string(data, 0, 0, "seed")
        .expect("seed string");
    builder.write_number(data, 0, 1, 2.5).expect("seed number");
    builder
        .write_boolean(data, 0, 2, true)
        .expect("seed boolean");
    builder
        .write_formula(data, 0, 3, "A1*2")
        .expect("seed formula");
    builder
        .write_string(data, 1, 0, "second row")
        .expect("second row");
    builder
        .write_number(data, 1, 1, 4.0)
        .expect("second number");

    let archive = builder.add_worksheet("Archive").expect("Archive worksheet");
    builder
        .write_string(archive, 0, 0, "untouched")
        .expect("archive string");
    let package = builder.build().expect("build XLS fixture");
    let mut edit = package.edit().expect("start rich edit");
    edit.set_value(
        "Data",
        Reference::new(0, 3).expect("formula reference"),
        Value::FormulaCache(FormulaCache::Number(6.0)),
    )
    .expect("set formula cache");
    edit.commit()
        .expect("commit formula cache")
        .into_package()
        .to_bytes()
        .expect("serialize XLS fixture")
}

fn many_cells(count: usize) -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let sheet = builder.add_worksheet("Data").expect("Data worksheet");
    for row in 0..count {
        builder
            .write_number(sheet, row as u32, 0, (row + 1) as f64)
            .expect("write numeric cell");
    }
    builder
        .build()
        .expect("build many-cell fixture")
        .to_bytes()
        .expect("serialize many-cell fixture")
}

fn formula_fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let sheet = builder.add_worksheet("Data").expect("Data worksheet");
    builder
        .write_number(sheet, 0, 0, 3.0)
        .expect("formula input");
    builder
        .write_formula(sheet, 0, 3, "A1*2")
        .expect("formula cell");
    let package = builder.build().expect("build formula fixture");
    let mut edit = package.edit().expect("formula cache edit");
    edit.set_value(
        "Data",
        Reference::new(0, 3).expect("formula coordinate"),
        Value::FormulaCache(FormulaCache::Number(6.0)),
    )
    .expect("set formula cache");
    edit.commit()
        .expect("commit formula cache")
        .into_package()
        .to_bytes()
        .expect("serialize formula fixture")
}

fn fixture_with_blank() -> Vec<u8> {
    let source = fixture();
    let package = litchi_xls::Package::from_bytes(&source).expect("open blank fixture source");
    let mut edit = package.edit().expect("edit blank fixture");
    edit.insert_cell(
        Selector::Position(Position::new(0)),
        Reference::new(1, 2).expect("blank reference"),
        Value::Blank,
    )
    .expect("insert formatting-only blank");
    edit.commit()
        .expect("commit blank fixture")
        .package()
        .to_bytes()
        .expect("serialize blank fixture")
}

fn rich_fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let sheet = builder.add_worksheet("Data").expect("Data worksheet");
    builder
        .write_number(sheet, 0, 1, 2.0)
        .expect("rich neighbor");
    let bytes = builder
        .build()
        .expect("build rich fixture")
        .to_bytes()
        .expect("serialize rich fixture");
    let package = litchi_xls::Package::from_bytes(&bytes).expect("open rich fixture");
    let mut edit = package.edit().expect("edit rich fixture");
    edit.insert_rich_text_cell(
        Selector::Position(Position::new(0)),
        Reference::new(0, 0).expect("rich reference"),
        "rich".to_owned(),
        vec![litchi_xls::SharedStringFormatRun {
            character_index: 0,
            font_index: 0,
        }],
    )
    .expect("insert rich cell");
    edit.commit()
        .expect("commit rich fixture")
        .package()
        .to_bytes()
        .expect("serialize rich fixture")
}

struct Owner {
    file: File,
    // A source-backed File retains the pathname, while this guard retains its
    // parent directory for the complete lifetime of the owner and its clones.
    _source_guard: Option<tempfile::TempDir>,
}

impl Deref for Owner {
    type Target = File;

    fn deref(&self) -> &Self::Target {
        &self.file
    }
}

fn open_owner(bytes: &[u8], eager: bool) -> Owner {
    if eager {
        return Owner {
            file: litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("open eager XLS owner"),
            _source_guard: None,
        };
    }

    let source_guard = tempfile::tempdir().expect("create source directory");
    let path = source_guard.path().join("source.xls");
    std::fs::write(&path, bytes).expect("write source");
    Owner {
        file: litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .expect("open source-backed XLS owner"),
        _source_guard: Some(source_guard),
    }
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .map(std::borrow::Cow::into_owned)
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

fn batch_remove(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

fn workbook_cell_query(expression: &'static str) -> Query<'static> {
    Query::parse(expression)
        .expect("parse workbook cell query")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn signed_fixture() -> Vec<u8> {
    let source = fixture();
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default())
            .expect("open XLS fixture");
    let mut writer = litchi_cfb::OleWriter::new();
    for path in ole.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..refs.len() {
            writer
                .create_storage(&refs[..depth])
                .expect("copy XLS storage");
        }
        let data = ole.open_stream(&refs).expect("copy XLS stream");
        writer.create_stream(&refs, &data).expect("copy XLS stream");
    }
    writer
        .create_stream(&["_signatures"], b"synthetic signature")
        .expect("add signature marker");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("serialize signed XLS");
    output.into_inner()
}

#[cfg(feature = "encryption")]
fn encrypted_fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let sheet = builder.add_worksheet("Secret").expect("Secret worksheet");
    builder
        .write_string(sheet, 0, 0, "protected")
        .expect("protected value");
    builder
        .set_password("secret", litchi_xls::EncryptionProfile::OfficeBinaryRc4)
        .expect("encryption profile");
    builder.set_encryption_entropy(
        litchi_xls::EncryptionEntropy::generate().expect("encryption entropy"),
    );
    builder
        .build()
        .expect("build encrypted XLS")
        .to_encrypted_bytes()
        .expect("serialize encrypted XLS")
}

#[test]
fn distinct_direct_sets_and_removals_share_one_native_commit() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "changed" })?;
                transaction.set(B1, props! { "value" => 7.25 })?;
                transaction.remove(C1)?;
                transaction.set(A2, props! { "value" => "second changed" })?;
                transaction.remove(B2)?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("coalesced mixed XLS edit");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 5);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("changed"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("7.25"));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some(""));
        assert_eq!(
            cell_text(commit.file(), A2).as_deref(),
            Some("second changed")
        );
        assert_eq!(cell_text(commit.file(), B2).as_deref(), Some(""));
        assert_eq!(cell_text(commit.file(), D1).as_deref(), Some("6"));
        assert_eq!(
            cell_text(commit.file(), ARCHIVE_A1).as_deref(),
            Some("untouched")
        );
        assert_eq!(original.to_bytes().expect("publish source"), before);

        assert_eq!(
            commit
                .patch()
                .apply(&original)
                .expect("replay coalesced edit")
                .to_bytes()
                .expect("publish replay"),
            commit.file().to_bytes().expect("publish target")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse coalesced edit")
                .to_bytes()
                .expect("publish inverse"),
            before
        );

        // Record 0362: a separately opened copy of the base takes the patch;
        // other content is refused.
        let foreign = open_owner(&before, eager);
        assert_eq!(
            commit
                .patch()
                .apply(&foreign)
                .expect("a byte-identical copy takes the patch")
                .to_bytes()
                .expect("publish replay"),
            commit.file().to_bytes().expect("publish commit")
        );
        assert_eq!(
            commit
                .patch()
                .apply(commit.file())
                .expect_err("other content must be refused")
                .code(),
            Code::InvalidInput
        );
    }
}

#[test]
fn one_thousand_distinct_direct_operations_do_not_escape_the_native_run() {
    let bytes = many_cells(1_000);
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish many-cell source");
        let commit = original
            .edit(|transaction| {
                for row in 1..=500 {
                    let path = format!("/sheet[1]/A{row}");
                    transaction.set(path.as_str(), props! { "value" => (row + 1_000) as f64 })?;
                }
                for row in 501..=1_000 {
                    let path = format!("/sheet[1]/A{row}");
                    transaction.remove(path.as_str())?;
                }
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("coalesced thousand-operation edit");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1_000);
        assert_eq!(commit.diagnostics().leaf_operations(), 1_000);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A1").as_deref(),
            Some("1001")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A500").as_deref(),
            Some("1500")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A501").as_deref(),
            Some("")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A1000").as_deref(),
            Some("")
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged source"),
            before
        );
    }
}

#[test]
fn repeated_coordinates_close_runs_and_preserve_last_value() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "first" })?;
                transaction.set(A1, props! { "value" => "last" })?;
                transaction.set(B1, props! { "value" => 11 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeated-coordinate edit");
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(
            commit.diagnostics().commits(),
            2,
            "a repeated coordinate is a native-run barrier"
        );
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("last"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("11"));

        let removed = original
            .edit(|transaction| {
                transaction.remove(A1)?;
                transaction.remove(A1)?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeated removal edit");
        assert_eq!(removed.diagnostics().operations(), 2);
        assert!(removed.diagnostics().commits() >= 1);
        assert_eq!(cell_text(removed.file(), A1).as_deref(), Some(""));
    }
}

#[test]
fn reads_queries_ranges_and_renames_are_native_run_barriers() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let read = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before read" })?;
                let _ = transaction.get(B1)?.view().text()?;
                transaction.set(B1, props! { "value" => 9 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("read barrier edit");
        assert_eq!(read.diagnostics().commits(), 2);
        assert_eq!(cell_text(read.file(), B1).as_deref(), Some("9"));

        let range = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before range" })?;
                transaction.set("Data!B1:C1", props! { "value" => 8 })?;
                transaction.set(A2, props! { "value" => "after range" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("range barrier edit");
        assert_eq!(range.diagnostics().commits(), 3);
        assert_eq!(cell_text(range.file(), B1).as_deref(), Some("8"));
        assert_eq!(cell_text(range.file(), C1).as_deref(), Some("8"));

        let query = workbook_cell_query("cell[value=2.5]");
        let queried = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before query" })?;
                transaction.set(query.clone(), props! { "value" => 12 })?;
                transaction.set(A2, props! { "value" => "after query" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("query barrier edit");
        assert_eq!(queried.diagnostics().commits(), 3);
        assert_eq!(cell_text(queried.file(), B1).as_deref(), Some("12"));

        let renamed = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before rename" })?;
                transaction.set("/sheet[1]", props! { "name" => "Input" })?;
                transaction.set("/sheet[1]/B1", props! { "value" => 13 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("rename barrier edit");
        assert_eq!(renamed.diagnostics().commits(), 3);
        assert_eq!(
            cell_text(renamed.file(), "/sheet[1]/B1").as_deref(),
            Some("13")
        );
        assert_eq!(
            renamed
                .file()
                .get("/sheet[1]")
                .expect("resolve renamed sheet")
                .view()
                .format()
                .expect("read renamed sheet")
                .find(|property| property.key() == "name")
                .map(|property| property.value().to_string())
                .as_deref(),
            Some("Input")
        );
    }
}

#[test]
fn atomic_and_best_effort_batches_attribute_formula_admission_failures() {
    for eager in [false, true] {
        let original = open_owner(&formula_fixture(), eager);
        let before = original.to_bytes().expect("publish formula source");
        let atomic = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "kept")),
                    batch_remove(D1),
                    batch_set(B1, Props::new().with("value", 9)),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must reject formula dependency");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::UnsupportedCapability);
        assert_eq!(
            original.to_bytes().expect("publish rollback source"),
            before
        );

        let best_effort = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "kept")),
                    batch_remove(D1),
                    batch_set(B1, Props::new().with("value", 9)),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::UnsupportedCapability
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(cell_text(best_effort.file(), A1).as_deref(), Some("kept"));
        assert_eq!(cell_text(best_effort.file(), B1).as_deref(), Some("9"));
        assert_eq!(cell_text(best_effort.file(), D1).as_deref(), Some("6"));
        assert_eq!(
            best_effort
                .patch()
                .inverse()
                .apply(best_effort.file())
                .expect("inverse best-effort batch")
                .to_bytes()
                .expect("publish inverse batch"),
            before
        );
    }
}

#[test]
fn atomic_batch_reports_each_scalar_noop_before_changed_items() {
    for eager in [false, true] {
        let original = open_owner(&fixture_with_blank(), eager);
        let before = original.to_bytes().expect("publish no-op batch source");
        let commit = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "seed")),
                    batch_set(B1, Props::new().with("value", 2.5)),
                    batch_set(C1, Props::new().with("value", true)),
                    batch_set("/sheet[1]/C2", Props::new().with("value", PropValue::Null)),
                    batch_set("/sheet[1]/Z99", Props::new().with("clear", true)),
                    batch_remove("/sheet[1]/Z100"),
                    batch_set(A1, Props::new().with("value", "changed")),
                    batch_remove(B1),
                ],
                BatchMode::Atomic,
            )
            .expect("mixed no-op and changed atomic batch");
        let changed = commit
            .value()
            .iter()
            .map(|result| match result {
                ItemResult::Applied(applied) => applied.changed,
                other => panic!("unexpected result: {other:?}"),
            })
            .collect::<Vec<_>>();
        assert_eq!(
            changed,
            [false, false, false, false, false, false, true, true]
        );
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 8);
        assert!(commit.diagnostics().commits() >= 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("changed"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some(""));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some("TRUE"));
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/C2").as_deref(),
            Some("")
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged source"),
            before
        );
    }
}

#[test]
fn deferred_rich_sst_dependency_failure_keeps_the_original_wire_index() {
    for eager in [false, true] {
        let original = open_owner(&rich_fixture(), eager);
        let before = original.to_bytes().expect("publish rich source");
        let atomic = original
            .batch(
                [
                    batch_set("/sheet[1]/B1", Props::new().with("value", 3)),
                    batch_remove("/sheet[1]/A1"),
                    batch_set("/sheet[1]/C1", Props::new().with("value", 4)),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must report the deferred rich dependency");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::UnsupportedCapability);
        assert_eq!(original.to_bytes().expect("publish rich rollback"), before);

        let best_effort = original
            .batch(
                [
                    batch_set("/sheet[1]/B1", Props::new().with("value", 3)),
                    batch_remove("/sheet[1]/A1"),
                    batch_set("/sheet[1]/C1", Props::new().with("value", 4)),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort rich dependency batch");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::UnsupportedCapability
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(cell_text(best_effort.file(), B1).as_deref(), Some("3"));
        assert_eq!(
            cell_text(best_effort.file(), "/sheet[1]/C1").as_deref(),
            Some("4")
        );
        assert_eq!(cell_text(best_effort.file(), A1).as_deref(), Some("rich"));
    }
}

#[test]
fn operation_and_target_budgets_cannot_be_bypassed_by_coalescing() {
    let original = open_owner(&fixture(), false);
    let before = original.to_bytes().expect("publish limit source");
    let operation_limits = SemanticLimits::builder()
        .max_operations(2)
        .expect("operation limit")
        .build()
        .expect("build operation limit");
    let operation_error = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(operation_limits),
            |transaction| {
                transaction.set(A1, props! { "value" => 1 })?;
                transaction.set(B1, props! { "value" => 2 })?;
                transaction.set(C1, props! { "value" => 3 })?;
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect_err("operation limit must count staged commands");
    assert_eq!(operation_error.code(), Code::LimitExceeded);
    assert_eq!(operation_error.item(), Some(2));
    assert!(matches!(
        operation_error.leaf(),
        Source::Transaction(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: 3,
            maximum: 2,
        })
    ));
    assert_eq!(original.to_bytes().expect("publish after limit"), before);

    let staged_limits = SemanticLimits::builder()
        .max_staged_bytes(1)
        .expect("staged byte limit")
        .build()
        .expect("build staged byte limit");
    let staged_error = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(staged_limits),
            |transaction| {
                transaction.set(A1, props! { "value" => "staged" })?;
                transaction.set(B1, props! { "value" => "also staged" })?;
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect_err("staged byte budget must close or refuse a native run");
    assert_eq!(staged_error.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after staged limit"),
        before
    );

    let probe = original
        .edit(|transaction| transaction.set(A1, props! { "value" => "probe" }))
        .expect("measure retained snapshot charge");
    assert!(probe.diagnostics().snapshot_bytes() > 0);
    let snapshot_limits = SemanticLimits::builder()
        .max_snapshot_bytes(probe.diagnostics().snapshot_bytes() - 1)
        .expect("snapshot byte limit")
        .build()
        .expect("build snapshot byte limit");
    let snapshot_error = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(snapshot_limits),
            |transaction| transaction.set(A1, props! { "value" => "snapshot" }),
        )
        .expect_err("retained snapshot budget must be enforced");
    assert_eq!(snapshot_error.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after snapshot limit"),
        before
    );

    let target_limits = SemanticLimits::builder()
        .max_expanded_targets(2)
        .expect("target limit")
        .build()
        .expect("build target limit");
    let target_error = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(target_limits),
            |transaction| transaction.set("Data!A1:C2", props! { "value" => "too many" }),
        )
        .expect_err("range target limit must be enforced");
    assert_eq!(target_error.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after target limit"),
        before
    );
}

#[test]
fn an_all_noop_edit_collapses_to_the_original_snapshot() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish no-op source");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "seed" })?;
                transaction.remove("/sheet[1]/Z999")?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("no-op edit");
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.file().to_bytes().expect("publish no-op result"),
            before
        );
        assert_eq!(original.to_bytes().expect("publish no-op source"), before);
    }
}

#[test]
fn stale_files_are_rejected_while_a_native_run_is_pending() {
    let directory = tempfile::tempdir().expect("stale source directory");
    let path = directory.path().join("stale.xls");
    std::fs::write(&path, fixture()).expect("write stale source");
    let original = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
        .expect("open source-backed XLS");
    let error = original
        .edit::<()>(|transaction| {
            transaction.set(A1, props! { "value" => "must not publish" })?;
            std::fs::write(&path, b"changed while the run was pending")
                .expect("change physical source");
            Ok(())
        })
        .expect_err("stale source must reject the pending run");
    assert_eq!(error.code(), Code::IoError);
    assert_eq!(
        std::fs::read(&path).expect("read changed source"),
        b"changed while the run was pending"
    );
}

/// Record 0337: items inside a coalesced run add no source check of their
/// own; the next item's resolution and leaf read still observe a source that
/// changed after the run's first item, and nothing is published.
#[test]
fn stale_files_are_rejected_inside_coalesced_set_and_remove_runs() {
    for remove in [false, true] {
        let directory = tempfile::tempdir().expect("stale source directory");
        let path = directory.path().join("stale.xls");
        std::fs::write(&path, fixture()).expect("write stale source");
        let original =
            litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                .expect("open source-backed XLS");
        let error = original
            .edit::<()>(|transaction| {
                if remove {
                    transaction.remove(A1)?;
                } else {
                    transaction.set(A1, props! { "value" => 7 })?;
                }
                std::fs::write(&path, b"changed inside the run").expect("change physical source");
                if remove {
                    transaction.remove(A2)?;
                } else {
                    transaction.set(A2, props! { "value" => 8 })?;
                }
                Ok(())
            })
            .expect_err("a source changed inside the run must be refused");
        assert_eq!(error.code(), Code::IoError, "remove={remove}");
        assert_eq!(
            std::fs::read(&path).expect("read changed source"),
            b"changed inside the run"
        );
    }
}

#[test]
fn signed_and_encrypted_security_wrappers_survive_batch_admission() {
    let signed_bytes = signed_fixture();
    let signed = open_owner(&signed_bytes, false);
    assert!(signed.is_signed());
    assert_eq!(
        signed
            .batch(
                [batch_set(A1, Props::new().with("value", "blocked"))],
                BatchMode::Atomic,
            )
            .expect_err("signed batch must be refused")
            .code(),
        Code::DocumentProtected
    );
    assert_eq!(
        signed.to_bytes().expect("publish signed source"),
        signed_bytes
    );

    #[cfg(feature = "encryption")]
    {
        let locked = litchi_ops::open_bytes(encrypted_fixture()).expect("open encrypted XLS");
        let File::Locked(locked) = locked else {
            panic!("encrypted XLS must remain locked");
        };
        let unlocked = locked
            .unlock(litchi_ops::Password::new("secret".to_owned()))
            .expect("unlock encrypted XLS");
        let error = unlocked
            .batch(
                [batch_set(
                    "/Secret/A1",
                    Props::new().with("value", "blocked"),
                )],
                BatchMode::Atomic,
            )
            .expect_err("unlocked encrypted batch must refuse plaintext editing");
        assert_eq!(error.code(), Code::DocumentProtected);
    }
}

#[test]
fn native_staging_limit_closes_runs_without_limiting_the_public_batch() {
    const COUNT: usize = 4_100;
    let bytes = many_cells(COUNT);
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for remove in [false, true] {
            let commit = original
                .edit(|transaction| {
                    for row in 1..=COUNT {
                        let path = format!("/sheet[1]/A{row}");
                        if remove {
                            transaction.remove(path.as_str())?;
                        } else {
                            transaction.set(path.as_str(), props! { "value" => -1.0 })?;
                        }
                    }
                    Ok::<_, litchi_ops::Error>(())
                })
                .expect("bounded native runs cover the complete public batch");
            assert_eq!(commit.diagnostics().operations(), COUNT);
            assert_eq!(commit.diagnostics().commits(), 2);
            let expected = if remove { "" } else { "-1" };
            assert_eq!(cell_text(commit.file(), A1).as_deref(), Some(expected));
            assert_eq!(
                cell_text(commit.file(), "/sheet[1]/A4100").as_deref(),
                Some(expected)
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes
            );
            assert_eq!(original.to_bytes().unwrap(), bytes);
        }
    }
}

#[test]
fn mutation_budget_does_not_count_native_string_resource_operations() {
    let bytes = many_cells(3);
    let limits = SemanticLimits::builder()
        .max_operations(2)
        .unwrap()
        .build()
        .unwrap();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set("/sheet[1]/A1", props! { "value" => "first new string" })?;
                    transaction.set("/sheet[1]/A2", props! { "value" => "second new string" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect("two calls may stage multiple native resource operations");
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            cell_text(commit.file(), A1).as_deref(),
            Some("first new string")
        );
        assert_eq!(
            cell_text(commit.file(), A2).as_deref(),
            Some("second new string")
        );
    }
}
