//! Unified XLS cell and worksheet writes through the public [`File`] API.
//!
//! The fixture is authored by the native XLS builder and each behavior is run
//! through both the retained positional owner and the eager byte owner.  The
//! tests deliberately keep the input handle immutable: a successful operation
//! produces a new handle, while failed operations leave the input publishable
//! byte-for-byte.
#![cfg(feature = "xls")]

use std::io::Cursor;
use std::ops::Deref;

use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, PropValue, Props, SourceMode, props,
};
use litchi_xls::cell_values::{FormulaCache, Reference, Selector, Value as XlsValue};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const A2: &str = "/sheet[1]/A2";
const B2: &str = "/sheet[1]/B2";
const SECOND_A1: &str = "/sheet[2]/A1";
const FORMULA_B1: &str = "/sheet[1]/B1";

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
        .write_string(data, 1, 0, "second row")
        .expect("second row");
    builder
        .write_number(data, 1, 1, 4.0)
        .expect("second number");

    let archive = builder.add_worksheet("Archive").expect("Archive worksheet");
    builder
        .write_string(archive, 0, 0, "untouched")
        .expect("archive string");
    builder
        .write_number(archive, 0, 1, 9.0)
        .expect("archive number");
    builder
        .build()
        .expect("build XLS fixture")
        .to_bytes()
        .expect("serialize XLS fixture")
}

fn formula_fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").expect("Data worksheet");
    builder
        .write_number(data, 0, 0, 3.0)
        .expect("formula input");
    builder
        .write_formula(data, 0, 1, "A1*2")
        .expect("formula cell");
    let package = builder.build().expect("build formula fixture");
    let mut edit = package.edit().expect("start formula cache edit");
    edit.set_value(
        "Data",
        Reference::new(0, 1).expect("formula reference"),
        XlsValue::FormulaCache(FormulaCache::Number(6.0)),
    )
    .expect("set formula cache");
    edit.commit()
        .expect("commit formula cache")
        .into_package()
        .to_bytes()
        .expect("serialize formula fixture")
}

fn styled_fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").expect("Data worksheet");
    let format = builder
        .register_number_format("0.00\"kg\"")
        .expect("custom number format");
    let format = builder
        .add_cell_format(litchi_xls::writer::formatting::ExtendedFormat {
            format_index: format,
            ..Default::default()
        })
        .expect("custom cell format");
    builder
        .write_number_with_format(data, 0, 0, 12.5, format)
        .expect("styled number");
    builder
        .build()
        .expect("build styled fixture")
        .to_bytes()
        .expect("serialize styled fixture")
}

fn native_cell_style(bytes: &[u8], row: u32, column: u32) -> Option<u16> {
    let snapshot = litchi_xls::Package::from_bytes(bytes)
        .and_then(|package| package.cell_values())
        .expect("open native XLS snapshot");
    snapshot
        .worksheet(Selector::Position(Position::new(0)))
        .expect("resolve native worksheet")
        .and_then(|worksheet| {
            worksheet
                .cell(Reference::new(row, column).expect("native cell reference"))
                .expect("resolve native cell")
                .map(|cell| cell.style().get())
        })
}

fn native_cell_value(bytes: &[u8], row: u32, column: u32) -> Option<XlsValue> {
    let snapshot = litchi_xls::Package::from_bytes(bytes)
        .and_then(|package| package.cell_values())
        .expect("open native XLS snapshot");
    snapshot
        .worksheet(Selector::Position(Position::new(0)))
        .expect("resolve native worksheet")
        .and_then(|worksheet| {
            worksheet
                .cell(Reference::new(row, column).expect("native cell reference"))
                .expect("resolve native cell")
                .map(|cell| cell.value().clone())
        })
}

struct Owner {
    file: File,
    source_guard: Option<tempfile::TempDir>,
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
            source_guard: None,
        };
    }

    let source_guard = tempfile::tempdir().expect("create XLS source directory");
    let path = source_guard.path().join("source.xls");
    std::fs::write(&path, bytes).expect("write XLS source");
    Owner {
        file: litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .expect("open source-backed XLS owner"),
        source_guard: Some(source_guard),
    }
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

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

fn workbook_cell_query() -> Query<'static> {
    Query::parse("cell")
        .expect("parse cell selector")
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
        .expect("add XLS signature marker");
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
fn scalar_values_round_trip_for_source_and_eager_owners() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        let updated = original
            .set(A1, props! { "value" => "changed" })
            .unwrap_or_else(|error| {
                let source = std::error::Error::source(&error);
                panic!("set text value: {error}; source={source:?}");
            })
            .set(B1, props! { "value" => 7.25 })
            .expect("set numeric value")
            .set(C1, props! { "value" => false })
            .expect("set boolean value");

        assert_eq!(cell_text(&updated, A1).as_deref(), Some("changed"));
        assert_eq!(cell_text(&updated, B1).as_deref(), Some("7.25"));
        assert_eq!(cell_text(&updated, C1).as_deref(), Some("FALSE"));
        assert_eq!(
            original.to_bytes().expect("publish detached original"),
            original_bytes,
            "the {eager:?} input owner changed"
        );
    }
}

#[test]
fn created_workbook_accepts_its_first_text_cell() {
    let original = litchi_ops::create(FileFormat::Xls).expect("create XLS");
    let before = original.to_bytes().expect("publish empty workbook");
    let updated = original
        .set(A1, props! { "value" => "first text" })
        .expect("insert first text");
    assert_eq!(cell_text(&updated, A1).as_deref(), Some("first text"));
    assert_eq!(
        original.to_bytes().expect("publish unchanged source"),
        before
    );
    let reopened = open_owner(&updated.to_bytes().expect("publish text workbook"), true);
    assert_eq!(cell_text(&reopened, A1).as_deref(), Some("first text"));
}

#[test]
fn missing_cells_insert_and_clear_for_source_and_eager_owners() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let inserted = original
            .set("/sheet[1]/D4", props! { "value" => 11 })
            .expect("insert missing cell");
        assert_eq!(cell_text(&inserted, "/sheet[1]/D4").as_deref(), Some("11"));
        let cleared = inserted
            .set("/sheet[1]/D4", props! { "clear" => true })
            .expect("clear inserted cell");
        assert_eq!(cell_text(&cleared, "/sheet[1]/D4").as_deref(), Some(""));
        assert_eq!(
            cell_property(&cleared, "/sheet[1]/D4", "empty").as_deref(),
            Some("true")
        );
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn null_value_and_clear_are_explicit_blank_operations() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let cleared = original
            .set(A1, Props::new().with("value", PropValue::Null))
            .expect("null clears text")
            .set(B1, props! { "clear" => true })
            .expect("clear clears number");
        for path in [A1, B1] {
            assert_eq!(cell_text(&cleared, path).as_deref(), Some(""), "{path}");
            assert_eq!(
                cell_property(&cleared, path, "empty").as_deref(),
                Some("true"),
                "{path} blank property"
            );
        }
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn exact_cell_noops_keep_original_bytes() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        assert_eq!(
            original
                .set(A1, props! { "value" => "seed" })
                .expect("repeat string")
                .to_bytes()
                .expect("publish string no-op"),
            before
        );
        assert_eq!(
            original
                .set(B1, props! { "value" => 2.5 })
                .unwrap_or_else(|error| {
                    let source = std::error::Error::source(&error);
                    panic!("repeat number: {error}; source={source:?}");
                })
                .to_bytes()
                .expect("publish number no-op"),
            before
        );
    }
}

#[test]
fn empty_cell_props_are_true_noops_for_both_owner_routes() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");

        let one_shot = original
            .set(A1, Props::new())
            .expect("empty cell props are a no-op");
        assert_eq!(one_shot.to_bytes().expect("publish one-shot no-op"), before);

        let commit = original
            .edit(|transaction| {
                transaction.set(A1, Props::new())?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("empty cell props transaction");
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.file().to_bytes().expect("publish transaction no-op"),
            before
        );
        assert_eq!(original.to_bytes().expect("publish detached owner"), before);
    }
}

#[test]
fn formula_scalar_set_and_clear_are_refused_without_losing_cached_formula() {
    for eager in [false, true] {
        let original = open_owner(&formula_fixture(), eager);
        let before = original.to_bytes().expect("publish formula owner");
        let cached_text = cell_text(&original, FORMULA_B1);
        let cached_type = cell_property(&original, FORMULA_B1, "type");

        for (key, properties) in [
            ("value", Props::new().with("value", "replacement")),
            ("clear", Props::new().with("clear", true)),
        ] {
            let error = original
                .set(FORMULA_B1, properties)
                .expect_err("formula scalar mutation must be refused");
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert_eq!(error.property(), Some(key));
            assert_eq!(
                original.to_bytes().expect("publish refused formula owner"),
                before
            );
            assert_eq!(cell_text(&original, FORMULA_B1), cached_text);
            assert_eq!(cell_property(&original, FORMULA_B1, "type"), cached_type);
        }
    }
}

#[test]
fn scalar_type_replacements_retain_the_native_cell_style() {
    for eager in [false, true] {
        let original = open_owner(&styled_fixture(), eager);
        let before = original.to_bytes().expect("publish styled owner");
        let style = native_cell_style(&before, 0, 0).expect("styled source cell");

        let number = original
            .set(A1, Props::new().with("value", PropValue::Number(17.25)))
            .expect("number replacement");
        let number_bytes = number.to_bytes().expect("publish number replacement");
        assert_eq!(native_cell_style(&number_bytes, 0, 0), Some(style));

        let text = number
            .set(A1, Props::new().with("value", "styled text"))
            .unwrap_or_else(|error| {
                panic!(
                    "text replacement: {error}; source={:?}",
                    std::error::Error::source(&error)
                )
            });
        let text_bytes = text.to_bytes().expect("publish text replacement");
        assert_eq!(native_cell_style(&text_bytes, 0, 0), Some(style));

        let boolean = text
            .set(A1, Props::new().with("value", true))
            .expect("boolean replacement");
        let boolean_bytes = boolean.to_bytes().expect("publish boolean replacement");
        assert_eq!(native_cell_style(&boolean_bytes, 0, 0), Some(style));

        let blank = boolean
            .set(A1, Props::new().with("clear", true))
            .expect("blank replacement");
        let blank_bytes = blank.to_bytes().expect("publish blank replacement");
        assert_eq!(native_cell_style(&blank_bytes, 0, 0), Some(style));
        assert_eq!(cell_text(&blank, A1).as_deref(), Some(""));
        assert_eq!(cell_property(&blank, A1, "empty").as_deref(), Some("true"));
        assert_eq!(
            original.to_bytes().expect("publish detached styled owner"),
            before
        );
    }
}

#[test]
fn exactly_representable_unsigned_scalars_are_written_as_numbers() {
    let values = [
        PropValue::Unsigned((1_u64 << 53) - 1),
        PropValue::Unsigned(1_u64 << 53),
        PropValue::Unsigned(1_u64 << 63),
    ];

    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let mut updated = original.file.clone();
        for value in values.iter().cloned() {
            let expected = match value {
                PropValue::Unsigned(value) => value as f64,
                _ => unreachable!("unsigned regression fixture only"),
            };
            updated = updated
                .set(A1, Props::new().with("value", value))
                .expect("exact unsigned scalar");
            match native_cell_value(&updated.to_bytes().expect("publish unsigned scalar"), 0, 0) {
                Some(XlsValue::Number(observed)) => assert_eq!(observed, expected),
                other => panic!("expected native Number({expected}), got {other:?}"),
            }
            assert_eq!(
                cell_property(&updated, A1, "type").as_deref(),
                Some("Number")
            );
        }
    }
}

#[test]
fn worksheet_name_writes_and_queries_preserve_other_tabs() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let renamed = original
            .set("/sheet[1]", props! { "name" => "Input" })
            .expect("rename worksheet");
        assert_eq!(
            sheet_property(&renamed, "/sheet[1]", "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(&renamed, "/sheet[2]", "name").as_deref(),
            Some("Archive")
        );
        let matches = renamed
            .query("sheet[name=Input]")
            .expect("query renamed worksheet")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("collect worksheet query");
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].path().to_string(), "/sheet[1]");
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn ranges_cover_existing_and_missing_cells_on_one_worksheet() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let updated = original
            .set("Data!A1:C2", props! { "value" => "range" })
            .expect("set worksheet range");
        for path in [A1, B1, C1, A2, B2, "/sheet[1]/C2"] {
            assert_eq!(
                cell_text(&updated, path).as_deref(),
                Some("range"),
                "{path}"
            );
        }
        assert_eq!(cell_text(&updated, SECOND_A1).as_deref(), Some("untouched"));
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn workbook_query_fanout_edits_all_stored_cells_atomically() {
    let selector = workbook_cell_query();
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let matches = original
            .query(selector.clone())
            .expect("compile workbook query")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("walk workbook cells");
        assert_eq!(matches.len(), 7);

        let commit = original
            .edit(|transaction| transaction.set(selector.clone(), props! { "value" => "query" }))
            .expect("edit query matches");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("query"));
        assert_eq!(
            cell_text(commit.file(), SECOND_A1).as_deref(),
            Some("query")
        );
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn transactions_replay_inverse_and_accept_reopened_copies() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "staged" })?;
                transaction.set(SECOND_A1, props! { "value" => 13 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit XLS transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(
            commit
                .patch()
                .apply(&original)
                .expect("replay patch")
                .to_bytes()
                .expect("publish replay"),
            commit.file().to_bytes().expect("publish commit")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse patch")
                .to_bytes()
                .expect("publish inverse"),
            before
        );
        let reopened = open_owner(&before, eager);
        assert_eq!(
            commit
                .patch()
                .apply(&reopened)
                .expect("a byte-identical copy takes the patch")
                .to_bytes()
                .expect("publish reopened replay"),
            commit.file().to_bytes().expect("publish committed file")
        );
    }
}

#[test]
fn atomic_and_best_effort_batches_keep_item_attribution_and_rollback() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        let atomic = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "atomic")),
                    batch_set(A2, Props::new().with("unsupported", true)),
                    batch_set(SECOND_A1, Props::new().with("value", "suffix")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must roll back unsupported property");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::UnsupportedProperty);
        assert_eq!(original.to_bytes().expect("publish rollback"), before);

        let best_effort = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "best")),
                    batch_set(A2, Props::new().with("unsupported", true)),
                    batch_set(SECOND_A1, Props::new().with("value", "suffix")),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch");
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
fn target_staged_and_snapshot_limits_fail_atomically() {
    let original = open_owner(&fixture(), false);
    let before = original.to_bytes().expect("publish original owner");

    let expanded_limits = SemanticLimits::builder()
        .max_expanded_targets(2)
        .expect("expanded target bound")
        .build()
        .expect("expanded target profile");
    let expanded = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(expanded_limits),
            |transaction| transaction.set("Data!A1:C2", props! { "value" => "too many" }),
        )
        .expect_err("range exceeds expanded-target limit");
    assert_eq!(expanded.code(), Code::LimitExceeded);
    assert_eq!(
        original
            .to_bytes()
            .expect("publish after expansion refusal"),
        before
    );

    let staged_limits = SemanticLimits::builder()
        .max_staged_bytes(1)
        .expect("staged byte bound")
        .build()
        .expect("staged byte profile");
    let staged = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(staged_limits),
            |transaction| transaction.set(A1, props! { "value" => "staged" }),
        )
        .expect_err("staged payload exceeds bound");
    assert_eq!(staged.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after staged refusal"),
        before
    );

    let probe = original
        .edit(|transaction| transaction.set(A1, props! { "value" => "snapshot" }))
        .expect("measure snapshot charge");
    let charge = probe.diagnostics().snapshot_bytes();
    assert!(charge > 0);
    let snapshot_limits = SemanticLimits::builder()
        .max_snapshot_bytes(charge - 1)
        .expect("snapshot byte bound")
        .build()
        .expect("snapshot byte profile");
    let snapshot = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(snapshot_limits),
            |transaction| transaction.set(A1, props! { "value" => "snapshot" }),
        )
        .expect_err("snapshot charge exceeds bound");
    assert_eq!(snapshot.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after snapshot refusal"),
        before
    );
}

#[test]
fn saved_cell_write_reopens_and_preserves_origin_for_source_and_eager_routes() {
    let bytes = fixture();
    let directory = tempfile::tempdir().expect("save directory");
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(A1, props! { "value" => "saved" })
            .expect("set saved value");
        let direct = updated.to_bytes().expect("publish updated XLS");
        let reopened = open_owner(&direct, eager);
        assert_eq!(cell_text(&reopened, A1).as_deref(), Some("saved"));

        let destination = directory.path().join(format!("saved-{eager}.xls"));
        updated.save_as(&destination).expect("save updated XLS");
        assert_eq!(std::fs::read(&destination).expect("read saved XLS"), direct);
        assert_eq!(
            cell_text(
                &litchi_ops::open(&destination).expect("reopen saved XLS"),
                A1
            )
            .as_deref(),
            Some("saved")
        );

        let origin = directory.path().join(format!("origin-{eager}.xls"));
        std::fs::write(&origin, &bytes).expect("write origin XLS");
        let path_owner = litchi_ops::open_with(
            &origin,
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .expect("open origin XLS");
        path_owner
            .set(A1, props! { "value" => "origin saved" })
            .expect("edit origin XLS")
            .save()
            .expect("save origin XLS");
        assert_eq!(
            cell_text(&litchi_ops::open(&origin).expect("reopen origin XLS"), A1).as_deref(),
            Some("origin saved")
        );
    }
}

#[test]
fn stale_source_is_refused_without_clobbering_destination() {
    let directory = tempfile::tempdir().expect("stale directory");
    let origin = directory.path().join("stale.xls");
    std::fs::write(&origin, fixture()).expect("write stale origin");
    let stale = litchi_ops::open(&origin).expect("open source-backed XLS");
    std::fs::write(&origin, b"source changed after open").expect("change source");

    assert_eq!(
        stale
            .set(A1, props! { "value" => "must refuse" })
            .expect_err("stale XLS edit")
            .code(),
        Code::IoError
    );
    let destination = directory.path().join("destination.xls");
    std::fs::write(&destination, b"keep destination").expect("write destination");
    assert_eq!(
        stale
            .save_as(&destination)
            .expect_err("stale XLS save")
            .code(),
        Code::IoError
    );
    assert_eq!(
        std::fs::read(&destination).expect("read destination"),
        b"keep destination"
    );
}

#[test]
fn signed_xls_refuses_writes_until_explicit_unsign() {
    let bytes = signed_fixture();
    let signed = open_owner(&bytes, false);
    assert!(signed.is_signed());
    assert_eq!(signed.to_bytes().expect("publish signed XLS"), bytes);
    assert_eq!(
        signed
            .set(A1, props! { "value" => "must refuse" })
            .expect_err("signed XLS edit")
            .code(),
        Code::DocumentProtected
    );

    let Owner {
        file: signed_file,
        source_guard: _source_guard,
    } = signed;
    let File::Signed(wrapper) = signed_file else {
        panic!("synthetic XLS signature must open as File::Signed");
    };
    let unsigned = wrapper.unsign();
    assert!(!unsigned.is_signed());
    let changed = unsigned
        .set(A1, props! { "value" => "unsigned" })
        .expect("edit explicitly unsigned XLS");
    let reopened = litchi_ops::open_bytes(changed.to_bytes().expect("publish unsigned XLS"))
        .expect("reopen unsigned XLS");
    assert!(!reopened.is_signed());
    assert_eq!(cell_text(&reopened, A1).as_deref(), Some("unsigned"));
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_xls_refuses_plain_writes_and_publication() {
    let locked = litchi_ops::open_bytes(encrypted_fixture()).expect("open encrypted XLS");
    let File::Locked(locked) = locked else {
        panic!("encrypted XLS must remain locked");
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .expect("unlock encrypted XLS");
    assert_eq!(
        cell_text(&unlocked, "/Secret/A1").as_deref(),
        Some("protected")
    );
    assert_eq!(
        unlocked
            .set("/Secret/A1", props! { "value" => "must refuse" })
            .expect_err("encrypted XLS edit")
            .code(),
        Code::DocumentProtected
    );
    assert_eq!(
        unlocked
            .to_bytes()
            .expect_err("encrypted publication")
            .code(),
        Code::DocumentProtected
    );
}

#[test]
fn invalid_cell_properties_are_atomic() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original owner");
        for properties in [
            Props::new().with("unsupported", true),
            Props::new().with("clear", false),
            Props::new().with("value", PropValue::Number(f64::NAN)),
            Props::new().with("value", "bad\u{1}text"),
        ] {
            let error = original
                .set(A1, properties)
                .expect_err("invalid XLS cell properties");
            assert!(matches!(
                error.code(),
                Code::UnsupportedProperty | Code::InvalidValue
            ));
            assert_eq!(
                original.to_bytes().expect("publish unchanged owner"),
                before
            );
        }
    }
}

#[test]
fn xls_creation_and_write_capability_are_advertised() {
    let capabilities = litchi_ops::capabilities::Capabilities::of(FileFormat::Xls);
    assert!(
        capabilities
            .verb(litchi_ops::capabilities::Verb::Set)
            .is_supported()
    );
    assert!(
        capabilities
            .verb(litchi_ops::capabilities::Verb::Batch)
            .is_supported()
    );
    assert!(
        capabilities
            .verb(litchi_ops::capabilities::Verb::Remove)
            .is_supported()
    );
    assert!(
        capabilities
            .verb(litchi_ops::capabilities::Verb::Save)
            .is_supported()
    );
    let cell = capabilities.element("cell").expect("cell capability");
    assert!(cell.property("value").expect("value property").set());
    assert!(cell.property("clear").expect("clear property").set());
    assert!(cell.operations().remove);
}
