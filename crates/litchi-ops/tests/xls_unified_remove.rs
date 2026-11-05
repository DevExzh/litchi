//! Unified XLS cell removal through the public [`File`] and transaction APIs.
//!
//! Legacy XLS removal is sparse: deleting a stored record leaves every other
//! coordinate and its formatting in place, while a virtual coordinate is an
//! exact no-op.  The same contract is exercised through retained source and
//! eager byte owners.
#![cfg(feature = "xls")]

use std::io::Cursor;
use std::ops::Deref;

use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, SourceMode};
use litchi_xls::cell_values::{Reference, Selector, Value as XlsValue};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const D1: &str = "/sheet[1]/D1";
const A2: &str = "/sheet[1]/A2";
const B2: &str = "/sheet[1]/B2";
const ARCHIVE_A1: &str = "/sheet[2]/A1";
const ARCHIVE_B1: &str = "/sheet[2]/B1";

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
        .expect("formula cell");
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
    let package = builder.build().expect("build XLS fixture");
    let mut edit = package.edit().expect("start formula cache edit");
    edit.set_value(
        "Data",
        Reference::new(0, 3).expect("formula reference"),
        XlsValue::FormulaCache(litchi_xls::cell_values::FormulaCache::Number(6.0)),
    )
    .expect("set formula cache");
    edit.commit()
        .expect("commit formula cache")
        .into_package()
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
        XlsValue::FormulaCache(litchi_xls::cell_values::FormulaCache::Number(6.0)),
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
        .expect("styled first cell");
    builder
        .write_number_with_format(data, 0, 1, 13.5, format)
        .expect("styled neighboring cell");
    builder
        .build()
        .expect("build styled fixture")
        .to_bytes()
        .expect("serialize styled fixture")
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

fn native_cell_value(bytes: &[u8], sheet: usize, row: u32, column: u32) -> Option<XlsValue> {
    let snapshot = litchi_xls::Package::from_bytes(bytes)
        .and_then(|package| package.cell_values())
        .expect("open native XLS snapshot");
    snapshot
        .worksheet(Selector::Position(Position::new(sheet)))
        .expect("resolve native worksheet")
        .and_then(|worksheet| {
            worksheet
                .cell(Reference::new(row, column).expect("native cell reference"))
                .expect("resolve native cell")
                .map(|cell| cell.value().clone())
        })
}

fn native_cell_style(bytes: &[u8], sheet: usize, row: u32, column: u32) -> Option<u16> {
    let snapshot = litchi_xls::Package::from_bytes(bytes)
        .and_then(|package| package.cell_values())
        .expect("open native XLS snapshot");
    snapshot
        .worksheet(Selector::Position(Position::new(sheet)))
        .expect("resolve native worksheet")
        .and_then(|worksheet| {
            worksheet
                .cell(Reference::new(row, column).expect("native cell reference"))
                .expect("resolve native cell")
                .map(|cell| cell.style().get())
        })
}

fn workbook_cell_query(selector: &'static str) -> Query<'static> {
    Query::parse(selector)
        .expect("parse workbook cell selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_remove(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
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
fn stored_cells_are_physically_removed_without_shifting_neighbors() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLS");
        let removed = original.remove(B1).expect("remove numeric cell");

        assert_eq!(cell_text(&removed, B1).as_deref(), Some(""));
        assert!(cell_property(&removed, B1, "empty").is_none());
        assert_eq!(cell_text(&removed, A1).as_deref(), Some("seed"));
        assert_eq!(cell_text(&removed, C1).as_deref(), Some("TRUE"));
        assert_eq!(cell_text(&removed, A2).as_deref(), Some("second row"));
        assert_eq!(cell_text(&removed, B2).as_deref(), Some("4"));

        let after = removed.to_bytes().expect("publish removed XLS");
        assert!(native_cell_value(&after, 0, 0, 1).is_none());
        assert_eq!(
            native_cell_value(&after, 0, 0, 2),
            Some(XlsValue::Boolean(true))
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            before
        );
    }
}

#[test]
fn neighboring_styles_survive_physical_removal() {
    for eager in [false, true] {
        let original = open_owner(&styled_fixture(), eager);
        let before = original.to_bytes().expect("publish styled source");
        let neighbor_style = native_cell_style(&before, 0, 0, 1).expect("neighbor style");
        let removed = original.remove(A1).expect("remove styled cell");
        let after = removed.to_bytes().expect("publish styled removal");

        assert!(native_cell_value(&after, 0, 0, 0).is_none());
        assert_eq!(native_cell_style(&after, 0, 0, 1), Some(neighbor_style));
        assert_eq!(cell_text(&removed, B1).as_deref(), Some("13.5"));
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn ranges_remove_existing_cells_and_include_virtual_holes() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original XLS");
        let removed = original
            .remove("Data!A1:C2")
            .expect("remove worksheet range");

        for path in [A1, B1, C1, A2, B2, "/sheet[1]/C2"] {
            assert_eq!(cell_text(&removed, path).as_deref(), Some(""), "{path}");
        }
        assert!(cell_property(&removed, A1, "empty").is_none());
        assert_eq!(cell_text(&removed, D1).as_deref(), Some("6"));
        assert_eq!(
            cell_text(&removed, ARCHIVE_A1).as_deref(),
            Some("untouched")
        );
        assert_eq!(cell_text(&removed, ARCHIVE_B1).as_deref(), Some("9"));
        assert_eq!(original.to_bytes().expect("publish original"), before);
    }
}

#[test]
fn virtual_paths_and_empty_queries_are_exact_noops() {
    let selector = workbook_cell_query("cell[value=missing]");
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish original XLS");
        for target in ["/sheet[1]/Z99", "Data!Y99:Z100"] {
            assert_eq!(
                original
                    .remove(target)
                    .expect("virtual removal no-op")
                    .to_bytes()
                    .expect("publish virtual no-op"),
                before,
                "{target}"
            );
        }
        assert_eq!(
            original
                .remove(selector.clone())
                .expect("empty query no-op")
                .to_bytes()
                .expect("publish empty query no-op"),
            before
        );
    }
}

#[test]
fn duplicate_query_matches_are_removed_once_across_worksheets() {
    let selector = workbook_cell_query("cell[value=seed], cell[value=seed], cell[value=untouched]");
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let commit = original
            .edit(|transaction| transaction.remove(selector.clone()))
            .expect("remove duplicate query matches");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some(""));
        assert_eq!(cell_text(commit.file(), ARCHIVE_A1).as_deref(), Some(""));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("2.5"));
        assert_eq!(cell_text(commit.file(), ARCHIVE_B1).as_deref(), Some("9"));
    }
}

#[test]
fn formula_cells_are_refused_and_other_targets_remain_untouched() {
    for eager in [false, true] {
        let original = open_owner(&formula_fixture(), eager);
        let before = original.to_bytes().expect("publish formula source");
        let cached_text = cell_text(&original, "/sheet[1]/B1");
        let error = original
            .remove("/sheet[1]/B1")
            .expect_err("pure formula removal must be refused");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(original.to_bytes().expect("publish refused source"), before);
        assert_eq!(cell_text(&original, "/sheet[1]/B1"), cached_text);
    }
}

#[test]
fn occupied_row_and_column_removals_are_granular_capability_errors() {
    for target in ["/sheet[1]/row[1]", "/sheet[1]/col[1]"] {
        let original = open_owner(&fixture(), false);
        let before = original.to_bytes().expect("publish original");
        let error = original
            .remove(target)
            .expect_err("unsupported XLS removal target");
        assert_eq!(
            error.code(),
            Code::UnsupportedCapability,
            "{target}: {error}"
        );
        assert_eq!(original.to_bytes().expect("publish unchanged"), before);
    }
}

#[test]
fn create_set_remove_save_and_reopen_keeps_the_cell_absent() {
    let original = litchi_ops::create(FileFormat::Xls).expect("create XLS");
    let inserted = original
        .set(A1, litchi_ops::props! { "value" => "created" })
        .expect("set created cell");
    let removed = inserted.remove(A1).expect("remove created cell");
    assert_eq!(cell_text(&removed, A1).as_deref(), Some(""));
    assert!(cell_property(&removed, A1, "empty").is_none());

    let directory = tempfile::tempdir().expect("save directory");
    let path = directory.path().join("removed.xls");
    removed.save_as(&path).expect("save removed XLS");
    let reopened = litchi_ops::open(&path).expect("reopen removed XLS");
    assert_eq!(cell_text(&reopened, A1).as_deref(), Some(""));
    assert!(cell_property(&reopened, A1, "empty").is_none());
    let published = reopened.to_bytes().expect("publish reopened XLS");
    assert!(native_cell_value(&published, 0, 0, 0).is_none());
}

#[test]
fn mixed_set_remove_transactions_replay_inverse_and_accept_reopened_copies() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish transaction source");
        let commit = original
            .edit(|transaction| {
                transaction.remove(A1)?;
                transaction.set(B1, litchi_ops::props! { "value" => 8.25 })?;
                transaction.remove(C1)?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mixed XLS transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some(""));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("8.25"));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some(""));

        assert_eq!(
            commit
                .patch()
                .apply(&original)
                .expect("replay mixed patch")
                .to_bytes()
                .expect("publish replay"),
            commit.file().to_bytes().expect("publish committed file")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse mixed patch")
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
fn removal_batches_are_atomic_or_attribute_failures_per_item() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let before = original.to_bytes().expect("publish batch source");
        let atomic = original
            .batch(
                [
                    batch_remove(A1),
                    batch_remove("/sheet[99]/A1"),
                    batch_remove(ARCHIVE_A1),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic removal batch must roll back");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::PathNotFound);
        assert_eq!(
            original.to_bytes().expect("publish rollback source"),
            before
        );

        let best_effort = original
            .batch(
                [
                    batch_remove(A1),
                    batch_remove("/sheet[99]/A1"),
                    batch_remove(ARCHIVE_A1),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort removal batch");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(result) if result.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error) if error.item() == Some(1) && error.code() == Code::PathNotFound
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(result) if result.changed
        ));
        assert_eq!(cell_text(best_effort.file(), A1).as_deref(), Some(""));
        assert_eq!(
            cell_text(best_effort.file(), ARCHIVE_A1).as_deref(),
            Some("")
        );
        assert_eq!(
            best_effort
                .patch()
                .inverse()
                .apply(best_effort.file())
                .expect("inverse removal batch")
                .to_bytes()
                .expect("publish inverse batch"),
            before
        );
    }
}

#[test]
fn removal_limits_cover_target_staged_and_retained_snapshot_charges() {
    let original = open_owner(&fixture(), false);
    let before = original.to_bytes().expect("publish limit source");

    let expanded_limits = SemanticLimits::builder()
        .max_expanded_targets(2)
        .expect("expanded target limit")
        .build()
        .expect("build expanded limit");
    let expanded = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(expanded_limits),
            |transaction| transaction.remove("Data!A1:C2"),
        )
        .expect_err("range removal exceeds target limit");
    assert_eq!(expanded.code(), Code::LimitExceeded);
    assert_eq!(expanded.item(), Some(0));
    assert!(matches!(
        expanded.leaf(),
        Source::Transaction(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: 6,
            maximum: 2,
        })
    ));
    assert_eq!(
        original.to_bytes().expect("publish after target refusal"),
        before
    );

    let staged_limits = SemanticLimits::builder()
        .max_staged_bytes(1)
        .expect("staged byte limit")
        .build()
        .expect("build staged limit");
    let staged = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(staged_limits),
            |transaction| transaction.remove(A1),
        )
        .expect_err("removal exceeds staged limit");
    assert_eq!(staged.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after staged refusal"),
        before
    );

    let probe = original
        .edit(|transaction| transaction.remove(A1))
        .expect("measure removal snapshot charge");
    let charge = probe.diagnostics().snapshot_bytes();
    assert!(charge > 0);
    let snapshot_limits = SemanticLimits::builder()
        .max_snapshot_bytes(charge - 1)
        .expect("snapshot byte limit")
        .build()
        .expect("build snapshot limit");
    let snapshot = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(snapshot_limits),
            |transaction| transaction.remove(A1),
        )
        .expect_err("removal exceeds snapshot limit");
    assert_eq!(snapshot.code(), Code::LimitExceeded);
    assert_eq!(
        original.to_bytes().expect("publish after snapshot refusal"),
        before
    );
}

#[test]
fn stale_source_refuses_removal_without_mutating_the_origin() {
    let directory = tempfile::tempdir().expect("stale source directory");
    let origin = directory.path().join("stale.xls");
    std::fs::write(&origin, fixture()).expect("write source");
    let stale = litchi_ops::open(&origin).expect("open source-backed XLS");
    std::fs::write(&origin, b"source changed after open").expect("change source");

    assert_eq!(
        stale.remove(A1).expect_err("stale XLS removal").code(),
        Code::IoError
    );
    assert_eq!(
        std::fs::read(&origin).expect("read changed origin"),
        b"source changed after open"
    );
}

#[test]
fn signed_xls_removal_requires_explicit_signature_consumption() {
    let bytes = signed_fixture();
    let signed = open_owner(&bytes, false);
    assert!(signed.is_signed());
    assert_eq!(
        signed.remove(A1).expect_err("signed XLS removal").code(),
        Code::DocumentProtected
    );
    assert_eq!(signed.to_bytes().expect("publish signed source"), bytes);

    let Owner {
        file: signed_file,
        source_guard: _source_guard,
    } = signed;
    let File::Signed(wrapper) = signed_file else {
        panic!("synthetic XLS signature must open as File::Signed");
    };
    let unsigned = wrapper.unsign();
    let removed = unsigned
        .remove(A1)
        .expect("remove explicitly unsigned cell");
    assert!(!removed.is_signed());
    let reopened = litchi_ops::open_bytes(removed.to_bytes().expect("publish unsigned XLS"))
        .expect("reopen unsigned XLS");
    assert!(!reopened.is_signed());
    assert_eq!(cell_text(&reopened, A1).as_deref(), Some(""));
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_xls_refuses_cell_removal_and_plain_publication() {
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
            .remove("/Secret/A1")
            .expect_err("encrypted XLS removal")
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
fn xls_capabilities_advertise_cell_removal_and_granular_gaps() {
    let capabilities = Capabilities::of(FileFormat::Xls);
    assert!(capabilities.verb(Verb::Remove).is_supported());
    let cell = capabilities.element("cell").expect("cell capability");
    assert!(cell.operations().remove);
    for element in ["sheet", "row", "col"] {
        assert!(
            capabilities
                .element(element)
                .expect("structural capability")
                .operations()
                .remove,
            "{element}"
        );
    }
}
