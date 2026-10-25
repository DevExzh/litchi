//! XLSX transaction coalescing and barrier semantics.
#![cfg(feature = "xlsx")]

use std::fs;

use litchi_opc::PackURI;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, Props, SourceMode, props,
};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const D1: &str = "/sheet[1]/D1";

fn fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("find native worksheet")
            .expect("minimal workbook has Sheet1");
        sheet.set("A1", "seed").expect("seed text");
        sheet.set("B1", 2_i32).expect("seed number");
        sheet
            .set(
                "C1",
                litchi_xlsx::Formula::new("B1*2").expect("seed formula"),
            )
            .expect("seed formula cell");
        sheet.set("D1", true).expect("seed boolean");
    }
    edit.commit()
        .expect("commit native fixture")
        .into_package()
        .to_bytes()
        .expect("serialize fixture")
}

fn merged_payload_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open fixture");
    let package =
        litchi_xlsx::Package::from_bytes(&workbook.to_bytes().expect("serialize fixture"))
            .expect("open fixture package")
            .into_plain_opc();
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
    let sheet = std::str::from_utf8(package.get_part(&sheet_uri).expect("worksheet").blob())
        .expect("worksheet XML");
    let sheet = sheet.replace(
        "</worksheet>",
        r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#,
    );
    let mut package = package;
    package
        .get_part_mut(&sheet_uri)
        .expect("worksheet part")
        .set_blob(sheet.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt merged payload package")
        .to_plain_bytes()
        .expect("serialize merged payload package")
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

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .text()
        .expect("read cell text")
        .map(std::borrow::Cow::into_owned)
}

fn cell_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .format()
        .expect("read cell format")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

fn meta_item() -> BatchItem<'static> {
    let mut item = BatchItem::new("meta");
    item.dump_version = Some(2);
    item
}

#[test]
fn closure_direct_payload_formula_and_clear_share_one_native_commit() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "changed" })?;
                transaction.set(C1, props! { "formula" => "A1" })?;
                transaction.set(D1, props! { "clear" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("coalesced closure edit");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("changed"));
        assert_eq!(
            cell_property(commit.file(), C1, "formula").as_deref(),
            Some("A1")
        );
        assert_eq!(cell_text(commit.file(), D1).as_deref(), Some(""));

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay coalesced edit");
        assert_eq!(
            replayed.to_bytes().expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse coalesced edit");
        assert_eq!(
            restored.to_bytes().expect("serialize restored"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn atomic_batch_reports_per_item_noops_while_using_one_native_commit() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let commit = original
            .batch(
                [
                    batch_set(A1, Props::new().with("value", "seed")),
                    batch_set(B1, Props::new().with("value", 3)),
                    batch_set(C1, Props::new().with("clear", true)),
                    batch_set(D1, Props::new().with("value", true)),
                ],
                BatchMode::Atomic,
            )
            .expect("coalesced atomic batch");

        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert!(commit.diagnostics().changed());
        let changed = commit
            .value()
            .iter()
            .map(|result| match result {
                ItemResult::Applied(applied) => applied.changed,
                other => panic!("unexpected batch result: {other:?}"),
            })
            .collect::<Vec<_>>();
        assert_eq!(changed, [false, true, true, false]);
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("3"));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some(""));

        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse atomic batch");
        assert_eq!(
            restored.to_bytes().expect("serialize restored"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn repeated_cells_and_font_assignments_are_barriers() {
    let original = open_owner(&fixture(), false);
    let commit = original
        .edit(|transaction| {
            transaction.set(A1, props! { "value" => "first" })?;
            transaction.set(A1, props! { "value" => "second" })?;
            transaction.set(A1, props! { "font.bold" => true })?;
            transaction.set(B1, props! { "font.italic" => true })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("repeated cell/font barrier transaction");

    assert_eq!(commit.diagnostics().operations(), 4);
    assert_eq!(
        commit.diagnostics().commits(),
        4,
        "repeated targets and font assignments each close the payload run"
    );
    assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("second"));
    assert_eq!(
        cell_property(commit.file(), A1, "font.bold").as_deref(),
        Some("true")
    );
    assert_eq!(
        cell_property(commit.file(), B1, "font.italic").as_deref(),
        Some("true")
    );
}

#[test]
fn caught_coalesced_set_failure_poisoning_discards_the_pending_run() {
    for eager in [false, true] {
        let original = open_owner(&fixture(), eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let error = original
            .edit::<()>(|transaction| {
                transaction.set(A1, props! { "value" => "staged" })?;
                let failure = transaction
                    .set(B1, props! { "unknownCellProperty" => true })
                    .expect_err("unknown property must fail");
                assert_eq!(failure.item(), Some(1));
                Ok(())
            })
            .expect_err("a caught coalesced failure must still poison the edit");

        assert_eq!(error.item(), Some(1));
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn selectors_and_reads_close_a_direct_run() {
    let original = open_owner(&fixture(), false);
    let selector = "Sheet1!cell[value>0]";
    let commit = original
        .edit(|transaction| {
            transaction.set(A1, props! { "value" => "before selector" })?;
            transaction.set(selector, props! { "value" => 0 })?;
            let _ = transaction.get(A1)?.view().text()?;
            transaction.set(B1, props! { "value" => 9 })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("selector/read barrier transaction");

    assert_eq!(commit.diagnostics().operations(), 3);
    assert_eq!(
        commit.diagnostics().commits(),
        3,
        "direct run, selector, and post-read direct run are separate commits"
    );
    assert_eq!(
        cell_text(commit.file(), A1).as_deref(),
        Some("before selector")
    );
    assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("9"));
}

#[test]
fn metadata_closes_atomic_batch_coalescing() {
    let original = open_owner(&fixture(), false);
    let commit = original
        .batch(
            [
                batch_set(A1, Props::new().with("value", "before metadata")),
                meta_item(),
                batch_set(B1, Props::new().with("value", 9)),
            ],
            BatchMode::Atomic,
        )
        .expect("metadata barrier batch");

    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(commit.diagnostics().commits(), 2);
    assert!(matches!(
        &commit.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &commit.value()[1],
        ItemResult::Applied(applied) if !applied.changed
    ));
    assert!(matches!(
        &commit.value()[2],
        ItemResult::Applied(applied) if applied.changed
    ));
}

#[test]
fn best_effort_isolates_each_item_and_preserves_wire_indices() {
    let original = open_owner(&fixture(), false);
    let original_bytes = original.to_bytes().expect("serialize source");
    let commit = original
        .batch(
            [
                batch_set(A1, Props::new().with("value", "kept")),
                batch_set(B1, Props::new().with("unknownCellProperty", true)),
                batch_set(D1, Props::new().with("value", false)),
            ],
            BatchMode::BestEffort,
        )
        .expect("best-effort batch");

    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(commit.diagnostics().commits(), 2);
    assert!(matches!(
        &commit.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &commit.value()[1],
        ItemResult::Failed(error) if error.item() == Some(1)
    ));
    assert!(matches!(
        &commit.value()[2],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("kept"));
    assert_eq!(cell_text(commit.file(), D1).as_deref(), Some("FALSE"));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse best-effort batch")
            .to_bytes()
            .expect("serialize restored"),
        original_bytes
    );
}

#[test]
fn coalescing_limits_are_checked_before_native_publication() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let operation_limit = SemanticLimits::builder()
            .max_operations(2)
            .expect("valid operation limit")
            .build()
            .expect("build operation limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(operation_limit),
                |transaction| {
                    transaction.set(A1, props! { "value" => "one" })?;
                    transaction.set(B1, props! { "value" => "two" })?;
                    transaction.set(C1, props! { "value" => "three" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("operation limit");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(error.item(), Some(2));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::Operations,
                observed: 3,
                maximum: 2,
            })
        ));
        assert_eq!(
            original.to_bytes().expect("source unchanged"),
            original_bytes
        );

        let target_limit = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("valid target limit")
            .build()
            .expect("build target limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(target_limit),
                |transaction| transaction.set("Sheet1!A1:B1", props! { "value" => "too many" }),
            )
            .expect_err("expanded target limit");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(error.item(), Some(0));
        assert_eq!(
            original.to_bytes().expect("source unchanged"),
            original_bytes
        );
    }
}

#[test]
fn staged_bytes_budget_closes_a_payload_run() {
    let limits = SemanticLimits::builder()
        .max_staged_bytes(128)
        .expect("valid staged-byte limit")
        .build()
        .expect("build staged-byte limit");
    let original = open_owner(&fixture(), false);
    let commit = original
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |transaction| {
                transaction.set(A1, props! { "value" => false })?;
                transaction.set(B1, props! { "value" => false })?;
                transaction.set(D1, props! { "value" => false })?;
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect("staged-byte barrier edit");

    assert_eq!(commit.diagnostics().operations(), 3);
    assert_eq!(
        commit.diagnostics().commits(),
        2,
        "two 64-byte payload plans fit the run; the third starts another run"
    );
    assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("FALSE"));
    assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("FALSE"));
    assert_eq!(cell_text(commit.file(), D1).as_deref(), Some("FALSE"));
}

#[test]
fn source_change_while_a_payload_run_is_pending_refuses_the_stage_or_flush() {
    for mutate_before_next_stage in [true, false] {
        let directory = tempfile::tempdir().expect("temporary source directory");
        let path = directory.path().join("pending-run.xlsx");
        let bytes = fixture();
        fs::write(&path, &bytes).expect("write source fixture");
        let original =
            litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                .expect("open path-backed source owner");

        let error = original
            .edit::<()>(|transaction| {
                transaction.set(A1, props! { "value" => "staged" })?;
                fs::write(&path, b"changed physical source").expect("mutate source");
                if mutate_before_next_stage {
                    transaction.set(B1, props! { "value" => "rejected" })?;
                }
                Ok(())
            })
            .expect_err("a stale source must not publish a pending payload run");

        assert_eq!(error.code(), Code::IoError);
        assert_eq!(
            fs::read(&path).expect("read changed source"),
            b"changed physical source"
        );
    }
}

#[test]
fn late_native_payload_failure_reports_wire_index_and_discards_group() {
    let bytes = merged_payload_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize merged source");
        let error = original
            .batch(
                [
                    meta_item(),
                    batch_set(A1, Props::new().with("value", "must not publish")),
                    batch_set(B1, Props::new().with("value", "covered")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("covered native payload must fail at commit");
        assert_eq!(error.item(), Some(2));
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(
            original.to_bytes().expect("source unchanged"),
            original_bytes
        );
    }
}
