//! Unified XLSB cell removal through the public [`File`] and transaction APIs.
//!
//! Cell removal is deliberately a sparse operation: coordinates are never
//! shifted, a missing coordinate is an exact no-op, and a formula group is
//! refused until the native owner supports group lifecycle edits.  The same assertions run through retained source and eager owners.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::FileFormat;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, SourceMode};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const D1: &str = "/sheet[1]/D1";
const A2: &str = "/sheet[1]/A2";
const ARCHIVE_A1: &str = "/sheet[2]/A1";
const ARCHIVE_B1: &str = "/sheet[2]/B1";

fn fixture() -> Vec<u8> {
    use litchi_core::sheet::CellValue;
    use litchi_xlsb::builder::{self, Builder};
    use litchi_xlsb::formula::ParsedFormula;

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "seed");
    data.set_cell(0, 1, 2.5_f64);
    data.set_cell(0, 2, true);
    data.set_cell_formula_binary(
        0,
        3,
        CellValue::Float(8.0),
        ParsedFormula {
            rgce: vec![0x1e, 2, 0],
            rgcb: Vec::new(),
        },
        false,
        0,
    );
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

fn grouped_formula_fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_shared_formula(0, 4, 0, 5, "1")
        .expect("install shared formula group");
    data.set_cell(1, 0, "untouched");
    writer.add_worksheet(data);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build grouped formula fixture")
        .write_to(&mut output)
        .expect("serialize grouped formula fixture");
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

fn workbook_query(selector: &'static str) -> Query<'static> {
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
    let source = OpcPackage::from_bytes(&fixture(), litchi_opc::SourceLimits::default())
        .expect("parse XLSB fixture as OPC");
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .expect("copy package relationship");
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).expect("find OPC part");
        let mut copy = BlobPart::new(
            name.clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("copy part relationship");
        }
        package.try_add_part(Box::new(copy)).expect("copy OPC part");
    }

    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin URI");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    package
        .try_add_part(Box::new(origin_part))
        .expect("signature origin part");
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .expect("signature XML part");
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature origin relationship");
    PackageWriter::to_bytes(&package).expect("serialize signed XLSB fixture")
}

fn part(bytes: &[u8], path: &str) -> Option<Vec<u8>> {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).ok()?;
    let uri = PackURI::new(path).ok()?;
    package.get_part(&uri).ok().map(|part| part.blob().to_vec())
}

fn column_name(mut column: u32) -> String {
    let mut name = String::new();
    loop {
        let digit = u8::try_from(column % 26).expect("column digit");
        name.push(char::from(b'A' + digit));
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    name.chars().rev().collect()
}

fn cell_path(sheet: usize, row: u32, column: u32) -> String {
    format!("/sheet[{}]/{}{}", sheet + 1, column_name(column), row + 1)
}

/// Return a styled, directly stored cell from a corpus fixture.  The helper
/// keeps this test independent of the producer's sheet names and order while
/// still proving that deleting a styled record does not rewrite unrelated
/// opaque parts.
fn styled_corpus_fixture() -> (Vec<u8>, String, usize, litchi_xlsb::cell_values::Reference) {
    for name in [
        "universal-content.xlsb",
        "Simple.xlsb",
        "date.xlsb",
        "cond_format.xlsb",
        "comments.xlsb",
        "hyperlink.xlsb",
    ] {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ooxml/xlsb")
                .join(name),
        )
        .expect("XLSB corpus fixture");
        let package = litchi_xlsb::Package::from_bytes(&bytes).expect("native XLSB package");
        let workbook = package.workbook().expect("native workbook");
        for tab in workbook.sheets() {
            let Some(worksheet) = tab.worksheet_position() else {
                continue;
            };
            let Ok(snapshot) = workbook.cell_values(worksheet.get()) else {
                continue;
            };
            if let Some(cell) = snapshot
                .cells()
                .find(|cell| cell.style().get() != 0 && cell.formula().is_none())
            {
                return (
                    bytes,
                    cell_path(
                        tab.position().get(),
                        cell.reference().row(),
                        cell.reference().column(),
                    ),
                    worksheet.get(),
                    cell.reference(),
                );
            }
        }
    }
    panic!("the XLSB corpus has no directly stored styled cell")
}

#[test]
fn direct_ranges_and_virtual_coordinates_remove_without_shifting_cells() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");

        let removed = original.remove(B1).expect("remove numeric cell");
        assert_eq!(cell_text(&removed, B1).as_deref(), Some("(empty)"));
        assert_eq!(cell_text(&removed, C1).as_deref(), Some("TRUE"));
        assert!(cell_property(&removed, B1, "formula").is_none());
        assert!(cell_properties(&removed, B1).is_empty());

        let ranged = removed.remove("Data!C1:C2").expect("remove a sparse range");
        assert_eq!(cell_text(&ranged, C1).as_deref(), Some("(empty)"));
        assert_eq!(
            cell_text(&ranged, "/sheet[1]/C2").as_deref(),
            Some("(empty)")
        );
        assert_eq!(cell_text(&ranged, A1).as_deref(), Some("seed"));
        assert_eq!(cell_text(&ranged, A2).as_deref(), Some("second row"));
        assert_eq!(cell_text(&ranged, D1).as_deref(), Some("8"));
        assert!(cell_property(&ranged, D1, "formula").is_some());

        let virtual_noop = original
            .remove("/sheet[1]/Z99")
            .expect("missing virtual cell is an exact no-op");
        assert_eq!(
            virtual_noop.to_bytes().expect("virtual no-op bytes"),
            before
        );
        let virtual_range_noop = original
            .remove("Data!Y99:Z100")
            .expect("missing virtual range is an exact no-op");
        assert_eq!(
            virtual_range_noop.to_bytes().expect("virtual range bytes"),
            before
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            before
        );
    }
}

#[test]
fn query_removal_spans_worksheets_and_empty_queries_are_refused() {
    let bytes = fixture();
    let selector = workbook_query("cell[value=seed or value=untouched]");
    let empty = workbook_query("cell[value=does-not-exist]");

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        let removed = original
            .remove(selector.clone())
            .expect("remove matching cells on both worksheets");
        assert_eq!(cell_text(&removed, A1).as_deref(), Some("(empty)"));
        assert_eq!(cell_text(&removed, ARCHIVE_A1).as_deref(), Some("(empty)"));
        assert_eq!(cell_text(&removed, B1).as_deref(), Some("2.5"));
        assert_eq!(cell_text(&removed, ARCHIVE_B1).as_deref(), Some("9"));

        let error = original
            .remove(empty.clone())
            .expect_err("an empty query cannot be treated as a successful no-op");
        assert_eq!(error.code(), Code::ZeroMatches);
        assert_eq!(
            original.to_bytes().expect("source after empty query"),
            before
        );
    }
}

#[test]
fn ordinary_formula_and_styled_records_are_removed_without_rewriting_opaque_parts() {
    let (bytes, styled_path, worksheet, reference) = styled_corpus_fixture();
    let original = open_owner(&bytes, false);
    let before = original.to_bytes().expect("publish corpus fixture");
    let removed = original
        .remove(styled_path.as_str())
        .expect("remove a directly stored styled cell");
    assert_eq!(
        cell_text(&removed, &styled_path).as_deref(),
        Some("(empty)")
    );
    let after = removed.to_bytes().expect("publish removed corpus fixture");
    let native = litchi_xlsb::Package::from_bytes(&after).expect("native package");
    let workbook = native.workbook().expect("native projection");
    assert!(
        workbook
            .cell_values(worksheet)
            .expect("cells")
            .cell(reference)
            .expect("cell")
            .is_none()
    );
    let mut preserved = 0;
    for path in [
        "/xl/styles.bin",
        "/xl/drawings/drawing1.xml",
        "/xl/comments1.bin",
        "/xl/calcChain.bin",
        "/xl/theme/theme1.xml",
    ] {
        if let Some(expected) = part(&before, path) {
            assert_eq!(
                part(&after, path).as_deref(),
                Some(expected.as_slice()),
                "{path}"
            );
            preserved += 1;
        }
    }
    assert!(
        preserved > 0,
        "corpus fixture should contain opaque members"
    );

    let formula_path = D1;
    let formula_owner = open_owner(&fixture(), false);
    assert!(cell_property(&formula_owner, formula_path, "formula").is_some());
    let formula_removed = formula_owner
        .remove(formula_path)
        .expect("remove ordinary formula record");
    assert_eq!(
        cell_text(&formula_removed, formula_path).as_deref(),
        Some("(empty)")
    );
    assert!(cell_property(&formula_removed, formula_path, "formula").is_none());
}

#[test]
fn partial_formula_group_removal_is_a_typed_failure_and_is_atomic() {
    let bytes = grouped_formula_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish grouped fixture");
        let error = original
            .remove("Data!E1")
            .expect_err("partial shared-formula removal must be refused");
        assert_eq!(error.code(), Code::UnsupportedType);
        assert!(matches!(
            error.leaf(),
            Source::Xlsb(litchi_xlsb::package::PackageError::Refused(refusal))
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedTarget
        ));
        assert_eq!(
            original.to_bytes().expect("source after formula refusal"),
            before
        );
    }
}

#[test]
fn remove_transactions_coalesce_mixed_cell_operations_and_replay_inverse() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish transaction source");
        let commit = original
            .edit(|transaction| {
                transaction.remove(A1)?;
                transaction.set(B1, litchi_ops::props! { "value" => 8.25 })?;
                transaction.remove(C1)?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mixed remove/set transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("(empty)"));
        assert_eq!(cell_text(commit.file(), B1).as_deref(), Some("8.25"));
        assert_eq!(cell_text(commit.file(), C1).as_deref(), Some("(empty)"));

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay remove patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse remove patch");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            before
        );
    }
}

#[test]
fn remove_batches_are_atomic_or_attribute_failures_per_item() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
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
            .expect_err("atomic remove batch must discard its prefix");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::PathNotFound);
        assert_eq!(
            original.to_bytes().expect("source after atomic failure"),
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
            .expect("best-effort remove batch");
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
        assert_eq!(
            cell_text(best_effort.file(), A1).as_deref(),
            Some("(empty)")
        );
        assert_eq!(
            cell_text(best_effort.file(), ARCHIVE_A1).as_deref(),
            Some("(empty)")
        );
        let restored = best_effort
            .patch()
            .inverse()
            .apply(best_effort.file())
            .expect("inverse best-effort remove batch");
        assert_eq!(restored.to_bytes().expect("restored batch"), before);
    }
}

#[test]
fn remove_limits_refuse_before_native_publication() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish limit source");
        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("valid target limit")
            .build()
            .expect("build target limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| transaction.remove("Data!A1:B1"),
            )
            .expect_err("range removal must honor expanded-target limits");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(error.item(), Some(0));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: 2,
                maximum: 1,
            })
        ));
        assert_eq!(original.to_bytes().expect("source after limit"), before);
    }
}

#[test]
fn signed_xlsb_removal_requires_explicit_signature_consumption() {
    let bytes = signed_fixture();
    for eager in [false, true] {
        let signed = open_owner(&bytes, eager);
        assert!(signed.is_signed());
        assert_eq!(signed.to_bytes().expect("publish signed source"), bytes);
        assert_eq!(
            signed.capabilities().verb(Verb::Remove).code(),
            Some(Code::DocumentProtected)
        );
        let error = signed
            .remove(A1)
            .expect_err("signed cell removal must refuse");
        assert_eq!(error.code(), Code::DocumentProtected);
        assert_eq!(signed.to_bytes().expect("signed source unchanged"), bytes);
    }
}

#[test]
fn xlsb_capabilities_advertise_removal_for_cells_and_sheets() {
    let report = Capabilities::of(FileFormat::Xlsb);
    assert!(report.verb(Verb::Remove).is_supported());
    assert!(
        report
            .element("cell")
            .expect("cell capability")
            .operations()
            .remove
    );
    assert!(
        report
            .element("sheet")
            .expect("sheet capability")
            .operations()
            .remove
    );
    for element in ["range", "row", "col", "table", "workbook"] {
        assert!(
            !report
                .element(element)
                .expect("declared XLSB element")
                .operations()
                .remove,
            "XLSB removal must remain limited to cells and sheets for {element}"
        );
    }
    let removable = report
        .elements()
        .filter(|element| element.operations().remove)
        .map(|element| element.name())
        .collect::<Vec<_>>();
    // Record 0322 adds the worksheet feature nodes; record 0330 claims the
    // last defined name's removal (record 0261).
    assert_eq!(
        removable,
        [
            "sheet",
            "cell",
            "namedrange",
            "comment",
            "dataValidation",
            "conditionalFormatting",
            "autofilter",
            "hyperlink",
            "picture",
            "shape",
            "rowbreak",
            "colbreak"
        ]
    );
}
