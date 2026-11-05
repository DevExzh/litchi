#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "generated package fixtures fail immediately with local diagnostics"
)]

//! Shared-string closure coverage for package-owned scalar edits.

use std::io::Cursor;

use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_xlsb::Package;
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::{Reference, Scalar, Value};
use litchi_xlsb::raw::{Encoder, Kind, Records, kind};

const SHARED_STRINGS_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";

fn a1() -> Reference {
    Reference::new(0, 0).expect("A1")
}

fn b1() -> Reference {
    Reference::new(0, 1).expect("B1")
}

fn generated_package() -> Package {
    let mut first = builder::Worksheet::new("First");
    first.set_cell(0, 0, "drop");
    first.set_cell(0, 1, "keep");

    let mut second = builder::Worksheet::new("Second");
    second.set_cell(0, 0, "keep");
    second.set_cell(0, 1, "later");

    let mut writer = Builder::new();
    writer.add_worksheet(first);
    writer.add_worksheet(second);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write shared-string fixture");
    Package::from_bytes(&bytes.into_inner()).expect("read shared-string fixture")
}

fn generated_unique_package() -> Package {
    let mut first = builder::Worksheet::new("First");
    first.set_cell(0, 0, "drop");
    first.set_cell(0, 1, "keep");

    let mut second = builder::Worksheet::new("Second");
    second.set_cell(0, 0, "later");

    let mut writer = Builder::new();
    writer.add_worksheet(first);
    writer.add_worksheet(second);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write unique-string fixture");
    Package::from_bytes(&bytes.into_inner()).expect("read unique-string fixture")
}

fn workbook_part(opc: &OpcPackage) -> PackURI {
    opc.main_document_part()
        .expect("main document")
        .partname()
        .clone()
}

fn relationship_target(opc: &OpcPackage, reltype: &str) -> PackURI {
    let workbook = opc.get_part(&workbook_part(opc)).expect("workbook part");
    workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == reltype)
        .expect("relationship")
        .target_partname()
        .expect("relationship target")
}

fn shared_strings_uri(opc: &OpcPackage) -> PackURI {
    let workbook = opc.get_part(&workbook_part(opc)).expect("workbook part");
    workbook
        .rels()
        .iter()
        .find(|relationship| {
            relationship.reltype() == litchi_opc::constants::relationship_type::SHARED_STRINGS
                || relationship.reltype()
                    == litchi_opc::constants::relationship_type::STRICT_SHARED_STRINGS
        })
        .expect("shared-string relationship")
        .target_partname()
        .expect("shared-string target")
}

fn sst_counts(opc: &OpcPackage) -> (u32, u32, usize) {
    let uri = shared_strings_uri(opc);
    let bytes = opc.get_part(&uri).expect("shared-string part").blob();
    let mut counts = None;
    let mut items = 0_usize;
    for item in Records::new(bytes, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("shared-string record");
        match record.kind() {
            kind::BEGIN_SST => {
                let payload = record.payload();
                counts = Some((
                    u32::from_le_bytes(payload[..4].try_into().expect("total")),
                    u32::from_le_bytes(payload[4..8].try_into().expect("unique")),
                ));
            },
            kind::SST_ITEM => items = items.saturating_add(1),
            _ => {},
        }
    }
    let (total, unique) = counts.expect("SST counts");
    (total, unique, items)
}

fn cell_value(package: &Package, sheet: usize, reference: Reference) -> Value {
    package
        .workbook()
        .expect("workbook")
        .cell_values(sheet)
        .expect("cell values")
        .cell(reference)
        .expect("cell lookup")
        .expect("stored cell")
        .value()
        .clone()
}

#[test]
fn scalar_conversion_closes_global_sst_and_reindexes_other_worksheets() {
    let source = generated_package();
    assert_eq!(sst_counts(&source.clone().into_opc()), (4, 3, 3));

    let mut edit = source
        .edit_cell_values(0)
        .expect("select first worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("first worksheet");
    edit.set_scalar(a1(), Scalar::Number(9.5))
        .expect("convert shared string to number");
    let package = edit.commit().expect("commit").into_package();

    // The old `drop` entry remains valid source data: total=3 still permits
    // all three unique entries, so closure avoids gratuitous index changes.
    assert_eq!(sst_counts(&package.clone().into_opc()), (3, 3, 3));
    assert!(matches!(cell_value(&package, 0, a1()), Value::Number(value) if value == 9.5));
    assert!(matches!(
        cell_value(&package, 0, b1()),
        Value::SharedStringIndex(1)
    ));
    assert!(matches!(
        cell_value(&package, 1, a1()),
        Value::SharedStringIndex(1)
    ));
    assert!(matches!(
        cell_value(&package, 1, b1()),
        Value::SharedStringIndex(2)
    ));
}

#[test]
fn scalar_conversion_compacts_only_when_unique_would_exceed_total() {
    let source = generated_unique_package();
    assert_eq!(sst_counts(&source.clone().into_opc()), (3, 3, 3));
    let mut edit = source
        .edit_cell_values(0)
        .expect("select first worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("first worksheet");
    edit.set_scalar(a1(), Scalar::Number(9.5))
        .expect("convert sole shared string");
    let package = edit.commit().expect("commit").into_package();

    assert_eq!(sst_counts(&package.clone().into_opc()), (2, 2, 2));
    assert!(matches!(
        cell_value(&package, 0, b1()),
        Value::SharedStringIndex(0)
    ));
    assert!(matches!(
        cell_value(&package, 1, a1()),
        Value::SharedStringIndex(1)
    ));
}

#[test]
fn scalar_patch_inverse_restores_the_exact_preclosure_package() {
    let source = generated_package();
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("select worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Blank)
        .expect("clear SST cell");
    let commit = edit.commit().expect("commit");
    let changed = commit.package().clone();
    let restored = changed
        .apply_cell_values(&commit.patch().inverse())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect("inverse patch");
    assert_eq!(restored.to_bytes().expect("restored bytes"), before);
}

#[test]
fn custom_shared_string_relationship_target_is_reconciled_without_uri_assumptions() {
    let source = generated_package();
    let mut opc = source.into_opc();
    let old_uri = shared_strings_uri(&opc);
    let custom_uri = PackURI::new("/xl/resources/strings/custom.bin").expect("custom URI");
    let bytes = opc
        .get_part(&old_uri)
        .expect("original SST")
        .blob()
        .to_vec();
    opc.remove_part(&old_uri);
    opc.try_add_part(Box::new(BlobPart::new(
        custom_uri.clone(),
        SHARED_STRINGS_CONTENT_TYPE.to_owned(),
        bytes,
    )))
    .expect("custom SST part");

    let workbook_uri = workbook_part(&opc);
    let relationship_id = {
        let workbook = opc.get_part(&workbook_uri).expect("workbook part");
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype() == litchi_opc::constants::relationship_type::SHARED_STRINGS
                    || relationship.reltype()
                        == litchi_opc::constants::relationship_type::STRICT_SHARED_STRINGS
            })
            .expect("shared-string relationship");
        relationship.r_id().to_owned()
    };
    let target_ref = "resources/strings/custom.bin".to_owned();
    opc.get_part_mut(&workbook_uri)
        .expect("mutable workbook")
        .rels_mut()
        .retarget(&relationship_id, target_ref)
        .expect("retarget SST relationship");
    let custom = Package::from_opc(opc).expect("custom-target package");

    let mut edit = custom
        .edit_cell_values(0)
        .expect("select worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(2.0))
        .expect("convert custom-target SST cell");
    let package = edit.commit().expect("commit").into_package();
    assert_eq!(shared_strings_uri(&package.clone().into_opc()), custom_uri);
    assert_eq!(sst_counts(&package.clone().into_opc()), (3, 3, 3));
}

#[test]
fn opaque_sst_suffix_survives_entry_compaction() {
    let source = generated_package();
    let mut opc = source.into_opc();
    let uri = shared_strings_uri(&opc);
    let mut bytes = opc.get_part(&uri).expect("SST").blob().to_vec();
    let opaque_kind = Kind::new(0x0f00).expect("opaque kind");
    Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
        .write_record(opaque_kind, b"opaque-sst-suffix")
        .expect("opaque suffix");
    opc.get_part_mut(&uri).expect("mutable SST").set_blob(bytes);
    let source = Package::from_opc(opc).expect("opaque suffix package");

    let mut edit = source
        .edit_cell_values(0)
        .expect("select worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(3.0))
        .expect("convert SST cell");
    let package = edit.commit().expect("commit").into_package();
    let uri = shared_strings_uri(&package.clone().into_opc());
    let bytes = package
        .clone()
        .into_opc()
        .get_part(&uri)
        .expect("SST")
        .blob()
        .to_vec();
    assert!(
        Records::new(&bytes, litchi_xlsb::raw::Limits::DEFAULT).any(|record| {
            let record = record.expect("opaque record");
            record.kind() == opaque_kind && record.payload() == b"opaque-sst-suffix"
        })
    );
}

#[test]
fn unsupported_low_cell_family_is_refused_before_remapping() {
    let source = generated_package();
    let mut opc = source.into_opc();
    let worksheet_uri =
        relationship_target(&opc, litchi_opc::constants::relationship_type::WORKSHEET);
    let source_bytes = opc
        .get_part(&worksheet_uri)
        .expect("worksheet")
        .blob()
        .to_vec();
    let opaque_cell_family = Kind::new(0x000c).expect("future cell kind");
    let mut payload = Vec::from([0_u8; 12]);
    payload[8..12].copy_from_slice(&0_u32.to_le_bytes());
    let mut bytes = Vec::new();
    for item in Records::new(&source_bytes, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("worksheet record");
        if record.kind() == kind::END_SHEET_DATA {
            Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
                .write_record(opaque_cell_family, &payload)
                .expect("unsupported cell record");
        }
        Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
            .write_record(record.kind(), record.payload())
            .expect("copy worksheet record");
    }
    opc.get_part_mut(&worksheet_uri)
        .expect("mutable worksheet")
        .set_blob(bytes);
    let source = Package::from_opc(opc).expect("unsupported record package");
    let mut edit = source
        .edit_cell_values(0)
        .expect("select worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    // Relationship iteration may select either worksheet. A reserved family
    // in the edited sheet refuses at admission; one in another sheet refuses
    // when SST reconciliation reaches the full dependency closure.
    let before = source.to_bytes().expect("source bytes");
    let result = edit
        .set_scalar(a1(), Scalar::Number(4.0))
        .and_then(|_| edit.commit());
    assert!(matches!(
        result,
        Err(litchi_xlsb::package::PackageError::Refused(_))
    ));
    assert_eq!(source.to_bytes().expect("source remains"), before);
}

#[test]
fn shared_string_cell_inside_opaque_frt_is_refused() {
    let source = generated_package();
    let mut opc = source.into_opc();
    let worksheet_uri =
        relationship_target(&opc, litchi_opc::constants::relationship_type::WORKSHEET);
    let source_bytes = opc
        .get_part(&worksheet_uri)
        .expect("worksheet")
        .blob()
        .to_vec();
    let opaque_cell_family = kind::CELL_ISST;
    let mut payload = Vec::from([0_u8; 12]);
    payload[..4].copy_from_slice(&100_u32.to_le_bytes());
    payload[8..12].copy_from_slice(&0_u32.to_le_bytes());
    let mut bytes = Vec::new();
    for item in Records::new(&source_bytes, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("worksheet record");
        if record.kind() == kind::END_SHEET_DATA {
            Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
                .write_record(kind::FRT_BEGIN, &[0; 4])
                .expect("FRT begin");
            Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
                .write_record(opaque_cell_family, &payload)
                .expect("unsupported cell record");
            Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
                .write_record(kind::FRT_END, &[])
                .expect("FRT end");
        }
        Encoder::new(&mut bytes, litchi_xlsb::raw::Limits::DEFAULT)
            .write_record(record.kind(), record.payload())
            .expect("copy worksheet record");
    }
    opc.get_part_mut(&worksheet_uri)
        .expect("mutable worksheet")
        .set_blob(bytes);
    let source = Package::from_opc(opc).expect("unsupported record package");
    let mut edit = source
        .edit_cell_values(0)
        .expect("select worksheet")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(4.0))
        .expect("convert SST cell");
    let error = edit.commit().expect_err("unsupported cell family refusal");
    assert!(error.to_string().contains("opaque FRT block"), "{error:?}");
}

#[test]
fn inconsistent_source_sst_count_is_not_silently_repaired() {
    let mut opc = generated_package().into_opc();
    let uri = shared_strings_uri(&opc);
    let mut output = Vec::new();
    for item in Records::new(
        opc.get_part(&uri).expect("SST").blob(),
        litchi_xlsb::raw::Limits::DEFAULT,
    ) {
        let record = item.expect("record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BEGIN_SST {
            payload[..4].copy_from_slice(&100_u32.to_le_bytes());
        }
        Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT)
            .write_record(record.kind(), &payload)
            .expect("record");
    }
    opc.get_part_mut(&uri).expect("SST").set_blob(output);
    let source = Package::from_opc(opc).expect("count fixture");
    let before = source.to_bytes().expect("source");
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(4.0)).expect("scalar");
    let error = edit.commit().expect_err("refuse inconsistent count");
    assert!(error.to_string().contains("source count"), "{error:?}");
    assert_eq!(source.to_bytes().expect("unchanged"), before);
}
