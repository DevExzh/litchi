#![allow(
    clippy::expect_used,
    clippy::float_cmp,
    clippy::unwrap_used,
    reason = "focused generated fixtures use immediate diagnostics"
)]

//! Focused coverage for the package-owned scalar cell assignment.

use std::io::Cursor;

use litchi_core::sheet::CellValue;
use litchi_opc::PackURI;
use litchi_xlsb::Package;
use litchi_xlsb::cell_values::{Limits, Reference, Scalar, StyleIndex, TransferLimits, Value};
use litchi_xlsb::formula::ParsedFormula;
use litchi_xlsb::raw::{Records, Writer, kind};
use litchi_xlsb::writer::{Builder, MutableWorksheet};

fn a1() -> Reference {
    Reference::new(0, 0).expect("A1")
}

fn b1() -> Reference {
    Reference::new(0, 1).expect("B1")
}

fn c1() -> Reference {
    Reference::new(0, 2).expect("C1")
}

fn d1() -> Reference {
    Reference::new(0, 3).expect("D1")
}

fn e1() -> Reference {
    Reference::new(0, 4).expect("E1")
}

fn style_zero() -> StyleIndex {
    StyleIndex::new(0).expect("default style")
}

fn package_with_scalars() -> Package {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, 1.0_f64);
    sheet.set_cell(0, 1, true);
    sheet.set_cell(0, 2, "same");
    sheet.set_cell(0, 3, CellValue::Empty);
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write fixture");
    Package::from_bytes(bytes.into_inner()).expect("read fixture")
}

fn package_with_grouped_formula() -> Package {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet
        .set_shared_formula(0, 0, 0, 1, "1")
        .expect("formula fixture");
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write formula fixture");
    Package::from_bytes(bytes.into_inner()).expect("read formula fixture")
}

fn package_with_missing_group_follower() -> Package {
    let package = package_with_grouped_formula();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").expect("worksheet URI");
    let mut opc = package.into_opc();
    let source = opc
        .get_part(&worksheet)
        .expect("worksheet part")
        .blob()
        .to_vec();
    let mut output = Vec::new();
    let mut row = None;
    for item in Records::new(&source) {
        let record = item.expect("worksheet record");
        if record.kind() == kind::ROW_HDR {
            row = Some(u32::from_le_bytes(
                record.payload()[..4].try_into().expect("row payload"),
            ));
        }
        let drop_follower = row == Some(0)
            && matches!(
                record.kind(),
                kind::FMLA_STRING | kind::FMLA_NUM | kind::FMLA_BOOL | kind::FMLA_ERROR
            )
            && u32::from_le_bytes(
                record.payload()[..4]
                    .try_into()
                    .expect("formula cell payload"),
            ) == 1;
        if !drop_follower {
            Writer::new(&mut output)
                .write_record(record.kind(), record.payload())
                .expect("copy worksheet record");
        }
    }
    opc.get_part_mut(&worksheet)
        .expect("worksheet part")
        .set_blob(output);
    Package::from_opc(opc).expect("read missing-follower fixture")
}

fn package_with_ordinary_formula() -> Package {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let formula = ParsedFormula {
        rgce: vec![0x1E, 2, 0],
        rgcb: Vec::new(),
    };
    sheet.set_cell_formula_binary(0, 0, CellValue::Float(2.0), formula, false, 0);
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write ordinary formula fixture");
    Package::from_bytes(bytes.into_inner()).expect("read ordinary formula fixture")
}

fn package_with_typed_formula_caches() -> Package {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let formula = ParsedFormula {
        rgce: vec![0x1E, 1, 0],
        rgcb: Vec::new(),
    };
    sheet.set_cell_formula_binary(0, 0, CellValue::Float(1.0), formula.clone(), false, 0);
    sheet.set_cell_formula_binary(0, 1, CellValue::Bool(true), formula.clone(), false, 0);
    sheet.set_cell_formula_binary(
        0,
        2,
        CellValue::String("same".to_owned()),
        formula,
        false,
        0,
    );
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("write typed formula fixture");
    Package::from_bytes(bytes.into_inner()).expect("read typed formula fixture")
}

fn snapshot_cell(package: &Package, reference: Reference) -> Value {
    package
        .workbook()
        .expect("workbook")
        .cell_values(0)
        .expect("cell values")
        .cell(reference)
        .expect("cell lookup")
        .expect("stored cell")
        .value()
        .clone()
}

#[test]
fn assigns_scalars_and_inserts_missing_cells() {
    let source = package_with_scalars();
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(2.5)).expect("number");
    edit.set_scalar(b1(), Scalar::Boolean(false))
        .expect("boolean");
    edit.set_scalar(c1(), Scalar::Text("changed".to_owned()))
        .expect("text");
    edit.set_scalar(d1(), Scalar::Text("inserted".to_owned()))
        .expect("missing text");
    edit.set_scalar(e1(), Scalar::Blank)
        .expect("missing blank no-op");
    let package = edit.commit().expect("commit").into_package();

    assert!(matches!(snapshot_cell(&package, a1()), Value::Number(value) if value == 2.5));
    assert!(matches!(
        snapshot_cell(&package, b1()),
        Value::Boolean(false)
    ));
    assert!(
        matches!(snapshot_cell(&package, c1()), Value::InlineString(value) if value == "changed")
    );
    assert!(
        matches!(snapshot_cell(&package, d1()), Value::InlineString(value) if value == "inserted")
    );
    assert!(
        package
            .workbook()
            .expect("workbook")
            .cell_values(0)
            .expect("cell values")
            .cell(e1())
            .expect("cell lookup")
            .is_none()
    );
}

#[test]
fn same_value_sst_assignment_is_an_exact_noop() {
    let source = package_with_scalars();
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(c1(), Scalar::Text("same".to_owned()))
        .expect("same shared string");
    let commit = edit.commit().expect("commit");
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.patch().before(), original.as_slice());
    assert_eq!(commit.patch().after(), original.as_slice());
}

#[test]
fn family_changes_keep_style_and_phonetic_flag() {
    let source = package_with_scalars();
    let mut seed = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    seed.set_style(a1(), style_zero()).expect("style");
    seed.set_show_phonetic(a1(), true).expect("phonetic flag");
    let source = seed.commit().expect("seed commit").into_package();

    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Text("text".to_owned()))
        .expect("number to text");
    let package = edit.commit().expect("text commit").into_package();
    let workbook = package.workbook().expect("workbook");
    let snapshot = workbook.cell_values(0).expect("cell values");
    let cell = snapshot.cell(a1()).expect("cell lookup").expect("A1");
    assert!(matches!(cell.value(), Value::InlineString(value) if value == "text"));
    assert_eq!(cell.style(), style_zero());
    assert!(cell.show_phonetic());

    let mut clear = package
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    clear.set_scalar(a1(), Scalar::Blank).expect("clear scalar");
    let cleared = clear.commit().expect("clear commit").into_package();
    let snapshot = cleared
        .workbook()
        .expect("workbook")
        .cell_values(0)
        .expect("cell values");
    let cell = snapshot.cell(a1()).expect("cell lookup").expect("A1");
    assert!(matches!(cell.value(), Value::Blank));
    assert_eq!(cell.style(), style_zero());
    assert!(cell.show_phonetic());
}

#[test]
fn invalid_scalar_is_rejected_before_mutation() {
    let source = package_with_scalars();
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values_with_limits(
            0,
            Limits::new(512 * 1024, 128, 512 * 1024, 4),
            TransferLimits::DEFAULT,
        )
        .expect("selector")
        .expect("worksheet");
    assert!(edit.set_scalar(a1(), Scalar::Number(f64::NAN)).is_err());
    assert!(
        edit.set_scalar(a1(), Scalar::Text("12345".to_owned()))
            .is_err()
    );
    let commit = edit.commit().expect("unchanged commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), original.as_slice());
    assert_eq!(commit.patch().after(), original.as_slice());
}

#[test]
fn ordinary_formula_equal_scalar_materializes_direct_value() {
    let source = package_with_ordinary_formula();
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(2.0))
        .expect("equal cached value");
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.patch().after(), original.as_slice());
    let package = commit.into_package();
    assert!(matches!(snapshot_cell(&package, a1()), Value::Number(value) if value == 2.0));
}

#[test]
fn ordinary_formula_changed_scalar_and_clear_materialize_direct_values() {
    let source = package_with_ordinary_formula();
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(3.0))
        .expect("changed cached value");
    let changed = edit.commit().expect("changed commit").into_package();
    assert!(matches!(snapshot_cell(&changed, a1()), Value::Number(value) if value == 3.0));

    let mut clear = changed
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    clear
        .set_scalar(a1(), Scalar::Blank)
        .expect("clear ordinary formula");
    let cleared = clear.commit().expect("clear commit").into_package();
    assert!(matches!(snapshot_cell(&cleared, a1()), Value::Blank));
}

#[test]
fn formula_cache_equal_scalars_materialize_direct_typed_values() {
    let source = package_with_typed_formula_caches();
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(1.0))
        .expect("equal numeric cache");
    edit.set_scalar(b1(), Scalar::Boolean(true))
        .expect("equal Boolean cache");
    edit.set_scalar(c1(), Scalar::Text("same".to_owned()))
        .expect("equal string cache");
    let package = edit.commit().expect("commit").into_package();
    assert!(matches!(snapshot_cell(&package, a1()), Value::Number(value) if value == 1.0));
    assert!(matches!(
        snapshot_cell(&package, b1()),
        Value::Boolean(true)
    ));
    assert!(matches!(snapshot_cell(&package, c1()), Value::InlineString(value) if value == "same"));
}

#[test]
fn scalar_assignment_refuses_grouped_formula_cells() {
    let source = package_with_grouped_formula();
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    assert!(edit.set_scalar(a1(), Scalar::Number(4.0)).is_err());
    assert!(edit.set_scalar(b1(), Scalar::Blank).is_err());
    let commit = edit.commit().expect("unchanged commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), original.as_slice());
    assert_eq!(commit.patch().after(), original.as_slice());
}

#[test]
fn scalar_assignment_refuses_missing_group_follower() {
    let source = package_with_missing_group_follower();
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .expect("worksheet");
    assert!(edit.set_scalar(b1(), Scalar::Number(4.0)).is_err());
    assert!(edit.set_scalar(b1(), Scalar::Blank).is_err());
    let commit = edit.commit().expect("unchanged commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), original.as_slice());
    assert_eq!(commit.patch().after(), original.as_slice());
}

#[test]
fn formula_assignment_refuses_transitional_and_strict_calculation_chains() {
    for relationship_type in [
        litchi_opc::constants::relationship_type::CALC_CHAIN,
        litchi_opc::constants::relationship_type::STRICT_CALC_CHAIN,
    ] {
        let mut opc = package_with_ordinary_formula().into_opc();
        let workbook_uri = opc
            .main_document_part()
            .expect("workbook")
            .partname()
            .clone();
        opc.add_part(Box::new(litchi_opc::BlobPart::new(
            PackURI::new("/xl/calcChain.bin").expect("URI"),
            "application/vnd.ms-excel.calcChain".to_string(),
            Vec::new(),
        )));
        opc.get_part_mut(&workbook_uri)
            .expect("workbook")
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_string(),
                "calcChain.bin".to_string(),
                "rIdCalcChain".to_string(),
                litchi_opc::TargetMode::Internal,
            )
            .expect("chain relationship");
        let source = Package::from_opc(opc).expect("opaque chain package");
        let mut edit = source
            .edit_cell_values(0)
            .expect("edit")
            .expect("worksheet");
        let error = edit
            .set_scalar(a1(), Scalar::Number(2.0))
            .expect_err("chain guard");
        assert!(error.to_string().contains("calculation chain"));
    }
}
