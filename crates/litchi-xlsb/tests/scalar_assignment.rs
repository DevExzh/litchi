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
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::{Limits, Reference, Scalar, StyleIndex, Value};
use litchi_xlsb::formula::ParsedFormula;
use litchi_xlsb::raw::{Encoder, Records, kind};

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
    let mut sheet = builder::Worksheet::new("Sheet1");
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
    Package::from_bytes(&bytes.into_inner()).expect("read fixture")
}

fn package_with_grouped_formula() -> Package {
    let mut sheet = builder::Worksheet::new("Sheet1");
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
    Package::from_bytes(&bytes.into_inner()).expect("read formula fixture")
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
    for item in Records::new(&source, litchi_xlsb::raw::Limits::DEFAULT) {
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
            Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT)
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
    let mut sheet = builder::Worksheet::new("Sheet1");
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
    Package::from_bytes(&bytes.into_inner()).expect("read ordinary formula fixture")
}

fn package_with_typed_formula_caches() -> Package {
    let mut sheet = builder::Worksheet::new("Sheet1");
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
    Package::from_bytes(&bytes.into_inner()).expect("read typed formula fixture")
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
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    seed.set_style(a1(), style_zero()).expect("style");
    seed.set_show_phonetic(a1(), true).expect("phonetic flag");
    let source = seed.commit().expect("seed commit").into_package();

    let mut edit = source
        .edit_cell_values(0)
        .expect("selector")
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
        .edit_cell_values_with_options(
            0,
            litchi_xlsb::cell_values::EditOptions::new().limits(Limits::new(
                512 * 1024,
                128,
                512 * 1024,
                4,
            )),
        )
        .expect("selector")
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    edit.set_scalar(a1(), Scalar::Number(3.0))
        .expect("changed cached value");
    let changed = edit.commit().expect("changed commit").into_package();
    assert!(matches!(snapshot_cell(&changed, a1()), Value::Number(value) if value == 3.0));

    let mut clear = changed
        .edit_cell_values(0)
        .expect("selector")
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
        .resolve()
        .expect("worksheet resolution")
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
            .resolve()
            .expect("worksheet resolution")
            .expect("worksheet");
        let error = edit
            .set_scalar(a1(), Scalar::Number(2.0))
            .expect_err("chain guard");
        assert!(error.to_string().contains("calculation chain"));
        let error = edit.remove(a1()).expect_err("removal chain guard");
        assert!(error.to_string().contains("calculation chain"));
        edit.insert_formula(
            b1(),
            style_zero(),
            Value::FormulaNumberCache(1.0),
            litchi_xlsb::cell_values::CellFormula::new(
                0,
                vec![0x1E, 1, 0],
                Vec::new(),
                litchi_xlsb::raw::Limits::DEFAULT,
            )
            .expect("formula"),
        )
        .expect("staged insertion");
        assert!(edit.remove(b1()).expect("cancel insertion"));
        assert!(edit.commit().expect("unchanged").patch().is_noop());
    }
}

#[test]
fn removal_is_idempotent_and_inverse_restores_exact_package() {
    let source = package_with_scalars();
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    assert!(!edit.remove(e1()).expect("absent coordinate"));
    assert!(edit.remove(a1()).expect("number"));
    assert!(!edit.remove(a1()).expect("repeated removal"));
    assert!(edit.remove(c1()).expect("last shared string"));
    let (after, patch) = edit.commit().expect("commit").into_parts();
    let workbook = after.workbook().expect("workbook");
    let cells = workbook.cell_values(0).expect("read");
    assert!(cells.cell(a1()).expect("lookup").is_none());
    assert!(cells.cell(c1()).expect("lookup").is_none());
    assert!(cells.cell(b1()).expect("neighbour").is_some());
    assert!(workbook.shared_strings().is_empty());
    let sst = PackURI::new("/xl/sharedStrings.bin").expect("URI");
    let part = after
        .opc_package()
        .get_part(&sst)
        .expect("retained SST envelope");
    let begin = Records::new(part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .next()
        .expect("SST begin")
        .expect("record");
    assert_eq!(&begin.payload()[..8], &[0; 8]);
    assert_eq!(source.to_bytes().expect("source remains"), before);
    assert_eq!(
        after
            .apply_cell_values(&patch.inverse())
            .expect("inverse")
            .package()
            .to_bytes()
            .expect("restored"),
        before
    );
}

#[test]
fn removal_of_missing_cells_and_cancelled_insertions_is_exact() {
    let source = package_with_scalars();
    let before = source.to_bytes().expect("bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    assert!(!edit.remove(e1()).expect("missing"));
    edit.set_scalar(e1(), Scalar::Number(9.0)).expect("insert");
    assert!(edit.remove(e1()).expect("remove insertion"));
    assert!(!edit.remove(e1()).expect("missing again"));
    let commit = edit.commit().expect("commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().expect("bytes"), before);
}

#[test]
fn removal_refuses_formula_group_members_including_absent_followers() {
    for source in [
        package_with_grouped_formula(),
        package_with_missing_group_follower(),
    ] {
        let before = source.to_bytes().expect("bytes");
        let mut edit = source
            .edit_cell_values(0)
            .expect("edit")
            .resolve()
            .expect("worksheet resolution")
            .expect("sheet");
        for reference in [a1(), b1()] {
            assert!(edit.remove(reference).is_err());
        }
        assert!(edit.commit().expect("unchanged commit").patch().is_noop());
        assert_eq!(source.to_bytes().expect("source"), before);
    }
}

#[test]
fn ordinary_formula_removal_drops_the_complete_record() {
    let source = package_with_typed_formula_caches();
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    for reference in [a1(), b1(), c1()] {
        assert!(edit.remove(reference).expect("formula removal"));
    }
    let after = edit.commit().expect("commit").into_package();
    let workbook = after.workbook().expect("workbook");
    let cells = workbook.cell_values(0).expect("read");
    for reference in [a1(), b1(), c1()] {
        assert!(cells.cell(reference).expect("lookup").is_none());
    }
}

#[test]
fn semantic_mutation_refuses_unmodeled_cell_families_without_guessing_coordinates() {
    let source = package_with_scalars();
    let mut opc = source.into_opc();
    let uri = PackURI::new("/xl/worksheets/sheet1.bin").expect("URI");
    let bytes = opc.get_part(&uri).expect("sheet").blob().to_vec();
    let mut output = Vec::new();
    let mut inserted = false;
    for item in Records::new(&bytes, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("record");
        Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT)
            .write_record(record.kind(), record.payload())
            .expect("copy");
        if !inserted && record.kind() == kind::ROW_HDR {
            Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT)
                .write_record(
                    litchi_xlsb::raw::Kind::new(0x000C).expect("reserved kind"),
                    &[0; 12],
                )
                .expect("reserved record");
            inserted = true;
        }
    }
    opc.get_part_mut(&uri).expect("sheet").set_blob(output);
    let source = Package::from_opc(opc).expect("opaque source");
    let before = source.to_bytes().expect("bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    assert!(matches!(
        edit.remove(e1()),
        Err(litchi_xlsb::package::PackageError::Refused(_))
    ));
    assert!(matches!(
        edit.set_scalar(e1(), Scalar::Blank),
        Err(litchi_xlsb::package::PackageError::Refused(_))
    ));
    assert_eq!(source.to_bytes().expect("unchanged"), before);
}
