//! Projected cell-state views for transaction-scoped worksheet edits.

use super::super::*;
use super::support::two_sheet_workbook;

use crate::formula::Formula;
use crate::{Cell, Error, FontChange, State, Value};
use litchi_opc::{BlobPart, PackURI, TargetMode};

#[test]
fn projected_state_distinguishes_missing_empty_payload_and_formula() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");

    assert!(matches!(
        sheet.state("A1").expect("missing state"),
        State::Missing
    ));

    sheet.clear("B1").expect("clear missing");
    assert_eq!(
        sheet.state("B1").expect("missing stays missing"),
        State::Missing
    );
    sheet.set("A1", "temporary").expect("set");
    assert!(matches!(
        sheet.state("A1").expect("set state"),
        State::Cell {
            content: Cell::Value(Value::Text(value)),
            shared_string: None,
            ..
        } if value.as_str() == "temporary"
    ));

    sheet.clear("A1").expect("clear");
    assert!(matches!(
        sheet.state("A1").expect("empty state"),
        State::Cell {
            content: Cell::Empty,
            shared_string: None,
            ..
        }
    ));

    sheet
        .set("A1", Formula::new("1+1").expect("formula"))
        .expect("formula set");
    assert!(matches!(
        sheet.state("A1").expect("formula state"),
        State::Cell {
            content: Cell::Formula(formula),
            shared_string: None,
            ..
        } if formula.text() == "1+1"
    ));
    let before = sheet.state("A1").expect("formula before repeat");
    sheet
        .set("A1", Formula::new("1+1").expect("same formula"))
        .expect("repeat");
    assert_eq!(before, sheet.state("A1").expect("exact repeated state"));
    sheet.remove("A1").expect("remove");
    assert_eq!(sheet.state("A1").expect("removed"), State::Missing);
}

#[test]
fn projected_state_refuses_deferred_style_lowering() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("font request");
    sheet
        .border(
            "B1",
            &crate::style::BorderChange::new().top_line(Some(crate::style::Line::Thin)),
        )
        .expect("border request");
    assert!(matches!(
        sheet.state("B1"),
        Err(Error::Unsupported {
            feature: "projecting cell state with pending style changes"
        })
    ));

    assert!(matches!(
        sheet.state("A1"),
        Err(Error::Unsupported {
            feature: "projecting cell state with pending style changes"
        })
    ));
}

#[test]
fn projected_state_retains_rich_shared_string_identity_for_staged_transfer() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    let shared_uri = PackURI::new("/xl/sharedStrings.xml").expect("shared strings URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            shared_uri,
            litchi_opc::constants::content_type::SML_SHARED_STRINGS.to_owned(),
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t> text</t></r></si></sst>"#.to_vec(),
        )))
        .expect("shared strings part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::SHARED_STRINGS.to_owned(),
            "sharedStrings.xml".to_owned(),
            "rIdSharedStrings".to_owned(),
            TargetMode::Internal,
        )
        .expect("shared strings relationship");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich-string workbook");

    let mut edit = source.edit().expect("edit");
    edit.copy_cells("Sheet1", "A1", "Sheet2", "B2")
        .expect("copy rich string");
    let sheet = edit.sheet("Sheet2").expect("lookup").expect("worksheet");
    assert!(matches!(
        sheet.state("B2").expect("projected transfer"),
        State::Cell {
            content: Cell::Value(Value::Text(value)),
            shared_string: Some(_),
            ..
        } if value.as_str() == "Rich text"
    ));
}
