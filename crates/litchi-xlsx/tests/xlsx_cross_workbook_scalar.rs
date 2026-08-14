#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::{Cell, DurablePatch, Error, Formula, Number, Value, Workbook};

const WORKBOOK_URI: &str = "/xl/workbook.xml";
const WORKSHEET_URI: &str = "/xl/worksheets/sheet1.xml";
const STYLES_URI: &str = "/xl/styles.xml";
const SHARED_STRINGS_URI: &str = "/xl/sharedStrings.xml";

fn workbook_with_sheet_xml(sheet: &str, shared_strings: Option<&str>) -> Workbook {
    let mut package = OpcPackage::new();
    let workbook = BlobPart::new(
        PackURI::new(WORKBOOK_URI).unwrap(),
        ct::SML_SHEET_MAIN.to_owned(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
    );
    package.try_add_part(Box::new(workbook)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(WORKSHEET_URI).unwrap(),
            ct::SML_WORKSHEET.to_owned(),
            sheet.as_bytes().to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(STYLES_URI).unwrap(),
            ct::SML_STYLES.to_owned(),
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.to_vec(),
        )))
        .unwrap();
    if let Some(shared_strings) = shared_strings {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SHARED_STRINGS_URI).unwrap(),
                ct::SML_SHARED_STRINGS.to_owned(),
                shared_strings.as_bytes().to_vec(),
            )))
            .unwrap();
    }
    let workbook = package
        .get_part_mut(&PackURI::new(WORKBOOK_URI).unwrap())
        .unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.to_owned(),
            "styles.xml".to_owned(),
            "rId2".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if shared_strings.is_some() {
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::SHARED_STRINGS.to_owned(),
                "sharedStrings.xml".to_owned(),
                "rId3".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.relate_to(WORKBOOK_URI.trim_start_matches('/'), rt::OFFICE_DOCUMENT);
    Workbook::from_bytes(PackageWriter::to_bytes(&package).unwrap()).unwrap()
}

fn workbook_with_scalar_values() -> Workbook {
    let workbook = Workbook::new().unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", 7_i32)
        .unwrap()
        .set("B1", true)
        .unwrap()
        .set("C1", Number::new("-0.000").unwrap())
        .unwrap()
        .set("D1", Value::date("2026-08-14").unwrap())
        .unwrap();
    edit.commit().unwrap().into_workbook()
}

#[test]
fn cross_workbook_scalar_copy_is_immutable_reversible_and_reopenable() {
    let donor = workbook_with_scalar_values();
    let donor_before = donor.to_plain_bytes().unwrap();
    let target = Workbook::new().unwrap();

    let mut edit = target.edit().unwrap();
    edit.copy_cells_from(&donor, "Sheet1", "A1:D1", "Sheet1", "F3")
        .unwrap();
    let commit = edit.commit().unwrap();
    let foreign_memory = Workbook::from_bytes(target.to_plain_bytes().unwrap()).unwrap();
    assert!(matches!(
        foreign_memory.apply(commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert!(commit.workbook().apply(&commit.patch().inverse()).is_ok());
    let durable = commit.patch().durable().unwrap();
    let changed = durable.apply(&target).unwrap();
    let sheet = changed.sheet("Sheet1").unwrap().unwrap();

    assert!(matches!(
        sheet.cell("F3").unwrap().stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "7"
    ));
    assert!(matches!(
        sheet.cell("G3").unwrap().stored(),
        Some(Cell::Value(Value::Bool(true)))
    ));
    assert!(matches!(
        sheet.cell("H3").unwrap().stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "-0.000"
    ));
    assert!(matches!(
        sheet.cell("I3").unwrap().stored(),
        Some(Cell::Value(Value::Date(date))) if date.as_str() == "2026-08-14"
    ));

    let reopened = Workbook::from_bytes(changed.to_plain_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("H3")
            .unwrap()
            .stored()
            .is_some()
    );
    assert_eq!(donor.to_plain_bytes().unwrap(), donor_before);

    let restored = durable.inverse().apply(&changed).unwrap();
    assert_eq!(
        restored.to_plain_bytes().unwrap(),
        target.to_plain_bytes().unwrap()
    );
    let foreign =
        DurablePatch::from_deterministic_json(&durable.to_deterministic_json().unwrap()).unwrap();
    assert!(matches!(
        foreign.apply(&donor),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn cross_workbook_scalar_copy_refuses_dependencies_and_occupied_targets_atomically() {
    let donor = workbook_with_scalar_values();
    let target = Workbook::new().unwrap();

    let mut occupied = target.edit().unwrap();
    occupied
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("F3", 99_i32)
        .unwrap();
    assert!(matches!(
        occupied.copy_cells_from(&donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));
    let still_original = occupied.commit().unwrap().into_workbook();
    assert!(matches!(
        still_original
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("F3")
            .unwrap()
            .stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "99"
    ));

    let mut formula_edit = Workbook::new().unwrap().edit().unwrap();
    formula_edit
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", Formula::new("1+1").unwrap())
        .unwrap();
    let formula_donor = formula_edit.commit().unwrap().into_workbook();
    let mut formula_copy = target.edit().unwrap();
    assert!(matches!(
        formula_copy.copy_cells_from(&formula_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));

    let style_donor = {
        let style_base = Workbook::new().unwrap();
        let style = style_base.styles().unwrap().base().unwrap();
        let mut edit = style_base.edit().unwrap();
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .set("A1", 1_i32)
            .unwrap()
            .style("A1", &style)
            .unwrap();
        edit.commit().unwrap().into_workbook()
    };
    let mut style_copy = target.edit().unwrap();
    assert!(matches!(
        style_copy.copy_cells_from(&style_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));

    let shared_donor = workbook_with_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#,
        Some(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>shared donor text</t></si></sst>"#,
        ),
    );
    let mut shared_copy = target.edit().unwrap();
    assert!(matches!(
        shared_copy.copy_cells_from(&shared_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));

    let rich_inline_donor = workbook_with_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>rich donor text</t></r></is></c></row></sheetData></worksheet>"#,
        None,
    );
    let mut rich_inline_copy = target.edit().unwrap();
    assert!(matches!(
        rich_inline_copy.copy_cells_from(&rich_inline_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));

    let hyperlink_donor = workbook_with_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/><hyperlinks><hyperlink ref="A1" location="Sheet1!A1"/></hyperlinks></worksheet>"#,
        None,
    );
    let mut hyperlink_copy = target.edit().unwrap();
    assert!(matches!(
        hyperlink_copy.copy_cells_from(&hyperlink_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));

    for (index, sheet) in [
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" future="1"><v>1</v></c></row></sheetData></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v><future/></c></row></sheetData></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><drawing/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><legacyDrawing/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><oleObjects/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><controls/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetProtection sheet="1"/><sheetData/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext uri="urn:future"/></extLst></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback/></mc:AlternateContent><sheetData/></worksheet>"#,
    ]
    .into_iter()
    .enumerate()
    {
        let hostile_donor = workbook_with_sheet_xml(sheet, None);
        let mut hostile_copy = target.edit().unwrap();
        let result =
            hostile_copy.copy_cells_from(&hostile_donor, "Sheet1", "A1", "Sheet1", "F3");
        assert!(
            matches!(&result, Err(Error::Unsupported { .. })),
            "hostile worksheet fixture {index} returned {result:?}"
        );
    }

    let merged_donor = {
        let mut edit = Workbook::new().unwrap().edit().unwrap();
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .merge("A1:B1")
            .unwrap()
            .set("A1", 1_i32)
            .unwrap();
        edit.commit().unwrap().into_workbook()
    };
    let mut merged_copy = target.edit().unwrap();
    assert!(matches!(
        merged_copy.copy_cells_from(&merged_donor, "Sheet1", "A1", "Sheet1", "F3"),
        Err(Error::Unsupported { .. })
    ));
}

#[test]
fn cross_workbook_scalar_copy_refuses_formula_owned_ranges_beyond_selected_cells() {
    let target = Workbook::new().unwrap();
    let scalar_donor = workbook_with_scalar_values();

    let source_fixtures = [
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B1">1</f><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="dataTable" ref="A1:B1"/><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:B1">1</f><v>1</v></c><c r="B1"><f t="shared" si="0"/><v>2</v></c></row></sheetData></worksheet>"#,
    ];
    for (index, source_xml) in source_fixtures.into_iter().enumerate() {
        let donor = workbook_with_sheet_xml(source_xml, None);
        let mut edit = target.edit().unwrap();
        let result = edit.copy_cells_from(&donor, "Sheet1", "B1", "Sheet1", "F3");
        assert!(
            matches!(&result, Err(Error::Unsupported { .. })),
            "source formula-owned fixture {index} returned {result:?}"
        );
    }

    let target_array = workbook_with_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B1">1</f><v>1</v></c></row></sheetData></worksheet>"#,
        None,
    );
    let mut target_edit = target_array.edit().unwrap();
    let result = target_edit.copy_cells_from(&scalar_donor, "Sheet1", "A1", "Sheet1", "B1");
    assert!(matches!(result, Err(Error::Unsupported { .. })));
}

#[test]
fn same_workbook_copy_from_delegates_to_existing_transfer_semantics() {
    let workbook = Workbook::new().unwrap();
    let style = workbook.styles().unwrap().base().unwrap();
    let mut prepare = workbook.edit().unwrap();
    prepare
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", Formula::new("B1+1").unwrap())
        .unwrap()
        .style("A1", &style)
        .unwrap();
    let prepared = prepare.commit().unwrap().into_workbook();

    let mut edit = prepared.edit().unwrap();
    edit.copy_cells_from(&prepared, "Sheet1", "A1", "Sheet1", "D4")
        .unwrap();
    let copied = edit.commit().unwrap().into_workbook();
    assert!(matches!(
        copied
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("D4")
            .unwrap()
            .stored(),
        Some(Cell::Formula(formula)) if formula.text() == "E4+1"
    ));
}

#[test]
fn cross_workbook_scalar_copies_join_when_targets_are_disjoint_and_conflict_when_overlapping() {
    let donor = workbook_with_scalar_values();
    let target = Workbook::new().unwrap();

    let mut left = target.edit().unwrap();
    left.copy_cells_from(&donor, "Sheet1", "A1", "Sheet1", "F3")
        .unwrap();
    let mut right = target.edit().unwrap();
    right
        .copy_cells_from(&donor, "Sheet1", "B1", "Sheet1", "H3")
        .unwrap();
    left.join(right).unwrap();
    let joined = left.commit().unwrap().into_workbook();
    assert!(
        joined
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("F3")
            .unwrap()
            .stored()
            .is_some()
    );
    assert!(
        joined
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("H3")
            .unwrap()
            .stored()
            .is_some()
    );

    let mut conflict_left = target.edit().unwrap();
    conflict_left
        .copy_cells_from(&donor, "Sheet1", "A1", "Sheet1", "F3")
        .unwrap();
    let mut conflict_right = target.edit().unwrap();
    conflict_right
        .copy_cells_from(&donor, "Sheet1", "C1", "Sheet1", "F3")
        .unwrap();
    assert!(conflict_left.join(conflict_right).is_err());
}
