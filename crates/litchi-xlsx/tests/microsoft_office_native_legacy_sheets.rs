//! Native macro-sheet metadata stays inert during ordinary worksheet edits.
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-legacy-sheets.xlsm"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

use litchi_xlsx::{Error, Package, WorksheetKind};
#[test]
fn native_macro_tab_and_explicit_nonworksheet_access_boundary() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let wb = p.workbook();
    assert_eq!(
        wb.sheets()
            .map(|s| s.name().to_string())
            .collect::<Vec<_>>(),
        ["Macro_016", "Sheet1"]
    );
    assert_eq!(wb.flavor(), litchi_xlsx::Flavor::MacroWorkbook);
    assert_eq!(wb.active_sheet().unwrap().name(), "Macro_016");
    let macro_sheet = wb.sheet("Macro_016").unwrap().unwrap();
    assert_eq!(macro_sheet.kind(), WorksheetKind::Macro);
    assert!(
        matches!(macro_sheet.cell("A1"),Err(Error::NotWorksheet {sheet}) if sheet=="Macro_016")
    );
    assert!(matches!(macro_sheet.extents(),Err(Error::NotWorksheet {sheet}) if sheet=="Macro_016"));
    assert!(
        matches!(macro_sheet.cells("A1:A3"),Err(Error::NotWorksheet {sheet}) if sheet=="Macro_016")
    );
    let sheet = wb.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(sheet.kind(), WorksheetKind::Worksheet);
    assert_eq!(
        sheet.cell("A1").unwrap().stored(),
        Some(&litchi_xlsx::Cell::Value(litchi_xlsx::Value::text(
            "EXCEL_LEGACY_SHEETS_016"
        )))
    );
    let parts = members(&bytes);
    assert_eq!(parts["xl/macrosheets/sheet1.xml"].len(), 1407);
    assert_eq!(
        hash(&parts["xl/macrosheets/sheet1.xml"]),
        "a913330c3e94155894d6db87cd5ee24db30f6bdc6ed742e7303972ec24258e6b"
    );
    assert!(!parts.keys().any(|name| name.contains("vbaProject")));
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_ordinary_worksheet_edit_keeps_macro_part_and_workbook_kind() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("D4", "Reviewed ordinary worksheet")
        .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsm").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened.workbook().flavor(),
        litchi_xlsx::Flavor::MacroWorkbook
    );
    assert_eq!(
        reopened.workbook().active_sheet().unwrap().kind(),
        WorksheetKind::Macro
    );
    assert_eq!(
        reopened
            .workbook()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("D4")
            .unwrap()
            .stored(),
        Some(&litchi_xlsx::Cell::Value(litchi_xlsx::Value::text(
            "Reviewed ordinary worksheet"
        )))
    );
    let original = members(&bytes);
    let changed = members(&after);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    let diffs = original
        .iter()
        .filter(|(n, b)| changed[*n] != **b)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(diffs, ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]);
    assert_eq!(
        workbook_events_without_calculation_properties(&original["xl/workbook.xml"]),
        workbook_events_without_calculation_properties(&changed["xl/workbook.xml"])
    );
    let calculation = reopened.workbook().calculation_metadata().unwrap();
    let calculation = calculation.properties().unwrap();
    assert_eq!(calculation.calculation_id(), 0);
    assert!(calculation.full_calculation_on_load());
    assert!(calculation.force_full_calculation());
    assert!(changed["xl/macrosheets/sheet1.xml"] == original["xl/macrosheets/sheet1.xml"]);
    let restored = reopened.apply(&commit.patch().inverse()).unwrap();
    let restored = members(&restored.package().to_bytes().unwrap());
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        restored.keys().collect::<Vec<_>>()
    );
    let inverse_diff = original
        .iter()
        .filter(|(n, b)| restored[*n] != **b)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert!(inverse_diff.is_empty(), "inverse changed {inverse_diff:?}");
}

fn hash(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

fn workbook_events_without_calculation_properties(bytes: &[u8]) -> Vec<Vec<u8>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut events = Vec::new();
    loop {
        let event = reader.read_event().unwrap();
        match &event {
            Event::Eof => break,
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Empty(element) if element.local_name().as_ref() == b"calcPr" => {
                let mut kept = std::collections::BTreeMap::new();
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if !matches!(
                        attribute.key.as_ref(),
                        b"calcId"
                            | b"fullCalcOnLoad"
                            | b"forceFullCalc"
                            | b"calcCompleted"
                            | b"calcOnSave"
                    ) {
                        kept.insert(attribute.key.as_ref().to_vec(), attribute.value.to_vec());
                    }
                }
                if !kept.is_empty() {
                    events.push(format!("calcPr:{kept:?}").into_bytes());
                }
            },
            _ => events.push(event.as_ref().to_vec()),
        }
    }
    events
}
