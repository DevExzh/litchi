use litchi_opc::PackURI;
use litchi_xlsx::{
    Package, parse_conditional_formattings, parse_page_margins, parse_print_options,
    parse_worksheet_page_setup,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn worksheet_uri() -> PackURI {
    PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI")
}

fn assert_compact_xml(xml: &[u8]) {
    assert!(
        !xml.iter()
            .any(|byte| matches!(*byte, b'\n' | b'\r' | b'\t'))
    );
    assert!(!xml.windows(3).any(|window| window == b"> <"));
}

#[test]
fn documented_read_only_worksheet_projections_parse_compact_source_without_rewriting_it() {
    let xml = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><printOptions gridLines="1" gridLinesSet="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape"/></worksheet>"#
    )
    .into_bytes();
    let before = xml.clone();

    let conditional = parse_conditional_formattings(&xml, 0).expect("conditional formatting");
    assert_eq!(conditional.len(), 1);
    assert!(parse_print_options(&xml).expect("print options").is_some());
    assert!(parse_page_margins(&xml).expect("page margins").is_some());
    assert!(
        parse_worksheet_page_setup(&xml)
            .expect("page setup")
            .is_some()
    );
    assert_eq!(xml, before);
    assert_compact_xml(&xml);
}

#[test]
fn unowned_conditional_hyperlink_and_page_break_xml_stays_byte_exact_through_save() {
    let worksheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><dimension ref="A1"/><sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><hyperlinks><hyperlink ref="A1" location="Sheet1!A1" display="local"/></hyperlinks><rowBreaks count="1" manualBreakCount="1"><brk id="1" min="0" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="1" min="0" max="1048575" man="1"/></colBreaks></worksheet>"#
    )
    .into_bytes();
    assert_compact_xml(&worksheet);

    let uri = worksheet_uri();
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    raw.get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(worksheet.clone());
    let package = Package::from_opc(raw).expect("validated package");
    let reopened = Package::from_bytes(package.to_bytes().expect("serialized package"))
        .expect("reopened package")
        .into_plain_opc();

    assert_eq!(
        reopened.get_part(&uri).expect("reopened worksheet").blob(),
        worksheet
    );
}

#[test]
fn defined_names_are_inert_catalog_records_and_custom_props_use_the_package_facade() {
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    let workbook_uri = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let workbook_xml = std::str::from_utf8(
        raw.get_part(&workbook_uri)
            .expect("workbook part")
            .blob(),
    )
    .expect("workbook UTF-8")
    .replace(
        "</sheets>",
        "</sheets><definedNames><definedName name=\"Rate\" comment=\"inert\">Sheet1!$A$1</definedName></definedNames>",
    )
    .into_bytes();
    assert_compact_xml(&workbook_xml);
    raw.get_part_mut(&workbook_uri)
        .expect("mutable workbook part")
        .set_blob(workbook_xml);

    let package = Package::from_opc(raw).expect("defined-name package");
    let workbook = package.workbook().expect("workbook snapshot");
    assert_eq!(workbook.defined_names().len(), 1);
    assert_eq!(workbook.defined_names()[0].name, "Rate");
    assert_eq!(workbook.defined_names()[0].reference, "Sheet1!$A$1");
    assert_eq!(
        workbook.defined_names()[0].comment.as_deref(),
        Some("inert")
    );
    assert!(
        package
            .custom_props()
            .expect("custom properties")
            .is_empty()
    );
}
