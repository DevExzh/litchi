//! A generated workbook for the XLSX allocation test and timing probe
//! (`tests/xlsx_query_allocations.rs`, `examples/xlsx_unified_bench.rs`).
//!
//! One worksheet, `Data`: a bold header row (`Region`, `Sales`, `Status`)
//! and `rows` data rows under a `Sales` table, written as Excel writes them
//! (shared strings, a styles part, a dimension). The data repeat with
//! period 4, so a workbook of any size has the same proportion of matches.

use std::io::{Cursor, Write as _};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml";

const STRINGS: [&str; 9] = [
    "Region", "Sales", "Status", "West", "East", "North", "South", "open", "closed",
];

/// The `Sales` values, repeating: half of them exceed 1000.
const SALES: [&str; 4] = ["1200", "800", "450.5", "2000"];

/// The package bytes of a workbook with `rows` data rows.
pub(crate) fn workbook(rows: usize) -> Vec<u8> {
    let last = rows + 1;
    let mut sheet = String::with_capacity(160 * rows + 1024);
    sheet.push_str(&format!(
        r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><dimension ref="A1:C{last}"/><sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="1" t="s"><v>2</v></c></row>"#
    ));
    for index in 0..rows {
        let row = index + 2;
        sheet.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}" t="s"><v>{}</v></c><c r="B{row}"><v>{}</v></c><c r="C{row}" t="s"><v>{}</v></c></row>"#,
            3 + index % 4,
            SALES[index % 4],
            7 + index % 2,
        ));
    }
    sheet.push_str(
        r#"</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#,
    );

    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="{CT}.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="{CT}.table+xml"/></Types>"#
    );
    let root = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    );
    let workbook_rels = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
    );
    let sheet_rels = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/table" Target="../tables/table1.xml"/></Relationships>"#
    );
    let table = format!(
        r#"<table xmlns="{SML}" id="1" name="Sales" displayName="Sales" ref="A1:C{last}"><autoFilter ref="A1:C{last}"/><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Sales"/><tableColumn id="3" name="Status"/></tableColumns></table>"#
    );
    let strings: String = STRINGS
        .iter()
        .map(|text| format!("<si><t>{text}</t></si>"))
        .collect();
    let shared = format!(
        r#"<sst xmlns="{SML}" count="{0}" uniqueCount="{0}">{strings}</sst>"#,
        STRINGS.len()
    );
    let styles = format!(
        r#"<styleSheet xmlns="{SML}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"#
    );

    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/sharedStrings.xml", shared.as_bytes()),
        ("xl/styles.xml", styles.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels.as_bytes()),
        ("xl/tables/table1.xml", table.as_bytes()),
    ] {
        writer.start_file(name, options).expect("start a member");
        writer.write_all(bytes).expect("write a member");
    }
    writer.finish().expect("finish the archive");
    output.into_inner()
}
