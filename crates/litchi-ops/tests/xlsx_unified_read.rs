//! Public `File` dispatch for the XLSX unified read adapter (M1 R4b).
#![cfg(feature = "xlsx")]

use std::io::{Cursor, Write as _};

use litchi_ops::query::Plan;
use litchi_ops::view::Mode;
use litchi_ops::{File, OpenOptions, SourceMode};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn fixture() -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#;
    let root_rels = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="{PACKAGE_REL}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    );
    let workbook_rels = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
    );
    let sheet = format!(
        r#"<worksheet xmlns="{MAIN}"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Ada</t></is></c><c r="B2"><v>42</v></c></row></sheetData></worksheet>"#
    );
    let styles = format!(
        r#"<styleSheet xmlns="{MAIN}"><numFmts count="0"/><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
    );
    let core = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Unified XLSX</dc:title></cp:coreProperties>"#;

    let mut output = Cursor::new(Vec::new());
    let mut archive = zip::ZipWriter::new(&mut output);
    let options = zip::write::SimpleFileOptions::default();
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
        ("xl/styles.xml", styles.as_bytes()),
        ("docProps/core.xml", core.as_bytes()),
    ] {
        archive
            .start_file(name, options)
            .expect("start package part");
        archive.write_all(contents).expect("write package part");
    }
    archive.finish().expect("finish package");
    output.into_inner()
}

fn files() -> [File; 2] {
    let bytes = fixture();
    [
        litchi_ops::open_bytes(bytes.clone()).expect("source-backed XLSX"),
        litchi_ops::open_bytes_with(
            bytes,
            &OpenOptions::default().source_mode(SourceMode::Eager),
        )
        .expect("eager XLSX"),
    ]
}

#[test]
fn file_dispatches_xlsx_root_get_query_and_view() {
    for file in files() {
        assert_eq!(file.format(), Some(litchi_core::FileFormat::Xlsx));
        assert_eq!(file.root().expect("root").path().to_string(), "/");

        let cell = file.get("/data/a2").expect("canonical cell");
        assert_eq!(cell.path().to_string(), "/sheet[1]/A2");
        assert_eq!(cell.kind(), "cell");
        assert_eq!(
            cell.view().text().expect("cell text").as_deref(),
            Some("Ada")
        );
        assert_eq!(
            file.get("Data!b2").expect("A1 notation").path().to_string(),
            "/sheet[1]/B2"
        );
        assert!(file.find("/DATA/B3").expect("find").is_some());

        let query = file.query("cell[value=42]").expect("query");
        let matches = query
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("matches");
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].path().to_string(), "/sheet[1]/B2");

        let text = file.view(Mode::Text).expect("text view");
        assert!(text.contains("=== Sheet: Data ==="));
        assert!(text.contains("A2=Ada"));
        let json = file.view_json(Mode::Outline).expect("outline JSON");
        assert!(
            serde_json::to_string(&json)
                .expect("serialize view")
                .contains("Data")
        );
    }
}

#[test]
fn xlsx_plans_compile_through_the_public_query_entry_point() {
    let plan = Plan::compile(
        &litchi_ops::query::Query::parse("cell[type=Number]").expect("query"),
        litchi_core::FileFormat::Xlsx,
    )
    .expect("XLSX plan");
    for file in files() {
        let matches = file
            .query(&plan)
            .expect("query plan")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("matches");
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].path().to_string(), "/sheet[1]/B2");
    }
}
