//! Tests of the XLSX adapter (M1 step R4b, record 0010 §4.7, §5.4, §6, §9.2):
//! path resolution, OfficeCLI-shaped node views (golden JSON), queries and
//! header rows, the view modes, decode counts and the path round trip.

use std::fs;
use std::io::{Cursor, Write as _};
use std::path::PathBuf;

use super::{Nav, Xlsx};
use crate::error::Result;
use crate::node::{Node, ViewOptions};
use crate::options::{OpenOptions, SourceMode};
use crate::path::{IntoPath, resolve};
use crate::view::mode::{Mode, Options, SheetRange};

mod decode;
mod paths;
mod policy;
mod queries;
mod round_trip;
mod views;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_PREFIX: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml";

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/")).join(name)
}

/// Opens a fixture through the unified opener with `mode`.
fn open_path_with(name: &str, mode: SourceMode) -> Xlsx {
    let options = OpenOptions {
        source_mode: mode,
        ..OpenOptions::default()
    };
    match crate::open_with(fixture(name), &options).expect("opens") {
        crate::File::Xlsx(xlsx) => xlsx,
        other => panic!("{name} opened as {other:?}"),
    }
}

fn open_path(name: &str) -> Xlsx {
    open_path_with(name, SourceMode::Auto)
}

/// Opens package bytes with `mode`: the source-backed owner, or the eager
/// snapshot.
fn open_bytes_with(bytes: Vec<u8>, mode: SourceMode) -> Xlsx {
    let options = OpenOptions {
        source_mode: mode,
        ..OpenOptions::default()
    };
    match crate::open_bytes_with(bytes, &options).expect("opens") {
        crate::File::Xlsx(xlsx) => xlsx,
        other => panic!("opened as {other:?}"),
    }
}

/// Both owner states of the same package.
fn both(bytes: &[u8]) -> [Xlsx; 2] {
    let source = open_bytes_with(bytes.to_vec(), SourceMode::Auto);
    let eager = open_bytes_with(bytes.to_vec(), SourceMode::Eager);
    assert!(source.is_source_backed() && !eager.is_source_backed());
    [source, eager]
}

fn get<'f>(file: &'f Xlsx, path: &str) -> Result<Node<'f>> {
    let path = path.into_path()?;
    resolve::get(&Nav::new(file), &path).map(Node::Xlsx)
}

fn find<'f>(file: &'f Xlsx, path: &str) -> Result<Option<Node<'f>>> {
    let path = path.into_path()?;
    Ok(resolve::find(&Nav::new(file), &path)?.map(Node::Xlsx))
}

fn path_of(node: &Node<'_>) -> String {
    node.path().to_string()
}

fn query<'f>(file: &'f Xlsx, selector: &str) -> Result<Vec<Node<'f>>> {
    let nav = Nav::new(file);
    let walk = crate::query::run(&nav, selector)?;
    walk.map(|node| node.map(Node::Xlsx)).collect()
}

fn query_paths(file: &Xlsx, selector: &str) -> Vec<String> {
    query(file, selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .iter()
        .map(path_of)
        .collect()
}

fn json(node: &Node<'_>, options: ViewOptions) -> String {
    let mut sink = Vec::new();
    node.view_with(options)
        .serialize_into(&mut serde_json::Serializer::pretty(&mut sink))
        .expect("serializes");
    String::from_utf8(sink).expect("UTF-8")
}

fn check_golden(name: &str, actual: &str) {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden/node_view/xlsx")
        .join(format!("{name}.json"));
    let actual = format!("{actual}\n");
    if std::env::var_os("LITCHI_BLESS").is_some() {
        fs::create_dir_all(path.parent().expect("a directory"))
            .expect("create the golden directory");
        fs::write(&path, &actual).expect("bless the golden file");
        return;
    }
    let expected = fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("read {}: {error}", path.display()));
    assert_eq!(
        actual, expected,
        "{name}: the node view differs from the OfficeCLI-shaped golden file (LITCHI_BLESS=1 rewrites it)"
    );
}

fn view(file: &Xlsx, mode: Mode, options: &Options) -> String {
    let mut out = String::new();
    super::view_to(file, mode, options, &mut out).expect("renders");
    out
}

fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in entries {
        writer.start_file(*name, options).expect("start member");
        writer.write_all(bytes).expect("write member");
    }
    writer.finish().expect("finish archive");
    output.into_inner()
}

/// One worksheet of a generated workbook.
struct SheetPart {
    name: &'static str,
    state: Option<&'static str>,
    xml: String,
    table: Option<String>,
}

/// A workbook package with shared strings, styles, core properties, the
/// given defined names and sheets (each optionally owning one table).
fn workbook(sheets: &[SheetPart], defined_names: &str, strings: &[&str], styles: &str) -> Vec<u8> {
    let mut overrides = String::new();
    let mut entries = String::new();
    let mut relationships = String::new();
    let mut parts: Vec<(String, String)> = Vec::new();
    for (index, sheet) in sheets.iter().enumerate() {
        let number = index + 1;
        overrides.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType="{CT_PREFIX}.worksheet+xml"/>"#
        ));
        let state = sheet
            .state
            .map_or(String::new(), |state| format!(r#" state="{state}""#));
        entries.push_str(&format!(
            r#"<sheet name="{}" sheetId="{number}"{state} r:id="rId{number}"/>"#,
            sheet.name
        ));
        relationships.push_str(&format!(
            r#"<Relationship Id="rId{number}" Type="{REL}/worksheet" Target="worksheets/sheet{number}.xml"/>"#
        ));
        parts.push((
            format!("xl/worksheets/sheet{number}.xml"),
            sheet.xml.clone(),
        ));
        if let Some(table) = &sheet.table {
            overrides.push_str(&format!(
                r#"<Override PartName="/xl/tables/table{number}.xml" ContentType="{CT_PREFIX}.table+xml"/>"#
            ));
            parts.push((
                format!("xl/worksheets/_rels/sheet{number}.xml.rels"),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/table" Target="../tables/table{number}.xml"/></Relationships>"#
                ),
            ));
            parts.push((format!("xl/tables/table{number}.xml"), table.clone()));
        }
    }
    let count = sheets.len();
    relationships.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId{}" Type="{REL}/styles" Target="styles.xml"/>"#,
        count + 1,
        count + 2
    ));
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{CT_PREFIX}.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="{CT_PREFIX}.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="{CT_PREFIX}.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>{overrides}</Types>"#
    );
    let root = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="{PACKAGE_REL}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#
    );
    let core = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Kitchen sink</dc:title><dc:creator>Litchi</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-05T08:00:00Z</dcterms:created></cp:coreProperties>"#;
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>{entries}</sheets>{defined_names}</workbook>"#
    );
    let workbook_relationships =
        format!(r#"<Relationships xmlns="{PACKAGE_REL}">{relationships}</Relationships>"#);
    let shared: String = strings
        .iter()
        .map(|text| format!("<si><t>{text}</t></si>"))
        .collect();
    let shared_strings = format!(
        r#"<sst xmlns="{SML}" count="{0}" uniqueCount="{0}">{shared}</sst>"#,
        strings.len()
    );
    let mut all: Vec<(&str, &[u8])> = vec![
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root.as_bytes()),
        ("docProps/core.xml", core.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        ),
        ("xl/sharedStrings.xml", shared_strings.as_bytes()),
        ("xl/styles.xml", styles.as_bytes()),
    ];
    for (name, xml) in &parts {
        all.push((name.as_str(), xml.as_bytes()));
    }
    zip(&all)
}

const STRINGS: &[&str] = &[
    "Region", "Sales", "Status", "West", "East", "North", "South", "open", "closed", "Merged",
    "Name", "Score", "Ann", "Bob",
];

fn styles() -> String {
    format!(
        r#"<styleSheet xmlns="{SML}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts><fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><i/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font><font><strike/><vertAlign val="superscript"/><sz val="10.5"/><color theme="4" tint="0.4"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="lightGray"><fgColor theme="4" tint="0.4"/><bgColor rgb="FF0000FF"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border diagonalUp="1"><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="medium"/><bottom style="double"/><diagonal style="thin"/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="2" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyAlignment="1" applyProtection="1" quotePrefix="1"><alignment textRotation="45" indent="2" readingOrder="2" shrinkToFit="1"/><protection locked="0" hidden="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
    )
}

/// The sheet `Data`: a `Sales` table over A1:C5 with formatted cells, every
/// value kind, a merge and formulas; a hidden row and two column records.
fn data_sheet() -> SheetPart {
    let xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><dimension ref="A1:F7"/><cols><col min="1" max="1" width="12.5" customWidth="1"/><col min="3" max="3" width="9" hidden="1"/></cols><sheetData><row r="1" ht="20" customHeight="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="1" t="s"><v>2</v></c><c r="E1" t="s"><v>9</v></c></row><row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2" s="2"><v>1200</v></c><c r="C2" t="s"><v>7</v></c><c r="D2" s="2"/><c r="F2"/></row><row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3"><v>800</v></c><c r="C3" t="inlineStr"><is><t>closed</t></is></c><c r="D3" t="b"><v>1</v></c></row><row r="4"><c r="A4" t="s"><v>5</v></c><c r="B4" s="3"><v>450.5</v></c><c r="C4" t="s"><v>7</v></c><c r="D4" t="e"><v>#DIV/0!</v></c></row><row r="5"><c r="A5" t="s"><v>6</v></c><c r="B5" s="4"><v>2000</v></c><c r="C5" t="s"><v>8</v></c><c r="D5" t="str"><f>A5&amp;"!"</f><v>South!</v></c></row><row r="6"><c r="B6"><f>SUM(B2:B5)</f><v>4450.5</v></c><c r="C6"><f>B6*2</f></c><c r="D6"><f t="array" ref="D6">MAX(B2:B5)</f><v>2000</v></c></row><row r="7" hidden="1" outlineLevel="1"><c r="A7" t="s"><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="E1:F2"/></mergeCells><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#
    );
    let table = format!(
        r#"<table xmlns="{SML}" id="1" name="Sales" displayName="Sales" ref="A1:C5" totalsRowShown="0"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Sales"/><tableColumn id="3" name="Status"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>"#
    );
    SheetPart {
        name: "Data",
        state: None,
        xml,
        table: Some(table),
    }
}

/// The hidden sheet `Notes`: a header-row block without a table (B2:C4).
fn notes_sheet() -> SheetPart {
    SheetPart {
        name: "Notes",
        state: Some("hidden"),
        xml: format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="2"><c r="B2" t="s"><v>10</v></c><c r="C2" t="s"><v>11</v></c></row><row r="3"><c r="B3" t="s"><v>12</v></c><c r="C3"><v>90</v></c></row><row r="4"><c r="B4" t="s"><v>13</v></c><c r="C4"><v>75</v></c></row></sheetData></worksheet>"#
        ),
        table: None,
    }
}

fn small_sheet(name: &'static str, xml: &str) -> SheetPart {
    SheetPart {
        name,
        state: None,
        xml: format!(r#"<worksheet xmlns="{SML}"><sheetData>{xml}</sheetData></worksheet>"#),
        table: None,
    }
}

/// The kitchen-sink workbook: `Data`, hidden `Notes`, `My Data` and
/// `Ärger`, three defined names.
fn kitchen() -> Vec<u8> {
    workbook(
        &[
            data_sheet(),
            notes_sheet(),
            small_sheet(
                "My Data",
                r#"<row r="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="B2" t="s"><v>4</v></c></row>"#,
            ),
            small_sheet(
                "Ärger",
                r#"<row r="1"><c r="A1" t="inlineStr"><is><t>ä</t></is></c></row>"#,
            ),
        ],
        r#"<definedNames><definedName name="Local" localSheetId="0" comment="sheet scoped">Data!$A$1</definedName><definedName name="Total">Data!$B$2:$B$5</definedName><definedName name="Volatile" function="1">NOW()</definedName></definedNames>"#,
        STRINGS,
        &styles(),
    )
}

#[test]
fn kitchen_sink_opens_on_both_owners() {
    for file in both(&kitchen()) {
        assert_eq!(
            file.worksheet_names().expect("names"),
            ["Data", "Notes", "My Data", "Ärger"]
        );
    }
}

#[test]
fn plans_compile_for_xlsx_and_run_on_any_workbook() {
    let query = crate::query::Query::parse("cell[type=Number]").expect("parses");
    let plan =
        crate::query::Plan::compile(&query, litchi_core::FileFormat::Xlsx).expect("compiles");
    assert_eq!(plan.format(), litchi_core::FileFormat::Xlsx);
    for file in both(&kitchen()) {
        let nav = Nav::new(&file);
        let walk =
            crate::query::walk::Walk::new(&nav, crate::query::walk::PlanRef::Borrowed(&plan))
                .expect("runs");
        let paths: Vec<String> = walk
            .map(|node| Node::Xlsx(node.expect("reads")).path().to_string())
            .collect();
        assert_eq!(
            paths,
            [
                "/sheet[1]/B2",
                "/sheet[1]/D2",
                "/sheet[1]/F2",
                "/sheet[1]/B3",
                "/sheet[1]/B4",
                "/sheet[1]/B5",
                "/sheet[1]/B6",
                "/sheet[1]/C6",
                "/sheet[1]/D6",
                "/sheet[2]/C3",
                "/sheet[2]/C4",
                "/sheet[3]/A1"
            ]
        );
    }
    // Pages body tables (record 0299) and Keynote slide tables (record 0298)
    // have cells: the plan compiles for them too.
    let other = crate::query::Plan::compile(&query, litchi_core::FileFormat::Keynote)
        .expect("Keynote slide-table cells");
    assert_eq!(other.format(), litchi_core::FileFormat::Keynote);
}

#[test]
fn view_options_reach_the_view() {
    // A smoke test of the shared mode types through the XLSX views.
    let file = &both(&kitchen())[0];
    let options = Options {
        range: Some(SheetRange::parse("data!A1:B2").expect("a range")),
        ..Options::default()
    };
    assert_eq!(
        view(file, Mode::Text, &options),
        "=== Sheet: Data ===\n[/sheet[1]/row[1]] A1=Region\tB1=Sales\n[/sheet[1]/row[2]] A2=West\tB2=1200"
    );
}
