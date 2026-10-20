//! Unified-read tests of the DOCX adapter (M1 step R4a): paths, the node
//! view against OfficeCLI's `DocumentNode` shapes, queries, views, decode
//! counts, the transaction ordinal of record 0010 §4.5.1, and a path round
//! trip over the DOCX corpus.

use std::fmt::Write as _;
use std::fs;
use std::io::{Cursor, Write as _};
use std::path::{Path, PathBuf};

use super::{Docx, DocxOwner, Nav};
use crate::error::{Code, Result};
use crate::node::{Node, ViewOptions};
use crate::path::{IntoPath, resolve};
use crate::query::walk::{PlanRef, Walk};
use crate::query::{Plan, Query};
use crate::view::mode::{Mode, Options};
use crate::{File, OpenOptions, SourceMode};
use litchi_core::FileFormat;

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

fn test_data() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

fn open(name: &str) -> File {
    crate::open(test_data().join(name)).unwrap_or_else(|error| panic!("open {name}: {error}"))
}

fn docx(file: &File) -> &Docx {
    match file {
        File::Docx(docx) => docx,
        other => panic!("expected DOCX, got {other:?}"),
    }
}

const MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/// A DOCX package holding `document` as its main part, with `relationships`
/// as the main part's relationships when given.
fn package(document: &str, relationships: Option<&str>) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{MAIN}"/></Types>"#
    );
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    writer.start_file("[Content_Types].xml", options).unwrap();
    writer.write_all(content_types.as_bytes()).unwrap();
    writer.start_file("_rels/.rels", options).unwrap();
    writer.write_all(root.as_bytes()).unwrap();
    writer.start_file("word/document.xml", options).unwrap();
    writer.write_all(document.as_bytes()).unwrap();
    if let Some(relationships) = relationships {
        writer
            .start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        writer.write_all(relationships.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14""#;

/// A report exercising every body axis: a heading with an identity and
/// direct formatting, hyperlinks and a comment-reference run, a table with
/// a merged cell and a nested table, a content control, a `w:customXml`
/// wrapper, a display equation, a duplicated `w14:paraId` and a section
/// break.
fn report_xml() -> String {
    format!(
        r#"<w:document {NAMESPACES}><w:body><w:p w14:paraId="1A2B3C4D" w14:textId="77AA0001"><w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/><w:spacing w:before="240" w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Quarterly report: </w:t></w:r><w:r><w:rPr><w:i w:val="0"/><w:color w:val="auto" w:themeColor="accent1"/><w:sz w:val="21"/><w:u w:val="single"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>revenue grew</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="rIdLink" w:tooltip="Open the report"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>the report</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> or </w:t></w:r><w:hyperlink w:anchor="Intro"><w:r><w:t>the introduction</w:t></w:r></w:hyperlink><w:r><w:commentReference w:id="0"/></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/><w:tblLook w:firstRow="1" w:lastRow="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Sales</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p w14:paraId="0C0FFEE0"><w:r><w:t>West</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:alias w:val="Summary"/><w:tag w:val="sum"/><w:id w:val="-42"/></w:sdtPr><w:sdtContent><w:p w14:paraId="5D7C0001"><w:r><w:t>Inside the control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:customXml w:uri="urn:example" w:element="note"><w:p w14:paraId="00000002"><w:r><w:t>Wrapped</w:t></w:r></w:p></w:customXml><w:p><m:oMathPara><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara></w:p><w:p w14:paraId="1A2B3C4D"><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Duplicate identity</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    )
}

const REPORT_RELATIONSHIPS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/report" TargetMode="External"/></Relationships>"#;

fn report_bytes() -> Vec<u8> {
    package(&report_xml(), Some(REPORT_RELATIONSHIPS))
}

/// The report, opened source-backed (the default route).
fn report() -> File {
    crate::open_bytes(report_bytes()).expect("open the report")
}

/// The report, opened through the eager owner.
fn eager_report() -> File {
    crate::open_bytes_with(
        report_bytes(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open the eager report")
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn get<'f>(docx: &'f Docx, path: &str) -> Result<Node<'f>> {
    let path = path.into_path()?;
    resolve::get(&Nav::new(docx), &path).map(Node::Docx)
}

fn find<'f>(docx: &'f Docx, path: &str) -> Result<Option<Node<'f>>> {
    let path = path.into_path()?;
    Ok(resolve::find(&Nav::new(docx), &path)?.map(Node::Docx))
}

fn path(node: &Node<'_>) -> String {
    node.path().to_string()
}

fn text(node: &Node<'_>) -> Option<String> {
    node.view()
        .text()
        .expect("readable text")
        .map(|text| text.into_owned())
}

fn json(node: &Node<'_>, depth: u32) -> String {
    let mut sink = Vec::new();
    node.view_with(ViewOptions::default().depth(depth))
        .serialize_into(&mut serde_json::Serializer::pretty(&mut sink))
        .expect("serializes");
    String::from_utf8(sink).expect("UTF-8")
}

fn check_golden(name: &str, actual: &str) {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden/node_view/docx")
        .join(format!("{name}.json"));
    let actual = format!("{actual}\n");
    if std::env::var_os("LITCHI_BLESS").is_some() {
        fs::create_dir_all(path.parent().expect("golden directory")).expect("golden directory");
        fs::write(&path, &actual).expect("bless the golden file");
        return;
    }
    let expected = fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("read {}: {error}", path.display()));
    assert_eq!(
        actual, expected,
        "{name}: the node view differs from the OfficeCLI-shaped golden file \
         (LITCHI_BLESS=1 rewrites it)"
    );
}

/// The matches of `selector` and their paths, as the walker writes them.
fn query<'f>(docx: &'f Docx, selector: &str) -> Result<Vec<(Node<'f>, String)>> {
    let query = Query::parse(selector)?;
    let plan = Plan::compile(&query, FileFormat::Docx)?;
    let nav = Nav::new(docx);
    let mut walk = Walk::new(&nav, PlanRef::Borrowed(&plan))?;
    let mut matches = Vec::new();
    while let Some(hit) = walk.next_hit() {
        let hit = hit?;
        let written = hit.path().to_string();
        matches.push((Node::Docx(hit.node().clone()), written));
    }
    Ok(matches)
}

fn query_paths(docx: &Docx, selector: &str) -> Vec<String> {
    query(docx, selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .into_iter()
        .map(|(node, written)| {
            assert_eq!(
                path(&node),
                written,
                "{selector}: node and walk paths agree"
            );
            written
        })
        .collect()
}

fn cold_loads(file: &File) -> u64 {
    match docx(file).owner() {
        DocxOwner::Source(package) => package.cache_diagnostics().cold_loads,
        DocxOwner::Eager(_) => panic!("expected the source-backed owner"),
    }
}

fn view(docx: &Docx, mode: Mode, options: &Options) -> String {
    let mut out = String::new();
    super::views::view_to(docx, mode, options, &mut out)
        .unwrap_or_else(|error| panic!("{mode:?} view: {error}"));
    out
}

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

#[test]
fn body_paths_resolve_on_the_visible_view() {
    let file = report();
    let docx = docx(&file);
    let root = get(docx, "/").unwrap();
    assert_eq!((root.kind(), path(&root)), ("document", "/".to_owned()));
    assert_eq!(get(docx, "/body").unwrap().kind(), "body");
    // `/body/p[N]` counts the body's own paragraphs: the customXml wrapper
    // is flattened and the display equation is `oMathPara[1]` instead.
    let expected = [
        ("/body/p[1]", "Quarterly report: revenue grew"),
        ("/body/p[2]", "See the report or the introduction"),
        ("/body/p[3]", "Wrapped"),
        ("/body/p[4]", "Duplicate identity"),
    ];
    for (at, wanted) in expected {
        let node = get(docx, at).unwrap();
        assert_eq!(node.kind(), "paragraph", "{at}");
        assert_eq!(text(&node).as_deref(), Some(wanted), "{at}");
    }
    assert_eq!(
        find(docx, "/body/p[5]").unwrap().map(|node| path(&node)),
        None
    );
    assert_eq!(
        get(docx, "/body/p[last()]")
            .unwrap()
            .positional_path()
            .to_string(),
        "/body/p[4]"
    );
    assert_eq!(get(docx, "/body/oMathPara[1]").unwrap().kind(), "equation");
    assert_eq!(get(docx, "/body/equation[1]").unwrap().kind(), "equation");
    assert_eq!(get(docx, "/body/section[1]").unwrap().kind(), "section");
    assert_eq!(get(docx, "/BODY/Paragraph[2]").unwrap().kind(), "paragraph");
}

#[test]
fn identities_require_raw_access_and_paths_stay_positional() {
    let file = report();
    let docx = docx(&file);
    for identity in [
        "/body/p[@paraId=1A2B3C4D]",
        "/body/p[@paraId=1a2b3c4d]",
        "/body/p[@textId=77AA0001]",
        "/body/p[@paraId=1A2B3C4D][@textId=77AA0001]",
    ] {
        assert_eq!(get(docx, identity).unwrap_err().code(), Code::InvalidPath);
        let raw = crate::path::raw::path(identity).unwrap();
        let node = Node::Docx(resolve::get(&Nav::new(docx), &raw).unwrap());
        assert_eq!(path(&node), "/body/p[1]");
    }
    let raw = crate::path::raw::path("/body/sdt[@sdtId=-42]/p[1]").unwrap();
    let inside = Node::Docx(resolve::get(&Nav::new(docx), &raw).unwrap());
    assert_eq!(text(&inside).as_deref(), Some("Inside the control"));
    assert_eq!(path(&inside), "/body/sdt[1]/p[1]");
    for invalid in ["/body/p[@paraId=nothex]", "/body/sdt[@sdtId=notint]"] {
        let raw = crate::path::raw::path(invalid).unwrap();
        assert_eq!(
            resolve::get(&Nav::new(docx), &raw).unwrap_err().code(),
            Code::InvalidPath
        );
    }
    for missing in ["/body/p[@paraId=0BADF00D]", "/body/sdt[@sdtId=42]"] {
        let raw = crate::path::raw::path(missing).unwrap();
        assert!(resolve::find(&Nav::new(docx), &raw).unwrap().is_none());
    }
    assert_eq!(path(&get(docx, "/body/p[4]").unwrap()), "/body/p[4]");
    assert_eq!(
        path(&get(docx, "/body/p[1]/r[2]").unwrap()),
        "/body/p[1]/r[2]"
    );
}

#[test]
fn table_paths_reach_cells_and_their_own_blocks() {
    let file = report();
    let docx = docx(&file);
    let cell = get(docx, "/body/tbl[1]/tr[1]/tc[2]").unwrap();
    assert_eq!(cell.kind(), "cell");
    assert_eq!(text(&cell).as_deref(), Some("Sales"));
    let merged = get(docx, "/body/table[1]/row[2]/cell[1]").unwrap();
    assert_eq!(path(&merged), "/body/tbl[1]/tr[2]/tc[1]");
    // A cell's `p[N]` counts its own paragraphs, never a nested table's.
    let own = get(docx, "/body/tbl[1]/tr[2]/tc[1]/p[2]").unwrap();
    assert_eq!(text(&own).as_deref(), Some(""));
    assert_eq!(
        path(&get(docx, "/body/tbl[1]/tr[2]/tc[1]/p[1]").unwrap()),
        "/body/tbl[1]/tr[2]/tc[1]/p[1]"
    );
    let nested = get(docx, "/body/tbl[1]/tr[2]/tc[1]/tbl[1]/tr[1]/tc[1]/p[1]").unwrap();
    assert_eq!(text(&nested).as_deref(), Some("nested"));
    assert!(find(docx, "/body/tbl[1]/tr[3]").unwrap().is_none());
    assert_eq!(
        get(docx, "/body/tbl[1]/tr[last()]/tc[last()]")
            .unwrap()
            .kind(),
        "cell"
    );
}

#[test]
fn runs_follow_officecli_run_set_and_hyperlinks_resolve() {
    let file = report();
    let docx = docx(&file);
    // OfficeCLI's `r[N]` set: runs inside hyperlinks count, the
    // comment-reference run does not.
    let texts: Vec<String> = (1..=4)
        .map(|index| text(&get(docx, &format!("/body/p[2]/r[{index}]")).unwrap()).unwrap())
        .collect();
    assert_eq!(texts, ["See ", "the report", " or ", "the introduction"]);
    assert!(find(docx, "/body/p[2]/r[5]").unwrap().is_none());
    let link = get(docx, "/body/p[2]/hyperlink[1]").unwrap();
    assert_eq!(link.kind(), "hyperlink");
    assert_eq!(text(&link).as_deref(), Some("the report"));
    let format = format_map(&link);
    // The link's attributes, then its first text run's formatting.
    assert_eq!(
        format,
        [
            ("url", "https://example.com/report".to_owned()),
            ("tooltip", "Open the report".to_owned()),
            ("rStyle", "Hyperlink".to_owned()),
        ]
    );
    let anchor = get(docx, "/body/p[2]/hyperlink[2]").unwrap();
    assert_eq!(format_map(&anchor), [("anchor", "Intro".to_owned())]);
}

fn format_map(node: &Node<'_>) -> Vec<(&'static str, String)> {
    node.view()
        .format()
        .expect("readable format")
        .map(|prop| (prop.key(), prop.to_string()))
        .collect()
}

#[test]
fn misses_and_foreign_axes_have_their_codes() {
    let file = report();
    let docx = docx(&file);
    let cases = [
        ("/body/p[9]", Code::PathNotFound),
        ("/body/tbl[2]", Code::PathNotFound),
        ("/slide[1]", Code::UnsupportedPath),
        ("/Sheet1/A1", Code::UnsupportedPath),
        ("/body/customXml[1]", Code::UnsupportedPath),
        ("/body/tr[1]", Code::UnsupportedPath),
        ("/body/p[1]/tc[1]", Code::UnsupportedPath),
        ("/body/bogus[1]", Code::InvalidPath),
        ("/body/p[@style=Heading1]", Code::InvalidPath),
        ("/body/p[text~=x]", Code::InvalidPath),
        // Pictures are read since record 0251d; the report has none.
        ("/body/p[1]/picture[1]", Code::PathNotFound),
        // The report has no styles or numbering part (record 0319).
        ("/styles/Heading1", Code::PathNotFound),
        ("/numbering", Code::PathNotFound),
        ("/theme", Code::UnsupportedCapability),
    ];
    for (at, code) in cases {
        let error = get(docx, at).expect_err(at);
        assert_eq!(error.code(), code, "{at}: {error}");
    }
    let error = get(docx, "/body/p[9]").unwrap_err();
    let suggestion = error.suggestion().unwrap_or_default();
    assert!(
        suggestion.contains("p[1]..p[4]") && suggestion.contains("tbl[1]"),
        "{suggestion}"
    );
}

// ---------------------------------------------------------------------------
// The transaction ordinal (record 0010 §4.5.1)
// ---------------------------------------------------------------------------

#[test]
fn transaction_positions_map_visible_paths_to_raw_ordinals() {
    for file in [report(), eager_report()] {
        let docx = docx(&file);
        let position = |at: &str| {
            let node = get(docx, at).unwrap();
            node.as_docx().unwrap().transaction_position()
        };
        // Raw direct `w:p` children: p1, p2, the equation's `w:p`, p4.
        assert_eq!(position("/body/p[1]").unwrap().get(), 0);
        assert_eq!(position("/body/p[2]").unwrap().get(), 1);
        assert_eq!(position("/body/oMathPara[1]").unwrap().get(), 2);
        assert_eq!(position("/body/p[4]").unwrap().get(), 3);
        assert_eq!(position("/body/tbl[1]").unwrap().get(), 0);
        assert_eq!(position("/body/sdt[1]").unwrap().get(), 0);
        // The wrapped paragraph has no transaction position: refused, not
        // guessed.
        assert_eq!(
            position("/body/p[3]").unwrap_err().code(),
            Code::UnsupportedCapability
        );
        assert_eq!(
            position("/body/p[1]/r[1]").unwrap_err().code(),
            Code::UnsupportedCapability
        );
    }
}

#[test]
fn transaction_positions_are_refused_when_markup_compatibility_selected_content() {
    let document = format!(
        r#"<w:document {NAMESPACES}><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>chosen</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"#
    );
    let file = crate::open_bytes(package(&document, None)).expect("open");
    let docx = docx(&file);
    // Reads see the selected branch …
    let texts: Vec<String> = (1..=3)
        .map(|index| text(&get(docx, &format!("/body/p[{index}]")).unwrap()).unwrap())
        .collect();
    assert_eq!(texts, ["before", "chosen", "after"]);
    // … but raw and visible positions are not proven to agree, so no
    // paragraph offers a transaction position.
    for index in 1..=3 {
        let node = get(docx, &format!("/body/p[{index}]")).unwrap();
        let error = node.as_docx().unwrap().transaction_position().unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

// ---------------------------------------------------------------------------
// The node view (OfficeCLI `DocumentNode` shapes)
// ---------------------------------------------------------------------------

#[test]
fn node_views_match_officecli_shapes() {
    let file = report();
    let docx = docx(&file);
    check_golden("paragraph", &json(&get(docx, "/body/p[1]").unwrap(), 1));
    check_golden("run", &json(&get(docx, "/body/p[1]/r[2]").unwrap(), 0));
    check_golden("table", &json(&get(docx, "/body/tbl[1]").unwrap(), 2));
    check_golden(
        "cell",
        &json(&get(docx, "/body/tbl[1]/tr[2]/tc[1]").unwrap(), 1),
    );
    check_golden(
        "hyperlink",
        &json(&get(docx, "/body/p[2]/hyperlink[1]").unwrap(), 0),
    );
    check_golden("sdt", &json(&get(docx, "/body/sdt[1]").unwrap(), 1));
    // An equation's text is its LaTeX, which needs the `formula` feature.
    let body = if cfg!(feature = "formula") {
        "body"
    } else {
        "body_without_formula"
    };
    check_golden(body, &json(&get(docx, "/body").unwrap(), 1));
}

#[test]
fn node_views_of_a_real_document_match_officecli_shapes() {
    let file = open("ooxml/docx/drawing.docx");
    let docx = docx(&file);
    check_golden(
        "drawing_paragraph",
        &json(&get(docx, "/body/p[10]").unwrap(), 1),
    );
    check_golden(
        "drawing_table",
        &json(&get(docx, "/body/tbl[1]").unwrap(), 1),
    );
}

#[test]
fn the_eager_and_source_owners_produce_the_same_views() {
    let (source, eager) = (report(), eager_report());
    for at in [
        "/body",
        "/body/p[1]",
        "/body/p[2]",
        "/body/tbl[1]",
        "/body/tbl[1]/tr[2]/tc[1]",
        "/body/sdt[1]",
        "/body/p[2]/hyperlink[1]",
    ] {
        let source = json(&get(docx(&source), at).unwrap(), 2);
        let eager = json(&get(docx(&eager), at).unwrap(), 2);
        assert_eq!(source, eager, "{at}");
    }
}

#[test]
fn the_root_lists_the_body_without_reading_it() {
    let file = report();
    let docx = docx(&file);
    let root = get(docx, "/").unwrap();
    let before = cold_loads(&file);
    let view = json(&root, 1);
    // The root's own properties come from the core-properties part (absent
    // here); the body is listed without its blocks, so its `childCount` is
    // omitted (M1-Δ1) and the main part is not read.
    assert!(view.contains("\"path\": \"/body\""), "{view}");
    assert!(!view.contains("\"childCount\": 7"), "{view}");
    assert_eq!(cold_loads(&file), before, "listing the body reads no part");
}

// ---------------------------------------------------------------------------
// Decode counts
// ---------------------------------------------------------------------------

#[test]
fn gets_and_queries_read_only_the_main_document_part() {
    for name in [
        "ooxml/docx/headerFooter.docx",
        "ooxml/docx/footnotes.docx",
        "ooxml/docx/comment.docx",
        "ooxml/docx/Hyperlink.docx",
    ] {
        let file = open(name);
        let docx = docx(&file);
        assert!(docx.is_source_backed(), "{name}");
        assert_eq!(cold_loads(&file), 0, "{name}: open reads no payload");
        let node = get(docx, "/body/p[1]").unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(cold_loads(&file), 1, "{name}: one part, the main document");
        // Deeper steps, properties, hyperlinks and queries reuse it.
        let _ = node.view_with(ViewOptions::default().depth(2)).format();
        for at in [
            "/body/p[last()]",
            "/body/p[1]/r[1]",
            "/body/p[1]/hyperlink[1]",
        ] {
            let _ = find(docx, at).unwrap_or_else(|error| panic!("{name} {at}: {error}"));
        }
        let _ = query(docx, "p:has(hyperlink)").unwrap_or_else(|error| panic!("{name}: {error}"));
        let _ = query(docx, "r[bold=true]").unwrap();
        assert_eq!(cold_loads(&file), 1, "{name}: still only the main document");
    }
}

// ---------------------------------------------------------------------------
// Queries
// ---------------------------------------------------------------------------

#[test]
fn queries_select_by_style_format_text_and_structure() {
    let file = report();
    let docx = docx(&file);
    assert_eq!(query_paths(docx, "p[style=Heading1]"), ["/body/p[1]"]);
    assert_eq!(query_paths(docx, "paragraph[style=heading1]").len(), 1);
    assert_eq!(
        query_paths(docx, "r[bold=true]"),
        ["/body/p[1]/r[1]", "/body/tbl[1]/tr[1]/tc[2]/p[1]/r[1]",]
    );
    assert_eq!(query_paths(docx, "r[bold=false]"), ["/body/p[4]/r[1]"]);
    assert_eq!(query_paths(docx, "r[size=10.5pt]").len(), 1);
    assert_eq!(query_paths(docx, "r[color=accent1]").len(), 1);
    assert_eq!(query_paths(docx, "r[underline=single]").len(), 1);
    assert_eq!(query_paths(docx, "p:contains(REVENUE)"), ["/body/p[1]"]);
    assert_eq!(
        query_paths(docx, "p:has(r[bold=true])"),
        ["/body/p[1]", "/body/tbl[1]/tr[1]/tc[2]/p[1]",]
    );
    assert_eq!(query_paths(docx, "tbl:has(r[bold=true])"), ["/body/tbl[1]"]);
    assert_eq!(query_paths(docx, "p:has(hyperlink[url~=example])").len(), 1);
    assert_eq!(
        query_paths(docx, "tc[colspan=2]"),
        ["/body/tbl[1]/tr[2]/tc[1]"]
    );
    assert_eq!(query_paths(docx, "sdt[tag=sum] > p").len(), 1);
    assert_eq!(query_paths(docx, "p[numId=3]").len(), 1);
    assert_eq!(query_paths(docx, "equation").len(), 1);
    // `*` lists the body's direct children.
    assert_eq!(query_paths(docx, "*").len(), 8);
}

#[test]
fn queries_stop_at_limit_and_stream_in_document_order() {
    let file = report();
    let docx = docx(&file);
    let all = query_paths(docx, "r");
    assert!(all.len() >= 10, "{all:?}");
    let query = Query::parse("r").unwrap();
    let plan = Plan::compile(&query, FileFormat::Docx).unwrap();
    let nav = Nav::new(docx);
    let first: Vec<String> = Walk::new(&nav, PlanRef::Borrowed(&plan))
        .unwrap()
        .limit(2)
        .map(|node| path(&Node::Docx(node.unwrap())))
        .collect();
    assert_eq!(first, all[..2]);
    // `[N]` on the subject is the Nth match overall.
    assert_eq!(query_paths(docx, "r[3]"), [all[2].clone()]);
    assert_eq!(query_paths(docx, "r[last()]"), [all[all.len() - 1].clone()]);
}

#[test]
fn query_diagnostics_and_refusals() {
    let unknown = Query::parse("p[colour=red]").unwrap();
    let plan = Plan::compile(&unknown, FileFormat::Docx).unwrap();
    assert_eq!(plan.diagnostics().len(), 1);
    let gap = Query::parse("ole").unwrap();
    assert_eq!(
        Plan::compile(&gap, FileFormat::Docx).unwrap_err().code(),
        Code::UnsupportedCapability
    );
    let foreign = Query::parse("slide").unwrap();
    assert_eq!(
        Plan::compile(&foreign, FileFormat::Docx)
            .unwrap_err()
            .code(),
        Code::UnsupportedPath
    );
    // `styleName` is read since record 0319.
    let style_name = Query::parse("p[styleName=Title]").unwrap();
    assert!(
        Plan::compile(&style_name, FileFormat::Docx)
            .unwrap()
            .diagnostics()
            .is_empty()
    );
}

#[test]
fn queries_on_real_documents_agree_with_leaf_loops() {
    for name in [
        "ooxml/docx/drawing.docx",
        "poi/test-data/document/saut_page.docx",
        "poi/test-data/document/heading123.docx",
    ] {
        let file = open(name);
        let docx = docx(&file);
        let blocks = super::nav::load_body(docx).unwrap();
        // Every body paragraph whose own `w:pPr` names Heading1.
        let mut expected = 0;
        for ordinal in 0..blocks.count(litchi_docx::BlockKind::Paragraph) {
            let paragraph = blocks
                .paragraph(litchi_core::Position::new(ordinal))
                .unwrap();
            if paragraph.direct_properties().unwrap().style.as_deref() == Some("Heading1") {
                expected += 1;
            }
        }
        let headings = query(docx, "/body/p[style=Heading1]").unwrap();
        assert_eq!(headings.len(), expected, "{name}");
        // Every run with direct bold anywhere in the main story's body.
        let bold = query(docx, "r[bold=true]").unwrap();
        for (node, written) in &bold {
            assert_eq!(&path(node), written, "{name}");
            let run = node.as_docx().unwrap().as_run().unwrap().unwrap();
            assert_eq!(run.direct_properties().unwrap().bold, Some(true), "{name}");
        }
    }
}

// ---------------------------------------------------------------------------
// Views
// ---------------------------------------------------------------------------

#[test]
fn text_view_lists_body_elements_by_their_paths() {
    let file = report();
    let docx = docx(&file);
    let text = view(docx, Mode::Text, &Options::default());
    // With the `formula` feature an equation shows its LaTeX, as OfficeCLI's
    // equation node does.
    let equation = if cfg!(feature = "formula") {
        "[Equation] x=1"
    } else {
        "[Equation]"
    };
    let expected = format!(
        "\
[/body/p[1]] Quarterly report: revenue grew
[/body/p[2]] See the report or the introduction
[/body/tbl[1]] [Table: 2 rows]
[/body/sdt[1]/p[1]] [sdt:Summary] Inside the control
[/body/p[3]] Wrapped
[/body/oMathPara[1]] {equation}
[/body/p[4]] Duplicate identity"
    );
    assert_eq!(text, expected);
    let windowed = view(
        docx,
        Mode::Text,
        &Options {
            start: Some(2),
            max_lines: Some(2),
            ..Options::default()
        },
    );
    assert_eq!(
        windowed,
        "[/body/p[2]] See the report or the introduction\n[/body/tbl[1]] [Table: 2 rows]\n\
         ... (showed 2 rows, 8 total, use --start/--end to view more)"
    );
    let json = super::views::view_json(docx, Mode::Text, &Options::default()).unwrap();
    let json = serde_json::to_value(&json).unwrap();
    assert_eq!(json["totalElements"], 8);
    assert_eq!(json["elements"][3]["sdt"], "Summary");
    assert_eq!(json["elements"][2]["type"], "table");
}

#[test]
fn annotated_outline_and_stats_views() {
    let file = report();
    let docx = docx(&file);
    let annotated = view(docx, Mode::Annotated, &Options::default());
    assert!(
        annotated.contains(
            "[/body/p[1]] \u{300c}Quarterly report: \u{300d} \u{2190} Heading1 | Calibri 14pt bold"
        ),
        "{annotated}"
    );
    assert!(
        annotated.contains("[/body/tbl[1]] [Table: 2\u{d7}2]"),
        "{annotated}"
    );
    assert!(
        annotated.contains("\u{2190} Normal | (default) | [sdt:Summary]"),
        "{annotated}"
    );
    let outline = view(docx, Mode::Outline, &Options::default());
    assert!(
        // The equation's `w:p` is a paragraph to OfficeCLI's counts.
        outline
            .starts_with("6 paragraphs | 1 tables | 0 images | 1 equations | 1 content controls"),
        "{outline}"
    );
    assert!(
        outline
            .contains("\u{251c}\u{2500}\u{2500} [1] \"Quarterly report: revenue grew\" (Heading1)"),
        "{outline}"
    );
    let stats = view(docx, Mode::Stats, &Options::default());
    assert!(stats.starts_with("Paragraphs: 6 | Words: "), "{stats}");
    assert!(
        stats.contains("Style Distribution:\n  Normal: 5\n  Heading1: 1"),
        "{stats}"
    );
    assert!(stats.contains("Empty Paragraphs: 0"), "{stats}");
    assert!(stats.contains("Font Usage:\n  (default): "), "{stats}");
    assert!(stats.ends_with("Consecutive Spaces: 0"), "{stats}");
    let json = super::views::view_json(docx, Mode::Stats, &Options::default()).unwrap();
    let json = serde_json::to_value(&json).unwrap();
    assert_eq!(json["paragraphs"], 6);
    assert_eq!(json["styleDistribution"]["Heading1"], 1);
    let html = super::views::view_to(docx, Mode::Html, &Options::default(), &mut String::new());
    assert_eq!(html.unwrap_err().code(), Code::UnsupportedCapability);
    let range = Options {
        columns: Some(vec![0]),
        ..Options::default()
    };
    assert_eq!(
        super::views::view_to(docx, Mode::Text, &range, &mut String::new())
            .unwrap_err()
            .code(),
        Code::InvalidArgument
    );
}

#[test]
fn views_of_real_documents_read_every_body_element() {
    for name in [
        "ooxml/docx/drawing.docx",
        "poi/test-data/document/saut_page.docx",
        "poi/test-data/document/heading123.docx",
    ] {
        let file = open(name);
        let docx = docx(&file);
        for mode in [Mode::Text, Mode::Annotated, Mode::Outline, Mode::Stats] {
            let out = view(docx, mode, &Options::default());
            assert!(!out.is_empty(), "{name} {mode:?}");
            assert_eq!(
                out.trim_end(),
                out,
                "{name} {mode:?}: trailing white space is trimmed"
            );
        }
        // Every path the text view prints resolves to the element it names.
        let text = view(docx, Mode::Text, &Options::default());
        for line in text.lines() {
            let Some(at) = line
                .strip_prefix('[')
                .and_then(|rest| rest.split_once("] "))
            else {
                continue;
            };
            get(docx, at.0)
                .unwrap_or_else(|error| panic!("{name}: {} does not resolve: {error}", at.0));
        }
    }
}

// ---------------------------------------------------------------------------
// The corpus: paths round-trip
// ---------------------------------------------------------------------------

fn docx_corpus() -> Vec<PathBuf> {
    fn visit(directory: &Path, out: &mut Vec<PathBuf>) {
        let Ok(entries) = fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                visit(&path, out);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
            {
                out.push(path);
            }
        }
    }
    let mut out = Vec::new();
    visit(&test_data(), &mut out);
    out.sort();
    out
}

/// Visits `node` and its descendants down to `depth` levels, at most
/// `budget` nodes, checking that each node's canonical and positional paths
/// resolve back to it.
fn round_trip(
    docx: &Docx,
    node: &Node<'_>,
    depth: u32,
    budget: &mut usize,
    name: &str,
) -> Result<()> {
    if *budget == 0 {
        return Ok(());
    }
    *budget -= 1;
    for written in [node.path().to_string(), node.positional_path().to_string()] {
        let back = get(docx, &written)
            .unwrap_or_else(|error| panic!("{name}: {written} does not resolve: {error}"));
        assert_eq!(back.kind(), node.kind(), "{name}: {written}");
        assert_eq!(
            back.positional_path().to_string(),
            node.positional_path().to_string(),
            "{name}: {written}"
        );
        assert_eq!(
            back.path().to_string(),
            node.path().to_string(),
            "{name}: {written}"
        );
    }
    if depth == 0 {
        return Ok(());
    }
    for child in node.children() {
        round_trip(docx, &child?, depth - 1, budget, name)?;
    }
    Ok(())
}

#[test]
fn paths_round_trip_over_the_docx_corpus() {
    let mut opened = 0;
    let mut walked = 0;
    let mut refused = Vec::new();
    let mut summary = String::new();
    for path in docx_corpus() {
        let name = path.display().to_string();
        let Ok(file) = crate::open(&path) else {
            continue;
        };
        let File::Docx(docx) = &file else {
            continue;
        };
        opened += 1;
        let root = get(docx, "/").unwrap();
        // The root, the body, its blocks, their rows or runs, and the
        // cells, as deep as a table's paragraphs; at most 400 nodes a file.
        let mut budget = 400;
        match round_trip(docx, &root, 5, &mut budget, &name) {
            Ok(()) => walked += 1,
            Err(error) => {
                assert_ne!(error.code(), Code::InternalError, "{name}: {error}");
                refused.push((name.clone(), error.code()));
            },
        }
        let _ = writeln!(summary, "{name}: {} nodes", 400 - budget);
    }
    println!(
        "{opened} DOCX files opened, {walked} walked, {} refused",
        refused.len()
    );
    assert!(opened >= 40, "the corpus opened {opened} DOCX files");
    assert!(walked >= opened - refused.len(), "{refused:?}");
    // Only malformed or refused documents may stop a walk, with a typed
    // error.
    assert!(refused.len() <= 4, "{refused:?}\n{summary}");
}

#[test]
fn cached_reads_and_iterators_refuse_changed_sources() {
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("revision.docx");
    fs::write(&source, report_bytes()).unwrap();
    let file = crate::open(&source).unwrap();
    let docx = docx(&file);
    let paragraph = get(docx, "/body/p[1]").unwrap();
    let table = get(docx, "/body/tbl[1]").unwrap();
    let nav = Nav::new(docx);
    let query = Query::parse("p").unwrap();
    let plan = Plan::compile(&query, FileFormat::Docx).unwrap();
    let mut walk = Walk::new(&nav, PlanRef::Borrowed(&plan)).unwrap();
    assert!(walk.next().unwrap().is_ok());
    let paragraph_view = paragraph.view();
    let mut children = paragraph_view.children();
    assert!(children.next().unwrap().is_ok());
    fs::write(&source, b"changed source revision").unwrap();
    assert_eq!(get(docx, "/body/p[1]").unwrap_err().code(), Code::IoError);
    assert_eq!(
        paragraph
            .as_docx()
            .unwrap()
            .as_paragraph()
            .unwrap_err()
            .code(),
        Code::IoError
    );
    assert_eq!(
        paragraph
            .as_docx()
            .unwrap()
            .transaction_position()
            .unwrap_err()
            .code(),
        Code::IoError
    );
    assert_eq!(paragraph.view().text().unwrap_err().code(), Code::IoError);
    assert_eq!(
        table.as_docx().unwrap().table().unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(walk.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(walk.next().is_none());
    assert_eq!(children.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(children.next().is_none());
    assert_eq!(
        super::views::view_json(docx, Mode::Text, &Options::default())
            .unwrap_err()
            .code(),
        Code::IoError
    );
}

#[test]
fn invalid_direct_properties_are_not_projected_as_absent() {
    for content in [
        r#"<w:p><w:pPr><w:jc w:val="SECRET_BAD"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:rPr><w:b w:val="SECRET_BAD"/></w:rPr><w:t>text</w:t></w:r></w:p>"#,
    ] {
        let xml = format!("<w:document {NAMESPACES}><w:body>{content}</w:body></w:document>");
        let file = crate::open_bytes(package(&xml, None)).unwrap();
        let path = if content.contains("pPr") {
            "/body/p[1]"
        } else {
            "/body/p[1]/r[1]"
        };
        let node = get(docx(&file), path).unwrap();
        let error = node.view().format().unwrap_err();
        assert_eq!(error.code(), Code::CorruptFile);
        assert!(!format!("{error:?} {error}").contains("SECRET_BAD"));
        let mode = if content.contains("pPr") {
            Mode::Outline
        } else {
            Mode::Annotated
        };
        let error = super::views::view_json(docx(&file), mode, &Options::default()).unwrap_err();
        assert_eq!(error.code(), Code::CorruptFile);
    }
}

#[test]
fn zero_limit_reads_no_document_payload() {
    let file = report();
    let owner = docx(&file);
    let before = cold_loads(&file);
    let query = Query::parse("p").unwrap();
    let plan = Plan::compile(&query, FileFormat::Docx).unwrap();
    let nav = Nav::new(owner);
    let mut walk = Walk::new(&nav, PlanRef::Borrowed(&plan)).unwrap().limit(0);
    assert!(walk.next().is_none());
    assert_eq!(walk.visited(), 0);
    assert_eq!(cold_loads(&file), before);
}

#[test]
fn effective_source_owner_resolves_retained_hyperlinks() {
    let source = litchi_docx::source_backed::Package::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(report_bytes()),
    ))
    .unwrap();
    let mut edit = source.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "revised")
        .unwrap();
    let current = edit.commit().unwrap().into_package();
    let owner = Docx::from_source(current, &OpenOptions::default()).unwrap();
    let link = get(&owner, "/body/p[2]/hyperlink[1]").unwrap();
    assert_eq!(text(&link).as_deref(), Some("the report"));
    let output = json(&link, 0);
    assert!(output.contains("https://example.com/report"));
    assert_eq!(
        text(&get(&owner, "/body/p[1]").unwrap()).as_deref(),
        Some("revised")
    );
}

#[test]
fn native_query_predicates_require_raw_admission() {
    let file = report();
    let owner = docx(&file);
    let nav = Nav::new(owner);
    for (selector, expected) in [("p[paraId=1A2B3C4D]", 2), ("sdt[sdtId=-42]", 1)] {
        assert_eq!(
            crate::query::run(&nav, selector).err().unwrap().code(),
            Code::InvalidSelector
        );
        let raw = crate::query::raw::query(selector).unwrap();
        let nodes = crate::query::run(&nav, raw)
            .unwrap()
            .collect::<Result<Vec<_>>>()
            .unwrap();
        assert_eq!(nodes.len(), expected);
        for node in nodes {
            let serialized = json(&Node::Docx(node), 0);
            assert!(!serialized.contains("paraId"));
            assert!(!serialized.contains("sdtId"));
        }
    }
}

#[test]
fn capabilities_match_live_paths_queries_and_views() {
    let file = report();
    let caps = file.capabilities();
    // The report has no comments part; comment paths are checked on a copy
    // with one added (record 0251c).
    let commented = file
        .add(
            "/body/p[1]",
            crate::NewNode::new("comment")
                .prop("author", "Ada")
                .prop("text", "note"),
        )
        .expect("add a comment")
        .into_file();
    // Likewise for pictures (record 0251d).
    let pictured = file
        .add(
            "/body/p[1]",
            crate::NewNode::new("picture").prop(
                "src",
                "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC",
            ),
        )
        .expect("add a picture")
        .into_file();
    // And charts (record 0283).
    let charted = file
        .add(
            "/body/p[1]",
            crate::NewNode::new("chart").prop("data", "A:1,2"),
        )
        .expect("add a chart")
        .into_file();
    // And SmartArt (record 0295).
    let smartart = file
        .add(
            "/body/p[1]",
            crate::NewNode::new("smartart").prop("text", "A\n\tA1\nB"),
        )
        .expect("add a SmartArt graphic")
        .into_file();
    // And the marks (record 0319): a tab stop, a break, a bookmark, a
    // simple field, a table of contents and a text box in the main story,
    // and a watermark in a built document's header.
    let marked = crate::open_bytes(package(
        &format!(
            r#"<w:document {NAMESPACES} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:bookmarkStart w:id="1" w:name="Intro"/><w:r><w:t>One</w:t></w:r><w:bookmarkEnd w:id="1"/><w:r><w:br w:type="page"/></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h "><w:r><w:t>Contents</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:drawing><wp:anchor><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Box 7"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr id="7" name="Box 7"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>"#
        ),
        None,
    ))
    .expect("open the marked document");
    let watermarked = {
        let mut builder = litchi_docx::Builder::new().expect("builder");
        builder
            .document()
            .expect("document")
            .add_paragraph_with_text("Body");
        builder
            .document()
            .expect("document")
            .set_watermark(litchi_docx::Watermark::text("DRAFT"));
        let package = builder.build().expect("build");
        crate::open_bytes(package.to_bytes().expect("bytes"))
            .expect("open the watermarked document")
    };
    // And the part elements (record 0319).
    let parted = file
        .add("/styles", crate::NewNode::new("style").prop("id", "Accent"))
        .or_else(|_| file.add("/", crate::NewNode::new("style").prop("id", "Accent")))
        .expect("add a style")
        .into_file();
    let parted = parted
        .add("/numbering", crate::NewNode::new("abstractNum"))
        .or_else(|_| parted.add("/", crate::NewNode::new("abstractNum")))
        .expect("add a numbering definition")
        .into_file();
    let parted = parted
        .add(
            "/numbering",
            crate::NewNode::new("num").prop("abstractNumId", 0),
        )
        .expect("add a numbering instance")
        .into_file();
    let parted = parted
        .add("/", crate::NewNode::new("header").prop("text", "Head"))
        .expect("add a header")
        .into_file();
    let parted = parted
        .add("/", crate::NewNode::new("footer").prop("text", "Foot"))
        .expect("add a footer")
        .into_file();
    let parted = parted
        .add(
            "/body/p[1]",
            crate::NewNode::new("footnote").prop("text", "Note"),
        )
        .expect("add a footnote")
        .into_file();
    let parted = parted
        .add(
            "/body/p[1]",
            crate::NewNode::new("endnote").prop("text", "End"),
        )
        .expect("add an endnote")
        .into_file();
    let file_for = |name: &str| {
        if matches!(
            name,
            "styles"
                | "style"
                | "numbering"
                | "abstractNum"
                | "num"
                | "level"
                | "header"
                | "footer"
                | "footnotes"
                | "footnote"
                | "endnotes"
                | "endnote"
        ) {
            &parted
        } else if matches!(name, "comments" | "comment") {
            &commented
        } else if name == "picture" {
            &pictured
        } else if matches!(name, "chart" | "series" | "axis") {
            &charted
        } else if matches!(name, "smartart" | "node") {
            &smartart
        } else if matches!(
            name,
            "tab" | "break" | "bookmark" | "field" | "toc" | "textbox"
        ) {
            &marked
        } else if name == "watermark" {
            &watermarked
        } else {
            &file
        }
    };
    for element in caps.elements() {
        if !element.operations().get {
            continue;
        }
        for (template, _) in element.paths() {
            // Native identities need raw paths; they are checked elsewhere.
            if template.contains("Id=") || template.contains("[@id=") {
                continue;
            }
            let path = template
                .replace("<id or name>", "Accent")
                .replace("[@type=T]", "[@type=default]")
                .replace("[@name=X]", "[@name=Intro]")
                .replace("[@role=ROLE]", "[@role=category]")
                .replace("[L]", "[1]")
                .replace("[N]", "[1]")
                .replace("[M]", "[1]")
                .replace("[K]", "[1]")
                .replace("[J]", "[1]")
                .replace("/body/p[1]/hyperlink", "/body/p[2]/hyperlink")
                .replace(
                    "/body/tbl[1]/tr[1]/tc[1]/tbl",
                    "/body/tbl[1]/tr[2]/tc[1]/tbl",
                );
            // Story paragraphs (record 0325) live in the parts fixture.
            let owner = if path.starts_with("/header") || path.starts_with("/footnote") {
                file_for("header")
            } else {
                file_for(element.name())
            };
            let node = owner
                .get(path.as_str())
                .unwrap_or_else(|error| panic!("{path}: {error}"));
            assert_eq!(node.kind(), element.name());
            for entry in node.view().format().unwrap() {
                assert!(
                    element.property(entry.key()).is_some(),
                    "{path}: {}",
                    entry.key()
                );
            }
        }
        if element.operations().query {
            assert!(
                file_for(element.name())
                    .query(element.name())
                    .unwrap()
                    .first()
                    .unwrap()
                    .is_some(),
                "{}",
                element.name()
            );
        }
    }
    for &mode in Mode::ALL {
        let result = file.view(mode);
        assert_eq!(caps.view(mode).is_supported(), result.is_ok());
        if let Err(error) = result {
            assert_eq!(caps.view(mode).code(), Some(error.code()));
        }
    }
}
