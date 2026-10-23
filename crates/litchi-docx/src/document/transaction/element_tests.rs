//! Path-addressed element edits: nested paragraphs, runs and hyperlinks,
//! table structure, exact inverses and durable replay.

use super::*;
use crate::properties::Justification;
use litchi_core::patch::{BlobLimits, PatchLimits};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn snapshot(body: &str) -> Snapshot {
    Snapshot::from_xml(
        format!(r#"<w:document xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}"><w:body>{body}</w:body></w:document>"#)
            .into_bytes(),
    )
    .unwrap()
}

fn limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(64, 32 * 1024 * 1024, 64 * 1024 * 1024),
        64 * 1024 * 1024,
        4_096,
        16,
        16 * 1024 * 1024,
        64 * 1024 * 1024,
    )
}

/// Commit and prove the patch exact both ways, in memory and durably.
fn commit_and_check(base: &Snapshot, edit: Edit) -> Commit {
    let commit = edit.commit().unwrap();
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), base.xml_bytes());
    if !commit.patch().is_noop() {
        let durable = commit.patch().to_durable(limits()).unwrap();
        let json = durable.to_deterministic_json().unwrap();
        let durable =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &json,
                limits(),
            )
            .unwrap();
        let replayed = base.apply_durable(&durable).unwrap();
        assert_eq!(replayed.xml_bytes(), commit.snapshot().xml_bytes());
        let inverse = commit.patch().inverse().to_durable(limits()).unwrap();
        let undone = commit.snapshot().apply_durable(&inverse).unwrap();
        assert_eq!(undone.xml_bytes(), base.xml_bytes());
    }
    commit
}

fn xml(commit: &Commit) -> String {
    String::from_utf8(commit.snapshot().xml_bytes().to_vec()).unwrap()
}

fn cell(table: usize, row: usize, cell: usize) -> ElementPath {
    ElementPath::new([
        ElementStep::Table(Position::new(table)),
        ElementStep::Row(Position::new(row)),
        ElementStep::Cell(Position::new(cell)),
    ])
}

const TABLE: &str = r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>"#;

#[test]
fn nested_paragraph_and_run_formatting_is_exact_and_reversible() {
    let base = snapshot(&format!("<w:p/>{TABLE}<w:sectPr/>"));
    let mut edit = base.edit();
    edit.format_paragraph_at(
        &cell(0, 0, 0).child(ElementStep::Paragraph(Position::new(0))),
        |paragraph| {
            paragraph.justification = Some(Justification::Center);
        },
    )
    .unwrap();
    edit.format_run_at(
        &cell(0, 0, 1)
            .child(ElementStep::Paragraph(Position::new(0)))
            .child(ElementStep::Hyperlink(Position::new(0)))
            .child(ElementStep::Run(Position::new(0))),
        |run| run.bold = Some(true),
    )
    .unwrap();
    let commit = commit_and_check(&base, edit);
    let text = xml(&commit);
    assert!(text.contains(r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t>"#));
    assert!(text.contains(r#"<w:hyperlink r:id="rId1"><w:r><w:rPr><w:b/></w:rPr><w:t>link</w:t>"#));
    // Unchanged values are exact no-ops.
    let mut noop = commit.snapshot().edit();
    noop.format_paragraph_at(
        &cell(0, 0, 0).child(ElementStep::Paragraph(Position::new(0))),
        |paragraph| {
            paragraph.justification = Some(Justification::Center);
        },
    )
    .unwrap();
    assert!(noop.commit().unwrap().patch().is_noop());
}

#[test]
fn nested_text_replacement_reaches_paragraphs_runs_hyperlinks_and_controls() {
    let base = snapshot(&format!(
        r#"{TABLE}<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>ctl</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
    ));
    let mut edit = base.edit();
    let first = cell(0, 0, 0).child(ElementStep::Paragraph(Position::new(0)));
    edit.replace_text_at(&first, "A").unwrap();
    let paragraph = cell(0, 0, 1).child(ElementStep::Paragraph(Position::new(0)));
    edit.replace_text_at(
        &paragraph.child(ElementStep::Hyperlink(Position::new(0))),
        "LINK",
    )
    .unwrap();
    edit.replace_text_at(&paragraph.child(ElementStep::Run(Position::new(0))), "B")
        .unwrap();
    let control = ElementPath::new([
        ElementStep::ContentControl(Position::new(0)),
        ElementStep::Paragraph(Position::new(0)),
    ]);
    edit.replace_text_at(&control, "CTL").unwrap();
    let commit = commit_and_check(&base, edit);
    let text = xml(&commit);
    for expected in [
        "<w:t>A</w:t>",
        "<w:t>LINK</w:t>",
        "<w:t>B</w:t>",
        "<w:t>CTL</w:t>",
    ] {
        assert!(text.contains(expected), "{expected} in {text}");
    }
}

#[test]
fn runs_insert_into_nested_paragraphs_and_empty_paragraphs() {
    let base = snapshot(r#"<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#);
    let mut edit = base.edit();
    let paragraph = cell(0, 0, 0).child(ElementStep::Paragraph(Position::new(0)));
    let mut run = NewRun::new("one");
    run.properties_mut().italic = Some(true);
    assert_eq!(
        edit.insert_run_at(&paragraph, None, &run).unwrap(),
        Position::new(0)
    );
    assert_eq!(
        edit.insert_run_at(&paragraph, Some(Position::new(0)), &NewRun::new("zero"))
            .unwrap(),
        Position::new(0)
    );
    let commit = commit_and_check(&base, edit);
    let text = xml(&commit);
    assert!(text.contains(
        r#"<w:tc><w:p><w:r><w:t>zero</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>one</w:t></w:r></w:p></w:tc>"#
    ));
}

#[test]
fn paragraphs_insert_and_remove_inside_cells_but_a_cell_keeps_one() {
    let base = snapshot(TABLE);
    let mut edit = base.edit();
    let container = cell(0, 0, 0);
    let position = edit
        .insert_paragraph_at(
            &container,
            ElementAnchor::End,
            &NewParagraph::new().with_run(NewRun::new("second")),
        )
        .unwrap();
    assert_eq!(position, Position::new(1));
    edit.remove_element_at(&container.child(ElementStep::Paragraph(Position::new(0))))
        .unwrap();
    assert!(matches!(
        edit.remove_element_at(&container.child(ElementStep::Paragraph(Position::new(0)))),
        Err(Error::Refused {
            reason: Refusal::ComplexContent,
            ..
        })
    ));
    let commit = commit_and_check(&base, edit);
    assert!(xml(&commit).contains(r#"</w:tcPr><w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc>"#));
}

#[test]
fn tables_insert_grow_shrink_move_and_remove_exactly() {
    let base = snapshot(r#"<w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/>"#);
    let mut edit = base.edit();
    let table = NewTable::new(2, 2)
        .unwrap()
        .with_cell_text(0, 0, "h1")
        .with_cell_text(1, 1, "x & y");
    assert_eq!(
        edit.insert_table_at(&ElementPath::body(), ElementAnchor::End, &table)
            .unwrap(),
        Position::new(0)
    );
    let path = ElementPath::new([ElementStep::Table(Position::new(0))]);
    assert_eq!(
        edit.insert_row_at(&path, None, None, &["r3".to_owned()])
            .unwrap(),
        Position::new(2)
    );
    edit.remove_element_at(&path.child(ElementStep::Row(Position::new(0))))
        .unwrap();
    // Move the table before the paragraph.
    let moved = edit
        .move_element_at(
            &path,
            &ElementPath::body(),
            ElementAnchor::Before(ElementStep::Paragraph(Position::new(0))),
        )
        .unwrap();
    assert_eq!(moved, path);
    let commit = commit_and_check(&base, edit);
    let text = xml(&commit);
    let table_start = text.find("<w:tbl>").unwrap();
    assert!(table_start < text.find("before").unwrap());
    assert!(
        text.contains(r#"<w:t xml:space="preserve">x &amp; y</w:t>"#)
            || text.contains("<w:t>x &amp; y</w:t>")
    );
    assert!(text.contains("<w:t>r3</w:t>"));
    assert!(!text.contains("h1"));
    assert!(text.ends_with("<w:sectPr/></w:body></w:document>"));
    assert_eq!(commit.snapshot().table_count(), 1);

    let mut removal = commit.snapshot().edit();
    removal.remove_element_at(&path).unwrap();
    let removed = removal.commit().unwrap();
    assert_eq!(removed.snapshot().table_count(), 0);
}

#[test]
fn cell_text_replacement_keeps_cell_and_first_paragraph_properties() {
    let base = snapshot(
        r#"<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>old</w:t></w:r></w:p><w:p><w:r><w:t>more</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut edit = base.edit();
    edit.replace_cell_text_at(&cell(0, 0, 0), "new").unwrap();
    let commit = commit_and_check(&base, edit);
    assert!(xml(&commit).contains(
        r#"<w:tc><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>new</w:t></w:r></w:p></w:tc>"#
    ));
}

#[test]
fn cell_text_replacement_refuses_content_it_would_drop() {
    // Record 0304: a later block that is not plain text, and a first
    // paragraph the text rewrite refuses, are refused rather than dropped.
    for content in [
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:pict/></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>"#,
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:hyperlink r:id="rId1"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>"#,
        r#"<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:r><w:pict/></w:r></w:p>"#,
    ] {
        let base = snapshot(&format!(
            "<w:tbl><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>"
        ));
        let mut edit = base.edit();
        assert!(
            matches!(
                edit.replace_cell_text_at(&cell(0, 0, 0), "new"),
                Err(Error::Refused { .. })
            ),
            "{content}"
        );
        assert!(edit.commit().unwrap().patch().is_noop());
    }
    // Plain later paragraphs, with formatting, tabs and breaks, are text.
    let base = snapshot(
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t><w:tab/><w:br/></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut edit = base.edit();
    edit.replace_cell_text_at(&cell(0, 0, 0), "new").unwrap();
    commit_and_check(&base, edit);
    // An empty first paragraph has nothing to lose.
    let base = snapshot(r#"<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#);
    let mut edit = base.edit();
    edit.replace_cell_text_at(&cell(0, 0, 0), "new").unwrap();
    commit_and_check(&base, edit);
}

#[test]
fn paragraph_text_replacement_preserves_bookmark_markers() {
    let base = snapshot(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
    );
    let mut edit = base.edit();
    edit.replace_paragraph_text(Position::new(0), "new")
        .unwrap();
    let commit = commit_and_check(&base, edit);
    let output = xml(&commit);
    assert!(output.contains(r#"<w:bookmarkStart w:id="1" w:name="b"/>"#));
    assert!(output.contains(r#"<w:t>new</w:t>"#));
    assert!(output.contains(r#"<w:bookmarkEnd w:id="1"/>"#));
}

#[test]
fn selected_paragraph_routes_replay_and_inverse_bookmarked_text() {
    let base = snapshot(
        r#"<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:bookmarkStart w:id="10" w:name="block"/><w:r><w:t>block old</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:bookmarkStart w:id="20" w:name="cell"/><w:r><w:t>cell old</w:t></w:r><w:bookmarkEnd w:id="20"/></w:p><w:tbl><w:tr><w:tc><w:p><w:bookmarkStart w:id="30" w:name="nested"/><w:r><w:t>nested old</w:t></w:r><w:bookmarkEnd w:id="30"/></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>"#,
    );
    let controls = [Position::new(0)];
    let table_path = [
        TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
        TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
    ];
    let mut edit = base.edit();
    edit.replace_block_content_control_paragraph_text(&controls, Position::new(0), "block new")
        .unwrap()
        .replace_table_cell_paragraph_text(
            Position::new(0),
            Position::new(0),
            Position::new(0),
            Position::new(0),
            "cell new",
        )
        .unwrap()
        .replace_nested_table_cell_paragraph_text(&table_path, Position::new(0), "nested new")
        .unwrap();
    let commit = commit_and_check(&base, edit);
    let output = xml(&commit);
    for retained in [
        r#"<w:bookmarkStart w:id="10" w:name="block"/>"#,
        r#"<w:t>block new</w:t>"#,
        r#"<w:bookmarkEnd w:id="10"/>"#,
        r#"<w:bookmarkStart w:id="20" w:name="cell"/>"#,
        r#"<w:t>cell new</w:t>"#,
        r#"<w:bookmarkEnd w:id="20"/>"#,
        r#"<w:bookmarkStart w:id="30" w:name="nested"/>"#,
        r#"<w:t>nested new</w:t>"#,
        r#"<w:bookmarkEnd w:id="30"/>"#,
    ] {
        assert!(
            output.contains(retained),
            "missing retained XML: {retained}"
        );
    }
}

#[test]
fn removals_that_would_orphan_ranges_are_refused() {
    let base = snapshot(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="b"/></w:p><w:tbl><w:tr><w:tc><w:p><w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut edit = base.edit();
    assert!(matches!(
        edit.remove_element_at(&ElementPath::new([ElementStep::Table(Position::new(0))])),
        Err(Error::Refused {
            reason: Refusal::OpenRange,
            ..
        })
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
}

#[test]
fn stale_element_patches_are_refused() {
    let base = snapshot(TABLE);
    let mut edit = base.edit();
    edit.replace_text_at(
        &cell(0, 0, 0).child(ElementStep::Paragraph(Position::new(0))),
        "x",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let other = snapshot(&TABLE.replace(">a<", ">c<"));
    assert!(commit.patch().apply(&other).is_err());
    let durable = commit.patch().to_durable(limits()).unwrap();
    assert!(other.apply_durable(&durable).is_err());
}

#[test]
fn element_paths_round_trip_their_durable_spelling() {
    let path = cell(1, 2, 3).child(ElementStep::ContentControl(Position::new(4)));
    assert_eq!(path.to_string(), "tbl:1/tr:2/tc:3/sdt:4");
    assert_eq!(ElementPath::parse(&path.to_string()), Some(path));
    assert_eq!(ElementPath::parse(""), Some(ElementPath::body()));
    assert_eq!(ElementPath::parse("x:1"), None);
}

/// Record 0342: table borders and row properties, spliced in schema order,
/// exact both ways; other properties are refused, and an unchanged value
/// is a no-op.
#[test]
fn table_borders_and_row_properties_are_exact_and_reversible() {
    use crate::properties::{
        Border, BorderStyle, EighthPoints, Height, HeightRule, HexColor, Twips,
    };
    let base = snapshot(
        r#"<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tblPrEx><w:tblW w:w="0" w:type="auto"/></w:tblPrEx><w:tc><w:p/></w:tc></w:tr><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#,
    );
    let table = ElementPath::new([ElementStep::Table(Position::new(0))]);
    let mut edit = base.edit();
    edit.format_table_at(&table, |table| {
        let mut border = Border::new(BorderStyle::from_xml("double").unwrap());
        border.size = Some(EighthPoints(8));
        border.color = Some(HexColor::Rgb([0x1F, 0x38, 0x64]));
        let borders = table.borders.get_or_insert_default();
        borders.top = Some(border);
        borders.inside_vertical = Some(border);
    })
    .unwrap();
    edit.format_row_at(&table.child(ElementStep::Row(Position::new(0))), |row| {
        row.header = Some(true);
        row.height = Some(Height::new(Twips(400), Some(HeightRule::Exact)));
    })
    .unwrap();
    edit.format_row_at(&table.child(ElementStep::Row(Position::new(1))), |row| {
        row.header = Some(true);
    })
    .unwrap();
    let commit = commit_and_check(&base, edit);
    let xml = xml(&commit);
    assert!(xml.contains(r#"<w:tblStyle w:val="Grid"/><w:tblBorders><w:top w:val="double" w:color="1F3864" w:sz="8"/><w:insideV w:val="double" w:color="1F3864" w:sz="8"/></w:tblBorders><w:tblLook w:val="04A0"/>"#), "{xml}");
    // A new `w:trPr` follows the row's `w:tblPrEx`.
    assert!(xml.contains(r#"</w:tblPrEx><w:trPr><w:trHeight w:val="400" w:hRule="exact"/><w:tblHeader/></w:trPr>"#), "{xml}");
    assert!(
        xml.contains(r#"<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>"#),
        "{xml}"
    );

    let mut edit = commit.snapshot().edit();
    edit.format_table_at(&table, |table| table.layout = None)
        .map(|_| ())
        .expect("an unchanged layout is a no-op");
    edit.format_table_at(&table, |table| table.caption = Some("caption".to_owned()))
        .map(|_| ())
        .expect_err("only the style and borders are written");
    edit.format_row_at(&table.child(ElementStep::Row(Position::new(1))), |row| {
        row.header = Some(true);
    })
    .unwrap();
    assert!(edit.commit().unwrap().patch().is_noop());

    let mut edit = commit.snapshot().edit();
    edit.format_table_at(&table, |table| table.style = Some("Other".to_owned()))
        .unwrap();
    let styled = commit_and_check(commit.snapshot(), edit);
    assert!(
        String::from_utf8_lossy(styled.snapshot().xml_bytes())
            .contains(r#"<w:tblPr><w:tblStyle w:val="Other"/><w:tblBorders>"#)
    );
}
