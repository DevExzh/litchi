use super::super::Snapshot;
use super::{FieldKind, TocEntry, Update};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn document(body: &str) -> String {
    format!(
        r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>"#
    )
}

const TOC_PLACEHOLDER: &str = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Update field to see table of contents</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;

fn heading(level: u8, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

/// Commit `updates`, then check the patch replays onto the source and its
/// inverse restores it exactly.
fn commit(source: &Snapshot, updates: impl Fn(&super::Fields) -> Vec<Update>) -> Snapshot {
    let mut edit = source.edit();
    let fields = edit.fields().unwrap();
    let updates = updates(&fields);
    edit.refresh_fields(&fields, &updates).unwrap();
    let commit = edit.commit().unwrap();
    let (target, patch) = commit.into_parts();
    let replayed = patch.apply(source).unwrap();
    assert_eq!(replayed.xml_bytes(), target.xml_bytes(), "replay");
    let restored = patch.inverse().apply(&target).unwrap();
    assert_eq!(restored.xml_bytes(), source.xml_bytes(), "inverse");
    target
}

fn text(snapshot: &Snapshot) -> String {
    String::from_utf8(snapshot.xml_bytes().to_vec()).unwrap()
}

#[test]
fn the_index_reads_paragraphs_fields_and_bookmarks() {
    let xml = document(&format!(
        r#"{TOC_PLACEHOLDER}{}<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:bookmarkStart w:id="4" w:name="Intro"/><w:r><w:t>Body</w:t></w:r><w:r><w:tab/></w:r><w:bookmarkEnd w:id="4"/></w:p><w:p><w:fldSimple w:instr=" PAGE " w:fldLock="1"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>"#,
        heading(1, "One")
    ));
    let snapshot = Snapshot::from_xml(xml).unwrap();
    let fields = snapshot.edit().fields().unwrap();
    let paragraphs = fields.paragraphs();
    assert_eq!(paragraphs.len(), 4);
    assert_eq!(
        paragraphs[0].text(),
        "Update field to see table of contents"
    );
    assert_eq!(paragraphs[1].style(), Some("Heading1"));
    assert_eq!(paragraphs[2].outline_level(), Some(1));
    assert_eq!(paragraphs[2].text(), "Body\t");
    assert_eq!(fields.bookmarks()[0].name(), "Intro");
    assert_eq!(fields.bookmarks()[0].paragraph(), Some(2));
    assert_eq!(fields.bookmarks()[0].text(), "Body\t");
    assert_eq!(fields.text_width(), Some(9360));
    let toc = &fields.fields()[0];
    assert_eq!(toc.kind(), FieldKind::Complex);
    assert_eq!(toc.instruction().trim(), r#"TOC \o "1-3" \h"#);
    assert_eq!(toc.result(), "Update field to see table of contents");
    assert!(fields.accepts_text(0) && fields.accepts_paragraphs(0));
    let page = &fields.fields()[1];
    assert_eq!(page.kind(), FieldKind::Simple);
    assert!(page.is_locked());
    assert_eq!(page.result(), "3");
}

#[test]
fn a_text_update_keeps_the_code_and_formatting() {
    let xml = document(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:fldSimple w:instr=" SEQ Figure "/></w:p><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    let source = Snapshot::from_xml(xml).unwrap();
    let target = commit(&source, |_| {
        vec![
            Update::Text {
                field: 0,
                text: "1".into(),
            },
            Update::Text {
                field: 1,
                text: "2".into(),
            },
            Update::Text {
                field: 2,
                text: "3".into(),
            },
        ]
    });
    let out = text(&target);
    assert!(
        out.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r>"#),
        "{out}"
    );
    assert!(out.contains(r#"<w:fldSimple w:instr=" SEQ Figure "><w:r><w:t xml:space="preserve">2</w:t></w:r></w:fldSimple>"#), "{out}");
    let fields = target.edit().fields().unwrap();
    let results: Vec<&str> = fields.fields().iter().map(super::Field::result).collect();
    assert_eq!(results, ["1", "2", "3"]);
    assert!(fields.fields()[2].has_result());
}

#[test]
fn an_unchanged_text_update_is_a_no_op() {
    let xml = document(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ A </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    let source = Snapshot::from_xml(xml).unwrap();
    let mut edit = source.edit();
    let fields = edit.fields().unwrap();
    let outcome = edit
        .refresh_fields(
            &fields,
            &[Update::Text {
                field: 0,
                text: "1".into(),
            }],
        )
        .unwrap();
    assert!(outcome.is_empty());
    assert!(edit.commit().unwrap().patch().is_noop());
}

#[test]
fn a_body_level_toc_is_rebuilt_with_bookmarks() {
    let xml = document(&format!(
        "{TOC_PLACEHOLDER}{}{}<w:p><w:r><w:t>tail</w:t></w:r></w:p>",
        heading(1, "One"),
        heading(2, "Two")
    ));
    let source = Snapshot::from_xml(xml).unwrap();
    let target = commit(&source, |fields| {
        let width = fields.text_width().unwrap();
        vec![
            Update::Bookmark {
                paragraph: 1,
                name: "_Toc100000001".into(),
            },
            Update::Bookmark {
                paragraph: 2,
                name: "_Toc100000002".into(),
            },
            Update::Toc {
                field: 0,
                entries: vec![
                    TocEntry::new(1, "One")
                        .anchor("_Toc100000001")
                        .hyperlink(true)
                        .page("")
                        .tab(width),
                    TocEntry::new(2, "Two")
                        .anchor("_Toc100000002")
                        .hyperlink(true)
                        .page("4")
                        .tab(width),
                ],
                dirty: true,
            },
        ]
    });
    let out = text(&target);
    assert!(
        out.contains(r#"<w:fldChar w:fldCharType="begin" w:dirty="true"/>"#),
        "{out}"
    );
    assert!(out.contains(r#"<w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>"#), "{out}");
    assert!(
        out.contains(r#"<w:hyperlink w:anchor="_Toc100000001" w:history="1">"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<w:bookmarkStart w:id="0" w:name="_Toc100000001"/>"#),
        "{out}"
    );
    assert!(!out.contains("Update field"), "{out}");
    let fields = target.edit().fields().unwrap();
    let texts: Vec<&str> = fields
        .paragraphs()
        .iter()
        .map(super::Paragraph::text)
        .collect();
    assert_eq!(texts, ["", "One\t", "Two\t4", "", "One", "Two", "tail"]);
    let toc = &fields.fields()[0];
    assert_eq!(toc.paragraphs(), (0, 3));
    // The PAGEREF fields nest in the TOC.
    assert_eq!(fields.fields()[2].parent(), Some(0));
    assert_eq!(
        fields.fields()[2].instruction(),
        r" PAGEREF _Toc100000002 \h "
    );
    assert_eq!(fields.fields()[2].result(), "4");
    // A second refresh with the same entries changes only nothing.
    assert!(fields.accepts_paragraphs(0));
}

#[test]
fn a_multi_paragraph_toc_in_a_content_control_is_one_element_replacement() {
    let body = format!(
        r#"<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Old</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Older</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>after</w:t></w:r></w:p></w:sdtContent></w:sdt>{}"#,
        heading(1, "New")
    );
    let source = Snapshot::from_xml(document(&body)).unwrap();
    let target = commit(&source, |_| {
        vec![Update::Toc {
            field: 0,
            entries: vec![TocEntry::new(1, "New")],
            dirty: false,
        }]
    });
    let fields = target.edit().fields().unwrap();
    let texts: Vec<&str> = fields
        .paragraphs()
        .iter()
        .map(super::Paragraph::text)
        .collect();
    assert_eq!(texts, ["", "New", "after", "New"]);
    assert_eq!(fields.paragraphs()[2].style(), None);
}

#[test]
fn a_simple_toc_becomes_a_complex_field() {
    let source = Snapshot::from_xml(document(&format!(
        r#"<w:p><w:r><w:t>before </w:t></w:r><w:fldSimple w:instr="TOC \o &quot;1-1&quot;"><w:r><w:t>Right-click</w:t></w:r></w:fldSimple></w:p>{}"#,
        heading(1, "H")
    )))
    .unwrap();
    let target = commit(&source, |_| {
        vec![Update::Toc {
            field: 0,
            entries: vec![TocEntry::new(1, "H")],
            dirty: false,
        }]
    });
    let fields = target.edit().fields().unwrap();
    assert_eq!(fields.fields()[0].kind(), FieldKind::Complex);
    assert_eq!(fields.fields()[0].instruction(), r#"TOC \o "1-1""#);
    assert_eq!(fields.fields()[0].result(), "H");
    assert_eq!(fields.paragraphs()[0].text(), "before ");
}

#[test]
fn unsupported_shapes_are_refused_without_change() {
    // The separator sits inside a hyperlink: the paragraph cannot be cut.
    let source = Snapshot::from_xml(document(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:hyperlink w:anchor="x"><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ))
    .unwrap();
    let mut edit = source.edit();
    let fields = edit.fields().unwrap();
    assert!(!fields.accepts_paragraphs(0));
    assert!(!fields.accepts_text(0));
    assert!(
        edit.refresh_fields(
            &fields,
            &[Update::Toc {
                field: 0,
                entries: Vec::new(),
                dirty: false
            }]
        )
        .is_err()
    );
    assert!(edit.commit().unwrap().patch().is_noop());
    // A result holding a nested field is not plain text.
    let nested = Snapshot::from_xml(document(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF a</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ))
    .unwrap();
    assert!(!nested.edit().fields().unwrap().accepts_text(0));
}

#[test]
fn stale_indexes_and_bad_text_are_refused() {
    let source = Snapshot::from_xml(document(TOC_PLACEHOLDER)).unwrap();
    let mut edit = source.edit();
    let fields = edit.fields().unwrap();
    assert!(
        edit.refresh_fields(
            &fields,
            &[Update::Text {
                field: 0,
                text: "\u{1}".into()
            }]
        )
        .is_err()
    );
    edit.refresh_fields(
        &fields,
        &[Update::Text {
            field: 0,
            text: "x".into(),
        }],
    )
    .unwrap();
    assert!(
        edit.refresh_fields(
            &fields,
            &[Update::Text {
                field: 0,
                text: "y".into()
            }]
        )
        .is_err()
    );
}

#[test]
fn instructions_bookmark_names_and_removals_are_spliced() {
    // Record 0325: field instruction sets, bookmark renames and removals.
    let xml = document(&format!(
        r#"{TOC_PLACEHOLDER}<w:p><w:bookmarkStart w:id="4" w:name="Intro"/><w:r><w:t>Body</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>"#
    ));
    let source = Snapshot::from_xml(xml).unwrap();
    let target = commit(&source, |_| {
        vec![
            Update::Instruction {
                field: 0,
                instruction: r#" TOC \o "1-2" \h "#.into(),
            },
            Update::Instruction {
                field: 1,
                instruction: " NUMPAGES ".into(),
            },
            Update::RenameBookmark {
                bookmark: 0,
                name: "Start & end".into(),
            },
        ]
    });
    let fields = target.edit().fields().unwrap();
    assert_eq!(fields.fields()[0].instruction(), r#" TOC \o "1-2" \h "#);
    assert!(fields.fields()[0].is_dirty());
    assert_eq!(fields.fields()[1].instruction(), " NUMPAGES ");
    assert_eq!(fields.fields()[1].result(), "3");
    assert_eq!(fields.bookmarks()[0].name(), "Start & end");
    // An equal instruction and name change nothing.
    let mut edit = target.edit();
    let index = edit.fields().unwrap();
    let outcome = edit
        .refresh_fields(
            &index,
            &[
                Update::Instruction {
                    field: 1,
                    instruction: " NUMPAGES ".into(),
                },
                Update::RenameBookmark {
                    bookmark: 0,
                    name: "Start & end".into(),
                },
            ],
        )
        .unwrap();
    assert!(outcome.is_empty());
    // Removing a bookmark keeps its content.
    let removed = commit(&target, |_| vec![Update::RemoveBookmark { bookmark: 0 }]);
    let text = text(&removed);
    assert!(!text.contains("bookmarkStart") && !text.contains("bookmarkEnd"));
    assert!(text.contains("<w:t>Body</w:t>"));
}
