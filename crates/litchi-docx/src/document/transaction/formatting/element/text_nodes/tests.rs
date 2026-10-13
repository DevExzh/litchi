#![allow(
    clippy::single_range_in_vec_init,
    reason = "one isolated range per rewrite"
)]

use litchi_core::Position;

use super::super::super::super::{Error, Refusal, Snapshot};
use super::{ElementPath, ElementStep, ParagraphRewrite, TextNodeRewrite};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn snapshot(body: &str) -> Snapshot {
    Snapshot::from_xml(format!(
        r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#
    ))
    .unwrap()
}

fn body(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

fn nodes(snapshot: &Snapshot, path: &ElementPath) -> Vec<(String, usize)> {
    snapshot
        .edit()
        .text_nodes(path)
        .unwrap()
        .into_iter()
        .map(|node| (node.text, node.group))
        .collect()
}

fn xml(snapshot: &Snapshot) -> String {
    String::from_utf8(snapshot.xml_bytes().to_vec()).unwrap()
}

fn rewrite(
    snapshot: &Snapshot,
    rewrites: &[ParagraphRewrite],
) -> Result<(Snapshot, Vec<ElementPath>), Error> {
    let mut edit = snapshot.edit();
    let paths = edit.rewrite_text_nodes(rewrites)?;
    let commit = edit.commit()?;
    Ok((commit.snapshot().clone(), paths))
}

#[test]
fn markers_are_transparent_and_boundaries_separate_groups() {
    let snapshot = snapshot(
        r#"<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>a</w:t></w:r><w:bookmarkStart w:id="0" w:name="x"/><w:proofErr w:type="spellStart"/><w:r><w:lastRenderedPageBreak/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r><w:bookmarkEnd w:id="0"/><w:hyperlink w:anchor="x"><w:r><w:t>d</w:t></w:r></w:hyperlink><w:r><w:t>e</w:t></w:r></w:p>"#,
    );
    let found = nodes(&snapshot, &body(0));
    let texts: Vec<&str> = found.iter().map(|(text, _)| text.as_str()).collect();
    assert_eq!(texts, ["a", "b", "c", "d", "e"]);
    assert_eq!(
        found[0].1, found[1].1,
        "bookmarks and proofing marks join text"
    );
    assert_ne!(found[1].1, found[2].1, "a tab separates text");
    assert_ne!(found[2].1, found[3].1, "a hyperlink starts a group");
    assert_ne!(found[3].1, found[4].1, "a hyperlink ends a group");
}

#[test]
fn field_results_deleted_text_and_simple_fields_are_not_listed() {
    let snapshot = snapshot(
        r#"<w:p><w:r><w:t>keep</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:del w:id="1" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del><w:fldSimple w:instr=" DATE "><w:r><w:t>today</w:t></w:r></w:fldSimple><w:ins w:id="2" w:author="a"><w:r><w:t>new</w:t></w:r></w:ins></w:p>"#,
    );
    let texts: Vec<String> = nodes(&snapshot, &body(0))
        .into_iter()
        .map(|(text, _)| text)
        .collect();
    assert_eq!(texts, ["keep", "new"]);
}

#[test]
fn rewriting_keeps_every_other_byte() {
    let snapshot = snapshot(
        r#"<w:p><w:r w:rsidR="00AB"><w:rPr><w:b/></w:rPr><w:t>Hello </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">wor</w:t></w:r><w:r><w:t>ld &amp; more</w:t></w:r></w:p><w:p><w:r><w:t>untouched</w:t></w:r></w:p>"#,
    );
    let before = xml(&snapshot);
    let (after, paths) = rewrite(
        &snapshot,
        &[ParagraphRewrite::new(
            body(0),
            vec![
                TextNodeRewrite::new(0, "Hello <earth>"),
                TextNodeRewrite::new(1, ""),
                TextNodeRewrite::new(2, " & more"),
            ],
        )],
    )
    .unwrap();
    assert!(paths.is_empty());
    let after = xml(&after);
    assert!(after.contains(
        r#"<w:r w:rsidR="00AB"><w:rPr><w:b/></w:rPr><w:t>Hello &lt;earth&gt;</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"></w:t></w:r><w:r><w:t xml:space="preserve"> &amp; more</w:t></w:r>"#
    ));
    // The untouched paragraph and everything around the paragraph stay.
    let tail = "<w:p><w:r><w:t>untouched</w:t></w:r></w:p></w:body></w:document>";
    assert!(before.ends_with(tail) && after.ends_with(tail));
}

#[test]
fn isolation_splits_runs_and_reports_their_paths() {
    let snapshot = snapshot(
        r#"<w:p><w:r><w:t>one</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>two three</w:t><w:tab/><w:t>four</w:t></w:r><w:hyperlink w:anchor="a"><w:r><w:t>link text</w:t></w:r></w:hyperlink><w:r><w:t>last</w:t></w:r></w:p>"#,
    );
    let (after, paths) = rewrite(
        &snapshot,
        &[ParagraphRewrite::new(
            body(0),
            vec![
                TextNodeRewrite::new(1, "two 3").isolating(vec![4..5]),
                TextNodeRewrite::new(2, "four").isolating(vec![0..2]),
                TextNodeRewrite::new(3, "link text").isolating(vec![5..9]),
                TextNodeRewrite::new(4, "last").isolating(vec![0..4]),
            ],
        )],
    )
    .unwrap();
    let paragraph = body(0);
    let hyperlink = paragraph.child(ElementStep::Hyperlink(Position::new(0)));
    let run =
        |base: &ElementPath, position: usize| base.child(ElementStep::Run(Position::new(position)));
    assert_eq!(
        paths,
        [
            run(&paragraph, 2),
            run(&paragraph, 4),
            run(&hyperlink, 1),
            run(&paragraph, 6),
        ]
    );
    let after_xml = xml(&after);
    assert!(after_xml.contains(
        r#"<w:r><w:t>one</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">two </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>3</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>fo</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ur</w:t></w:r>"#
    ));
    assert!(after_xml.contains(
        r#"<w:hyperlink w:anchor="a"><w:r><w:t xml:space="preserve">link </w:t></w:r><w:r><w:t>text</w:t></w:r></w:hyperlink><w:r><w:t>last</w:t></w:r>"#
    ));
    // Every reported path names the isolated text.
    let edit = after.edit();
    for (path, text) in paths.iter().zip(["3", "fo", "text", "last"]) {
        let bytes = std::str::from_utf8(edit.element_xml(path).unwrap()).unwrap();
        assert!(bytes.contains(&format!(">{text}</w:t>")), "{bytes}");
    }
}

#[test]
fn isolation_needs_an_addressable_run() {
    let snapshot = snapshot(
        r#"<w:p><w:ins w:id="1" w:author="a"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>"#,
    );
    // Text-only rewrites work inside a tracked insertion.
    let (after, _) = rewrite(
        &snapshot,
        &[ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(0, "changed")],
        )],
    )
    .unwrap();
    assert!(xml(&after).contains("<w:t>changed</w:t>"));
    let error = rewrite(
        &snapshot,
        &[ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(0, "inserted").isolating(vec![0..2])],
        )],
    )
    .unwrap_err();
    assert!(matches!(
        error,
        Error::Refused {
            reason: Refusal::ComplexRun,
            ..
        }
    ));
}

#[test]
fn nested_and_body_paragraphs_rewrite_together_and_invert() {
    let snapshot = snapshot(
        r#"<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>beta</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>gamma</w:t></w:r></w:p>"#,
    );
    let cell = ElementPath::new([
        ElementStep::Table(Position::new(0)),
        ElementStep::Row(Position::new(0)),
        ElementStep::Cell(Position::new(0)),
        ElementStep::Paragraph(Position::new(0)),
    ]);
    assert_eq!(nodes(&snapshot, &cell)[0].0, "beta");
    let mut edit = snapshot.edit();
    edit.rewrite_text_nodes(&[
        ParagraphRewrite::new(body(1), vec![TextNodeRewrite::new(0, "GAMMA")]),
        ParagraphRewrite::new(cell.clone(), vec![TextNodeRewrite::new(0, "BETA")]),
        ParagraphRewrite::new(body(0), vec![TextNodeRewrite::new(0, "ALPHA")]),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    let text = xml(commit.snapshot());
    for word in ["ALPHA", "BETA", "GAMMA"] {
        assert!(text.contains(&format!("<w:t>{word}</w:t>")), "{text}");
    }
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), snapshot.xml_bytes());
}

#[test]
fn malformed_requests_are_refused_without_change() {
    let snapshot = snapshot(r#"<w:p><w:r><w:t>ab</w:t></w:r></w:p>"#);
    let cases = [
        vec![ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(1, "x")],
        )],
        vec![ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(0, "a\tb")],
        )],
        vec![ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(0, "ab").isolating(vec![1..1])],
        )],
        vec![ParagraphRewrite::new(
            body(0),
            vec![TextNodeRewrite::new(0, "é").isolating(vec![0..1])],
        )],
        vec![
            ParagraphRewrite::new(body(0), vec![TextNodeRewrite::new(0, "x")]),
            ParagraphRewrite::new(body(0), vec![TextNodeRewrite::new(0, "y")]),
        ],
        vec![ParagraphRewrite::new(
            body(3),
            vec![TextNodeRewrite::new(0, "x")],
        )],
    ];
    for rewrites in cases {
        let mut edit = snapshot.edit();
        assert!(edit.rewrite_text_nodes(&rewrites).is_err());
        assert_eq!(
            edit.commit().unwrap().snapshot().xml_bytes(),
            snapshot.xml_bytes()
        );
    }
}
