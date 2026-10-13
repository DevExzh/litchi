use std::sync::Arc;

use super::{FieldKind, OBJECT, Replacement, Story, part_story, scan, substitute_part};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn header(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W}">{body}</w:hdr>"#
    )
}

fn document(body: &str) -> String {
    format!(r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#)
}

fn story(xml: &str) -> Story {
    part_story("/word/header1.xml", Arc::new(xml.as_bytes().to_vec())).expect("story")
}

/// Replace every `{{key}}` of every paragraph with `value(key)`.
fn fill(xml: &str, value: impl Fn(&str) -> String) -> String {
    let story = story(xml);
    let mut replacements = Vec::new();
    for (index, paragraph) in story.paragraphs().iter().enumerate() {
        let text = paragraph.text();
        let mut from = 0;
        while let Some(start) = text[from..].find("{{").map(|offset| from + offset) {
            let Some(end) = text[start..].find("}}").map(|offset| start + offset + 2) else {
                break;
            };
            replacements.push(Replacement::new(
                index,
                start..end,
                value(&text[start + 2..end - 2]),
            ));
            from = end;
        }
    }
    let (after, _) = substitute_part(&story, xml.as_bytes(), &replacements).expect("substitute");
    String::from_utf8(after.unwrap_or_else(|| xml.as_bytes().to_vec())).expect("UTF-8")
}

#[test]
fn a_placeholder_split_across_runs_keeps_each_runs_formatting() {
    let xml = header(
        r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello {{na</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>me}}!</w:t></w:r></w:p>"#,
    );
    assert_eq!(story(&xml).paragraphs()[0].text(), "Hello {{name}}!");
    let out = fill(&xml, |_| "Ada".to_owned());
    assert!(
        out.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Hello Ada</w:t></w:r>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<w:r><w:rPr><w:i/></w:rPr><w:t>!</w:t></w:r>"#),
        "{out}"
    );
}

#[test]
fn a_placeholder_spanning_three_runs_empties_the_middle_one() {
    let xml = header(
        r#"<w:p><w:r><w:t>{{</w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>key</w:t></w:r><w:r><w:t>}} tail</w:t></w:r></w:p>"#,
    );
    let out = fill(&xml, |key| format!("<{key}> & co"));
    assert!(out.contains("<w:t>&lt;key&gt; &amp; co</w:t>"), "{out}");
    assert!(
        out.contains(r#"<w:rPr><w:u w:val="single"/></w:rPr><w:t></w:t>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<w:t xml:space="preserve"> tail</w:t>"#),
        "{out}"
    );
    assert_eq!(story(&out).paragraphs()[0].text(), "<key> & co tail");
}

#[test]
fn fixed_content_is_never_text() {
    let xml = header(r#"<w:p><w:r><w:t>{{a</w:t><w:tab/><w:t>}}</w:t><w:drawing/></w:r></w:p>"#);
    let story = story(&xml);
    let paragraph = &story.paragraphs()[0];
    assert_eq!(paragraph.text(), format!("{{{{a\t}}}}{OBJECT}"));
    assert!(!paragraph.is_text(0..paragraph.text().len()));
    assert!(paragraph.is_text(0..3));
    let error = substitute_part(&story, xml.as_bytes(), &[Replacement::new(0, 0..6, "x")]);
    assert!(error.is_err());
}

#[test]
fn tabs_and_line_feeds_become_run_content() {
    let xml = header(r#"<w:p><w:r><w:t>[{{v}}]</w:t></w:r></w:p>"#);
    let out = fill(&xml, |_| "a\tb\nc".to_owned());
    assert!(
        out.contains("<w:t>[a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c]</w:t>"),
        "{out}"
    );
    assert_eq!(story(&out).paragraphs()[0].text(), "[a\tb\nc]");
}

#[test]
fn a_complex_merge_field_becomes_one_formatted_run() {
    let xml = header(concat!(
        r#"<w:p><w:r><w:t xml:space="preserve">To </w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve"> MERGEFIELD First </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:rPr><w:b/></w:rPr><w:t>«First»</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:t>!</w:t></w:r></w:p>"#,
    ));
    let story = story(&xml);
    let paragraph = &story.paragraphs()[0];
    let [field] = paragraph.fields() else {
        panic!("one field");
    };
    assert_eq!(field.kind(), FieldKind::Complex);
    assert_eq!(field.instruction(), " MERGEFIELD First ");
    assert!(field.is_replaceable());
    let range = field.range();
    assert_eq!(
        &paragraph.text()[range.clone()],
        format!("{OBJECT}{OBJECT}«First»{OBJECT}")
    );
    let (after, outcome) =
        substitute_part(&story, xml.as_bytes(), &[Replacement::new(0, range, "Ada")]).unwrap();
    assert_eq!(outcome.fields(), 1);
    let out = String::from_utf8(after.unwrap()).unwrap();
    assert!(
        out.contains(r#"<w:t xml:space="preserve">To </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Ada</w:t></w:r><w:r><w:t>!</w:t></w:r>"#),
        "{out}"
    );
    assert!(!out.contains("fldChar"));
}

#[test]
fn a_simple_field_is_replaced_whole() {
    let xml = header(
        r#"<w:p><w:fldSimple w:instr=" MERGEFIELD &quot;Last Name&quot; "><w:r><w:rPr><w:i/></w:rPr><w:t>«Last Name»</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let story = story(&xml);
    let field = &story.paragraphs()[0].fields()[0];
    assert_eq!(field.kind(), FieldKind::Simple);
    assert_eq!(field.instruction(), r#" MERGEFIELD "Last Name" "#);
    assert!(field.is_replaceable());
    let (after, _) = substitute_part(
        &story,
        xml.as_bytes(),
        &[Replacement::new(0, field.range(), " Lovelace")],
    )
    .unwrap();
    let out = String::from_utf8(after.unwrap()).unwrap();
    assert!(
        out.contains(r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> Lovelace</w:t></w:r></w:p>"#),
        "{out}"
    );
}

#[test]
fn fields_with_other_content_are_not_replaceable() {
    // A bookmark inside the field, a field spanning a hyperlink, and a
    // nested field.
    let bookmark = header(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD a</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>x</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
    ));
    assert!(!story(&bookmark).paragraphs()[0].fields()[0].is_replaceable());
    let hyperlink = header(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD a</w:instrText></w:r>"#,
        r#"<w:hyperlink w:anchor="x"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>"#,
    ));
    assert!(!story(&hyperlink).paragraphs()[0].fields()[0].is_replaceable());
    let nested = header(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD a</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:instrText> = 1 "y" "n"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ));
    let story = story(&nested);
    let fields = story.paragraphs()[0].fields();
    assert_eq!(fields.len(), 2);
    let outer = fields
        .iter()
        .find(|field| field.instruction().starts_with("IF"))
        .unwrap();
    let inner = fields
        .iter()
        .find(|field| field.instruction() == "MERGEFIELD a")
        .unwrap();
    assert!(!outer.is_replaceable());
    assert!(inner.is_replaceable());
}

#[test]
fn a_field_that_spans_paragraphs_is_not_listed() {
    let xml = header(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:p>"#,
        r#"<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ));
    let story = story(&xml);
    assert!(
        story
            .paragraphs()
            .iter()
            .all(|paragraph| paragraph.fields().is_empty())
    );
}

#[test]
fn text_box_paragraphs_are_their_own_paragraphs() {
    let xml = header(concat!(
        r#"<w:p><w:r><w:t>{{a}}</w:t></w:r><w:r><w:drawing><w:txbxContent>"#,
        r#"<w:p><w:r><w:t>{{b}}</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>"#,
    ));
    let story = story(&xml);
    let texts: Vec<&str> = story
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text())
        .collect();
    assert_eq!(texts, ["{{b}}", format!("{{{{a}}}}{OBJECT}").as_str()]);
    let out = fill(&xml, str::to_uppercase);
    assert!(
        out.contains("<w:t>A</w:t>") && out.contains("<w:t>B</w:t>"),
        "{out}"
    );
}

#[test]
fn main_stories_list_only_addressable_body_blocks() {
    let xml = document(concat!(
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        r#"<w:customXml><w:p><w:r><w:t>hidden</w:t></w:r></w:p></w:customXml>"#,
        r#"<w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    ));
    let paragraphs = scan(xml.as_bytes(), true).unwrap();
    let texts: Vec<&str> = paragraphs
        .iter()
        .map(|paragraph| paragraph.text())
        .collect();
    assert_eq!(texts, ["one", "cell", "control"]);
}

#[test]
fn invalid_plans_are_refused() {
    let xml = header(r#"<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>"#);
    let story = story(&xml);
    let bytes = xml.as_bytes();
    // Overlapping, out of order, out of range, empty, control characters.
    for replacements in [
        vec![
            Replacement::new(0, 0..3, "x"),
            Replacement::new(0, 2..4, "y"),
        ],
        vec![
            Replacement::new(0, 3..4, "x"),
            Replacement::new(0, 0..1, "y"),
        ],
        vec![Replacement::new(0, 0..9, "x")],
        vec![Replacement::new(0, 1..1, "x")],
        vec![Replacement::new(0, 0..1, "\u{1}")],
        vec![Replacement::new(1, 0..1, "x")],
    ] {
        assert!(substitute_part(&story, bytes, &replacements).is_err());
    }
    // A stale story.
    assert!(substitute_part(&story, b"<w:hdr/>", &[Replacement::new(0, 0..1, "x")]).is_err());
    // Nothing to do.
    assert_eq!(substitute_part(&story, bytes, &[]).unwrap().0, None);
}

mod edits {
    use litchi_core::Position;

    use crate::document::ElementPath;
    use crate::document::substitute::Replacement;

    fn plan(story: &crate::document::substitute::Story) -> Vec<Replacement> {
        let mut replacements = Vec::new();
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            if let (Some(start), Some(end)) =
                (paragraph.text().find("{{"), paragraph.text().find("}}"))
            {
                replacements.push(Replacement::new(index, start..end + 2, "Ada"));
            }
        }
        replacements
    }

    #[test]
    fn main_document_substitution_commits_reversibly() {
        let mut builder = crate::Builder::new().unwrap();
        {
            let document = builder.document().unwrap();
            document.add_paragraph_with_text("Dear {{name}},");
            document.add_paragraph_with_text("unchanged");
            document.add_paragraph_with_text("Bye {{name}}");
        }
        let package = builder.build().unwrap();
        let mut edit = package.edit().unwrap();
        let story = edit.story().unwrap();
        let outcome = edit.substitute(&story, &plan(&story)).unwrap();
        assert_eq!((outcome.paragraphs(), outcome.texts()), (2, 2));
        // The story is now stale.
        assert!(edit.substitute(&story, &plan(&story)).is_err());
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (next, patch) = commit.into_parts();
        assert_eq!(
            next.document().unwrap().text().unwrap(),
            "Dear Ada,unchangedBye Ada"
        );
        let restored = next.apply(&patch.inverse()).unwrap();
        assert_eq!(
            restored.document().unwrap().text().unwrap(),
            package.document().unwrap().text().unwrap()
        );
        let mut bytes = next.clone();
        let reopened = crate::Package::from_bytes(&bytes.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.document().unwrap().text().unwrap(),
            "Dear Ada,unchangedBye Ada"
        );
        let _ = ElementPath::new([crate::document::ElementStep::Paragraph(Position::new(0))]);
    }

    #[test]
    fn an_empty_plan_is_an_exact_no_op() {
        let mut builder = crate::Builder::new().unwrap();
        builder.document().unwrap().add_paragraph_with_text("plain");
        let package = builder.build().unwrap();
        let mut edit = package.edit().unwrap();
        let story = edit.story().unwrap();
        assert!(edit.substitute(&story, &[]).unwrap().is_empty());
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
    }
}
