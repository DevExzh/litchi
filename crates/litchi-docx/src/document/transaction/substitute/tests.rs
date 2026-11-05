use std::sync::Arc;

use super::{FieldKind, OBJECT, Outcome, Replacement, Story, part_story, scan_paragraphs};
use crate::document::Result;

fn substitute_part(
    story: &Story,
    current: &[u8],
    replacements: &[Replacement],
) -> Result<(Option<Vec<u8>>, Outcome)> {
    super::substitute_part(story, current, replacements, |_run| {})
}

fn scan(xml: &[u8], main: bool) -> Result<Vec<super::Paragraph>> {
    scan_paragraphs(xml, main)
}

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
fn a_field_that_spans_paragraphs_is_listed_but_never_replaceable() {
    let xml = header(concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:p>"#,
        r#"<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ));
    let story = story(&xml);
    let [field] = story.paragraphs()[0].fields() else {
        panic!("the field is listed where it begins");
    };
    assert_eq!(field.instruction(), "TOC");
    assert!(!field.is_replaceable());
    assert!(story.paragraphs()[1].fields().is_empty());
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
        let restored = next.apply(&patch.inverse()).unwrap().into_package();
        assert_eq!(
            restored.document().unwrap().text().unwrap(),
            package.document().unwrap().text().unwrap()
        );
        let bytes = next.clone();
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

/// Record 0285: one primitive for find and merge, field state across
/// paragraphs, segments, scoped stories and formatted replacements.
mod unified {
    use litchi_core::Position;

    use super::super::{Replacement, Segment, Story};
    use super::{OBJECT, W};
    use crate::document::{ElementPath, ElementStep, Snapshot};

    fn snapshot(body: &str) -> Snapshot {
        Snapshot::from_xml(format!(
            r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
        ))
        .unwrap()
    }

    fn body(position: usize) -> ElementPath {
        ElementPath::new([ElementStep::Paragraph(Position::new(position))])
    }

    fn xml(snapshot: &Snapshot) -> String {
        String::from_utf8(snapshot.xml_bytes().to_vec()).unwrap()
    }

    fn segments(story: &Story, paragraph: usize) -> Vec<(String, usize, bool)> {
        let paragraph = &story.paragraphs()[paragraph];
        paragraph
            .segments()
            .map(|segment: Segment| {
                (
                    paragraph.text()[segment.range()].to_owned(),
                    segment.group(),
                    segment.is_scoped(),
                )
            })
            .collect()
    }

    /// A table of contents whose result spans three paragraphs, a heading
    /// after it, and a nested field inside an instruction.
    const TOC: &str = concat!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Intro</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>Intro body</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>Intro end</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"#,
    );

    #[test]
    fn field_results_read_as_text_in_every_paragraph_they_span() {
        let snapshot = snapshot(TOC);
        let story = snapshot.edit().story().unwrap();
        let texts: Vec<&str> = story
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect();
        assert_eq!(
            texts,
            [
                format!("{OBJECT}{OBJECT}Intro\t1").as_str(),
                "Intro body",
                format!("Intro end{OBJECT}").as_str(),
                "Intro",
            ]
        );
        // Every paragraph of the result reads the same through a scoped
        // story, which starts inside the field.
        for (position, text) in [(1, "Intro body"), (2, "Intro end")] {
            let scoped = snapshot.edit().story_within(&[body(position)]).unwrap();
            assert_eq!(scoped.paragraphs().len(), 1);
            assert!(scoped.paragraphs()[0].text().starts_with(text));
        }
    }

    #[test]
    fn rewriting_a_result_that_spans_paragraphs_keeps_the_field() {
        let snapshot = snapshot(TOC);
        let mut edit = snapshot.edit();
        let story = edit.story().unwrap();
        let mut replacements = Vec::new();
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            let text = paragraph.text();
            let mut from = 0;
            while let Some(at) = text[from..].find("Intro").map(|at| at + from) {
                replacements.push(Replacement::new(index, at..at + 5, "Preface"));
                from = at + 5;
            }
        }
        assert_eq!(replacements.len(), 4);
        edit.substitute(&story, &replacements).unwrap();
        let commit = edit.commit().unwrap();
        let after = xml(commit.snapshot());
        // Markers and the instruction are byte-for-byte where they were.
        for marker in [
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Preface</w:t></w:r>"#,
            r#"<w:r><w:t>Preface end</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ] {
            assert!(after.contains(marker), "{after}");
        }
        assert_eq!(after.matches("Intro").count(), 0);
        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.xml_bytes(), snapshot.xml_bytes());
    }

    #[test]
    fn text_inside_an_instruction_is_fixed_even_across_paragraphs() {
        // An IF whose instruction holds a nested field result, and whose
        // instruction continues into the next paragraph.
        let snapshot = snapshot(concat!(
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD a</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>nested</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>still code</w:t></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>shown</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        ));
        let edit = snapshot.edit();
        let story = edit.story().unwrap();
        assert!(story.paragraphs()[0].segments().next().is_none());
        let shown: Vec<String> = segments(&story, 1)
            .into_iter()
            .map(|(text, _, _)| text)
            .collect();
        assert_eq!(shown, ["shown"]);
        let scoped = edit.story_within(&[body(1)]).unwrap();
        assert_eq!(segments(&scoped, 0)[0].0, "shown");
        assert!(!scoped.paragraphs()[0].is_text(0..1));
    }

    #[test]
    fn segments_group_text_that_no_boundary_separates() {
        let snapshot = snapshot(
            r#"<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>a</w:t></w:r><w:bookmarkStart w:id="0" w:name="x"/><w:proofErr w:type="spellStart"/><w:r><w:lastRenderedPageBreak/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r><w:bookmarkEnd w:id="0"/><w:hyperlink w:anchor="x"><w:r><w:t>d</w:t></w:r></w:hyperlink><w:r><w:t>e</w:t></w:r><w:fldSimple w:instr=" DATE "><w:r><w:t>today</w:t></w:r></w:fldSimple><w:del w:id="1" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins w:id="2" w:author="a"><w:r><w:t>new</w:t></w:r></w:ins></w:p>"#,
        );
        let story = snapshot.edit().story().unwrap();
        let found = segments(&story, 0);
        let texts: Vec<&str> = found.iter().map(|(text, _, _)| text.as_str()).collect();
        assert_eq!(texts, ["a", "b", "c", "d", "e", "today", "new"]);
        let group = |index: usize| found[index].1;
        assert_eq!(group(0), group(1), "bookmarks and proofing marks join text");
        assert_ne!(group(1), group(2), "a tab separates text");
        assert_ne!(group(2), group(3), "a hyperlink starts a group");
        assert_ne!(group(3), group(4), "a hyperlink ends a group");
        assert_ne!(group(4), group(5), "a simple field starts a group");
        assert_ne!(group(5), group(6), "deleted text and insertions separate");
        assert!(found.iter().all(|(_, _, scoped)| *scoped));
    }

    #[test]
    fn scoped_stories_list_only_what_the_paths_name() {
        let snapshot = snapshot(concat!(
            r#"<w:p><w:r><w:t>zero</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>one </w:t></w:r><w:hyperlink w:anchor="a"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t> two</w:t></w:r></w:p>"#,
            r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p><w:p><w:r><w:t>cell two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        ));
        let edit = snapshot.edit();
        let hyperlink = body(1).child(ElementStep::Hyperlink(Position::new(0)));
        let story = edit.story_within(&[hyperlink]).unwrap();
        let scoped: Vec<(String, bool)> = segments(&story, 0)
            .into_iter()
            .map(|(text, _, scoped)| (text, scoped))
            .collect();
        assert_eq!(
            scoped,
            [
                ("one ".to_owned(), false),
                ("link".to_owned(), true),
                (" two".to_owned(), false)
            ]
        );
        let table = ElementPath::new([ElementStep::Table(Position::new(0))]);
        let story = edit.story_within(&[table, body(0), body(0)]).unwrap();
        let texts: Vec<&str> = story
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect();
        assert_eq!(texts, ["zero", "cell", "cell two"]);
        let missing = edit.story_within(&[body(9)]);
        assert!(missing.is_err());
    }

    #[test]
    fn formatted_replacements_split_and_format_runs() {
        let snapshot = snapshot(
            r#"<w:p><w:r><w:t>one</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>two three</w:t><w:tab/><w:t>four</w:t></w:r><w:hyperlink w:anchor="a"><w:r><w:t>link text</w:t></w:r></w:hyperlink><w:ins w:id="1" w:author="a"><w:r><w:t>last</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>untouched</w:t></w:r></w:p>"#,
        );
        let mut edit = snapshot.edit();
        let story = edit.story().unwrap();
        let text = story.paragraphs()[0].text().to_owned();
        let find = |needle: &str| {
            let at = text.find(needle).unwrap();
            at..at + needle.len()
        };
        let outcome = edit
            .substitute_formatted(
                &story,
                &[
                    Replacement::new(0, find("three"), "3").formatted(),
                    Replacement::keep(0, find("fo")).formatted(),
                    Replacement::keep(0, find("text")),
                    Replacement::keep(0, find("last")),
                ],
                |run| run.bold = Some(true),
            )
            .unwrap();
        assert_eq!(outcome.formatted(), 4);
        assert_eq!(outcome.paragraphs(), 1);
        let after = xml(edit.commit().unwrap().snapshot());
        assert!(after.contains(concat!(
            r#"<w:r><w:t>one</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">two </w:t></w:r>"#,
            r#"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>3</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:tab/></w:r>"#,
            r#"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>fo</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ur</w:t></w:r>"#,
        )), "{after}");
        assert!(after.contains(concat!(
            r#"<w:hyperlink w:anchor="a"><w:r><w:t xml:space="preserve">link </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r></w:hyperlink>"#,
            r#"<w:ins w:id="1" w:author="a"><w:r><w:rPr><w:b/></w:rPr><w:t>last</w:t></w:r></w:ins>"#,
        )), "{after}");
        assert!(after.contains("<w:p><w:r><w:t>untouched</w:t></w:r></w:p>"));
    }

    #[test]
    fn nested_and_body_paragraphs_change_together_and_invert() {
        let snapshot = snapshot(
            r#"<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>beta</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>gamma</w:t></w:r><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>delta</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>"#,
        );
        let mut edit = snapshot.edit();
        let story = edit.story().unwrap();
        let replacements: Vec<Replacement> = story
            .paragraphs()
            .iter()
            .enumerate()
            .map(|(index, paragraph)| {
                let end = paragraph.text().trim_end_matches(OBJECT).len();
                Replacement::new(index, 0..end, paragraph.text()[..end].to_uppercase())
            })
            .collect();
        assert_eq!(replacements.len(), 4);
        edit.substitute(&story, &replacements).unwrap();
        let commit = edit.commit().unwrap();
        let text = xml(commit.snapshot());
        for word in ["ALPHA", "BETA", "GAMMA", "DELTA"] {
            assert!(text.contains(&format!("<w:t>{word}</w:t>")), "{text}");
        }
        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.xml_bytes(), snapshot.xml_bytes());
    }

    #[test]
    fn part_stories_number_their_items_and_format_runs() {
        let xml = format!(
            r#"<w:comments xmlns:w="{W}"><w:comment w:id="4"><w:p><w:r><w:t>first note</w:t></w:r></w:p></w:comment><w:comment w:id="9"><w:p><w:r><w:t>second note</w:t></w:r></w:p><w:p><w:r><w:t>more</w:t></w:r></w:p></w:comment></w:comments>"#
        );
        let story = super::super::part_story(
            "/word/comments.xml",
            std::sync::Arc::new(xml.as_bytes().to_vec()),
        )
        .unwrap();
        let items: Vec<Option<usize>> = story
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.item())
            .collect();
        assert_eq!(items, [Some(0), Some(1), Some(1)]);
        let (after, outcome) = super::super::substitute_part(
            &story,
            xml.as_bytes(),
            &[Replacement::new(1, 0..6, "2nd").formatted()],
            |run| run.italic = Some(true),
        )
        .unwrap();
        assert_eq!(outcome.formatted(), 1);
        let after = String::from_utf8(after.unwrap()).unwrap();
        assert!(
            after.contains(r#"<w:r><w:rPr><w:i/></w:rPr><w:t>2nd</w:t></w:r><w:r><w:t xml:space="preserve"> note</w:t></w:r>"#),
            "{after}"
        );
    }

    #[test]
    fn malformed_requests_are_refused_without_change() {
        let snapshot = snapshot(r#"<w:p><w:r><w:t>ab</w:t></w:r></w:p>"#);
        let cases = [
            vec![Replacement::new(0, 1..1, "x")],
            vec![Replacement::new(0, 0..1, "\u{1}")],
            vec![Replacement::new(1, 0..1, "x")],
            vec![Replacement::new(0, 0..3, "x")],
            vec![
                Replacement::new(0, 1..2, "x"),
                Replacement::new(0, 0..1, "y"),
            ],
        ];
        for replacements in cases {
            let mut edit = snapshot.edit();
            let story = edit.story().unwrap();
            assert!(edit.substitute(&story, &replacements).is_err());
            assert_eq!(
                edit.commit().unwrap().snapshot().xml_bytes(),
                snapshot.xml_bytes()
            );
        }
    }
}

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn alternate_content_is_one_opaque_object() {
    // Record 0307: a text box with a fallback, each branch holding its own
    // paragraph. Neither branch's paragraph is listed, and the run reads the
    // wrapper as one object.
    let wrapper = r#"<mc:AlternateContent xmlns:wps="urn:wps"><mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>CHOICE {{a}}</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>FALLBACK {{a}}</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent>"#;
    let body = format!(
        r#"<w:p><w:r><w:t>ab</w:t>{wrapper}<w:t>cd</w:t></w:r><w:r><mc:AlternateContent/></w:r></w:p><w:p><w:r><w:t>{{{{a}}}}</w:t></w:r></w:p>"#
    );
    for (xml, main) in [
        (
            document(&body).replace("<w:document ", &format!(r#"<w:document xmlns:mc="{MC}" "#)),
            true,
        ),
        (
            header(&body).replace("<w:hdr ", &format!(r#"<w:hdr xmlns:mc="{MC}" "#)),
            false,
        ),
    ] {
        let paragraphs = scan(xml.as_bytes(), main).unwrap();
        assert_eq!(paragraphs.len(), 2, "{xml}");
        assert_eq!(paragraphs[0].text(), format!("ab{OBJECT}cd{OBJECT}"));
        assert_eq!(paragraphs[0].alternate_content(), [2, 7]);
        assert!(paragraphs[1].alternate_content().is_empty());
        // The two text pieces around the wrapper are separate groups.
        let groups: Vec<usize> = paragraphs[0]
            .segments()
            .map(|segment| segment.group())
            .collect();
        assert_eq!(groups.len(), 2);
        assert_ne!(groups[0], groups[1]);
    }
    // A merge fills the placeholder outside the wrapper and nowhere else.
    let xml = header(&body).replace("<w:hdr ", &format!(r#"<w:hdr xmlns:mc="{MC}" "#));
    let filled = fill(&xml, |_| "X".to_owned());
    assert!(filled.contains("CHOICE {{a}}") && filled.contains("FALLBACK {{a}}"));
    assert!(filled.contains("<w:t>X</w:t>"), "{filled}");
}
