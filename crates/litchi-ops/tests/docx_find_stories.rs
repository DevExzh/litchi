//! DOCX find/replace over fields and every story (record 0285): field
//! results are searched as OfficeCLI's find reads them, instructions never
//! are, a table of contents spanning paragraphs keeps its structure, and the
//! document root also searches headers, footers, footnotes and endnotes;
//! comments are searched through their own paths. Every case runs on both
//! owners and survives a save and reopen.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::{BatchItem, File, OpenOptions, Props, SourceMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORD: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/// A package with `body`, and a header, footer, footnotes, endnotes and
/// comments, each holding the word `draft`.
fn package(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr></w:body></w:document>"#
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Header draft </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>"#
    );
    let footer =
        format!(r#"<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Footer draft</w:t></w:r></w:p></w:ftr>"#);
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="{W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note draft</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    );
    let endnotes = format!(
        r#"<w:endnotes xmlns:w="{W}"><w:endnote w:id="1"><w:p><w:r><w:t>End draft</w:t></w:r></w:p></w:endnote></w:endnotes>"#
    );
    let comments = format!(
        r#"<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>First draft remark</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="B"><w:p><w:r><w:t>Second draft remark</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="{WORD}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="{WORD}.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="{WORD}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="{WORD}.endnotes+xml"/><Override PartName="/word/comments.xml" ContentType="{WORD}.comments+xml"/></Types>"#
    );
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="{R}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="{R}/footer" Target="footer1.xml"/><Relationship Id="rIdFn" Type="{R}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEn" Type="{R}/endnotes" Target="endnotes.xml"/><Relationship Id="rIdC" Type="{R}/comments" Target="comments.xml"/></Relationships>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", relationships.as_str()),
        ("word/document.xml", document.as_str()),
        ("word/header1.xml", header.as_str()),
        ("word/footer1.xml", footer.as_str()),
        ("word/footnotes.xml", footnotes.as_str()),
        ("word/endnotes.xml", endnotes.as_str()),
        ("word/comments.xml", comments.as_str()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(contents.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A table of contents whose result spans three paragraphs, then the
/// headings it lists.
const TOC: &str = concat!(
    r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Introduction</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:p>"#,
    r#"<w:p><w:r><w:t>Background</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>2</w:t></w:r></w:p>"#,
    r#"<w:p><w:r><w:t>Introduction to TOC</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    r#"<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>"#,
    r#"<w:p><w:r><w:t>Background</w:t></w:r></w:p>"#,
);

fn open(eager: bool, bytes: Vec<u8>) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .unwrap()
    } else {
        litchi_ops::open_bytes(bytes).unwrap()
    }
}

fn member(file: &File, name: &str) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name(name).unwrap();
    let mut contents = String::new();
    member.read_to_string(&mut contents).unwrap();
    contents
}

fn find(file: &File, path: &str, props: Props<'_>) -> (File, Option<usize>) {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    let commit = file.apply(&item).unwrap_or_else(|error| {
        let mut chain = String::new();
        let mut source = std::error::Error::source(&error);
        while let Some(inner) = source {
            chain.push_str(&format!(" <- {inner}"));
            source = inner.source();
        }
        panic!("{path}: {error}{chain}")
    });
    (commit.file().clone(), commit.value().matched())
}

fn replace(find_text: &str, replace_text: &str) -> Props<'static> {
    Props::new()
        .with("find", find_text.to_owned())
        .with("replace", replace_text.to_owned())
}

/// Field markers and instructions, in order.
fn field_markup(xml: &str) -> Vec<String> {
    let mut found = Vec::new();
    let mut rest = xml;
    while let Some(at) = rest
        .find("<w:fldChar")
        .or_else(|| rest.find("<w:instrText"))
    {
        let marker = rest.find("<w:fldChar");
        let instruction = rest.find("<w:instrText");
        let at = match (marker, instruction) {
            (Some(marker), Some(instruction)) => marker.min(instruction),
            _ => at,
        };
        let end = if rest[at..].starts_with("<w:instrText") {
            rest[at..].find("</w:instrText>").unwrap() + "</w:instrText>".len()
        } else {
            rest[at..].find("/>").unwrap() + 2
        };
        found.push(rest[at..at + end].to_owned());
        rest = &rest[at + end..];
    }
    found
}

#[test]
fn a_table_of_contents_spanning_paragraphs_keeps_its_field() {
    for eager in [false, true] {
        let file = open(eager, package(TOC));
        let before = member(&file, "word/document.xml");
        // The 0279 finding: a find over the body reaches the table's
        // entries. They are field results, searched as OfficeCLI and Word
        // search them, consistently in every paragraph of the result.
        let (after, matched) = find(&file, "/body", replace("Introduction", "Overview"));
        assert_eq!(matched, Some(3));
        let xml = member(&after, "word/document.xml");
        assert_eq!(xml.matches("Introduction").count(), 0, "{xml}");
        assert_eq!(xml.matches("Overview").count(), 3, "{xml}");
        // The field's markers and instruction are byte-for-byte unchanged.
        assert_eq!(field_markup(&xml), field_markup(&before));
        assert_eq!(field_markup(&xml).len(), 4);
        // An instruction is never searched, even where its text matches.
        let (after, matched) = find(&file, "/body", replace("TOC", "LIST"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains(r#"<w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText>"#)
        );
        assert!(xml.contains("<w:t>Introduction to LIST</w:t>"), "{xml}");
        // Paths inside the field's result read its state too.
        let (after, matched) = find(&file, "/body/p[2]", replace("Background", "Context"));
        assert_eq!(matched, Some(1));
        assert_eq!(
            field_markup(&member(&after, "word/document.xml")),
            field_markup(&before)
        );
        let reopened = open(eager, after.to_bytes().unwrap());
        assert!(
            member(&reopened, "word/document.xml").contains("<w:t>Context</w:t><w:t>")
                || member(&reopened, "word/document.xml").contains("<w:t>Context</w:t>")
        );
    }
}

#[test]
fn simple_field_results_are_searched_and_instructions_are_not() {
    for eager in [false, true] {
        let file = open(
            eager,
            package(
                r#"<w:p><w:r><w:t xml:space="preserve">Printed </w:t></w:r><w:fldSimple w:instr=" DATE \@ &quot;yyyy&quot; "><w:r><w:rPr><w:b/></w:rPr><w:t>2024</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> DATE</w:t></w:r></w:p>"#,
            ),
        );
        // `DATE` occurs in the instruction attribute and once as text.
        let (after, matched) = find(&file, "/body", replace("DATE", "date"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains(r#"w:instr=" DATE \@ &quot;yyyy&quot; ""#),
            "{xml}"
        );
        assert!(
            xml.contains(r#"<w:t xml:space="preserve"> date</w:t>"#),
            "{xml}"
        );
        // The result is text; the field edge separates it from the text
        // around it, so `Printed 2024` is not one match.
        let (after, matched) = find(&file, "/body", replace("2024", "2025"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains(r#"<w:fldSimple w:instr=" DATE \@ &quot;yyyy&quot; "><w:r><w:rPr><w:b/></w:rPr><w:t>2025</w:t></w:r></w:fldSimple>"#),
            "{xml}"
        );
        let (_, matched) = find(&file, "/body", replace("Printed 2024", "x"));
        assert_eq!(matched, Some(0));
    }
}

#[test]
fn complex_field_results_spanning_runs_are_one_text() {
    for eager in [false, true] {
        let file = open(
            eager,
            package(concat!(
                r#"<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> PAGE</w:t></w:r></w:p>"#,
            )),
        );
        let before = member(&file, "word/document.xml");
        // A match across the result's runs collapses into its first run.
        let (after, matched) = find(&file, "/body/p[1]", replace("12", "34"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>34</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t></w:t></w:r>"#),
            "{xml}"
        );
        assert_eq!(field_markup(&xml), field_markup(&before));
        // The instruction is not text; the text after the field is.
        let (after, matched) = find(&file, "/", replace("PAGE", "SIDE"));
        // Only the body's trailing word: the header's `PAGE` is a simple
        // field's instruction attribute.
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(xml.contains(r#"<w:instrText xml:space="preserve"> PAGE </w:instrText>"#));
        assert!(xml.contains(r#"<w:t xml:space="preserve"> SIDE</w:t>"#));
        // A match never crosses a field marker.
        let (_, matched) = find(&file, "/body", replace("Page 12", "x"));
        assert_eq!(matched, Some(0));
    }
}

#[test]
fn the_root_also_searches_headers_footers_and_notes() {
    for eager in [false, true] {
        let file = open(
            eager,
            package(r#"<w:p><w:r><w:t>Body draft</w:t></w:r></w:p>"#),
        );
        let (after, matched) = find(&file, "/", replace("draft", "final"));
        // Body, header, footer, footnote, endnote; never the comments.
        assert_eq!(matched, Some(5));
        assert!(member(&after, "word/document.xml").contains("<w:t>Body final</w:t>"));
        assert!(
            member(&after, "word/header1.xml")
                .contains(r#"<w:t xml:space="preserve">Header final </w:t>"#)
        );
        assert!(member(&after, "word/footer1.xml").contains("<w:t>Footer final</w:t>"));
        assert!(
            member(&after, "word/footnotes.xml")
                .contains(r#"<w:t xml:space="preserve"> Note final</w:t>"#)
        );
        assert!(member(&after, "word/endnotes.xml").contains("<w:t>End final</w:t>"));
        assert_eq!(
            member(&after, "word/comments.xml"),
            member(&file, "word/comments.xml")
        );
        // The body is only the main story.
        let (after, matched) = find(&file, "/body", replace("draft", "final"));
        assert_eq!(matched, Some(1));
        assert_eq!(
            member(&after, "word/header1.xml"),
            member(&file, "word/header1.xml")
        );
        // Save and reopen.
        let (after, _) = find(&file, "/", replace("draft", "final"));
        let reopened = open(eager, after.to_bytes().unwrap());
        assert!(member(&reopened, "word/footer1.xml").contains("Footer final"));
        let (_, matched) = find(&reopened, "/", replace("draft", "x"));
        assert_eq!(matched, Some(0));
    }
}

#[test]
fn format_properties_apply_in_every_story() {
    for eager in [false, true] {
        let file = open(
            eager,
            package(r#"<w:p><w:r><w:t>Body draft</w:t></w:r></w:p>"#),
        );
        let (after, matched) = find(
            &file,
            "/",
            Props::new().with("find", "draft").with("bold", true),
        );
        assert_eq!(matched, Some(5));
        let header = member(&after, "word/header1.xml");
        assert!(
            header.contains(r#"<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Header </w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">draft</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>"#),
            "{header}"
        );
        assert!(
            member(&after, "word/document.xml")
                .contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>draft</w:t></w:r>"#)
        );
        // Values are admitted before anything is read.
        let mut item = BatchItem::new("set");
        item.path = Some("/".into());
        item.props = Some(Props::new().with("find", "absent").with("bold", "maybe"));
        assert!(file.apply(&item).is_err());
    }
}

#[test]
fn comments_are_searched_through_their_paths() {
    for eager in [false, true] {
        let file = open(eager, package(r#"<w:p><w:r><w:t>Body</w:t></w:r></w:p>"#));
        let (after, matched) = find(&file, "/comments", replace("draft", "final"));
        assert_eq!(matched, Some(2));
        let comments = member(&after, "word/comments.xml");
        assert!(
            comments.contains("First final remark") && comments.contains("Second final remark")
        );
        let (after, matched) = find(&file, "/comments/comment[2]", replace("draft", "final"));
        assert_eq!(matched, Some(1));
        let comments = member(&after, "word/comments.xml");
        assert!(
            comments.contains("First draft remark") && comments.contains("Second final remark"),
            "{comments}"
        );
        assert_eq!(
            member(&after, "word/document.xml"),
            member(&file, "word/document.xml")
        );
    }
}

#[test]
fn text_boxes_are_searched_with_their_paragraph_scope_left_alone() {
    let body = concat!(
        r#"<w:p><w:r><w:t>Outer draft</w:t></w:r><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>Boxed draft</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>"#,
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell draft</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    for eager in [false, true] {
        let file = open(eager, package(body));
        // The body holds the outer paragraph, its text box and the cell, as
        // OfficeCLI's root find reads every descendant paragraph.
        let (after, matched) = find(&file, "/body", replace("draft", "final"));
        assert_eq!(matched, Some(3));
        let xml = member(&after, "word/document.xml");
        assert!(!xml.contains("draft"), "{xml}");
        // A paragraph's own scope leaves its text boxes alone.
        let (after, matched) = find(&file, "/body/p[1]", replace("draft", "final"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains("Outer final") && xml.contains("Boxed draft"),
            "{xml}"
        );
        // A cell scope reads only the cell.
        let (after, matched) = find(&file, "/body/tbl[1]/tr[1]/tc[1]", replace("draft", "final"));
        assert_eq!(matched, Some(1));
        let xml = member(&after, "word/document.xml");
        assert!(
            xml.contains("Cell final") && xml.contains("Outer draft"),
            "{xml}"
        );
        let reopened = open(eager, after.to_bytes().unwrap());
        assert!(member(&reopened, "word/document.xml").contains("Cell final"));
    }
}
