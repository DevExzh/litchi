//! No silent content loss on `text` sets (record 0304).
//!
//! Setting the `text` of a container (paragraph, run, shape, cell, note)
//! replaces its text. Content inside the container that is not text — a
//! picture, a frame, a field, a bookmark, a comment anchor, a note
//! reference — must either survive the set or make it fail with a typed
//! error naming it. Each module sets text on a container holding such
//! content in one format and asserts which of the two happens.

#![allow(clippy::unwrap_used, reason = "tests")]

#[cfg(any(
    feature = "odt",
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "doc",
    feature = "rtf",
    feature = "ppt",
    feature = "odp",
    feature = "odg",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
#[allow(dead_code, reason = "each format uses a subset of the helpers")]
mod support {
    use litchi_ops::{File, open_bytes};

    pub(crate) fn fixture(relative: &str) -> Vec<u8> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative);
        std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
    }

    pub(crate) fn open(relative: &str) -> File {
        open_bytes(fixture(relative)).unwrap_or_else(|error| panic!("{relative}: {error}"))
    }

    /// Publish and reopen, so the assertions read the saved bytes.
    pub(crate) fn reopen(file: &File) -> File {
        open_bytes(file.to_bytes().unwrap()).unwrap()
    }

    pub(crate) fn text(file: &File, path: &str) -> String {
        file.get(path)
            .unwrap_or_else(|error| panic!("{path}: {error}"))
            .view()
            .text()
            .unwrap()
            .unwrap_or_default()
            .into_owned()
    }

    /// One member of a ZIP package, as text.
    pub(crate) fn member(file: &File, name: &str) -> String {
        use std::io::Read as _;
        let bytes = file.to_bytes().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut text = String::new();
        entry.read_to_string(&mut text).unwrap();
        text
    }

    /// One member of a ZIP package, as bytes.
    pub(crate) fn member_bytes(file: &File, name: &str) -> Vec<u8> {
        use std::io::Read as _;
        let bytes = file.to_bytes().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut out = Vec::new();
        entry.read_to_end(&mut out).unwrap();
        out
    }

    /// `bytes` (a ZIP package) with member `name` rewritten by `change`.
    pub(crate) fn rewrite(bytes: &[u8], name: &str, change: impl Fn(&str) -> String) -> Vec<u8> {
        use std::io::{Read as _, Write as _};
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).unwrap();
            let entry_name = entry.name().to_owned();
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).unwrap();
            if entry_name == name {
                contents = change(std::str::from_utf8(&contents).unwrap()).into_bytes();
            }
            zip.start_file(entry_name, options).unwrap();
            zip.write_all(&contents).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    /// The error code of a refused set, after checking the snapshot is
    /// untouched.
    pub(crate) fn refused(file: &File, path: &str, value: &str) -> litchi_ops::Error {
        let before = file.to_bytes().unwrap();
        let error = match file.set(path, litchi_ops::props! { "text" => value }) {
            Ok(_) => panic!("{path}: the text set was admitted"),
            Err(error) => error,
        };
        assert_eq!(file.to_bytes().unwrap(), before);
        error
    }

    pub(crate) fn count(file: &File, query: &str) -> usize {
        file.query(query)
            .unwrap_or_else(|error| panic!("{query}: {error}"))
            .count()
    }
}

/// ODT, OTH and ODM share the ODT body edit: frames, notes, bookmarks and
/// other elements without text keep their character offset; text fields are
/// refused.
#[cfg(feature = "odt")]
mod odt_family {
    use litchi_ops::{Code, props};

    use super::support::{count, member, open, reopen, text};

    #[test]
    fn odt_paragraph_text_keeps_its_frames() {
        let file = open("odf/corpus/writer-images-frames.odt");
        let frames = count(&file, "frame");
        assert_eq!(frames, 5);
        let old = text(&file, "/body/p[2]");
        let changed = reopen(
            &file
                .set("/body/p[2]", props! { "text" => "New text" })
                .unwrap(),
        );
        assert_eq!(text(&changed, "/body/p[2]"), "New text");
        assert_eq!(count(&changed, "frame"), frames);
        assert_eq!(count(&changed, "/body/p[2]/frame"), 1);
        // Setting the old text back restores the source content.
        let restored = changed.set("/body/p[2]", props! { "text" => old }).unwrap();
        assert_eq!(
            member(&restored, "content.xml"),
            member(&file, "content.xml")
        );
    }

    #[cfg(all(feature = "oth", feature = "odm"))]
    #[test]
    fn oth_and_odm_paragraph_text_keeps_inline_frames() {
        for name in ["unified-ops/0294/frames.oth", "unified-ops/0294/frames.odm"] {
            let file = open(name);
            assert_eq!(count(&file, "/body/p[1]/frame"), 3, "{name}");
            let changed = reopen(
                &file
                    .set("/body/p[1]", props! { "text" => "Other text" })
                    .unwrap(),
            );
            assert_eq!(text(&changed, "/body/p[1]"), "Other text", "{name}");
            assert_eq!(count(&changed, "/body/p[1]/frame"), 3, "{name}");
            assert!(
                member(&changed, "content.xml").contains("Inside the box"),
                "{name}"
            );
            // The frames sit after "Frames", at offset 6, in both texts.
            let restored = changed
                .set("/body/p[1]", props! { "text" => "Frames" })
                .unwrap();
            assert_eq!(
                member(&restored, "content.xml"),
                member(&file, "content.xml"),
                "{name}: the old text restores the source"
            );
        }
    }

    #[test]
    fn odt_paragraph_text_refuses_text_fields() {
        let file = open("odfdo/tests/samples/variable.odt");
        // "Today is <text:date>…</text:date> <text:time>…</text:time>."
        let error = file
            .set("/body/p[1]", props! { "text" => "Today" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(error.message().contains("text:date"), "{}", error.message());
        // The same refusal inside a batch leaves the snapshot untouched.
        let before = file.to_bytes().unwrap();
        let result = file.edit(|tx| {
            tx.set("/body/p[1]", props! { "text" => "Today" })?;
            Ok(())
        });
        assert!(result.is_err());
        assert_eq!(file.to_bytes().unwrap(), before);
        // A paragraph whose only element is a zero-text reference mark is
        // set, and the mark is kept; the field after it is refused.
        let error = file
            .set("/body/p[4]", props! { "text" => "Page" })
            .unwrap_err();
        assert!(
            error.message().contains("text:reference-ref"),
            "{}",
            error.message()
        );
    }
}

/// DOCX: a paragraph or run text set fills the existing text slots and
/// keeps every other run child (drawings, pictures, note and comment
/// references); paragraph-level wrappers and anchors (bookmarks, comment
/// ranges, simple fields, hyperlinks, content controls) and complex fields
/// are refused. A cell text set keeps only its first paragraph and is
/// refused when a later block or that paragraph holds anything but text.
#[cfg(feature = "docx")]
mod docx {
    use litchi_ops::{Code, File, open_bytes, props};

    use super::support::{fixture, member, refused, rewrite, text};

    const PICT: &str = r#"<w:r><w:pict><v:rect style="width:10pt;height:10pt"/></w:pict></w:r>"#;

    /// `comment.docx` with `body` as its body (the comments part stays).
    fn document(body: &str) -> File {
        let bytes = rewrite(
            &fixture("ooxml/docx/comment.docx"),
            "word/document.xml",
            |xml| {
                let start = xml.find("<w:body>").unwrap() + "<w:body>".len();
                let end = xml.find("<w:sectPr>").unwrap();
                format!("{}{body}{}", &xml[..start], &xml[end..])
            },
        );
        open_bytes(bytes).unwrap()
    }

    fn body(file: &File) -> String {
        member(file, "word/document.xml")
    }

    #[test]
    fn paragraph_text_keeps_pictures_in_runs() {
        let file = document(&format!(
            "<w:p><w:r><w:t>Hello </w:t></w:r>{PICT}<w:r><w:t>world</w:t></w:r></w:p>"
        ));
        let changed = file
            .set("/body/p[1]", props! { "text" => "Changed" })
            .unwrap();
        assert_eq!(text(&changed, "/body/p[1]"), "Changed");
        assert!(body(&changed).contains("<w:pict>"), "{}", body(&changed));
        let run = file
            .set("/body/p[1]/r[1]", props! { "text" => "Hi " })
            .unwrap();
        assert!(body(&run).contains("<w:pict>"));
    }

    #[test]
    fn paragraph_text_keeps_a_comment_reference_in_a_run() {
        let file = document(r#"<w:p><w:r><w:t>Hi</w:t><w:commentReference w:id="0"/></w:r></w:p>"#);
        let changed = file
            .set("/body/p[1]", props! { "text" => "Annotated" })
            .unwrap();
        assert!(body(&changed).contains(r#"<w:commentReference w:id="0"/>"#));
        assert_eq!(text(&changed, "/body/p[1]"), "Annotated");
        // The fixture's own paragraph has a reference and no text slot.
        let file = open_bytes(fixture("ooxml/docx/comment.docx")).unwrap();
        let error = refused(&file, "/body/p[1]", "Annotated");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
    }

    #[test]
    fn paragraph_anchors_and_fields_are_refused() {
        // Direct paragraph bookmark markers have no visible text and are
        // retained byte-for-byte by the paragraph writer.
        let bookmarked = document(
            r#"<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
        );
        let changed = bookmarked
            .set("/body/p[1]", props! { "text" => "Replaced" })
            .unwrap();
        assert_eq!(text(&changed, "/body/p[1]"), "Replaced");
        assert!(body(&changed).contains(r#"<w:bookmarkStart w:id="1" w:name="b"/>"#));
        assert!(body(&changed).contains(r#"<w:bookmarkEnd w:id="1"/>"#));

        for paragraph in [
            r#"<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        ] {
            let file = document(paragraph);
            let error = refused(&file, "/body/p[1]", "Replaced");
            assert_eq!(
                error.code(),
                Code::UnsupportedCapability,
                "{paragraph}: {error}"
            );
        }
    }

    #[test]
    fn cell_text_refuses_content_it_would_drop() {
        let cell = |content: &str| {
            document(&format!(
                "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>\
                 <w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>{content}</w:tc></w:tr></w:tbl><w:p/>"
            ))
        };
        // Plain paragraphs are text: the set merges them into one.
        let file = cell("<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>");
        let changed = file
            .set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "New" })
            .unwrap();
        assert_eq!(text(&changed, "/body/tbl[1]/tr[1]/tc[1]"), "New");
        // A picture in a later paragraph, a nested table, a bookmark in the
        // first paragraph: refused, not dropped.
        for content in [
            format!("<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p>{PICT}</w:p>"),
            "<w:p><w:r><w:t>one</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>".to_owned(),
            r#"<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#.to_owned(),
        ] {
            let file = cell(&content);
            let error = refused(&file, "/body/tbl[1]/tr[1]/tc[1]", "New");
            assert_ne!(error.code(), Code::InternalError, "{content}: {error}");
        }
        // A picture in the first paragraph is kept.
        let file = cell(&format!("<w:p><w:r><w:t>one</w:t></w:r>{PICT}</w:p>"));
        let changed = file
            .set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "New" })
            .unwrap();
        assert!(body(&changed).contains("<w:pict>"));
    }
}

/// PPTX: shape, paragraph and cell text sets replace runs and line breaks;
/// a field (`a:fld`) in the replaced range is refused, never turned into
/// plain text or blanked. A run text set on the field itself sets its
/// cached result.
#[cfg(feature = "pptx")]
mod pptx {
    use litchi_ops::{Code, File, OpenOptions, SourceMode, props};

    use super::support::{member, refused, rewrite, text};

    const FIELD: &str = r#"<a:fld id="{385F08E6-D5BD-471D-B4C9-27BB3DE84C9D}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>"#;

    /// A builder deck (one text box "alpha", one 2×2 table) with `change`
    /// applied to its slide XML, opened through both owner routes.
    fn decks(change: impl Fn(&str) -> String) -> Vec<File> {
        let mut builder = litchi_pptx::builder::Builder::new().unwrap();
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box("alpha", 0, 0, 1_828_800, 914_400);
        let package = builder.build().unwrap();
        let mut edit = package.edit().unwrap();
        edit.add_table(
            0usize,
            &[
                vec!["a1".to_owned(), "b1".to_owned()],
                vec!["a2".to_owned(), "b2".to_owned()],
            ],
            (0, 1_000_000, 2_000_000, 600_000),
        )
        .unwrap();
        let bytes = edit.commit().unwrap().into_package().to_bytes().unwrap();
        let bytes = rewrite(&bytes, "ppt/slides/slide1.xml", change);
        [SourceMode::Eager, SourceMode::Auto]
            .into_iter()
            .map(|mode| {
                litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().source_mode(mode))
                    .unwrap()
            })
            .collect()
    }

    fn slide(file: &File) -> String {
        member(file, "ppt/slides/slide1.xml")
    }

    /// Put a field after the run holding `text`.
    fn field_after(text: &str) -> impl Fn(&str) -> String + '_ {
        move |xml: &str| {
            let at = xml.find(&format!("{text}</a:t></a:r>")).unwrap()
                + text.len()
                + "</a:t></a:r>".len();
            format!("{}{FIELD}{}", &xml[..at], &xml[at..])
        }
    }

    #[test]
    fn shape_and_paragraph_text_refuse_fields() {
        for file in decks(field_after("alpha")) {
            assert_eq!(text(&file, "/slide[1]/shape[1]"), "alpha7");
            for path in ["/slide[1]/shape[1]", "/slide[1]/shape[1]/paragraph[1]"] {
                let error = refused(&file, path, "beta");
                assert_eq!(error.code(), Code::UnsupportedCapability, "{path}: {error}");
            }
            // The run before the field is set; the field stays.
            let changed = file
                .set(
                    "/slide[1]/shape[1]/paragraph[1]/run[1]",
                    props! { "text" => "beta" },
                )
                .unwrap();
            assert_eq!(text(&changed, "/slide[1]/shape[1]"), "beta7");
            assert!(slide(&changed).contains(r#"type="slidenum""#));
        }
    }

    #[test]
    fn cell_text_refuses_fields() {
        for file in decks(field_after("a2")) {
            let error = refused(&file, "/slide[1]/table[1]/tr[2]/tc[1]", "x");
            assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        }
    }

    #[test]
    fn line_breaks_and_paragraphs_are_text() {
        let breaks = |xml: &str| {
            xml.replacen(
                "alpha</a:t></a:r>",
                "alpha</a:t></a:r><a:br/><a:r><a:t>omega</a:t></a:r>",
                1,
            )
        };
        for file in decks(breaks) {
            assert_eq!(
                text(&file, "/slide[1]/shape[1]/paragraph[1]"),
                "alpha\nomega"
            );
            let changed = file
                .set(
                    "/slide[1]/shape[1]/paragraph[1]",
                    props! { "text" => "one\ntwo" },
                )
                .unwrap();
            assert_eq!(
                text(&changed, "/slide[1]/shape[1]/paragraph[1]"),
                "one\ntwo"
            );
            let changed = file
                .set(
                    "/slide[1]/table[1]/tr[2]/tc[1]",
                    props! { "text" => "x\ny" },
                )
                .unwrap();
            // Each line is a paragraph of the cell.
            assert!(slide(&changed).contains("<a:t>y</a:t>"));
        }
    }
}

/// ODP: shape, cell and notes text sets rebuild the story's paragraphs,
/// and a paragraph or run text set replaces its content, spans and their
/// formatting included; a field, a hyperlink, a frame, a note or an
/// annotation inside the replaced content is refused.
#[cfg(feature = "odp")]
mod odp {
    use litchi_ops::{Code, open_bytes, props};

    use super::support::{fixture, refused, rewrite, text};

    #[test]
    fn fields_and_links_are_refused() {
        let bytes = rewrite(
            &fixture("odf/corpus/impress-basic.odp"),
            "content.xml",
            |xml| {
                xml.replacen(
                r#"<text:p text:style-name="P1">OL</text:p>"#,
                r#"<text:p text:style-name="P1">OL <text:page-number>1</text:page-number></text:p>"#,
                1,
            )
            .replacen(
                r#"<text:span text:style-name="T1">Eins</text:span>"#,
                r#"<text:span text:style-name="T1">Eins <text:a xlink:href="https://example.com/">link</text:a></text:span>"#,
                1,
            )
            },
        );
        let file = open_bytes(bytes).unwrap();
        assert_eq!(text(&file, "/slide[1]/shape[1]"), "OL 1");
        for path in [
            "/slide[1]/shape[1]",
            "/slide[1]/shape[1]/paragraph[1]",
            "/slide[1]/shape[2]",
            "/slide[1]/shape[2]/paragraph[1]",
            "/slide[1]/shape[2]/paragraph[1]/run[1]",
        ] {
            let error = refused(&file, path, "New");
            assert_eq!(error.code(), Code::UnsupportedCapability, "{path}: {error}");
        }
        // Spans are formatting of the text: replaced.
        let changed = file
            .set(
                "/slide[1]/shape[2]/paragraph[2]",
                props! { "text" => "Two" },
            )
            .unwrap();
        assert_eq!(text(&changed, "/slide[1]/shape[2]/paragraph[2]"), "Two");
    }
}

/// ODS: a cell value set replaces the cell's text. A hyperlink in the cell
/// (`text:a`) is refused unless `clear` removes it explicitly; rich
/// content the flat cell model cannot carry (spans, annotations, frames,
/// several paragraphs) is refused by the leaf for the whole row.
#[cfg(feature = "ods")]
mod ods {
    use litchi_ops::{Code, open_bytes, props};

    use super::support::{fixture, member, rewrite};

    const LINKED: &str = "libreoffice-core/sc/qa/unit/data/ods/external_hyperlink.ods";

    #[test]
    fn cell_values_refuse_to_drop_hyperlinks() {
        let file = open_bytes(fixture(LINKED)).unwrap();
        let before = file.to_bytes().unwrap();
        let error = file
            .set("/sheet[1]/A1", props! { "value" => "plain" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        assert_eq!(file.to_bytes().unwrap(), before);
        // The same value is an exact no-op that keeps the link.
        let same = file
            .set("/sheet[1]/A1", props! { "value" => "hyperlink with #" })
            .unwrap();
        assert!(member(&same, "content.xml").contains("<text:a "));
        // `clear` removes the link explicitly; the value set then applies.
        let cleared = file
            .set("/sheet[1]/A1", props! { "clear" => true })
            .unwrap()
            .set("/sheet[1]/A1", props! { "value" => "plain" })
            .unwrap();
        assert!(!member(&cleared, "content.xml").contains("<text:a "));
    }

    #[test]
    fn rich_cells_are_refused() {
        let bytes = rewrite(&fixture(LINKED), "content.xml", |xml| {
            let start = xml.find("<text:a ").unwrap();
            let end = xml.find("</text:a>").unwrap() + "</text:a>".len();
            format!(
                "{}<text:span text:style-name=\"T1\">rich</text:span> text{}",
                &xml[..start],
                &xml[end..]
            )
        });
        let file = open_bytes(bytes).unwrap();
        let error = file
            .set("/sheet[1]/A1", props! { "value" => "plain" })
            .unwrap_err();
        assert_ne!(error.code(), Code::InternalError, "{error}");
    }
}

/// XLSX and XLSB: a cell value set replaces the cell's value. The runs of
/// a rich-text cell are the formatting of that value and are replaced with
/// it (the new value is plain). Comments, notes and hyperlinks are parts
/// and records of the sheet, not of the cell value, and survive.
#[cfg(any(feature = "xlsx", feature = "xlsb"))]
mod workbook {
    use litchi_ops::props;

    use super::support::{member, open, text};

    #[cfg(feature = "xlsx")]
    #[test]
    fn xlsx_rich_cells_become_plain_and_comments_stay() {
        let file = open("ooxml/xlsx/TextColor.xlsx");
        assert_eq!(text(&file, "/sheet[1]/D6"), "Red Green");
        let changed = file
            .set("/sheet[1]/D6", props! { "value" => "Plain" })
            .unwrap();
        assert_eq!(text(&changed, "/sheet[1]/D6"), "Plain");
        let file = open("ooxml/xlsx/SimpleWithComments.xlsx");
        let changed = file
            .set("/sheet[1]/B1", props! { "value" => "uno" })
            .unwrap();
        assert_eq!(text(&changed, "/sheet[1]/B1"), "uno");
        for part in ["xl/comments1.xml", "xl/drawings/vmlDrawing1.vml"] {
            assert_eq!(member(&changed, part), member(&file, part), "{part}");
        }
    }

    #[cfg(feature = "xlsb")]
    #[test]
    fn xlsb_cells_keep_comments_and_hyperlinks() {
        let file = open("ooxml/xlsb/comments.xlsb");
        let changed = file
            .set("/sheet[1]/B1", props! { "value" => "uno" })
            .unwrap();
        assert_eq!(text(&changed, "/sheet[1]/B1"), "uno");
        assert_eq!(
            super::support::member_bytes(&changed, "xl/comments1.bin"),
            super::support::member_bytes(&file, "xl/comments1.bin")
        );
        let file = open("ooxml/xlsb/hyperlink.xlsb");
        let changed = file
            .set("/sheet[1]/A1", props! { "value" => "Tika" })
            .unwrap();
        assert_eq!(text(&changed, "/sheet[1]/A1"), "Tika");
        // The relationship part is rewritten, the hyperlink target kept.
        assert!(
            member(&changed, "xl/worksheets/_rels/sheet1.bin.rels")
                .contains(r#"Target="http://tika.apache.org/" TargetMode="External""#)
        );
    }
}

/// DOC: a paragraph text set is refused when the paragraph holds a
/// picture, a field, a note reference or another control character (the
/// leaf's dependency refusals); bookmarks are kept and shifted.
#[cfg(feature = "doc")]
mod doc {
    use litchi_ops::{Code, props};

    use super::support::{open, refused, text};

    #[test]
    fn pictures_fields_and_notes_are_refused() {
        for (name, path) in [
            ("ole/doc/testPictures.doc", "/body/p[1]"), // "Sunset: \x01 …"
            ("ole/doc/testPictures.doc", "/body/p[6]"), // SHAPE field around a picture
            ("ole/doc/PngPicture.doc", "/body/p[2]"),   // a floating picture
            ("ole/doc/hyperlink.doc", "/body/p[1]"),    // HYPERLINK field
            ("ole/doc/inline-endnote-and-footnote.doc", "/body/p[4]"), // note references
        ] {
            let file = open(name);
            let error = refused(&file, path, "Replaced");
            assert_eq!(
                error.code(),
                Code::UnsupportedCapability,
                "{name} {path}: {error}"
            );
        }
        // A paragraph of plain text is set.
        let file = open("ole/doc/testPictures.doc");
        let changed = file
            .set("/body/p[5]", props! { "text" => "Shape:" })
            .unwrap();
        assert_eq!(text(&changed, "/body/p[5]"), "Shape:");
    }

    /// Record 0304 F4: `FloatingPictures.doc` tags its style names as smart
    /// tags whose ends lie inside paragraphs 5 onwards. The set is a typed
    /// refusal naming the range, not a corrupt file (record 0308).
    #[test]
    fn a_smart_tag_boundary_inside_the_paragraph_is_refused() {
        let file = open("ole/doc/FloatingPictures.doc");
        let error = refused(&file, "/body/p[5]", "Replaced");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        let chain =
            std::iter::successors(std::error::Error::source(&error), |cause| cause.source())
                .map(ToString::to_string)
                .collect::<Vec<_>>()
                .join(" / ");
        assert!(chain.contains("smart tag"), "{chain}");
        // A tag that ends at its paragraph's edge stays editable.
        let changed = file
            .set("/body/p[1]", props! { "text" => "Plain" })
            .unwrap();
        assert_eq!(text(&changed, "/body/p[1]"), "Plain");
    }
}

/// RTF: a paragraph text set is refused while the body holds fields,
/// pictures, bookmarks, notes or annotations.
#[cfg(feature = "rtf")]
mod rtf {
    use litchi_ops::Code;

    use super::support::{open, refused};

    #[test]
    fn structured_bodies_are_refused() {
        for name in [
            "rtf/hyperlink.rtf",
            "rtf/watermark.rtf",
            "rtf/repeated-deftab.rtf",
        ] {
            let file = open(name);
            let error = refused(&file, "/body/p[1]", "Replaced");
            assert_eq!(error.code(), Code::UnsupportedCapability, "{name}: {error}");
        }
    }
}

/// PPT: a shape or paragraph text set that changes the length of a body
/// with formatting runs, interactions, bookmarks or fields is refused; a
/// same-length set that would write over a metacharacter field (`*` for a
/// date, slide number, header or footer) is refused too.
#[cfg(feature = "ppt")]
mod ppt {
    use litchi_ops::{Code, props};

    use super::support::{open, refused, text};

    #[test]
    fn metacharacter_fields_are_refused() {
        let file = open("poi/test-data/slideshow/datetime.ppt");
        assert_eq!(text(&file, "/slide[1]/shape[1]"), "*");
        let error = refused(&file, "/slide[1]/shape[1]", "X");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        let error = refused(&file, "/slide[1]/shape[1]", "XX");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        // The same text is an exact no-op.
        let same = file
            .set("/slide[1]/shape[1]", props! { "text" => "*" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        // Plain shapes are set.
        let file = open("ole/ppt/text_shapes.ppt");
        let changed = file
            .set("/slide[1]/shape[3]", props! { "text" => "Square" })
            .unwrap();
        assert_eq!(text(&changed, "/slide[1]/shape[3]"), "Square");
    }
}

/// Pages: a body paragraph text set keeps the attachments anchored at the
/// paragraph's start or end and refuses one between its text runs; a
/// section text set that would remove an attachment or a footnote mark is
/// refused.
#[cfg(feature = "pages")]
mod pages {
    use litchi_ops::{Code, props};

    use super::support::{open, refused, text};

    const ANCHOR: char = '\u{fffc}';

    #[test]
    fn anchored_objects_survive_or_refuse() {
        for name in [
            "iwork/pages/image-adjustments-native.pages",
            "iwork/pages/chart-data-native.pages",
        ] {
            let file = open(name);
            let anchors = text(&file, "/section[1]").matches(ANCHOR).count();
            assert_eq!(anchors, 1, "{name}");
            let changed = file
                .set("/body/p[1]", props! { "text" => "Litchi" })
                .unwrap();
            assert_eq!(text(&changed, "/body/p[1]"), "Litchi", "{name}");
            assert_eq!(
                text(&changed, "/section[1]").matches(ANCHOR).count(),
                anchors,
                "{name}: the anchored object is kept"
            );
            let error = refused(&file, "/section[1]", "Litchi");
            assert_ne!(error.code(), Code::InternalError, "{name}: {error}");
        }
    }
}

/// Keynote: a placeholder text set is refused when it would remove an
/// object reference (an inline attachment, a field) of the storage; the
/// leaf's `tests/slide_text.rs` covers those refusals on synthetic
/// storages, since no corpus deck has inline objects in a placeholder.
#[cfg(feature = "keynote")]
mod keynote {
    use litchi_ops::props;

    use super::support::{open, text};

    #[test]
    fn plain_placeholders_are_set() {
        let file = open("iwork/keynote/basic.key");
        let changed = file
            .set("/slide[1]/placeholder[1]", props! { "text" => "Litchi" })
            .unwrap();
        assert_eq!(text(&changed, "/slide[1]/placeholder[1]"), "Litchi");
        assert_eq!(
            text(&changed, "/slide[1]/placeholder[2]"),
            text(&file, "/slide[1]/placeholder[2]")
        );
    }
}

/// ODG: a shape text set replaces only a single plain text span; a text
/// box with several paragraphs, spans, links or fields is refused.
#[cfg(feature = "odg")]
mod odg {
    use litchi_ops::{Code, props};

    use super::support::{open, refused, text};

    #[test]
    fn links_are_refused_and_plain_text_is_set() {
        let file = open("libreoffice-core/desktop/qa/data/3page.odg");
        // A text box the leaf cannot replace losslessly is an unsupported
        // capability, as in the other formats (record 0308).
        let error = refused(&file, "/page[2]/shape[1]", "No link");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        let changed = file
            .set("/page[1]/shape[1]", props! { "text" => "Other" })
            .unwrap();
        assert_eq!(text(&changed, "/page[1]/shape[1]"), "Other");
    }
}
