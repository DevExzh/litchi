#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use super::super::*;
use super::codec::{Charset, RtfWriter, WriterOptions};
use litchi_codepage::Mbcs;
use std::io;

fn write_body_story_for_test(
    shapes: &[Shape<'_>],
    drawing_order: &[StoryDrawing],
    story_events: &[BodyStoryEvent],
) -> io::Result<Vec<u8>> {
    let bookmarks = BookmarkTable::new();
    let mut output = Vec::new();
    {
        let mut writer = RtfWriter::new(&mut output);
        writer.write_blocks_with_markup(
            &[],
            &[],
            &bookmarks,
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            shapes,
            &[],
            drawing_order,
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            &[],
            story_events,
            &[],
        )?;
    }
    Ok(output)
}

#[test]
fn body_story_writer_rejects_out_of_range_resource_references() {
    let error = write_body_story_for_test(
        &[],
        &[],
        &[BodyStoryEvent::Drawing(StoryDrawing::Shape(usize::MAX))],
    )
    .unwrap_err();

    assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
}

#[test]
fn body_story_writer_rejects_duplicate_resource_references() {
    let shapes = [Shape::new(ShapeType::Rectangle)];
    let drawing_order = [StoryDrawing::Shape(0)];
    let story_events = [
        BodyStoryEvent::Drawing(StoryDrawing::Shape(0)),
        BodyStoryEvent::Drawing(StoryDrawing::Shape(0)),
    ];

    let error = write_body_story_for_test(&shapes, &drawing_order, &story_events).unwrap_err();

    assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
}

#[test]
fn header_footer_writer_rejects_non_utf8_story_boundaries() {
    let mut header = HeaderFooter::new(HeaderFooterType::Header);
    header.add_paragraph(HeaderFooterParagraph::new(
        std::borrow::Cow::Borrowed("é"),
        Formatting::default(),
        Paragraph::default(),
    ));
    header
        .story_events
        .push(StoryEvent::PageBreak(PageBreak::new(1)));
    let mut output = Vec::new();

    let error = RtfWriter::new(&mut output)
        .write_header_footer(&header)
        .unwrap_err();

    assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
}

#[test]
fn test_simple_document() {
    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);

    writer.write_document_header().unwrap();
    writer.write_text("Hello World").unwrap();
    writer.write_str("}").unwrap();

    let result = String::from_utf8(output).unwrap();
    assert!(result.contains("rtf1"));
    assert!(result.contains("Hello World"));
}

#[test]
fn plain_ascii_text_is_written_as_one_chunk() {
    #[derive(Default)]
    struct CountingSink {
        bytes: Vec<u8>,
        writes: usize,
    }

    impl io::Write for CountingSink {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            self.writes += 1;
            self.bytes.extend_from_slice(buffer);
            Ok(buffer.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    let mut sink = CountingSink::default();
    RtfWriter::new(&mut sink)
        .write_text("A long ordinary ASCII fragment 0123456789")
        .unwrap();

    assert_eq!(sink.writes, 1);
    assert_eq!(sink.bytes, b"A long ordinary ASCII fragment 0123456789");
}

#[test]
fn chunked_text_writer_preserves_all_escape_spellings() {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_text("plain\\{}\n\t\r—é")
        .unwrap();

    assert_eq!(
        String::from_utf8(output).unwrap(),
        "plain\\\\\\{\\}\\par \\tab \\'0d\\emdash \\u233?"
    );
}

#[test]
fn checked_charsets_write_their_exact_header_controls() {
    assert!(Charset::ansi(1200).is_err());
    assert!(Charset::ansi(65000).is_err());
    for (charset, expected) in [
        (Charset::Ansi(Mbcs::WINDOWS_1252), "\\ansi\\ansicpg1252"),
        (Charset::Mac, "\\mac"),
        (Charset::Pc, "\\pc"),
        (Charset::Pca, "\\pca"),
    ] {
        let mut output = Vec::new();
        let options = WriterOptions {
            charset,
            ..WriterOptions::default()
        };
        let mut writer = RtfWriter::with_options(&mut output, options);
        writer.write_document_header().unwrap();
        writer.write_str("}").unwrap();

        let result = String::from_utf8(output).unwrap();
        assert!(result.contains(expected), "{charset:?}: {result}");
        assert!(RtfDocument::parse(&result).is_ok(), "{charset:?}: {result}");
    }
}

#[test]
fn test_control_words() {
    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);

    writer.write_control_word("test", Some(42)).unwrap();
    writer.write_control_word("flag", None).unwrap();

    let result = String::from_utf8(output).unwrap();
    assert_eq!(result, "\\test42\\flag");
}

#[test]
fn equation_writer_uses_an_empty_cached_result_without_evaluation() {
    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);
    writer.write_document_header().unwrap();
    writer.write_equation(r"\f(1,2)").unwrap();
    writer.write_str("}").unwrap();

    let serialized = String::from_utf8(output).unwrap();
    assert!(serialized.contains(r"EQ \\f(1,2)"));
    assert!(serialized.contains(r"{\fldrslt}"));

    let document = RtfDocument::parse(&serialized).unwrap();
    let equations = document.equations();
    assert_eq!(equations.len(), 1);
    assert_eq!(equations[0].expression(), r"\f(1,2)");
    assert_eq!(equations[0].cached_result(), None);
}

#[test]
fn document_writer_round_trips_caller_authored_eq_fields() {
    let mut document = RtfDocument::parse(r"{\rtf1\ansi BeforeAfter}").unwrap();
    let mut equation = Field::new_equation(r"\f(1,2)").unwrap();
    equation.owner = FieldOwner::Body;
    equation.position = "Before".len();
    equation.range_end = equation.position;
    document.push_field(equation).unwrap();

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(output).unwrap();
    assert!(serialized.contains(r"{\fldrslt}"));

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.text(), "BeforeAfter");
    assert_eq!(reparsed.equation_count(), 1);
    assert_eq!(reparsed.equations()[0].expression(), r"\f(1,2)");
}

#[test]
fn document_info_writer_round_trips() {
    let mut info = DocumentInfo::new().with_title(std::borrow::Cow::Borrowed("Résumé 你"));
    info.author = Some(std::borrow::Cow::Borrowed("Ada"));
    info.creation_time = Some(std::borrow::Cow::Borrowed("2026-07-15T12:34:56"));
    info.pages = Some(3);
    info.characters_with_spaces = Some(42);

    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);
    writer.write_document_header().unwrap();
    writer.write_document_info(&info).unwrap();
    writer.write_text("Body").unwrap();
    writer.write_str("}").unwrap();

    let rtf = String::from_utf8(output).unwrap();
    let parsed = RtfDocument::parse(&rtf).unwrap();
    assert_eq!(parsed.info().title.as_deref(), Some("Résumé 你"));
    assert_eq!(parsed.info().author.as_deref(), Some("Ada"));
    assert_eq!(
        parsed.info().creation_time.as_deref(),
        Some("2026-07-15T12:34:56")
    );
    assert_eq!(parsed.info().pages, Some(3));
    assert_eq!(parsed.info().characters_with_spaces, Some(42));
    assert_eq!(parsed.text(), "Body");
}

#[test]
fn document_writer_round_trips_bookmark_ranges() {
    let source = r"{\rtf1\ansi Start {\*\bkmkstart\bkmkcolf2\bkmkcoll4\bkmkpub Link}R\'e9sum\'e9 \u20320?{\*\bkmkend Link} end}";
    let document = RtfDocument::parse(source).unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    let bookmark = reparsed.bookmarks().get("Link").unwrap();
    assert_eq!(bookmark.content, "Résumé 你");
    assert_eq!(bookmark.first_column, Some(2));
    assert_eq!(bookmark.last_column, Some(4));
    assert!(bookmark.is_public);
}

#[test]
fn document_writer_preserves_bookmark_in_empty_body() {
    let document = RtfDocument::parse(r"{\rtf1{\*\bkmkstart Empty}{\*\bkmkend Empty}}").unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    let bookmark = reparsed.bookmarks().get("Empty").unwrap();
    assert_eq!(bookmark.position, 0);
    assert!(bookmark.content.is_empty());
    assert!(reparsed.text().is_empty());
}

#[test]
fn document_writer_round_trips_annotations() {
    let source = r"{\rtf1\ansi Before {\*\atrfstart 12}range{\*\atrfend 12}{\*\atnid AM}{\*\atnauthor Ada M}\chatn{\*\annotation{\*\atnref 12}{\*\atndate 12345}{\*\atnparent 4}{\*\atnicn 3}{\*\atntime 99}Review \u20320? now} after}";
    let document = RtfDocument::parse(source).unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.annotations().len(), 1);
    let annotation = &reparsed.annotations()[0];
    assert_eq!(annotation.id, 12);
    assert_eq!(annotation.author, "Ada M");
    assert_eq!(annotation.initials, "AM");
    assert_eq!(annotation.date.as_deref(), Some("12345"));
    assert_eq!(annotation.parent_id.as_deref(), Some("4"));
    assert_eq!(annotation.icon.as_deref(), Some("3"));
    assert_eq!(annotation.time.as_deref(), Some("99"));
    assert_eq!(annotation.text, "Review 你 now");
    assert_eq!(annotation.position, "Before ".len());
    assert_eq!(annotation.range_end, "Before range".len());
}

#[test]
fn document_writer_round_trips_headers_and_footers() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi\sectd\sbkodd\pgwsxn11000\pghsxn15000\marglsxn910\margrsxn810\margtsxn710\margbsxn610\guttersxn130\headery310\footery410\lndscpsxn\cols3\colsx370\pgnstarts6\pgnlcltr\vertalb\linemod1\lineppage{\header Header \u20320? one\par Header two}{\footer Footer}Body}",
    )
    .unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), "Body");
    assert_eq!(reparsed.sections().len(), 1);
    let section = &reparsed.sections()[0];
    assert_eq!(section.properties, document.sections()[0].properties);
    assert_eq!(
        section.get_header(HeaderFooterType::Header).unwrap().text(),
        "Header 你 one\nHeader two"
    );
    assert_eq!(
        section.get_header(HeaderFooterType::Footer).unwrap().text(),
        "Footer"
    );
}

#[test]
fn document_writer_round_trips_stylesheets() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\stylesheet{\s0\snext0 Normal;}{\s1\b\qc\sbasedon0\snext0\slink2\sautoupd\shidden\slocked\ssemihidden\sunhideused\sqformat\spriority9\styrsid42 Heading \u20320?;}{\*\cs2\i\additive\slink1 Emphasis;}{\*\ds3 Section;}{\*\ts4 Table;}}Body}",
    )
    .unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), "Body");
    assert_eq!(reparsed.stylesheet().styles().len(), 5);
    let heading = reparsed
        .stylesheet()
        .get_typed(StyleType::Paragraph, 1)
        .unwrap();
    assert_eq!(heading.name, "Heading 你");
    assert!(heading.formatting.bold);
    assert_eq!(heading.paragraph.unwrap().alignment, Alignment::Center);
    assert_eq!(heading.linked_style, Some(2));
    assert!(heading.auto_update);
    assert!(heading.hidden);
    assert!(heading.locked);
    assert!(heading.semi_hidden);
    assert!(heading.unhide_when_used);
    assert!(heading.quick_format);
    assert_eq!(heading.priority, Some(9));
    assert_eq!(heading.revision_id, Some(42));

    let character = reparsed
        .stylesheet()
        .get_typed(StyleType::Character, 2)
        .unwrap();
    assert!(character.additive);
    assert!(character.formatting.italic);
    assert!(
        reparsed
            .stylesheet()
            .get_typed(StyleType::Section, 3)
            .is_some()
    );
    assert!(
        reparsed
            .stylesheet()
            .get_typed(StyleType::Table, 4)
            .is_some()
    );
}

#[test]
fn document_writer_round_trips_list_tables() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\*\listtable{\list\listtemplateid42\listhybrid{\listlevel\levelnfc0\leveljc2\levelfollow1\levelstartat3\levelspace120\levelindent360{\leveltext\'02\'00.;}{\levelnumbers\'01;}\f2}{\listlevel\levelnfc77\leveljc0\levelfollow2\levelstartat1{\leveltext\'01\u8226?;}{\levelnumbers;}}{\listname Outline;}\listid77}}{\*\listoverridetable{\listoverride\listid77\listoverridecount1{\lfolevel\listoverridestartat\levelstartat9}\ls4}}\pard\ls4\ilvl1 Body}",
    )
    .unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), "Body");
    let paragraph = reparsed.blocks().last().unwrap().paragraph;
    assert_eq!(paragraph.list_override, Some(4));
    assert_eq!(paragraph.list_level, Some(1));
    let list = reparsed.list_table().get(77).unwrap();
    assert_eq!(list.template_id, 42);
    assert!(list.hybrid);
    assert_eq!(list.name, "Outline");
    assert_eq!(list.levels.len(), 2);
    assert_eq!(list.levels[0].number_text, "\0.");
    assert_eq!(list.levels[0].follow, ListFollow::Space);
    assert_eq!(list.levels[1].level_type, ListLevelType::Other(77));
    assert_eq!(list.levels[1].number_text, "•");
    assert_eq!(list.levels[1].follow, ListFollow::Nothing);
    let list_override = reparsed.list_override_table().get(4).unwrap();
    assert_eq!(list_override.list_id, 77);
    assert_eq!(list_override.level_count_override, Some(1));
    assert_eq!(list_override.start_at_override, Some(9));
}

#[test]
fn document_writer_round_trips_tracked_revision_ranges() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\*\revtbl{Unknown;}{Ada;}}Before {\deleted\revauthdel1\revdttmdel123 old}{\revised\revauth1\revdttm-456 new \u20320?} after}",
    )
    .unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let reparsed = RtfDocument::from_bytes(&output).unwrap_or_else(|error| {
        panic!(
            "failed to parse revision writer output: {error}\n{}",
            String::from_utf8_lossy(&output)
        )
    });
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.revisions().len(), 2);
    for (actual, expected) in reparsed.revisions().iter().zip(document.revisions()) {
        assert_eq!(actual.revision_type, expected.revision_type);
        assert_eq!(actual.id, expected.id);
        assert_eq!(actual.author, expected.author);
        assert_eq!(actual.date, expected.date);
        assert_eq!(actual.content, expected.content);
        assert_eq!(actual.position, expected.position);
        assert_eq!(actual.range_end, expected.range_end);
    }
}

#[test]
fn document_writer_round_trips_multiple_section_boundaries() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi\sectd\sbkpage\pgwsxn10000{\header First}One\sect\sectd\sbknone\pgwsxn12000{\header Second}Two\sect\sectd\sbkeven\pgwsxn14000{\header Third}Three}",
    )
    .unwrap();
    assert_eq!(document.text(), "OneTwoThree");
    assert_eq!(document.sections().len(), 3);
    assert_eq!(
        document.section_breaks().copied().collect::<Vec<_>>(),
        vec![
            SectionBreak::new("One".len(), Some(1)),
            SectionBreak::new("OneTwo".len(), Some(2)),
        ]
    );

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8_lossy(&output);
    assert!(serialized.contains("\\sect\\sectd"));

    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.sections().len(), 3);
    assert_eq!(
        reparsed.section_breaks().copied().collect::<Vec<_>>(),
        document.section_breaks().copied().collect::<Vec<_>>(),
    );
    assert_eq!(
        reparsed.sections()[0].properties.break_type,
        SectionBreakType::Page
    );
    assert_eq!(reparsed.sections()[0].properties.page_width, 10000);
    assert_eq!(
        reparsed.sections()[1].properties.break_type,
        SectionBreakType::Continuous
    );
    assert_eq!(reparsed.sections()[1].properties.page_width, 12000);
    assert_eq!(
        reparsed.sections()[2].properties.break_type,
        SectionBreakType::EvenPage
    );
    assert_eq!(reparsed.sections()[2].properties.page_width, 14000);
    for (section, expected_header) in reparsed.sections().iter().zip(["First", "Second", "Third"]) {
        assert_eq!(
            section.get_header(HeaderFooterType::Header).unwrap().text(),
            expected_header
        );
    }
}

#[test]
fn document_writer_round_trips_boundary_to_first_explicit_section() {
    let document =
        RtfDocument::parse(r"{\rtf1\ansi Before\sect\sectd\sbknone{\header Second}After}").unwrap();
    assert_eq!(document.text(), "BeforeAfter");
    assert_eq!(document.sections().len(), 1);
    assert_eq!(
        document.section_breaks().copied().collect::<Vec<_>>(),
        vec![SectionBreak::new("Before".len(), Some(0))]
    );

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.section_breaks().copied().collect::<Vec<_>>(),
        document.section_breaks().copied().collect::<Vec<_>>(),
    );
    assert_eq!(reparsed.sections().len(), 1);
    assert_eq!(
        reparsed.sections()[0]
            .get_header(HeaderFooterType::Header)
            .unwrap()
            .text(),
        "Second"
    );
}

#[test]
fn document_writer_preserves_inherited_section_boundary() {
    let document = RtfDocument::parse(r"{\rtf1\ansi\sectd\sbknone One\sect Two}").unwrap();
    assert_eq!(document.text(), "OneTwo");
    assert_eq!(document.sections().len(), 1);
    assert_eq!(
        document.section_breaks().copied().collect::<Vec<_>>(),
        vec![SectionBreak::new("One".len(), None)]
    );

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.sections().len(), 1);
    assert_eq!(
        reparsed.section_breaks().copied().collect::<Vec<_>>(),
        document.section_breaks().copied().collect::<Vec<_>>(),
    );
}

#[test]
fn document_writer_preserves_page_then_section_break_order() {
    let document =
        RtfDocument::parse(r"{\rtf1\ansi\sectd One\page\sect\sectd\sbknone Two}").unwrap();
    assert_eq!(document.text(), "OneTwo");
    assert_eq!(document.sections().len(), 2);
    assert_eq!(
        document.body_story_events(),
        [
            BodyStoryEvent::PageBreak(PageBreak::new("One".len())),
            BodyStoryEvent::SectionBreak(SectionBreak::new("One".len(), Some(1),)),
        ]
    );

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.body_story_events(), document.body_story_events());
}

#[test]
fn document_writer_round_trips_libreoffice_multi_section_fixture() {
    let document = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf94043.rtf"
    )))
    .unwrap();
    assert!(document.sections().len() >= 3);
    assert!(document.section_breaks().count() >= 3);

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.sections().len(), document.sections().len());
    assert_eq!(
        reparsed.section_breaks().copied().collect::<Vec<_>>(),
        document.section_breaks().copied().collect::<Vec<_>>(),
    );
}
