#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "fixture assertions panic on construction failure by design"
)]

//! Leniency acceptance tests for the litchi-ppt findings of
//! `docs/LENIENCY_AUDIT.md`.
//!
//! Each fixture is a complete `.ppt` written by the crate's writer, then
//! patched in place (same length, so record framing and the persist
//! directory stay valid unless a test breaks them on purpose) to carry
//! exactly one cataloged non-conformity. Every fixture must open through the
//! default path, keep what it can, and enumerate the repair; the unpatched
//! control must report nothing.

use litchi_core::OwnedSource;
use litchi_ppt::shapes::ShapeEnum;
use litchi_ppt::text_prop::parse_style_text_prop_atom_strict;
use litchi_ppt::writer::Writer;
use litchi_ppt::writer::text_format::Paragraph;
use litchi_ppt::{
    Defect, DefectLocation, Package, PersistDefect, Presentation, Shape as _, ShapeDefect,
    SlideListDefect, SourceBackedPackage, TextPropDefect, ToleranceReport, ToleratedDefect,
};
use std::io::Cursor;
use std::sync::Arc;

/// `StyleTextPropAtom` record header prefix: version/instance 0, type 4001.
const STYLE_TEXT_PROP_HEADER: [u8; 4] = [0x00, 0x00, 0xA1, 0x0F];
/// The text every fixture carries; the atom covers its length plus one.
const TEXT: &str = "abcd";

/// A written presentation together with the location of its single
/// `StyleTextPropAtom` payload.
struct Fixture {
    bytes: Vec<u8>,
    payload_offset: usize,
    payload_len: usize,
}

impl Fixture {
    fn new(paragraph: Paragraph, expected_payload_prefix: &[u8]) -> Self {
        let mut writer = Writer::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_rich_textbox(slide, 10, 10, 200, 40, vec![paragraph])
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let bytes = output.into_inner();

        let headers: Vec<usize> = bytes
            .windows(STYLE_TEXT_PROP_HEADER.len())
            .enumerate()
            .filter(|(_, window)| *window == STYLE_TEXT_PROP_HEADER)
            .map(|(offset, _)| offset)
            .collect();
        assert_eq!(headers.len(), 1, "fixture must hold one StyleTextPropAtom");
        let header = headers[0];
        let payload_len =
            u32::from_le_bytes(bytes[header + 4..header + 8].try_into().unwrap()) as usize;
        let payload_offset = header + 8;
        // Proves the payload is stored contiguously in the compound file, so
        // in-place patches land in the atom and nowhere else.
        assert!(
            bytes[payload_offset..].starts_with(expected_payload_prefix),
            "unexpected writer payload {:02x?}",
            &bytes[payload_offset..payload_offset + payload_len]
        );
        Self {
            bytes,
            payload_offset,
            payload_len,
        }
    }

    /// Plain paragraph: one `TextPFRun` with an empty mask (10 bytes), then
    /// one `TextCFRun` (count at 10, mask at 14, 6 property bytes).
    fn plain() -> Self {
        Self::new(
            Paragraph::new(TEXT),
            &[5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 0, 0, 0, 0, 0, 7, 0],
        )
    }

    /// Bulleted paragraph: the `TextPFRun` mask selects only the bullet
    /// flags, whose `bulletFlags` word sits at payload offset 10.
    fn bulleted() -> Self {
        Self::new(
            Paragraph::new(TEXT).bullet_enabled(true),
            &[5, 0, 0, 0, 0, 0, 0x01, 0, 0, 0, 0x01, 0x00],
        )
    }

    /// Word-wrapped paragraph: the `TextPFRun` mask selects only the word
    /// wrap flag, whose `wrapFlags` word sits at payload offset 10.
    fn word_wrapped() -> Self {
        Self::new(
            Paragraph::new(TEXT).word_wrap(true),
            &[5, 0, 0, 0, 0, 0, 0, 0, 0x04, 0, 0x02, 0x00],
        )
    }

    fn patch_u16(mut self, payload_index: usize, value: u16) -> Self {
        let at = self.payload_offset + payload_index;
        self.bytes[at..at + 2].copy_from_slice(&value.to_le_bytes());
        self
    }

    fn patch_u32(mut self, payload_index: usize, value: u32) -> Self {
        let at = self.payload_offset + payload_index;
        self.bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
        self
    }

    fn payload(&self) -> &[u8] {
        &self.bytes[self.payload_offset..self.payload_offset + self.payload_len]
    }

    /// Open through the default (`Leniency::Strict`) owned-reader path.
    fn open(&self) -> Presentation {
        Package::from_reader(Cursor::new(self.bytes.clone()))
            .unwrap()
            .presentation()
            .unwrap()
    }

    /// Open through the default source-backed path.
    fn open_source_backed(&self) -> Presentation {
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(self.bytes.clone())))
            .unwrap()
            .presentation()
            .unwrap()
    }
}

/// The repairs made while reading the fixture's only slide, plus that
/// slide's persist ID and its text box's shape ID and paragraph formatting.
struct Read {
    report: ToleranceReport,
    collected: ToleranceReport,
    persist_id: u32,
    shape_id: u32,
    text: String,
    paragraph: litchi_ppt::ParagraphRunFormatting,
}

fn read(presentation: &Presentation) -> Read {
    // The text-box fixtures carry no open-time defect.
    assert!(presentation.tolerance_report().is_clean());

    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 1);
    let slide = &slides[0];
    let shapes = slide.shapes().unwrap();
    let [ShapeEnum::TextBox(textbox)] = shapes else {
        panic!("expected one text box, got {shapes:?}");
    };
    let paragraph = textbox.paragraph_runs()[0].formatting.clone();

    // Threading a report through the read API appends the cached repairs on
    // every call, so the per-slide accessor and the threaded report agree.
    let mut threaded = ToleranceReport::default();
    slide.shapes_with_report(&mut threaded).unwrap();
    let report = slide.tolerance_report().unwrap().clone();
    assert_eq!(threaded, report);

    Read {
        collected: presentation.collect_tolerance_report().unwrap(),
        report,
        persist_id: slide.persist_id(),
        shape_id: textbox.id(),
        text: textbox.text().to_owned(),
        paragraph,
    }
}

fn located(read: &Read, defect: TextPropDefect, context: u32) -> ToleratedDefect {
    ToleratedDefect {
        defect: Defect::TextProp(defect),
        context,
        location: DefectLocation {
            persist_id: Some(read.persist_id),
            shape_id: Some(read.shape_id),
        },
    }
}

/// Opens `fixture` through both default paths and checks it reports exactly
/// `expected` (built from the read's own location) and keeps its text.
fn assert_repaired(fixture: &Fixture, expected: impl Fn(&Read) -> Vec<ToleratedDefect>) -> Read {
    let source_backed = read(&fixture.open_source_backed());
    let owned = read(&fixture.open());
    for read in [&owned, &source_backed] {
        assert_eq!(read.text, TEXT);
        assert_eq!(read.report.defects(), expected(read).as_slice());
        assert_eq!(read.report.unrecorded(), 0);
        assert_eq!(read.collected, read.report);
    }
    owned
}

#[test]
fn conformant_controls_report_nothing() {
    for fixture in [
        Fixture::plain(),
        Fixture::bulleted(),
        Fixture::word_wrapped(),
    ] {
        parse_style_text_prop_atom_strict(fixture.payload(), TEXT.len()).unwrap();
        let read = assert_repaired(&fixture, |_| Vec::new());
        assert!(read.report.is_clean());
        assert!(read.collected.is_clean());
        assert_eq!(read.collected.total(), 0);
    }
}

#[test]
fn overcounted_paragraph_coverage_opens_and_is_recorded() {
    // TextPFRun 0 claims 9 characters of a 5-unit (text + 1) atom.
    let fixture = Fixture::plain().patch_u32(0, 9);
    assert!(parse_style_text_prop_atom_strict(fixture.payload(), TEXT.len()).is_err());

    assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::StyleTextPropCoverage, 0)]
    });
}

#[test]
fn overcounted_character_coverage_opens_and_is_recorded() {
    // TextCFRun 0 (run index 1 after the one paragraph run) claims 7.
    let fixture = Fixture::plain().patch_u32(10, 7);
    assert!(parse_style_text_prop_atom_strict(fixture.payload(), TEXT.len()).is_err());

    assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::StyleTextPropCoverage, 1)]
    });
}

#[test]
fn trailing_style_bytes_open_and_are_recorded() {
    // Clearing the TextCFRun mask leaves its six property bytes as a tail
    // after full coverage; the tail is recorded at the next run index (2).
    let fixture = Fixture::plain().patch_u32(14, 0);
    let error = parse_style_text_prop_atom_strict(fixture.payload(), TEXT.len()).unwrap_err();
    assert!(error.to_string().contains("trailing bytes"), "{error}");

    assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::StyleTextPropCoverage, 2)]
    });
}

#[test]
fn reserved_bullet_flag_bits_are_masked_and_recorded() {
    let fixture = Fixture::bulleted().patch_u16(10, 0xF0F1);

    let read = assert_repaired(&fixture, |read| {
        vec![located(
            read,
            TextPropDefect::ReservedBulletFlagBits,
            0xF0F1,
        )]
    });
    // The defined bit still applies; the raw word is kept for round-trip.
    assert_eq!(read.paragraph.bullet_enabled, Some(true));
    assert_eq!(read.paragraph.bullet_flags_raw, Some(0xF0F1));
}

#[test]
fn reserved_wrap_flag_bits_are_masked_and_recorded() {
    let fixture = Fixture::word_wrapped().patch_u16(10, 0xFFFA);

    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::ReservedWrapFlagBits, 0xFFFA)]
    });
    assert_eq!(read.paragraph.word_wrap, Some(true));
    assert_eq!(read.paragraph.character_wrap, None);
    assert_eq!(read.paragraph.overflow, None);
    assert_eq!(read.paragraph.wrap_flags_raw, Some(0xFFFA));
}

#[test]
fn style_atom_framing_corruption_stays_fatal() {
    // A StyleTextPropAtom whose record header overruns its ClientTextbox is
    // structural corruption, not a cataloged non-conformity.
    let mut fixture = Fixture::plain();
    let length_at = fixture.payload_offset - 4;
    fixture.bytes[length_at..length_at + 4].copy_from_slice(&0x0000_FFFFu32.to_le_bytes());

    let presentation = fixture.open();
    let slides = presentation.slides().unwrap();
    let error = slides[0].shapes().unwrap_err();
    assert!(error.to_string().contains("ClientTextbox"), "{error}");
    assert!(slides[0].tolerance_report().is_err());
    assert!(presentation.collect_tolerance_report().is_err());
}

// ---------------------------------------------------------------------------
// F5.3: text enumeration / ColorIndexStruct / indent-level range checks.
// ---------------------------------------------------------------------------

use litchi_ppt::writer::text_format::{
    TabAlign, TabStop, TextAlign, TextDirection, TextFontAlign, TextRun,
};
use litchi_ppt::{ParagraphAlignment, ParagraphFontAlignment, ParagraphTabAlignment};

/// One paragraph property selected by the `TextPFRun` mask; its first value
/// word sits at payload offset 10, right after the run header.
fn single_paragraph_property(paragraph: Paragraph, mask: [u8; 4], value: &[u8]) -> Fixture {
    let mut prefix = vec![5, 0, 0, 0, 0, 0];
    prefix.extend_from_slice(&mask);
    prefix.extend_from_slice(value);
    Fixture::new(paragraph, &prefix)
}

#[test]
fn f5_3_conformant_controls_report_nothing() {
    let fixtures = [
        single_paragraph_property(
            Paragraph::new(TEXT).align(TextAlign::Center),
            [0, 8, 0, 0],
            &[1, 0],
        ),
        single_paragraph_property(
            Paragraph::new(TEXT).font_alignment(TextFontAlign::Hanging),
            [0, 0, 1, 0],
            &[1, 0],
        ),
        single_paragraph_property(
            Paragraph::new(TEXT).text_direction(TextDirection::RightToLeft),
            [0, 0, 32, 0],
            &[1, 0],
        ),
        single_paragraph_property(
            Paragraph::new(TEXT).tab_stops(vec![TabStop::new(360, TabAlign::Center)]),
            [0, 0, 16, 0],
            &[1, 0, 0x68, 0x01, 1, 0],
        ),
        single_paragraph_property(
            Paragraph::new(TEXT).bullet_color_scheme(3),
            [36, 0, 0, 0],
            &[4, 0, 0, 0, 0, 3],
        ),
        scheme_colored_run(),
    ];
    for fixture in fixtures {
        let read = assert_repaired(&fixture, |_| Vec::new());
        assert!(read.collected.is_clean());
    }
}

#[test]
fn undefined_text_alignment_is_left_unset_and_recorded() {
    let fixture = single_paragraph_property(
        Paragraph::new(TEXT).align(TextAlign::Center),
        [0, 8, 0, 0],
        &[1, 0],
    )
    .patch_u16(10, 9);
    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::UnknownTextAlignment, 9)]
    });
    assert_eq!(read.paragraph.alignment, None);

    let control = single_paragraph_property(
        Paragraph::new(TEXT).align(TextAlign::Center),
        [0, 8, 0, 0],
        &[1, 0],
    );
    assert_eq!(
        assert_repaired(&control, |_| Vec::new())
            .paragraph
            .alignment,
        Some(ParagraphAlignment::Center)
    );
}

#[test]
fn undefined_font_alignment_is_left_unset_and_recorded() {
    let control = single_paragraph_property(
        Paragraph::new(TEXT).font_alignment(TextFontAlign::Hanging),
        [0, 0, 1, 0],
        &[1, 0],
    );
    assert_eq!(
        assert_repaired(&control, |_| Vec::new())
            .paragraph
            .font_alignment,
        Some(ParagraphFontAlignment::Hanging)
    );

    let fixture = control.patch_u16(10, 7);
    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::UnknownFontAlignment, 7)]
    });
    assert_eq!(read.paragraph.font_alignment, None);
}

#[test]
fn undefined_text_direction_is_left_unset_and_recorded() {
    let fixture = single_paragraph_property(
        Paragraph::new(TEXT).text_direction(TextDirection::RightToLeft),
        [0, 0, 32, 0],
        &[1, 0],
    )
    .patch_u16(10, 5);
    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::UnknownTextDirection, 5)]
    });
    assert_eq!(read.paragraph.text_direction, None);
}

#[test]
fn undefined_tab_type_keeps_a_left_tab_and_is_recorded() {
    // Tab count at 10, first tab position at 12, its type at 14.
    let fixture = single_paragraph_property(
        Paragraph::new(TEXT).tab_stops(vec![TabStop::new(360, TabAlign::Center)]),
        [0, 0, 16, 0],
        &[1, 0, 0x68, 0x01, 1, 0],
    )
    .patch_u16(14, 6);
    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::UnknownTabType, 6)]
    });
    let tabs = read.paragraph.tab_stops.expect("tab stops");
    assert_eq!(tabs.len(), 1);
    assert_eq!(tabs[0].position, 360);
    assert_eq!(tabs[0].alignment, ParagraphTabAlignment::Left);
}

#[test]
fn out_of_range_indent_level_is_clamped_and_recorded() {
    let fixture = Fixture::plain().patch_u16(4, 7);
    let read = assert_repaired(&fixture, |read| {
        vec![located(read, TextPropDefect::IndentLevelOutOfRange, 7)]
    });
    assert_eq!(read.paragraph.indent_level, 4);
}

#[test]
fn undefined_bullet_color_index_reads_as_unset_and_is_recorded() {
    // bulletFlags at 10, bullet ColorIndexStruct at 12 (index byte at 15).
    let fixture = single_paragraph_property(
        Paragraph::new(TEXT).bullet_color_scheme(3),
        [36, 0, 0, 0],
        &[4, 0, 0, 0, 0, 3],
    )
    .patch_u32(12, 0x0900_0000);
    let read = assert_repaired(&fixture, |read| {
        vec![located(
            read,
            TextPropDefect::InvalidColorIndex,
            0x0900_0000,
        )]
    });
    assert_eq!(read.paragraph.bullet_color, None);
    assert_eq!(read.paragraph.bullet_scheme_color, None);
    assert_eq!(read.paragraph.bullet_color_raw, Some(0x0900_0000));
}

/// A run colored with scheme slot 3: one empty `TextPFRun` (10 bytes), then
/// a `TextCFRun` whose mask selects font index, size, and color; the color
/// `ColorIndexStruct` sits at payload offset 22.
fn scheme_colored_run() -> Fixture {
    Fixture::new(
        Paragraph::with_runs(vec![TextRun::new(TEXT).color_scheme(3)]),
        &[
            5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 0, 0, 0, 0, 0, 7, 0, 0, 0, 18, 0, 0, 0, 0, 3,
        ],
    )
}

#[test]
fn undefined_text_color_index_reads_as_unset_and_is_recorded() {
    let fixture = scheme_colored_run().patch_u32(22, 0x0A00_0000);
    assert_repaired(&fixture, |read| {
        vec![located(
            read,
            TextPropDefect::InvalidColorIndex,
            0x0A00_0000,
        )]
    });
    let presentation = fixture.open();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();
    let ShapeEnum::TextBox(textbox) = &shapes[0] else {
        panic!("expected a text box");
    };
    let run = &textbox.runs()[0];
    assert_eq!(run.formatting.font_color, None);
    assert_eq!(run.formatting.font_scheme_color, None);
    assert_eq!(run.formatting.font_color_raw, Some(0x0A00_0000));
}

// ---------------------------------------------------------------------------
// F5.11: per-shape error isolation.
// ---------------------------------------------------------------------------

/// `OutlineTextRefAtom` record type: strictly a four-byte atom.
const OUTLINE_TEXT_REF_ATOM: u16 = 0x0F9E;

/// A slide with two plain text boxes, `first` and `second`.
fn two_textboxes() -> Vec<u8> {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    for text in ["first", "second"] {
        writer
            .add_rich_textbox(slide, 10, 10, 200, 40, vec![Paragraph::new(text)])
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn style_atom_headers(bytes: &[u8]) -> Vec<usize> {
    bytes
        .windows(STYLE_TEXT_PROP_HEADER.len())
        .enumerate()
        .filter(|(_, window)| *window == STYLE_TEXT_PROP_HEADER)
        .map(|(offset, _)| offset)
        .collect()
}

fn open_both(bytes: &[u8]) -> [Presentation; 2] {
    [
        Package::from_reader(Cursor::new(bytes.to_vec()))
            .unwrap()
            .presentation()
            .unwrap(),
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
            .unwrap()
            .presentation()
            .unwrap(),
    ]
}

fn textbox_texts(shapes: &[ShapeEnum<'static>]) -> Vec<(u32, String)> {
    shapes
        .iter()
        .map(|shape| {
            let ShapeEnum::TextBox(textbox) = shape else {
                panic!("expected text boxes, got {shape:?}");
            };
            (textbox.id(), textbox.text().to_owned())
        })
        .collect()
}

#[test]
fn one_unreadable_text_box_is_dropped_and_its_sibling_kept() {
    let control = two_textboxes();
    let [control_presentation, _] = open_both(&control);
    let slides = control_presentation.slides().unwrap();
    let control_shapes = textbox_texts(slides[0].shapes().unwrap());
    assert_eq!(
        control_shapes
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["first", "second"]
    );
    assert!(slides[0].tolerance_report().unwrap().is_clean());
    assert!(
        control_presentation
            .collect_tolerance_report()
            .unwrap()
            .is_clean()
    );

    // Retype the first text box's StyleTextPropAtom as an
    // OutlineTextRefAtom: still well framed, but not a four-byte atom, so
    // the first shape's content is unreadable.
    let mut bytes = control;
    let headers = style_atom_headers(&bytes);
    assert_eq!(headers.len(), 2);
    bytes[headers[0] + 2..headers[0] + 4].copy_from_slice(&OUTLINE_TEXT_REF_ATOM.to_le_bytes());

    for presentation in open_both(&bytes) {
        assert!(presentation.tolerance_report().is_clean());
        let slides = presentation.slides().unwrap();
        let slide = &slides[0];
        assert_eq!(
            textbox_texts(slide.shapes().unwrap()),
            [control_shapes[1].clone()]
        );
        let expected = [ToleratedDefect {
            defect: Defect::Shape(ShapeDefect::UnreadableShape),
            context: 0,
            location: DefectLocation {
                persist_id: Some(slide.persist_id()),
                shape_id: Some(control_shapes[0].0),
            },
        }];
        assert_eq!(slide.tolerance_report().unwrap().defects(), expected);
        assert_eq!(
            presentation.collect_tolerance_report().unwrap().defects(),
            expected
        );
    }
}

// ---------------------------------------------------------------------------
// Corpus: conformant files repair nothing.
// ---------------------------------------------------------------------------

fn corpus_ppt_files() -> Vec<std::path::PathBuf> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut pending = vec![root];
    let mut files = Vec::new();
    while let Some(directory) = pending.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                pending.push(path);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("ppt"))
            {
                files.push(path);
            }
        }
    }
    files.sort();
    files
}

/// Every unencrypted `.ppt` in `test-data` opens through the default path
/// with an empty report, open-time and per-slide/notes alike.
#[test]
fn corpus_presentations_report_nothing() {
    let files = corpus_ppt_files();
    assert!(files.len() >= 20, "test-data corpus not found");
    let mut opened = 0;
    for path in &files {
        let presentation = match Package::open(path).and_then(|mut package| package.presentation())
        {
            Ok(presentation) => presentation,
            // Encrypted presentations are out of scope (no password).
            Err(error) => {
                let message = error.to_string().to_ascii_lowercase();
                assert!(
                    message.contains("encrypt") || message.contains("password"),
                    "{} failed to open: {error}",
                    path.display()
                );
                continue;
            },
        };
        opened += 1;
        let report = presentation
            .collect_tolerance_report()
            .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        assert!(
            report.is_clean(),
            "{} reported {:?}",
            path.display(),
            report.defects()
        );
    }
    assert!(opened >= 20, "only {opened} corpus files opened");
}

// ---------------------------------------------------------------------------
// Open-time slide-list and persist-chain repairs (F5.4-F5.7).
// ---------------------------------------------------------------------------

/// `SlidePersistAtom` record header: version/instance 0, type 1011, 20 bytes.
const SLIDE_PERSIST_ATOM_HEADER: [u8; 8] = [0x00, 0x00, 0xF3, 0x03, 0x14, 0, 0, 0];

/// A writer-built deck of `count` titled slides with the persist and slide
/// IDs the control reads back.
struct Deck {
    bytes: Vec<u8>,
    /// `(persist_id, slide_id)` per slide, in presentation order.
    slides: Vec<(u32, u32)>,
}

impl Deck {
    fn new(count: usize) -> Self {
        let mut writer = Writer::new();
        for index in 0..count {
            let slide = writer.add_slide().unwrap();
            writer
                .add_rich_textbox(
                    slide,
                    10,
                    10,
                    200,
                    40,
                    vec![Paragraph::new(format!("slide {index}"))],
                )
                .unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let bytes = output.into_inner();
        let mut slides = Vec::new();
        for presentation in open_both(&bytes) {
            assert!(presentation.tolerance_report().is_clean());
            assert!(presentation.collect_tolerance_report().unwrap().is_clean());
            slides = presentation
                .slides()
                .unwrap()
                .iter()
                .map(|slide| (slide.persist_id(), slide.slide_id()))
                .collect();
        }
        assert_eq!(slides.len(), count);
        Self { bytes, slides }
    }

    /// Payload offset of the presentation `SlidePersistAtom` of `slide`.
    fn slide_persist_atom(&self, slide: usize) -> usize {
        let (persist_id, slide_id) = self.slides[slide];
        let offsets: Vec<usize> = self
            .bytes
            .windows(SLIDE_PERSIST_ATOM_HEADER.len() + 16)
            .enumerate()
            .filter(|(_, window)| {
                window.starts_with(&SLIDE_PERSIST_ATOM_HEADER)
                    && window[8..12] == persist_id.to_le_bytes()
                    && window[20..24] == slide_id.to_le_bytes()
            })
            .map(|(offset, _)| offset + SLIDE_PERSIST_ATOM_HEADER.len())
            .collect();
        assert_eq!(offsets.len(), 1, "SlidePersistAtom of slide {slide}");
        offsets[0]
    }

    fn patch_u32(mut self, at: usize, value: u32) -> Self {
        self.bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
        self
    }

    fn patch_u16(mut self, at: usize, value: u16) -> Self {
        self.bytes[at..at + 2].copy_from_slice(&value.to_le_bytes());
        self
    }

    /// Open through both default paths; each must open, list
    /// `expected_slides` (indexes into the control's slides), and report
    /// exactly `expected` at open time.
    fn assert_opens(&self, expected_slides: &[usize], expected: &[ToleratedDefect]) {
        for presentation in open_both(&self.bytes) {
            let slides: Vec<_> = presentation
                .slides()
                .unwrap()
                .iter()
                .map(|slide| (slide.persist_id(), slide.slide_id()))
                .collect();
            let wanted: Vec<_> = expected_slides
                .iter()
                .map(|&index| self.slides[index])
                .collect();
            assert_eq!(slides, wanted);
            assert_eq!(presentation.tolerance_report().defects(), expected);
            assert_eq!(presentation.tolerance_report().unrecorded(), 0);
            // Reading every remaining slide repairs nothing further.
            assert_eq!(
                presentation.collect_tolerance_report().unwrap().defects(),
                expected
            );
        }
    }
}

fn open_time(defect: impl Into<Defect>, context: u32, persist_id: Option<u32>) -> ToleratedDefect {
    ToleratedDefect {
        defect: defect.into(),
        context,
        location: DefectLocation {
            persist_id,
            shape_id: None,
        },
    }
}

#[test]
fn duplicate_slide_id_keeps_the_first_entry_and_is_recorded() {
    let deck = Deck::new(3);
    let (_, first_slide) = deck.slides[0];
    let (second_persist, _) = deck.slides[1];
    let at = deck.slide_persist_atom(1);
    let deck = deck.patch_u32(at + 12, first_slide);
    deck.assert_opens(
        &[0, 2],
        &[open_time(
            SlideListDefect::DuplicateSlideId,
            first_slide,
            Some(second_persist),
        )],
    );
}

#[test]
fn duplicate_slide_persist_ref_keeps_the_first_entry_and_is_recorded() {
    let deck = Deck::new(3);
    let (first_persist, _) = deck.slides[0];
    let (_, third_slide) = deck.slides[2];
    let at = deck.slide_persist_atom(2);
    let deck = deck.patch_u32(at, first_persist);
    deck.assert_opens(
        &[0, 1],
        &[open_time(
            SlideListDefect::DuplicateSlidePersistRef,
            third_slide,
            Some(first_persist),
        )],
    );
}

#[test]
fn null_slide_id_drops_the_entry_and_is_recorded() {
    let deck = Deck::new(2);
    let (persist_id, _) = deck.slides[0];
    let at = deck.slide_persist_atom(0);
    let deck = deck.patch_u32(at + 12, 0);
    deck.assert_opens(
        &[1],
        &[open_time(
            SlideListDefect::NullSlideReference,
            0,
            Some(persist_id),
        )],
    );
}

#[test]
fn excessive_text_placeholder_count_is_clamped_and_recorded() {
    let deck = Deck::new(2);
    let (persist_id, _) = deck.slides[1];
    let at = deck.slide_persist_atom(1);
    let deck = deck.patch_u32(at + 8, 12);
    deck.assert_opens(
        &[0, 1],
        &[open_time(
            SlideListDefect::TextPlaceholderCount,
            12,
            Some(persist_id),
        )],
    );
}

#[test]
fn nonconforming_slide_persist_atom_header_is_read_by_offset_and_recorded() {
    let deck = Deck::new(2);
    let (persist_id, _) = deck.slides[0];
    let at = deck.slide_persist_atom(0);
    // recInstance 2 on the first slide's atom.
    let deck = deck.patch_u16(at - 8, 0x0020);
    deck.assert_opens(
        &[0, 1],
        &[open_time(
            SlideListDefect::SlidePersistAtomHeader,
            20,
            Some(persist_id),
        )],
    );
}

#[test]
fn unmapped_slide_persist_ref_skips_the_slide_and_is_recorded() {
    let deck = Deck::new(3);
    let (_, slide_id) = deck.slides[1];
    let at = deck.slide_persist_atom(1);
    let deck = deck.patch_u32(at, 999);
    deck.assert_opens(
        &[0, 2],
        &[open_time(
            SlideListDefect::StaleSlidePersistRef,
            slide_id,
            Some(999),
        )],
    );
}

#[test]
fn slide_persist_ref_to_a_non_slide_skips_the_slide_and_is_recorded() {
    let deck = Deck::new(2);
    let document_persist_id = open_both(&deck.bytes)[0]
        .slide_directory()
        .document_persist_id();
    let (_, slide_id) = deck.slides[0];
    let at = deck.slide_persist_atom(0);
    let deck = deck.patch_u32(at, document_persist_id);
    deck.assert_opens(
        &[1],
        &[open_time(
            SlideListDefect::StaleSlidePersistRef,
            slide_id,
            Some(document_persist_id),
        )],
    );
}

/// `CurrentUserAtom.headerToken` of an unencrypted presentation.
const CURRENT_USER_TOKEN: [u8; 4] = 0xE391_C05Fu32.to_le_bytes();

impl Deck {
    /// File offset of the Current User stream's `offsetToCurrentEdit`.
    fn current_edit_field(&self) -> usize {
        let tokens: Vec<usize> = self
            .bytes
            .windows(4)
            .enumerate()
            .filter(|(_, window)| *window == CURRENT_USER_TOKEN)
            .map(|(offset, _)| offset)
            .collect();
        assert_eq!(tokens.len(), 1, "one CurrentUserAtom");
        tokens[0] + 4
    }

    fn current_edit_offset(&self) -> u32 {
        let at = self.current_edit_field();
        u32::from_le_bytes(self.bytes[at..at + 4].try_into().unwrap())
    }

    /// Rename the "Current User" CFB directory entry (same length), so the
    /// package no longer has that stream.
    fn without_current_user(mut self) -> Self {
        let name: Vec<u8> = "Current User"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let entries: Vec<usize> = self
            .bytes
            .windows(name.len())
            .enumerate()
            .filter(|(_, window)| *window == name.as_slice())
            .map(|(offset, _)| offset)
            .collect();
        assert_eq!(entries.len(), 1, "one Current User directory entry");
        // "Current User" -> "Current Usex"
        let last = entries[0] + name.len() - 2;
        self.bytes[last] = b'x';
        self
    }
}

#[test]
fn zero_current_edit_offset_recovers_the_live_user_edit_by_scan() {
    let deck = Deck::new(2);
    let live = deck.current_edit_offset();
    let at = deck.current_edit_field();
    let deck = deck.patch_u32(at, 0);
    deck.assert_opens(
        &[0, 1],
        &[open_time(
            PersistDefect::UnusableCurrentEditOffset,
            live,
            None,
        )],
    );
}

#[test]
fn current_edit_offset_off_a_user_edit_recovers_by_scan() {
    let deck = Deck::new(2);
    let live = deck.current_edit_offset();
    let at = deck.current_edit_field();
    // Offset 8 is the first child record of the DocumentContainer.
    let deck = deck.patch_u32(at, 8);
    deck.assert_opens(
        &[0, 1],
        &[open_time(
            PersistDefect::UnusableCurrentEditOffset,
            live,
            None,
        )],
    );
}

#[test]
fn missing_current_user_stream_recovers_by_scan() {
    let deck = Deck::new(2);
    let live = deck.current_edit_offset();
    let deck = deck.without_current_user();
    deck.assert_opens(
        &[0, 1],
        &[open_time(PersistDefect::MissingCurrentUser, live, None)],
    );
}

#[test]
fn user_edit_chain_corruption_stays_fatal() {
    // With the Current User offset unusable *and* no UserEditAtom left in
    // the stream, there is no chain to recover: the open still fails.
    let deck = Deck::new(2);
    let user_edit_header: [u8; 4] = [0x00, 0x00, 0xF5, 0x0F];
    let headers: Vec<usize> = deck
        .bytes
        .windows(4)
        .enumerate()
        .filter(|(_, window)| *window == user_edit_header)
        .map(|(offset, _)| offset)
        .collect();
    assert_eq!(headers.len(), 1);
    let at = deck.current_edit_field();
    let mut deck = deck.patch_u32(at, 0);
    deck.bytes[headers[0] + 2] = 0xF4;
    for result in [
        Package::from_reader(Cursor::new(deck.bytes.clone()))
            .and_then(|mut package| package.presentation())
            .err(),
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(deck.bytes.clone())))
            .and_then(|package| package.presentation())
            .err(),
    ] {
        let error = result.expect("a stream without a UserEditAtom stays fatal");
        assert!(error.to_string().contains("UserEditAtom"), "{error}");
    }
}

/// File offsets of every record header with type `record_type`.
fn record_headers(bytes: &[u8], record_type: u16) -> Vec<usize> {
    let [low, high] = record_type.to_le_bytes();
    bytes
        .windows(4)
        .enumerate()
        .filter(|(_, window)| window[2] == low && window[3] == high && window[0] & 0x0f == 0)
        .map(|(offset, _)| offset)
        .collect()
}

/// A two-slide deck whose second slide's text was edited through the
/// incremental text-edit path: the stream then carries two `UserEditAtom`s,
/// each with a full persist directory `[1..=4]`, and the newer directory
/// maps the second slide (persist ID 4) to its edited copy.
struct EditedDeck {
    deck: Deck,
    /// File offset of the newer directory's payload.
    directory: usize,
    /// File offset of the newer `UserEditAtom`'s payload.
    user_edit: usize,
}

impl EditedDeck {
    fn new() -> Self {
        use litchi_ppt::text_edit::{Position, Snapshot, Target};

        let deck = Deck::new(2);
        let mut edit = Snapshot::from_bytes(deck.bytes.clone())
            .unwrap()
            .edit_text(Target::new(Position::new(1), Position::new(0)))
            .unwrap();
        edit.set_text("edited").unwrap();
        let bytes = edit.commit().unwrap().snapshot().bytes().to_vec();
        let directories = record_headers(&bytes, 0x1772);
        let user_edits = record_headers(&bytes, 0x0FF5);
        assert_eq!((directories.len(), user_edits.len()), (2, 2));
        let edited = Self {
            deck: Deck {
                bytes,
                slides: deck.slides,
            },
            directory: directories[1] + 8,
            user_edit: user_edits[1] + 8,
        };
        // The newer directory is one full run of four persist IDs.
        assert_eq!(edited.directory_word(0), 0x0040_0001);
        assert_eq!(edited.second_slide_text(), "edited");
        edited
    }

    fn directory_word(&self, index: usize) -> u32 {
        let at = self.directory + 4 * index;
        u32::from_le_bytes(self.deck.bytes[at..at + 4].try_into().unwrap())
    }

    /// Rewrite the newer directory's five words in place.
    fn with_directory(mut self, words: [u32; 5]) -> Self {
        for (index, word) in words.into_iter().enumerate() {
            let at = self.directory + 4 * index;
            self.deck.bytes[at..at + 4].copy_from_slice(&word.to_le_bytes());
        }
        self
    }

    /// The second slide's text, read through both default paths.
    fn second_slide_text(&self) -> String {
        let texts: Vec<String> = open_both(&self.deck.bytes)
            .iter()
            .map(|presentation| {
                let slides = presentation.slides().unwrap();
                let shapes = slides[1].shapes().unwrap();
                let ShapeEnum::TextBox(textbox) = &shapes[0] else {
                    panic!("expected a text box");
                };
                textbox.text().to_owned()
            })
            .collect();
        assert_eq!(texts[0], texts[1]);
        texts[0].clone()
    }
}

#[test]
fn edited_deck_control_reports_nothing() {
    let edited = EditedDeck::new();
    edited.deck.assert_opens(&[0, 1], &[]);
}

#[test]
fn duplicate_persist_id_keeps_the_first_offset_and_is_recorded() {
    let edited = EditedDeck::new();
    let (slide, edited_offset) = (edited.directory_word(3), edited.directory_word(4));
    // [4] -> edited copy, [3, 4] -> slide 1 and a bogus offset for 4.
    let edited = edited.with_directory([0x0010_0004, edited_offset, 0x0020_0003, slide, 1]);
    edited.deck.assert_opens(
        &[0, 1],
        &[ToleratedDefect {
            defect: Defect::Persist(PersistDefect::DuplicatePersistId),
            context: 4,
            location: DefectLocation {
                persist_id: Some(4),
                shape_id: None,
            },
        }],
    );
    assert_eq!(edited.second_slide_text(), "edited");
}

#[test]
fn zero_count_persist_run_is_skipped_and_recorded() {
    let edited = EditedDeck::new();
    let (slide, edited_offset) = (edited.directory_word(3), edited.directory_word(4));
    let edited = edited.with_directory([0x0010_0004, edited_offset, 0x0010_0003, slide, 7]);
    edited.deck.assert_opens(
        &[0, 1],
        &[open_time(PersistDefect::ZeroCountPersistRun, 7, None)],
    );
    assert_eq!(edited.second_slide_text(), "edited");
}

#[test]
fn persist_directory_pointer_to_another_record_is_ignored_and_recorded() {
    let mut edited = EditedDeck::new();
    // Point the newer edit's offsetPersistDirectory (payload word 3) at the
    // older UserEditAtom (its offsetLastEdit, payload word 2).
    let at = edited.user_edit;
    let previous: [u8; 4] = edited.deck.bytes[at + 8..at + 12].try_into().unwrap();
    edited.deck.bytes[at + 12..at + 16].copy_from_slice(&previous);
    edited.deck.assert_opens(
        &[0, 1],
        &[open_time(
            PersistDefect::PersistDirectoryRecordType,
            0x0FF5,
            None,
        )],
    );
    // The older directory now supplies persist ID 4: the pre-edit slide.
    assert_eq!(edited.second_slide_text(), "slide 1");
}

#[test]
fn truncated_persist_directory_stays_fatal() {
    let edited = EditedDeck::new();
    // A run claiming 0xFFF offsets overruns the directory.
    let edited = edited.with_directory([0xFFF0_0001, 0, 0, 0, 0]);
    for result in [
        Package::from_reader(Cursor::new(edited.deck.bytes.clone()))
            .and_then(|mut package| package.presentation())
            .err(),
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(edited.deck.bytes.clone())))
            .and_then(|package| package.presentation())
            .err(),
    ] {
        assert!(
            result.is_some(),
            "a truncated persist directory stays fatal"
        );
    }
}

#[test]
fn out_of_range_current_user_name_length_reads_as_no_username() {
    let deck = Deck::new(2);
    // lenUserName follows headerToken and offsetToCurrentEdit.
    let at = deck.current_edit_field() + 4;
    let deck = deck.patch_u16(at, 300);
    deck.assert_opens(
        &[0, 1],
        &[open_time(PersistDefect::CurrentUserNameLength, 300, None)],
    );
}

#[test]
fn nonconforming_user_edit_length_is_read_and_recorded() {
    // Point Current User at the older UserEditAtom of the edited deck and
    // give it a 30-byte payload: the two extra bytes overlap the record
    // that follows, but every field the reader uses is in the first 28.
    let edited = EditedDeck::new();
    let user_edits = record_headers(&edited.deck.bytes, 0x0FF5);
    let older = user_edits[0];
    // The newer atom's offsetLastEdit is the older atom's stream offset.
    let at = edited.user_edit;
    let older_offset = u32::from_le_bytes(edited.deck.bytes[at + 8..at + 12].try_into().unwrap());
    let field = edited.deck.current_edit_field();
    let mut deck = edited.deck.patch_u32(field, older_offset);
    deck.bytes[older + 4..older + 8].copy_from_slice(&30u32.to_le_bytes());
    deck.assert_opens(
        &[0, 1],
        &[open_time(PersistDefect::UserEditAtomLength, 30, None)],
    );
    // The older edit is the one read: the second slide's pre-edit text.
    let presentation = Package::from_reader(Cursor::new(deck.bytes.clone()))
        .unwrap()
        .presentation()
        .unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[1].shapes().unwrap();
    let ShapeEnum::TextBox(textbox) = &shapes[0] else {
        panic!("expected a text box");
    };
    assert_eq!(textbox.text(), "slide 1");
}
