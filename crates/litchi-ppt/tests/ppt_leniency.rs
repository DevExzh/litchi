#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "fixture assertions panic on construction failure by design"
)]

//! Leniency acceptance tests for findings F5.1 and F5.2.
//!
//! Each fixture is a complete `.ppt` written by the crate's writer with one
//! rich text box, whose `StyleTextPropAtom` payload is then patched in place
//! (same length, so record framing and the persist directory stay valid) to
//! carry exactly one cataloged non-conformity. Every fixture must open
//! through the default path, keep its text, apply the defined formatting,
//! and enumerate the repair; the unpatched control must report nothing.

use litchi_core::OwnedSource;
use litchi_ppt::shapes::ShapeEnum;
use litchi_ppt::text_prop::parse_style_text_prop_atom_strict;
use litchi_ppt::writer::Writer;
use litchi_ppt::writer::text_format::Paragraph;
use litchi_ppt::{
    Defect, DefectLocation, Package, Presentation, Shape as _, SourceBackedPackage, TextPropDefect,
    ToleranceReport, ToleratedDefect,
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
    // Open-time parses repair nothing in this crate yet.
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
