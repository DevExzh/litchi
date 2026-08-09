//! `StyleTextPropAtom` per-shape text formatting runs (MS-PPT 2.9.44).
//!
//! A `StyleTextPropAtom` record carries the character-level and
//! paragraph-level formatting of the text that precedes it, as an array of
//! `TextPFRun` structures (MS-PPT 2.9.45) followed by an array of
//! `TextCFRun` structures (MS-PPT 2.9.46). It is inert: font references are
//! never resolved and formatting is never applied.

use super::records::record::Record;
use super::text_format_exception::{TextCFException, TextPFException};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// `RT_StyleTextPropAtom` record type (MS-PPT 2.9.44).
const STYLE_TEXT_PROP_TYPE: u16 = 0x0FA1;

/// Largest valid `IndentLevel` value (MS-PPT 2.2.13).
const MAX_INDENT_LEVEL: u16 = 4;

/// `PFMasks` bits that MUST be FALSE inside a `TextPFRun` (MS-PPT 2.9.45):
/// `leftMargin`, `indent`, `defaultTabSize`, and `tabStops`.
const PF_RUN_FORBIDDEN_MASKS: u32 = 0x0000_0100 | 0x0000_0400 | 0x0000_8000 | 0x0010_0000;

/// A `TextPFRun` paragraph-formatting run (MS-PPT 2.9.45).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextPFRun {
    count: u32,
    indent_level: u16,
    paragraph: TextPFException,
}

impl TextPFRun {
    /// Number of characters of the corresponding text this run covers.
    #[must_use]
    pub const fn count(&self) -> u32 {
        self.count
    }
    /// Paragraph indentation level; within 0..=4 (MS-PPT 2.2.13).
    #[must_use]
    pub const fn indent_level(&self) -> u16 {
        self.indent_level
    }
    /// The `TextPFException` paragraph formatting of the run.
    #[must_use]
    pub const fn paragraph(&self) -> &TextPFException {
        &self.paragraph
    }

    /// Parse one `TextPFRun` from the start of `data`.
    ///
    /// Returns the decoded run and the number of bytes consumed.
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 6, "TextPFRun header")?;
        let count = read_u32(data, 0);
        if count == 0 {
            return Err(corrupted("TextPFRun covers zero characters"));
        }
        let indent_level = read_u16(data, 4);
        if indent_level > MAX_INDENT_LEVEL {
            return Err(corrupted("TextPFRun indent level is out of range"));
        }
        let (paragraph, paragraph_len) = TextPFException::parse_prefix(&data[6..])?;
        if paragraph.masks() & PF_RUN_FORBIDDEN_MASKS != 0 {
            return Err(corrupted("TextPFRun has forbidden paragraph masks set"));
        }
        Ok((
            Self {
                count,
                indent_level,
                paragraph,
            },
            6 + paragraph_len,
        ))
    }

    /// Serialize the run without any record header.
    fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(10);
        data.extend_from_slice(&self.count.to_le_bytes());
        data.extend_from_slice(&self.indent_level.to_le_bytes());
        data.extend_from_slice(&self.paragraph.to_payload());
        data
    }
}

/// A `TextCFRun` character-formatting run (MS-PPT 2.9.46).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextCFRun {
    count: u32,
    character: TextCFException,
}

impl TextCFRun {
    /// Number of characters of the corresponding text this run covers.
    #[must_use]
    pub const fn count(&self) -> u32 {
        self.count
    }
    /// The `TextCFException` character formatting of the run.
    #[must_use]
    pub const fn character(&self) -> &TextCFException {
        &self.character
    }

    /// Parse one `TextCFRun` from the start of `data`.
    ///
    /// Returns the decoded run and the number of bytes consumed.
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextCFRun header")?;
        let count = read_u32(data, 0);
        if count == 0 {
            return Err(corrupted("TextCFRun covers zero characters"));
        }
        let (character, character_len) = TextCFException::parse_prefix(&data[4..])?;
        Ok((Self { count, character }, 4 + character_len))
    }

    /// Serialize the run without any record header.
    fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(8);
        data.extend_from_slice(&self.count.to_le_bytes());
        data.extend_from_slice(&self.character.to_payload());
        data
    }
}

/// A parsed `StyleTextPropAtom` record (MS-PPT 2.9.44) with the
/// paragraph-level and character-level formatting runs of the corresponding
/// text.
#[allow(
    clippy::module_name_repetitions,
    reason = "`StyleTextPropAtom` is the established public API name mirroring the MS-PPT record; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleTextPropAtom {
    paragraph_runs: Vec<TextPFRun>,
    character_runs: Vec<TextCFRun>,
}

impl StyleTextPropAtom {
    /// The `TextPFRun` paragraph-formatting runs.
    #[must_use]
    pub fn paragraph_runs(&self) -> &[TextPFRun] {
        &self.paragraph_runs
    }
    /// The `TextCFRun` character-formatting runs.
    #[must_use]
    pub fn character_runs(&self) -> &[TextCFRun] {
        &self.character_runs
    }

    /// Parse a complete `StyleTextPropAtom` record (MS-PPT 2.9.44).
    ///
    /// See [`Self::parse`] for the meaning of `text_length`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record, text_length: usize) -> Result<Self> {
        if record.record_type != RecordType::StyleTextPropAtom
            || record.record_type_raw != STYLE_TEXT_PROP_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("StyleTextPropAtom has an invalid header"));
        }
        Self::parse(&record.data, text_length)
    }

    /// Parse the whole payload of a `StyleTextPropAtom`.
    ///
    /// `text_length` is the number of UTF-16 code units in the corresponding
    /// text, as specified by the most closely preceding `TextCharsAtom` or
    /// `TextBytesAtom`. The runs of each array must cover
    /// `text_length + 1` characters in total: the extra character is the
    /// paragraph mark that terminates the text. Because neither run array is
    /// counted, the boundary between them is only well-defined when the
    /// paragraph runs exactly cover that length.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(data: &[u8], text_length: usize) -> Result<Self> {
        let style_length = u32::try_from(text_length)
            .ok()
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| corrupted("StyleTextPropAtom text length exceeds u32"))?;

        let mut offset = 0usize;
        let mut paragraph_runs = Vec::new();
        let mut paragraph_coverage = 0u32;
        while paragraph_coverage < style_length {
            let (run, consumed) = TextPFRun::parse_prefix(&data[offset..])?;
            if run.count > style_length - paragraph_coverage {
                return Err(corrupted("TextPFRun has invalid character coverage"));
            }
            paragraph_coverage += run.count;
            offset += consumed;
            paragraph_runs.push(run);
        }

        let mut character_runs = Vec::new();
        let mut character_coverage = 0u32;
        while character_coverage < style_length {
            let (run, consumed) = TextCFRun::parse_prefix(&data[offset..])?;
            if run.count > style_length - character_coverage {
                return Err(corrupted("TextCFRun has invalid character coverage"));
            }
            character_coverage += run.count;
            offset += consumed;
            character_runs.push(run);
        }

        if offset != data.len() {
            return Err(corrupted("StyleTextPropAtom has trailing bytes"));
        }
        Ok(Self {
            paragraph_runs,
            character_runs,
        })
    }

    /// Serialize the complete `StyleTextPropAtom`, including its header.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the runs are parsed from a PPT record payload whose length is a u32, so the re-serialized payload length always fits u32"
    )]
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut payload = Vec::new();
        for run in &self.paragraph_runs {
            payload.extend_from_slice(&run.to_payload());
        }
        for run in &self.character_runs {
            payload.extend_from_slice(&run.to_payload());
        }
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&STYLE_TEXT_PROP_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(&payload);
        data
    }
}

fn corrupted(message: impl Into<String>) -> Error {
    Error::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn require_bytes(data: &[u8], offset: usize, needed: usize, field: &str) -> Result<()> {
    if data.len() < offset.saturating_add(needed) {
        return Err(corrupted(format!("{field} is truncated")));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::slide_show_settings::ColorIndexKind;
    use crate::text_run::ParagraphAlignment;

    fn record(data: &[u8]) -> Record {
        Record {
            record_type: RecordType::StyleTextPropAtom,
            record_type_raw: STYLE_TEXT_PROP_TYPE,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    // CFMasks bits used by the sample payloads (MS-PPT 2.9.15).
    const CF_MASK_BOLD: u32 = 0x0000_0001;
    const CF_MASK_TYPEFACE: u32 = 0x0001_0000;
    const CF_MASK_SIZE: u32 = 0x0002_0000;
    const CF_MASK_COLOR: u32 = 0x0004_0000;

    // PFMasks bits used by the sample payloads (MS-PPT 2.9.21).
    const PF_MASK_ALIGN: u32 = 0x0000_0800;
    const PF_MASK_LINE_SPACING: u32 = 0x0000_1000;

    /// Two paragraph runs and two character runs covering 10 characters, so
    /// the corresponding text is 9 UTF-16 code units long.
    fn sample_payload() -> Vec<u8> {
        let mut data = Vec::new();
        // TextPFRun: 6 characters, indent level 2, align + lineSpacing.
        data.extend_from_slice(&6u32.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&(PF_MASK_ALIGN | PF_MASK_LINE_SPACING).to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes()); // align right
        data.extend_from_slice(&150i16.to_le_bytes()); // lineSpacing, percent
        // TextPFRun: 4 characters, indent level 0, no formatting.
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        // TextCFRun: 6 characters, bold + typeface.
        data.extend_from_slice(&6u32.to_le_bytes());
        data.extend_from_slice(&(CF_MASK_BOLD | CF_MASK_TYPEFACE).to_le_bytes());
        data.extend_from_slice(&0x0001u16.to_le_bytes()); // fontStyle: bold
        data.extend_from_slice(&5u16.to_le_bytes()); // fontRef
        // TextCFRun: 4 characters, size + color.
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&(CF_MASK_SIZE | CF_MASK_COLOR).to_le_bytes());
        data.extend_from_slice(&2400i16.to_le_bytes()); // fontSize
        data.extend_from_slice(&[0x12, 0x34, 0x56, 0xFE]); // sRGB color
        data
    }

    #[test]
    fn parses_runs_and_round_trips() {
        let payload = sample_payload();
        let parsed = StyleTextPropAtom::parse_record(&record(&payload), 9).unwrap();

        let paragraphs = parsed.paragraph_runs();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].count(), 6);
        assert_eq!(paragraphs[0].indent_level(), 2);
        assert_eq!(
            paragraphs[0].paragraph().text_alignment(),
            Some(ParagraphAlignment::Right)
        );
        assert_eq!(paragraphs[0].paragraph().line_spacing(), Some(150));
        assert_eq!(paragraphs[1].count(), 4);
        assert_eq!(paragraphs[1].paragraph().masks(), 0);

        let characters = parsed.character_runs();
        assert_eq!(characters.len(), 2);
        assert_eq!(characters[0].count(), 6);
        assert!(characters[0].character().font_style().unwrap().bold());
        assert_eq!(characters[0].character().font_ref(), Some(5));
        assert_eq!(characters[1].count(), 4);
        assert_eq!(characters[1].character().font_size(), Some(2400));
        let color = characters[1].character().color().unwrap();
        assert_eq!(color.kind, ColorIndexKind::Srgb);
        assert_eq!(color.red, 0x12);

        let serialized = parsed.to_bytes();
        assert_eq!(serialized[8..], payload[..]);
        assert_eq!(
            StyleTextPropAtom::parse(&serialized[8..], 9).unwrap(),
            parsed
        );
    }

    #[test]
    fn parses_minimal_runs_and_round_trips() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let parsed = StyleTextPropAtom::parse(&payload, 1).unwrap();
        assert_eq!(parsed.paragraph_runs().len(), 1);
        assert_eq!(parsed.character_runs().len(), 1);
        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn rejects_malformed_style_text_prop() {
        let payload = sample_payload();
        // Wrong record type.
        let mut wrong_type = record(&payload);
        wrong_type.record_type = RecordType::TxCFStyleAtom;
        assert!(StyleTextPropAtom::parse_record(&wrong_type, 9).is_err());
        // Nonzero version.
        let mut wrong_version = record(&payload);
        wrong_version.version = 0xF;
        assert!(StyleTextPropAtom::parse_record(&wrong_version, 9).is_err());
        // Nonzero instance.
        let mut wrong_instance = record(&payload);
        wrong_instance.instance = 1;
        assert!(StyleTextPropAtom::parse_record(&wrong_instance, 9).is_err());
        // Every truncation is rejected.
        for length in 0..payload.len() {
            assert!(StyleTextPropAtom::parse(&payload[..length], 9).is_err());
        }
        // Trailing bytes.
        let mut trailing = payload.clone();
        trailing.push(0);
        assert!(StyleTextPropAtom::parse(&trailing, 9).is_err());
        // Text length that does not fit a u32.
        assert!(StyleTextPropAtom::parse(&payload, usize::MAX).is_err());
        // Wrong text length: run coverage no longer matches.
        assert!(StyleTextPropAtom::parse(&payload, 8).is_err());
        assert!(StyleTextPropAtom::parse(&payload, 10).is_err());
    }

    #[test]
    fn rejects_invalid_run_fields() {
        // Zero-length paragraph run.
        let mut zero_length_run = Vec::new();
        zero_length_run.extend_from_slice(&0u32.to_le_bytes());
        zero_length_run.extend_from_slice(&0u16.to_le_bytes());
        zero_length_run.extend_from_slice(&0u32.to_le_bytes());
        assert!(StyleTextPropAtom::parse(&zero_length_run, 1).is_err());

        // Indent level above the maximum.
        let mut excessive_indent_run = Vec::new();
        excessive_indent_run.extend_from_slice(&2u32.to_le_bytes());
        excessive_indent_run.extend_from_slice(&5u16.to_le_bytes());
        excessive_indent_run.extend_from_slice(&0u32.to_le_bytes());
        excessive_indent_run.extend_from_slice(&2u32.to_le_bytes());
        excessive_indent_run.extend_from_slice(&0u32.to_le_bytes());
        assert!(StyleTextPropAtom::parse(&excessive_indent_run, 1).is_err());

        // Forbidden leftMargin mask inside a TextPFRun.
        let mut left_margin_run = Vec::new();
        left_margin_run.extend_from_slice(&2u32.to_le_bytes());
        left_margin_run.extend_from_slice(&0u16.to_le_bytes());
        left_margin_run.extend_from_slice(&0x0000_0100u32.to_le_bytes());
        left_margin_run.extend_from_slice(&288i16.to_le_bytes());
        left_margin_run.extend_from_slice(&2u32.to_le_bytes());
        left_margin_run.extend_from_slice(&0u32.to_le_bytes());
        assert!(StyleTextPropAtom::parse(&left_margin_run, 1).is_err());

        // Forbidden tabStops mask inside a TextPFRun.
        let mut tab_stops_run = Vec::new();
        tab_stops_run.extend_from_slice(&2u32.to_le_bytes());
        tab_stops_run.extend_from_slice(&0u16.to_le_bytes());
        tab_stops_run.extend_from_slice(&0x0010_0000u32.to_le_bytes());
        tab_stops_run.extend_from_slice(&1u16.to_le_bytes());
        tab_stops_run.extend_from_slice(&100i16.to_le_bytes());
        tab_stops_run.extend_from_slice(&0u16.to_le_bytes());
        tab_stops_run.extend_from_slice(&2u32.to_le_bytes());
        tab_stops_run.extend_from_slice(&0u32.to_le_bytes());
        assert!(StyleTextPropAtom::parse(&tab_stops_run, 1).is_err());

        // Forbidden pp10ext mask inside a TextCFRun.
        let mut pp10ext_run = Vec::new();
        pp10ext_run.extend_from_slice(&2u32.to_le_bytes());
        pp10ext_run.extend_from_slice(&0u16.to_le_bytes());
        pp10ext_run.extend_from_slice(&0u32.to_le_bytes());
        pp10ext_run.extend_from_slice(&2u32.to_le_bytes());
        pp10ext_run.extend_from_slice(&0x0010_0000u32.to_le_bytes());
        assert!(StyleTextPropAtom::parse(&pp10ext_run, 1).is_err());
    }
}
