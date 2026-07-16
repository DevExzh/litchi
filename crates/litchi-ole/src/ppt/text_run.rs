use super::super::consts::PptRecordType;
/// TextRun parsing for PowerPoint presentations.
///
/// Based on Apache POI's HSLF TextRun and related classes, this module
/// provides proper text extraction with formatting from PPT files.
use super::package::{PptError, Result};
use super::records::PptRecord;
use super::text::extractor::{decode_text_bytes, from_utf16le_lossy};

/// Text formatting properties for a text run.
///
/// Based on Apache POI's TextPropCollection and CharacterPropertyBags.
#[derive(Debug, Clone, Default)]
pub struct TextRunFormatting {
    /// Font size in points
    pub font_size: Option<u16>,
    /// Font color (RGB)
    pub font_color: Option<u32>,
    /// Raw PowerPoint `ColorIndexStruct` value
    pub font_color_raw: Option<u32>,
    /// PowerPoint color-scheme index when the color is not direct sRGB
    pub font_scheme_color: Option<u8>,
    /// Bold formatting
    pub bold: bool,
    /// Italic formatting
    pub italic: bool,
    /// Underline formatting
    pub underline: bool,
    /// Shadow formatting
    pub shadow: bool,
    /// Embossed/relief formatting
    pub embossed: bool,
    /// Baseline position as a percentage of line height
    pub baseline_position: Option<i16>,
    /// Font name
    pub font_name: Option<String>,
    /// Zero-based font reference in the PowerPoint font collection
    pub font_index: Option<u16>,
    /// East Asian font reference
    pub asian_font_index: Option<u16>,
    /// ANSI font reference
    pub ansi_font_index: Option<u16>,
    /// Symbol font reference
    pub symbol_font_index: Option<u16>,
}

/// A text run with formatting.
///
/// Based on Apache POI's RichTextRun.
///
/// # Example
///
/// ```rust,ignore
/// let run = TextRun::new("x^2 + y^2 = z^2".to_string(), 0);
/// println!("Text: {}", run.text);
///
/// // Check for embedded MTEF formulas
/// if let Some(formula_ast) = run.mtef_formula_ast() {
///     println!("MTEF formula AST with {} nodes", formula_ast.len());
/// }
/// ```
#[derive(Debug, Clone)]
pub struct TextRun {
    /// Text content
    pub text: String,
    /// Formatting properties
    pub formatting: TextRunFormatting,
    /// Start index in the full text
    pub start_index: usize,
    /// Length in characters
    pub length: usize,
    /// Parsed MTEF formula AST (if this run contains a formula)
    #[cfg(feature = "formula")]
    mtef_formula_ast: Option<Vec<litchi_formula::MathNode<'static>>>,
    /// Parsed MTEF formula AST placeholder (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    mtef_formula_ast: Option<Vec<()>>,
}

impl TextRun {
    /// Create a new text run.
    pub fn new(text: String, start_index: usize) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting: TextRunFormatting::default(),
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Create a text run with formatting.
    pub fn with_formatting(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Create a text run with MTEF formula AST.
    #[cfg(feature = "formula")]
    pub fn with_mtef_formula(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
        mtef_ast: Vec<litchi_formula::MathNode<'static>>,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: Some(mtef_ast),
        }
    }

    /// Create a text run with MTEF formula AST fallback (when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub fn with_mtef_formula(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
        _mtef_ast: Vec<()>,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Check if this text run contains an MTEF formula.
    ///
    /// Returns true if this run contains a parsed MTEF formula AST.
    pub fn has_mtef_formula(&self) -> bool {
        self.mtef_formula_ast.is_some()
    }

    /// Get the MTEF formula AST if this run contains a formula.
    ///
    /// Returns the parsed MTEF formula as AST nodes if this run contains a MathType equation,
    /// None otherwise.
    #[cfg(feature = "formula")]
    pub fn mtef_formula_ast(&self) -> Option<&Vec<litchi_formula::MathNode<'static>>> {
        self.mtef_formula_ast.as_ref()
    }

    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast(&self) -> Option<&Vec<()>> {
        self.mtef_formula_ast.as_ref()
    }

    /// Get a mutable reference to the MTEF formula AST.
    ///
    /// This allows for modification of the formula AST if needed.
    #[cfg(feature = "formula")]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<litchi_formula::MathNode<'static>>> {
        &mut self.mtef_formula_ast
    }

    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<()>> {
        &mut self.mtef_formula_ast
    }
}

/// Text run extractor for PowerPoint slides.
///
/// Based on Apache POI's TextRun, StyleTextPropAtom, and related classes.
pub struct TextRunExtractor {
    /// Full text content
    text: String,
    /// Text runs with formatting
    runs: Vec<TextRun>,
}

impl TextRunExtractor {
    /// Create a new text run extractor.
    pub fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
        }
    }

    /// Extract text runs from PPT records.
    ///
    /// Based on Apache POI's TextExtractor and SlideShow text parsing logic.
    ///
    /// # Arguments
    ///
    /// * `records` - PPT records to extract text from
    pub fn extract_from_records(&mut self, records: &[PptRecord]) -> Result<()> {
        for record in records {
            self.process_record(record)?;
        }
        Ok(())
    }

    /// Process a single PPT record.
    fn process_record(&mut self, record: &PptRecord) -> Result<()> {
        match record.record_type {
            PptRecordType::TextCharsAtom => {
                // UTF-16LE text
                let text = from_utf16le_lossy(&record.data);
                if !text.is_empty() {
                    let start_index = self.text.chars().count();
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            },
            PptRecordType::TextBytesAtom => {
                // Low bytes of UTF-16 characters
                let text = decode_text_bytes(&record.data);
                if !text.is_empty() {
                    let start_index = self.text.chars().count();
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            },
            PptRecordType::StyleTextPropAtom => {
                // Text formatting properties
                self.apply_style_properties(record)?;
            },
            _ => {
                // Recursively process child records
                for child in &record.children {
                    self.process_record(child)?;
                }
            },
        }

        Ok(())
    }

    /// Apply style properties from StyleTextPropAtom.
    ///
    /// Based on Apache POI's StyleTextPropAtom parsing.
    fn apply_style_properties(&mut self, record: &PptRecord) -> Result<()> {
        if record.data.len() < 8 {
            return Ok(()); // Not enough data
        }

        let Some(source_run) = self.runs.pop() else {
            return Ok(());
        };
        let text_length = source_run.text.encode_utf16().count();
        let (_paragraph_styles, character_styles) =
            super::text_prop::parse_style_text_prop_atom(&record.data, text_length);

        if character_styles.is_empty() {
            self.runs.push(source_run);
            return Ok(());
        }

        let mut remaining = source_run.text.as_str();
        let mut character_offset = 0usize;
        for char_style in &character_styles {
            if remaining.is_empty() {
                break;
            }

            let requested_units = char_style.characters_covered as usize;
            let (byte_count, character_count) = utf16_prefix(remaining, requested_units);
            if byte_count == 0 {
                continue;
            }

            let text = remaining[..byte_count].to_string();
            let formatting = formatting_from_style(char_style)?;
            self.runs.push(TextRun::with_formatting(
                text,
                source_run.start_index + character_offset,
                formatting,
            ));
            remaining = &remaining[byte_count..];
            character_offset += character_count;
        }

        if !remaining.is_empty() {
            self.runs.push(TextRun::with_formatting(
                remaining.to_string(),
                source_run.start_index + character_offset,
                source_run.formatting,
            ));
        }

        Ok(())
    }

    /// Get the full extracted text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get all text runs.
    pub fn runs(&self) -> &[TextRun] {
        &self.runs
    }

    /// Get the number of runs.
    pub fn run_count(&self) -> usize {
        self.runs.len()
    }
}

fn formatting_from_style(
    style: &super::text_prop::TextPropCollection,
) -> Result<TextRunFormatting> {
    let font_color_raw = style.get_value("font.color").map(|color| color as u32);
    if font_color_raw.is_some_and(|raw| !matches!((raw >> 24) as u8, 0x00..=0x07 | 0xFE | 0xFF)) {
        return Err(PptError::Corrupted(
            "TextCFRun has an invalid ColorIndexStruct index".to_string(),
        ));
    }
    let (font_color, font_scheme_color) = font_color_raw
        .map(decode_color_index_struct)
        .unwrap_or((None, None));
    let font_size = style
        .get_value("font.size")
        .map(|size| {
            if (1..=4000).contains(&size) {
                Ok(size as u16)
            } else {
                Err(PptError::Corrupted(
                    "TextCFRun font size is outside the 1..=4000 point range".to_string(),
                ))
            }
        })
        .transpose()?;
    let font_index = |name| -> Result<Option<u16>> {
        style
            .get_value(name)
            .map(|index| {
                u16::try_from(index).map_err(|_| {
                    PptError::Corrupted("TextCFRun has an invalid font index".to_string())
                })
            })
            .transpose()
    };
    let baseline_position = style
        .get_value("superscript")
        .map(|position| {
            if (-100..=100).contains(&position) {
                Ok(position as i16)
            } else {
                Err(PptError::Corrupted(
                    "TextCFRun baseline position is outside the -100..=100 range".to_string(),
                ))
            }
        })
        .transpose()?;
    let mut formatting = TextRunFormatting {
        font_size,
        font_color,
        font_color_raw,
        font_scheme_color,
        font_index: font_index("font.index")?,
        asian_font_index: font_index("asian.font.index")?,
        ansi_font_index: font_index("ansi.font.index")?,
        symbol_font_index: font_index("symbol.font.index")?,
        baseline_position,
        ..TextRunFormatting::default()
    };

    if let Some(flags) = style.get_value("char.flags") {
        let (bold, italic, underline) = super::text_prop::extract_char_flags(flags);
        formatting.bold = bold;
        formatting.italic = italic;
        formatting.underline = underline;
        formatting.shadow = flags & 0x0010 != 0;
        formatting.embossed = flags & 0x0200 != 0;
    }
    Ok(formatting)
}

fn decode_color_index_struct(raw: u32) -> (Option<u32>, Option<u8>) {
    let red = raw & 0xFF;
    let green = (raw >> 8) & 0xFF;
    let blue = (raw >> 16) & 0xFF;
    match (raw >> 24) as u8 {
        0xFE => (Some((red << 16) | (green << 8) | blue), None),
        index @ 0x00..=0x07 => (None, Some(index)),
        _ => (None, None),
    }
}

fn utf16_prefix(text: &str, requested_units: usize) -> (usize, usize) {
    if requested_units == 0 {
        return (0, 0);
    }

    let mut units = 0usize;
    let mut byte_count = 0usize;
    let mut character_count = 0usize;
    for (offset, character) in text.char_indices() {
        let next_units = units + character.len_utf16();
        if next_units > requested_units && byte_count != 0 {
            break;
        }
        units = next_units;
        byte_count = offset + character.len_utf8();
        character_count += 1;
        if units >= requested_units {
            break;
        }
    }
    (byte_count, character_count)
}

impl Default for TextRunExtractor {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_run_creation() {
        let run = TextRun::new("Hello".to_string(), 0);
        assert_eq!(run.text, "Hello");
        assert_eq!(run.start_index, 0);
        assert_eq!(run.length, 5);
    }

    #[test]
    fn test_text_run_extractor() {
        let mut extractor = TextRunExtractor::new();

        // Create a simple TextCharsAtom record
        let text_data = vec![
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
            0x00, 0x00, // null terminator
        ];

        let record = PptRecord {
            record_type: PptRecordType::TextCharsAtom,
            record_type_raw: 4000,
            version: 0,
            instance: 0,
            data_length: text_data.len() as u32,
            data: text_data,
            children: Vec::new(),
        };

        extractor.extract_from_records(&[record]).unwrap();
        assert_eq!(extractor.text(), "Hello");
        assert_eq!(extractor.run_count(), 1);
    }

    #[test]
    fn text_run_extractor_uses_ppt_unicode_encodings_and_character_offsets() {
        let unicode_record = PptRecord {
            record_type: PptRecordType::TextCharsAtom,
            record_type_raw: 4000,
            version: 0,
            instance: 0,
            data_length: 4,
            data: vec![0x3D, 0xD8, 0x00, 0xDE],
            children: Vec::new(),
        };
        let byte_record = PptRecord {
            record_type: PptRecordType::TextBytesAtom,
            record_type_raw: 4008,
            version: 0,
            instance: 0,
            data_length: 2,
            data: vec![0x80, 0xE9],
            children: Vec::new(),
        };
        let mut extractor = TextRunExtractor::new();

        extractor
            .extract_from_records(&[unicode_record, byte_record])
            .unwrap();

        assert_eq!(extractor.text(), "😀\u{80}é");
        assert_eq!(extractor.runs()[0].start_index, 0);
        assert_eq!(extractor.runs()[0].length, 1);
        assert_eq!(extractor.runs()[1].start_index, 1);
        assert_eq!(extractor.runs()[1].length, 2);
    }

    #[test]
    fn style_atom_splits_text_into_character_runs() {
        let text_record = PptRecord {
            record_type: PptRecordType::TextBytesAtom,
            record_type_raw: 4008,
            version: 0,
            instance: 0,
            data_length: 4,
            data: b"abcd".to_vec(),
            children: Vec::new(),
        };
        let mut style_data = Vec::new();
        style_data.extend_from_slice(&5u32.to_le_bytes());
        style_data.extend_from_slice(&0i16.to_le_bytes());
        style_data.extend_from_slice(&0u32.to_le_bytes());
        style_data.extend_from_slice(&2u32.to_le_bytes());
        style_data.extend_from_slice(&0x0001u32.to_le_bytes());
        style_data.extend_from_slice(&0x0001i16.to_le_bytes());
        style_data.extend_from_slice(&3u32.to_le_bytes());
        style_data.extend_from_slice(&0x0002u32.to_le_bytes());
        style_data.extend_from_slice(&0x0002i16.to_le_bytes());
        let style_record = PptRecord {
            record_type: PptRecordType::StyleTextPropAtom,
            record_type_raw: 4001,
            version: 0,
            instance: 0,
            data_length: style_data.len() as u32,
            data: style_data,
            children: Vec::new(),
        };
        let mut extractor = TextRunExtractor::new();

        extractor
            .extract_from_records(&[text_record, style_record])
            .unwrap();

        assert_eq!(extractor.text(), "abcd");
        assert_eq!(extractor.run_count(), 2);
        assert_eq!(extractor.runs()[0].text, "ab");
        assert!(extractor.runs()[0].formatting.bold);
        assert!(!extractor.runs()[0].formatting.italic);
        assert_eq!(extractor.runs()[1].text, "cd");
        assert!(!extractor.runs()[1].formatting.bold);
        assert!(extractor.runs()[1].formatting.italic);
        assert_eq!(extractor.runs()[1].start_index, 2);
    }

    #[test]
    fn style_spans_count_utf16_code_units_without_splitting_surrogates() {
        assert_eq!(utf16_prefix("😀x", 2), ("😀".len(), 1));
        assert_eq!(utf16_prefix("😀x", 3), ("😀x".len(), 2));
    }

    #[test]
    fn decodes_direct_and_scheme_color_index_structs() {
        assert_eq!(
            decode_color_index_struct(0xFE33_2211),
            (Some(0x0011_2233), None)
        );
        assert_eq!(decode_color_index_struct(0x0400_0000), (None, Some(4)));
        assert_eq!(decode_color_index_struct(0xFF00_0000), (None, None));
    }

    #[test]
    fn rejects_invalid_text_cf_font_sizes() {
        let text_record = PptRecord {
            record_type: PptRecordType::TextBytesAtom,
            record_type_raw: 4008,
            version: 0,
            instance: 0,
            data_length: 1,
            data: b"x".to_vec(),
            children: Vec::new(),
        };
        let mut style_data = Vec::new();
        style_data.extend_from_slice(&2u32.to_le_bytes());
        style_data.extend_from_slice(&0i16.to_le_bytes());
        style_data.extend_from_slice(&0u32.to_le_bytes());
        style_data.extend_from_slice(&2u32.to_le_bytes());
        style_data.extend_from_slice(&0x0002_0000u32.to_le_bytes());
        style_data.extend_from_slice(&0i16.to_le_bytes());
        let style_record = PptRecord {
            record_type: PptRecordType::StyleTextPropAtom,
            record_type_raw: 4001,
            version: 0,
            instance: 0,
            data_length: style_data.len() as u32,
            data: style_data,
            children: Vec::new(),
        };
        let mut extractor = TextRunExtractor::new();

        let error = extractor
            .extract_from_records(&[text_record, style_record])
            .unwrap_err();
        assert!(error.to_string().contains("font size"));
    }

    #[test]
    fn rejects_invalid_text_cf_color_and_baseline_values() {
        let mut invalid_color = super::super::text_prop::TextPropCollection::new(
            1,
            super::super::text_prop::TextPropType::Character,
        );
        let mut color = super::super::text_prop::TextProp::new("font.color", 4, 0x40000);
        color.value = 0x0800_0000;
        invalid_color.properties.push(color);
        assert!(formatting_from_style(&invalid_color).is_err());

        let mut invalid_position = super::super::text_prop::TextPropCollection::new(
            1,
            super::super::text_prop::TextPropType::Character,
        );
        let mut position = super::super::text_prop::TextProp::new("superscript", 2, 0x80000);
        position.value = 101;
        invalid_position.properties.push(position);
        assert!(formatting_from_style(&invalid_position).is_err());
    }
}
