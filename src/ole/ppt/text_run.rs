/// TextRun parsing for PowerPoint presentations.
///
/// Based on Apache POI's HSLF TextRun and related classes, this module
/// provides proper text extraction with formatting from PPT files.
use super::package::Result;
use super::super::consts::PptRecordType;
use super::record_parser::PptRecord;
use crate::ole::binary::{parse_utf16le_string, parse_windows1252_string_len};

/// Text formatting properties for a text run.
///
/// Based on Apache POI's TextPropCollection and CharacterPropertyBags.
#[derive(Debug, Clone, Default)]
pub struct TextRunFormatting {
    /// Font size in points
    pub font_size: Option<u16>,
    /// Font color (RGB)
    pub font_color: Option<u32>,
    /// Bold formatting
    pub bold: bool,
    /// Italic formatting
    pub italic: bool,
    /// Underline formatting
    pub underline: bool,
    /// Font name
    pub font_name: Option<String>,
}

/// A text run with formatting.
///
/// Based on Apache POI's RichTextRun.
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
        }
    }

    /// Create a text run with formatting.
    pub fn with_formatting(text: String, start_index: usize, formatting: TextRunFormatting) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
        }
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
                let text = parse_utf16le_string(&record.data);
                if !text.is_empty() {
                    let start_index = self.text.len();
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            }
            PptRecordType::TextBytesAtom => {
                // Windows-1252 text
                let text = parse_windows1252_string_len(&record.data, 0, record.data.len());
                if !text.is_empty() {
                    let start_index = self.text.len();
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            }
            PptRecordType::StyleTextPropAtom => {
                // Text formatting properties
                self.apply_style_properties(record)?;
            }
            _ => {
                // Recursively process child records
                for child in &record.children {
                    self.process_record(child)?;
                }
            }
        }

        Ok(())
    }

    /// Apply style properties from StyleTextPropAtom.
    ///
    /// Based on Apache POI's StyleTextPropAtom parsing.
    fn apply_style_properties(&mut self, record: &PptRecord) -> Result<()> {
        if record.data.len() < 10 {
            return Ok(()); // Not enough data
        }

        // Parse the StyleTextPropAtom using POI's logic
        let (_paragraph_styles, character_styles) = super::text_prop::parse_style_text_prop_atom(
            &record.data,
            self.text.chars().count(),
        );

        // Apply character styles to runs
        for (style_idx, char_style) in character_styles.iter().enumerate() {
            // Find the run that corresponds to this style
            if style_idx < self.runs.len() {
                let run = &mut self.runs[style_idx];

                // Extract formatting from character properties
                let mut formatting = TextRunFormatting::default();

                // Font size
                if let Some(size) = char_style.get_value("font.size") {
                    formatting.font_size = Some(size as u16);
                }

                // Font color
                if let Some(color) = char_style.get_value("font.color") {
                    formatting.font_color = Some(color as u32);
                }

                // Character flags (bold, italic, underline)
                if let Some(flags) = char_style.get_value("char.flags") {
                    let (bold, italic, underline) = super::text_prop::extract_char_flags(flags);
                    formatting.bold = bold;
                    formatting.italic = italic;
                    formatting.underline = underline;
                }

                run.formatting = formatting;
            }
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
}

