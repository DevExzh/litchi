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
    /// Original `CFMasks` value.
    pub property_mask: u32,
    /// Raw `CFStyle` value, when present.
    pub font_style_raw: Option<u16>,
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
    /// Explicit bold value, or `None` when inherited.
    pub bold_explicit: Option<bool>,
    /// Italic formatting
    pub italic: bool,
    /// Explicit italic value, or `None` when inherited.
    pub italic_explicit: Option<bool>,
    /// Underline formatting
    pub underline: bool,
    /// Explicit underline value, or `None` when inherited.
    pub underline_explicit: Option<bool>,
    /// Shadow formatting
    pub shadow: bool,
    /// Explicit shadow value, or `None` when inherited.
    pub shadow_explicit: Option<bool>,
    /// Whether the run originated from double-byte input, when specified.
    pub fe_hint: Option<bool>,
    /// Whether Kumimoji formatting is active, when specified.
    pub kumi: Option<bool>,
    /// De facto legacy strikethrough value from the MS-PPT `unused3` bit.
    pub legacy_strikethrough: Option<bool>,
    /// Embossed/relief formatting
    pub embossed: bool,
    /// Explicit emboss value, or `None` when inherited.
    pub embossed_explicit: Option<bool>,
    /// PowerPoint 9 additional-property run grouping identifier.
    pub pp9_run_id: Option<u8>,
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

/// Paragraph alignment stored by `TextPFException`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphAlignment {
    /// Left for horizontal text, top for vertical text.
    Left,
    /// Center for horizontal text, middle for vertical text.
    Center,
    /// Right for horizontal text, bottom for vertical text.
    Right,
    /// Flush both horizontal or vertical edges.
    Justify,
    /// Distribute space between characters.
    Distributed,
    /// Thai distributed justification.
    ThaiDistributed,
    /// Low Kashida justification.
    JustifyLow,
}

/// Vertical placement of characters within the line height.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphFontAlignment {
    /// Place characters on the font baseline.
    Roman,
    /// Hang characters from the top of the line.
    Hanging,
    /// Center characters within the line height.
    Center,
    /// Anchor characters to the bottom of the line.
    UpholdFixed,
}

/// Paragraph text direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTextDirection {
    /// Left-to-right text flow.
    LeftToRight,
    /// Right-to-left text flow.
    RightToLeft,
}

/// Alignment at a paragraph tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTabAlignment {
    /// Left-aligned tab stop.
    Left,
    /// Center-aligned tab stop.
    Center,
    /// Right-aligned tab stop.
    Right,
    /// Decimal-point-aligned tab stop.
    Decimal,
}

/// A paragraph tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphTabStop {
    /// Signed offset in PowerPoint master units.
    pub position: i16,
    /// How text aligns at the stop.
    pub alignment: ParagraphTabAlignment,
}

/// Formatting explicitly carried by one PowerPoint paragraph run.
#[derive(Debug, Clone, Default)]
pub struct ParagraphRunFormatting {
    /// Original `PFMasks` value.
    pub property_mask: u32,
    /// Paragraph indentation level.
    pub indent_level: u16,
    /// Raw `BulletFlags` value, when present.
    pub bullet_flags_raw: Option<u16>,
    /// Whether the paragraph has a bullet, when explicitly specified.
    pub bullet_enabled: Option<bool>,
    /// Whether the bullet font override is active, when explicitly specified.
    pub bullet_font_enabled: Option<bool>,
    /// Whether the bullet color override is active, when explicitly specified.
    pub bullet_color_enabled: Option<bool>,
    /// Whether the bullet size override is active, when explicitly specified.
    pub bullet_size_enabled: Option<bool>,
    /// Raw UTF-16 code unit used as the bullet character.
    pub bullet_character: Option<u16>,
    /// Zero-based bullet font reference.
    pub bullet_font_index: Option<u16>,
    /// Raw `BulletSize` value.
    pub bullet_size: Option<i16>,
    /// Normalized direct bullet color in `0xRRGGBB` form.
    pub bullet_color: Option<u32>,
    /// Raw bullet `ColorIndexStruct` value.
    pub bullet_color_raw: Option<u32>,
    /// Bullet color-scheme index when the color is not direct sRGB.
    pub bullet_scheme_color: Option<u8>,
    /// Paragraph alignment.
    pub alignment: Option<ParagraphAlignment>,
    /// Raw line-spacing value.
    pub line_spacing: Option<i16>,
    /// Raw space-before value.
    pub space_before: Option<i16>,
    /// Raw space-after value.
    pub space_after: Option<i16>,
    /// Left margin in master units.
    pub left_margin: Option<i16>,
    /// First-line indent in master units.
    pub indent: Option<i16>,
    /// Default tab size in master units.
    pub default_tab_size: Option<i16>,
    /// Explicit paragraph tab stops.
    pub tab_stops: Option<Vec<ParagraphTabStop>>,
    /// Character alignment within the line height.
    pub font_alignment: Option<ParagraphFontAlignment>,
    /// Whether East Asian character wrapping is active, when explicitly specified.
    pub character_wrap: Option<bool>,
    /// Whether wrapping occurs at word boundaries, when explicitly specified.
    pub word_wrap: Option<bool>,
    /// Whether hanging punctuation is allowed, when explicitly specified.
    pub overflow: Option<bool>,
    /// Raw `PFWrapFlags` value, when present.
    pub wrap_flags_raw: Option<u16>,
    /// Paragraph text direction.
    pub text_direction: Option<ParagraphTextDirection>,
}

impl ParagraphRunFormatting {
    /// Decode the bullet UTF-16 code unit as a Unicode scalar value.
    ///
    /// Returns `None` when no bullet character is present or the stored unit is
    /// an unpaired surrogate.
    pub fn bullet_char(&self) -> Option<char> {
        char::from_u32(u32::from(self.bullet_character?))
    }
}

/// A text range carrying paragraph-level formatting.
#[derive(Debug, Clone)]
pub struct ParagraphRun {
    /// Text covered by this paragraph style, including any stored paragraph marker.
    pub text: String,
    /// Paragraph formatting properties.
    pub formatting: ParagraphRunFormatting,
    /// Start index in Unicode scalar values in the full text.
    pub start_index: usize,
    /// Length in Unicode scalar values.
    pub length: usize,
}

impl ParagraphRun {
    /// Create a paragraph run with explicit formatting.
    pub fn with_formatting(
        text: String,
        start_index: usize,
        formatting: ParagraphRunFormatting,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
        }
    }
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
    /// Paragraph-level formatting runs
    paragraph_runs: Vec<ParagraphRun>,
    /// Most recently encountered text atom awaiting its style atom
    pending_text: Option<(String, usize)>,
}

impl TextRunExtractor {
    /// Create a new text run extractor.
    pub fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
            paragraph_runs: Vec::new(),
            pending_text: None,
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
                let start_index = self.text.chars().count();
                self.pending_text = Some((text.clone(), start_index));
                if !text.is_empty() {
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            },
            PptRecordType::TextBytesAtom => {
                // Low bytes of UTF-16 characters
                let text = decode_text_bytes(&record.data);
                let start_index = self.text.chars().count();
                self.pending_text = Some((text.clone(), start_index));
                if !text.is_empty() {
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

        let Some((source_text, start_index)) = self.pending_text.take() else {
            return Ok(());
        };
        let source_run = if source_text.is_empty() {
            TextRun::new(String::new(), start_index)
        } else {
            self.runs
                .pop()
                .unwrap_or_else(|| TextRun::new(source_text.clone(), start_index))
        };
        let text_length = source_text.encode_utf16().count();
        let (paragraph_styles, character_styles) =
            super::text_prop::parse_style_text_prop_atom_strict(&record.data, text_length)?;

        self.apply_paragraph_styles(&source_text, start_index, &paragraph_styles)?;

        if character_styles.is_empty() {
            if !source_run.text.is_empty() {
                self.runs.push(source_run);
            }
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

    fn apply_paragraph_styles(
        &mut self,
        source_text: &str,
        start_index: usize,
        paragraph_styles: &[super::text_prop::TextPropCollection],
    ) -> Result<()> {
        if paragraph_styles.is_empty() {
            return Ok(());
        }

        let mut remaining = source_text;
        let mut character_offset = 0usize;
        for style in paragraph_styles {
            let (byte_count, character_count) =
                utf16_prefix(remaining, style.characters_covered as usize);
            if byte_count == 0 && !remaining.is_empty() {
                continue;
            }
            let text = remaining[..byte_count].to_string();
            self.paragraph_runs.push(ParagraphRun::with_formatting(
                text,
                start_index + character_offset,
                paragraph_formatting_from_style(style)?,
            ));
            remaining = &remaining[byte_count..];
            character_offset += character_count;
        }

        if !remaining.is_empty() {
            self.paragraph_runs.push(ParagraphRun::with_formatting(
                remaining.to_string(),
                start_index + character_offset,
                ParagraphRunFormatting::default(),
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

    /// Get paragraph-level formatting runs.
    pub fn paragraph_runs(&self) -> &[ParagraphRun] {
        &self.paragraph_runs
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
        property_mask: style.property_mask,
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
        let flags = flags as u16;
        formatting.font_style_raw = Some(flags);
        let (bold, italic, underline) = super::text_prop::extract_char_flags(i32::from(flags));
        formatting.bold = bold;
        formatting.italic = italic;
        formatting.underline = underline;
        formatting.shadow = flags & 0x0010 != 0;
        formatting.embossed = flags & 0x0200 != 0;
        let explicit =
            |mask: u32, bit: u16| (style.property_mask & mask != 0).then_some(flags & bit != 0);
        formatting.bold_explicit = explicit(0x0001, 0x0001);
        formatting.italic_explicit = explicit(0x0002, 0x0002);
        formatting.underline_explicit = explicit(0x0004, 0x0004);
        formatting.shadow_explicit = explicit(0x0010, 0x0010);
        formatting.fe_hint = explicit(0x0020, 0x0020);
        formatting.kumi = explicit(0x0080, 0x0080);
        formatting.legacy_strikethrough = explicit(0x0100, 0x0100);
        formatting.embossed_explicit = explicit(0x0200, 0x0200);
        if style.property_mask & 0x3C00 != 0 {
            formatting.pp9_run_id = Some(((flags >> 10) & 0x0F) as u8);
        }
    }
    Ok(formatting)
}

fn paragraph_formatting_from_style(
    style: &super::text_prop::TextPropCollection,
) -> Result<ParagraphRunFormatting> {
    if style.indent_level > 4 {
        return Err(PptError::Corrupted(
            "TextPFRun indent level exceeds the 0..=4 range".to_string(),
        ));
    }
    let property_mask = style.property_mask;
    let bullet_flags_raw = style.get_value("paragraph.flags").map(|value| value as u16);
    if bullet_flags_raw.is_some_and(|flags| flags & !0x000F != 0) {
        return Err(PptError::Corrupted(
            "TextPFRun has reserved BulletFlags bits set".to_string(),
        ));
    }
    let bullet_flag = |mask: u32, bit: u16| {
        (property_mask & mask != 0).then(|| bullet_flags_raw.is_some_and(|flags| flags & bit != 0))
    };

    let bullet_color_raw = style.get_value("bullet.color").map(|value| value as u32);
    if bullet_color_raw.is_some_and(|raw| !matches!((raw >> 24) as u8, 0x00..=0x07 | 0xFE | 0xFF)) {
        return Err(PptError::Corrupted(
            "TextPFRun has an invalid bullet ColorIndexStruct index".to_string(),
        ));
    }
    let (bullet_color, bullet_scheme_color) = bullet_color_raw
        .map(decode_color_index_struct)
        .unwrap_or((None, None));

    let alignment = style
        .get_value("alignment")
        .map(|value| match value {
            0 => Ok(ParagraphAlignment::Left),
            1 => Ok(ParagraphAlignment::Center),
            2 => Ok(ParagraphAlignment::Right),
            3 => Ok(ParagraphAlignment::Justify),
            4 => Ok(ParagraphAlignment::Distributed),
            5 => Ok(ParagraphAlignment::ThaiDistributed),
            6 => Ok(ParagraphAlignment::JustifyLow),
            _ => Err(PptError::Corrupted(
                "TextPFRun has an invalid TextAlignmentEnum value".to_string(),
            )),
        })
        .transpose()?;
    let font_alignment = style
        .get_value("fontAlignment")
        .map(|value| match value {
            0 => Ok(ParagraphFontAlignment::Roman),
            1 => Ok(ParagraphFontAlignment::Hanging),
            2 => Ok(ParagraphFontAlignment::Center),
            3 => Ok(ParagraphFontAlignment::UpholdFixed),
            _ => Err(PptError::Corrupted(
                "TextPFRun has an invalid TextFontAlignmentEnum value".to_string(),
            )),
        })
        .transpose()?;
    let text_direction = style
        .get_value("textDirection")
        .map(|value| match value {
            0 => Ok(ParagraphTextDirection::LeftToRight),
            1 => Ok(ParagraphTextDirection::RightToLeft),
            _ => Err(PptError::Corrupted(
                "TextPFRun has an invalid TextDirectionEnum value".to_string(),
            )),
        })
        .transpose()?;

    let tab_stops = if property_mask & 0x0010_0000 != 0 {
        Some(
            style
                .tab_stops
                .iter()
                .map(|tab| {
                    let alignment = match tab.alignment {
                        0 => ParagraphTabAlignment::Left,
                        1 => ParagraphTabAlignment::Center,
                        2 => ParagraphTabAlignment::Right,
                        3 => ParagraphTabAlignment::Decimal,
                        _ => {
                            return Err(PptError::Corrupted(
                                "TextPFRun has an invalid TextTabTypeEnum value".to_string(),
                            ));
                        },
                    };
                    Ok(ParagraphTabStop {
                        position: tab.position,
                        alignment,
                    })
                })
                .collect::<Result<Vec<_>>>()?,
        )
    } else {
        None
    };

    let wrap_flags_raw = style.get_value("wrapFlags").map(|value| value as u16);
    if wrap_flags_raw.is_some_and(|flags| flags & !0x0007 != 0) {
        return Err(PptError::Corrupted(
            "TextPFRun has reserved PFWrapFlags bits set".to_string(),
        ));
    }
    let wrap_flag = |mask: u32, bit: u16| {
        (property_mask & mask != 0).then(|| wrap_flags_raw.is_some_and(|flags| flags & bit != 0))
    };

    Ok(ParagraphRunFormatting {
        property_mask,
        indent_level: style.indent_level,
        bullet_flags_raw,
        bullet_enabled: bullet_flag(0x0001, 0x0001),
        bullet_font_enabled: bullet_flag(0x0002, 0x0002),
        bullet_color_enabled: bullet_flag(0x0004, 0x0004),
        bullet_size_enabled: bullet_flag(0x0008, 0x0008),
        bullet_character: style.get_value("bullet.char").map(|value| value as u16),
        bullet_font_index: style.get_value("bullet.font").map(|value| value as u16),
        bullet_size: style.get_value("bullet.size").map(|value| value as i16),
        bullet_color,
        bullet_color_raw,
        bullet_scheme_color,
        alignment,
        line_spacing: style.get_value("linespacing").map(|value| value as i16),
        space_before: style.get_value("spacebefore").map(|value| value as i16),
        space_after: style.get_value("spaceafter").map(|value| value as i16),
        left_margin: style.get_value("text.offset").map(|value| value as i16),
        indent: style.get_value("bullet.offset").map(|value| value as i16),
        default_tab_size: style.get_value("defaultTabSize").map(|value| value as i16),
        tab_stops,
        font_alignment,
        character_wrap: wrap_flag(0x0002_0000, 0x0001),
        word_wrap: wrap_flag(0x0004_0000, 0x0002),
        overflow: wrap_flag(0x0008_0000, 0x0004),
        wrap_flags_raw,
        text_direction,
    })
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
    fn exposes_complete_paragraph_runs_with_utf16_coverage() {
        let text = "😀a\rb";
        let text_record = PptRecord {
            record_type: PptRecordType::TextCharsAtom,
            record_type_raw: 4000,
            version: 0,
            instance: 0,
            data_length: (text.encode_utf16().count() * 2) as u32,
            data: text.encode_utf16().flat_map(u16::to_le_bytes).collect(),
            children: Vec::new(),
        };
        let mask: u32 = 0x000F
            | 0x0010
            | 0x0020
            | 0x0040
            | 0x0080
            | 0x0100
            | 0x0400
            | 0x0800
            | 0x1000
            | 0x2000
            | 0x4000
            | 0x8000
            | 0x0001_0000
            | 0x0002_0000
            | 0x0004_0000
            | 0x0008_0000
            | 0x0010_0000
            | 0x0020_0000;
        let mut style_data = Vec::new();
        style_data.extend_from_slice(&4u32.to_le_bytes());
        style_data.extend_from_slice(&2i16.to_le_bytes());
        style_data.extend_from_slice(&mask.to_le_bytes());
        style_data.extend_from_slice(&0x000Fu16.to_le_bytes());
        style_data.extend_from_slice(&0x2022u16.to_le_bytes());
        style_data.extend_from_slice(&65_535u16.to_le_bytes());
        style_data.extend_from_slice(&(-24i16).to_le_bytes());
        style_data.extend_from_slice(&0xFE33_2211u32.to_le_bytes());
        style_data.extend_from_slice(&4u16.to_le_bytes());
        style_data.extend_from_slice(&120i16.to_le_bytes());
        style_data.extend_from_slice(&(-10i16).to_le_bytes());
        style_data.extend_from_slice(&20i16.to_le_bytes());
        style_data.extend_from_slice(&720i16.to_le_bytes());
        style_data.extend_from_slice(&(-360i16).to_le_bytes());
        style_data.extend_from_slice(&144i16.to_le_bytes());
        style_data.extend_from_slice(&2u16.to_le_bytes());
        style_data.extend_from_slice(&(-20i16).to_le_bytes());
        style_data.extend_from_slice(&1u16.to_le_bytes());
        style_data.extend_from_slice(&720i16.to_le_bytes());
        style_data.extend_from_slice(&3u16.to_le_bytes());
        style_data.extend_from_slice(&3u16.to_le_bytes());
        style_data.extend_from_slice(&5u16.to_le_bytes());
        style_data.extend_from_slice(&1u16.to_le_bytes());
        style_data.extend_from_slice(&2u32.to_le_bytes());
        style_data.extend_from_slice(&1i16.to_le_bytes());
        style_data.extend_from_slice(&0x0800u32.to_le_bytes());
        style_data.extend_from_slice(&2u16.to_le_bytes());
        style_data.extend_from_slice(&6u32.to_le_bytes());
        style_data.extend_from_slice(&0u32.to_le_bytes());
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

        assert_eq!(extractor.paragraph_runs().len(), 2);
        let first = &extractor.paragraph_runs()[0];
        assert_eq!(first.text, "😀a\r");
        assert_eq!(first.start_index, 0);
        assert_eq!(first.length, 3);
        assert_eq!(first.formatting.property_mask, mask);
        assert_eq!(first.formatting.indent_level, 2);
        assert_eq!(first.formatting.bullet_flags_raw, Some(0x000F));
        assert_eq!(first.formatting.bullet_enabled, Some(true));
        assert_eq!(first.formatting.bullet_font_enabled, Some(true));
        assert_eq!(first.formatting.bullet_color_enabled, Some(true));
        assert_eq!(first.formatting.bullet_size_enabled, Some(true));
        assert_eq!(first.formatting.bullet_character, Some(0x2022));
        assert_eq!(first.formatting.bullet_font_index, Some(65_535));
        assert_eq!(first.formatting.bullet_size, Some(-24));
        assert_eq!(first.formatting.bullet_color, Some(0x0011_2233));
        assert_eq!(first.formatting.bullet_color_raw, Some(0xFE33_2211));
        assert_eq!(
            first.formatting.alignment,
            Some(ParagraphAlignment::Distributed)
        );
        assert_eq!(first.formatting.line_spacing, Some(120));
        assert_eq!(first.formatting.space_before, Some(-10));
        assert_eq!(first.formatting.space_after, Some(20));
        assert_eq!(first.formatting.left_margin, Some(720));
        assert_eq!(first.formatting.indent, Some(-360));
        assert_eq!(first.formatting.default_tab_size, Some(144));
        assert_eq!(
            first.formatting.tab_stops,
            Some(vec![
                ParagraphTabStop {
                    position: -20,
                    alignment: ParagraphTabAlignment::Center,
                },
                ParagraphTabStop {
                    position: 720,
                    alignment: ParagraphTabAlignment::Decimal,
                },
            ])
        );
        assert_eq!(
            first.formatting.font_alignment,
            Some(ParagraphFontAlignment::UpholdFixed)
        );
        assert_eq!(first.formatting.wrap_flags_raw, Some(5));
        assert_eq!(first.formatting.character_wrap, Some(true));
        assert_eq!(first.formatting.word_wrap, Some(false));
        assert_eq!(first.formatting.overflow, Some(true));
        assert_eq!(
            first.formatting.text_direction,
            Some(ParagraphTextDirection::RightToLeft)
        );

        let second = &extractor.paragraph_runs()[1];
        assert_eq!(second.text, "b");
        assert_eq!(second.start_index, 3);
        assert_eq!(second.length, 1);
        assert_eq!(second.formatting.indent_level, 1);
        assert_eq!(second.formatting.alignment, Some(ParagraphAlignment::Right));
    }

    #[test]
    fn rejects_invalid_paragraph_enumerations() {
        let mut style = super::super::text_prop::TextPropCollection::new(
            1,
            super::super::text_prop::TextPropType::Paragraph,
        );
        style.property_mask = 0x0010_0000;
        style.tab_stops.push(super::super::text_prop::TextTabStop {
            position: 0,
            alignment: 4,
        });

        let error = paragraph_formatting_from_style(&style).unwrap_err();
        assert!(error.to_string().contains("TextTabTypeEnum"));

        style.indent_level = 5;
        let error = paragraph_formatting_from_style(&style).unwrap_err();
        assert!(error.to_string().contains("indent level"));
    }

    #[test]
    fn preserves_formatting_for_an_empty_paragraph() {
        let text_record = PptRecord {
            record_type: PptRecordType::TextBytesAtom,
            record_type_raw: 4008,
            version: 0,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: Vec::new(),
        };
        let mut style_data = Vec::new();
        style_data.extend_from_slice(&1u32.to_le_bytes());
        style_data.extend_from_slice(&0i16.to_le_bytes());
        style_data.extend_from_slice(&0x0800u32.to_le_bytes());
        style_data.extend_from_slice(&1u16.to_le_bytes());
        style_data.extend_from_slice(&1u32.to_le_bytes());
        style_data.extend_from_slice(&0u32.to_le_bytes());
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

        assert!(extractor.runs().is_empty());
        assert_eq!(extractor.paragraph_runs().len(), 1);
        assert_eq!(extractor.paragraph_runs()[0].text, "");
        assert_eq!(
            extractor.paragraph_runs()[0].formatting.alignment,
            Some(ParagraphAlignment::Center)
        );
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
