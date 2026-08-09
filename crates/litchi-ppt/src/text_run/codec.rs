//! `PowerPoint` text-run extraction and bounded formatting decoding.

use super::model::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphRun, ParagraphRunFormatting,
    ParagraphTabAlignment, ParagraphTabStop, ParagraphTextDirection, TextRun, TextRunFormatting,
};
use super::package::TextRunExtractor;
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use crate::text::extractor::{decode_text_bytes, from_utf16le_lossy};

impl TextRunExtractor {
    /// Extract text runs from PPT records.
    ///
    /// Based on Apache POI's `TextExtractor` and `SlideShow` text parsing logic.
    ///
    /// # Arguments
    ///
    /// * `records` - PPT records to extract text from
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn extract_from_records(&mut self, records: &[Record]) -> Result<()> {
        for record in records {
            self.process_record(record)?;
        }
        Ok(())
    }

    /// Process a single PPT record.
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "`RecordType` mirrors the full MS-PPT record-type enumeration; every record other \
                  than the three text atoms is handled uniformly by recursing into child records"
    )]
    fn process_record(&mut self, record: &Record) -> Result<()> {
        match record.record_type {
            RecordType::TextCharsAtom => {
                // UTF-16LE text
                let text = from_utf16le_lossy(&record.data);
                let start_index = self.text.chars().count();
                self.pending_text = Some((text.clone(), start_index));
                if !text.is_empty() {
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            },
            RecordType::TextBytesAtom => {
                // Low bytes of UTF-16 characters
                let text = decode_text_bytes(&record.data);
                let start_index = self.text.chars().count();
                self.pending_text = Some((text.clone(), start_index));
                if !text.is_empty() {
                    self.text.push_str(&text);
                    self.runs.push(TextRun::new(text, start_index));
                }
            },
            RecordType::StyleTextPropAtom => {
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

    /// Apply style properties from `StyleTextPropAtom`.
    ///
    /// Based on Apache POI's `StyleTextPropAtom` parsing.
    fn apply_style_properties(&mut self, record: &Record) -> Result<()> {
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
            crate::text_prop::parse_style_text_prop_atom_strict(&record.data, text_length)?;

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
        paragraph_styles: &[crate::text_prop::TextPropCollection],
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
}

pub(super) fn formatting_from_style(
    style: &crate::text_prop::TextPropCollection,
) -> Result<TextRunFormatting> {
    let font_color_raw = style.get_value("font.color").map(i32::cast_unsigned);
    if font_color_raw.is_some_and(|raw| !matches!((raw >> 24) as u8, 0x00..=0x07 | 0xFE | 0xFF)) {
        return Err(Error::Corrupted(
            "TextCFRun has an invalid ColorIndexStruct index".to_string(),
        ));
    }
    let (font_color, font_scheme_color) =
        font_color_raw.map_or((None, None), decode_color_index_struct);
    let font_size = style
        .get_value("font.size")
        .map(|size| {
            u16::try_from(size)
                .ok()
                .filter(|points| (1..=4000).contains(points))
                .ok_or_else(|| {
                    Error::Corrupted(
                        "TextCFRun font size is outside the 1..=4000 point range".to_string(),
                    )
                })
        })
        .transpose()?;
    let font_index = |name| -> Result<Option<u16>> {
        style
            .get_value(name)
            .map(|index| {
                u16::try_from(index).map_err(|_err| {
                    Error::Corrupted("TextCFRun has an invalid font index".to_string())
                })
            })
            .transpose()
    };
    let baseline_position = style
        .get_value("superscript")
        .map(|position| {
            i16::try_from(position)
                .ok()
                .filter(|percent| (-100..=100).contains(percent))
                .ok_or_else(|| {
                    Error::Corrupted(
                        "TextCFRun baseline position is outside the -100..=100 range".to_string(),
                    )
                })
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

    if let Some(raw_flags) = style.get_value("char.flags") {
        let flags = u16::try_from(raw_flags).map_err(|_err| {
            Error::Corrupted("TextCFRun has an invalid CFStyle value".to_string())
        })?;
        formatting.font_style_raw = Some(flags);
        let (bold, italic, underline) = crate::text_prop::extract_char_flags(i32::from(flags));
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

pub(super) fn paragraph_formatting_from_style(
    style: &crate::text_prop::TextPropCollection,
) -> Result<ParagraphRunFormatting> {
    if style.indent_level > 4 {
        return Err(Error::Corrupted(
            "TextPFRun indent level exceeds the 0..=4 range".to_string(),
        ));
    }
    let property_mask = style.property_mask;
    let u16_property = |name: &str| -> Result<Option<u16>> {
        style
            .get_value(name)
            .map(|value| {
                u16::try_from(value).map_err(|_err| {
                    Error::Corrupted(format!("TextPFRun has an invalid {name} value"))
                })
            })
            .transpose()
    };
    let i16_property = |name: &str| -> Result<Option<i16>> {
        style
            .get_value(name)
            .map(|value| {
                i16::try_from(value).map_err(|_err| {
                    Error::Corrupted(format!("TextPFRun has an invalid {name} value"))
                })
            })
            .transpose()
    };
    let bullet_flags_raw = u16_property("paragraph.flags")?;
    if bullet_flags_raw.is_some_and(|flags| flags & !0x000F != 0) {
        return Err(Error::Corrupted(
            "TextPFRun has reserved BulletFlags bits set".to_string(),
        ));
    }
    let bullet_flag = |mask: u32, bit: u16| {
        (property_mask & mask != 0).then(|| bullet_flags_raw.is_some_and(|flags| flags & bit != 0))
    };

    let bullet_color_raw = style.get_value("bullet.color").map(i32::cast_unsigned);
    if bullet_color_raw.is_some_and(|raw| !matches!((raw >> 24) as u8, 0x00..=0x07 | 0xFE | 0xFF)) {
        return Err(Error::Corrupted(
            "TextPFRun has an invalid bullet ColorIndexStruct index".to_string(),
        ));
    }
    let (bullet_color, bullet_scheme_color) =
        bullet_color_raw.map_or((None, None), decode_color_index_struct);

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
            _ => Err(Error::Corrupted(
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
            _ => Err(Error::Corrupted(
                "TextPFRun has an invalid TextFontAlignmentEnum value".to_string(),
            )),
        })
        .transpose()?;
    let text_direction = style
        .get_value("textDirection")
        .map(|value| match value {
            0 => Ok(ParagraphTextDirection::LeftToRight),
            1 => Ok(ParagraphTextDirection::RightToLeft),
            _ => Err(Error::Corrupted(
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
                    let tab_alignment = match tab.alignment {
                        0 => ParagraphTabAlignment::Left,
                        1 => ParagraphTabAlignment::Center,
                        2 => ParagraphTabAlignment::Right,
                        3 => ParagraphTabAlignment::Decimal,
                        _ => {
                            return Err(Error::Corrupted(
                                "TextPFRun has an invalid TextTabTypeEnum value".to_string(),
                            ));
                        },
                    };
                    Ok(ParagraphTabStop {
                        position: tab.position,
                        alignment: tab_alignment,
                    })
                })
                .collect::<Result<Vec<_>>>()?,
        )
    } else {
        None
    };

    let wrap_flags_raw = u16_property("wrapFlags")?;
    if wrap_flags_raw.is_some_and(|flags| flags & !0x0007 != 0) {
        return Err(Error::Corrupted(
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
        bullet_character: u16_property("bullet.char")?,
        bullet_font_index: u16_property("bullet.font")?,
        bullet_size: i16_property("bullet.size")?,
        bullet_color,
        bullet_color_raw,
        bullet_scheme_color,
        alignment,
        line_spacing: i16_property("linespacing")?,
        space_before: i16_property("spacebefore")?,
        space_after: i16_property("spaceafter")?,
        left_margin: i16_property("text.offset")?,
        indent: i16_property("bullet.offset")?,
        default_tab_size: i16_property("defaultTabSize")?,
        tab_stops,
        font_alignment,
        character_wrap: wrap_flag(0x0002_0000, 0x0001),
        word_wrap: wrap_flag(0x0004_0000, 0x0002),
        overflow: wrap_flag(0x0008_0000, 0x0004),
        wrap_flags_raw,
        text_direction,
    })
}

pub(super) fn decode_color_index_struct(raw: u32) -> (Option<u32>, Option<u8>) {
    let red = raw & 0xFF;
    let green = (raw >> 8) & 0xFF;
    let blue = (raw >> 16) & 0xFF;
    match (raw >> 24) as u8 {
        0xFE => (Some((red << 16) | (green << 8) | blue), None),
        index @ 0x00..=0x07 => (None, Some(index)),
        _ => (None, None),
    }
}

pub(super) fn utf16_prefix(text: &str, requested_units: usize) -> (usize, usize) {
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
