//! Versioned PowerPoint text-formatting extensions.

use litchi_core::binary::{read_i16_le, read_u16_le, read_u32_le};

use super::package::{PptError, Result};

/// PowerPoint 9 paragraph extensions from `TextPFException9`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextParagraphExtension9 {
    /// Original `PFMasks` value.
    pub mask: u32,
    /// Picture-bullet reference, when present.
    pub bullet_blip_ref: Option<i16>,
    /// Whether automatic numbering is active, when explicitly present.
    pub auto_numbered: Option<bool>,
    /// Raw `TextAutoNumberSchemeEnum` value.
    pub auto_number_scheme: Option<u16>,
    /// First number in the sequence.
    pub auto_number_start: Option<i16>,
}

impl TextParagraphExtension9 {
    /// Parse one `TextPFException9` from the start of `data`.
    ///
    /// Returns the decoded extension and the number of bytes consumed.
    pub fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextPFException9 mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        if mask & !0x0380_0000 != 0 {
            return Err(PptError::Corrupted(
                "TextPFException9 has unsupported paragraph mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;

        let bullet_blip_ref = if mask & 0x0080_0000 != 0 {
            require_bytes(data, offset, 2, "TextPFException9 bulletBlipRef")?;
            let value = read_i16_le(data, offset).unwrap_or(0);
            if value < -1 {
                return Err(PptError::Corrupted(
                    "TextPFException9 has an invalid picture-bullet reference".to_string(),
                ));
            }
            offset += 2;
            Some(value)
        } else {
            None
        };
        let auto_numbered = if mask & 0x0200_0000 != 0 {
            require_bytes(data, offset, 2, "TextPFException9 auto-number flag")?;
            let value = read_i16_le(data, offset).unwrap_or(0);
            offset += 2;
            match value {
                0 => Some(false),
                1 => Some(true),
                _ => {
                    return Err(PptError::Corrupted(
                        "TextPFException9 has an invalid auto-number flag".to_string(),
                    ));
                },
            }
        } else {
            None
        };
        let (auto_number_scheme, auto_number_start) = if mask & 0x0100_0000 != 0 {
            require_bytes(data, offset, 4, "TextPFException9 auto-number scheme")?;
            let scheme = read_u16_le(data, offset).unwrap_or(0);
            let start = read_i16_le(data, offset + 2).unwrap_or(0);
            if scheme > 0x0028 {
                return Err(PptError::Corrupted(
                    "TextPFException9 has an invalid auto-number scheme".to_string(),
                ));
            }
            if start < 1 {
                return Err(PptError::Corrupted(
                    "TextPFException9 auto-number start must be positive".to_string(),
                ));
            }
            offset += 4;
            (Some(scheme), Some(start))
        } else {
            (None, None)
        };

        Ok((
            Self {
                mask,
                bullet_blip_ref,
                auto_numbered,
                auto_number_scheme,
                auto_number_start,
            },
            offset,
        ))
    }

    /// Effective auto-number scheme, applying the MS-PPT default.
    pub fn effective_auto_number_scheme(&self) -> Option<u16> {
        self.auto_numbered?
            .then_some(self.auto_number_scheme.unwrap_or(0x0003))
    }

    /// Effective starting number, applying the MS-PPT default.
    pub fn effective_auto_number_start(&self) -> Option<i16> {
        self.auto_numbered?
            .then_some(self.auto_number_start.unwrap_or(1))
    }
}

/// PowerPoint 9 character extensions from `TextCFException9`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextCharacterExtension9 {
    /// Original `CFMasks` value.
    pub mask: u32,
    /// Raw optional PP10 extension word.
    pub pp10_extension: Option<u32>,
    /// Four-bit run identifier used by `StyleTextProp10Atom`.
    pub pp10_run_id: Option<u8>,
}

impl TextCharacterExtension9 {
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextCFException9 mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        // Bits 3, 6, 8, 14, and 15 are undefined fields that readers ignore.
        // All other fields except pp10ext MUST be zero in TextCFException9.
        if mask & !0x0010_c148 != 0 {
            return Err(PptError::Corrupted(
                "TextCFException9 has unsupported character mask bits".to_string(),
            ));
        }

        let pp10_extension = if mask & 0x0010_0000 != 0 {
            require_bytes(data, 4, 4, "TextCFException9 PP10 extension")?;
            Some(read_u32_le(data, 4).unwrap_or(0))
        } else {
            None
        };
        let consumed = if pp10_extension.is_some() { 8 } else { 4 };
        Ok((
            Self {
                mask,
                pp10_extension,
                pp10_run_id: pp10_extension.map(|value| (value & 0x0f) as u8),
            },
            consumed,
        ))
    }
}

/// PowerPoint 9 special information from the `StyleTextProp9` SI subset.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextSpecialInfoExtension9 {
    /// Original `TextSIException` mask.
    pub mask: u32,
    /// Whether the text contains bidirectional characters, when present.
    pub bidi: Option<bool>,
    /// Raw optional PP10 text extension word.
    pub pp10_extension: Option<u32>,
    /// Four-bit run identifier used by `StyleTextProp11Atom`.
    pub pp10_run_id: Option<u8>,
    /// Whether a grammar error is present, when the PP10 extension exists.
    pub grammar_error: Option<bool>,
}

impl TextSpecialInfoExtension9 {
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "StyleTextProp9 TextSIException mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        // StyleTextProp9 forbids spell, language, alternate language, and smart
        // tags. Bits 3, 4, and 7 are undefined and are ignored.
        if mask & !0x0000_00f8 != 0 {
            return Err(PptError::Corrupted(
                "StyleTextProp9 has unsupported special-info mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;

        let bidi = if mask & 0x40 != 0 {
            require_bytes(data, offset, 2, "StyleTextProp9 bidi flag")?;
            let value = read_i16_le(data, offset).unwrap_or(0);
            offset += 2;
            match value {
                0 => Some(false),
                1 => Some(true),
                _ => {
                    return Err(PptError::Corrupted(
                        "StyleTextProp9 has an invalid bidi flag".to_string(),
                    ));
                },
            }
        } else {
            None
        };

        let pp10_extension = if mask & 0x20 != 0 {
            require_bytes(data, offset, 4, "StyleTextProp9 PP10 text extension")?;
            let value = read_u32_le(data, offset).unwrap_or(0);
            if value & 0x7fff_fff0 != 0 {
                return Err(PptError::Corrupted(
                    "StyleTextProp9 PP10 text extension has reserved bits".to_string(),
                ));
            }
            offset += 4;
            Some(value)
        } else {
            None
        };

        Ok((
            Self {
                mask,
                bidi,
                pp10_extension,
                pp10_run_id: pp10_extension.map(|value| (value & 0x0f) as u8),
                grammar_error: pp10_extension.map(|value| value & 0x8000_0000 != 0),
            },
            offset,
        ))
    }
}

/// Additional formatting for one PowerPoint 9 text-style run group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextStyleExtension9Run {
    /// Additional paragraph formatting.
    pub paragraph: TextParagraphExtension9,
    /// Additional character formatting.
    pub character: TextCharacterExtension9,
    /// Additional special text information.
    pub special_info: TextSpecialInfoExtension9,
}

/// Parsed payload of a `StyleTextProp9Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStyleExtension9 {
    /// Formatting tuples indexed by the base character run's `pp9rt` value.
    pub runs: Vec<TextStyleExtension9Run>,
}

impl TextStyleExtension9 {
    /// Parse a complete `StyleTextProp9Atom` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut offset = 0usize;
        let mut runs = Vec::new();
        while offset < data.len() {
            let (paragraph, consumed) = TextParagraphExtension9::parse_prefix(&data[offset..])?;
            offset += consumed;
            let (character, consumed) = TextCharacterExtension9::parse_prefix(&data[offset..])?;
            offset += consumed;
            let (special_info, consumed) =
                TextSpecialInfoExtension9::parse_prefix(&data[offset..])?;
            offset += consumed;
            runs.push(TextStyleExtension9Run {
                paragraph,
                character,
                special_info,
            });
        }
        Ok(Self { runs })
    }
}

fn require_bytes(data: &[u8], offset: usize, size: usize, field: &str) -> Result<()> {
    let end = offset
        .checked_add(size)
        .ok_or_else(|| PptError::Corrupted(format!("{field} offset overflow")))?;
    if end > data.len() {
        return Err(PptError::Corrupted(format!("Truncated {field}")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_powerpoint_9_auto_numbering() {
        let mask = 0x0380_0000u32;
        let mut data = Vec::new();
        data.extend_from_slice(&mask.to_le_bytes());
        data.extend_from_slice(&(-1i16).to_le_bytes());
        data.extend_from_slice(&1i16.to_le_bytes());
        data.extend_from_slice(&0x0028u16.to_le_bytes());
        data.extend_from_slice(&7i16.to_le_bytes());

        let (extension, consumed) = TextParagraphExtension9::parse_prefix(&data).unwrap();

        assert_eq!(consumed, data.len());
        assert_eq!(extension.bullet_blip_ref, Some(-1));
        assert_eq!(extension.auto_numbered, Some(true));
        assert_eq!(extension.auto_number_scheme, Some(0x0028));
        assert_eq!(extension.auto_number_start, Some(7));
        assert_eq!(extension.effective_auto_number_scheme(), Some(0x0028));
        assert_eq!(extension.effective_auto_number_start(), Some(7));
    }

    #[test]
    fn applies_auto_number_defaults_and_rejects_invalid_values() {
        let mut defaults = Vec::new();
        defaults.extend_from_slice(&0x0200_0000u32.to_le_bytes());
        defaults.extend_from_slice(&1i16.to_le_bytes());
        let (extension, _) = TextParagraphExtension9::parse_prefix(&defaults).unwrap();
        assert_eq!(extension.effective_auto_number_scheme(), Some(0x0003));
        assert_eq!(extension.effective_auto_number_start(), Some(1));

        let mut invalid_flag = defaults;
        invalid_flag[4..6].copy_from_slice(&2i16.to_le_bytes());
        assert!(TextParagraphExtension9::parse_prefix(&invalid_flag).is_err());

        let mut invalid_scheme = Vec::new();
        invalid_scheme.extend_from_slice(&0x0100_0000u32.to_le_bytes());
        invalid_scheme.extend_from_slice(&0x0029u16.to_le_bytes());
        invalid_scheme.extend_from_slice(&0i16.to_le_bytes());
        assert!(TextParagraphExtension9::parse_prefix(&invalid_scheme).is_err());

        assert!(TextParagraphExtension9::parse_prefix(&[0, 0, 0]).is_err());

        let unsupported = 1u32.to_le_bytes();
        assert!(TextParagraphExtension9::parse_prefix(&unsupported).is_err());

        let mut invalid_blip = Vec::new();
        invalid_blip.extend_from_slice(&0x0080_0000u32.to_le_bytes());
        invalid_blip.extend_from_slice(&(-2i16).to_le_bytes());
        assert!(TextParagraphExtension9::parse_prefix(&invalid_blip).is_err());
    }

    #[test]
    fn parses_complete_style_text_prop9_tuples() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0200_0000u32.to_le_bytes());
        data.extend_from_slice(&1i16.to_le_bytes());
        data.extend_from_slice(&0x0010_0000u32.to_le_bytes());
        data.extend_from_slice(&0xdead_beeau32.to_le_bytes());
        data.extend_from_slice(&0x60u32.to_le_bytes());
        data.extend_from_slice(&1i16.to_le_bytes());
        data.extend_from_slice(&0x8000_000bu32.to_le_bytes());

        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());

        let style = TextStyleExtension9::parse(&data).unwrap();

        assert_eq!(style.runs.len(), 2);
        assert_eq!(style.runs[0].paragraph.auto_numbered, Some(true));
        assert_eq!(style.runs[0].character.pp10_run_id, Some(10));
        assert_eq!(style.runs[0].special_info.bidi, Some(true));
        assert_eq!(style.runs[0].special_info.pp10_run_id, Some(11));
        assert_eq!(style.runs[0].special_info.grammar_error, Some(true));
        assert_eq!(style.runs[1].character.pp10_run_id, None);
    }

    #[test]
    fn rejects_malformed_style_text_prop9_tuples() {
        let mut unsupported_cf = Vec::new();
        unsupported_cf.extend_from_slice(&0u32.to_le_bytes());
        unsupported_cf.extend_from_slice(&1u32.to_le_bytes());
        unsupported_cf.extend_from_slice(&0u32.to_le_bytes());
        assert!(TextStyleExtension9::parse(&unsupported_cf).is_err());

        let mut unsupported_si = Vec::new();
        unsupported_si.extend_from_slice(&0u32.to_le_bytes());
        unsupported_si.extend_from_slice(&0u32.to_le_bytes());
        unsupported_si.extend_from_slice(&1u32.to_le_bytes());
        assert!(TextStyleExtension9::parse(&unsupported_si).is_err());

        let mut invalid_bidi = Vec::new();
        invalid_bidi.extend_from_slice(&0u32.to_le_bytes());
        invalid_bidi.extend_from_slice(&0u32.to_le_bytes());
        invalid_bidi.extend_from_slice(&0x40u32.to_le_bytes());
        invalid_bidi.extend_from_slice(&2i16.to_le_bytes());
        assert!(TextStyleExtension9::parse(&invalid_bidi).is_err());

        let mut reserved_pp10 = Vec::new();
        reserved_pp10.extend_from_slice(&0u32.to_le_bytes());
        reserved_pp10.extend_from_slice(&0u32.to_le_bytes());
        reserved_pp10.extend_from_slice(&0x20u32.to_le_bytes());
        reserved_pp10.extend_from_slice(&0x10u32.to_le_bytes());
        assert!(TextStyleExtension9::parse(&reserved_pp10).is_err());

        assert!(TextStyleExtension9::parse(&[0; 11]).is_err());
    }
}
