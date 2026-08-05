//! Versioned PowerPoint text-formatting extensions.

use litchi_core::binary::{read_i16_le, read_u16_le, read_u32_le};

use super::package::{Error, Result};
use super::records::Record;

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
            return Err(Error::Corrupted(
                "TextPFException9 has unsupported paragraph mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;

        let bullet_blip_ref = if mask & 0x0080_0000 != 0 {
            require_bytes(data, offset, 2, "TextPFException9 bulletBlipRef")?;
            let value = read_i16_le(data, offset).unwrap_or(0);
            if value < -1 {
                return Err(Error::Corrupted(
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
                    return Err(Error::Corrupted(
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
                return Err(Error::Corrupted(
                    "TextPFException9 has an invalid auto-number scheme".to_string(),
                ));
            }
            if start < 1 {
                return Err(Error::Corrupted(
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
            return Err(Error::Corrupted(
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
            return Err(Error::Corrupted(
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
                    return Err(Error::Corrupted(
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
                return Err(Error::Corrupted(
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

/// Additional character formatting from one `TextCFException10`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextCharacterExtension10 {
    /// Original `CFMasks` value.
    pub mask: u32,
    /// Font index in `FontCollection10Container` for East Asian text.
    pub new_east_asian_font_ref: Option<u16>,
    /// Font index in `FontCollection10Container` for complex-script text.
    pub complex_script_font_ref: Option<u16>,
    /// Raw undefined `pp11ext` word, preserved when present.
    pub pp11_extension: Option<u32>,
}

impl TextCharacterExtension10 {
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextCFException10 mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        // Bits 3, 6, 8, 14, and 15 are undefined fields that readers ignore.
        // Only the three PowerPoint 10 fields may otherwise be set.
        if mask & !0x0700_c148 != 0 {
            return Err(Error::Corrupted(
                "TextCFException10 has unsupported character mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;
        let new_east_asian_font_ref =
            read_optional_u16(data, &mut offset, mask & 0x0100_0000 != 0, "new EA font")?;
        let complex_script_font_ref = read_optional_u16(
            data,
            &mut offset,
            mask & 0x0200_0000 != 0,
            "complex-script font",
        )?;
        let pp11_extension = if mask & 0x0400_0000 != 0 {
            require_bytes(data, offset, 4, "TextCFException10 PP11 extension")?;
            let value = read_u32_le(data, offset).unwrap_or(0);
            offset += 4;
            Some(value)
        } else {
            None
        };
        Ok((
            Self {
                mask,
                new_east_asian_font_ref,
                complex_script_font_ref,
                pp11_extension,
            },
            offset,
        ))
    }
}

/// Parsed payload of a `StyleTextProp10Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStyleExtension10 {
    /// Character extensions indexed through PowerPoint 9 `pp10runid` values.
    pub runs: Vec<TextCharacterExtension10>,
}

impl TextStyleExtension10 {
    /// Parse a complete `StyleTextProp10Atom` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut offset = 0usize;
        let mut runs = Vec::new();
        while offset < data.len() {
            let (run, consumed) = TextCharacterExtension10::parse_prefix(&data[offset..])?;
            offset += consumed;
            runs.push(run);
        }
        Ok(Self { runs })
    }
}

/// Smart-tag information from one PowerPoint 11 `StyleTextProp11` run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextSpecialInfoExtension11 {
    /// Original `TextSIException` mask.
    pub mask: u32,
    /// Indices into the document-wide smart-tag store.
    pub smart_tag_indices: Vec<u32>,
}

impl TextSpecialInfoExtension11 {
    fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "StyleTextProp11 TextSIException mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        // StyleTextProp11 permits only smart tags. Bits 3, 4, and 7 are
        // undefined and ignored; all other TextSIException fields MUST be zero.
        if mask & !0x0000_0298 != 0 {
            return Err(Error::Corrupted(
                "StyleTextProp11 has unsupported special-info mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;
        let mut smart_tag_indices = Vec::new();
        if mask & 0x0200 != 0 {
            require_bytes(data, offset, 4, "StyleTextProp11 smart-tag count")?;
            let count = read_u32_le(data, offset).unwrap_or(0);
            offset += 4;
            let count = usize::try_from(count).map_err(|_| {
                Error::Corrupted("StyleTextProp11 smart-tag count overflow".to_string())
            })?;
            let byte_count = count.checked_mul(4).ok_or_else(|| {
                Error::Corrupted("StyleTextProp11 smart-tag size overflow".to_string())
            })?;
            require_bytes(
                data,
                offset,
                byte_count,
                "StyleTextProp11 smart-tag indices",
            )?;
            smart_tag_indices.reserve(count);
            for _ in 0..count {
                smart_tag_indices.push(read_u32_le(data, offset).unwrap_or(0));
                offset += 4;
            }
        }
        Ok((
            Self {
                mask,
                smart_tag_indices,
            },
            offset,
        ))
    }
}

/// Parsed payload of a `StyleTextProp11Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStyleExtension11 {
    /// Smart-tag runs indexed through PowerPoint 9 SI `pp10runid` values.
    pub runs: Vec<TextSpecialInfoExtension11>,
}

impl TextStyleExtension11 {
    /// Parse a complete `StyleTextProp11Atom` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut offset = 0usize;
        let mut runs = Vec::new();
        while offset < data.len() {
            let (run, consumed) = TextSpecialInfoExtension11::parse_prefix(&data[offset..])?;
            offset += consumed;
            runs.push(run);
        }
        Ok(Self { runs })
    }
}

/// Additional PowerPoint 9 formatting for one master indent level.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextMasterStyleExtension9Level {
    /// Additional paragraph formatting.
    pub paragraph: TextParagraphExtension9,
    /// Additional character formatting.
    pub character: TextCharacterExtension9,
}

/// Parsed payload of a `TextMasterStyle9Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextMasterStyleExtension9 {
    /// `TextTypeEnum` value from the record instance.
    pub text_type: u16,
    /// Additional formatting for indent levels `0..levels.len()`.
    pub levels: Vec<TextMasterStyleExtension9Level>,
}

impl TextMasterStyleExtension9 {
    /// Parse a complete `TextMasterStyle9Atom` payload.
    pub fn parse(data: &[u8], text_type: u16) -> Result<Self> {
        validate_text_type(text_type, "TextMasterStyle9Atom")?;
        require_bytes(data, 0, 2, "TextMasterStyle9Atom level count")?;
        let level_count = read_u16_le(data, 0).unwrap_or(0);
        if level_count > 5 {
            return Err(Error::Corrupted(
                "TextMasterStyle9Atom has more than five levels".to_string(),
            ));
        }
        let mut offset = 2usize;
        let mut levels = Vec::with_capacity(level_count as usize);
        for _ in 0..level_count {
            let (paragraph, consumed) = TextParagraphExtension9::parse_prefix(&data[offset..])?;
            offset += consumed;
            let (character, consumed) = TextCharacterExtension9::parse_prefix(&data[offset..])?;
            offset += consumed;
            levels.push(TextMasterStyleExtension9Level {
                paragraph,
                character,
            });
        }
        if offset != data.len() {
            return Err(Error::Corrupted(
                "TextMasterStyle9Atom has trailing bytes".to_string(),
            ));
        }
        Ok(Self { text_type, levels })
    }
}

/// Parsed payload of a `TextMasterStyle10Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextMasterStyleExtension10 {
    /// `TextTypeEnum` value from the record instance.
    pub text_type: u16,
    /// Additional character formatting for each indent level.
    pub levels: Vec<TextCharacterExtension10>,
}

impl TextMasterStyleExtension10 {
    /// Parse a complete `TextMasterStyle10Atom` payload.
    pub fn parse(data: &[u8], text_type: u16) -> Result<Self> {
        validate_text_type(text_type, "TextMasterStyle10Atom")?;
        require_bytes(data, 0, 2, "TextMasterStyle10Atom level count")?;
        let level_count = read_u16_le(data, 0).unwrap_or(0);
        if level_count > 5 {
            return Err(Error::Corrupted(
                "TextMasterStyle10Atom has more than five levels".to_string(),
            ));
        }
        let mut offset = 2usize;
        let mut levels = Vec::with_capacity(level_count as usize);
        for _ in 0..level_count {
            let (level, consumed) = TextCharacterExtension10::parse_prefix(&data[offset..])?;
            offset += consumed;
            levels.push(level);
        }
        if offset != data.len() {
            return Err(Error::Corrupted(
                "TextMasterStyle10Atom has trailing bytes".to_string(),
            ));
        }
        Ok(Self { text_type, levels })
    }
}

/// Parsed payload of a document-level `TextDefaults9Atom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextDefaultsExtension9 {
    /// Default PowerPoint 9 character formatting.
    pub character: TextCharacterExtension9,
    /// Default PowerPoint 9 paragraph formatting.
    pub paragraph: TextParagraphExtension9,
}

impl TextDefaultsExtension9 {
    /// Parse a complete `TextDefaults9Atom` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (character, consumed) = TextCharacterExtension9::parse_prefix(data)?;
        let (paragraph, paragraph_size) = TextParagraphExtension9::parse_prefix(&data[consumed..])?;
        if consumed + paragraph_size != data.len() {
            return Err(Error::Corrupted(
                "TextDefaults9Atom has trailing bytes".to_string(),
            ));
        }
        Ok(Self {
            character,
            paragraph,
        })
    }
}

/// Parsed payload of a document-level `TextDefaults10Atom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextDefaultsExtension10 {
    /// Default PowerPoint 10 character formatting.
    pub character: TextCharacterExtension10,
}

impl TextDefaultsExtension10 {
    /// Parse a complete `TextDefaults10Atom` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (character, consumed) = TextCharacterExtension10::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(Error::Corrupted(
                "TextDefaults10Atom has trailing bytes".to_string(),
            ));
        }
        Ok(Self { character })
    }
}

fn validate_text_type(text_type: u16, record: &str) -> Result<()> {
    if matches!(text_type, 0 | 1 | 2 | 4 | 5 | 6 | 7 | 8) {
        Ok(())
    } else {
        Err(Error::Corrupted(format!(
            "{record} has an invalid TextTypeEnum instance"
        )))
    }
}

/// Versioned text master styles collected from a PPT record tree.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct VersionedTextMasterStyles {
    /// PowerPoint 9 paragraph and character master extensions.
    pub powerpoint9: Vec<TextMasterStyleExtension9>,
    /// PowerPoint 10 character master extensions.
    pub powerpoint10: Vec<TextMasterStyleExtension10>,
}

impl VersionedTextMasterStyles {
    /// Collect and parse all versioned master-style atoms below `root`.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut result = Self::default();
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != crate::consts::RecordType::TextMasterStyle9Atom {
                continue;
            }
            validate_atom_header(&record, "TextMasterStyle9Atom", false)?;
            result.powerpoint9.push(TextMasterStyleExtension9::parse(
                &record.data,
                record.instance,
            )?);
        }
        for record in root.versioned_binary_tag_records(10)? {
            if record.record_type != crate::consts::RecordType::TextMasterStyle10Atom {
                continue;
            }
            validate_atom_header(&record, "TextMasterStyle10Atom", false)?;
            result.powerpoint10.push(TextMasterStyleExtension10::parse(
                &record.data,
                record.instance,
            )?);
        }
        Ok(result)
    }
}

/// Versioned document-wide text defaults collected from a PPT record tree.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct VersionedTextDefaults {
    /// PowerPoint 9 default paragraph and character extensions.
    pub powerpoint9: Option<TextDefaultsExtension9>,
    /// PowerPoint 10 default character extensions.
    pub powerpoint10: Option<TextDefaultsExtension10>,
}

impl VersionedTextDefaults {
    /// Collect and parse document-level text-default atoms below `root`.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut result = Self::default();
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != crate::consts::RecordType::TextDefaults9Atom {
                continue;
            }
            validate_atom_header(&record, "TextDefaults9Atom", true)?;
            if result
                .powerpoint9
                .replace(TextDefaultsExtension9::parse(&record.data)?)
                .is_some()
            {
                return Err(Error::Corrupted(
                    "Record tree contains multiple TextDefaults9Atom records".to_string(),
                ));
            }
        }
        for record in root.versioned_binary_tag_records(10)? {
            if record.record_type != crate::consts::RecordType::TextDefaults10Atom {
                continue;
            }
            validate_atom_header(&record, "TextDefaults10Atom", true)?;
            if result
                .powerpoint10
                .replace(TextDefaultsExtension10::parse(&record.data)?)
                .is_some()
            {
                return Err(Error::Corrupted(
                    "Record tree contains multiple TextDefaults10Atom records".to_string(),
                ));
            }
        }
        Ok(result)
    }
}

fn validate_atom_header(record: &Record, name: &str, zero_instance: bool) -> Result<()> {
    if record.version != 0 || zero_instance && record.instance != 0 {
        return Err(Error::Corrupted(format!(
            "{name} has an invalid record header"
        )));
    }
    Ok(())
}

fn read_optional_u16(
    data: &[u8],
    offset: &mut usize,
    present: bool,
    field: &str,
) -> Result<Option<u16>> {
    if !present {
        return Ok(None);
    }
    require_bytes(data, *offset, 2, field)?;
    let value = read_u16_le(data, *offset).unwrap_or(0);
    *offset += 2;
    Ok(Some(value))
}

fn require_bytes(data: &[u8], offset: usize, size: usize, field: &str) -> Result<()> {
    let end = offset
        .checked_add(size)
        .ok_or_else(|| Error::Corrupted(format!("{field} offset overflow")))?;
    if end > data.len() {
        return Err(Error::Corrupted(format!("Truncated {field}")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ppt_record_bytes(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&record_type.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> Record {
        let tag_name: Vec<u8> = format!("___PPT{version}")
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = ppt_record_bytes(0, 0, 4026, &tag_name);
        let blob = ppt_record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = ppt_record_bytes(0x0f, 0, 0x138a, &tag_payload);
        Record {
            record_type: crate::consts::RecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

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

    #[test]
    fn parses_powerpoint_10_character_extensions() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0700_0000u32.to_le_bytes());
        data.extend_from_slice(&65_535u16.to_le_bytes());
        data.extend_from_slice(&42u16.to_le_bytes());
        data.extend_from_slice(&0xdead_beefu32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());

        let style = TextStyleExtension10::parse(&data).unwrap();

        assert_eq!(style.runs.len(), 2);
        assert_eq!(style.runs[0].new_east_asian_font_ref, Some(65_535));
        assert_eq!(style.runs[0].complex_script_font_ref, Some(42));
        assert_eq!(style.runs[0].pp11_extension, Some(0xdead_beef));
        assert_eq!(style.runs[1].mask, 0);
    }

    #[test]
    fn rejects_malformed_powerpoint_10_character_extensions() {
        assert!(TextStyleExtension10::parse(&[0; 3]).is_err());
        assert!(TextStyleExtension10::parse(&1u32.to_le_bytes()).is_err());

        let truncated = 0x0100_0000u32.to_le_bytes();
        assert!(TextStyleExtension10::parse(&truncated).is_err());
    }

    #[test]
    fn parses_powerpoint_11_smart_tag_runs() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0200u32.to_le_bytes());
        data.extend_from_slice(&3u32.to_le_bytes());
        data.extend_from_slice(&7u32.to_le_bytes());
        data.extend_from_slice(&11u32.to_le_bytes());
        data.extend_from_slice(&13u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());

        let style = TextStyleExtension11::parse(&data).unwrap();

        assert_eq!(style.runs.len(), 2);
        assert_eq!(style.runs[0].smart_tag_indices, vec![7, 11, 13]);
        assert!(style.runs[1].smart_tag_indices.is_empty());
    }

    #[test]
    fn rejects_malformed_powerpoint_11_smart_tag_runs() {
        assert!(TextStyleExtension11::parse(&[0; 3]).is_err());
        assert!(TextStyleExtension11::parse(&1u32.to_le_bytes()).is_err());

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0200u32.to_le_bytes());
        truncated.extend_from_slice(&2u32.to_le_bytes());
        truncated.extend_from_slice(&1u32.to_le_bytes());
        assert!(TextStyleExtension11::parse(&truncated).is_err());
    }

    #[test]
    fn parses_powerpoint_9_and_10_master_style_extensions() {
        let mut master9 = Vec::new();
        master9.extend_from_slice(&1u16.to_le_bytes());
        master9.extend_from_slice(&0x0200_0000u32.to_le_bytes());
        master9.extend_from_slice(&1i16.to_le_bytes());
        master9.extend_from_slice(&0x0010_0000u32.to_le_bytes());
        master9.extend_from_slice(&0x1234_567au32.to_le_bytes());
        let style9 = TextMasterStyleExtension9::parse(&master9, 1).unwrap();
        assert_eq!(style9.levels.len(), 1);
        assert_eq!(style9.levels[0].paragraph.auto_numbered, Some(true));
        assert_eq!(style9.levels[0].character.pp10_run_id, Some(10));

        let mut master10 = Vec::new();
        master10.extend_from_slice(&2u16.to_le_bytes());
        master10.extend_from_slice(&0x0100_0000u32.to_le_bytes());
        master10.extend_from_slice(&65_535u16.to_le_bytes());
        master10.extend_from_slice(&0u32.to_le_bytes());
        let style10 = TextMasterStyleExtension10::parse(&master10, 5).unwrap();
        assert_eq!(style10.levels.len(), 2);
        assert_eq!(style10.levels[0].new_east_asian_font_ref, Some(65_535));
        assert_eq!(style10.levels[1].mask, 0);
    }

    #[test]
    fn rejects_malformed_versioned_master_style_extensions() {
        assert!(TextMasterStyleExtension9::parse(&[], 1).is_err());
        assert!(TextMasterStyleExtension9::parse(&6u16.to_le_bytes(), 1).is_err());
        assert!(TextMasterStyleExtension9::parse(&0u16.to_le_bytes(), 3).is_err());

        let mut trailing9 = 0u16.to_le_bytes().to_vec();
        trailing9.push(0);
        assert!(TextMasterStyleExtension9::parse(&trailing9, 1).is_err());

        let mut truncated10 = Vec::new();
        truncated10.extend_from_slice(&1u16.to_le_bytes());
        truncated10.extend_from_slice(&0x0100_0000u32.to_le_bytes());
        assert!(TextMasterStyleExtension10::parse(&truncated10, 1).is_err());
        assert!(TextMasterStyleExtension10::parse(&0u16.to_le_bytes(), 9).is_err());
    }

    #[test]
    fn parses_and_validates_document_text_defaults() {
        let mut defaults9 = Vec::new();
        defaults9.extend_from_slice(&0u32.to_le_bytes());
        defaults9.extend_from_slice(&0x0200_0000u32.to_le_bytes());
        defaults9.extend_from_slice(&1i16.to_le_bytes());
        let defaults9 = TextDefaultsExtension9::parse(&defaults9).unwrap();
        assert_eq!(defaults9.character.mask, 0);
        assert_eq!(defaults9.paragraph.auto_numbered, Some(true));

        let mut defaults10 = Vec::new();
        defaults10.extend_from_slice(&0x0200_0000u32.to_le_bytes());
        defaults10.extend_from_slice(&37u16.to_le_bytes());
        let defaults10 = TextDefaultsExtension10::parse(&defaults10).unwrap();
        assert_eq!(defaults10.character.complex_script_font_ref, Some(37));

        let mut trailing = defaults10.character.mask.to_le_bytes().to_vec();
        trailing.extend_from_slice(&37u16.to_le_bytes());
        trailing.push(0);
        assert!(TextDefaultsExtension10::parse(&trailing).is_err());
        assert!(TextDefaultsExtension9::parse(&[0; 7]).is_err());
    }

    #[test]
    fn discovers_typed_master_styles_and_defaults_in_versioned_tags() {
        let mut ppt9_blob = Vec::new();
        ppt9_blob.extend_from_slice(&ppt_record_bytes(0, 1, 4013, &0u16.to_le_bytes()));
        ppt9_blob.extend_from_slice(&ppt_record_bytes(0, 0, 4016, &[0; 8]));

        let mut ppt10_blob = Vec::new();
        ppt10_blob.extend_from_slice(&ppt_record_bytes(0, 5, 4018, &0u16.to_le_bytes()));
        ppt10_blob.extend_from_slice(&ppt_record_bytes(0, 0, 4020, &[0; 4]));

        let root = Record {
            record_type: crate::consts::RecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![
                prog_tags_record(9, &ppt9_blob),
                prog_tags_record(10, &ppt10_blob),
            ],
        };

        let styles = VersionedTextMasterStyles::parse(&root).unwrap();
        assert_eq!(styles.powerpoint9.len(), 1);
        assert_eq!(styles.powerpoint9[0].text_type, 1);
        assert_eq!(styles.powerpoint10.len(), 1);
        assert_eq!(styles.powerpoint10[0].text_type, 5);

        let defaults = VersionedTextDefaults::parse(&root).unwrap();
        assert_eq!(defaults.powerpoint9.unwrap().paragraph.mask, 0);
        assert_eq!(defaults.powerpoint10.unwrap().character.mask, 0);

        assert!(root.versioned_binary_tag_records(8).is_err());
        assert!(root.versioned_binary_tag_records(13).is_err());
    }
}
