//! XLSB rich shared-string parsing.

use crate::package::error::{Error, Result};
use crate::package::records::decode_string;
use litchi_core::binary;

/// A font change within an XLSB shared string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SharedStringRun {
    /// UTF-16 character index where this formatting run begins.
    pub character_index: u16,
    /// Index into the workbook font table.
    pub font_id: u16,
}

/// A mapping from phonetic text to a range in the base string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PhoneticRun {
    /// UTF-16 character index where this run begins in the phonetic text.
    pub phonetic_character_index: u16,
    /// UTF-16 character index where the corresponding base-text range begins.
    pub base_character_index: u16,
    /// Number of UTF-16 characters in the corresponding base-text range.
    pub base_character_count: u16,
}

/// The type of phonetic conversion stored with a shared string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhoneticType {
    /// Half-width Katakana.
    HalfWidthKatakana,
    /// Full-width Katakana.
    FullWidthKatakana,
    /// Hiragana.
    Hiragana,
    /// Preserve the phonetic text without conversion.
    NoConversion,
}

impl PhoneticType {
    fn from_bits(value: u16) -> Self {
        match value & 3 {
            0 => Self::HalfWidthKatakana,
            1 => Self::FullWidthKatakana,
            2 => Self::Hiragana,
            _ => Self::NoConversion,
        }
    }
}

/// Horizontal alignment of phonetic text over its base string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhoneticAlignment {
    /// Use the application default.
    NoControl,
    /// Align phonetic text to the left.
    Left,
    /// Center phonetic text.
    Center,
    /// Distribute phonetic text across the base range.
    Distributed,
}

impl PhoneticAlignment {
    fn from_bits(value: u16) -> Self {
        match value & 3 {
            0 => Self::NoControl,
            1 => Self::Left,
            2 => Self::Center,
            _ => Self::Distributed,
        }
    }
}

/// Optional pronunciation metadata stored with an XLSB shared string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PhoneticString {
    /// Phonetic text associated with the base shared string.
    pub text: String,
    /// Mappings from phonetic-text positions to base-text ranges.
    pub runs: Vec<PhoneticRun>,
    /// Index into the workbook font table used for phonetic text.
    pub font_id: u16,
    /// Phonetic conversion type encoded by `phType`.
    pub phonetic_type: PhoneticType,
    /// Horizontal alignment encoded by `alcH`.
    pub alignment: PhoneticAlignment,
}

/// An XLSB shared-string entry with optional formatting and pronunciation data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SharedString {
    /// Plain base text used as the cell value.
    pub text: String,
    /// Formatting runs, empty for an unformatted string.
    pub runs: Vec<SharedStringRun>,
    /// Optional phonetic text, ranges, and display settings.
    pub phonetic: Option<PhoneticString>,
}

impl SharedString {
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.is_empty() {
            return Err(Error::InvalidLength {
                expected: 1,
                found: 0,
            });
        }
        let rich = data[0] & 1 != 0;
        let extended = data[0] & 2 != 0;
        let (text, consumed) = decode_string(&data[1..])?;
        let text_len = text.encode_utf16().count();
        if text_len > 0x7FFF {
            return Err(Error::Unrecognized {
                typ: "RichStr text length".to_string(),
                val: text_len.to_string(),
            });
        }
        let mut offset = 1 + consumed;
        let mut runs = Vec::new();
        if rich {
            let count = Self::read_count(data, &mut offset, "StrRun")?;
            let byte_count = count
                .checked_mul(4)
                .ok_or_else(|| Error::Encoding("StrRun byte count overflow".to_string()))?;
            let end = offset
                .checked_add(byte_count)
                .ok_or_else(|| Error::Encoding("StrRun offset overflow".to_string()))?;
            if end > data.len() {
                return Err(Error::InvalidLength {
                    expected: end,
                    found: data.len(),
                });
            }
            runs.reserve(count);
            let mut previous = None;
            for chunk in data[offset..end].chunks_exact(4) {
                let character_index = binary::read_u16_le_at(chunk, 0)?;
                if usize::from(character_index) >= text_len
                    || previous.is_some_and(|value| character_index <= value)
                {
                    return Err(Error::Unrecognized {
                        typ: "StrRun ich".to_string(),
                        val: character_index.to_string(),
                    });
                }
                previous = Some(character_index);
                runs.push(SharedStringRun {
                    character_index,
                    font_id: binary::read_u16_le_at(chunk, 2)?,
                });
            }
            offset = end;
        }

        let phonetic = if extended {
            let (phonetic_text, consumed) = decode_string(&data[offset..])?;
            offset += consumed;
            let phonetic_len = phonetic_text.encode_utf16().count();
            let count = Self::read_count(data, &mut offset, "PhRun")?;
            let byte_count = count
                .checked_mul(6)
                .ok_or_else(|| Error::Encoding("PhRun byte count overflow".to_string()))?;
            let runs_end = offset
                .checked_add(byte_count)
                .ok_or_else(|| Error::Encoding("PhRun offset overflow".to_string()))?;
            let end = runs_end
                .checked_add(4)
                .ok_or_else(|| Error::Encoding("phonetic settings offset overflow".to_string()))?;
            if end > data.len() {
                return Err(Error::InvalidLength {
                    expected: end,
                    found: data.len(),
                });
            }
            let mut phonetic_runs = Vec::with_capacity(count);
            let mut previous_phonetic = None;
            let mut previous_base_end = None;
            for chunk in data[offset..runs_end].chunks_exact(6) {
                let phonetic_character_index = binary::read_u16_le_at(chunk, 0)?;
                let base_character_index = binary::read_u16_le_at(chunk, 2)?;
                let base_character_count = binary::read_u16_le_at(chunk, 4)?;
                let base_end = usize::from(base_character_index)
                    .checked_add(usize::from(base_character_count))
                    .ok_or_else(|| Error::Encoding("PhRun range overflow".to_string()))?;
                if usize::from(phonetic_character_index) >= phonetic_len
                    || usize::from(base_character_index) >= text_len
                    || base_end > text_len
                    || previous_phonetic.is_some_and(|value| phonetic_character_index <= value)
                    || previous_base_end
                        .is_some_and(|value| usize::from(base_character_index) < value)
                {
                    return Err(Error::Unrecognized {
                        typ: "PhRun index".to_string(),
                        val: format!("{phonetic_character_index}/{base_character_index}"),
                    });
                }
                previous_phonetic = Some(phonetic_character_index);
                previous_base_end = Some(base_end);
                phonetic_runs.push(PhoneticRun {
                    phonetic_character_index,
                    base_character_index,
                    base_character_count,
                });
            }
            let font_id = binary::read_u16_le_at(data, runs_end)?;
            let flags = binary::read_u16_le_at(data, runs_end + 2)?;
            offset = end;
            Some(PhoneticString {
                text: phonetic_text,
                runs: phonetic_runs,
                font_id,
                phonetic_type: PhoneticType::from_bits(flags),
                alignment: PhoneticAlignment::from_bits(flags >> 2),
            })
        } else {
            None
        };
        if offset != data.len() {
            return Err(Error::Unrecognized {
                typ: "RichStr".to_string(),
                val: format!("{} trailing bytes", data.len() - offset),
            });
        }
        Ok(Self {
            text,
            runs,
            phonetic,
        })
    }

    fn read_count(data: &[u8], offset: &mut usize, context: &str) -> Result<usize> {
        let end = offset
            .checked_add(4)
            .ok_or_else(|| Error::Encoding(format!("{context} count offset overflow")))?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let count = binary::read_u32_le_at(data, *offset)? as usize;
        *offset = end;
        if count > 0x7FFF {
            return Err(Error::Unrecognized {
                typ: format!("{context} count"),
                val: count.to_string(),
            });
        }
        Ok(count)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_rich_and_phonetic_shared_string() {
        let mut data = vec![3];
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&[b'A', 0, b'B', 0]);
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&[0, 0, 1, 0, 1, 0, 2, 0]);
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&[b'a', 0, b'b', 0]);
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&[0, 0, 0, 0, 2, 0]);
        data.extend_from_slice(&[3, 0, 6, 0]);

        let value = SharedString::parse(&data).unwrap();
        assert_eq!(value.text, "AB");
        assert_eq!(value.runs.len(), 2);
        assert_eq!(value.runs[1].font_id, 2);
        let phonetic = value.phonetic.unwrap();
        assert_eq!(phonetic.runs[0].base_character_count, 2);
        assert_eq!(phonetic.font_id, 3);
        assert_eq!(phonetic.alignment, PhoneticAlignment::Left);
    }
}
