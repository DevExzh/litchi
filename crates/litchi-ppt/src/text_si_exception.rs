//! `TextSIExceptionAtom` default special information (MS-PPT 2.9.31/2.9.32).
//!
//! The atom carries default language and spelling settings for the text in a
//! slide, notes, or master. It is inert: languages are never applied to any
//! spell-checking or formatting behavior.

use super::records::record::PptRecord;
use crate::consts::PptRecordType;
use crate::package::{PptError, Result};

/// `RT_TextSpecialInfoDefaultAtom` record type.
const TEXT_SPECIAL_INFO_DEFAULT_TYPE: u16 = 0x0FA9;

// TextSIException mask bits (MS-PPT 2.9.32).
const MASK_SPELL: u32 = 0x0001;
const MASK_LANG: u32 = 0x0002;
const MASK_ALT_LANG: u32 = 0x0004;
/// Bits that must be zero in a `TextSIExceptionAtom`: `fPp10ext`, `fBidi`,
/// `smartTag`, `reserved1`, and `reserved2`.
const MASK_FORBIDDEN: u32 = 0xFFFF_FF78;

/// `SpellingFlags.grammar`: forbidden in a `TextSIExceptionAtom`.
const SPELL_GRAMMAR: u16 = 0x0004;
/// `SpellingFlags.error`.
const SPELL_ERROR: u16 = 0x0001;
/// `SpellingFlags.clean`.
const SPELL_CLEAN: u16 = 0x0002;

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

/// Spelling status defaults from a `SpellingFlags` structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointSpellingFlags {
    /// Whether the text is spelled incorrectly.
    error: bool,
    /// Whether the text needs rechecking.
    clean: bool,
}

impl PowerPointSpellingFlags {
    /// Build spelling flags from validated bits (`TextSIException`,
    /// MS-PPT 2.9.33).
    pub(crate) const fn from_bits(error: bool, clean: bool) -> Self {
        Self { error, clean }
    }
    /// Whether the text is spelled incorrectly.
    pub const fn error(&self) -> bool {
        self.error
    }
    /// Whether the text needs rechecking.
    pub const fn clean(&self) -> bool {
        self.clean
    }
}

/// Default special information for a text body (`TextSIExceptionAtom`,
/// MS-PPT 2.9.31).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointTextSpecialInfoDefaults {
    spelling: Option<PowerPointSpellingFlags>,
    /// Primary language identifier (`TxLCID`).
    language: Option<u16>,
    /// Alternate language identifier (`TxLCID`).
    alternate_language: Option<u16>,
}

impl PowerPointTextSpecialInfoDefaults {
    /// Spelling status defaults, when present.
    pub const fn spelling(&self) -> Option<PowerPointSpellingFlags> {
        self.spelling
    }
    /// Primary language identifier, when present.
    pub const fn language(&self) -> Option<u16> {
        self.language
    }
    /// Alternate language identifier, when present.
    pub const fn alternate_language(&self) -> Option<u16> {
        self.alternate_language
    }

    /// Parse a complete `TextSIExceptionAtom` record.
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::TextSpecialInfoDefaultAtom
            || record.record_type_raw != TEXT_SPECIAL_INFO_DEFAULT_TYPE
            || record.version != 0
        {
            return Err(corrupted("TextSIExceptionAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the `TextSIException` payload.
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(corrupted("TextSIException mask is truncated"));
        }
        let mask = u32::from_le_bytes(data[..4].try_into().expect("length checked"));
        // TextSIExceptionAtom forbids fPp10ext, fBidi, smartTag, reserved1,
        // and reserved2 (MS-PPT 2.9.31).
        if mask & MASK_FORBIDDEN != 0 {
            return Err(corrupted(
                "TextSIExceptionAtom has forbidden special-info mask bits",
            ));
        }
        let mut offset = 4usize;
        let mut take_u16 = |mask_bit: u32, name: &str| -> Result<Option<u16>> {
            if mask & mask_bit == 0 {
                return Ok(None);
            }
            let Some(bytes) = data.get(offset..offset + 2) else {
                return Err(corrupted(format!("TextSIException {name} is truncated")));
            };
            offset += 2;
            Ok(Some(u16::from_le_bytes(
                bytes.try_into().expect("length checked"),
            )))
        };
        let spelling = take_u16(MASK_SPELL, "spellInfo")?
            .map(|raw| {
                if raw & SPELL_GRAMMAR != 0 {
                    return Err(corrupted(
                        "TextSIExceptionAtom spellInfo.grammar must be zero",
                    ));
                }
                Ok(PowerPointSpellingFlags {
                    error: raw & SPELL_ERROR != 0,
                    clean: raw & SPELL_CLEAN != 0,
                })
            })
            .transpose()?;
        let language = take_u16(MASK_LANG, "lid")?;
        let alternate_language = take_u16(MASK_ALT_LANG, "altLid")?;
        if offset != data.len() {
            return Err(corrupted(
                "TextSIException mask does not consume its payload exactly",
            ));
        }
        Ok(Self {
            spelling,
            language,
            alternate_language,
        })
    }

    /// Serialize the complete record, including its header.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut mask = 0u32;
        if self.spelling.is_some() {
            mask |= MASK_SPELL;
        }
        if self.language.is_some() {
            mask |= MASK_LANG;
        }
        if self.alternate_language.is_some() {
            mask |= MASK_ALT_LANG;
        }
        let mut data = Vec::with_capacity(14);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&TEXT_SPECIAL_INFO_DEFAULT_TYPE.to_le_bytes());
        data.extend_from_slice(&(4 + 2 * mask.count_ones()).to_le_bytes());
        data.extend_from_slice(&mask.to_le_bytes());
        if let Some(spelling) = self.spelling {
            let raw = u16::from(spelling.error) | (u16::from(spelling.clean) << 1);
            data.extend_from_slice(&raw.to_le_bytes());
        }
        if let Some(language) = self.language {
            data.extend_from_slice(&language.to_le_bytes());
        }
        if let Some(alternate_language) = self.alternate_language {
            data.extend_from_slice(&alternate_language.to_le_bytes());
        }
        Ok(data)
    }
}

/// A validated index into the sequence of `TextHeaderAtom` records that
/// follows a slide's persist record (`OutlineTextRefAtom`, MS-PPT 2.9.78).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PowerPointOutlineTextRef(u32);

impl PowerPointOutlineTextRef {
    /// A validated non-negative outline text reference index.
    pub fn new(index: i32) -> Result<Self> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_| corrupted("OutlineTextRefAtom index is negative"))
    }

    /// The zero-based outline text index.
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Parse a complete `OutlineTextRefAtom` record.
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::OutlineTextRefAtom
            || record.record_type_raw != 0x0F9E
            || record.version != 0
            || record.data.len() != 4
        {
            return Err(corrupted(
                "OutlineTextRefAtom has an invalid header or length",
            ));
        }
        Self::new(i32::from_le_bytes(
            record.data[..4].try_into().expect("length checked"),
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn atom(data: &[u8]) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::TextSpecialInfoDefaultAtom,
            record_type_raw: TEXT_SPECIAL_INFO_DEFAULT_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    #[test]
    fn parses_defaults_and_round_trips() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0007u32.to_le_bytes());
        data.extend_from_slice(&0x0002u16.to_le_bytes()); // clean
        data.extend_from_slice(&0x0409u16.to_le_bytes()); // en-US
        data.extend_from_slice(&0x0809u16.to_le_bytes()); // en-GB
        let defaults = PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).unwrap();
        assert!(defaults.spelling().unwrap().clean());
        assert!(!defaults.spelling().unwrap().error());
        assert_eq!(defaults.language(), Some(0x0409));
        assert_eq!(defaults.alternate_language(), Some(0x0809));
        assert_eq!(defaults.to_bytes().unwrap()[8..], data[..]);
    }

    #[test]
    fn parses_empty_and_language_only_defaults() {
        let data = 0u32.to_le_bytes();
        let defaults = PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).unwrap();
        assert_eq!(defaults.spelling(), None);
        assert_eq!(defaults.language(), None);
        assert_eq!(defaults.to_bytes().unwrap()[8..], data[..]);

        let mut data = Vec::new();
        data.extend_from_slice(&0x0002u32.to_le_bytes());
        data.extend_from_slice(&0x0002u16.to_le_bytes());
        let defaults = PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).unwrap();
        assert_eq!(defaults.spelling(), None);
        assert_eq!(defaults.language(), Some(0x0002));
        assert_eq!(defaults.to_bytes().unwrap()[8..], data[..]);
    }

    #[test]
    fn rejects_malformed_defaults() {
        // Truncated mask.
        assert!(PowerPointTextSpecialInfoDefaults::parse_record(&atom(&[1, 0])).is_err());
        // Forbidden mask bits (fBidi / smartTag / reserved).
        let mut data = Vec::new();
        data.extend_from_slice(&0x0042u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        assert!(PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).is_err());
        // Grammar flag set.
        let mut data = Vec::new();
        data.extend_from_slice(&0x0001u32.to_le_bytes());
        data.extend_from_slice(&0x0004u16.to_le_bytes());
        assert!(PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).is_err());
        // Mask count not consuming the payload.
        let mut data = Vec::new();
        data.extend_from_slice(&0x0002u32.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        assert!(PowerPointTextSpecialInfoDefaults::parse_record(&atom(&data)).is_err());
    }

    #[test]
    fn outline_text_ref_validates_and_parses() {
        let record = PptRecord {
            record_type: PptRecordType::OutlineTextRefAtom,
            record_type_raw: 0x0F9E,
            version: 0,
            instance: 0,
            data_length: 4,
            data: 5i32.to_le_bytes().to_vec(),
            children: Vec::new(),
        };
        assert_eq!(
            PowerPointOutlineTextRef::parse_record(&record)
                .unwrap()
                .get(),
            5
        );

        let mut negative = record.clone();
        negative.data = (-1i32).to_le_bytes().to_vec();
        assert!(PowerPointOutlineTextRef::parse_record(&negative).is_err());

        let mut long = record.clone();
        long.data.extend_from_slice(&[0; 4]);
        assert!(PowerPointOutlineTextRef::parse_record(&long).is_err());
    }
}
