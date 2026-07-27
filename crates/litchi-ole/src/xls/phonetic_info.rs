//! BIFF8 `PhoneticInfo` record (MS-XLS 2.4.192): default phonetic-string
//! format and visible phonetic cell ranges.
//!
//! The record may span `Continue` records when the `SqRef` range list is
//! long; continuations are concatenated before parsing.

use super::{XlsError, XlsResult};

/// Record type of the `PhoneticInfo` record.
pub(crate) const PHONETIC_INFO_RECORD_TYPE: u16 = 0x00EF;

/// Largest legal range count (`SqRef.cref`).
const MAX_RANGES: usize = 0x2000;
/// Size in bytes of the `Phs` structure plus the `SqRef` count.
const HEADER_LEN: usize = 6;
/// Size in bytes of one `Ref8` range.
const REF8_LEN: usize = 6;

// Phs flag bits.
const PHONETIC_TYPE_MASK: u16 = 0x0003;
const ALIGNMENT_SHIFT: u16 = 2;
const ALIGNMENT_MASK: u16 = 0x0003;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: PHONETIC_INFO_RECORD_TYPE,
        message: message.into(),
    }
}

/// The phonetic character type of a `Phs` structure (`phType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPhoneticType {
    /// Narrow (half-width) Katakana characters.
    NarrowKatakana,
    /// Wide (full-width) Katakana characters.
    WideKatakana,
    /// Hiragana characters.
    Hiragana,
    /// Any type of characters.
    Any,
}

impl XlsPhoneticType {
    fn from_code(value: u16) -> Self {
        match value & PHONETIC_TYPE_MASK {
            0 => Self::NarrowKatakana,
            1 => Self::WideKatakana,
            2 => Self::Hiragana,
            _ => Self::Any,
        }
    }

    fn code(self) -> u16 {
        match self {
            Self::NarrowKatakana => 0,
            Self::WideKatakana => 1,
            Self::Hiragana => 2,
            Self::Any => 3,
        }
    }
}

/// The phonetic-string alignment of a `Phs` structure (`alcH`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPhoneticAlignment {
    /// General alignment.
    General,
    /// Left aligned.
    Left,
    /// Center aligned.
    Center,
    /// Distributed alignment.
    Distributed,
}

impl XlsPhoneticAlignment {
    fn from_code(value: u16) -> Self {
        match (value >> ALIGNMENT_SHIFT) & ALIGNMENT_MASK {
            0 => Self::General,
            1 => Self::Left,
            2 => Self::Center,
            _ => Self::Distributed,
        }
    }

    fn code(self) -> u16 {
        match self {
            Self::General => 0,
            Self::Left => 1,
            Self::Center => 2,
            Self::Distributed => 3,
        }
    }
}

/// The default phonetic-string format of a sheet (`Phs`, MS-XLS 2.5.201).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPhoneticFormat {
    font_index: u16,
    phonetic_type: XlsPhoneticType,
    alignment: XlsPhoneticAlignment,
    /// Raw unused bits of the `Phs` bitfield, preserved verbatim.
    unused_flags: u16,
}

impl XlsPhoneticFormat {
    /// A phonetic format with no unused flag bits.
    pub fn new(
        font_index: u16,
        phonetic_type: XlsPhoneticType,
        alignment: XlsPhoneticAlignment,
    ) -> Self {
        Self { font_index, phonetic_type, alignment, unused_flags: 0 }
    }

    pub const fn font_index(&self) -> u16 {
        self.font_index
    }
    pub const fn phonetic_type(&self) -> XlsPhoneticType {
        self.phonetic_type
    }
    pub const fn alignment(&self) -> XlsPhoneticAlignment {
        self.alignment
    }
}

/// One cell range with visible phonetic strings (`Ref8`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPhoneticRange {
    first_row: u16,
    last_row: u16,
    first_col: u8,
    last_col: u8,
}

impl XlsPhoneticRange {
    /// A range; the last row and column must not precede the first.
    pub fn new(first_row: u16, last_row: u16, first_col: u8, last_col: u8) -> XlsResult<Self> {
        if last_row < first_row || last_col < first_col {
            return Err(invalid("phonetic range is reversed"));
        }
        Ok(Self { first_row, last_row, first_col, last_col })
    }

    pub const fn first_row(&self) -> u16 {
        self.first_row
    }
    pub const fn last_row(&self) -> u16 {
        self.last_row
    }
    pub const fn first_col(&self) -> u8 {
        self.first_col
    }
    pub const fn last_col(&self) -> u8 {
        self.last_col
    }
}

/// Typed `PhoneticInfo` record content (MS-XLS 2.4.192).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPhoneticInfo {
    format: XlsPhoneticFormat,
    ranges: Vec<XlsPhoneticRange>,
}

impl XlsPhoneticInfo {
    /// A record with the default format and the visible phonetic ranges.
    pub fn try_new(format: XlsPhoneticFormat, ranges: Vec<XlsPhoneticRange>) -> XlsResult<Self> {
        if ranges.len() > MAX_RANGES {
            return Err(invalid("phonetic range count exceeds 0x2000"));
        }
        Ok(Self { format, ranges })
    }

    pub const fn format(&self) -> XlsPhoneticFormat {
        self.format
    }
    pub fn ranges(&self) -> &[XlsPhoneticRange] {
        &self.ranges
    }

    /// Parse the concatenated payloads of a `PhoneticInfo` record and its
    /// `Continue` records.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < HEADER_LEN {
            return Err(XlsError::InvalidLength {
                expected: HEADER_LEN,
                found: data.len(),
            });
        }
        let flags = u16::from_le_bytes([data[2], data[3]]);
        let range_count = usize::from(u16::from_le_bytes([data[4], data[5]]));
        if range_count > MAX_RANGES {
            return Err(invalid("phonetic range count exceeds 0x2000"));
        }
        let expected = HEADER_LEN
            .checked_add(range_count.checked_mul(REF8_LEN).ok_or_else(|| invalid("range overflow"))?)
            .ok_or_else(|| invalid("range overflow"))?;
        if data.len() != expected {
            return Err(XlsError::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        let mut ranges = Vec::with_capacity(range_count);
        for chunk in data[HEADER_LEN..].chunks_exact(REF8_LEN) {
            ranges.push(XlsPhoneticRange::new(
                u16::from_le_bytes([chunk[0], chunk[1]]),
                u16::from_le_bytes([chunk[2], chunk[3]]),
                chunk[4],
                chunk[5],
            )?);
        }
        Ok(Self {
            format: XlsPhoneticFormat {
                font_index: u16::from_le_bytes([data[0], data[1]]),
                phonetic_type: XlsPhoneticType::from_code(flags),
                alignment: XlsPhoneticAlignment::from_code(flags),
                unused_flags: flags & !0x000F,
            },
            ranges,
        })
    }

    /// Serialize back to the concatenated record payload.
    pub(crate) fn to_payload(&self) -> Vec<u8> {
        let flags = self.format.phonetic_type.code()
            | (self.format.alignment.code() << ALIGNMENT_SHIFT)
            | self.format.unused_flags;
        let mut payload = Vec::with_capacity(HEADER_LEN + self.ranges.len() * REF8_LEN);
        payload.extend_from_slice(&self.format.font_index.to_le_bytes());
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&(self.ranges.len() as u16).to_le_bytes());
        for range in &self.ranges {
            payload.extend_from_slice(&range.first_row.to_le_bytes());
            payload.extend_from_slice(&range.last_row.to_le_bytes());
            payload.push(range.first_col);
            payload.push(range.last_col);
        }
        payload
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> XlsPhoneticInfo {
        XlsPhoneticInfo::try_new(
            XlsPhoneticFormat::new(4, XlsPhoneticType::Hiragana, XlsPhoneticAlignment::Center),
            vec![
                XlsPhoneticRange::new(1, 3, 0, 5).unwrap(),
                XlsPhoneticRange::new(7, 7, 2, 2).unwrap(),
            ],
        )
        .unwrap()
    }

    #[test]
    fn round_trips() {
        let payload = sample().to_payload();
        let parsed = XlsPhoneticInfo::parse(&payload).unwrap();
        assert_eq!(parsed, sample());
        assert_eq!(parsed.format().font_index(), 4);
        assert_eq!(parsed.format().phonetic_type(), XlsPhoneticType::Hiragana);
        assert_eq!(parsed.format().alignment(), XlsPhoneticAlignment::Center);
        assert_eq!(parsed.ranges().len(), 2);
        assert_eq!(parsed.ranges()[1].last_col(), 2);
    }

    #[test]
    fn preserves_unused_flag_bits() {
        let mut payload = sample().to_payload();
        payload[2] |= 0xF0;
        let parsed = XlsPhoneticInfo::parse(&payload).unwrap();
        assert_eq!(parsed.format().unused_flags & 0xF0, 0xF0);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated header.
        assert!(XlsPhoneticInfo::parse(&[0; 5]).is_err());
        // Count/payload mismatch.
        let mut payload = sample().to_payload();
        payload.truncate(HEADER_LEN + REF8_LEN);
        assert!(XlsPhoneticInfo::parse(&payload).is_err());
        // Reversed range.
        assert!(XlsPhoneticRange::new(5, 2, 0, 1).is_err());
        // Count cap.
        assert!(XlsPhoneticInfo::try_new(
            XlsPhoneticFormat::new(0, XlsPhoneticType::Any, XlsPhoneticAlignment::General),
            vec![XlsPhoneticRange::new(0, 0, 0, 0).unwrap(); MAX_RANGES + 1],
        )
        .is_err());
    }
}
