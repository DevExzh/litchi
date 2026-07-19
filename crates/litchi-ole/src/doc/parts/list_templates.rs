//! Word `SttbRgtplc` table for list-level template codes.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use std::collections::HashSet;

const FIB_INDEX: usize = 96;
const MAX_ENTRY_COUNT: usize = 0x7FF0;
const LEVEL_COUNT: usize = 9;
const ENTRY_CODE_UNITS: u16 = 0x12;
const ENTRY_BYTES: usize = LEVEL_COUNT * 4;
const MAX_TABLE_BYTES: usize = 6 + MAX_ENTRY_COUNT * (2 + ENTRY_BYTES);

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Application-defined language identifier stored in a built-in `Tplc`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ListTemplateLanguageId(u16);

impl ListTemplateLanguageId {
    pub const fn new(raw: u16) -> Self {
        Self(raw)
    }

    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// Defined `TplcBuildIn.ilgpdM1` formats from [MS-DOC] 2.9.329.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum BuiltInListTemplate {
    BlackBullet = 0x0000,
    WhiteBullet = 0x0001,
    BlackSquare = 0x0002,
    WhiteSquare = 0x0003,
    Diamond = 0x0004,
    ArrowHead = 0x0005,
    Arrow = 0x0006,
    ArabicPeriod = 0x0007,
    ArabicParenthesis = 0x0008,
    UpperRomanPeriod = 0x0009,
    UpperLetterPeriod = 0x000A,
    LowerLetterParenthesis = 0x000B,
    LowerLetterPeriod = 0x000C,
    LowerRomanPeriod = 0x000D,
    None = 0x7FFF,
}

impl TryFrom<u16> for BuiltInListTemplate {
    type Error = u16;

    fn try_from(value: u16) -> std::result::Result<Self, Self::Error> {
        Ok(match value {
            0x0000 => Self::BlackBullet,
            0x0001 => Self::WhiteBullet,
            0x0002 => Self::BlackSquare,
            0x0003 => Self::WhiteSquare,
            0x0004 => Self::Diamond,
            0x0005 => Self::ArrowHead,
            0x0006 => Self::Arrow,
            0x0007 => Self::ArabicPeriod,
            0x0008 => Self::ArabicParenthesis,
            0x0009 => Self::UpperRomanPeriod,
            0x000A => Self::UpperLetterPeriod,
            0x000B => Self::LowerLetterParenthesis,
            0x000C => Self::LowerLetterPeriod,
            0x000D => Self::LowerRomanPeriod,
            0x7FFF => Self::None,
            invalid => return Err(invalid),
        })
    }
}

/// One strongly typed 32-bit `Tplc` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ListTemplateCode {
    BuiltIn {
        format: BuiltInListTemplate,
        language: ListTemplateLanguageId,
    },
    UserDefined {
        random_id: u32,
    },
}

impl ListTemplateCode {
    pub fn user_defined(random_id: u32) -> Result<Self> {
        if random_id > 0x7FFF_FFFF {
            return Err(corrupted("TplcUser random identifier exceeds 31 bits"));
        }
        Ok(Self::UserDefined { random_id })
    }

    pub fn from_raw(raw: u32) -> Result<Self> {
        if raw & 1 == 0 {
            return Self::user_defined(raw >> 1);
        }
        let format = BuiltInListTemplate::try_from(((raw >> 1) & 0x7FFF) as u16)
            .map_err(|invalid| corrupted(format!("invalid TplcBuildIn format {invalid:#06x}")))?;
        Ok(Self::BuiltIn {
            format,
            language: ListTemplateLanguageId::new((raw >> 16) as u16),
        })
    }

    pub const fn raw(self) -> u32 {
        match self {
            Self::BuiltIn { format, language } => {
                1 | ((format as u32) << 1) | ((language.raw() as u32) << 16)
            },
            Self::UserDefined { random_id } => random_id << 1,
        }
    }
}

/// One optional nine-level entry parallel to an `LSTF` definition.
pub type ListTemplateEntry = Option<[ListTemplateCode; LEVEL_COUNT]>;

/// Ordered `SttbRgtplc` entries parallel to `PlfLst.rgLstf`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ListTemplateTable {
    entries: Vec<ListTemplateEntry>,
}

impl ListTemplateTable {
    pub fn try_new(entries: Vec<ListTemplateEntry>) -> Result<Self> {
        validate_entries(&entries)?;
        Ok(Self { entries })
    }

    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start =
            usize::try_from(offset).map_err(|_| corrupted("SttbRgtplc offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("SttbRgtplc length is too large"))?;
        if length > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbRgtplc exceeds its specification-derived size cap",
            ));
        }
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("SttbRgtplc range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("SttbRgtplc extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 6
            || data.len() > MAX_TABLE_BYTES
            || read_u16(data, 0, "SttbRgtplc fExtend")? != 0xFFFF
            || read_u16(data, 4, "SttbRgtplc cbExtra")? != 0
        {
            return Err(corrupted("SttbRgtplc has an invalid header"));
        }
        let count = usize::from(read_u16(data, 2, "SttbRgtplc cData")?);
        if count > MAX_ENTRY_COUNT {
            return Err(corrupted("SttbRgtplc contains more than 0x7FF0 entries"));
        }
        let mut entries = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let unit_count = read_u16(data, offset, &format!("SttbRgtplc entry {index} length"))?;
            offset = offset
                .checked_add(2)
                .ok_or_else(|| corrupted("SttbRgtplc entry offset overflows"))?;
            match unit_count {
                0 => entries.push(None),
                ENTRY_CODE_UNITS => {
                    let end = offset
                        .checked_add(ENTRY_BYTES)
                        .ok_or_else(|| corrupted("SttbRgtplc entry range overflows"))?;
                    if end > data.len() {
                        return Err(corrupted(format!("SttbRgtplc entry {index} is truncated")));
                    }
                    let mut codes = [ListTemplateCode::UserDefined { random_id: 0 }; LEVEL_COUNT];
                    for (level, code) in codes.iter_mut().enumerate() {
                        *code = ListTemplateCode::from_raw(read_u32(
                            data,
                            offset + level * 4,
                            &format!("SttbRgtplc entry {index} level {level}"),
                        )?)?;
                    }
                    entries.push(Some(codes));
                    offset = end;
                },
                invalid => {
                    return Err(corrupted(format!(
                        "SttbRgtplc entry {index} has invalid cchData {invalid:#06x}"
                    )));
                },
            }
        }
        if offset != data.len() {
            return Err(corrupted("SttbRgtplc has trailing bytes"));
        }
        Self::try_new(entries)
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn entries(&self) -> &[ListTemplateEntry] {
        &self.entries
    }

    pub fn get(&self, definition: usize) -> Option<&[ListTemplateCode; LEVEL_COUNT]> {
        self.entries.get(definition)?.as_ref()
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_entries(&self.entries)?;
        let mut data = Vec::with_capacity(
            6 + self.entries.len() * 2
                + self.entries.iter().filter(|entry| entry.is_some()).count() * ENTRY_BYTES,
        );
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(self.entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for entry in &self.entries {
            if let Some(codes) = entry {
                data.extend_from_slice(&ENTRY_CODE_UNITS.to_le_bytes());
                for code in codes {
                    data.extend_from_slice(&code.raw().to_le_bytes());
                }
            } else {
                data.extend_from_slice(&0u16.to_le_bytes());
            }
        }
        Ok(data)
    }
}

fn validate_entries(entries: &[ListTemplateEntry]) -> Result<()> {
    if entries.len() > MAX_ENTRY_COUNT {
        return Err(corrupted("SttbRgtplc contains more than 0x7FF0 entries"));
    }
    let mut user_ids = HashSet::new();
    for code in entries.iter().flatten().flatten() {
        if let ListTemplateCode::UserDefined { random_id } = code {
            if *random_id > 0x7FFF_FFFF {
                return Err(corrupted("TplcUser random identifier exceeds 31 bits"));
            }
            if !user_ids.insert(*random_id) {
                return Err(corrupted(format!(
                    "SttbRgtplc repeats user-defined template identifier {random_id}"
                )));
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn built_in(format: BuiltInListTemplate, language: u16) -> ListTemplateCode {
        ListTemplateCode::BuiltIn {
            format,
            language: ListTemplateLanguageId::new(language),
        }
    }

    #[test]
    fn tplc_codes_round_trip_and_reject_unknown_built_ins() {
        for code in [
            built_in(BuiltInListTemplate::BlackBullet, 0x0409),
            built_in(BuiltInListTemplate::None, 0),
            ListTemplateCode::user_defined(0x1234_5678).unwrap(),
        ] {
            assert_eq!(ListTemplateCode::from_raw(code.raw()).unwrap(), code);
        }
        assert!(ListTemplateCode::from_raw(1 | (0x1234 << 1)).is_err());
        assert!(ListTemplateCode::user_defined(0x8000_0000).is_err());
    }

    #[test]
    fn table_round_trips_empty_and_complete_parallel_slots() {
        let codes = [
            built_in(BuiltInListTemplate::ArabicPeriod, 0x0409),
            built_in(BuiltInListTemplate::ArabicParenthesis, 0x0409),
            built_in(BuiltInListTemplate::UpperRomanPeriod, 0x0409),
            built_in(BuiltInListTemplate::UpperLetterPeriod, 0x0409),
            built_in(BuiltInListTemplate::LowerLetterParenthesis, 0x0409),
            built_in(BuiltInListTemplate::LowerLetterPeriod, 0x0409),
            built_in(BuiltInListTemplate::LowerRomanPeriod, 0x0409),
            built_in(BuiltInListTemplate::Arrow, 0x0409),
            built_in(BuiltInListTemplate::None, 0),
        ];
        let table = ListTemplateTable::try_new(vec![Some(codes), None]).unwrap();
        let bytes = table.to_bytes().unwrap();
        assert_eq!(ListTemplateTable::parse_bytes(&bytes).unwrap(), table);
        assert_eq!(table.get(0), Some(&codes));
        assert!(table.get(1).is_none());
    }

    #[test]
    fn rejects_invalid_lengths_trailing_bytes_and_duplicate_user_ids() {
        let mut invalid = vec![0xFF, 0xFF, 1, 0, 0, 0, 1, 0];
        assert!(ListTemplateTable::parse_bytes(&invalid).is_err());
        invalid[6..8].copy_from_slice(&ENTRY_CODE_UNITS.to_le_bytes());
        assert!(ListTemplateTable::parse_bytes(&invalid).is_err());
        let user = ListTemplateCode::user_defined(7).unwrap();
        assert!(ListTemplateTable::try_new(vec![Some([user; LEVEL_COUNT])]).is_err());
        let mut empty = ListTemplateTable::default().to_bytes().unwrap();
        empty.push(0);
        assert!(ListTemplateTable::parse_bytes(&empty).is_err());
    }
}
