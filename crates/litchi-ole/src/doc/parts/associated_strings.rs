//! Word `SttbfAssoc` associated-document string table.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

const FIB_INDEX: usize = 32;
const STRING_COUNT: usize = 18;
const MAX_GENERAL_UNITS: usize = 255;
const MAX_PASSWORD_UNITS: usize = 15;
const MAX_TABLE_BYTES: usize = 6 + 17 * (2 + MAX_GENERAL_UNITS * 2) + (2 + MAX_PASSWORD_UNITS * 2);

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Fixed slot indexes in an `SttbfAssoc` table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum AssociatedStringSlot {
    Unused0 = 0,
    TemplatePath = 1,
    Title = 2,
    Subject = 3,
    Keywords = 4,
    Unused5 = 5,
    Author = 6,
    LastRevisedBy = 7,
    MailMergeDataSourcePath = 8,
    MailMergeHeaderPath = 9,
    Unused10 = 10,
    Unused11 = 11,
    Unused12 = 12,
    Unused13 = 13,
    Unused14 = 14,
    Unused15 = 15,
    Unused16 = 16,
    WriteReservationPassword = 17,
}

impl AssociatedStringSlot {
    pub const ALL: [Self; STRING_COUNT] = [
        Self::Unused0,
        Self::TemplatePath,
        Self::Title,
        Self::Subject,
        Self::Keywords,
        Self::Unused5,
        Self::Author,
        Self::LastRevisedBy,
        Self::MailMergeDataSourcePath,
        Self::MailMergeHeaderPath,
        Self::Unused10,
        Self::Unused11,
        Self::Unused12,
        Self::Unused13,
        Self::Unused14,
        Self::Unused15,
        Self::Unused16,
        Self::WriteReservationPassword,
    ];

    pub const fn index(self) -> usize {
        self as usize
    }
}

/// The 18 associated strings stored with a Word document.
///
/// Path values are inert metadata. Parsing this table never opens or resolves them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentAssociatedStrings {
    values: [String; STRING_COUNT],
}

impl Default for DocumentAssociatedStrings {
    fn default() -> Self {
        Self {
            values: std::array::from_fn(|_| String::new()),
        }
    }
}

impl DocumentAssociatedStrings {
    pub fn try_new(values: [String; STRING_COUNT]) -> Result<Self> {
        validate_values(&values)?;
        Ok(Self { values })
    }

    /// Parse the optional FIB index-32 table range.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start =
            usize::try_from(offset).map_err(|_| corrupted("SttbfAssoc offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("SttbfAssoc length is too large"))?;
        if length > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbfAssoc exceeds its specification-derived size cap",
            ));
        }
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("SttbfAssoc range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("SttbfAssoc extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `SttbfAssoc` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbfAssoc exceeds its specification-derived size cap",
            ));
        }
        if data.len() < 6
            || read_u16(data, 0, "SttbfAssoc fExtend")? != 0xFFFF
            || usize::from(read_u16(data, 2, "SttbfAssoc cData")?) != STRING_COUNT
            || read_u16(data, 4, "SttbfAssoc cbExtra")? != 0
        {
            return Err(corrupted("SttbfAssoc has an invalid header"));
        }

        let mut values = Vec::with_capacity(STRING_COUNT);
        let mut offset = 6usize;
        for index in 0..STRING_COUNT {
            let unit_count = usize::from(read_u16(
                data,
                offset,
                &format!("SttbfAssoc string {index} length"),
            )?);
            let maximum = if index == AssociatedStringSlot::WriteReservationPassword.index() {
                MAX_PASSWORD_UNITS
            } else {
                MAX_GENERAL_UNITS
            };
            if unit_count > maximum {
                return Err(corrupted(format!(
                    "SttbfAssoc string {index} exceeds {maximum} UTF-16 code units"
                )));
            }
            offset = offset
                .checked_add(2)
                .ok_or_else(|| corrupted("SttbfAssoc string offset overflows"))?;
            let byte_count = unit_count
                .checked_mul(2)
                .ok_or_else(|| corrupted("SttbfAssoc string length overflows"))?;
            let end = offset
                .checked_add(byte_count)
                .ok_or_else(|| corrupted("SttbfAssoc string range overflows"))?;
            let bytes = data
                .get(offset..end)
                .ok_or_else(|| corrupted(format!("SttbfAssoc string {index} is truncated")))?;
            let units = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            values.push(String::from_utf16(&units).map_err(|_| {
                corrupted(format!("SttbfAssoc string {index} contains invalid UTF-16"))
            })?);
            offset = end;
        }
        if offset != data.len() {
            return Err(corrupted("SttbfAssoc has trailing bytes"));
        }
        let values = values
            .try_into()
            .map_err(|_| corrupted("SttbfAssoc did not contain exactly 18 strings"))?;
        Ok(Self { values })
    }

    pub fn get(&self, slot: AssociatedStringSlot) -> &str {
        &self.values[slot.index()]
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = (AssociatedStringSlot, &str)> {
        AssociatedStringSlot::ALL
            .into_iter()
            .map(|slot| (slot, self.get(slot)))
    }

    /// Associated template path as inert metadata.
    pub fn template_path(&self) -> &str {
        self.get(AssociatedStringSlot::TemplatePath)
    }
    pub fn title(&self) -> &str {
        self.get(AssociatedStringSlot::Title)
    }
    pub fn subject(&self) -> &str {
        self.get(AssociatedStringSlot::Subject)
    }
    pub fn keywords(&self) -> &str {
        self.get(AssociatedStringSlot::Keywords)
    }
    pub fn author(&self) -> &str {
        self.get(AssociatedStringSlot::Author)
    }
    pub fn last_revised_by(&self) -> &str {
        self.get(AssociatedStringSlot::LastRevisedBy)
    }
    /// Associated mail-merge data source path as inert metadata.
    pub fn mail_merge_data_source_path(&self) -> &str {
        self.get(AssociatedStringSlot::MailMergeDataSourcePath)
    }
    /// Associated mail-merge header path as inert metadata.
    pub fn mail_merge_header_path(&self) -> &str {
        self.get(AssociatedStringSlot::MailMergeHeaderPath)
    }
    pub fn write_reservation_password(&self) -> &str {
        self.get(AssociatedStringSlot::WriteReservationPassword)
    }

    /// Serialize deterministically, preserving every unused slot.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let size = validate_values(&self.values)?;
        let mut data = Vec::with_capacity(size);
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(STRING_COUNT as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for value in &self.values {
            let units = value.encode_utf16().collect::<Vec<_>>();
            let count = u16::try_from(units.len())
                .map_err(|_| corrupted("SttbfAssoc string length exceeds u16"))?;
            data.extend_from_slice(&count.to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        Ok(data)
    }
}

fn validate_values(values: &[String; STRING_COUNT]) -> Result<usize> {
    let mut size = 6usize;
    for (index, value) in values.iter().enumerate() {
        let units = value.encode_utf16().count();
        let maximum = if index == AssociatedStringSlot::WriteReservationPassword.index() {
            MAX_PASSWORD_UNITS
        } else {
            MAX_GENERAL_UNITS
        };
        if units > maximum {
            return Err(corrupted(format!(
                "SttbfAssoc string {index} exceeds {maximum} UTF-16 code units"
            )));
        }
        size = size
            .checked_add(2 + units * 2)
            .ok_or_else(|| corrupted("SttbfAssoc serialized size overflows"))?;
    }
    if size > MAX_TABLE_BYTES {
        return Err(corrupted(
            "SttbfAssoc exceeds its specification-derived size cap",
        ));
    }
    Ok(size)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(values: &[&str; STRING_COUNT]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(STRING_COUNT as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for value in values {
            let units = value.encode_utf16().collect::<Vec<_>>();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        data
    }

    #[test]
    fn typed_unicode_slots_round_trip_exactly() {
        let mut values = [""; STRING_COUNT];
        values[1] = "C:\\模板\\Normal.dot";
        values[2] = "Quarterly 😀";
        values[6] = "张三";
        values[7] = "Alice";
        values[17] = "reserve";
        let bytes = table(&values);
        let parsed = DocumentAssociatedStrings::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.template_path(), "C:\\模板\\Normal.dot");
        assert_eq!(parsed.title(), "Quarterly 😀");
        assert_eq!(parsed.author(), "张三");
        assert_eq!(parsed.last_revised_by(), "Alice");
        assert_eq!(parsed.iter().len(), STRING_COUNT);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_header_lengths_utf16_and_trailing_data() {
        let empty = [""; STRING_COUNT];
        let mut wrong_count = table(&empty);
        wrong_count[2..4].copy_from_slice(&17u16.to_le_bytes());
        assert!(DocumentAssociatedStrings::parse_bytes(&wrong_count).is_err());

        let mut wrong_extra = table(&empty);
        wrong_extra[4..6].copy_from_slice(&1u16.to_le_bytes());
        assert!(DocumentAssociatedStrings::parse_bytes(&wrong_extra).is_err());

        let mut truncated = table(&empty);
        truncated.pop();
        assert!(DocumentAssociatedStrings::parse_bytes(&truncated).is_err());

        let mut invalid_utf16 = table(&empty);
        invalid_utf16[6..8].copy_from_slice(&1u16.to_le_bytes());
        invalid_utf16.insert(8, 0x00);
        invalid_utf16.insert(9, 0xD8);
        assert!(DocumentAssociatedStrings::parse_bytes(&invalid_utf16).is_err());

        let mut trailing = table(&empty);
        trailing.push(0);
        assert!(DocumentAssociatedStrings::parse_bytes(&trailing).is_err());

        let mut oversized_password = std::array::from_fn(|_| String::new());
        oversized_password[17] = "x".repeat(16);
        assert!(DocumentAssociatedStrings::try_new(oversized_password).is_err());
    }
}
