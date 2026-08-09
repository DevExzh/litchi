//! Bounded binary codecs and FIB routing for mail-merge parts.
//! Every slice is consumed exactly; no external data source is opened.

use super::model::{
    DocumentMailMerge, FieldMapInfo, FieldMapping, FilterComparison, FilterCondition,
    FilterDataItem, Fnpi, MailMergeDocumentType, MergeDataSourceKind, OdsoProperty, Pmfs, Pms,
    RecipientEntry, RecipientInfo, Rfs, SortColumnAndDirection, SortDirection, SttbfRfs, Wpms,
};
use super::validation::{
    CB_COUNT, COUNT_MARKER, FC_ODSO, FC_PMS, FC_PMS_NEW, FIELD_MAP_COLUMN_INDEX,
    FIELD_MAP_COLUMN_NAME, FIELD_MAP_COLUMN_NIL, FIELD_MAP_COUNT, FIELD_MAP_FIELD_NAME,
    FIELD_MAP_MAPPED, FIELD_MAP_MAPPED_VALUE, FILTER_ITEM_HEADER_LEN, IREC_MAX, IREC_NIL,
    ITEM_TERMINATOR, LIST_SIZE_MARKER, LIST_SIZE_OVERFLOW, MAX_COLUMN_INDEX, MAX_FILTER_CHARS,
    MAX_SORT_KEYS, ODSO_ID_COLUMN_DELIMITER, ODSO_ID_CONNECTION_STRING, ODSO_ID_CONNECTION_TYPE,
    ODSO_ID_DATA_SOURCE_FILE, ODSO_ID_DATA_TABLE, ODSO_ID_FIELD_MAP, ODSO_ID_FIRST_ROW_IS_HEADER,
    ODSO_ID_RECIPIENT_FILTERS, ODSO_ID_RECIPIENTS, ODSO_ID_SORT_ORDER, ODSO_ID_WIZARD_STEP,
    ODSO_LARGE, PMFS_LEN, PMFS_LINK_TO_CONNECTION, PMFS_LINK_TO_FILE, PMFS_NO_PROMPT_QT,
    PMFS_QUERY, PMS_HEADER_LEN, RECIPIENT_HASH, RECIPIENT_INCLUDED, RECIPIENT_UNIQUE_COLUMN,
    RECIPIENT_UNIQUE_VALUE, SORT_KEY_LEN, SQL_MAX_BYTES, SQL_MIN_BYTES, STTB_F_EXTEND,
    STTBF_RFS_CB_EXTRA, STTBF_RFS_MAX_CHARS, STTBF_RFS_MAX_STRINGS, STTBF_RFS_MIN_STRINGS,
    WIZARD_STEP_MAX, WIZARD_STEP_MIN, WPMSDT_DOC_TYPE_MASK, corrupted,
};
use crate::package::Result;
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};

/// A bounds-checked cursor over one structure's byte range.
struct Reader<'a> {
    data: &'a [u8],
    position: usize,
    context: &'static str,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8], context: &'static str) -> Self {
        Reader {
            data,
            position: 0,
            context,
        }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.position
    }

    fn bytes(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| corrupted(format!("{} range overflows", self.context)))?;
        let slice = self
            .data
            .get(self.position..end)
            .ok_or_else(|| corrupted(format!("{} is truncated", self.context)))?;
        self.position = end;
        Ok(slice)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.bytes(1)?[0])
    }

    fn u16(&mut self) -> Result<u16> {
        let bytes = self.bytes(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn u32(&mut self) -> Result<u32> {
        let bytes = self.bytes(4)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    /// Require that the whole range was consumed exactly.
    fn finish(&self) -> Result<()> {
        if self.remaining() != 0 {
            return Err(corrupted(format!("{} has trailing bytes", self.context)));
        }
        Ok(())
    }
}

/// Decode a UTF-16LE string that must occupy `bytes` exactly.
fn decode_utf16(bytes: &[u8], context: &str) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(corrupted(format!("{context} has an odd byte length")));
    }
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{context} is not valid UTF-16")))
}

impl Pmfs {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        debug_assert_eq!(data.len(), PMFS_LEN);
        let flags = data[1];
        Ok(Pmfs {
            source_kind: MergeDataSourceKind::parse(data[0])?,
            link_to_file: flags & PMFS_LINK_TO_FILE != 0,
            link_to_connection: flags & PMFS_LINK_TO_CONNECTION != 0,
            no_prompt_query_tools: flags & PMFS_NO_PROMPT_QT != 0,
            uses_query: flags & PMFS_QUERY != 0,
            field_token: i16::from_le_bytes([data[2], data[3]]),
            record_token: i16::from_le_bytes([data[4], data[5]]),
            file_name: Fnpi {
                raw: u16::from_le_bytes([data[6], data[7]]),
            },
        })
    }
}

impl SttbfRfs {
    fn parse(reader: &mut Reader<'_>) -> Result<Self> {
        if reader.u16()? != STTB_F_EXTEND {
            return Err(corrupted("SttbfRfs.fExtend is not 0xFFFF"));
        }
        let count = reader.u16()?;
        if !(STTBF_RFS_MIN_STRINGS..=STTBF_RFS_MAX_STRINGS).contains(&count) {
            return Err(corrupted("SttbfRfs.cData is not 4 or 5"));
        }
        if reader.u16()? != STTBF_RFS_CB_EXTRA {
            return Err(corrupted("SttbfRfs.cbExtra is not zero"));
        }
        let mut strings = Vec::with_capacity(usize::from(count));
        for _ in 0..count {
            let chars = reader.u16()?;
            if chars > STTBF_RFS_MAX_CHARS {
                return Err(corrupted("SttbfRfs string exceeds 255 characters"));
            }
            let raw = reader.bytes(usize::from(chars) * 2)?;
            strings.push(decode_utf16(raw, "SttbfRfs string")?);
        }
        Ok(SttbfRfs { strings })
    }
}

impl Pms {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < PMS_HEADER_LEN {
            return Err(corrupted("Pms is truncated"));
        }
        let mut reader = Reader::new(data, "Pms");
        let state = Wpms::parse(reader.u16()?)?;
        let header_source_index = reader.u8()?;
        let fetch_source_index = reader.u8()?;
        if header_source_index > 1 || fetch_source_index > 1 {
            return Err(corrupted("Pms source index is not 0 or 1"));
        }
        let current_record = match reader.u32()? {
            IREC_NIL => None,
            value if value <= IREC_MAX => Some(value),
            _ => return Err(corrupted("Pms.iRecCur is out of range")),
        };
        let sources = [
            Pmfs::parse(reader.bytes(PMFS_LEN)?)?,
            Pmfs::parse(reader.bytes(PMFS_LEN)?)?,
        ];
        let filter = Rfs::parse(reader.u32()?)?;
        let sql_length = reader.u16()?;
        let sql_query = if sql_length == 0 {
            None
        } else {
            if sql_length % 2 != 0 {
                return Err(corrupted("Pms.cblszSqlStr is not even"));
            }
            if !(SQL_MIN_BYTES..=SQL_MAX_BYTES).contains(&sql_length) {
                return Err(corrupted("Pms.cblszSqlStr is out of range"));
            }
            let raw = reader.bytes(usize::from(sql_length))?;
            let (text, terminator) = raw.split_at(raw.len() - 2);
            if terminator != [0, 0] {
                return Err(corrupted("Pms.lxszSqlStr is not null-terminated"));
            }
            Some(decode_utf16(text, "Pms.lxszSqlStr")?)
        };
        let strings = if filter.has_string_table {
            Some(SttbfRfs::parse(&mut reader)?)
        } else {
            None
        };
        let document_type = match reader.remaining() {
            0 => None,
            4 => Some(MailMergeDocumentType::parse(
                reader.u32()? & WPMSDT_DOC_TYPE_MASK,
            )?),
            _ => return Err(corrupted("Pms has a partial Wpmsdt")),
        };
        reader.finish()?;
        Ok(Pms {
            state,
            header_source_index,
            fetch_source_index,
            current_record,
            sources,
            filter,
            sql_query,
            strings,
            document_type,
        })
    }
}

impl FilterDataItem {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "FilterDataItem");
        let column = reader.u32()?;
        if column > MAX_COLUMN_INDEX {
            return Err(corrupted("FilterDataItem.iColumn is out of range"));
        }
        let comparison = FilterComparison::parse(reader.u32()?)?;
        let condition = FilterCondition::parse(reader.u32()?)?;
        let raw = reader.bytes(reader.remaining())?;
        if raw.len() < 2 {
            return Err(corrupted("FilterDataItem string is truncated"));
        }
        let (text, terminator) = raw.split_at(raw.len() - 2);
        if terminator != [0, 0] {
            return Err(corrupted("FilterDataItem string is not null-terminated"));
        }
        let value = decode_utf16(text, "FilterDataItem string")?;
        if value.chars().count() > MAX_FILTER_CHARS {
            return Err(corrupted("FilterDataItem string exceeds 212 characters"));
        }
        reader.finish()?;
        Ok(FilterDataItem {
            column,
            comparison,
            condition,
            value,
        })
    }

    pub(super) fn parse_list(data: &[u8]) -> Result<Vec<Self>> {
        let mut reader = Reader::new(data, "recipient filter list");
        let mut filters = Vec::new();
        while reader.remaining() > 0 {
            let item_size = usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("FilterDataItem.cbItem exceeds usize"))?;
            if item_size < FILTER_ITEM_HEADER_LEN as usize + 2 {
                return Err(corrupted("FilterDataItem.cbItem is too small"));
            }
            filters.push(Self::parse(reader.bytes(item_size - 4)?)?);
        }
        Ok(filters)
    }
}

impl SortColumnAndDirection {
    pub(super) fn parse_list(data: &[u8]) -> Result<Vec<Self>> {
        if !data.len().is_multiple_of(SORT_KEY_LEN) {
            return Err(corrupted("sort key list has a partial item"));
        }
        if data.len() / SORT_KEY_LEN > MAX_SORT_KEYS {
            return Err(corrupted("sort key list exceeds three items"));
        }
        let mut keys = Vec::with_capacity(data.len() / SORT_KEY_LEN);
        for chunk in data.chunks_exact(SORT_KEY_LEN) {
            let column = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
            if column > MAX_COLUMN_INDEX {
                return Err(corrupted("SortColumnAndDirection.iColumn is out of range"));
            }
            let direction =
                SortDirection::parse(u32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]))?;
            keys.push(SortColumnAndDirection { column, direction });
        }
        Ok(keys)
    }
}

impl RecipientInfo {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "RecipientInfo");
        if reader.u16()? != COUNT_MARKER {
            return Err(corrupted("RecipientInfo.countMarker is not zero"));
        }
        if reader.u16()? != CB_COUNT {
            return Err(corrupted("RecipientInfo.cbCount is not 4"));
        }
        let count = usize::try_from(reader.u32()?)
            .map_err(|_| corrupted("RecipientInfo.cRecipients exceeds usize"))?;
        if reader.u16()? != LIST_SIZE_MARKER {
            return Err(corrupted("RecipientInfo list size marker is not 1"));
        }
        let short_size = reader.u16()?;
        let list_size = if short_size == LIST_SIZE_OVERFLOW {
            usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("RecipientInfo list size exceeds usize"))?
        } else {
            usize::from(short_size)
        };
        let list = reader.bytes(list_size)?;
        reader.finish()?;
        let mut items = Reader::new(list, "RecipientInfo recipients");
        let mut recipients = Vec::with_capacity(count);
        for _ in 0..count {
            let mut recipient = RecipientEntry {
                included: true,
                ..RecipientEntry::default()
            };
            loop {
                let id = items.u16()?;
                let size = usize::from(items.u16()?);
                if id == ITEM_TERMINATOR {
                    if size != 0 {
                        return Err(corrupted("RecipientTerminator has data"));
                    }
                    break;
                }
                let value = items.bytes(size)?;
                match id {
                    RECIPIENT_INCLUDED => {
                        if size != 4 {
                            return Err(corrupted("recipient inclusion item is not 4 bytes"));
                        }
                        recipient.included =
                            match u32::from_le_bytes([value[0], value[1], value[2], value[3]]) {
                                0 => false,
                                1 => true,
                                _ => {
                                    return Err(corrupted(
                                        "recipient inclusion value is not 0 or 1",
                                    ));
                                },
                            };
                    },
                    RECIPIENT_UNIQUE_COLUMN | RECIPIENT_HASH => {
                        if size != 4 {
                            return Err(corrupted("recipient integer item is not 4 bytes"));
                        }
                        let number = u32::from_le_bytes([value[0], value[1], value[2], value[3]]);
                        if id == RECIPIENT_UNIQUE_COLUMN {
                            recipient.unique_column = Some(number);
                        } else {
                            recipient.record_hash = Some(number);
                        }
                    },
                    RECIPIENT_UNIQUE_VALUE => {
                        recipient.unique_value =
                            Some(decode_utf16(value, "recipient unique value")?);
                    },
                    _ => return Err(corrupted("recipient item id is not defined")),
                }
            }
            recipients.push(recipient);
        }
        items.finish()?;
        Ok(RecipientInfo { recipients })
    }
}

impl FieldMapInfo {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "FieldMapInfo");
        if reader.u16()? != COUNT_MARKER {
            return Err(corrupted("FieldMapInfo.countMarker is not zero"));
        }
        if reader.u16()? != CB_COUNT {
            return Err(corrupted("FieldMapInfo.cbCount is not 4"));
        }
        if reader.u32()? != FIELD_MAP_COUNT {
            return Err(corrupted("FieldMapInfo.cFields is not 30"));
        }
        if reader.u16()? != LIST_SIZE_MARKER {
            return Err(corrupted("FieldMapInfo list size marker is not 1"));
        }
        let short_size = reader.u16()?;
        let list_size = if short_size == LIST_SIZE_OVERFLOW {
            usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("FieldMapInfo list size exceeds usize"))?
        } else {
            usize::from(short_size)
        };
        let list = reader.bytes(list_size)?;
        reader.finish()?;
        let mut items = Reader::new(list, "FieldMapInfo mappings");
        let mut mappings = Vec::with_capacity(FIELD_MAP_COUNT as usize);
        for _ in 0..FIELD_MAP_COUNT {
            let mut mapping = FieldMapping::default();
            loop {
                let id = items.u16()?;
                let size = usize::from(items.u16()?);
                if id == ITEM_TERMINATOR {
                    if size != 0 {
                        return Err(corrupted("FieldMapTerminator has data"));
                    }
                    break;
                }
                let value = items.bytes(size)?;
                match id {
                    FIELD_MAP_MAPPED => {
                        if size != 4
                            || u32::from_le_bytes([value[0], value[1], value[2], value[3]])
                                != FIELD_MAP_MAPPED_VALUE
                        {
                            return Err(corrupted("field map mapped flag is not 1"));
                        }
                    },
                    FIELD_MAP_COLUMN_NAME => {
                        mapping.column_name = Some(decode_utf16(value, "field map column name")?);
                    },
                    FIELD_MAP_FIELD_NAME => {
                        // The standard field name is ignored by definition
                        // (MS-DOC 2.9.84); only its framing is validated.
                    },
                    FIELD_MAP_COLUMN_INDEX => {
                        if size != 4 {
                            return Err(corrupted("field map column index is not 4 bytes"));
                        }
                        let index = u32::from_le_bytes([value[0], value[1], value[2], value[3]]);
                        if index != FIELD_MAP_COLUMN_NIL {
                            mapping.column_index = Some(index);
                        }
                    },
                    _ => return Err(corrupted("field map item id is not defined")),
                }
            }
            mappings.push(mapping);
        }
        items.finish()?;
        Ok(FieldMapInfo { mappings })
    }
}

impl OdsoProperty {
    pub(super) fn decode(id: u16, value: &[u8]) -> Result<Self> {
        Ok(match id {
            ODSO_ID_CONNECTION_STRING => {
                Self::ConnectionString(decode_utf16(value, "ODSO connection string")?)
            },
            ODSO_ID_DATA_TABLE => Self::DataTable(decode_utf16(value, "ODSO data table")?),
            ODSO_ID_DATA_SOURCE_FILE => {
                Self::DataSourceFile(decode_utf16(value, "ODSO data source file")?)
            },
            ODSO_ID_CONNECTION_TYPE => Self::ConnectionType(expect_u32(value, "ODSO property")?),
            ODSO_ID_COLUMN_DELIMITER => {
                if value.len() != 2 {
                    return Err(corrupted("ODSO column delimiter is not 2 bytes"));
                }
                Self::ColumnDelimiter(u16::from_le_bytes([value[0], value[1]]))
            },
            ODSO_ID_FIRST_ROW_IS_HEADER => match expect_u32(value, "ODSO property")? {
                0 => Self::FirstRowIsHeader(false),
                1 => Self::FirstRowIsHeader(true),
                _ => return Err(corrupted("ODSO first-row flag is not 0 or 1")),
            },
            ODSO_ID_RECIPIENT_FILTERS => Self::RecipientFilters(FilterDataItem::parse_list(value)?),
            ODSO_ID_SORT_ORDER => Self::SortOrder(SortColumnAndDirection::parse_list(value)?),
            ODSO_ID_RECIPIENTS => Self::Recipients(RecipientInfo::parse(value)?),
            ODSO_ID_FIELD_MAP => Self::FieldMap(FieldMapInfo::parse(value)?),
            ODSO_ID_WIZARD_STEP => {
                if value.len() != 2 {
                    return Err(corrupted("ODSO wizard step is not 2 bytes"));
                }
                let step = u16::from_le_bytes([value[0], value[1]]);
                if !(WIZARD_STEP_MIN..=WIZARD_STEP_MAX).contains(&step) {
                    return Err(corrupted("ODSO wizard step is not between 1 and 6"));
                }
                Self::WizardStep(step)
            },
            _ => Self::Unknown {
                id,
                data: value.to_vec(),
            },
        })
    }
}

fn expect_u32(value: &[u8], context: &str) -> Result<u32> {
    if value.len() != 4 {
        return Err(corrupted(format!("{context} is not 4 bytes")));
    }
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

impl DocumentMailMerge {
    /// Parse the mail-merge state addressed by the FIB, or `None` when the
    /// document carries neither a `Pms` nor ODSO data.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentMailMerge>> {
        // The Word 6/95 FIB table-pointer layout assigns these indices to
        // unrelated structures, so they only carry merge state from Word 97
        // on.
        if fib.version() < WORD_97_NFIB {
            return Ok(None);
        }
        let state = optional_slice(fib, table_stream, FC_PMS, "Pms")?
            .map(Pms::parse)
            .transpose()?;
        let new_state = optional_slice(fib, table_stream, FC_PMS_NEW, "PmsNew")?
            .map(Pms::parse)
            .transpose()?;
        let odso_properties = match optional_slice(fib, table_stream, FC_ODSO, "ODSO data")? {
            Some(data) => parse_odso_properties(data)?,
            None => Vec::new(),
        };
        if state.is_none() && new_state.is_none() && odso_properties.is_empty() {
            return Ok(None);
        }
        Ok(Some(DocumentMailMerge {
            state,
            new_state,
            odso_properties,
        }))
    }
}

/// Parse the ODSO property bag, which is a sequence of variable-length
/// `ODSOPropertyBase` items filling the byte range exactly.
pub(super) fn parse_odso_properties(data: &[u8]) -> Result<Vec<OdsoProperty>> {
    let mut reader = Reader::new(data, "ODSO property set");
    let mut properties = Vec::new();
    while reader.remaining() > 0 {
        let id = reader.u16()?;
        let size = reader.u16()?;
        let value = if size == ODSO_LARGE {
            let large_size = usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("ODSO property size exceeds usize"))?;
            reader.bytes(large_size)?
        } else {
            reader.bytes(usize::from(size))?
        };
        properties.push(OdsoProperty::decode(id, value)?);
    }
    Ok(properties)
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}
