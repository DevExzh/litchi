//! Binary codecs for `CAPI`, `SttbfCaption`, and `SttbfAutoCaption`.

use super::Format;
use super::model::{
    AutoEntry, AutoTable, Definition, Heading, Info, LabelTable, Location, Numbering, Separator,
};
use super::validation::{
    MAX_ENTRIES, MAX_LABEL_UNITS, MAX_STRING_UNITS, corrupted, validate_auto_entries,
    validate_definitions, validate_table_size,
};
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;

/// FIB index of `fcSttbfCaption`/`lcbSttbfCaption` (MS-DOC 2.5.6).
pub(super) const CAPTION_FIB_INDEX: usize = 52;
/// FIB index of `fcSttbfAutoCaption`/`lcbSttbfAutoCaption` (MS-DOC 2.5.6).
pub(super) const AUTO_CAPTION_FIB_INDEX: usize = 53;
/// Serialized size of one `CAPI` extra-data value (MS-DOC 2.9.24).
pub(super) const CAPI_SIZE: usize = 6;
/// Serialized size of one `SttbfAutoCaption` extra-data value.
pub(super) const AUTO_CAPTION_EXTRA_SIZE: u16 = 2;

/// Slice one optional caption table from a FIB/table-stream pair.
pub(super) fn parse_fib_table<'a>(
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
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], context: &str) -> Result<String> {
    if !data.len().is_multiple_of(2) {
        return Err(corrupted(format!("{context} has an odd byte length")));
    }
    char::decode_utf16(
        data.chunks_exact(2)
            .map(|unit| u16::from_le_bytes([unit[0], unit[1]])),
    )
    .collect::<std::result::Result<String, _>>()
    .map_err(|error| corrupted(format!("invalid {context}: {error}")))
}

fn parse_sttb_header(data: &[u8], extra: u16, name: &str) -> Result<usize> {
    validate_table_size(data, name)?;
    if data.len() < 6
        || read_u16(data, 0, &format!("{name} fExtend"))? != u16::MAX
        || read_u16(data, 4, &format!("{name} cbExtra"))? != extra
    {
        return Err(corrupted(format!("{name} has an invalid header")));
    }
    Ok(usize::from(read_u16(data, 2, &format!("{name} cData"))?))
}

fn parse_string(
    data: &[u8],
    offset: usize,
    max_units: usize,
    table: &str,
    index: usize,
) -> Result<(String, usize)> {
    let units = usize::from(read_u16(
        data,
        offset,
        &format!("{table} string {index} length"),
    )?);
    if units > max_units {
        return Err(corrupted(format!(
            "{table} string {index} exceeds {max_units} UTF-16 code units"
        )));
    }
    let start = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{table} string offset overflows")))?;
    let byte_len = units
        .checked_mul(2)
        .ok_or_else(|| corrupted(format!("{table} string size overflows")))?;
    let end = start
        .checked_add(byte_len)
        .ok_or_else(|| corrupted(format!("{table} string range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{table} string {index} is truncated")))?;
    Ok((
        decode_utf16(bytes, &format!("{table} string {index}"))?,
        end,
    ))
}

fn append_string(out: &mut Vec<u8>, value: &str, table: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    let units = u16::try_from(units)
        .map_err(|_| corrupted(format!("{table} string exceeds the STTB length field")))?;
    out.extend_from_slice(&units.to_le_bytes());
    for unit in value.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn serialized_string_len(value: &str, table: &str) -> Result<usize> {
    let units = value.encode_utf16().count();
    u16::try_from(units)
        .map_err(|_| corrupted(format!("{table} string exceeds the STTB length field")))?;
    units
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(2))
        .ok_or_else(|| corrupted(format!("{table} serialized length overflows")))
}

impl Info {
    /// Serialized size of one `CAPI` value.
    pub const SIZE: usize = CAPI_SIZE;

    /// Decode exactly one `CAPI` value.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("CAPI must be exactly 6 bytes"));
        }
        let flags = read_u16(data, 0, "CAPI flags")?;
        let numbering = if flags & 0x4 != 0 {
            Some(Numbering::new(
                Heading::from_raw(((flags >> 3) & 0xF) as u8)?,
                Separator::from_raw(read_u16(data, 4, "CAPI xchSeparator")?)?,
            ))
        } else {
            None
        };
        let mut value = Self::new(
            Location::from_raw((flags & 0x3) as u8)?,
            numbering,
            flags & 0x8000 != 0,
            parse_format(read_u16(data, 2, "CAPI nfc")?)?,
        );
        value.raw_flags = flags;
        value.raw_separator = read_u16(data, 4, "CAPI xchSeparator")?;
        Ok(value)
    }

    /// Serialize while retaining undefined CAPI bits and ignored fields.
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let mut flags = self.raw_flags;
        flags = (flags & !0x0003) | self.location() as u16;
        flags = (flags & !0x8000) | u16::from(self.omit_label()) << 15;
        if let Some(numbering) = self.numbering() {
            flags = (flags & !0x007C) | 0x0004 | (numbering.heading() as u16) << 3;
        } else {
            flags &= !0x0004;
        }
        let separator = self
            .numbering()
            .map_or(self.raw_separator, |numbering| numbering.separator() as u16);
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&flags.to_le_bytes());
        data[2..4].copy_from_slice(&(self.number_format() as u8 as u16).to_le_bytes());
        data[4..6].copy_from_slice(&separator.to_le_bytes());
        data
    }
}

fn parse_format(value: u16) -> Result<Format> {
    let value = u8::try_from(value)
        .map_err(|_| corrupted(format!("CAPI nfc has a nonzero high byte: 0x{value:04X}")))?;
    Format::try_from(value)
        .map_err(|invalid| corrupted(format!("CAPI nfc has invalid MSONFC value 0x{invalid:02X}")))
}

impl LabelTable {
    /// Decode one complete `SttbfCaption` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let count = parse_sttb_header(data, CAPI_SIZE as u16, "SttbfCaption")?;
        if count > MAX_ENTRIES {
            return Err(corrupted("SttbfCaption count exceeds 65535 entries"));
        }
        let mut definitions = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let (label, next) = parse_string(data, offset, MAX_LABEL_UNITS, "SttbfCaption", index)?;
            let end = next
                .checked_add(CAPI_SIZE)
                .ok_or_else(|| corrupted("SttbfCaption CAPI range overflows"))?;
            let extra = data.get(next..end).ok_or_else(|| {
                corrupted(format!("SttbfCaption entry {index} CAPI is truncated"))
            })?;
            definitions.push(Definition::try_new(label, Info::from_bytes(extra)?)?);
            offset = end;
        }
        if offset != data.len() {
            return Err(corrupted("SttbfCaption has trailing bytes"));
        }
        Self::try_new(definitions)
    }

    /// Serialize one complete `SttbfCaption` payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_definitions(self.definitions())?;
        let mut capacity = 6usize;
        for definition in self.definitions() {
            capacity = capacity
                .checked_add(serialized_string_len(definition.label(), "SttbfCaption")?)
                .and_then(|value| value.checked_add(CAPI_SIZE))
                .ok_or_else(|| corrupted("SttbfCaption serialized length overflows"))?;
        }
        if capacity > super::validation::MAX_TABLE_BYTES {
            return Err(corrupted("SttbfCaption exceeds the table size cap"));
        }
        let mut data = Vec::with_capacity(capacity);
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(self.len() as u16).to_le_bytes());
        data.extend_from_slice(&(CAPI_SIZE as u16).to_le_bytes());
        for definition in self.definitions() {
            append_string(&mut data, definition.label(), "SttbfCaption")?;
            data.extend_from_slice(&definition.info().to_bytes());
        }
        Ok(data)
    }
}

impl AutoTable {
    /// Decode one complete `SttbfAutoCaption` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let count = parse_sttb_header(data, AUTO_CAPTION_EXTRA_SIZE, "SttbfAutoCaption")?;
        if count > MAX_ENTRIES {
            return Err(corrupted("SttbfAutoCaption count exceeds 65535 entries"));
        }
        let mut entries = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let (prog_id, next) =
                parse_string(data, offset, MAX_STRING_UNITS, "SttbfAutoCaption", index)?;
            let caption_index = read_u16(
                data,
                next,
                &format!("SttbfAutoCaption entry {index} caption index"),
            )?;
            let end = next
                .checked_add(usize::from(AUTO_CAPTION_EXTRA_SIZE))
                .ok_or_else(|| corrupted("SttbfAutoCaption extra-data range overflows"))?;
            if data.get(next..end).is_none() {
                return Err(corrupted(format!(
                    "SttbfAutoCaption entry {index} extra data is truncated"
                )));
            }
            entries.push(AutoEntry::try_new(prog_id, caption_index)?);
            offset = end;
        }
        if offset != data.len() {
            return Err(corrupted("SttbfAutoCaption has trailing bytes"));
        }
        Self::try_new(entries)
    }

    /// Serialize one complete `SttbfAutoCaption` payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_auto_entries(self.entries())?;
        let mut capacity = 6usize;
        for entry in self.entries() {
            capacity = capacity
                .checked_add(serialized_string_len(entry.prog_id(), "SttbfAutoCaption")?)
                .and_then(|value| value.checked_add(usize::from(AUTO_CAPTION_EXTRA_SIZE)))
                .ok_or_else(|| corrupted("SttbfAutoCaption serialized length overflows"))?;
        }
        if capacity > super::validation::MAX_TABLE_BYTES {
            return Err(corrupted("SttbfAutoCaption exceeds the table size cap"));
        }
        let mut data = Vec::with_capacity(capacity);
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(self.len() as u16).to_le_bytes());
        data.extend_from_slice(&AUTO_CAPTION_EXTRA_SIZE.to_le_bytes());
        for entry in self.entries() {
            append_string(&mut data, entry.prog_id(), "SttbfAutoCaption")?;
            data.extend_from_slice(&entry.caption_index().to_le_bytes());
        }
        Ok(data)
    }
}
