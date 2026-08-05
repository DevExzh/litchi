//! Strict BIFF8/MS-XLS HLink and HLinkTooltip codecs.

use super::model::{
    FileMoniker, Hyperlink, HyperlinkMoniker, HyperlinkRange, ItemMoniker, TOOLTIP_RECORD_TYPE,
    UrlMoniker,
};
use crate::error::{Error, Result};

pub(super) const URL_MONIKER_CLSID: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
pub(super) const FILE_MONIKER_CLSID: [u8; 16] =
    [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const COMPOSITE_MONIKER_CLSID: [u8; 16] =
    [0x09, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const ANTI_MONIKER_CLSID: [u8; 16] = [0x05, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const ITEM_MONIKER_CLSID: [u8; 16] = [0x04, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
pub(super) const URL_SERIAL_GUID: [u8; 16] = [
    0x79, 0x58, 0x81, 0xF4, 0x3B, 0x1D, 0x7F, 0x48, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63,
];

pub fn parse_hlink_record(data: &[u8]) -> Result<Hyperlink> {
    let mut cursor = Cursor::new(data);
    let range = cursor.range()?;
    let class_id = cursor.guid()?;
    // MS-XLS 2.4.140 defines hlinkClsid as the producer CLSID and does not
    // constrain it to the standard hyperlink class. Keep the value for
    // diagnostics and preserve hyperlinks emitted by other COM producers.
    let version = cursor.u32()?;
    if version != 2 {
        return invalid(format!("HLink stream version must be 2, got {version}"));
    }
    let flags = cursor.u32()?;
    if flags & !0x03FF != 0 {
        return invalid(format!("HLink contains reserved flag bits: {flags:#010x}"));
    }
    let has_moniker = flags & 1 != 0;
    let site_gave_display_name = flags & 4 != 0;
    let has_location = flags & 8 != 0;
    let has_display = flags & 0x10 != 0;
    let has_guid = flags & 0x20 != 0;
    let has_time = flags & 0x40 != 0;
    let has_frame = flags & 0x80 != 0;
    let string_moniker = flags & 0x100 != 0;
    if string_moniker && !has_moniker {
        return invalid("HLink string-moniker flag requires the moniker flag".to_string());
    }
    if site_gave_display_name && !has_display {
        return invalid("HLink site-display-name flag requires a display name".to_string());
    }
    if !has_moniker && !has_location {
        return invalid("HLink must contain a moniker or location".to_string());
    }
    let display_name = has_display.then(|| cursor.hyperlink_string()).transpose()?;
    let target_frame = has_frame.then(|| cursor.hyperlink_string()).transpose()?;
    let moniker = if string_moniker {
        Some(HyperlinkMoniker::String(cursor.hyperlink_string()?))
    } else if has_moniker {
        Some(parse_moniker(&mut cursor, 0)?)
    } else {
        None
    };
    let location = has_location
        .then(|| cursor.hyperlink_string())
        .transpose()?;
    let hyperlink_guid = has_guid.then(|| cursor.guid()).transpose()?;
    let creation_time = has_time.then(|| cursor.u64()).transpose()?;
    if cursor.remaining() != 0 {
        return invalid(format!(
            "HLink contains {} trailing bytes",
            cursor.remaining()
        ));
    }
    Ok(Hyperlink {
        range,
        class_id,
        absolute: flags & 2 != 0,
        site_gave_display_name,
        absolute_from_relative: flags & 0x200 != 0,
        display_name,
        target_frame,
        moniker,
        location,
        hyperlink_guid,
        creation_time,
        tooltip: None,
    })
}

pub(super) fn parse_tooltip(data: &[u8]) -> Result<(HyperlinkRange, String)> {
    if data.len() < 14 {
        return invalid(format!("HLinkTooltip payload is too short: {}", data.len()));
    }
    let mut cursor = Cursor::new(data);
    if cursor.u16()? != TOOLTIP_RECORD_TYPE {
        return invalid("HLinkTooltip internal record type must be 0x0800".to_string());
    }
    let range = cursor.range()?;
    if !cursor.remaining().is_multiple_of(2) {
        return invalid("HLinkTooltip string has an odd byte length".to_string());
    }
    let units = cursor.remaining() / 2;
    if !(2..=256).contains(&units) {
        return invalid(format!(
            "HLinkTooltip length must be 2..=256 UTF-16 units, got {units}"
        ));
    }
    let remaining = cursor.remaining();
    Ok((range, decode_terminated_utf16(cursor.take(remaining)?)?))
}

fn parse_moniker(cursor: &mut Cursor<'_>, depth: usize) -> Result<HyperlinkMoniker> {
    if depth >= 16 {
        return invalid("HLink composite-moniker nesting exceeds 16".to_string());
    }
    let clsid = cursor.guid()?;
    if clsid == URL_MONIKER_CLSID {
        return Ok(HyperlinkMoniker::Url(parse_url_moniker(cursor)?));
    }
    if clsid == FILE_MONIKER_CLSID {
        return Ok(HyperlinkMoniker::File(parse_file_moniker(cursor)?));
    }
    if clsid == COMPOSITE_MONIKER_CLSID {
        let count = cursor.u32()? as usize;
        if count == 0 || count > cursor.remaining() / 20 {
            return invalid(format!("invalid composite moniker count: {count}"));
        }
        let mut monikers = Vec::with_capacity(count.min(256));
        for _ in 0..count {
            monikers.push(parse_moniker(cursor, depth + 1)?);
        }
        return Ok(HyperlinkMoniker::Composite(monikers));
    }
    if clsid == ANTI_MONIKER_CLSID {
        let count = cursor.u32()?;
        if count > 1_048_576 {
            return invalid(format!("anti-moniker count exceeds 1048576: {count}"));
        }
        return Ok(HyperlinkMoniker::Anti { count });
    }
    if clsid == ITEM_MONIKER_CLSID {
        return Ok(HyperlinkMoniker::Item(parse_item_moniker(cursor)?));
    }
    invalid("HLink contains an unknown moniker CLSID".to_string())
}

fn parse_url_moniker(cursor: &mut Cursor<'_>) -> Result<UrlMoniker> {
    let length = cursor.u32()? as usize;
    if length < 2 || !length.is_multiple_of(2) {
        return invalid(format!("invalid URLMoniker length: {length}"));
    }
    let data = cursor.take(length)?;
    let has_tail = length >= 26 && data[length - 24..length - 8] == URL_SERIAL_GUID;
    let (url_data, serialization_uri_flags) = if has_tail {
        let serial_version = u32::from_le_bytes(data[length - 8..length - 4].try_into().unwrap());
        if serial_version != 0 {
            return invalid(format!(
                "URLMoniker serial version must be zero, got {serial_version}"
            ));
        }
        let uri_flags = u32::from_le_bytes(data[length - 4..].try_into().unwrap());
        if uri_flags & 0xFFFF_0000 != 0 {
            return invalid(format!(
                "URLMoniker URI flags contain reserved bits: {uri_flags:#010x}"
            ));
        }
        (&data[..length - 24], Some(uri_flags))
    } else {
        (data, None)
    };
    if url_data.len() < 2 || url_data.len() % 2 != 0 {
        return invalid("URLMoniker URL has an invalid byte length".to_string());
    }
    Ok(UrlMoniker {
        url: decode_terminated_utf16(url_data)?,
        serialization_uri_flags,
    })
}

fn parse_file_moniker(cursor: &mut Cursor<'_>) -> Result<FileMoniker> {
    let parent_directory_count = cursor.u16()?;
    let ansi_length = cursor.u32()? as usize;
    if ansi_length == 0 || ansi_length > 32_767 {
        return invalid(format!(
            "FileMoniker ANSI path length must be 1..=32767, got {ansi_length}"
        ));
    }
    let ansi_path = decode_terminated_ansi(cursor.take(ansi_length)?)?;
    let end_server = cursor.u16()?;
    let is_unc = ansi_path.starts_with("\\\\");
    let unc_server_character_count = if is_unc {
        if end_server == 0xFFFF || usize::from(end_server) > ansi_path.encode_utf16().count() {
            return invalid("FileMoniker UNC server length is invalid".to_string());
        }
        Some(end_server)
    } else {
        if end_server != 0xFFFF {
            return invalid("non-UNC FileMoniker endServer must be 0xFFFF".to_string());
        }
        None
    };
    let version = cursor.u16()?;
    if version != 0xDEAD {
        return invalid(format!(
            "FileMoniker version must be 0xDEAD, got {version:#06x}"
        ));
    }
    if cursor.take(16)?.iter().any(|&byte| byte != 0) || cursor.u32()? != 0 {
        return invalid("FileMoniker reserved fields must be zero".to_string());
    }
    let unicode_size = cursor.u32()? as usize;
    let unicode_path = if unicode_size == 0 {
        None
    } else {
        if unicode_size < 6 {
            return invalid(format!(
                "FileMoniker Unicode block is too short: {unicode_size}"
            ));
        }
        let unicode_bytes = cursor.u32()? as usize;
        if !unicode_bytes.is_multiple_of(2) || unicode_size != unicode_bytes + 6 {
            return invalid("FileMoniker Unicode size fields are inconsistent".to_string());
        }
        if cursor.u16()? != 3 {
            return invalid("FileMoniker Unicode key must equal 3".to_string());
        }
        Some(decode_unterminated_utf16(cursor.take(unicode_bytes)?)?)
    };
    Ok(FileMoniker {
        parent_directory_count,
        ansi_path,
        unicode_path,
        unc_server_character_count,
    })
}

fn parse_item_moniker(cursor: &mut Cursor<'_>) -> Result<ItemMoniker> {
    let delimiter_length = cursor.u32()? as usize;
    let (delimiter_ansi, delimiter_unicode) = parse_item_string(cursor.take(delimiter_length)?)?;
    let item_length = cursor.u32()? as usize;
    let (item_ansi, item_unicode) = parse_item_string(cursor.take(item_length)?)?;
    Ok(ItemMoniker {
        delimiter_ansi,
        delimiter_unicode,
        item_ansi,
        item_unicode,
    })
}
fn parse_item_string(data: &[u8]) -> Result<(String, Option<String>)> {
    let terminator = data.iter().position(|&byte| byte == 0).ok_or_else(|| {
        Error::InvalidData("ItemMoniker ANSI string is not NUL-terminated".to_string())
    })?;
    let ansi = data[..terminator]
        .iter()
        .map(|&byte| char::from(byte))
        .collect();
    let unicode_data = &data[terminator + 1..];
    if !unicode_data.len().is_multiple_of(2) {
        return invalid("ItemMoniker Unicode string has an odd byte length".to_string());
    }
    let unicode = (!unicode_data.is_empty())
        .then(|| decode_unterminated_utf16(unicode_data))
        .transpose()?;
    Ok((ansi, unicode))
}

fn decode_terminated_utf16(data: &[u8]) -> Result<String> {
    if data.len() < 2 || !data.len().is_multiple_of(2) {
        return invalid("invalid terminated UTF-16 byte length".to_string());
    }
    let units: Vec<u16> = data
        .chunks_exact(2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .collect();
    if units.last() != Some(&0) || units[..units.len() - 1].contains(&0) {
        return invalid("hyperlink string must contain exactly one trailing NUL".to_string());
    }
    String::from_utf16(&units[..units.len() - 1])
        .map_err(|_| Error::InvalidData("hyperlink string contains invalid UTF-16".to_string()))
}
fn decode_unterminated_utf16(data: &[u8]) -> Result<String> {
    if !data.len().is_multiple_of(2) {
        return invalid("invalid UTF-16 byte length".to_string());
    }
    let units: Vec<u16> = data
        .chunks_exact(2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .collect();
    if units.contains(&0) {
        return invalid("unterminated hyperlink string contains NUL".to_string());
    }
    String::from_utf16(&units)
        .map_err(|_| Error::InvalidData("hyperlink string contains invalid UTF-16".to_string()))
}
fn decode_terminated_ansi(data: &[u8]) -> Result<String> {
    if data.last() != Some(&0) || data[..data.len() - 1].contains(&0) {
        return invalid("FileMoniker ANSI path must contain exactly one trailing NUL".to_string());
    }
    Ok(data[..data.len() - 1]
        .iter()
        .map(|&byte| char::from(byte))
        .collect())
}

struct Cursor<'a> {
    data: &'a [u8],
    position: usize,
}
impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }
    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.position)
    }
    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| Error::InvalidData("hyperlink field size overflow".to_string()))?;
        let data = self
            .data
            .get(self.position..end)
            .ok_or_else(|| Error::InvalidData("truncated hyperlink record".to_string()))?;
        self.position = end;
        Ok(data)
    }
    fn u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }
    fn u64(&mut self) -> Result<u64> {
        Ok(u64::from_le_bytes(self.take(8)?.try_into().unwrap()))
    }
    fn guid(&mut self) -> Result<[u8; 16]> {
        Ok(self.take(16)?.try_into().unwrap())
    }
    fn range(&mut self) -> Result<HyperlinkRange> {
        let first_row = self.u16()?;
        let last_row = self.u16()?;
        let first_column = self.u16()?;
        let last_column = self.u16()?;
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return invalid("hyperlink contains an invalid or out-of-range Ref8U".to_string());
        }
        Ok(HyperlinkRange {
            first_row,
            last_row,
            first_column: first_column as u8,
            last_column: last_column as u8,
        })
    }
    fn hyperlink_string(&mut self) -> Result<String> {
        let units = self.u32()? as usize;
        if units == 0 {
            return invalid("HyperlinkString length must include a NUL terminator".to_string());
        }
        let bytes = units
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidData("HyperlinkString size overflow".to_string()))?;
        decode_terminated_utf16(self.take(bytes)?)
    }
}

pub(super) fn invalid<T>(message: String) -> Result<T> {
    Err(Error::InvalidData(message))
}
