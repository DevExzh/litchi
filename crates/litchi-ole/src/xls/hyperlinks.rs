//! Strict, inert BIFF8 worksheet hyperlink retention.

use crate::xls::error::{XlsError, XlsResult};

pub const RECORD_TYPE: u16 = 0x01B8;
pub const TOOLTIP_RECORD_TYPE: u16 = 0x0800;

const STANDARD_HLINK_CLSID: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const URL_MONIKER_CLSID: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const FILE_MONIKER_CLSID: [u8; 16] = [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const COMPOSITE_MONIKER_CLSID: [u8; 16] =
    [0x09, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const ANTI_MONIKER_CLSID: [u8; 16] = [0x05, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const ITEM_MONIKER_CLSID: [u8; 16] = [0x04, 0x03, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const URL_SERIAL_GUID: [u8; 16] = [
    0x79, 0x58, 0x81, 0xF4, 0x3B, 0x1D, 0x7F, 0x48, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63,
];

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsHyperlinkRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}
impl XlsHyperlinkRange {
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsUrlMoniker {
    url: String,
    serialization_uri_flags: Option<u32>,
}
impl XlsUrlMoniker {
    pub fn url(&self) -> &str {
        &self.url
    }
    pub fn serialization_uri_flags(&self) -> Option<u32> {
        self.serialization_uri_flags
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsFileMoniker {
    parent_directory_count: u16,
    ansi_path: String,
    unicode_path: Option<String>,
    unc_server_character_count: Option<u16>,
}
impl XlsFileMoniker {
    pub fn parent_directory_count(&self) -> u16 {
        self.parent_directory_count
    }
    pub fn ansi_path(&self) -> &str {
        &self.ansi_path
    }
    pub fn unicode_path(&self) -> Option<&str> {
        self.unicode_path.as_deref()
    }
    pub fn path(&self) -> &str {
        self.unicode_path.as_deref().unwrap_or(&self.ansi_path)
    }
    pub fn unc_server_character_count(&self) -> Option<u16> {
        self.unc_server_character_count
    }
    pub fn is_unc(&self) -> bool {
        self.unc_server_character_count.is_some()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsItemMoniker {
    delimiter_ansi: String,
    delimiter_unicode: Option<String>,
    item_ansi: String,
    item_unicode: Option<String>,
}
impl XlsItemMoniker {
    pub fn delimiter(&self) -> &str {
        self.delimiter_unicode
            .as_deref()
            .unwrap_or(&self.delimiter_ansi)
    }
    pub fn item(&self) -> &str {
        self.item_unicode.as_deref().unwrap_or(&self.item_ansi)
    }
    pub fn delimiter_ansi(&self) -> &str {
        &self.delimiter_ansi
    }
    pub fn item_ansi(&self) -> &str {
        &self.item_ansi
    }
}

/// Serialized moniker data retained without activation or resolution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsHyperlinkMoniker {
    String(String),
    Url(XlsUrlMoniker),
    File(XlsFileMoniker),
    Composite(Vec<XlsHyperlinkMoniker>),
    Anti { count: u32 },
    Item(XlsItemMoniker),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsHyperlinkTargetKind {
    Document,
    Url,
    Email,
    File,
    Unc,
    StringMoniker,
    Composite,
    Anti,
    Item,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsHyperlink {
    range: XlsHyperlinkRange,
    class_id: [u8; 16],
    absolute: bool,
    site_gave_display_name: bool,
    absolute_from_relative: bool,
    display_name: Option<String>,
    target_frame: Option<String>,
    moniker: Option<XlsHyperlinkMoniker>,
    location: Option<String>,
    hyperlink_guid: Option<[u8; 16]>,
    creation_time: Option<u64>,
    tooltip: Option<String>,
}
impl XlsHyperlink {
    pub fn range(&self) -> XlsHyperlinkRange {
        self.range
    }
    pub fn class_id(&self) -> &[u8; 16] {
        &self.class_id
    }
    pub fn absolute(&self) -> bool {
        self.absolute
    }
    pub fn site_gave_display_name(&self) -> bool {
        self.site_gave_display_name
    }
    pub fn absolute_from_relative(&self) -> bool {
        self.absolute_from_relative
    }
    pub fn display_name(&self) -> Option<&str> {
        self.display_name.as_deref()
    }
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }
    pub fn moniker(&self) -> Option<&XlsHyperlinkMoniker> {
        self.moniker.as_ref()
    }
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }
    pub fn hyperlink_guid(&self) -> Option<&[u8; 16]> {
        self.hyperlink_guid.as_ref()
    }
    pub fn creation_time(&self) -> Option<u64> {
        self.creation_time
    }
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }
    pub fn target_kind(&self) -> XlsHyperlinkTargetKind {
        match self.moniker.as_ref() {
            None => XlsHyperlinkTargetKind::Document,
            Some(XlsHyperlinkMoniker::Url(url))
                if starts_ascii_case_insensitive(url.url(), "mailto:") =>
            {
                XlsHyperlinkTargetKind::Email
            },
            Some(XlsHyperlinkMoniker::Url(_)) => XlsHyperlinkTargetKind::Url,
            Some(XlsHyperlinkMoniker::File(file)) if file.is_unc() => XlsHyperlinkTargetKind::Unc,
            Some(XlsHyperlinkMoniker::File(_)) => XlsHyperlinkTargetKind::File,
            Some(XlsHyperlinkMoniker::String(value)) if value.starts_with("\\\\") => {
                XlsHyperlinkTargetKind::Unc
            },
            Some(XlsHyperlinkMoniker::String(value))
                if starts_ascii_case_insensitive(value, "mailto:") =>
            {
                XlsHyperlinkTargetKind::Email
            },
            Some(XlsHyperlinkMoniker::String(_)) => XlsHyperlinkTargetKind::StringMoniker,
            Some(XlsHyperlinkMoniker::Composite(_)) => XlsHyperlinkTargetKind::Composite,
            Some(XlsHyperlinkMoniker::Anti { .. }) => XlsHyperlinkTargetKind::Anti,
            Some(XlsHyperlinkMoniker::Item(_)) => XlsHyperlinkTargetKind::Item,
        }
    }
    /// Serialized base address, without filesystem or network resolution.
    pub fn address(&self) -> Option<&str> {
        match self.moniker.as_ref() {
            Some(XlsHyperlinkMoniker::String(value)) => Some(value),
            Some(XlsHyperlinkMoniker::Url(url)) => Some(url.url()),
            Some(XlsHyperlinkMoniker::File(file)) => Some(file.path()),
            Some(XlsHyperlinkMoniker::Item(item)) => Some(item.item()),
            Some(XlsHyperlinkMoniker::Composite(_) | XlsHyperlinkMoniker::Anti { .. }) => None,
            None => self.location(),
        }
    }
}

#[derive(Debug, Default)]
pub(crate) struct HyperlinkCollector {
    hyperlinks: Vec<XlsHyperlink>,
    pending_tooltip_index: Option<usize>,
}
impl HyperlinkCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if record_type == TOOLTIP_RECORD_TYPE {
            let index = self.pending_tooltip_index.take().ok_or_else(|| {
                XlsError::InvalidData("HLinkTooltip must immediately follow an HLink".to_string())
            })?;
            let (range, tooltip) = parse_tooltip(data)?;
            if self.hyperlinks[index].range != range {
                return invalid(
                    "HLinkTooltip range does not match its preceding HLink".to_string(),
                );
            }
            self.hyperlinks[index].tooltip = Some(tooltip);
            return Ok(());
        }
        self.pending_tooltip_index = None;
        if record_type == RECORD_TYPE {
            self.hyperlinks.push(parse_hlink_record(data)?);
            self.pending_tooltip_index = Some(self.hyperlinks.len() - 1);
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> Vec<XlsHyperlink> {
        self.hyperlinks
    }
}

pub fn parse_hlink_record(data: &[u8]) -> XlsResult<XlsHyperlink> {
    let mut cursor = Cursor::new(data);
    let range = cursor.range()?;
    let class_id = cursor.guid()?;
    if class_id != STANDARD_HLINK_CLSID {
        return invalid("HLink contains an unsupported hyperlink CLSID".to_string());
    }
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
        Some(XlsHyperlinkMoniker::String(cursor.hyperlink_string()?))
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
    Ok(XlsHyperlink {
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

fn parse_tooltip(data: &[u8]) -> XlsResult<(XlsHyperlinkRange, String)> {
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

fn parse_moniker(cursor: &mut Cursor<'_>, depth: usize) -> XlsResult<XlsHyperlinkMoniker> {
    if depth >= 16 {
        return invalid("HLink composite-moniker nesting exceeds 16".to_string());
    }
    let clsid = cursor.guid()?;
    if clsid == URL_MONIKER_CLSID {
        return Ok(XlsHyperlinkMoniker::Url(parse_url_moniker(cursor)?));
    }
    if clsid == FILE_MONIKER_CLSID {
        return Ok(XlsHyperlinkMoniker::File(parse_file_moniker(cursor)?));
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
        return Ok(XlsHyperlinkMoniker::Composite(monikers));
    }
    if clsid == ANTI_MONIKER_CLSID {
        let count = cursor.u32()?;
        if count > 1_048_576 {
            return invalid(format!("anti-moniker count exceeds 1048576: {count}"));
        }
        return Ok(XlsHyperlinkMoniker::Anti { count });
    }
    if clsid == ITEM_MONIKER_CLSID {
        return Ok(XlsHyperlinkMoniker::Item(parse_item_moniker(cursor)?));
    }
    invalid("HLink contains an unknown moniker CLSID".to_string())
}

fn parse_url_moniker(cursor: &mut Cursor<'_>) -> XlsResult<XlsUrlMoniker> {
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
    Ok(XlsUrlMoniker {
        url: decode_terminated_utf16(url_data)?,
        serialization_uri_flags,
    })
}

fn parse_file_moniker(cursor: &mut Cursor<'_>) -> XlsResult<XlsFileMoniker> {
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
    Ok(XlsFileMoniker {
        parent_directory_count,
        ansi_path,
        unicode_path,
        unc_server_character_count,
    })
}

fn parse_item_moniker(cursor: &mut Cursor<'_>) -> XlsResult<XlsItemMoniker> {
    let delimiter_length = cursor.u32()? as usize;
    let (delimiter_ansi, delimiter_unicode) = parse_item_string(cursor.take(delimiter_length)?)?;
    let item_length = cursor.u32()? as usize;
    let (item_ansi, item_unicode) = parse_item_string(cursor.take(item_length)?)?;
    Ok(XlsItemMoniker {
        delimiter_ansi,
        delimiter_unicode,
        item_ansi,
        item_unicode,
    })
}
fn parse_item_string(data: &[u8]) -> XlsResult<(String, Option<String>)> {
    let terminator = data.iter().position(|&byte| byte == 0).ok_or_else(|| {
        XlsError::InvalidData("ItemMoniker ANSI string is not NUL-terminated".to_string())
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

fn decode_terminated_utf16(data: &[u8]) -> XlsResult<String> {
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
        .map_err(|_| XlsError::InvalidData("hyperlink string contains invalid UTF-16".to_string()))
}
fn decode_unterminated_utf16(data: &[u8]) -> XlsResult<String> {
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
        .map_err(|_| XlsError::InvalidData("hyperlink string contains invalid UTF-16".to_string()))
}
fn decode_terminated_ansi(data: &[u8]) -> XlsResult<String> {
    if data.last() != Some(&0) || data[..data.len() - 1].contains(&0) {
        return invalid("FileMoniker ANSI path must contain exactly one trailing NUL".to_string());
    }
    Ok(data[..data.len() - 1]
        .iter()
        .map(|&byte| char::from(byte))
        .collect())
}
fn starts_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|start| start.eq_ignore_ascii_case(prefix))
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
    fn take(&mut self, count: usize) -> XlsResult<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| XlsError::InvalidData("hyperlink field size overflow".to_string()))?;
        let data = self
            .data
            .get(self.position..end)
            .ok_or_else(|| XlsError::InvalidData("truncated hyperlink record".to_string()))?;
        self.position = end;
        Ok(data)
    }
    fn u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn u32(&mut self) -> XlsResult<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }
    fn u64(&mut self) -> XlsResult<u64> {
        Ok(u64::from_le_bytes(self.take(8)?.try_into().unwrap()))
    }
    fn guid(&mut self) -> XlsResult<[u8; 16]> {
        Ok(self.take(16)?.try_into().unwrap())
    }
    fn range(&mut self) -> XlsResult<XlsHyperlinkRange> {
        let first_row = self.u16()?;
        let last_row = self.u16()?;
        let first_column = self.u16()?;
        let last_column = self.u16()?;
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return invalid("hyperlink contains an invalid or out-of-range Ref8U".to_string());
        }
        Ok(XlsHyperlinkRange {
            first_row,
            last_row,
            first_column: first_column as u8,
            last_column: last_column as u8,
        })
    }
    fn hyperlink_string(&mut self) -> XlsResult<String> {
        let units = self.u32()? as usize;
        if units == 0 {
            return invalid("HyperlinkString length must include a NUL terminator".to_string());
        }
        let bytes = units
            .checked_mul(2)
            .ok_or_else(|| XlsError::InvalidData("HyperlinkString size overflow".to_string()))?;
        decode_terminated_utf16(self.take(bytes)?)
    }
}
fn invalid<T>(message: String) -> XlsResult<T> {
    Err(XlsError::InvalidData(message))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn range(data: &mut Vec<u8>, row: u16, column: u16) {
        for value in [row, row, column, column] {
            data.extend_from_slice(&value.to_le_bytes());
        }
    }
    fn string(data: &mut Vec<u8>, value: &str) {
        data.extend_from_slice(&(value.encode_utf16().count() as u32 + 1).to_le_bytes());
        for unit in value.encode_utf16().chain(std::iter::once(0)) {
            data.extend_from_slice(&unit.to_le_bytes());
        }
    }
    fn base(flags: u32) -> Vec<u8> {
        let mut data = Vec::new();
        range(&mut data, 4, 0);
        data.extend_from_slice(&STANDARD_HLINK_CLSID);
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data
    }
    fn url_link() -> Vec<u8> {
        let mut data = base(0x17);
        string(&mut data, "Example");
        data.extend_from_slice(&URL_MONIKER_CLSID);
        let mut url = Vec::new();
        for u in "https://example.com"
            .encode_utf16()
            .chain(std::iter::once(0))
        {
            url.extend_from_slice(&u.to_le_bytes());
        }
        data.extend_from_slice(&(url.len() as u32 + 24).to_le_bytes());
        data.extend_from_slice(&url);
        data.extend_from_slice(&URL_SERIAL_GUID);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0xABA5u32.to_le_bytes());
        data
    }
    #[test]
    fn parses_url_email_document_file_and_string_monikers() {
        let url = parse_hlink_record(&url_link()).unwrap();
        assert_eq!(url.target_kind(), XlsHyperlinkTargetKind::Url);
        assert_eq!(url.address(), Some("https://example.com"));
        let mut document = base(0x1C);
        string(&mut document, "place");
        string(&mut document, "Sheet1!A1");
        let document = parse_hlink_record(&document).unwrap();
        assert_eq!(document.target_kind(), XlsHyperlinkTargetKind::Document);
        assert_eq!(document.location(), Some("Sheet1!A1"));
        let mut file = base(1);
        file.extend_from_slice(&FILE_MONIKER_CLSID);
        file.extend_from_slice(&0u16.to_le_bytes());
        file.extend_from_slice(&9u32.to_le_bytes());
        file.extend_from_slice(b"file.xls\0");
        file.extend_from_slice(&0xFFFFu16.to_le_bytes());
        file.extend_from_slice(&0xDEADu16.to_le_bytes());
        file.extend_from_slice(&[0; 16]);
        file.extend_from_slice(&0u32.to_le_bytes());
        file.extend_from_slice(&0u32.to_le_bytes());
        let file = parse_hlink_record(&file).unwrap();
        assert_eq!(file.target_kind(), XlsHyperlinkTargetKind::File);
        assert_eq!(file.address(), Some("file.xls"));
        let mut unc = base(0x101);
        string(&mut unc, "\\\\server\\share");
        let unc = parse_hlink_record(&unc).unwrap();
        assert_eq!(unc.target_kind(), XlsHyperlinkTargetKind::Unc);
    }
    #[test]
    fn links_only_an_immediately_following_matching_tooltip() {
        let mut collector = HyperlinkCollector::new();
        collector.feed_record(RECORD_TYPE, &url_link()).unwrap();
        let mut tooltip = Vec::new();
        tooltip.extend_from_slice(&TOOLTIP_RECORD_TYPE.to_le_bytes());
        range(&mut tooltip, 4, 0);
        for u in "Open site".encode_utf16().chain(std::iter::once(0)) {
            tooltip.extend_from_slice(&u.to_le_bytes());
        }
        collector
            .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
            .unwrap();
        assert_eq!(collector.finish()[0].tooltip(), Some("Open site"));
        let mut collector = HyperlinkCollector::new();
        assert!(
            collector
                .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
                .is_err()
        );
    }
    #[test]
    fn rejects_bad_guid_flags_range_and_url_tail() {
        let mut data = url_link();
        data[8] ^= 1;
        assert!(parse_hlink_record(&data).is_err());
        let mut data = url_link();
        data[31] = 0x80;
        assert!(parse_hlink_record(&data).is_err());
        let mut data = url_link();
        data[0..2].copy_from_slice(&5u16.to_le_bytes());
        data[2..4].copy_from_slice(&4u16.to_le_bytes());
        assert!(parse_hlink_record(&data).is_err());
        let mut data = url_link();
        let last = data.len() - 24;
        data[last] ^= 1;
        assert!(parse_hlink_record(&data).is_err());
    }
    #[test]
    fn reads_poi_hyperlink_fixtures() {
        use crate::xls::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;
        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet")
                .join(name)
        };
        let workbook =
            XlsWorkbook::new(File::open(fixture("WithTwoHyperLinks.xls")).unwrap()).unwrap();
        let links = workbook.xls_worksheet(0).unwrap().hyperlinks();
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].range().first_row(), 4);
        assert_eq!(links[0].display_name(), Some("Foo"));
        assert_eq!(links[0].address(), Some("http://poi.apache.org/"));
        assert_eq!(links[1].range().first_column(), 1);
        let workbook =
            XlsWorkbook::new(File::open(fixture("HyperlinksOnManySheets.xls")).unwrap()).unwrap();
        assert_eq!(workbook.xls_worksheet(0).unwrap().hyperlinks().len(), 2);
        let email = &workbook.xls_worksheet(1).unwrap().hyperlinks()[0];
        assert_eq!(email.target_kind(), XlsHyperlinkTargetKind::Email);
        assert_eq!(email.address(), Some("mailto:dev@poi.apache.org"));
        let document = &workbook.xls_worksheet(2).unwrap().hyperlinks()[0];
        assert_eq!(document.target_kind(), XlsHyperlinkTargetKind::Document);
        assert_eq!(document.location(), Some("WebLinks!A1"));
    }

    #[test]
    fn test_parse_hlink_record_too_short() {
        assert!(parse_hlink_record(&[0; 10]).is_err());
    }

    #[test]
    fn test_parse_hlink_record_invalid_version() {
        let mut data = base(0x08);
        data[24..28].copy_from_slice(&99u32.to_le_bytes());
        string(&mut data, "Sheet1!A1");
        assert!(parse_hlink_record(&data).is_err());
    }

    #[test]
    fn test_hyperlink_target_url() {
        let link = parse_hlink_record(&url_link()).unwrap();
        assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Url);
        assert_eq!(link.address(), Some("https://example.com"));
        assert!(link.absolute());
    }

    #[test]
    fn test_hyperlink_target_document() {
        let mut data = base(0x08);
        string(&mut data, "Sheet1!A1");
        let link = parse_hlink_record(&data).unwrap();
        assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Document);
        assert_eq!(link.address(), Some("Sheet1!A1"));
    }

    #[test]
    fn test_hyperlink_target_unc() {
        let mut data = base(0x101);
        string(&mut data, "\\\\server\\share\\file.txt");
        let link = parse_hlink_record(&data).unwrap();
        assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Unc);
        assert_eq!(link.address(), Some("\\\\server\\share\\file.txt"));
    }

    #[test]
    fn test_hyperlink_target_file_with_long_name() {
        let mut data = base(0x01);
        data.extend_from_slice(&FILE_MONIKER_CLSID);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&13u32.to_le_bytes());
        data.extend_from_slice(b"LONGFI~1.TXT\0");
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&0xDEADu16.to_le_bytes());
        data.extend_from_slice(&[0; 16]);
        data.extend_from_slice(&0u32.to_le_bytes());
        let unicode: Vec<u8> = "long_filename.txt"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        data.extend_from_slice(&(unicode.len() as u32 + 6).to_le_bytes());
        data.extend_from_slice(&(unicode.len() as u32).to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&unicode);
        let link = parse_hlink_record(&data).unwrap();
        assert_eq!(link.address(), Some("long_filename.txt"));
        let XlsHyperlinkMoniker::File(file) = link.moniker().unwrap() else {
            panic!()
        };
        assert_eq!(file.ansi_path(), "LONGFI~1.TXT");
        assert_eq!(file.unicode_path(), Some("long_filename.txt"));
    }

    #[test]
    fn test_hyperlink_target_file_without_long_name() {
        let mut data = base(0x01);
        data.extend_from_slice(&FILE_MONIKER_CLSID);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&9u32.to_le_bytes());
        data.extend_from_slice(b"FILE.TXT\0");
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&0xDEADu16.to_le_bytes());
        data.extend_from_slice(&[0; 16]);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        let link = parse_hlink_record(&data).unwrap();
        assert_eq!(link.address(), Some("FILE.TXT"));
    }

    #[test]
    fn test_xls_hyperlink_clone() {
        let link = parse_hlink_record(&url_link()).unwrap();
        assert_eq!(link.clone(), link);
        assert_eq!(
            link.moniker().unwrap().clone(),
            link.moniker().unwrap().clone()
        );
    }

    #[test]
    fn test_xls_hyperlink_debug() {
        let link = parse_hlink_record(&url_link()).unwrap();
        let debug = format!("{link:?}");
        assert!(debug.contains("XlsHyperlink"));
        assert!(debug.contains("https://example.com"));
    }

    #[test]
    fn test_record_type_constant() {
        assert_eq!(RECORD_TYPE, 0x01B8);
        assert_eq!(TOOLTIP_RECORD_TYPE, 0x0800);
    }
}
