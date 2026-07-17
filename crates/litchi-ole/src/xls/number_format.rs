use std::collections::HashMap;

use super::error::{XlsError, XlsResult};
use super::records::Record;

const DATE1904_RECORD: u16 = 0x0022;
const FORMAT_RECORD: u16 = 0x041e;
const XF_RECORD: u16 = 0x00e0;
const XFCRC_RECORD: u16 = 0x087c;
const MAX_FORMAT_RECORDS: usize = 218;
const MIN_XF_RECORDS: usize = 16;
const MAX_XF_RECORDS: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum XlsDateSystem {
    #[default]
    Excel1900,
    Excel1904,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsNumberFormat {
    id: u16,
    code: String,
    date_time: bool,
}

impl XlsNumberFormat {
    pub fn id(&self) -> u16 {
        self.id
    }

    pub fn code(&self) -> &str {
        &self.code
    }

    pub fn is_builtin_override(&self) -> bool {
        self.id < 164
    }

    pub fn is_date_time(&self) -> bool {
        self.date_time
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsExtendedFormatKind {
    Cell { parent_style_xf: u16 },
    Style,
}

/// Local-application versus parent-inheritance semantics for the six XF property families.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsExtendedFormatApplications {
    number_format: bool,
    font: bool,
    alignment: bool,
    border: bool,
    fill: bool,
    protection: bool,
}

impl XlsExtendedFormatApplications {
    pub fn applies_number_format(&self) -> bool {
        self.number_format
    }
    pub fn applies_font(&self) -> bool {
        self.font
    }
    pub fn applies_alignment(&self) -> bool {
        self.alignment
    }
    pub fn applies_border(&self) -> bool {
        self.border
    }
    pub fn applies_fill(&self) -> bool {
        self.fill
    }
    pub fn applies_protection(&self) -> bool {
        self.protection
    }
    pub fn inherits_number_format(&self) -> bool {
        !self.number_format
    }
    pub fn inherits_font(&self) -> bool {
        !self.font
    }
    pub fn inherits_alignment(&self) -> bool {
        !self.alignment
    }
    pub fn inherits_border(&self) -> bool {
        !self.border
    }
    pub fn inherits_fill(&self) -> bool {
        !self.fill
    }
    pub fn inherits_protection(&self) -> bool {
        !self.protection
    }

    fn all_local() -> Self {
        Self {
            number_format: true,
            font: true,
            alignment: true,
            border: true,
            fill: true,
            protection: true,
        }
    }

    fn from_cell_bits(bits: u8) -> Self {
        Self {
            number_format: bits & 0x01 != 0,
            font: bits & 0x02 != 0,
            alignment: bits & 0x04 != 0,
            border: bits & 0x08 != 0,
            fill: bits & 0x10 != 0,
            protection: bits & 0x20 != 0,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsExtendedFormat {
    index: u16,
    font_index: u16,
    number_format_id: u16,
    kind: XlsExtendedFormatKind,
    applications: XlsExtendedFormatApplications,
    quote_prefix: bool,
    pivot_button: bool,
    has_xf_extension: bool,
    locked: bool,
    hidden: bool,
    alignment: crate::xls::alignment::XlsCellAlignment,
    borders: crate::xls::border_fill::XlsCellBorders,
    fill: crate::xls::border_fill::XlsCellFill,
}

impl XlsExtendedFormat {
    pub fn index(&self) -> u16 {
        self.index
    }

    pub fn number_format_id(&self) -> u16 {
        self.number_format_id
    }

    /// Returns the logical index of the global Font record used by this XF.
    pub fn font_index(&self) -> u16 {
        self.font_index
    }

    pub fn kind(&self) -> XlsExtendedFormatKind {
        self.kind
    }

    pub fn parent_style_xf_index(&self) -> Option<u16> {
        match self.kind {
            XlsExtendedFormatKind::Cell { parent_style_xf } => Some(parent_style_xf),
            XlsExtendedFormatKind::Style => None,
        }
    }

    pub fn applications(&self) -> XlsExtendedFormatApplications {
        self.applications
    }

    pub fn quote_prefix(&self) -> bool {
        self.quote_prefix
    }
    pub fn pivot_button(&self) -> bool {
        self.pivot_button
    }
    pub fn has_xf_extension(&self) -> bool {
        self.has_xf_extension
    }

    pub fn is_cell_format(&self) -> bool {
        matches!(self.kind, XlsExtendedFormatKind::Cell { .. })
    }

    pub fn locked(&self) -> bool {
        self.locked
    }

    pub fn hidden(&self) -> bool {
        self.hidden
    }

    pub fn alignment(&self) -> &crate::xls::alignment::XlsCellAlignment {
        &self.alignment
    }

    /// Returns the border metadata stored by this XF record.
    pub fn borders(&self) -> &crate::xls::border_fill::XlsCellBorders {
        &self.borders
    }

    /// Returns the fill pattern and colors stored by this XF record.
    pub fn fill(&self) -> &crate::xls::border_fill::XlsCellFill {
        &self.fill
    }
}

/// Borrowed effective formatting after applying a CellXF's parent StyleXF.
#[derive(Debug, Clone, Copy)]
pub struct XlsEffectiveExtendedFormat<'a> {
    direct: &'a XlsExtendedFormat,
    parent: Option<&'a XlsExtendedFormat>,
}

impl<'a> XlsEffectiveExtendedFormat<'a> {
    pub fn direct(&self) -> &'a XlsExtendedFormat {
        self.direct
    }
    pub fn parent_style(&self) -> Option<&'a XlsExtendedFormat> {
        self.parent
    }

    fn source(&self, local: bool) -> &'a XlsExtendedFormat {
        if local {
            self.direct
        } else {
            self.parent.unwrap_or(self.direct)
        }
    }

    pub fn number_format_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_number_format())
    }
    pub fn font_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_font())
    }
    pub fn alignment_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_alignment())
    }
    pub fn border_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_border())
    }
    pub fn fill_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_fill())
    }
    pub fn protection_source(&self) -> &'a XlsExtendedFormat {
        self.source(self.direct.applications.applies_protection())
    }

    pub fn number_format_id(&self) -> u16 {
        self.number_format_source().number_format_id
    }
    pub fn font_index(&self) -> u16 {
        self.font_source().font_index
    }
    pub fn alignment(&self) -> &'a crate::xls::alignment::XlsCellAlignment {
        &self.alignment_source().alignment
    }
    pub fn borders(&self) -> &'a crate::xls::border_fill::XlsCellBorders {
        &self.border_source().borders
    }
    pub fn fill(&self) -> &'a crate::xls::border_fill::XlsCellFill {
        &self.fill_source().fill
    }
    pub fn locked(&self) -> bool {
        self.protection_source().locked
    }
    pub fn hidden(&self) -> bool {
        self.protection_source().hidden
    }
    pub fn quote_prefix(&self) -> bool {
        self.direct.quote_prefix
    }
    pub fn pivot_button(&self) -> bool {
        self.direct.pivot_button
    }
    pub fn has_xf_extension(&self) -> bool {
        self.direct.has_xf_extension
    }
}

#[derive(Debug, Clone, Default)]
pub struct XlsFormatting {
    date_system: XlsDateSystem,
    number_formats: Vec<XlsNumberFormat>,
    extended_formats: Vec<XlsExtendedFormat>,
    format_by_id: HashMap<u16, usize>,
}

impl XlsFormatting {
    pub fn date_system(&self) -> XlsDateSystem {
        self.date_system
    }

    /// Explicit BIFF `Format` records in their original workbook order.
    pub fn number_formats(&self) -> &[XlsNumberFormat] {
        &self.number_formats
    }

    /// BIFF `XF` records in index order, including style-XF slots.
    pub fn extended_formats(&self) -> &[XlsExtendedFormat] {
        &self.extended_formats
    }

    pub fn number_format(&self, id: u16) -> Option<&XlsNumberFormat> {
        self.format_by_id
            .get(&id)
            .and_then(|index| self.number_formats.get(*index))
    }

    pub fn extended_format(&self, index: u16) -> Option<&XlsExtendedFormat> {
        self.extended_formats.get(index as usize)
    }

    pub fn effective_extended_format(&self, index: u16) -> Option<XlsEffectiveExtendedFormat<'_>> {
        let direct = self.extended_format(index)?;
        let parent = direct
            .parent_style_xf_index()
            .and_then(|parent| self.extended_format(parent));
        Some(XlsEffectiveExtendedFormat { direct, parent })
    }

    pub fn is_date_time_format(&self, id: u16) -> bool {
        self.number_format(id)
            .map(XlsNumberFormat::is_date_time)
            .unwrap_or_else(|| is_builtin_date_time(id))
    }

    pub(crate) fn cell_format(&self, index: u16) -> Option<&XlsExtendedFormat> {
        if index == 0 {
            return None;
        }
        self.extended_format(index)
            .filter(|format| format.is_cell_format())
    }

    pub(crate) fn validate_cell_xf(&self, index: u16) -> XlsResult<()> {
        if self.extended_formats.is_empty() {
            return Ok(());
        }
        if index == 0 {
            return Ok(());
        }
        if index < 15 {
            return Err(invalid(
                XF_RECORD,
                format!("cell references reserved style-XF slot {index}"),
            ));
        }
        match self.extended_format(index) {
            Some(format) if format.is_cell_format() => Ok(()),
            Some(_) => Err(invalid(
                XF_RECORD,
                format!("cell references style XF {index}"),
            )),
            None => Err(invalid(
                XF_RECORD,
                format!("cell references out-of-range XF {index}"),
            )),
        }
    }

    pub(crate) fn parse_globals(records: &[Record]) -> XlsResult<Self> {
        let mut date_system = None;
        let mut number_formats = Vec::new();
        let mut extended_formats = Vec::new();
        let mut format_by_id = HashMap::new();
        let mut xfcrc = None;

        for record in records {
            match record.header.record_type {
                DATE1904_RECORD => {
                    if date_system.is_some() {
                        return Err(invalid(DATE1904_RECORD, "duplicate Date1904 record"));
                    }
                    date_system = Some(parse_date_system(&record.data)?);
                },
                FORMAT_RECORD => {
                    if number_formats.len() == MAX_FORMAT_RECORDS {
                        return Err(invalid(FORMAT_RECORD, "more than 218 Format records"));
                    }
                    let format = parse_number_format(&record.data)?;
                    if format_by_id.contains_key(&format.id) {
                        return Err(invalid(
                            FORMAT_RECORD,
                            format!("duplicate number format identifier {}", format.id),
                        ));
                    }
                    format_by_id.insert(format.id, number_formats.len());
                    number_formats.push(format);
                },
                XF_RECORD => {
                    if extended_formats.len() == MAX_XF_RECORDS {
                        return Err(invalid(XF_RECORD, "more than 65,536 XF records"));
                    }
                    let index = u16::try_from(extended_formats.len()).map_err(|_| {
                        invalid(XF_RECORD, "XF index does not fit in the BIFF index field")
                    })?;
                    extended_formats.push(parse_xf(&record.data, index)?);
                },
                XFCRC_RECORD => {
                    if xfcrc.is_some() {
                        return Err(invalid(XFCRC_RECORD, "duplicate XFCRC record"));
                    }
                    xfcrc = Some(parse_xfcrc(&record.data)?);
                },
                _ => {},
            }
        }

        let date_system = date_system
            .ok_or_else(|| invalid(DATE1904_RECORD, "missing required Date1904 record"))?;
        if extended_formats.len() < MIN_XF_RECORDS {
            return Err(invalid(
                XF_RECORD,
                format!(
                    "workbook has {} XF records; expected at least 16",
                    extended_formats.len()
                ),
            ));
        }
        if let Some(count) = xfcrc {
            if count as usize != extended_formats.len() {
                return Err(invalid(
                    XFCRC_RECORD,
                    format!(
                        "XFCRC declares {count} XF records but {} were parsed",
                        extended_formats.len()
                    ),
                ));
            }
        }

        for (index, format) in extended_formats.iter().enumerate() {
            if index < 15 && !matches!(format.kind, XlsExtendedFormatKind::Style) {
                return Err(invalid(
                    XF_RECORD,
                    format!("mandatory XF slot {index} is not a style XF"),
                ));
            }
            if index == 15 && !format.is_cell_format() {
                return Err(invalid(XF_RECORD, "mandatory XF slot 15 is not a cell XF"));
            }
            if format.number_format_id > 81 {
                if !(164..=392).contains(&format.number_format_id) {
                    return Err(invalid(
                        XF_RECORD,
                        format!(
                            "XF {index} uses reserved number format identifier {}",
                            format.number_format_id
                        ),
                    ));
                }
                if !format_by_id.contains_key(&format.number_format_id) {
                    return Err(invalid(
                        XF_RECORD,
                        format!(
                            "XF {index} references missing custom number format {}",
                            format.number_format_id
                        ),
                    ));
                }
            }
            if let XlsExtendedFormatKind::Cell { parent_style_xf } = format.kind {
                let parent = extended_formats
                    .get(parent_style_xf as usize)
                    .ok_or_else(|| {
                        invalid(
                            XF_RECORD,
                            format!("XF {index} has out-of-range parent {parent_style_xf}"),
                        )
                    })?;
                if !matches!(parent.kind, XlsExtendedFormatKind::Style) {
                    return Err(invalid(
                        XF_RECORD,
                        format!("XF {index} parent {parent_style_xf} is not a style XF"),
                    ));
                }
            }
        }

        Ok(Self {
            date_system,
            number_formats,
            extended_formats,
            format_by_id,
        })
    }
}

fn parse_date_system(data: &[u8]) -> XlsResult<XlsDateSystem> {
    if data.len() != 2 {
        return Err(invalid(
            DATE1904_RECORD,
            format!("Date1904 payload has {} bytes; expected 2", data.len()),
        ));
    }
    match u16::from_le_bytes([data[0], data[1]]) {
        0 => Ok(XlsDateSystem::Excel1900),
        1 => Ok(XlsDateSystem::Excel1904),
        value => Err(invalid(
            DATE1904_RECORD,
            format!("Date1904 Boolean is {value}; expected 0 or 1"),
        )),
    }
}

fn parse_number_format(data: &[u8]) -> XlsResult<XlsNumberFormat> {
    if data.len() < 5 {
        return Err(invalid(FORMAT_RECORD, "truncated Format record"));
    }
    let id = u16::from_le_bytes([data[0], data[1]]);
    if !matches!(id, 5..=8 | 23..=26 | 41..=44 | 63..=66 | 164..=392) {
        return Err(invalid(
            FORMAT_RECORD,
            format!("number format identifier {id} is outside the permitted ranges"),
        ));
    }
    let code = parse_xl_unicode_string(&data[2..])?;
    let count = code.encode_utf16().count();
    if !(1..=255).contains(&count) {
        return Err(invalid(
            FORMAT_RECORD,
            format!("format string has {count} UTF-16 code units; expected 1 through 255"),
        ));
    }
    Ok(XlsNumberFormat {
        id,
        date_time: is_custom_date_time(&code),
        code,
    })
}

fn parse_xl_unicode_string(data: &[u8]) -> XlsResult<String> {
    if data.len() < 3 {
        return Err(invalid(FORMAT_RECORD, "truncated XLUnicodeString header"));
    }
    let cch = u16::from_le_bytes([data[0], data[1]]) as usize;
    let flags = data[2];
    if flags & !0x0d != 0 {
        return Err(invalid(
            FORMAT_RECORD,
            format!("XLUnicodeString has reserved option bits 0x{flags:02x}"),
        ));
    }
    let mut offset = 3usize;
    let rich_runs = if flags & 0x08 != 0 {
        let bytes = data
            .get(offset..offset + 2)
            .ok_or_else(|| invalid(FORMAT_RECORD, "truncated rich-run count"))?;
        offset += 2;
        u16::from_le_bytes([bytes[0], bytes[1]]) as usize
    } else {
        0
    };
    let extension_len = if flags & 0x04 != 0 {
        let bytes = data
            .get(offset..offset + 4)
            .ok_or_else(|| invalid(FORMAT_RECORD, "truncated extension length"))?;
        offset += 4;
        u32::from_le_bytes(bytes.try_into().unwrap()) as usize
    } else {
        0
    };
    let char_width = if flags & 0x01 != 0 { 2usize } else { 1usize };
    let char_bytes = cch
        .checked_mul(char_width)
        .ok_or_else(|| invalid(FORMAT_RECORD, "format string length overflow"))?;
    let chars = data
        .get(offset..offset + char_bytes)
        .ok_or_else(|| invalid(FORMAT_RECORD, "truncated format string characters"))?;
    offset += char_bytes;
    let trailing = rich_runs
        .checked_mul(4)
        .and_then(|value| value.checked_add(extension_len))
        .ok_or_else(|| invalid(FORMAT_RECORD, "format string optional-data overflow"))?;
    if offset.checked_add(trailing) != Some(data.len()) {
        return Err(invalid(
            FORMAT_RECORD,
            "XLUnicodeString optional data is truncated or has trailing bytes",
        ));
    }
    if char_width == 1 {
        Ok(chars.iter().map(|byte| char::from(*byte)).collect())
    } else {
        let units = chars
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        char::decode_utf16(units)
            .collect::<Result<String, _>>()
            .map_err(|_| invalid(FORMAT_RECORD, "format string contains invalid UTF-16"))
    }
}

fn parse_xf(data: &[u8], index: u16) -> XlsResult<XlsExtendedFormat> {
    if data.len() != 20 {
        return Err(invalid(
            XF_RECORD,
            format!("XF payload has {} bytes; expected 20", data.len()),
        ));
    }
    let number_format_id = u16::from_le_bytes([data[2], data[3]]);
    let flags = u16::from_le_bytes([data[4], data[5]]);
    let locked = flags & 0x0001 != 0;
    let hidden = flags & 0x0002 != 0;
    let style = flags & 0x0004 != 0;
    let prefix = flags & 0x0008 != 0;
    let parent = flags >> 4;
    if style && prefix {
        return Err(invalid(XF_RECORD, "style XF has the f123Prefix bit set"));
    }
    let rotation = data[7];
    if rotation > 180 && rotation != 255 {
        return Err(invalid(
            XF_RECORD,
            format!("XF text rotation {rotation} is outside the permitted range"),
        ));
    }
    let kind = if style {
        XlsExtendedFormatKind::Style
    } else {
        XlsExtendedFormatKind::Cell {
            parent_style_xf: parent,
        }
    };
    let font_index = u16::from_le_bytes([data[0], data[1]]);
    crate::xls::font::validate_font_index(font_index)?;
    let applications = if style {
        XlsExtendedFormatApplications::all_local()
    } else {
        XlsExtendedFormatApplications::from_cell_bits(data[9] >> 2)
    };
    let border2 = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let area = u16::from_le_bytes([data[18], data[19]]);
    let quote_prefix = !style && prefix;
    let has_xf_extension = !style && border2 & (1 << 25) != 0;
    let pivot_button = !style && area & (1 << 14) != 0;
    let alignment = crate::xls::alignment::XlsCellAlignment::parse(data[6], data[7], data[8])?;
    let (borders, fill) = crate::xls::border_fill::parse_xf_border_fill(data, style)?;

    Ok(XlsExtendedFormat {
        index,
        font_index,
        number_format_id,
        kind,
        applications,
        quote_prefix,
        pivot_button,
        has_xf_extension,
        locked,
        hidden,
        alignment,
        borders,
        fill,
    })
}

fn parse_xfcrc(data: &[u8]) -> XlsResult<u16> {
    if data.len() != 20 {
        return Err(invalid(
            XFCRC_RECORD,
            format!("XFCRC payload has {} bytes; expected 20", data.len()),
        ));
    }
    if u16::from_le_bytes([data[0], data[1]]) != XFCRC_RECORD {
        return Err(invalid(
            XFCRC_RECORD,
            "XFCRC FrtHeader has the wrong record type",
        ));
    }
    if data[2..14].iter().any(|byte| *byte != 0) {
        return Err(invalid(XFCRC_RECORD, "XFCRC reserved fields are nonzero"));
    }
    let count = u16::from_le_bytes([data[14], data[15]]);
    if !(16..=4050).contains(&count) {
        return Err(invalid(
            XFCRC_RECORD,
            format!("XFCRC count {count} is outside 16 through 4050"),
        ));
    }
    Ok(count)
}

fn is_builtin_date_time(id: u16) -> bool {
    matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58)
}

fn is_custom_date_time(code: &str) -> bool {
    let chars: Vec<char> = code.chars().collect();
    let mut index = 0usize;
    let mut token = false;
    while index < chars.len() {
        match chars[index] {
            '"' => {
                index += 1;
                while index < chars.len() && chars[index] != '"' {
                    index += 1;
                }
                if index == chars.len() {
                    return false;
                }
                index += 1;
            },
            '\\' | '_' | '*' => {
                index = index.saturating_add(2);
            },
            '[' => {
                let start = index + 1;
                index = start;
                while index < chars.len() && chars[index] != ']' {
                    index += 1;
                }
                if index == chars.len() {
                    return false;
                }
                let inner: String = chars[start..index]
                    .iter()
                    .collect::<String>()
                    .to_ascii_lowercase();
                if matches!(inner.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
                    token = true;
                }
                index += 1;
            },
            ch if matches!(ch.to_ascii_lowercase(), 'y' | 'm' | 'd' | 'h' | 's') => {
                token = true;
                index += 1;
            },
            _ => index += 1,
        }
    }
    token
}

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn semantic_xf(style: bool, parent: u16, application_bits: u8) -> [u8; 20] {
        let mut data = [0; 20];
        let flags = (parent << 4) | if style { 0x0004 } else { 0 };
        data[4..6].copy_from_slice(&flags.to_le_bytes());
        data[9] = application_bits << 2;
        data
    }

    #[test]
    fn decodes_cell_apply_polarity_and_style_local_semantics() {
        let cell = parse_xf(&semantic_xf(false, 0, 0b10_1010), 1).unwrap();
        let apply = cell.applications();
        assert!(apply.inherits_number_format());
        assert!(apply.applies_font());
        assert!(apply.inherits_alignment());
        assert!(apply.applies_border());
        assert!(apply.inherits_fill());
        assert!(apply.applies_protection());

        let style = parse_xf(&semantic_xf(true, 0x0fff, 0), 0).unwrap();
        let apply = style.applications();
        assert!(apply.applies_number_format());
        assert!(apply.applies_font());
        assert!(apply.applies_alignment());
        assert!(apply.applies_border());
        assert!(apply.applies_fill());
        assert!(apply.applies_protection());
    }

    #[test]
    fn decodes_cell_special_flags_and_ignores_style_reserved_overlays() {
        let mut cell_data = semantic_xf(false, 0, 0);
        cell_data[4] |= 0x08;
        let mut border2 = 1u32 << 25;
        cell_data[14..18].copy_from_slice(&border2.to_le_bytes());
        cell_data[18..20].copy_from_slice(&(1u16 << 14).to_le_bytes());
        let cell = parse_xf(&cell_data, 1).unwrap();
        assert!(cell.quote_prefix());
        assert!(cell.has_xf_extension());
        assert!(cell.pivot_button());

        let mut style_data = semantic_xf(true, 0x0fff, 0x3f);
        border2 |= 1 << 25;
        style_data[14..18].copy_from_slice(&border2.to_le_bytes());
        style_data[18..20].copy_from_slice(&(3u16 << 14).to_le_bytes());
        let style = parse_xf(&style_data, 0).unwrap();
        assert!(!style.has_xf_extension());
        assert!(!style.pivot_button());
        assert!(style.applications().applies_fill());
    }

    #[test]
    fn resolves_effective_components_by_borrowing_parent_or_cell() {
        let mut style_data = semantic_xf(true, 0x0fff, 0);
        style_data[2..4].copy_from_slice(&14u16.to_le_bytes());
        let style = parse_xf(&style_data, 0).unwrap();
        let mut cell_data = semantic_xf(false, 0, 0b00_0010);
        cell_data[0..2].copy_from_slice(&1u16.to_le_bytes());
        cell_data[2..4].copy_from_slice(&1u16.to_le_bytes());
        let cell = parse_xf(&cell_data, 1).unwrap();
        let formatting = XlsFormatting {
            extended_formats: vec![style, cell],
            ..XlsFormatting::default()
        };

        let effective = formatting.effective_extended_format(1).unwrap();
        assert_eq!(effective.parent_style().unwrap().index(), 0);
        assert_eq!(effective.number_format_id(), 14);
        assert_eq!(effective.number_format_source().index(), 0);
        assert_eq!(effective.font_index(), 1);
        assert_eq!(effective.font_source().index(), 1);
        assert!(std::ptr::eq(
            effective.alignment(),
            formatting.extended_formats()[0].alignment(),
        ));
    }
    use crate::xls::cell::XlsCell;
    use crate::xls::records::{CellRecord, RecordHeader};
    use crate::xls::workbook::XlsWorkbook;
    use litchi_core::sheet::{Cell, CellValue, Worksheet};
    use std::fs::{self, File};
    use std::io::Cursor;
    use std::path::{Path, PathBuf};

    fn record(record_type: u16, data: Vec<u8>) -> Record {
        Record {
            header: RecordHeader {
                record_type,
                data_len: data.len() as u16,
            },
            data,
        }
    }

    fn xf(style: bool, parent: u16, format_id: u16) -> Record {
        let mut data = vec![0u8; 20];
        data[2..4].copy_from_slice(&format_id.to_le_bytes());
        let flags = (parent << 4) | u16::from(style) << 2 | 1;
        data[4..6].copy_from_slice(&flags.to_le_bytes());
        record(XF_RECORD, data)
    }

    fn format_record(id: u16, code: &str) -> Record {
        let mut data = Vec::new();
        data.extend_from_slice(&id.to_le_bytes());
        data.extend_from_slice(&(code.len() as u16).to_le_bytes());
        data.push(0);
        data.extend_from_slice(code.as_bytes());
        record(FORMAT_RECORD, data)
    }

    fn fixture(relative: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative)
    }

    #[test]
    fn recognizes_date_formats_without_literal_false_positives() {
        for code in ["yyyy-mm-dd", "h:mm AM/PM", "[hh]:mm:ss", "dd mmmm"] {
            assert!(is_custom_date_time(code), "{code}");
        }
        for code in ["0.00E+00", "0.00 \"m\"", "0.0\\m", "[Red]0.00", "General"] {
            assert!(!is_custom_date_time(code), "{code}");
        }
    }

    #[test]
    fn parses_strict_unicode_format_records() {
        let mut data = vec![164, 0, 4, 0, 1];
        for unit in "日期mm".encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        let format = parse_number_format(&data).unwrap();
        assert_eq!(format.id(), 164);
        assert_eq!(format.code(), "日期mm");
        assert!(format.is_date_time());

        data.push(0);
        assert!(parse_number_format(&data).is_err());
    }

    #[test]
    fn rejects_invalid_date_xf_and_crc_shapes() {
        assert!(parse_date_system(&[2, 0]).is_err());
        assert!(parse_date_system(&[0]).is_err());
        assert!(parse_xf(&[0; 19], 0).is_err());

        let mut crc = [0u8; 20];
        crc[0..2].copy_from_slice(&XFCRC_RECORD.to_le_bytes());
        crc[14..16].copy_from_slice(&16u16.to_le_bytes());
        assert_eq!(parse_xfcrc(&crc).unwrap(), 16);
        crc[2] = 1;
        assert!(parse_xfcrc(&crc).is_err());
    }

    #[test]
    fn classifies_numeric_and_formula_caches_without_literal_false_positives() {
        let mut records = vec![record(DATE1904_RECORD, vec![0, 0])];
        records.push(format_record(164, "yyyy-mm-dd"));
        records.push(format_record(165, "0.00 \"m\""));
        for _ in 0..15 {
            records.push(xf(true, 0x0fff, 0));
        }
        records.push(xf(false, 0, 14));
        records.push(xf(false, 0, 164));
        records.push(xf(false, 0, 165));
        let formatting = XlsFormatting::parse_globals(&records).unwrap();

        let builtin = CellRecord::Number {
            row: 0,
            col: 0,
            xf_index: 15,
            value: 39_304.0,
        };
        let custom = CellRecord::Formula {
            row: 0,
            col: 1,
            xf_index: 16,
            value: crate::xls::records::FormulaValue::Number(45_000.5),
            formula: vec![0x1e, 1, 0],
        };
        let literal = CellRecord::Number {
            row: 0,
            col: 2,
            xf_index: 17,
            value: 12.5,
        };

        let builtin =
            XlsCell::from_record_with_formula_context(&builtin, None, None, Some(&formatting))
                .unwrap();
        let custom =
            XlsCell::from_record_with_formula_context(&custom, None, None, Some(&formatting))
                .unwrap();
        let literal =
            XlsCell::from_record_with_formula_context(&literal, None, None, Some(&formatting))
                .unwrap();
        assert_eq!(builtin.value(), &CellValue::DateTime(39_304.0));
        assert_eq!(custom.value(), &CellValue::DateTime(45_000.5));
        assert_eq!(literal.value(), &CellValue::Float(12.5));
        assert_eq!(literal.xf_index(), 17);
    }

    #[test]
    fn opens_poi_date_and_epoch_fixtures() {
        let dates = XlsWorkbook::new(
            File::open(fixture(
                "3rdparty/poi/test-data/spreadsheet/DateFormats.xls",
            ))
            .unwrap(),
        )
        .unwrap();
        let sheet = dates.xls_worksheet(0).unwrap();
        let mut found = false;
        for row in 0..sheet.row_count() as u32 {
            for col in 0..sheet.column_count() as u32 {
                if matches!(
                    sheet.get_cell(row, col).map(Cell::value),
                    Some(CellValue::DateTime(_))
                ) {
                    found = true;
                }
            }
        }
        assert!(
            found,
            "POI DateFormats.xls contains no classified date cells"
        );

        for (path, expected) in [
            (
                "3rdparty/poi/test-data/spreadsheet/1900DateWindowing.xls",
                XlsDateSystem::Excel1900,
            ),
            (
                "3rdparty/poi/test-data/spreadsheet/1904DateWindowing.xls",
                XlsDateSystem::Excel1904,
            ),
        ] {
            let mut bytes = fs::read(fixture(path)).unwrap();
            assert!(XlsWorkbook::new(Cursor::new(bytes.clone())).is_err());

            // Both POI windowing fixtures declare sector 10 as their sole FAT
            // sector but mark FAT[10] ENDOFCHAIN rather than FATSECT. Normalize
            // that one proven MS-CFB defect in the test copy only; production
            // container validation remains strict.
            let sector_shift = u16::from_le_bytes([bytes[0x1e], bytes[0x1f]]);
            let sector_size = 1usize << sector_shift;
            let fat_sector = u32::from_le_bytes(bytes[0x4c..0x50].try_into().unwrap()) as usize;
            assert_eq!(u32::from_le_bytes(bytes[0x2c..0x30].try_into().unwrap()), 1);
            let marker_offset = (fat_sector + 1) * sector_size + fat_sector * 4;
            assert_eq!(
                u32::from_le_bytes(bytes[marker_offset..marker_offset + 4].try_into().unwrap()),
                0xffff_fffe
            );
            bytes[marker_offset..marker_offset + 4].copy_from_slice(&0xffff_fffdu32.to_le_bytes());

            assert_eq!(u32::from_le_bytes(bytes[0x30..0x34].try_into().unwrap()), 0);
            for sid in 5..=7usize {
                let entry = sector_size + sid * 128;
                assert_eq!(bytes[entry + 66], 0);
                assert_eq!(
                    u16::from_le_bytes(bytes[entry + 64..entry + 66].try_into().unwrap()),
                    2
                );
                assert_eq!(&bytes[entry..entry + 2], &[0, 0]);
                bytes[entry + 64..entry + 66].copy_from_slice(&0u16.to_le_bytes());
            }

            let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
            assert_eq!(workbook.date_system(), expected, "{path}");
        }
    }

    #[test]
    fn opens_libreoffice_format_fixtures_with_ordered_xfs() {
        for path in [
            "3rdparty/libreoffice-core/sc/qa/unit/data/xls/formats.xls",
            "3rdparty/libreoffice-core/sc/qa/unit/data/xls/cellformat.xls",
        ] {
            let workbook = XlsWorkbook::new(File::open(fixture(path)).unwrap()).unwrap();
            assert!(workbook.extended_formats().len() >= 16, "{path}");
            for (index, format) in workbook.extended_formats().iter().enumerate() {
                assert_eq!(format.index() as usize, index, "{path}");
            }
        }
    }
}
