use std::collections::HashMap;

use super::error::{XlsError, XlsResult};
use super::leniency::{XlsFormattingDefect, XlsToleranceLog};
use super::records::Record;

const DATE1904_RECORD: u16 = 0x0022;
/// MS-XLS 2.4.126 `Format` record type.
pub(crate) const FORMAT_RECORD: u16 = 0x041e;
/// MS-XLS 2.4.353 `XF` record type.
pub(crate) const XF_RECORD: u16 = 0x00e0;
/// MS-XLS 2.4.354 `XFCRC` record type.
pub(crate) const XFCRC_RECORD: u16 = 0x087c;
const MAX_DXF_RECORDS: usize = 65_536;
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
    differential_formats: Vec<crate::xls::differential_format::XlsDifferentialFormat>,
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

    /// Global `DXF` records in zero-based reference order.
    pub fn differential_formats(
        &self,
    ) -> &[crate::xls::differential_format::XlsDifferentialFormat] {
        &self.differential_formats
    }

    pub fn differential_format(
        &self,
        id: crate::xls::table_styles::XlsDifferentialFormatId,
    ) -> Option<&crate::xls::differential_format::XlsDifferentialFormat> {
        self.differential_formats.get(id.index() as usize)
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

    /// Parse the formatting records of the workbook globals under `tolerance`.
    ///
    /// See [`XlsFormattingDefect`] for the exhaustive list of defects a lenient
    /// policy repairs here; every other validation is unchanged.
    pub(crate) fn parse_globals(
        records: &[Record],
        tolerance: &mut XlsToleranceLog,
    ) -> XlsResult<Self> {
        let mut date_system = None;
        let mut number_formats = Vec::new();
        let mut extended_formats = Vec::new();
        let mut differential_formats = Vec::new();
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
                    let ordinal = u32::try_from(number_formats.len()).map_err(|_| {
                        invalid(FORMAT_RECORD, "Format record ordinal does not fit in u32")
                    })?;
                    let format = parse_number_format(&record.data, ordinal, tolerance)?;
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
                    extended_formats.push(parse_xf(&record.data, index, tolerance)?);
                },
                XFCRC_RECORD => {
                    if xfcrc.is_some() {
                        return Err(invalid(XFCRC_RECORD, "duplicate XFCRC record"));
                    }
                    xfcrc = Some(parse_xfcrc(&record.data)?);
                },
                crate::xls::differential_format::DXF_RECORD_TYPE => {
                    if differential_formats.len() == MAX_DXF_RECORDS {
                        return Err(invalid(
                            crate::xls::differential_format::DXF_RECORD_TYPE,
                            "more than 65,536 DXF records",
                        ));
                    }
                    differential_formats.push(
                        crate::xls::differential_format::XlsDifferentialFormat::parse_payload(
                            &record.data,
                        )?,
                    );
                },
                _ => {},
            }
        }

        // MS-XLS 2.1 lists Date1904 in the globals grammar, but workbooks
        // written by other producers omit it and Excel still opens them. The
        // record only selects between two date systems and the absent value is
        // unambiguous, so fall back to the 1900 system rather than rejecting an
        // otherwise readable workbook.
        let date_system = date_system.unwrap_or(XlsDateSystem::Excel1900);
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
            let parsed = extended_formats.len();
            if count as usize != parsed {
                // XFCRC is a redundant integrity summary over the XF table.
                // The XF records themselves remain individually well-formed, so
                // a disagreeing cxfs costs no cell data.
                tolerance.tolerate(
                    XlsFormattingDefect::ExtendedFormatCountMismatch,
                    u32::try_from(parsed).unwrap_or(u32::MAX),
                    u32::from(count),
                    || {
                        invalid(
                            XFCRC_RECORD,
                            format!("XFCRC declares {count} XF records but {parsed} were parsed"),
                        )
                    },
                )?;
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

        for (dxf_index, dxf) in differential_formats.iter().enumerate() {
            for property in dxf.properties().properties() {
                if let crate::xls::differential_format::XlsXfProperty::NumberFormatId(id) = property
                {
                    if *id > 81 {
                        if !(164..=392).contains(id) {
                            return Err(invalid(
                                crate::xls::differential_format::DXF_RECORD_TYPE,
                                format!(
                                    "DXF {dxf_index} uses reserved number format identifier {id}"
                                ),
                            ));
                        }
                        if !format_by_id.contains_key(id) {
                            return Err(invalid(
                                crate::xls::differential_format::DXF_RECORD_TYPE,
                                format!(
                                    "DXF {dxf_index} references missing custom number format {id}"
                                ),
                            ));
                        }
                    }
                }
            }
        }

        Ok(Self {
            date_system,
            number_formats,
            extended_formats,
            differential_formats,
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

/// Smallest number of UTF-16 code units MS-XLS 2.4.126 permits in a format code.
const MIN_FORMAT_CODE_UNITS: usize = 1;
/// Largest number of UTF-16 code units MS-XLS 2.4.126 permits in a format code.
const MAX_FORMAT_CODE_UNITS: usize = 255;

fn parse_number_format(
    data: &[u8],
    ordinal: u32,
    tolerance: &mut XlsToleranceLog,
) -> XlsResult<XlsNumberFormat> {
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
    let (code, truncated) = parse_xl_unicode_string(&data[2..], ordinal, tolerance)?;
    let count = code.encode_utf16().count();
    // A truncated code was already recorded as `FormatStringOverrun`; its
    // surviving length is a consequence of the repair, not a separate defect.
    if !truncated && !(MIN_FORMAT_CODE_UNITS..=MAX_FORMAT_CODE_UNITS).contains(&count) {
        return Err(invalid(
            FORMAT_RECORD,
            format!(
                "format string has {count} UTF-16 code units; expected \
                 {MIN_FORMAT_CODE_UNITS} through {MAX_FORMAT_CODE_UNITS}"
            ),
        ));
    }
    Ok(XlsNumberFormat {
        id,
        date_time: is_custom_date_time(&code),
        code,
    })
}

/// `fHighByte`: the characters in `rgb` are UTF-16 rather than compressed.
pub(crate) const XL_UNICODE_STRING_HIGH_BYTE: u8 = 0x01;
/// Bytes per character when `fHighByte` is set.
pub(crate) const UTF16_CHAR_BYTES: usize = 2;
/// Bytes per character in a compressed (single-byte) string.
pub(crate) const COMPRESSED_CHAR_BYTES: usize = 1;
/// `cch` plus the option byte that precede an XLUnicodeString's characters.
const XL_UNICODE_STRING_HEADER_LEN: usize = 3;
/// Offset of the option byte within an XLUnicodeString.
const XL_UNICODE_STRING_FLAGS_OFFSET: usize = 2;

/// Decode a `Format` record's `XLUnicodeString`.
///
/// Returns the decoded code and whether the payload was shorter than `cch`
/// declared. A short payload is only reachable when `tolerance` permits
/// [`XlsFormattingDefect::FormatStringOverrun`]; a *longer* payload stays a hard
/// error, because trailing bytes mean the record framing is wrong rather than
/// merely the count.
fn parse_xl_unicode_string(
    data: &[u8],
    ordinal: u32,
    tolerance: &mut XlsToleranceLog,
) -> XlsResult<(String, bool)> {
    if data.len() < XL_UNICODE_STRING_HEADER_LEN {
        return Err(invalid(FORMAT_RECORD, "truncated XLUnicodeString header"));
    }
    let cch = u16::from_le_bytes([data[0], data[1]]) as usize;
    // MS-XLS 2.5.294: an XLUnicodeString is `cch`, one option byte, and `rgb`.
    // Only `fHighByte` is defined; every other bit is reserved and "MUST be
    // zero, and MUST be ignored". Unlike XLUnicodeRichExtendedString (2.5.293)
    // there are no `cRun`/`cbExtRst` fields, so the option byte never changes
    // the layout beyond the character width.
    let flags = data[XL_UNICODE_STRING_FLAGS_OFFSET];
    let char_width = if flags & XL_UNICODE_STRING_HIGH_BYTE != 0 {
        UTF16_CHAR_BYTES
    } else {
        COMPRESSED_CHAR_BYTES
    };
    let declared_bytes = cch
        .checked_mul(char_width)
        .ok_or_else(|| invalid(FORMAT_RECORD, "format string length overflow"))?;
    let available = data.len() - XL_UNICODE_STRING_HEADER_LEN;
    let truncated = declared_bytes > available;
    let char_bytes = if truncated {
        tolerance.tolerate(
            XlsFormattingDefect::FormatStringOverrun,
            ordinal,
            u32::try_from(cch).unwrap_or(u32::MAX),
            || invalid(FORMAT_RECORD, "truncated format string characters"),
        )?;
        // Keep only whole characters so a stray odd byte cannot split a
        // UTF-16 code unit.
        available - (available % char_width)
    } else {
        declared_bytes
    };
    let chars = data
        .get(XL_UNICODE_STRING_HEADER_LEN..XL_UNICODE_STRING_HEADER_LEN + char_bytes)
        .ok_or_else(|| invalid(FORMAT_RECORD, "truncated format string characters"))?;
    if !truncated && XL_UNICODE_STRING_HEADER_LEN + char_bytes != data.len() {
        return Err(invalid(
            FORMAT_RECORD,
            "XLUnicodeString has trailing bytes after its characters",
        ));
    }
    let code = if char_width == COMPRESSED_CHAR_BYTES {
        chars.iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units = chars
            .chunks_exact(UTF16_CHAR_BYTES)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        char::decode_utf16(units)
            .collect::<Result<String, _>>()
            .map_err(|_| invalid(FORMAT_RECORD, "format string contains invalid UTF-16"))?
    };
    Ok((code, truncated))
}

fn parse_xf(
    data: &[u8],
    index: u16,
    tolerance: &mut XlsToleranceLog,
) -> XlsResult<XlsExtendedFormat> {
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
    let alignment = crate::xls::alignment::XlsCellAlignment::parse(
        data[6], data[7], data[8], index, tolerance,
    )?;
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

/// Strict-mode shim used by both test modules: the production reader threads a
/// tolerance log, but these tests exercise the default reject-everything policy.
#[cfg(test)]
fn parse_globals_strict(records: &[Record]) -> XlsResult<XlsFormatting> {
    XlsFormatting::parse_globals(
        records,
        &mut XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::Strict),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Strict-mode shim: the production reader threads a tolerance log, but
    /// these tests exercise the default (reject-everything) policy.
    fn parse_xf(data: &[u8], index: u16) -> XlsResult<XlsExtendedFormat> {
        super::parse_xf(
            data,
            index,
            &mut XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::Strict),
        )
    }

    /// Strict-mode shim for [`super::parse_number_format`].
    fn parse_number_format(data: &[u8]) -> XlsResult<XlsNumberFormat> {
        super::parse_number_format(
            data,
            0,
            &mut XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::Strict),
        )
    }

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

    /// Build a `Format` payload whose `cch` claims `declared` characters while
    /// only `present` characters follow.
    fn overlong_format_record(declared: u16, present: &str) -> Vec<u8> {
        let mut data = vec![164, 0];
        data.extend_from_slice(&declared.to_le_bytes());
        data.push(XL_UNICODE_STRING_HIGH_BYTE);
        for unit in present.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    #[test]
    fn a_lenient_policy_truncates_a_format_string_that_overstates_its_payload() {
        const DECLARED: u16 = 40;
        const ORDINAL: u32 = 3;
        let data = overlong_format_record(DECLARED, "0.00");

        let mut tolerance =
            XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::TolerateFormattingDefects);
        let format = super::parse_number_format(&data, ORDINAL, &mut tolerance)
            .expect("a lenient policy decodes the characters that are present");
        assert_eq!(format.code(), "0.00");

        let report = tolerance.into_report();
        assert_eq!(report.count(XlsFormattingDefect::FormatStringOverrun), 1);
        assert_eq!(report.defects()[0].ordinal(), ORDINAL);
        assert_eq!(report.defects()[0].observed(), u32::from(DECLARED));
        assert_eq!(report.defects()[0].record_type(), FORMAT_RECORD);
    }

    #[test]
    fn a_truncated_format_string_drops_a_split_utf16_code_unit() {
        // One trailing byte cannot form a UTF-16 code unit; the repair must
        // discard it rather than decode half a character.
        let mut data = overlong_format_record(40, "0.00");
        data.push(0x30);

        let mut tolerance =
            XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::TolerateFormattingDefects);
        let format = super::parse_number_format(&data, 0, &mut tolerance)
            .expect("a lenient policy keeps only whole characters");
        assert_eq!(format.code(), "0.00");
        assert_eq!(
            tolerance
                .into_report()
                .count(XlsFormattingDefect::FormatStringOverrun),
            1
        );
    }

    #[test]
    fn a_lenient_policy_still_rejects_trailing_bytes_and_a_bad_format_identifier() {
        let mut tolerance =
            XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::TolerateFormattingDefects);
        // Trailing bytes past a satisfied `cch` are a framing defect.
        let mut trailing = overlong_format_record(4, "0.00");
        trailing.push(0);
        trailing.push(0);
        assert!(super::parse_number_format(&trailing, 0, &mut tolerance).is_err());
        // An identifier outside the permitted ranges is not a formatting defect.
        let mut bad_id = overlong_format_record(4, "0.00");
        bad_id[0..2].copy_from_slice(&0u16.to_le_bytes());
        assert!(super::parse_number_format(&bad_id, 0, &mut tolerance).is_err());
        assert!(tolerance.into_report().is_clean());
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
        let formatting = parse_globals_strict(&records).unwrap();

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
                "test-data/poi/test-data/spreadsheet/DateFormats.xls",
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
                "test-data/poi/test-data/spreadsheet/1900DateWindowing.xls",
                XlsDateSystem::Excel1900,
            ),
            (
                "test-data/poi/test-data/spreadsheet/1904DateWindowing.xls",
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
            "test-data/libreoffice-core/sc/qa/unit/data/xls/formats.xls",
            "test-data/libreoffice-core/sc/qa/unit/data/xls/cellformat.xls",
        ] {
            let workbook = XlsWorkbook::new(File::open(fixture(path)).unwrap()).unwrap();
            assert!(workbook.extended_formats().len() >= 16, "{path}");
            for (index, format) in workbook.extended_formats().iter().enumerate() {
                assert_eq!(format.index() as usize, index, "{path}");
            }
        }
    }
}

#[cfg(test)]
mod real_world_tolerance_tests {
    use super::*;
    use crate::xls::records::{Record, RecordHeader};

    fn record(record_type: u16, data: Vec<u8>) -> Record {
        Record {
            header: RecordHeader {
                record_type,
                data_len: data.len() as u16,
            },
            data,
        }
    }

    fn format_record(id: u16, code: &str) -> Record {
        let mut data = id.to_le_bytes().to_vec();
        data.extend_from_slice(&(code.len() as u16).to_le_bytes());
        data.push(0); // compressed characters, no reserved bits
        data.extend_from_slice(code.as_bytes());
        record(FORMAT_RECORD, data)
    }

    /// `MIN_XF_RECORDS` style XFs plus one cell XF, the smallest set
    /// `parse_globals` accepts.
    fn minimal_xf_records() -> Vec<Record> {
        fn xf(style: bool, parent: u16, format_id: u16) -> Record {
            let mut data = vec![0u8; 20];
            data[2..4].copy_from_slice(&format_id.to_le_bytes());
            let flags = (parent << 4) | u16::from(style) << 2 | 1;
            data[4..6].copy_from_slice(&flags.to_le_bytes());
            record(XF_RECORD, data)
        }
        let mut records: Vec<Record> = (0..MIN_XF_RECORDS - 1)
            .map(|_| xf(true, 0x0fff, 0))
            .collect();
        records.push(xf(false, 0, 164));
        records
    }

    /// MS-XLS 2.1 lists Date1904 in the globals grammar, but workbooks written
    /// by other producers omit it. The record selects between exactly two date
    /// systems, so the absent value is unambiguous and must not make the
    /// workbook unreadable.
    #[test]
    fn defaults_to_the_1900_date_system_when_date1904_is_absent() {
        let mut records = vec![format_record(164, "yyyy-mm-dd")];
        records.extend(minimal_xf_records());

        let formatting = parse_globals_strict(&records).expect("a missing Date1904 is not fatal");
        assert_eq!(formatting.date_system(), XlsDateSystem::Excel1900);
    }

    /// An explicit record still wins over the fallback.
    #[test]
    fn honours_an_explicit_1904_date_system() {
        let mut records = vec![record(DATE1904_RECORD, vec![1, 0])];
        records.push(format_record(164, "yyyy-mm-dd"));
        records.extend(minimal_xf_records());

        let formatting = parse_globals_strict(&records).expect("explicit Date1904 parses");
        assert_eq!(formatting.date_system(), XlsDateSystem::Excel1904);
    }

    /// MS-XLS 2.5.294 gives XLUnicodeString exactly one meaningful option bit;
    /// the rest are reserved and "MUST be zero, and MUST be ignored". A writer
    /// that leaves one set must not make the FORMAT record unreadable, and the
    /// bits must not be mistaken for the `fRichSt`/`fExtSt` of the unrelated
    /// XLUnicodeRichExtendedString, which would shift the character offset.
    #[test]
    fn ignores_reserved_option_bits_in_format_record_strings() {
        const CODE: &str = "0.00";
        for reserved in [0x02u8, 0x04, 0x08, 0x20, 0xf0] {
            let mut data = 164u16.to_le_bytes().to_vec();
            data.extend_from_slice(&(CODE.len() as u16).to_le_bytes());
            data.push(reserved); // fHighByte clear, reserved bits set
            data.extend_from_slice(CODE.as_bytes());

            let mut records = vec![
                record(DATE1904_RECORD, vec![0, 0]),
                record(FORMAT_RECORD, data),
            ];
            records.extend(minimal_xf_records());

            let formatting = parse_globals_strict(&records)
                .unwrap_or_else(|error| panic!("reserved bits {reserved:#04x} rejected: {error}"));
            assert_eq!(
                formatting.number_format(164).map(XlsNumberFormat::code),
                Some(CODE),
                "reserved bits {reserved:#04x} must not shift the characters"
            );
        }
    }

    /// Tolerating reserved bits must not turn a genuinely truncated record into
    /// a silently mis-parsed one.
    #[test]
    fn still_rejects_a_truncated_format_record() {
        let mut data = 164u16.to_le_bytes().to_vec();
        data.extend_from_slice(&10u16.to_le_bytes()); // claims 10 characters
        data.push(0);
        data.extend_from_slice(b"0.00"); // supplies 4
        let mut records = vec![
            record(DATE1904_RECORD, vec![0, 0]),
            record(FORMAT_RECORD, data),
        ];
        records.extend(minimal_xf_records());

        assert!(parse_globals_strict(&records).is_err());
    }

    #[test]
    fn a_lenient_policy_repairs_an_xfcrc_count_disagreement_only() {
        let mut records = vec![format_record(164, "yyyy-mm-dd")];
        records.extend(minimal_xf_records());
        let mut crc = [0u8; 20];
        crc[0..2].copy_from_slice(&XFCRC_RECORD.to_le_bytes());
        crc[16..20].copy_from_slice(&99u32.to_le_bytes());
        crc[14..16].copy_from_slice(&(MIN_XF_RECORDS as u16 + 1).to_le_bytes());
        records.push(record(XFCRC_RECORD, crc.to_vec()));

        assert!(parse_globals_strict(&records).is_err());

        let mut tolerance =
            XlsToleranceLog::new(crate::xls::leniency::XlsLeniency::TolerateFormattingDefects);
        let formatting = XlsFormatting::parse_globals(&records, &mut tolerance)
            .expect("a lenient policy trusts the XF records that were parsed");
        assert_eq!(formatting.extended_formats().len(), MIN_XF_RECORDS);

        let report = tolerance.into_report();
        assert_eq!(
            report.count(XlsFormattingDefect::ExtendedFormatCountMismatch),
            1
        );
        let entry = report.defects()[0];
        assert_eq!(entry.ordinal(), MIN_XF_RECORDS as u32);
        assert_eq!(entry.observed(), MIN_XF_RECORDS as u32 + 1);
    }
}
