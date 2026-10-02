//! BIFF8 codecs for workbook date, `Format`, `XF`, and `XFCRC` records.

use std::collections::HashMap;

use super::model::{
    DateSystem, ExtendedFormat, ExtendedFormatApplications, ExtendedFormatKind, Formatting,
    NumberFormat,
};
use super::{
    COMPRESSED_CHAR_BYTES, DATE1904_RECORD, FORMAT_RECORD, MAX_DXF_RECORDS, MAX_FORMAT_RECORDS,
    MAX_XF_RECORDS, MIN_XF_RECORDS, UTF16_CHAR_BYTES, XF_RECORD, XFCRC_RECORD,
    XL_UNICODE_STRING_HIGH_BYTE,
};
use crate::error::{Error, Result};
use crate::leniency::{CellFormatDefect, FormattingDefect, ToleranceLog, ToleranceReport};
use litchi_biff::RecordRef;

/// The mandatory first cell XF (MS-XLS 2.4.353 reserves slots 0–14 for style
/// XFs): the Excel-compatible substitute for a cell-XF reference that cannot
/// format a cell.
const DEFAULT_CELL_XF_INDEX: u16 = 15;

impl Formatting {
    /// The cell format `index` resolves to, applying the Excel-compatible
    /// substitution for a reference that cannot format a cell.
    ///
    /// A valid cell-XF index resolves to its own format. An index of zero
    /// resolves to no format, exactly as stored. Anything else — out of
    /// range, a reserved style slot, or a style XF — resolves to the default
    /// cell XF the way Excel renders it; [`Formatting::check_cell_xf`]
    /// records that substitution at the point the reference is read.
    pub(crate) fn cell_format(&self, index: u16) -> Option<&ExtendedFormat> {
        if index == 0 {
            return None;
        }
        self.extended_format(index)
            .filter(|format| format.is_cell_format())
            .or_else(|| self.default_cell_format())
    }

    /// The default cell XF: the mandatory first cell XF at
    /// [`DEFAULT_CELL_XF_INDEX`], which `parse_globals` guarantees for every
    /// parsed workbook. Absent only for tables built outside the reader.
    fn default_cell_format(&self) -> Option<&ExtendedFormat> {
        self.extended_format(DEFAULT_CELL_XF_INDEX)
            .filter(|format| format.is_cell_format())
    }

    /// Whether a cell, row, or column `ixfe` cannot format a cell and is
    /// repaired by the [`CellFormatDefect::XfReference`] substitution.
    ///
    /// Zero and every index into an empty table are accepted as stored, so
    /// they never count as repairs.
    pub(crate) fn cell_xf_needs_repair(&self, index: u16) -> bool {
        if self.extended_formats.is_empty() || index == 0 {
            return false;
        }
        index < DEFAULT_CELL_XF_INDEX
            || !self
                .extended_format(index)
                .is_some_and(|format| format.is_cell_format())
    }

    /// Record the Excel-compatible repair for `index` when it cannot format
    /// a cell. `ordinal` locates the carrying record (the zero-based row of a
    /// cell or `ROW` record, the first column of a `COLINFO` record).
    ///
    /// The substitution itself is applied by [`Formatting::cell_format`], so
    /// every consumer of the index sees the default cell XF.
    pub(crate) fn check_cell_xf(&self, index: u16, ordinal: u32, tolerance: &mut ToleranceReport) {
        if self.cell_xf_needs_repair(index) {
            tolerance.record(CellFormatDefect::XfReference, ordinal, u32::from(index));
        }
    }

    /// Parse the formatting records of the workbook globals under `tolerance`.
    ///
    /// See [`FormattingDefect`] for the exhaustive list of defects a lenient
    /// policy repairs here; every other validation is unchanged.
    pub(crate) fn parse_globals(
        records: &[RecordRef<'_>],
        tolerance: &mut ToleranceLog,
    ) -> Result<Self> {
        let mut date_system = None;
        let mut number_formats = Vec::new();
        let mut extended_formats = Vec::new();
        let mut differential_formats = Vec::new();
        let mut format_by_id = HashMap::new();
        let mut xfcrc = None;
        let mut xf_extensions = Vec::new();

        for record in records {
            match record.kind().get() {
                DATE1904_RECORD => {
                    if date_system.is_some() {
                        return Err(invalid(DATE1904_RECORD, "duplicate Date1904 record"));
                    }
                    date_system = Some(parse_date_system(record.payload())?);
                },
                FORMAT_RECORD => {
                    if number_formats.len() == MAX_FORMAT_RECORDS {
                        return Err(invalid(FORMAT_RECORD, "more than 218 Format records"));
                    }
                    let ordinal = u32::try_from(number_formats.len()).map_err(|_error| {
                        invalid(FORMAT_RECORD, "Format record ordinal does not fit in u32")
                    })?;
                    let format = parse_number_format(record.payload(), ordinal, tolerance)?;
                    if !is_permitted_format_id(format.id) {
                        // Kept, as LibreOffice keeps it: an XF naming the id
                        // still resolves to this record's code.
                        tolerance.report_mut().record(
                            CellFormatDefect::FormatIdentifier,
                            ordinal,
                            u32::from(format.id),
                        );
                    }
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
                    let index = u16::try_from(extended_formats.len()).map_err(|_error| {
                        invalid(XF_RECORD, "XF index does not fit in the BIFF index field")
                    })?;
                    extended_formats.push(parse_xf(record.payload(), index, tolerance)?);
                },
                XFCRC_RECORD => {
                    if xfcrc.is_some() {
                        return Err(invalid(XFCRC_RECORD, "duplicate XFCRC record"));
                    }
                    xfcrc = Some(parse_xfcrc(
                        record.payload(),
                        u32::try_from(record.offset()).unwrap_or(u32::MAX),
                        tolerance.report_mut(),
                    )?);
                },
                crate::xf_ext::XF_EXT_RECORD_TYPE => {
                    if xf_extensions.len() == MAX_XF_RECORDS {
                        return Err(invalid(
                            crate::xf_ext::XF_EXT_RECORD_TYPE,
                            "more than 65,536 XFExt records",
                        ));
                    }
                    xf_extensions.push(crate::xf_ext::XfExt::parse(record.payload())?);
                },
                crate::differential_format::DXF_RECORD_TYPE => {
                    if differential_formats.len() == MAX_DXF_RECORDS {
                        return Err(invalid(
                            crate::differential_format::DXF_RECORD_TYPE,
                            "more than 65,536 DXF records",
                        ));
                    }
                    differential_formats.push(
                        crate::differential_format::DifferentialFormat::parse_payload(
                            record.payload(),
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
        let date_system = date_system.unwrap_or(DateSystem::Excel1900);
        if extended_formats.len() < MIN_XF_RECORDS {
            return Err(invalid(
                XF_RECORD,
                format!(
                    "workbook has {} XF records; expected at least 16",
                    extended_formats.len()
                ),
            ));
        }
        if let Some(Some(count)) = xfcrc {
            let parsed = extended_formats.len();
            if count as usize != parsed {
                // XFCRC is a redundant integrity summary over the XF table.
                // The XF records themselves remain individually well-formed, so
                // a disagreeing cxfs costs no cell data.
                tolerance.tolerate(
                    FormattingDefect::ExtendedFormatCountMismatch,
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

        // An `ifmt` that names neither a built-in format nor a `Format` record
        // renders as General in Excel; POI never checks it at open.
        for format in &mut extended_formats {
            let id = format.number_format_id;
            if id > MAX_BUILT_IN_FORMAT_ID && !format_by_id.contains_key(&id) {
                tolerance.report_mut().record(
                    CellFormatDefect::NumberFormatReference,
                    u32::from(format.index),
                    u32::from(id),
                );
                format.number_format_id = GENERAL_FORMAT_ID;
            }
        }
        for (index, format) in extended_formats.iter().enumerate() {
            if index < 15 && !matches!(format.kind, ExtendedFormatKind::Style) {
                return Err(invalid(
                    XF_RECORD,
                    format!("mandatory XF slot {index} is not a style XF"),
                ));
            }
            if index == 15 && !format.is_cell_format() {
                return Err(invalid(XF_RECORD, "mandatory XF slot 15 is not a cell XF"));
            }
            if let ExtendedFormatKind::Cell { parent_style_xf } = format.kind {
                let parent = extended_formats
                    .get(parent_style_xf as usize)
                    .ok_or_else(|| {
                        invalid(
                            XF_RECORD,
                            format!("XF {index} has out-of-range parent {parent_style_xf}"),
                        )
                    })?;
                if !matches!(parent.kind, ExtendedFormatKind::Style) {
                    return Err(invalid(
                        XF_RECORD,
                        format!("XF {index} parent {parent_style_xf} is not a style XF"),
                    ));
                }
            }
        }

        for extension in &xf_extensions {
            if usize::from(extension.xf_index()) >= extended_formats.len() {
                return Err(invalid(
                    crate::xf_ext::XF_EXT_RECORD_TYPE,
                    format!(
                        "XFExt references XF index {} but only {} XF records exist",
                        extension.xf_index(),
                        extended_formats.len()
                    ),
                ));
            }
        }

        for (dxf_index, dxf) in (0_u32..).zip(differential_formats.iter_mut()) {
            for id in dxf.number_format_ids_mut() {
                if *id > MAX_BUILT_IN_FORMAT_ID && !format_by_id.contains_key(id) {
                    tolerance.report_mut().record(
                        CellFormatDefect::DxfNumberFormatReference,
                        dxf_index,
                        u32::from(*id),
                    );
                    *id = GENERAL_FORMAT_ID;
                }
            }
        }

        Ok(Self {
            date_system,
            number_formats,
            extended_formats,
            differential_formats,
            xf_extensions,
            format_by_id,
        })
    }
}

/// The General number format: the substitute for an unresolvable `ifmt`.
const GENERAL_FORMAT_ID: u16 = 0;
/// The largest number format identifier that never needs a `Format` record.
const MAX_BUILT_IN_FORMAT_ID: u16 = 81;

/// Whether MS-XLS 2.4.126 permits a `Format` record to define `id`.
fn is_permitted_format_id(id: u16) -> bool {
    matches!(id, 5..=8 | 23..=26 | 41..=44 | 63..=66 | 164..=392)
}

pub(super) fn parse_date_system(data: &[u8]) -> Result<DateSystem> {
    if data.len() != 2 {
        return Err(invalid(
            DATE1904_RECORD,
            format!("Date1904 payload has {} bytes; expected 2", data.len()),
        ));
    }
    match u16::from_le_bytes([data[0], data[1]]) {
        0 => Ok(DateSystem::Excel1900),
        1 => Ok(DateSystem::Excel1904),
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

/// `cch` plus the option byte that precede an `XLUnicodeString`'s characters.
const XL_UNICODE_STRING_HEADER_LEN: usize = 3;
/// Offset of the option byte within an `XLUnicodeString`.
const XL_UNICODE_STRING_FLAGS_OFFSET: usize = 2;

pub(super) fn parse_number_format(
    data: &[u8],
    ordinal: u32,
    tolerance: &mut ToleranceLog,
) -> Result<NumberFormat> {
    if data.len() < 5 {
        return Err(invalid(FORMAT_RECORD, "truncated Format record"));
    }
    // An identifier outside the permitted ranges is recorded by the caller as
    // `CellFormatDefect::FormatIdentifier` and the record kept.
    let id = u16::from_le_bytes([data[0], data[1]]);
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
    Ok(NumberFormat {
        id,
        date_time: is_custom_date_time(&code),
        code,
    })
}

/// Decode a `Format` record's `XLUnicodeString`.
///
/// Returns the decoded code and whether the payload was shorter than `cch`
/// declared. A short payload is only reachable when `tolerance` permits
/// [`FormattingDefect::FormatStringOverrun`]; a *longer* payload stays a hard
/// error, because trailing bytes mean the record framing is wrong rather than
/// merely the count.
fn parse_xl_unicode_string(
    data: &[u8],
    ordinal: u32,
    tolerance: &mut ToleranceLog,
) -> Result<(String, bool)> {
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
            FormattingDefect::FormatStringOverrun,
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
            .as_chunks::<UTF16_CHAR_BYTES>()
            .0
            .iter()
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        char::decode_utf16(units)
            .collect::<Result<String, _>>()
            .map_err(|_error| invalid(FORMAT_RECORD, "format string contains invalid UTF-16"))?
    };
    Ok((code, truncated))
}

pub(super) fn parse_xf(
    data: &[u8],
    index: u16,
    tolerance: &mut ToleranceLog,
) -> Result<ExtendedFormat> {
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
        // MS-XLS 2.4.353 leaves f123Prefix undefined on a style XF; it is
        // ignored there (the cell-only quote prefix is never derived from it).
        tolerance.report_mut().record(
            CellFormatDefect::StylePrefix,
            u32::from(index),
            u32::from(flags),
        );
    }
    let kind = if style {
        ExtendedFormatKind::Style
    } else {
        ExtendedFormatKind::Cell {
            parent_style_xf: parent,
        }
    };
    let mut font_index = u16::from_le_bytes([data[0], data[1]]);
    if crate::font::validate_font_index(font_index).is_err() {
        // Index 4 is reserved and indices above 1022 are out of range; Excel
        // renders such a cell in the default font.
        tolerance.report_mut().record(
            CellFormatDefect::FontReference,
            u32::from(index),
            u32::from(font_index),
        );
        font_index = 0;
    }
    let applications = if style {
        ExtendedFormatApplications::all_local()
    } else {
        ExtendedFormatApplications::from_cell_bits(data[9] >> 2)
    };
    let border2 = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let area = u16::from_le_bytes([data[18], data[19]]);
    let quote_prefix = !style && prefix;
    let has_xf_extension = !style && border2 & (1 << 25) != 0;
    let pivot_button = !style && area & (1 << 14) != 0;
    let alignment =
        crate::alignment::CellAlignment::parse(data[6], data[7], data[8], index, tolerance)?;
    let (borders, fill) =
        crate::border_fill::parse_xf_border_fill(data, index, tolerance.report_mut());

    Ok(ExtendedFormat {
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

/// Parse an `XFCRC` payload into its `cxfs`, or `None` when the count is
/// outside `16..=4050` and the record is therefore ignored.
///
/// `XFCRC` is a redundant integrity summary, so its header and reserved
/// fields are ignored ("MUST be ignored") and an impossible count discards
/// the summary; each repair is recorded against `record_offset`. Only a
/// payload of the wrong size stays a framing error.
pub(super) fn parse_xfcrc(
    data: &[u8],
    record_offset: u32,
    tolerance: &mut ToleranceReport,
) -> Result<Option<u16>> {
    if data.len() != 20 {
        return Err(invalid(
            XFCRC_RECORD,
            format!("XFCRC payload has {} bytes; expected 20", data.len()),
        ));
    }
    let header_type = u16::from_le_bytes([data[0], data[1]]);
    let reserved = data[2..14].iter().filter(|byte| **byte != 0).count();
    if header_type != XFCRC_RECORD || reserved != 0 {
        tolerance.record(
            CellFormatDefect::XfcrcReserved,
            record_offset,
            u32::try_from(reserved).unwrap_or(u32::MAX),
        );
    }
    let count = u16::from_le_bytes([data[14], data[15]]);
    if !(16..=4050).contains(&count) {
        tolerance.record(
            CellFormatDefect::XfcrcCount,
            record_offset,
            u32::from(count),
        );
        return Ok(None);
    }
    Ok(Some(count))
}

pub(super) fn is_custom_date_time(code: &str) -> bool {
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

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}
