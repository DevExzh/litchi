//! BIFF8 PivotTable OLAP extension records: the `SXVIEWEX` sequence of the
//! worksheet substream (MS-XLS 2.1).
//!
//! The grammar is:
//!
//! ```text
//! SXVIEWEX = SXViewEx *SXTH *SXPIEx *SXVDTEx
//! ```
//!
//! This module implements typed readers and writers for all four records:
//!
//! - **SXViewEx** (0x080E): counts of the OLAP extension records that
//!   follow, plus a bounded blob of future-version bytes.
//! - **SXTH** (0x00DB): properties of one OLAP pivot hierarchy.
//! - **SXPIEx** (0x080C): OLAP extension of one page-axis entry.
//! - **SXVDTEx** (0x080F): OLAP extension of one pivot field.
//!
//! Spec note (MS-XLS erratum preserved verbatim): the `FrtHeaderOld.rt`
//! field of these records does NOT always equal the containing record type.
//! The spec mandates `rt = 0x080C` inside SXViewEx (record type 0x080E),
//! `rt = 0x080E` inside SXPIEx (record type 0x080C), and `rt = 0x080D`
//! inside SXTH (record type 0x00DB). Only SXVDTEx carries its own record
//! type. The mandated values are validated on read and reproduced on write.
//!
//! Everything in this module is INERT: MDX unique names and member captions
//! are stored verbatim and no OLAP server is ever contacted.
//!
//! # References
//!
//! - MS-XLS 2.4.299 (SXPIEx), 2.4.308 (SXTH), 2.4.311 (SXVDTEx),
//!   2.4.314 (SXViewEx), 2.5.134 (FrtFlags), 2.5.136 (FrtHeaderOld),
//!   2.5.157 (HiddenMemberSet), 2.5.254 (SXAxis), 2.5.263 (SXVIFlags),
//!   2.5.294 (XLUnicodeString)

use super::{XlsError, XlsResult};

/// Record type of the `SXPIEx` record (MS-XLS 2.4.299).
pub(crate) const SXPI_EX_RECORD_TYPE: u16 = 0x080C;
/// Record type of the `SXViewEx` record (MS-XLS 2.4.314).
pub(crate) const SX_VIEW_EX_RECORD_TYPE: u16 = 0x080E;
/// Record type of the `SXVDTEx` record (MS-XLS 2.4.311).
pub(crate) const SXVDT_EX_RECORD_TYPE: u16 = 0x080F;
/// Record type of the `SXTH` record (MS-XLS 2.4.308).
pub(crate) const SXTH_RECORD_TYPE: u16 = 0x00DB;

/// `FrtHeaderOld.rt` mandated inside `SXViewEx` (MS-XLS 2.4.314). This is
/// NOT the `SXViewEx` record type; see the module documentation.
const SX_VIEW_EX_FRT_RT: u16 = 0x080C;
/// `FrtHeaderOld.rt` mandated inside `SXPIEx` (MS-XLS 2.4.299). This is
/// NOT the `SXPIEx` record type; see the module documentation.
const SXPI_EX_FRT_RT: u16 = 0x080E;
/// `FrtHeaderOld.rt` mandated inside `SXVDTEx` (MS-XLS 2.4.311).
const SXVDT_EX_FRT_RT: u16 = SXVDT_EX_RECORD_TYPE;
/// `FrtHeaderOld.rt` mandated inside `SXTH` (MS-XLS 2.4.308). This is NOT
/// the `SXTH` record type; see the module documentation.
const SXTH_FRT_RT: u16 = 0x080D;

/// Size in bytes of an `FrtHeaderOld` (MS-XLS 2.5.136).
const FRT_HEADER_OLD_LEN: usize = 4;
/// Maximum character count of the `XLUnicodeString` fields in these
/// records (MS-XLS 2.4.299, 2.4.308).
const MAX_OLAP_STRING_CHARS: usize = 255;
/// Maximum byte count of `SXViewEx.rgbFuture` (MS-XLS 2.4.314).
const MAX_FUTURE_BYTES: usize = 1_024;

/// `fHighByte` bit of an `XLUnicodeString` option byte (MS-XLS 2.5.294).
const HIGH_BYTE: u8 = 0x01;
/// Reserved bits of an `XLUnicodeString` option byte: all but `fHighByte`.
const STRING_OPTION_RESERVED: u8 = !HIGH_BYTE;

// `SXTH` flag-word bits (MS-XLS 2.4.308).
const TH_MEASURE: u32 = 0x0000_0001;
const TH_OUTLINE_MODE: u32 = 0x0000_0004;
const TH_MULTIPLE_PAGE_ITEMS: u32 = 0x0000_0008;
const TH_SUBTOTAL_AT_TOP: u32 = 0x0000_0010;
const TH_NAMED_SET: u32 = 0x0000_0020;
const TH_HIDDEN_FROM_FIELD_LIST: u32 = 0x0000_0040;
const TH_ATTRIBUTE_HIERARCHY: u32 = 0x0000_0080;
const TH_TIME_HIERARCHY: u32 = 0x0000_0100;
const TH_FILTER_INCLUSIVE: u32 = 0x0000_0200;
const TH_KEY_ATTRIBUTE_HIERARCHY: u32 = 0x0000_0800;
const TH_KPI: u32 = 0x0000_1000;

// `SXTH` drag-permission bits (MS-XLS 2.4.308).
const DRAG_TO_ROW: u16 = 0x0001;
const DRAG_TO_COLUMN: u16 = 0x0002;
const DRAG_TO_PAGE: u16 = 0x0004;
const DRAG_TO_DATA: u16 = 0x0008;
const DRAG_TO_HIDE: u16 = 0x0010;

// `SXAxis` bits (MS-XLS 2.5.254).
const AXIS_ROW: u16 = 0x0001;
const AXIS_COLUMN: u16 = 0x0002;
const AXIS_PAGE: u16 = 0x0004;
const AXIS_DATA: u16 = 0x0008;
/// Reserved bits of an `SXAxis` word: MUST be zero (MS-XLS 2.5.254).
const AXIS_RESERVED: u16 = 0xFFF0;

// `SXVDTEx` flag bits (MS-XLS 2.4.311).
const VDT_TENSOR_SORT: u16 = 0x0001;
const VDT_DRILLED_LEVEL: u16 = 0x0002;
const VDT_ITEMS_DRILLED_BY_DEFAULT: u16 = 0x0004;
const VDT_MEMBER_PROPERTY_IN_REPORT: u16 = 0x0008;
const VDT_MEMBER_PROPERTY_IN_TIP: u16 = 0x0010;
const VDT_MEMBER_PROPERTY_IN_CAPTION: u16 = 0x0020;
/// Reserved bits of the `SXVDTEx` flag word: MUST be zero (MS-XLS 2.4.311).
const VDT_RESERVED: u16 = 0xFFC0;

// `SXVIFlags` bits (MS-XLS 2.5.263).
const VI_DRILLED_MEMBER: u16 = 0x0001;
const VI_HAS_CHILDREN: u16 = 0x0004;
const VI_COLLAPSED_MEMBER: u16 = 0x0008;
const VI_HAS_CHILDREN_EST: u16 = 0x0010;
const VI_OLAP_FILTER_SELECTED: u16 = 0x0020;
/// Reserved bits of an `SXVIFlags` word: MUST be zero (MS-XLS 2.5.263).
const VI_RESERVED: u16 = 0xFFC2;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_i16(data: &[u8], offset: usize) -> i16 {
    read_u16(data, offset) as i16
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    read_u32(data, offset) as i32
}

/// Borrow `len` bytes at `offset`, or fail with a truncation error.
fn slice_at(data: &[u8], offset: usize, len: usize, record_type: u16) -> XlsResult<&[u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| invalid(record_type, "field offset overflow"))?;
    data.get(offset..end).ok_or(XlsError::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

/// Validate an `FrtHeaderOld` (MS-XLS 2.5.136) and return the record body
/// that follows it. `grbitFrt.fFrtRef` and `grbitFrt.fFrtAlert` MUST be
/// zero and the reserved bits MUST be zero, so the whole word must be 0.
fn frt_body(data: &[u8], record_type: u16, expected_rt: u16) -> XlsResult<&[u8]> {
    let header = slice_at(data, 0, FRT_HEADER_OLD_LEN, record_type)?;
    if read_u16(header, 0) != expected_rt {
        return Err(invalid(record_type, "FrtHeaderOld.rt mismatch"));
    }
    if read_u16(header, 2) != 0 {
        return Err(invalid(record_type, "FrtHeaderOld.grbitFrt must be zero"));
    }
    Ok(&data[FRT_HEADER_OLD_LEN..])
}

/// Parse an `XLUnicodeString` (MS-XLS 2.5.294) at the start of `data`.
/// Returns the string and the number of bytes consumed.
fn parse_olap_string(data: &[u8], record_type: u16) -> XlsResult<(String, usize)> {
    let header = slice_at(data, 0, 3, record_type)?;
    let char_count = usize::from(read_u16(header, 0));
    if char_count > MAX_OLAP_STRING_CHARS {
        return Err(invalid(record_type, "XLUnicodeString exceeds 255 characters"));
    }
    let options = header[2];
    if options & STRING_OPTION_RESERVED != 0 {
        return Err(invalid(record_type, "XLUnicodeString reserved option bits set"));
    }
    let wide = options & HIGH_BYTE != 0;
    let byte_len = char_count * if wide { 2 } else { 1 };
    let bytes = slice_at(data, 3, byte_len, record_type)?;
    let text = if wide {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units)
            .map_err(|_| invalid(record_type, "XLUnicodeString is not valid UTF-16LE"))?
    } else {
        bytes.iter().map(|&byte| char::from(byte)).collect()
    };
    Ok((text, 3 + byte_len))
}

/// Serialize an `XLUnicodeString` (MS-XLS 2.5.294) into `output`,
/// compressed when every character is in U+0000..=U+00FF and wide
/// otherwise.
fn append_olap_string(record_type: u16, value: &str, output: &mut Vec<u8>) -> XlsResult<()> {
    let compressible = value.chars().all(|ch| u32::from(ch) <= 0xFF);
    let char_count = if compressible {
        value.len()
    } else {
        value.encode_utf16().count()
    };
    if char_count > MAX_OLAP_STRING_CHARS {
        return Err(XlsError::InvalidData(format!(
            "record 0x{record_type:04X} XLUnicodeString exceeds {MAX_OLAP_STRING_CHARS} characters"
        )));
    }
    output.extend_from_slice(&(char_count as u16).to_le_bytes());
    if compressible {
        output.push(0u8); // fHighByte = 0
        output.extend(value.chars().map(|ch| ch as u8));
    } else {
        output.push(HIGH_BYTE);
        for unit in value.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(())
}

/// Validate an `XLUnicodeString` field length on the write path.
fn check_string_len(record_type: u16, value: &str, field: &str, allow_empty: bool) -> XlsResult<()> {
    let char_count = value.chars().count();
    if char_count > MAX_OLAP_STRING_CHARS {
        return Err(XlsError::InvalidData(format!(
            "record 0x{record_type:04X} {field} exceeds {MAX_OLAP_STRING_CHARS} characters"
        )));
    }
    if !allow_empty && char_count == 0 {
        return Err(XlsError::InvalidData(format!(
            "record 0x{record_type:04X} {field} must not be empty"
        )));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// SXViewEx (MS-XLS 2.4.314)
// ---------------------------------------------------------------------------

/// Typed `SXViewEx` record content (MS-XLS 2.4.314): the header of the
/// PivotTable view OLAP extension sequence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPivotViewOlapHeader {
    /// Number of `SXTH` records that follow (`csxth`). MUST be at least 1.
    pub hierarchy_count: u32,
    /// Number of `SXPIEx` records that follow the `SXTH` records (`csxpi`).
    pub page_extension_count: u32,
    /// Number of `SXVDTEx` records that follow the `SXPIEx` records
    /// (`csxvdtex`).
    pub field_extension_count: u32,
    /// Information from future versions (`rgbFuture`), at most 1024 bytes.
    pub future_bytes: Vec<u8>,
}

impl XlsPivotViewOlapHeader {
    /// Parse an `SXViewEx` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let body = frt_body(data, SX_VIEW_EX_RECORD_TYPE, SX_VIEW_EX_FRT_RT)?;
        let fixed = slice_at(body, 0, 16, SX_VIEW_EX_RECORD_TYPE)?;
        let hierarchy_count = read_i32(fixed, 0);
        let page_extension_count = read_i32(fixed, 4);
        let field_extension_count = read_i32(fixed, 8);
        if hierarchy_count < 1 {
            return Err(invalid(SX_VIEW_EX_RECORD_TYPE, "SXViewEx csxth must be at least 1"));
        }
        if page_extension_count < 0 || field_extension_count < 0 {
            return Err(invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx record counts must be non-negative",
            ));
        }
        let future_len = read_u32(fixed, 12) as usize;
        if future_len > MAX_FUTURE_BYTES {
            return Err(invalid(SX_VIEW_EX_RECORD_TYPE, "SXViewEx cbFuture exceeds 1024"));
        }
        let future_bytes = slice_at(body, 16, future_len, SX_VIEW_EX_RECORD_TYPE)?;
        if 16 + future_len != body.len() {
            return Err(invalid(SX_VIEW_EX_RECORD_TYPE, "SXViewEx cbFuture does not match the record size"));
        }
        Ok(XlsPivotViewOlapHeader {
            hierarchy_count: hierarchy_count as u32,
            page_extension_count: page_extension_count as u32,
            field_extension_count: field_extension_count as u32,
            future_bytes: future_bytes.to_vec(),
        })
    }

    /// Serialize back to a complete `SXViewEx` record payload.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        if self.hierarchy_count < 1 {
            return Err(XlsError::InvalidData(
                "SXViewEx hierarchy_count must be at least 1".to_string(),
            ));
        }
        if self.future_bytes.len() > MAX_FUTURE_BYTES {
            return Err(XlsError::InvalidData(format!(
                "SXViewEx future_bytes exceeds {MAX_FUTURE_BYTES} bytes"
            )));
        }
        let mut payload = Vec::with_capacity(FRT_HEADER_OLD_LEN + 16 + self.future_bytes.len());
        payload.extend_from_slice(&SX_VIEW_EX_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&(self.hierarchy_count as i32).to_le_bytes());
        payload.extend_from_slice(&(self.page_extension_count as i32).to_le_bytes());
        payload.extend_from_slice(&(self.field_extension_count as i32).to_le_bytes());
        payload.extend_from_slice(&(self.future_bytes.len() as u32).to_le_bytes());
        payload.extend_from_slice(&self.future_bytes);
        Ok(payload)
    }
}

// ---------------------------------------------------------------------------
// SXTH (MS-XLS 2.4.308)
// ---------------------------------------------------------------------------

/// The PivotTable axis or axes a pivot hierarchy is present on (`SXAxis`,
/// MS-XLS 2.5.254).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct XlsPivotHierarchyAxis {
    /// Whether the hierarchy is on the row axis (`sxaxisRw`).
    pub row: bool,
    /// Whether the hierarchy is on the column axis (`sxaxisCol`).
    pub column: bool,
    /// Whether the hierarchy is on the page axis (`sxaxisPage`).
    pub page: bool,
    /// Whether the hierarchy is on the data axis (`sxaxisData`).
    pub data: bool,
}

impl XlsPivotHierarchyAxis {
    fn from_bits(record_type: u16, bits: u16) -> XlsResult<Self> {
        if bits & AXIS_RESERVED != 0 {
            return Err(invalid(record_type, "SXAxis reserved bits set"));
        }
        Ok(XlsPivotHierarchyAxis {
            row: bits & AXIS_ROW != 0,
            column: bits & AXIS_COLUMN != 0,
            page: bits & AXIS_PAGE != 0,
            data: bits & AXIS_DATA != 0,
        })
    }

    fn bits(self) -> u16 {
        let mut bits = 0u16;
        if self.row {
            bits |= AXIS_ROW;
        }
        if self.column {
            bits |= AXIS_COLUMN;
        }
        if self.page {
            bits |= AXIS_PAGE;
        }
        if self.data {
            bits |= AXIS_DATA;
        }
        bits
    }
}

/// A `HiddenMemberSet` structure (MS-XLS 2.5.157): the OLAP members hidden
/// from the PivotTable view at one level of a pivot hierarchy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsHiddenMemberSet {
    /// Names of the hidden OLAP members (`rgMemberName`), each at most 255
    /// characters.
    pub member_names: Vec<String>,
}

/// Typed `SXTH` record content (MS-XLS 2.4.308): properties of one OLAP
/// pivot hierarchy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPivotHierarchy {
    /// Whether this hierarchy is an OLAP measure (`fMeasure`).
    pub is_measure: bool,
    /// Whether level fields are created with `SXVDEx.fOutline` set
    /// (`fOutlineMode`).
    pub outline_mode: bool,
    /// Whether multiple OLAP members can be selected on the page axis
    /// (`fEnableMultiplePageItems`).
    pub multiple_page_items: bool,
    /// Whether level fields are created with `SXVDEx.fSubtotalAtTop` set
    /// (`fSubtotalAtTop`).
    pub subtotal_at_top: bool,
    /// Whether this hierarchy is an OLAP named set (`fSet`).
    pub is_named_set: bool,
    /// Whether this hierarchy is hidden in the field list
    /// (`fDontShowFList`).
    pub hidden_from_field_list: bool,
    /// Whether this hierarchy is an attribute hierarchy
    /// (`fAttributeHierarchy`).
    pub is_attribute_hierarchy: bool,
    /// Whether this hierarchy is a time hierarchy (`fTimeHierarchy`).
    pub is_time_hierarchy: bool,
    /// Whether manual filters are inclusive rather than exclusive
    /// (`fFilterInclusive`).
    pub filter_inclusive: bool,
    /// Whether this is the key attribute hierarchy of its dimension
    /// (`fKeyAttributeHierarchy`).
    pub is_key_attribute_hierarchy: bool,
    /// Whether this hierarchy is a KPI hierarchy (`fKPI`).
    pub is_kpi: bool,
    /// The axis or axes this hierarchy is present on (`sxaxis`).
    pub axis: XlsPivotHierarchyAxis,
    /// The associated pivot field index (`isxvd`).
    pub pivot_field_index: i32,
    /// Number of pivot fields on PivotTable axes for this hierarchy
    /// (`csxvdXl`). Related to `level_fields` by the `stAll` rule (see
    /// [`XlsPivotHierarchy::parse`]).
    pub axis_field_count: i32,
    /// Whether this hierarchy can be placed on the row axis (`fDragToRow`).
    pub drag_to_row: bool,
    /// Whether this hierarchy can be placed on the column axis
    /// (`fDragToColumn`).
    pub drag_to_column: bool,
    /// Whether this hierarchy can be placed on the page axis (`fDragToPage`).
    pub drag_to_page: bool,
    /// Whether this hierarchy can be placed on the data axis (`fDragToData`).
    pub drag_to_data: bool,
    /// Whether this hierarchy can be removed from the view (`fDragToHide`).
    pub drag_to_hide: bool,
    /// MDX unique name of this hierarchy (`stUnique`), 1..=255 characters.
    pub unique_name: String,
    /// Display name of this hierarchy (`stDisplay`), 1..=255 characters.
    pub display_name: String,
    /// MDX unique name of the default member (`stDefault`), at most 255
    /// characters.
    pub default_member: String,
    /// Unique name of the ALL member (`stAll`); empty when the hierarchy
    /// has no ALL member.
    pub all_member: String,
    /// Unique name of the OLAP dimension this hierarchy belongs to
    /// (`stDimension`); MUST be empty for measures.
    pub dimension: String,
    /// Pivot fields associated with this hierarchy (`rgisxvd`); each
    /// element is a pivot field index or -1 for none.
    pub level_fields: Vec<i32>,
    /// Hidden OLAP members per level (`rgHiddenMemberSets`).
    pub hidden_member_sets: Vec<XlsHiddenMemberSet>,
}

impl XlsPivotHierarchy {
    /// Parse an `SXTH` record payload.
    ///
    /// Cross-field rules enforced (MS-XLS 2.4.308):
    ///
    /// - a measure cannot be a named set, cannot drag to the row, column,
    ///   or page axis, and has an empty `stDimension`;
    /// - `level_fields` (`cisxvd`) MUST be empty when the hierarchy is on
    ///   neither the row nor the column axis;
    /// - `axis_field_count` (`csxvdXl`) MUST equal the `level_fields` count
    ///   when `all_member` (`stAll`) is empty, and one less otherwise;
    /// - `hidden_member_sets` MUST be empty when `filter_inclusive` is set.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let body = frt_body(data, SXTH_RECORD_TYPE, SXTH_FRT_RT)?;
        let fixed = slice_at(body, 0, 18, SXTH_RECORD_TYPE)?;
        let flags = read_u32(fixed, 0);
        let is_measure = flags & TH_MEASURE != 0;
        let is_named_set = flags & TH_NAMED_SET != 0;
        let filter_inclusive = flags & TH_FILTER_INCLUSIVE != 0;
        let axis = XlsPivotHierarchyAxis::from_bits(SXTH_RECORD_TYPE, read_u16(fixed, 4))?;
        if read_u16(fixed, 6) != 0 {
            return Err(invalid(SXTH_RECORD_TYPE, "SXTH reserved word must be zero"));
        }
        let pivot_field_index = read_i32(fixed, 8);
        let axis_field_count = read_i32(fixed, 12);
        if axis_field_count < 0 {
            return Err(invalid(SXTH_RECORD_TYPE, "SXTH csxvdXl must be non-negative"));
        }
        let drag = read_u16(fixed, 16);
        let drag_to_row = drag & DRAG_TO_ROW != 0;
        let drag_to_column = drag & DRAG_TO_COLUMN != 0;
        let drag_to_page = drag & DRAG_TO_PAGE != 0;
        if is_measure && (is_named_set || drag_to_row || drag_to_column || drag_to_page) {
            return Err(invalid(
                SXTH_RECORD_TYPE,
                "SXTH measure cannot be a named set or drag to row/column/page",
            ));
        }

        let mut offset = 18;
        let (unique_name, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
        offset += used;
        let (display_name, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
        offset += used;
        let (default_member, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
        offset += used;
        let (all_member, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
        offset += used;
        let (dimension, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
        offset += used;
        if unique_name.is_empty() || display_name.is_empty() {
            return Err(invalid(SXTH_RECORD_TYPE, "SXTH stUnique/stDisplay must not be empty"));
        }
        if is_measure && !dimension.is_empty() {
            return Err(invalid(SXTH_RECORD_TYPE, "SXTH measure must have an empty stDimension"));
        }

        let level_count = read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0) as usize;
        offset += 4;
        if level_count > 0 && !axis.row && !axis.column {
            return Err(invalid(
                SXTH_RECORD_TYPE,
                "SXTH cisxvd must be zero off the row/column axes",
            ));
        }
        let expected_axis_fields = if all_member.is_empty() {
            level_count as i64
        } else {
            level_count as i64 - 1
        };
        if i64::from(axis_field_count) != expected_axis_fields {
            return Err(invalid(
                SXTH_RECORD_TYPE,
                "SXTH csxvdXl does not match cisxvd and stAll",
            ));
        }
        let mut level_fields = Vec::with_capacity(level_count.min(body.len() / 4));
        for _ in 0..level_count {
            let value = read_i32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0);
            offset += 4;
            if value < -1 {
                return Err(invalid(SXTH_RECORD_TYPE, "SXTH rgisxvd element must be -1 or a pivot field index"));
            }
            level_fields.push(value);
        }

        let hidden_set_count = read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0) as usize;
        offset += 4;
        if hidden_set_count > 0 && filter_inclusive {
            return Err(invalid(
                SXTH_RECORD_TYPE,
                "SXTH cHiddenMemberSets must be zero for inclusive filters",
            ));
        }
        let mut hidden_member_sets = Vec::new();
        // rgHiddenMemberSets exists iff cHiddenMemberSets > 0 and cisxvd > 0.
        if hidden_set_count > 0 && level_count > 0 {
            for _ in 0..hidden_set_count {
                let name_count =
                    read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0) as usize;
                offset += 4;
                let mut member_names = Vec::with_capacity(name_count.min(body.len() / 3));
                for _ in 0..name_count {
                    let (name, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
                    offset += used;
                    member_names.push(name);
                }
                hidden_member_sets.push(XlsHiddenMemberSet { member_names });
            }
        }
        if offset != body.len() {
            return Err(invalid(SXTH_RECORD_TYPE, "trailing bytes after SXTH"));
        }

        Ok(XlsPivotHierarchy {
            is_measure,
            outline_mode: flags & TH_OUTLINE_MODE != 0,
            multiple_page_items: flags & TH_MULTIPLE_PAGE_ITEMS != 0,
            subtotal_at_top: flags & TH_SUBTOTAL_AT_TOP != 0,
            is_named_set,
            hidden_from_field_list: flags & TH_HIDDEN_FROM_FIELD_LIST != 0,
            is_attribute_hierarchy: flags & TH_ATTRIBUTE_HIERARCHY != 0,
            is_time_hierarchy: flags & TH_TIME_HIERARCHY != 0,
            filter_inclusive,
            is_key_attribute_hierarchy: flags & TH_KEY_ATTRIBUTE_HIERARCHY != 0,
            is_kpi: flags & TH_KPI != 0,
            axis,
            pivot_field_index,
            axis_field_count,
            drag_to_row,
            drag_to_column,
            drag_to_page,
            drag_to_data: drag & DRAG_TO_DATA != 0,
            drag_to_hide: drag & DRAG_TO_HIDE != 0,
            unique_name,
            display_name,
            default_member,
            all_member,
            dimension,
            level_fields,
            hidden_member_sets,
        })
    }

    /// Serialize back to a complete `SXTH` record payload. The same
    /// cross-field rules as [`XlsPivotHierarchy::parse`] are enforced.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        if self.is_measure
            && (self.is_named_set || self.drag_to_row || self.drag_to_column || self.drag_to_page)
        {
            return Err(XlsError::InvalidData(
                "SXTH measure cannot be a named set or drag to row/column/page".to_string(),
            ));
        }
        if self.is_measure && !self.dimension.is_empty() {
            return Err(XlsError::InvalidData(
                "SXTH measure must have an empty dimension".to_string(),
            ));
        }
        if !self.level_fields.is_empty() && !self.axis.row && !self.axis.column {
            return Err(XlsError::InvalidData(
                "SXTH level_fields must be empty off the row/column axes".to_string(),
            ));
        }
        if self.axis_field_count < 0 {
            return Err(XlsError::InvalidData(
                "SXTH axis_field_count must be non-negative".to_string(),
            ));
        }
        let expected_axis_fields = if self.all_member.is_empty() {
            self.level_fields.len() as i64
        } else {
            self.level_fields.len() as i64 - 1
        };
        if i64::from(self.axis_field_count) != expected_axis_fields {
            return Err(XlsError::InvalidData(
                "SXTH axis_field_count does not match level_fields and all_member".to_string(),
            ));
        }
        if self.filter_inclusive && !self.hidden_member_sets.is_empty() {
            return Err(XlsError::InvalidData(
                "SXTH hidden_member_sets must be empty for inclusive filters".to_string(),
            ));
        }
        if !self.hidden_member_sets.is_empty() && self.level_fields.is_empty() {
            return Err(XlsError::InvalidData(
                "SXTH hidden_member_sets require non-empty level_fields".to_string(),
            ));
        }
        check_string_len(SXTH_RECORD_TYPE, &self.unique_name, "unique_name", false)?;
        check_string_len(SXTH_RECORD_TYPE, &self.display_name, "display_name", false)?;
        check_string_len(SXTH_RECORD_TYPE, &self.default_member, "default_member", true)?;
        check_string_len(SXTH_RECORD_TYPE, &self.all_member, "all_member", true)?;
        check_string_len(SXTH_RECORD_TYPE, &self.dimension, "dimension", true)?;
        for &field in &self.level_fields {
            if field < -1 {
                return Err(XlsError::InvalidData(
                    "SXTH level_fields element must be -1 or a pivot field index".to_string(),
                ));
            }
        }

        let mut flags = 0u32;
        if self.is_measure {
            flags |= TH_MEASURE;
        }
        if self.outline_mode {
            flags |= TH_OUTLINE_MODE;
        }
        if self.multiple_page_items {
            flags |= TH_MULTIPLE_PAGE_ITEMS;
        }
        if self.subtotal_at_top {
            flags |= TH_SUBTOTAL_AT_TOP;
        }
        if self.is_named_set {
            flags |= TH_NAMED_SET;
        }
        if self.hidden_from_field_list {
            flags |= TH_HIDDEN_FROM_FIELD_LIST;
        }
        if self.is_attribute_hierarchy {
            flags |= TH_ATTRIBUTE_HIERARCHY;
        }
        if self.is_time_hierarchy {
            flags |= TH_TIME_HIERARCHY;
        }
        if self.filter_inclusive {
            flags |= TH_FILTER_INCLUSIVE;
        }
        if self.is_key_attribute_hierarchy {
            flags |= TH_KEY_ATTRIBUTE_HIERARCHY;
        }
        if self.is_kpi {
            flags |= TH_KPI;
        }
        let mut drag = 0u16;
        if self.drag_to_row {
            drag |= DRAG_TO_ROW;
        }
        if self.drag_to_column {
            drag |= DRAG_TO_COLUMN;
        }
        if self.drag_to_page {
            drag |= DRAG_TO_PAGE;
        }
        if self.drag_to_data {
            drag |= DRAG_TO_DATA;
        }
        if self.drag_to_hide {
            drag |= DRAG_TO_HIDE;
        }

        let mut payload = Vec::new();
        payload.extend_from_slice(&SXTH_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&self.axis.bits().to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
        payload.extend_from_slice(&self.pivot_field_index.to_le_bytes());
        payload.extend_from_slice(&self.axis_field_count.to_le_bytes());
        payload.extend_from_slice(&drag.to_le_bytes());
        append_olap_string(SXTH_RECORD_TYPE, &self.unique_name, &mut payload)?;
        append_olap_string(SXTH_RECORD_TYPE, &self.display_name, &mut payload)?;
        append_olap_string(SXTH_RECORD_TYPE, &self.default_member, &mut payload)?;
        append_olap_string(SXTH_RECORD_TYPE, &self.all_member, &mut payload)?;
        append_olap_string(SXTH_RECORD_TYPE, &self.dimension, &mut payload)?;
        payload.extend_from_slice(&(self.level_fields.len() as u32).to_le_bytes());
        for &field in &self.level_fields {
            payload.extend_from_slice(&field.to_le_bytes());
        }
        payload.extend_from_slice(&(self.hidden_member_sets.len() as u32).to_le_bytes());
        if !self.level_fields.is_empty() {
            for set in &self.hidden_member_sets {
                payload.extend_from_slice(&(set.member_names.len() as u32).to_le_bytes());
                for name in &set.member_names {
                    check_string_len(SXTH_RECORD_TYPE, name, "hidden member name", true)?;
                    append_olap_string(SXTH_RECORD_TYPE, name, &mut payload)?;
                }
            }
        }
        Ok(payload)
    }
}

// ---------------------------------------------------------------------------
// SXPIEx (MS-XLS 2.4.299)
// ---------------------------------------------------------------------------

/// Typed `SXPIEx` record content (MS-XLS 2.4.299): the OLAP extension of
/// one page-axis entry of a PivotTable view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPivotPageItemOlapExt {
    /// Pivot hierarchy index of the hierarchy on the page axis (`isxth`).
    pub hierarchy_index: u32,
    /// Unique name of the OLAP member used for filtering (`stUnique`), at
    /// most 255 characters.
    pub unique_name: String,
    /// Caption of the OLAP member (`stDisplay`), at most 255 characters.
    pub display_name: String,
}

impl XlsPivotPageItemOlapExt {
    /// Parse an `SXPIEx` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let body = frt_body(data, SXPI_EX_RECORD_TYPE, SXPI_EX_FRT_RT)?;
        let hierarchy_index = read_u32(slice_at(body, 0, 4, SXPI_EX_RECORD_TYPE)?, 0);
        let (unique_name, used) = parse_olap_string(&body[4..], SXPI_EX_RECORD_TYPE)?;
        let tail = &body[4 + used..];
        let (display_name, used) = parse_olap_string(tail, SXPI_EX_RECORD_TYPE)?;
        if used != tail.len() {
            return Err(invalid(SXPI_EX_RECORD_TYPE, "trailing bytes after SXPIEx"));
        }
        Ok(XlsPivotPageItemOlapExt {
            hierarchy_index,
            unique_name,
            display_name,
        })
    }

    /// Serialize back to a complete `SXPIEx` record payload.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        check_string_len(SXPI_EX_RECORD_TYPE, &self.unique_name, "unique_name", true)?;
        check_string_len(SXPI_EX_RECORD_TYPE, &self.display_name, "display_name", true)?;
        let mut payload = Vec::new();
        payload.extend_from_slice(&SXPI_EX_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&self.hierarchy_index.to_le_bytes());
        append_olap_string(SXPI_EX_RECORD_TYPE, &self.unique_name, &mut payload)?;
        append_olap_string(SXPI_EX_RECORD_TYPE, &self.display_name, &mut payload)?;
        Ok(payload)
    }
}

// ---------------------------------------------------------------------------
// SXVDTEx (MS-XLS 2.4.311)
// ---------------------------------------------------------------------------

/// An `SXVIFlags` structure (MS-XLS 2.5.263): additional OLAP properties
/// of one pivot item.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct XlsPivotItemOlapFlags {
    /// Whether child elements of this item are collapsed
    /// (`fDrilledMember`).
    pub drilled_member: bool,
    /// Whether the item has child OLAP members (`fHasChildren`).
    pub has_children: bool,
    /// Whether the subnodes of this item are collapsed
    /// (`fCollapsedMember`).
    pub collapsed_member: bool,
    /// Whether `has_children` is considered correct (`fHasChildrenEst`).
    pub has_children_estimated: bool,
    /// Whether the item is selected for OLAP manual filtering
    /// (`fOlapFilterSelected`).
    pub olap_filter_selected: bool,
}

impl XlsPivotItemOlapFlags {
    fn from_bits(record_type: u16, bits: u16) -> XlsResult<Self> {
        if bits & VI_RESERVED != 0 {
            return Err(invalid(record_type, "SXVIFlags reserved bits set"));
        }
        Ok(XlsPivotItemOlapFlags {
            drilled_member: bits & VI_DRILLED_MEMBER != 0,
            has_children: bits & VI_HAS_CHILDREN != 0,
            collapsed_member: bits & VI_COLLAPSED_MEMBER != 0,
            has_children_estimated: bits & VI_HAS_CHILDREN_EST != 0,
            olap_filter_selected: bits & VI_OLAP_FILTER_SELECTED != 0,
        })
    }

    fn bits(self) -> u16 {
        let mut bits = 0u16;
        if self.drilled_member {
            bits |= VI_DRILLED_MEMBER;
        }
        if self.has_children {
            bits |= VI_HAS_CHILDREN;
        }
        if self.collapsed_member {
            bits |= VI_COLLAPSED_MEMBER;
        }
        if self.has_children_estimated {
            bits |= VI_HAS_CHILDREN_EST;
        }
        if self.olap_filter_selected {
            bits |= VI_OLAP_FILTER_SELECTED;
        }
        bits
    }
}

/// Typed `SXVDTEx` record content (MS-XLS 2.4.311): the OLAP extension of
/// one pivot field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPivotFieldOlapExt {
    /// Whether the sort order is determined by the OLAP source data
    /// (`fTensorSort`).
    pub tensor_sort: bool,
    /// Whether all pivot items of this field are expanded
    /// (`fDrilledLevel`).
    pub drilled_level: bool,
    /// Whether this attribute hierarchy is expanded by default
    /// (`fItemsDrilledByDefault`).
    pub items_drilled_by_default: bool,
    /// Whether this member property field is displayed in the report
    /// (`fMemPropDisplayInReport`).
    pub member_property_in_report: bool,
    /// Whether this member property field is displayed in a ToolTip
    /// (`fMemPropDisplayInTip`).
    pub member_property_in_tip: bool,
    /// Whether member property captions replace pivot item captions
    /// (`fMemPropDisplayInCaption`).
    pub member_property_in_caption: bool,
    /// The pivot hierarchy this field is associated with (`isxth`): a
    /// pivot hierarchy index, or -1 when the field is not part of a pivot
    /// hierarchy.
    pub hierarchy_index: i16,
    /// Zero-based index of the associated OLAP level (`isxtl`).
    pub olap_level_index: i32,
    /// Additional properties of the pivot items (`rgsxvi`); one element
    /// per pivot item of this field.
    pub item_flags: Vec<XlsPivotItemOlapFlags>,
}

impl XlsPivotFieldOlapExt {
    /// Parse an `SXVDTEx` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let body = frt_body(data, SXVDT_EX_RECORD_TYPE, SXVDT_EX_FRT_RT)?;
        let fixed = slice_at(body, 0, 12, SXVDT_EX_RECORD_TYPE)?;
        let flags = read_u16(fixed, 0);
        if flags & VDT_RESERVED != 0 {
            return Err(invalid(SXVDT_EX_RECORD_TYPE, "SXVDTEx reserved bits set"));
        }
        let hierarchy_index = read_i16(fixed, 2);
        if hierarchy_index < -1 {
            return Err(invalid(
                SXVDT_EX_RECORD_TYPE,
                "SXVDTEx isxth must be -1 or a pivot hierarchy index",
            ));
        }
        let olap_level_index = read_i32(fixed, 4);
        let item_count = read_i32(fixed, 8);
        if item_count < 0 {
            return Err(invalid(SXVDT_EX_RECORD_TYPE, "SXVDTEx csxvi must be non-negative"));
        }
        let item_count = item_count as usize;
        let items = slice_at(body, 12, item_count * 2, SXVDT_EX_RECORD_TYPE)?;
        if 12 + item_count * 2 != body.len() {
            return Err(invalid(
                SXVDT_EX_RECORD_TYPE,
                "SXVDTEx csxvi does not match the record size",
            ));
        }
        let mut item_flags = Vec::with_capacity(item_count);
        for chunk in items.chunks_exact(2) {
            item_flags.push(XlsPivotItemOlapFlags::from_bits(
                SXVDT_EX_RECORD_TYPE,
                u16::from_le_bytes([chunk[0], chunk[1]]),
            )?);
        }
        Ok(XlsPivotFieldOlapExt {
            tensor_sort: flags & VDT_TENSOR_SORT != 0,
            drilled_level: flags & VDT_DRILLED_LEVEL != 0,
            items_drilled_by_default: flags & VDT_ITEMS_DRILLED_BY_DEFAULT != 0,
            member_property_in_report: flags & VDT_MEMBER_PROPERTY_IN_REPORT != 0,
            member_property_in_tip: flags & VDT_MEMBER_PROPERTY_IN_TIP != 0,
            member_property_in_caption: flags & VDT_MEMBER_PROPERTY_IN_CAPTION != 0,
            hierarchy_index,
            olap_level_index,
            item_flags,
        })
    }

    /// Serialize back to a complete `SXVDTEx` record payload.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        if self.hierarchy_index < -1 {
            return Err(XlsError::InvalidData(
                "SXVDTEx hierarchy_index must be -1 or a pivot hierarchy index".to_string(),
            ));
        }
        let mut flags = 0u16;
        if self.tensor_sort {
            flags |= VDT_TENSOR_SORT;
        }
        if self.drilled_level {
            flags |= VDT_DRILLED_LEVEL;
        }
        if self.items_drilled_by_default {
            flags |= VDT_ITEMS_DRILLED_BY_DEFAULT;
        }
        if self.member_property_in_report {
            flags |= VDT_MEMBER_PROPERTY_IN_REPORT;
        }
        if self.member_property_in_tip {
            flags |= VDT_MEMBER_PROPERTY_IN_TIP;
        }
        if self.member_property_in_caption {
            flags |= VDT_MEMBER_PROPERTY_IN_CAPTION;
        }
        let mut payload = Vec::with_capacity(FRT_HEADER_OLD_LEN + 12 + self.item_flags.len() * 2);
        payload.extend_from_slice(&SXVDT_EX_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&self.hierarchy_index.to_le_bytes());
        payload.extend_from_slice(&self.olap_level_index.to_le_bytes());
        payload.extend_from_slice(&(self.item_flags.len() as i32).to_le_bytes());
        for item in &self.item_flags {
            payload.extend_from_slice(&item.bits().to_le_bytes());
        }
        Ok(payload)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a compressed XLUnicodeString.
    fn olap_string(text: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(text.len() as u16).to_le_bytes());
        out.push(0u8); // fHighByte = 0
        out.extend_from_slice(text.as_bytes());
        out
    }

    fn frt_header(rt: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&rt.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out
    }

    // -- SXViewEx -----------------------------------------------------------

    fn view_ex_payload(hierarchies: i32, pages: i32, fields: i32, future: &[u8]) -> Vec<u8> {
        let mut payload = frt_header(SX_VIEW_EX_FRT_RT);
        payload.extend_from_slice(&hierarchies.to_le_bytes());
        payload.extend_from_slice(&pages.to_le_bytes());
        payload.extend_from_slice(&fields.to_le_bytes());
        payload.extend_from_slice(&(future.len() as u32).to_le_bytes());
        payload.extend_from_slice(future);
        payload
    }

    #[test]
    fn view_ex_parses() {
        let header = XlsPivotViewOlapHeader::parse(&view_ex_payload(2, 1, 3, &[0xDE, 0xAD]))
            .expect("parse");
        assert_eq!(header.hierarchy_count, 2);
        assert_eq!(header.page_extension_count, 1);
        assert_eq!(header.field_extension_count, 3);
        assert_eq!(header.future_bytes, vec![0xDE, 0xAD]);
    }

    #[test]
    fn view_ex_rejects_bad_header_and_counts() {
        // Wrong FrtHeaderOld.rt (the record type itself is NOT the rt here).
        assert!(XlsPivotViewOlapHeader::parse(&view_ex_payload(1, 0, 0, &[])[..])
            .map(|_| ())
            .is_ok());
        let mut payload = view_ex_payload(1, 0, 0, &[]);
        payload[0] = 0x0E; // rt = 0x080E (the record type) is wrong per spec
        payload[1] = 0x08;
        assert!(XlsPivotViewOlapHeader::parse(&payload).is_err());

        // grbitFrt must be zero.
        let mut payload = view_ex_payload(1, 0, 0, &[]);
        payload[2] = 0x01;
        assert!(XlsPivotViewOlapHeader::parse(&payload).is_err());

        // csxth must be at least 1.
        assert!(XlsPivotViewOlapHeader::parse(&view_ex_payload(0, 0, 0, &[])).is_err());
        // Negative counts are illegal.
        assert!(XlsPivotViewOlapHeader::parse(&view_ex_payload(1, -1, 0, &[])).is_err());
        assert!(XlsPivotViewOlapHeader::parse(&view_ex_payload(1, 0, -1, &[])).is_err());
    }

    #[test]
    fn view_ex_rejects_future_blob_defects() {
        // cbFuture larger than the record.
        let mut payload = view_ex_payload(1, 0, 0, &[1, 2]);
        payload.truncate(payload.len() - 1);
        assert!(XlsPivotViewOlapHeader::parse(&payload).is_err());

        // cbFuture beyond the 1024-byte bound.
        let mut payload = view_ex_payload(1, 0, 0, &[]);
        payload[16] = 0x01; // cbFuture = 1025
        payload[17] = 0x04;
        assert!(XlsPivotViewOlapHeader::parse(&payload).is_err());

        // Truncated fixed part.
        assert!(XlsPivotViewOlapHeader::parse(&[]).is_err());
        assert!(XlsPivotViewOlapHeader::parse(&frt_header(SX_VIEW_EX_FRT_RT)).is_err());
    }

    #[test]
    fn view_ex_round_trips() {
        for future in [Vec::new(), vec![0xAA; 17]] {
            let header = XlsPivotViewOlapHeader {
                hierarchy_count: 1,
                page_extension_count: 2,
                field_extension_count: 4,
                future_bytes: future,
            };
            let payload = header.to_payload().expect("serialize");
            assert_eq!(XlsPivotViewOlapHeader::parse(&payload).expect("re-parse"), header);
        }
        // Zero hierarchies cannot be written.
        let bad = XlsPivotViewOlapHeader {
            hierarchy_count: 0,
            page_extension_count: 0,
            field_extension_count: 0,
            future_bytes: Vec::new(),
        };
        assert!(bad.to_payload().is_err());
        // An oversized future blob cannot be written.
        let bad = XlsPivotViewOlapHeader {
            hierarchy_count: 1,
            page_extension_count: 0,
            field_extension_count: 0,
            future_bytes: vec![0u8; MAX_FUTURE_BYTES + 1],
        };
        assert!(bad.to_payload().is_err());
    }

    // -- SXTH ---------------------------------------------------------------

    struct HierarchyBuilder {
        flags: u32,
        axis: u16,
        reserved: u16,
        isxvd: i32,
        csxvd_xl: i32,
        drag: u16,
        unique: String,
        display: String,
        default_member: String,
        all_member: String,
        dimension: String,
        levels: Vec<i32>,
        hidden_set_count: u32,
        hidden_sets: Vec<Vec<String>>,
    }

    impl HierarchyBuilder {
        fn new() -> Self {
            HierarchyBuilder {
                flags: TH_FILTER_INCLUSIVE | TH_ATTRIBUTE_HIERARCHY,
                axis: AXIS_ROW,
                reserved: 0,
                isxvd: 3,
                csxvd_xl: 2,
                drag: DRAG_TO_ROW | DRAG_TO_DATA | DRAG_TO_HIDE,
                unique: "[Product].[Category]".to_string(),
                display: "Category".to_string(),
                default_member: "[Product].[Category].&[1]".to_string(),
                all_member: "[Product].[Category].[All]".to_string(),
                dimension: "[Product]".to_string(),
                levels: vec![3, -1, 5],
                hidden_set_count: 0,
                hidden_sets: Vec::new(),
            }
        }

        fn build(self) -> Vec<u8> {
            let mut payload = frt_header(SXTH_FRT_RT);
            payload.extend_from_slice(&self.flags.to_le_bytes());
            payload.extend_from_slice(&self.axis.to_le_bytes());
            payload.extend_from_slice(&self.reserved.to_le_bytes());
            payload.extend_from_slice(&self.isxvd.to_le_bytes());
            payload.extend_from_slice(&self.csxvd_xl.to_le_bytes());
            payload.extend_from_slice(&self.drag.to_le_bytes());
            payload.extend_from_slice(&olap_string(&self.unique));
            payload.extend_from_slice(&olap_string(&self.display));
            payload.extend_from_slice(&olap_string(&self.default_member));
            payload.extend_from_slice(&olap_string(&self.all_member));
            payload.extend_from_slice(&olap_string(&self.dimension));
            payload.extend_from_slice(&(self.levels.len() as u32).to_le_bytes());
            for level in &self.levels {
                payload.extend_from_slice(&level.to_le_bytes());
            }
            payload.extend_from_slice(&self.hidden_set_count.to_le_bytes());
            for set in &self.hidden_sets {
                payload.extend_from_slice(&(set.len() as u32).to_le_bytes());
                for name in set {
                    payload.extend_from_slice(&olap_string(name));
                }
            }
            payload
        }
    }

    #[test]
    fn hierarchy_parses() {
        let mut builder = HierarchyBuilder::new();
        builder.hidden_set_count = 2;
        builder.hidden_sets = vec![vec!["&[4]".to_string()], vec!["&[7]".to_string(), "&[9]".to_string()]];
        // Inclusive filter forbids hidden sets; flip to exclusive.
        builder.flags &= !TH_FILTER_INCLUSIVE;
        let hierarchy = XlsPivotHierarchy::parse(&builder.build()).expect("parse");
        assert!(!hierarchy.is_measure);
        assert!(hierarchy.is_attribute_hierarchy);
        assert!(!hierarchy.filter_inclusive);
        assert!(!hierarchy.is_named_set);
        assert!(!hierarchy.is_kpi);
        assert_eq!(
            hierarchy.axis,
            XlsPivotHierarchyAxis { row: true, column: false, page: false, data: false }
        );
        assert_eq!(hierarchy.pivot_field_index, 3);
        assert_eq!(hierarchy.axis_field_count, 2);
        assert!(hierarchy.drag_to_row);
        assert!(!hierarchy.drag_to_column);
        assert!(hierarchy.drag_to_data);
        assert!(hierarchy.drag_to_hide);
        assert_eq!(hierarchy.unique_name, "[Product].[Category]");
        assert_eq!(hierarchy.display_name, "Category");
        assert_eq!(hierarchy.default_member, "[Product].[Category].&[1]");
        assert_eq!(hierarchy.all_member, "[Product].[Category].[All]");
        assert_eq!(hierarchy.dimension, "[Product]");
        assert_eq!(hierarchy.level_fields, vec![3, -1, 5]);
        assert_eq!(hierarchy.hidden_member_sets.len(), 2);
        assert_eq!(hierarchy.hidden_member_sets[0].member_names, vec!["&[4]".to_string()]);
        assert_eq!(
            hierarchy.hidden_member_sets[1].member_names,
            vec!["&[7]".to_string(), "&[9]".to_string()]
        );
    }

    #[test]
    fn hierarchy_parses_empty_all_member_relation() {
        // stAll empty => csxvdXl MUST equal cisxvd.
        let mut builder = HierarchyBuilder::new();
        builder.all_member = String::new();
        builder.csxvd_xl = 3;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_ok());
        // stAll non-empty => csxvdXl MUST equal cisxvd - 1 (3 - 1 = 2 ok, 3 bad).
        let mut builder = HierarchyBuilder::new();
        builder.csxvd_xl = 3;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());
    }

    #[test]
    fn hierarchy_rejects_header_and_reserved_defects() {
        // Wrong FrtHeaderOld.rt.
        let mut payload = HierarchyBuilder::new().build();
        payload[0] = 0xDB; // record type is NOT the mandated rt
        payload[1] = 0x00;
        assert!(XlsPivotHierarchy::parse(&payload).is_err());

        // Reserved word after sxaxis must be zero.
        let mut builder = HierarchyBuilder::new();
        builder.reserved = 1;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // SXAxis reserved bits must be zero.
        let mut builder = HierarchyBuilder::new();
        builder.axis |= 0x0010;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Negative csxvdXl.
        let mut builder = HierarchyBuilder::new();
        builder.csxvd_xl = -1;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Truncation and trailing bytes.
        assert!(XlsPivotHierarchy::parse(&[]).is_err());
        let payload = HierarchyBuilder::new().build();
        assert!(XlsPivotHierarchy::parse(&payload[..payload.len() - 1]).is_err());
        let mut payload = HierarchyBuilder::new().build();
        payload.push(0);
        assert!(XlsPivotHierarchy::parse(&payload).is_err());
    }

    #[test]
    fn hierarchy_rejects_measure_violations() {
        // Measure + named set.
        let mut builder = HierarchyBuilder::new();
        builder.flags |= TH_MEASURE | TH_NAMED_SET;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Measure + drag to page.
        let mut builder = HierarchyBuilder::new();
        builder.flags |= TH_MEASURE;
        builder.drag |= DRAG_TO_PAGE;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Measure with a non-empty dimension.
        let mut builder = HierarchyBuilder::new();
        builder.flags |= TH_MEASURE;
        builder.drag &= !(DRAG_TO_ROW);
        builder.axis = AXIS_DATA;
        builder.levels = Vec::new();
        builder.csxvd_xl = -1; // also invalid on its own
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());
        let mut builder = HierarchyBuilder::new();
        builder.flags |= TH_MEASURE;
        builder.drag &= !DRAG_TO_ROW;
        builder.axis = AXIS_DATA;
        builder.levels = Vec::new();
        builder.all_member = String::new();
        builder.csxvd_xl = 0;
        builder.dimension = String::new();
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_ok());
    }

    #[test]
    fn hierarchy_rejects_axis_and_filter_violations() {
        // cisxvd > 0 while on neither row nor column axis.
        let mut builder = HierarchyBuilder::new();
        builder.axis = AXIS_PAGE;
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Inclusive filter with hidden member sets.
        let mut builder = HierarchyBuilder::new();
        builder.hidden_set_count = 1;
        builder.hidden_sets = vec![vec!["&[4]".to_string()]];
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // Empty stUnique / stDisplay.
        let mut builder = HierarchyBuilder::new();
        builder.unique = String::new();
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());
        let mut builder = HierarchyBuilder::new();
        builder.display = String::new();
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());

        // rgisxvd element below -1.
        let mut builder = HierarchyBuilder::new();
        builder.levels = vec![3, -2, 5];
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());
    }

    #[test]
    fn hierarchy_rejects_string_option_defects() {
        // Reserved option bits in an XLUnicodeString.
        let payload = HierarchyBuilder::new().build();
        // stUnique option byte: frt(4) + fixed(18) + cch(2).
        let mut payload2 = payload.clone();
        payload2[4 + 18 + 2] = 0x02;
        assert!(XlsPivotHierarchy::parse(&payload2).is_err());

        // String longer than 255 characters.
        let mut builder = HierarchyBuilder::new();
        builder.unique = "x".repeat(300);
        assert!(XlsPivotHierarchy::parse(&builder.build()).is_err());
    }

    #[test]
    fn hierarchy_round_trips() {
        let hierarchy = XlsPivotHierarchy {
            is_measure: false,
            outline_mode: true,
            multiple_page_items: true,
            subtotal_at_top: false,
            is_named_set: false,
            hidden_from_field_list: true,
            is_attribute_hierarchy: true,
            is_time_hierarchy: false,
            filter_inclusive: false,
            is_key_attribute_hierarchy: true,
            is_kpi: false,
            axis: XlsPivotHierarchyAxis { row: true, column: true, page: false, data: false },
            pivot_field_index: -1,
            axis_field_count: 1,
            drag_to_row: true,
            drag_to_column: true,
            drag_to_page: true,
            drag_to_data: false,
            drag_to_hide: true,
            unique_name: "[Date].[Calendar]".to_string(),
            display_name: "Calendar €".to_string(), // forces the wide string path
            default_member: String::new(),
            all_member: "[Date].[Calendar].[All]".to_string(),
            dimension: "[Date]".to_string(),
            level_fields: vec![0, 2],
            hidden_member_sets: vec![XlsHiddenMemberSet {
                member_names: vec!["&[2001]".to_string()],
            }],
        };
        let payload = hierarchy.to_payload().expect("serialize");
        assert_eq!(XlsPivotHierarchy::parse(&payload).expect("re-parse"), hierarchy);

        // A measure hierarchy with no level fields.
        let measure = XlsPivotHierarchy {
            is_measure: true,
            outline_mode: false,
            multiple_page_items: false,
            subtotal_at_top: false,
            is_named_set: false,
            hidden_from_field_list: false,
            is_attribute_hierarchy: false,
            is_time_hierarchy: false,
            filter_inclusive: false,
            is_key_attribute_hierarchy: false,
            is_kpi: false,
            axis: XlsPivotHierarchyAxis { row: false, column: false, page: false, data: true },
            pivot_field_index: 7,
            axis_field_count: 0,
            drag_to_row: false,
            drag_to_column: false,
            drag_to_page: false,
            drag_to_data: true,
            drag_to_hide: false,
            unique_name: "[Measures].[Sales]".to_string(),
            display_name: "Sales".to_string(),
            default_member: String::new(),
            all_member: String::new(),
            dimension: String::new(),
            level_fields: Vec::new(),
            hidden_member_sets: Vec::new(),
        };
        let payload = measure.to_payload().expect("serialize");
        assert_eq!(XlsPivotHierarchy::parse(&payload).expect("re-parse"), measure);
    }

    #[test]
    fn hierarchy_serialize_rejects_inconsistent_fields() {
        let base = XlsPivotHierarchy::parse(&HierarchyBuilder::new().build()).expect("parse");

        let mut bad = base.clone();
        bad.unique_name = String::new();
        assert!(bad.to_payload().is_err());

        let mut bad = base.clone();
        bad.display_name = "y".repeat(256);
        assert!(bad.to_payload().is_err());

        let mut bad = base.clone();
        bad.is_measure = true; // dimension is non-empty
        assert!(bad.to_payload().is_err());

        let mut bad = base.clone();
        bad.axis_field_count = 99;
        assert!(bad.to_payload().is_err());

        let mut bad = base.clone();
        bad.filter_inclusive = true;
        bad.hidden_member_sets = vec![XlsHiddenMemberSet { member_names: vec!["x".to_string()] }];
        assert!(bad.to_payload().is_err());

        let mut bad = base.clone();
        bad.level_fields = Vec::new(); // breaks the csxvdXl relation
        assert!(bad.to_payload().is_err());
    }

    // -- SXPIEx -------------------------------------------------------------

    fn page_ext_payload(hierarchy: u32, unique: &str, display: &str) -> Vec<u8> {
        let mut payload = frt_header(SXPI_EX_FRT_RT);
        payload.extend_from_slice(&hierarchy.to_le_bytes());
        payload.extend_from_slice(&olap_string(unique));
        payload.extend_from_slice(&olap_string(display));
        payload
    }

    #[test]
    fn page_ext_parses() {
        let ext =
            XlsPivotPageItemOlapExt::parse(&page_ext_payload(1, "[Product].&[3]", "Bikes")).expect("parse");
        assert_eq!(ext.hierarchy_index, 1);
        assert_eq!(ext.unique_name, "[Product].&[3]");
        assert_eq!(ext.display_name, "Bikes");
    }

    #[test]
    fn page_ext_rejects_defects() {
        // Wrong FrtHeaderOld.rt (record type 0x080C is NOT the mandated rt).
        let mut payload = page_ext_payload(0, "u", "d");
        payload[0] = 0x0C;
        payload[1] = 0x08;
        assert!(XlsPivotPageItemOlapExt::parse(&payload).is_err());

        // Oversized string, truncation, trailing bytes.
        assert!(XlsPivotPageItemOlapExt::parse(&page_ext_payload(0, &"x".repeat(300), "d")).is_err());
        assert!(XlsPivotPageItemOlapExt::parse(&[]).is_err());
        let payload = page_ext_payload(0, "u", "d");
        assert!(XlsPivotPageItemOlapExt::parse(&payload[..payload.len() - 1]).is_err());
        let mut payload = page_ext_payload(0, "u", "d");
        payload.push(0);
        assert!(XlsPivotPageItemOlapExt::parse(&payload).is_err());
    }

    #[test]
    fn page_ext_round_trips() {
        let ext = XlsPivotPageItemOlapExt {
            hierarchy_index: 0x7FFF,
            unique_name: "[Date].&[2024]".to_string(),
            display_name: "Year 2024 €".to_string(), // wide string path
        };
        let payload = ext.to_payload().expect("serialize");
        assert_eq!(XlsPivotPageItemOlapExt::parse(&payload).expect("re-parse"), ext);

        let bad = XlsPivotPageItemOlapExt {
            hierarchy_index: 0,
            unique_name: "x".repeat(256),
            display_name: String::new(),
        };
        assert!(bad.to_payload().is_err());
    }

    // -- SXVDTEx ------------------------------------------------------------

    fn field_ext_payload(flags: u16, isxth: i16, isxtl: i32, items: &[u16]) -> Vec<u8> {
        let mut payload = frt_header(SXVDT_EX_FRT_RT);
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&isxth.to_le_bytes());
        payload.extend_from_slice(&isxtl.to_le_bytes());
        payload.extend_from_slice(&(items.len() as i32).to_le_bytes());
        for item in items {
            payload.extend_from_slice(&item.to_le_bytes());
        }
        payload
    }

    #[test]
    fn field_ext_parses() {
        let flags = VDT_TENSOR_SORT | VDT_MEMBER_PROPERTY_IN_REPORT;
        let items = [VI_HAS_CHILDREN | VI_OLAP_FILTER_SELECTED, VI_COLLAPSED_MEMBER];
        let ext = XlsPivotFieldOlapExt::parse(&field_ext_payload(flags, 2, 1, &items)).expect("parse");
        assert!(ext.tensor_sort);
        assert!(!ext.drilled_level);
        assert!(ext.member_property_in_report);
        assert!(!ext.member_property_in_tip);
        assert_eq!(ext.hierarchy_index, 2);
        assert_eq!(ext.olap_level_index, 1);
        assert_eq!(ext.item_flags.len(), 2);
        assert!(ext.item_flags[0].has_children);
        assert!(ext.item_flags[0].olap_filter_selected);
        assert!(!ext.item_flags[0].collapsed_member);
        assert!(ext.item_flags[1].collapsed_member);
    }

    #[test]
    fn field_ext_rejects_defects() {
        // Reserved flag bits.
        assert!(XlsPivotFieldOlapExt::parse(&field_ext_payload(0x0040, 0, 0, &[])).is_err());
        // Reserved SXVIFlags bits.
        assert!(XlsPivotFieldOlapExt::parse(&field_ext_payload(0, 0, 0, &[0x0002])).is_err());
        // isxth below -1.
        assert!(XlsPivotFieldOlapExt::parse(&field_ext_payload(0, -2, 0, &[])).is_err());
        // Negative csxvi.
        let mut payload = field_ext_payload(0, 0, 0, &[]);
        payload[12] = 0xFF;
        payload[13] = 0xFF;
        payload[14] = 0xFF;
        payload[15] = 0xFF;
        assert!(XlsPivotFieldOlapExt::parse(&payload).is_err());
        // csxvi not matching the record size.
        let mut payload = field_ext_payload(0, 0, 0, &[0]);
        payload.pop();
        assert!(XlsPivotFieldOlapExt::parse(&payload).is_err());
        let mut payload = field_ext_payload(0, 0, 0, &[]);
        payload.extend_from_slice(&[0, 0]);
        assert!(XlsPivotFieldOlapExt::parse(&payload).is_err());
        // Truncation.
        assert!(XlsPivotFieldOlapExt::parse(&[]).is_err());
        assert!(XlsPivotFieldOlapExt::parse(&frt_header(SXVDT_EX_FRT_RT)).is_err());
    }

    #[test]
    fn field_ext_round_trips() {
        let ext = XlsPivotFieldOlapExt {
            tensor_sort: true,
            drilled_level: true,
            items_drilled_by_default: false,
            member_property_in_report: false,
            member_property_in_tip: true,
            member_property_in_caption: true,
            hierarchy_index: -1,
            olap_level_index: 0x0000_7FFF,
            item_flags: vec![
                XlsPivotItemOlapFlags {
                    drilled_member: true,
                    has_children: true,
                    collapsed_member: false,
                    has_children_estimated: true,
                    olap_filter_selected: false,
                },
                XlsPivotItemOlapFlags::default(),
            ],
        };
        let payload = ext.to_payload().expect("serialize");
        assert_eq!(XlsPivotFieldOlapExt::parse(&payload).expect("re-parse"), ext);

        let bad = XlsPivotFieldOlapExt { hierarchy_index: -2, ..ext };
        assert!(bad.to_payload().is_err());
    }
}
