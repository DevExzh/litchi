//! BIFF payload codecs for PivotTable OLAP extension records.

use crate::error::{Error, Result};

use super::model::{
    HiddenMemberSet, PivotFieldOlapExt, PivotHierarchy, PivotHierarchyAxis, PivotItemOlapFlags,
    PivotPageItemOlapExt, PivotViewOlapHeader,
};
use super::validation;
use super::{
    FRT_HEADER_OLD_LEN, MAX_FUTURE_BYTES, MAX_OLAP_STRING_CHARS, SX_VIEW_EX_FRT_RT,
    SX_VIEW_EX_RECORD_TYPE, SXPI_EX_FRT_RT, SXPI_EX_RECORD_TYPE, SXTH_FRT_RT, SXTH_RECORD_TYPE,
    SXVDT_EX_FRT_RT, SXVDT_EX_RECORD_TYPE,
};

// `XLUnicodeString.fHighByte` and its reserved bits (MS-XLS 2.5.294).
const HIGH_BYTE: u8 = 0x01;
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

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
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
fn slice_at(data: &[u8], offset: usize, len: usize, record_type: u16) -> Result<&[u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| invalid(record_type, "field offset overflow"))?;
    data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

/// Validate an `FrtHeaderOld` (MS-XLS 2.5.136) and return the record body
/// that follows it. The four-byte header values are mandated for each record
/// by the MS-XLS erratum described by the owner facade.
fn frt_body(data: &[u8], record_type: u16, expected_rt: u16) -> Result<&[u8]> {
    let header = slice_at(data, 0, FRT_HEADER_OLD_LEN, record_type)?;
    if read_u16(header, 0) != expected_rt {
        return Err(invalid(record_type, "FrtHeaderOld.rt mismatch"));
    }
    if read_u16(header, 2) != 0 {
        return Err(invalid(record_type, "FrtHeaderOld.grbitFrt must be zero"));
    }
    Ok(&data[FRT_HEADER_OLD_LEN..])
}

/// Parse an `XLUnicodeString` at the start of `data`. Returns the value and
/// the number of bytes consumed.
fn parse_olap_string(data: &[u8], record_type: u16) -> Result<(String, usize)> {
    let header = slice_at(data, 0, 3, record_type)?;
    let char_count = usize::from(read_u16(header, 0));
    if char_count > MAX_OLAP_STRING_CHARS {
        return Err(invalid(
            record_type,
            "XLUnicodeString exceeds 255 characters",
        ));
    }
    let options = header[2];
    if options & STRING_OPTION_RESERVED != 0 {
        return Err(invalid(
            record_type,
            "XLUnicodeString reserved option bits set",
        ));
    }
    let wide = options & HIGH_BYTE != 0;
    let byte_len = char_count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| invalid(record_type, "XLUnicodeString length overflow"))?;
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

/// Serialize an `XLUnicodeString`, compressed when every character is in
/// U+0000..=U+00FF and wide otherwise.
fn append_olap_string(record_type: u16, value: &str, output: &mut Vec<u8>) -> Result<()> {
    let compressible = value.chars().all(|ch| u32::from(ch) <= 0xFF);
    let char_count = value.encode_utf16().count();
    if char_count > MAX_OLAP_STRING_CHARS {
        return Err(Error::InvalidData(format!(
            "record 0x{record_type:04X} XLUnicodeString exceeds {MAX_OLAP_STRING_CHARS} UTF-16 characters"
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

impl PivotViewOlapHeader {
    /// Parse an `SXViewEx` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = frt_body(data, SX_VIEW_EX_RECORD_TYPE, SX_VIEW_EX_FRT_RT)?;
        let fixed = slice_at(body, 0, 16, SX_VIEW_EX_RECORD_TYPE)?;
        let hierarchy_count = read_i32(fixed, 0);
        let page_extension_count = read_i32(fixed, 4);
        let field_extension_count = read_i32(fixed, 8);
        if hierarchy_count < 1 {
            return Err(invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx csxth must be at least 1",
            ));
        }
        if page_extension_count < 0 || field_extension_count < 0 {
            return Err(invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx record counts must be non-negative",
            ));
        }
        let future_len = usize::try_from(read_u32(fixed, 12)).map_err(|_| {
            invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx cbFuture cannot be represented",
            )
        })?;
        if future_len > MAX_FUTURE_BYTES {
            return Err(invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx cbFuture exceeds 1024",
            ));
        }
        let future_bytes = slice_at(body, 16, future_len, SX_VIEW_EX_RECORD_TYPE)?;
        if 16 + future_len != body.len() {
            return Err(invalid(
                SX_VIEW_EX_RECORD_TYPE,
                "SXViewEx cbFuture does not match the record size",
            ));
        }
        let value = PivotViewOlapHeader {
            hierarchy_count: hierarchy_count as u32,
            page_extension_count: page_extension_count as u32,
            field_extension_count: field_extension_count as u32,
            future_bytes: future_bytes.to_vec(),
        };
        validation::validate_view_header(&value, false)?;
        Ok(value)
    }

    /// Serialize back to a complete `SXViewEx` record payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validation::validate_view_header(self, true)?;
        let hierarchy_count = i32::try_from(self.hierarchy_count)
            .map_err(|_| Error::InvalidData("SXViewEx csxth exceeds i32".to_string()))?;
        let page_extension_count = i32::try_from(self.page_extension_count)
            .map_err(|_| Error::InvalidData("SXViewEx csxpi exceeds i32".to_string()))?;
        let field_extension_count = i32::try_from(self.field_extension_count)
            .map_err(|_| Error::InvalidData("SXViewEx csxvdtex exceeds i32".to_string()))?;
        let mut payload = Vec::with_capacity(FRT_HEADER_OLD_LEN + 16 + self.future_bytes.len());
        payload.extend_from_slice(&SX_VIEW_EX_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&hierarchy_count.to_le_bytes());
        payload.extend_from_slice(&page_extension_count.to_le_bytes());
        payload.extend_from_slice(&field_extension_count.to_le_bytes());
        payload.extend_from_slice(&(self.future_bytes.len() as u32).to_le_bytes());
        payload.extend_from_slice(&self.future_bytes);
        Ok(payload)
    }
}

impl PivotHierarchyAxis {
    pub(super) fn from_bits(record_type: u16, bits: u16) -> Result<Self> {
        if bits & AXIS_RESERVED != 0 {
            return Err(invalid(record_type, "SXAxis reserved bits set"));
        }
        Ok(PivotHierarchyAxis {
            row: bits & AXIS_ROW != 0,
            column: bits & AXIS_COLUMN != 0,
            page: bits & AXIS_PAGE != 0,
            data: bits & AXIS_DATA != 0,
        })
    }

    pub(super) fn bits(self) -> u16 {
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

impl PivotHierarchy {
    /// Parse an `SXTH` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = frt_body(data, SXTH_RECORD_TYPE, SXTH_FRT_RT)?;
        let fixed = slice_at(body, 0, 18, SXTH_RECORD_TYPE)?;
        let flags = read_u32(fixed, 0);
        let is_measure = flags & TH_MEASURE != 0;
        let is_named_set = flags & TH_NAMED_SET != 0;
        let filter_inclusive = flags & TH_FILTER_INCLUSIVE != 0;
        let axis = PivotHierarchyAxis::from_bits(SXTH_RECORD_TYPE, read_u16(fixed, 4))?;
        if read_u16(fixed, 6) != 0 {
            return Err(invalid(SXTH_RECORD_TYPE, "SXTH reserved word must be zero"));
        }
        let pivot_field_index = read_i32(fixed, 8);
        let axis_field_count = read_i32(fixed, 12);
        let drag = read_u16(fixed, 16);

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

        let level_count =
            usize::try_from(read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0))
                .map_err(|_| invalid(SXTH_RECORD_TYPE, "SXTH cisxvd cannot be represented"))?;
        offset += 4;
        let mut level_fields = Vec::with_capacity(level_count.min(body.len() / 4));
        for _ in 0..level_count {
            let value = read_i32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0);
            offset += 4;
            level_fields.push(value);
        }

        let hidden_set_count =
            usize::try_from(read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0)).map_err(
                |_| {
                    invalid(
                        SXTH_RECORD_TYPE,
                        "SXTH cHiddenMemberSets cannot be represented",
                    )
                },
            )?;
        offset += 4;
        if hidden_set_count > level_count {
            return Err(invalid(
                SXTH_RECORD_TYPE,
                "SXTH cHiddenMemberSets exceeds cisxvd",
            ));
        }
        let mut hidden_member_sets = Vec::with_capacity(hidden_set_count);
        // rgHiddenMemberSets exists iff cHiddenMemberSets > 0 and cisxvd > 0.
        if hidden_set_count > 0 && level_count > 0 {
            for _ in 0..hidden_set_count {
                let name_count =
                    usize::try_from(read_u32(slice_at(body, offset, 4, SXTH_RECORD_TYPE)?, 0))
                        .map_err(|_| invalid(SXTH_RECORD_TYPE, "hidden member count overflow"))?;
                offset += 4;
                let mut member_names = Vec::with_capacity(name_count.min(body.len() / 3));
                for _ in 0..name_count {
                    let (name, used) = parse_olap_string(&body[offset..], SXTH_RECORD_TYPE)?;
                    offset += used;
                    member_names.push(name);
                }
                hidden_member_sets.push(HiddenMemberSet { member_names });
            }
        }
        if offset != body.len() {
            return Err(invalid(SXTH_RECORD_TYPE, "trailing bytes after SXTH"));
        }

        let value = PivotHierarchy {
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
            drag_to_row: drag & DRAG_TO_ROW != 0,
            drag_to_column: drag & DRAG_TO_COLUMN != 0,
            drag_to_page: drag & DRAG_TO_PAGE != 0,
            drag_to_data: drag & DRAG_TO_DATA != 0,
            drag_to_hide: drag & DRAG_TO_HIDE != 0,
            unique_name,
            display_name,
            default_member,
            all_member,
            dimension,
            level_fields,
            hidden_member_sets,
        };
        validation::validate_hierarchy(&value, false)?;
        Ok(value)
    }

    /// Serialize back to a complete `SXTH` record payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validation::validate_hierarchy(self, true)?;
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

        let level_count = u32::try_from(self.level_fields.len())
            .map_err(|_| Error::InvalidData("SXTH cisxvd exceeds u32".to_string()))?;
        let hidden_set_count = u32::try_from(self.hidden_member_sets.len())
            .map_err(|_| Error::InvalidData("SXTH cHiddenMemberSets exceeds u32".to_string()))?;
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
        payload.extend_from_slice(&level_count.to_le_bytes());
        for &field in &self.level_fields {
            payload.extend_from_slice(&field.to_le_bytes());
        }
        payload.extend_from_slice(&hidden_set_count.to_le_bytes());
        for set in &self.hidden_member_sets {
            let name_count = u32::try_from(set.member_names.len()).map_err(|_| {
                Error::InvalidData("SXTH hidden member count exceeds u32".to_string())
            })?;
            payload.extend_from_slice(&name_count.to_le_bytes());
            for name in &set.member_names {
                append_olap_string(SXTH_RECORD_TYPE, name, &mut payload)?;
            }
        }
        Ok(payload)
    }
}

impl PivotPageItemOlapExt {
    /// Parse an `SXPIEx` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = frt_body(data, SXPI_EX_RECORD_TYPE, SXPI_EX_FRT_RT)?;
        let hierarchy_index = read_u32(slice_at(body, 0, 4, SXPI_EX_RECORD_TYPE)?, 0);
        let (unique_name, used) = parse_olap_string(&body[4..], SXPI_EX_RECORD_TYPE)?;
        let tail = &body[4 + used..];
        let (display_name, used) = parse_olap_string(tail, SXPI_EX_RECORD_TYPE)?;
        if used != tail.len() {
            return Err(invalid(SXPI_EX_RECORD_TYPE, "trailing bytes after SXPIEx"));
        }
        let value = PivotPageItemOlapExt {
            hierarchy_index,
            unique_name,
            display_name,
        };
        validation::validate_page_extension(&value, false)?;
        Ok(value)
    }

    /// Serialize back to a complete `SXPIEx` record payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validation::validate_page_extension(self, true)?;
        let mut payload = Vec::new();
        payload.extend_from_slice(&SXPI_EX_FRT_RT.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        payload.extend_from_slice(&self.hierarchy_index.to_le_bytes());
        append_olap_string(SXPI_EX_RECORD_TYPE, &self.unique_name, &mut payload)?;
        append_olap_string(SXPI_EX_RECORD_TYPE, &self.display_name, &mut payload)?;
        Ok(payload)
    }
}

impl PivotItemOlapFlags {
    pub(super) fn from_bits(record_type: u16, bits: u16) -> Result<Self> {
        if bits & VI_RESERVED != 0 {
            return Err(invalid(record_type, "SXVIFlags reserved bits set"));
        }
        Ok(PivotItemOlapFlags {
            drilled_member: bits & VI_DRILLED_MEMBER != 0,
            has_children: bits & VI_HAS_CHILDREN != 0,
            collapsed_member: bits & VI_COLLAPSED_MEMBER != 0,
            has_children_estimated: bits & VI_HAS_CHILDREN_EST != 0,
            olap_filter_selected: bits & VI_OLAP_FILTER_SELECTED != 0,
        })
    }

    pub(super) fn bits(self) -> u16 {
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

impl PivotFieldOlapExt {
    /// Parse an `SXVDTEx` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = frt_body(data, SXVDT_EX_RECORD_TYPE, SXVDT_EX_FRT_RT)?;
        let fixed = slice_at(body, 0, 12, SXVDT_EX_RECORD_TYPE)?;
        let flags = read_u16(fixed, 0);
        if flags & VDT_RESERVED != 0 {
            return Err(invalid(SXVDT_EX_RECORD_TYPE, "SXVDTEx reserved bits set"));
        }
        let hierarchy_index = read_i16(fixed, 2);
        let olap_level_index = read_i32(fixed, 4);
        let item_count = read_i32(fixed, 8);
        if item_count < 0 {
            return Err(invalid(
                SXVDT_EX_RECORD_TYPE,
                "SXVDTEx csxvi must be non-negative",
            ));
        }
        let item_count = usize::try_from(item_count)
            .map_err(|_| invalid(SXVDT_EX_RECORD_TYPE, "SXVDTEx csxvi overflow"))?;
        let item_bytes = item_count
            .checked_mul(2)
            .ok_or_else(|| invalid(SXVDT_EX_RECORD_TYPE, "SXVDTEx item byte count overflow"))?;
        let items = slice_at(body, 12, item_bytes, SXVDT_EX_RECORD_TYPE)?;
        if 12 + item_bytes != body.len() {
            return Err(invalid(
                SXVDT_EX_RECORD_TYPE,
                "SXVDTEx csxvi does not match the record size",
            ));
        }
        let mut item_flags = Vec::with_capacity(item_count);
        for chunk in items.chunks_exact(2) {
            item_flags.push(PivotItemOlapFlags::from_bits(
                SXVDT_EX_RECORD_TYPE,
                u16::from_le_bytes([chunk[0], chunk[1]]),
            )?);
        }
        let value = PivotFieldOlapExt {
            tensor_sort: flags & VDT_TENSOR_SORT != 0,
            drilled_level: flags & VDT_DRILLED_LEVEL != 0,
            items_drilled_by_default: flags & VDT_ITEMS_DRILLED_BY_DEFAULT != 0,
            member_property_in_report: flags & VDT_MEMBER_PROPERTY_IN_REPORT != 0,
            member_property_in_tip: flags & VDT_MEMBER_PROPERTY_IN_TIP != 0,
            member_property_in_caption: flags & VDT_MEMBER_PROPERTY_IN_CAPTION != 0,
            hierarchy_index,
            olap_level_index,
            item_flags,
        };
        validation::validate_field_extension(&value, false)?;
        Ok(value)
    }

    /// Serialize back to a complete `SXVDTEx` record payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validation::validate_field_extension(self, true)?;
        let item_count = i32::try_from(self.item_flags.len())
            .map_err(|_| Error::InvalidData("SXVDTEx csxvi exceeds i32".to_string()))?;
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
        payload.extend_from_slice(&item_count.to_le_bytes());
        for item in &self.item_flags {
            payload.extend_from_slice(&item.bits().to_le_bytes());
        }
        Ok(payload)
    }
}
