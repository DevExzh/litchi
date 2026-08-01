//! BIFF8 extended range-sort metadata.
//!
//! `SortData` ([MS-XLS] 2.4.264) stores its fixed fields in record 0x0895 and
//! places each `SortCond12` ([MS-XLS] 2.5.242) in a separate `ContinueFrt12`
//! record. This module deliberately models that record group as one value.

use crate::{XlsError, XlsResult};
use std::io::Write;

/// `SortData` record identifier.
pub const SORT_DATA_RECORD_TYPE: u16 = 0x0895;
/// `ContinueFrt12` record identifier.
pub const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087f;

const SORT_DATA_BODY_LEN: usize = 38;
const FRT_HEADER_LEN: usize = 12;
const SORT_CONDITION_FIXED_LEN: usize = 30;
const MAX_CONTINUE_RGB_LEN: usize = 8_212;
const MAX_ROW_INDEX: u32 = 0x000f_ffff;
const MAX_COLUMN_INDEX: u32 = 0x0000_3fff;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidData(message.into())
}

fn read_u16(data: &[u8], offset: usize) -> XlsResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or(XlsError::InvalidLength {
            expected: offset + 2,
            found: data.len(),
        })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> XlsResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or(XlsError::InvalidLength {
            expected: offset + 4,
            found: data.len(),
        })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i32(data: &[u8], offset: usize) -> XlsResult<i32> {
    Ok(read_u32(data, offset)? as i32)
}

/// A validated `RFX` cell range used by extended sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSortRange {
    first_row: u32,
    last_row: u32,
    first_column: u32,
    last_column: u32,
}

impl XlsSortRange {
    /// Creates a range, enforcing the `Rw12`, `Col12`, and ordering bounds.
    pub fn new(
        first_row: u32,
        last_row: u32,
        first_column: u32,
        last_column: u32,
    ) -> XlsResult<Self> {
        if first_row > last_row {
            return Err(invalid("SortData first row exceeds last row"));
        }
        if first_column > last_column {
            return Err(invalid("SortData first column exceeds last column"));
        }
        if last_row > MAX_ROW_INDEX {
            return Err(invalid("SortData row exceeds the Rw12 maximum"));
        }
        if last_column > MAX_COLUMN_INDEX {
            return Err(invalid("SortData column exceeds the Col12 maximum"));
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }

    pub fn first_row(self) -> u32 {
        self.first_row
    }

    pub fn last_row(self) -> u32 {
        self.last_row
    }

    pub fn first_column(self) -> u32 {
        self.first_column
    }

    pub fn last_column(self) -> u32 {
        self.last_column
    }

    fn parse(data: &[u8], offset: usize) -> XlsResult<Self> {
        Self::new(
            read_u32(data, offset)?,
            read_u32(data, offset + 4)?,
            read_u32(data, offset + 8)?,
            read_u32(data, offset + 12)?,
        )
    }

    fn write_to(self, output: &mut Vec<u8>) {
        output.extend_from_slice(&self.first_row.to_le_bytes());
        output.extend_from_slice(&self.last_row.to_le_bytes());
        output.extend_from_slice(&self.first_column.to_le_bytes());
        output.extend_from_slice(&self.last_column.to_le_bytes());
    }
}

/// Whether fields identify rows or columns to reorder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSortOrientation {
    Rows,
    Columns,
}

/// Character-order versus locale-specific alternate sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSortMethod {
    CharacterOrder,
    Alternate,
}

/// Object which owns the sort field (`sfp` and `idParent`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSortParent {
    Sheet,
    Table { id: u32 },
    AutoFilter,
    QueryTable { index: u32 },
}

/// A DXF table index used by color-based sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDifferentialFormatIndex(u32);

impl XlsDifferentialFormatIndex {
    pub const fn new(index: u32) -> Self {
        Self(index)
    }

    pub const fn index(self) -> u32 {
        self.0
    }
}

/// Icon sets allowed by the `KPISets` enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSortIconSet {
    NoIcon,
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}

impl XlsSortIconSet {
    fn code(self) -> u32 {
        match self {
            Self::NoIcon => u32::MAX,
            Self::ThreeArrows => 0,
            Self::ThreeArrowsGray => 1,
            Self::ThreeFlags => 2,
            Self::ThreeTrafficLights1 => 3,
            Self::ThreeTrafficLights2 => 4,
            Self::ThreeSigns => 5,
            Self::ThreeSymbols => 6,
            Self::ThreeSymbols2 => 7,
            Self::FourArrows => 8,
            Self::FourArrowsGray => 9,
            Self::FourRedToBlack => 10,
            Self::FourRating => 11,
            Self::FourTrafficLights => 12,
            Self::FiveArrows => 13,
            Self::FiveArrowsGray => 14,
            Self::FiveRating => 15,
            Self::FiveQuarters => 16,
        }
    }

    fn from_code(code: u32) -> XlsResult<Self> {
        Ok(match code {
            u32::MAX => Self::NoIcon,
            0 => Self::ThreeArrows,
            1 => Self::ThreeArrowsGray,
            2 => Self::ThreeFlags,
            3 => Self::ThreeTrafficLights1,
            4 => Self::ThreeTrafficLights2,
            5 => Self::ThreeSigns,
            6 => Self::ThreeSymbols,
            7 => Self::ThreeSymbols2,
            8 => Self::FourArrows,
            9 => Self::FourArrowsGray,
            10 => Self::FourRedToBlack,
            11 => Self::FourRating,
            12 => Self::FourTrafficLights,
            13 => Self::FiveArrows,
            14 => Self::FiveArrowsGray,
            15 => Self::FiveRating,
            16 => Self::FiveQuarters,
            _ => return Err(invalid("SortCond12 contains an unknown KPISets value")),
        })
    }
}

/// Icon ordinal used by icon-set sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSortIcon {
    NoIcon,
    First,
    Second,
    Third,
    Fourth,
    Fifth,
}

impl XlsSortIcon {
    fn code(self) -> i32 {
        match self {
            Self::NoIcon => -1,
            Self::First => 0,
            Self::Second => 1,
            Self::Third => 2,
            Self::Fourth => 3,
            Self::Fifth => 4,
        }
    }

    fn from_code(code: i32) -> XlsResult<Self> {
        Ok(match code {
            -1 => Self::NoIcon,
            0 => Self::First,
            1 => Self::Second,
            2 => Self::Third,
            3 => Self::Fourth,
            4 => Self::Fifth,
            _ => return Err(invalid("SortCond12 icon index is outside -1 through 4")),
        })
    }
}

/// The criterion used by a `SortCond12` entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsSortOn {
    Values {
        custom_list: Option<String>,
    },
    CellColor {
        differential_format: XlsDifferentialFormatIndex,
    },
    FontColor {
        differential_format: XlsDifferentialFormatIndex,
    },
    Icon {
        set: XlsSortIconSet,
        icon: XlsSortIcon,
    },
}

/// One extended sort condition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSortCondition {
    range: XlsSortRange,
    descending: bool,
    sort_on: XlsSortOn,
}

impl XlsSortCondition {
    pub fn new(range: XlsSortRange, descending: bool, sort_on: XlsSortOn) -> Self {
        Self {
            range,
            descending,
            sort_on,
        }
    }

    pub fn range(&self) -> XlsSortRange {
        self.range
    }

    pub fn is_descending(&self) -> bool {
        self.descending
    }

    pub fn sort_on(&self) -> &XlsSortOn {
        &self.sort_on
    }
}

/// Complete extended sorting metadata represented by one BIFF record group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSortData {
    range: XlsSortRange,
    orientation: XlsSortOrientation,
    case_sensitive: bool,
    method: XlsSortMethod,
    parent: XlsSortParent,
    conditions: Vec<XlsSortCondition>,
}

impl XlsSortData {
    pub fn new(range: XlsSortRange, parent: XlsSortParent) -> Self {
        Self {
            range,
            orientation: XlsSortOrientation::Rows,
            case_sensitive: false,
            method: XlsSortMethod::CharacterOrder,
            parent,
            conditions: Vec::new(),
        }
    }

    pub fn set_orientation(&mut self, orientation: XlsSortOrientation) {
        self.orientation = orientation;
    }

    pub fn set_case_sensitive(&mut self, case_sensitive: bool) {
        self.case_sensitive = case_sensitive;
    }

    pub fn set_method(&mut self, method: XlsSortMethod) {
        self.method = method;
    }

    pub fn add_condition(&mut self, condition: XlsSortCondition) {
        self.conditions.push(condition);
    }

    pub fn range(&self) -> XlsSortRange {
        self.range
    }

    pub fn orientation(&self) -> XlsSortOrientation {
        self.orientation
    }

    pub fn is_case_sensitive(&self) -> bool {
        self.case_sensitive
    }

    pub fn method(&self) -> XlsSortMethod {
        self.method
    }

    pub fn parent(&self) -> XlsSortParent {
        self.parent
    }

    pub fn conditions(&self) -> &[XlsSortCondition] {
        &self.conditions
    }

    /// Writes the `SortData` record followed by one `ContinueFrt12` per condition.
    pub fn write_biff_records<W: Write>(&self, writer: &mut W) -> XlsResult<()> {
        let condition_count = u32::try_from(self.conditions.len())
            .map_err(|_| invalid("SortData has more than u32::MAX conditions"))?;
        let (parent_kind, parent_id) = match self.parent {
            XlsSortParent::Sheet => (0u16, 0),
            XlsSortParent::Table { id } => (1, id),
            XlsSortParent::AutoFilter => (2, 0),
            XlsSortParent::QueryTable { index } => (3, index),
        };
        let flags = u16::from(self.orientation == XlsSortOrientation::Columns)
            | (u16::from(self.case_sensitive) << 1)
            | (u16::from(self.method == XlsSortMethod::Alternate) << 2)
            | (parent_kind << 3);

        write_record_header(writer, SORT_DATA_RECORD_TYPE, SORT_DATA_BODY_LEN)?;
        writer.write_all(&SORT_DATA_RECORD_TYPE.to_le_bytes())?;
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&[0; 8])?;
        writer.write_all(&flags.to_le_bytes())?;
        let mut range = Vec::with_capacity(16);
        self.range.write_to(&mut range);
        writer.write_all(&range)?;
        writer.write_all(&condition_count.to_le_bytes())?;
        writer.write_all(&parent_id.to_le_bytes())?;

        for condition in &self.conditions {
            let body = encode_condition(condition)?;
            write_record_header(
                writer,
                CONTINUE_FRT12_RECORD_TYPE,
                FRT_HEADER_LEN + body.len(),
            )?;
            writer.write_all(&CONTINUE_FRT12_RECORD_TYPE.to_le_bytes())?;
            writer.write_all(&0u16.to_le_bytes())?;
            writer.write_all(&[0; 8])?;
            writer.write_all(&body)?;
        }
        Ok(())
    }
}

fn write_record_header<W: Write>(writer: &mut W, record_type: u16, len: usize) -> XlsResult<()> {
    let len = u16::try_from(len).map_err(|_| invalid("BIFF record body exceeds u16::MAX"))?;
    writer.write_all(&record_type.to_le_bytes())?;
    writer.write_all(&len.to_le_bytes())?;
    Ok(())
}

fn encode_condition(condition: &XlsSortCondition) -> XlsResult<Vec<u8>> {
    let (sort_on, cond_data, custom_list) = match &condition.sort_on {
        XlsSortOn::Values { custom_list } => (0u16, [0u8; 8], custom_list.as_deref()),
        XlsSortOn::CellColor {
            differential_format,
        } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&differential_format.index().to_le_bytes());
            (1, data, None)
        },
        XlsSortOn::FontColor {
            differential_format,
        } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&differential_format.index().to_le_bytes());
            (2, data, None)
        },
        XlsSortOn::Icon { set, icon } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&set.code().to_le_bytes());
            data[4..].copy_from_slice(&icon.code().to_le_bytes());
            (3, data, None)
        },
    };
    let units = custom_list
        .map(|value| value.encode_utf16().collect::<Vec<_>>())
        .unwrap_or_default();
    let char_count = u32::try_from(units.len())
        .map_err(|_| invalid("SortCond12 custom list exceeds u32::MAX UTF-16 code units"))?;
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let string_bytes = if units.is_empty() {
        0
    } else if compressed {
        1 + units.len()
    } else {
        1 + units.len() * 2
    };
    let body_len = SORT_CONDITION_FIXED_LEN
        .checked_add(string_bytes)
        .ok_or_else(|| invalid("SortCond12 encoded length overflow"))?;
    if body_len > MAX_CONTINUE_RGB_LEN {
        return Err(invalid("SortCond12 exceeds the ContinueFrt12 rgb limit"));
    }

    let mut output = Vec::with_capacity(body_len);
    output.extend_from_slice(&((sort_on << 1) | u16::from(condition.descending)).to_le_bytes());
    condition.range.write_to(&mut output);
    output.extend_from_slice(&cond_data);
    output.extend_from_slice(&char_count.to_le_bytes());
    if !units.is_empty() {
        output.push(u8::from(!compressed));
        if compressed {
            output.extend(units.into_iter().map(|unit| unit as u8));
        } else {
            for unit in units {
                output.extend_from_slice(&unit.to_le_bytes());
            }
        }
    }
    Ok(output)
}

/// Parses a `SortData` payload and the following `ContinueFrt12` payloads.
///
/// Inputs exclude the standard four-byte BIFF record headers, consistent with
/// the rest of the XLS record parsers.
pub fn parse_sort_data(base: &[u8], continuations: &[&[u8]]) -> XlsResult<XlsSortData> {
    if base.len() != SORT_DATA_BODY_LEN {
        return Err(XlsError::InvalidLength {
            expected: SORT_DATA_BODY_LEN,
            found: base.len(),
        });
    }
    let echoed_type = read_u16(base, 0)?;
    if echoed_type != SORT_DATA_RECORD_TYPE {
        return Err(XlsError::UnexpectedRecordType {
            expected: SORT_DATA_RECORD_TYPE,
            found: echoed_type,
        });
    }
    if read_u16(base, 2)? & 0x0003 != 0 {
        return Err(invalid("SortData FrtHeader flags violate [MS-XLS] 2.5.135"));
    }
    let flags = read_u16(base, 12)?;
    let parent_kind = (flags >> 3) & 0x0007;
    if parent_kind > 3 {
        return Err(invalid("SortData sfp is outside 0 through 3"));
    }
    let range = XlsSortRange::parse(base, 14)?;
    let condition_count = usize::try_from(read_u32(base, 30)?)
        .map_err(|_| invalid("SortData condition count is not addressable"))?;
    if condition_count != continuations.len() {
        return Err(invalid(format!(
            "SortData declares {condition_count} conditions but {} continuations were supplied",
            continuations.len()
        )));
    }
    let parent_id = read_u32(base, 34)?;
    let parent = match parent_kind {
        0 => XlsSortParent::Sheet,
        1 => XlsSortParent::Table { id: parent_id },
        2 => XlsSortParent::AutoFilter,
        3 => XlsSortParent::QueryTable { index: parent_id },
        _ => unreachable!(),
    };

    let mut conditions = Vec::with_capacity(condition_count);
    for continuation in continuations {
        conditions.push(parse_continuation(continuation)?);
    }
    Ok(XlsSortData {
        range,
        orientation: if flags & 0x0001 != 0 {
            XlsSortOrientation::Columns
        } else {
            XlsSortOrientation::Rows
        },
        case_sensitive: flags & 0x0002 != 0,
        method: if flags & 0x0004 != 0 {
            XlsSortMethod::Alternate
        } else {
            XlsSortMethod::CharacterOrder
        },
        parent,
        conditions,
    })
}

#[derive(Debug)]
struct PendingSortData {
    base: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    expected_conditions: usize,
}

/// Sequential record-group assembler used by the normal worksheet parser.
#[derive(Debug, Default)]
pub(crate) struct SortDataCollector {
    pending: Option<PendingSortData>,
}

impl SortDataCollector {
    pub(crate) fn feed_record(
        &mut self,
        record_type: u16,
        data: &[u8],
    ) -> XlsResult<Option<XlsSortData>> {
        if let Some(pending) = self.pending.as_mut() {
            if record_type != CONTINUE_FRT12_RECORD_TYPE {
                return Err(XlsError::InvalidRecord {
                    record_type,
                    message: format!(
                        "SortData must be followed immediately by {} ContinueFrt12 records",
                        pending.expected_conditions
                    ),
                });
            }
            pending.continuations.push(data.to_vec());
            if pending.continuations.len() > pending.expected_conditions {
                return Err(invalid("SortData received too many ContinueFrt12 records"));
            }
            if pending.continuations.len() == pending.expected_conditions {
                let pending = self.pending.take().expect("pending SortData exists");
                let continuations = pending
                    .continuations
                    .iter()
                    .map(Vec::as_slice)
                    .collect::<Vec<_>>();
                return parse_sort_data(&pending.base, &continuations).map(Some);
            }
            return Ok(None);
        }

        if record_type != SORT_DATA_RECORD_TYPE {
            return Ok(None);
        }
        if data.len() != SORT_DATA_BODY_LEN {
            return Err(XlsError::InvalidLength {
                expected: SORT_DATA_BODY_LEN,
                found: data.len(),
            });
        }
        let expected_conditions = usize::try_from(read_u32(data, 30)?)
            .map_err(|_| invalid("SortData condition count is not addressable"))?;
        if expected_conditions == 0 {
            return parse_sort_data(data, &[]).map(Some);
        }
        self.pending = Some(PendingSortData {
            base: data.to_vec(),
            continuations: Vec::new(),
            expected_conditions,
        });
        Ok(None)
    }

    pub(crate) fn finish(self) -> XlsResult<()> {
        if let Some(pending) = self.pending {
            return Err(XlsError::InvalidRecord {
                record_type: SORT_DATA_RECORD_TYPE,
                message: format!(
                    "worksheet ended after {} of {} SortData conditions",
                    pending.continuations.len(),
                    pending.expected_conditions
                ),
            });
        }
        Ok(())
    }
}

fn parse_continuation(data: &[u8]) -> XlsResult<XlsSortCondition> {
    if data.len() < FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN {
        return Err(XlsError::InvalidLength {
            expected: FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN,
            found: data.len(),
        });
    }
    if data.len() - FRT_HEADER_LEN > MAX_CONTINUE_RGB_LEN {
        return Err(invalid("ContinueFrt12 rgb exceeds 8,212 bytes"));
    }
    let echoed_type = read_u16(data, 0)?;
    if echoed_type != CONTINUE_FRT12_RECORD_TYPE {
        return Err(XlsError::UnexpectedRecordType {
            expected: CONTINUE_FRT12_RECORD_TYPE,
            found: echoed_type,
        });
    }
    let frt_flags = read_u16(data, 2)?;
    if frt_flags & 0x0002 != 0 {
        return Err(invalid("ContinueFrt12 fFrtAlert must be zero"));
    }
    if frt_flags & 0x0001 == 0 && data[4..12].iter().any(|byte| *byte != 0) {
        return Err(invalid(
            "ContinueFrt12 has a reference while fFrtRef is zero",
        ));
    }

    let body = &data[FRT_HEADER_LEN..];
    let flags = read_u16(body, 0)?;
    if flags & 0xffe0 != 0 {
        return Err(invalid("SortCond12 reserved flag bits are nonzero"));
    }
    let sort_on_code = (flags >> 1) & 0x000f;
    let range = XlsSortRange::parse(body, 2)?;
    let data_value = read_u32(body, 18)?;
    let reserved_data = read_u32(body, 22)?;
    let char_count = read_i32(body, 26)?;
    if char_count < 0 {
        return Err(invalid("SortCond12 cchSt is negative"));
    }
    let char_count = char_count as usize;
    let sort_on = match sort_on_code {
        0 => {
            if data_value != 0 || reserved_data != 0 {
                return Err(invalid("value SortCond12 has nonzero CondDataValue fields"));
            }
            let custom_list = parse_custom_list(body, char_count)?;
            XlsSortOn::Values { custom_list }
        },
        1 | 2 => {
            if reserved_data != 0 || char_count != 0 || body.len() != SORT_CONDITION_FIXED_LEN {
                return Err(invalid("color SortCond12 has reserved or trailing data"));
            }
            let differential_format = XlsDifferentialFormatIndex::new(data_value);
            if sort_on_code == 1 {
                XlsSortOn::CellColor {
                    differential_format,
                }
            } else {
                XlsSortOn::FontColor {
                    differential_format,
                }
            }
        },
        3 => {
            if char_count != 0 || body.len() != SORT_CONDITION_FIXED_LEN {
                return Err(invalid(
                    "icon SortCond12 has a custom list or trailing data",
                ));
            }
            XlsSortOn::Icon {
                set: XlsSortIconSet::from_code(data_value)?,
                icon: XlsSortIcon::from_code(reserved_data as i32)?,
            }
        },
        _ => return Err(invalid("SortCond12 sortOn is outside 0 through 3")),
    };
    Ok(XlsSortCondition::new(range, flags & 0x0001 != 0, sort_on))
}

fn parse_custom_list(body: &[u8], char_count: usize) -> XlsResult<Option<String>> {
    if char_count == 0 {
        if body.len() != SORT_CONDITION_FIXED_LEN {
            return Err(invalid("SortCond12 has trailing bytes with zero cchSt"));
        }
        return Ok(None);
    }
    let flags = *body
        .get(SORT_CONDITION_FIXED_LEN)
        .ok_or(XlsError::InvalidLength {
            expected: SORT_CONDITION_FIXED_LEN + 1,
            found: body.len(),
        })?;
    if flags & !0x01 != 0 {
        return Err(invalid("XLUnicodeStringNoCch has unsupported flag bits"));
    }
    let wide = flags & 0x01 != 0;
    let encoded_len = char_count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| invalid("SortCond12 string byte length overflow"))?;
    let expected = SORT_CONDITION_FIXED_LEN + 1 + encoded_len;
    if body.len() != expected {
        return Err(XlsError::InvalidLength {
            expected,
            found: body.len(),
        });
    }
    let encoded = &body[SORT_CONDITION_FIXED_LEN + 1..];
    let value = if wide {
        let units = encoded
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| XlsError::Encoding("invalid UTF-16 in SortCond12 custom list".into()))?
    } else {
        encoded.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn split_records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
        let mut records = Vec::new();
        let mut offset = 0;
        while offset < bytes.len() {
            let record_type = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
            let len = u16::from_le_bytes([bytes[offset + 2], bytes[offset + 3]]) as usize;
            offset += 4;
            records.push((record_type, bytes[offset..offset + len].to_vec()));
            offset += len;
        }
        records
    }

    #[test]
    fn round_trips_every_sort_on_kind_and_unicode() {
        let range = XlsSortRange::new(2, 1_048_575, 1, 16_383).unwrap();
        let mut value = XlsSortData::new(range, XlsSortParent::Table { id: 41 });
        value.set_orientation(XlsSortOrientation::Columns);
        value.set_case_sensitive(true);
        value.set_method(XlsSortMethod::Alternate);
        value.add_condition(XlsSortCondition::new(
            XlsSortRange::new(2, 20, 1, 1).unwrap(),
            true,
            XlsSortOn::Values {
                custom_list: Some("High,中,Low".into()),
            },
        ));
        value.add_condition(XlsSortCondition::new(
            XlsSortRange::new(2, 20, 2, 2).unwrap(),
            false,
            XlsSortOn::CellColor {
                differential_format: XlsDifferentialFormatIndex::new(7),
            },
        ));
        value.add_condition(XlsSortCondition::new(
            XlsSortRange::new(2, 20, 3, 3).unwrap(),
            true,
            XlsSortOn::FontColor {
                differential_format: XlsDifferentialFormatIndex::new(11),
            },
        ));
        value.add_condition(XlsSortCondition::new(
            XlsSortRange::new(2, 20, 4, 4).unwrap(),
            false,
            XlsSortOn::Icon {
                set: XlsSortIconSet::FiveQuarters,
                icon: XlsSortIcon::Fifth,
            },
        ));

        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        assert_eq!(records.len(), 5);
        assert_eq!(records[0].0, SORT_DATA_RECORD_TYPE);
        assert!(
            records[1..]
                .iter()
                .all(|record| record.0 == CONTINUE_FRT12_RECORD_TYPE)
        );
        let continuations = records[1..]
            .iter()
            .map(|record| record.1.as_slice())
            .collect::<Vec<_>>();
        let parsed = parse_sort_data(&records[0].1, &continuations).unwrap();
        assert_eq!(parsed, value);
    }

    #[test]
    fn rejects_invalid_ranges_and_condition_count_mismatch() {
        assert!(XlsSortRange::new(8, 7, 0, 0).is_err());
        assert!(XlsSortRange::new(0, MAX_ROW_INDEX + 1, 0, 0).is_err());
        assert!(XlsSortRange::new(0, 0, 0, MAX_COLUMN_INDEX + 1).is_err());

        let value = XlsSortData::new(XlsSortRange::new(0, 0, 0, 0).unwrap(), XlsSortParent::Sheet);
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let mut records = split_records(&bytes);
        records[0].1[30..34].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[]).is_err());
    }

    #[test]
    fn rejects_malformed_continuation_fields_and_unicode() {
        let range = XlsSortRange::new(0, 5, 0, 0).unwrap();
        let mut value = XlsSortData::new(range, XlsSortParent::AutoFilter);
        value.add_condition(XlsSortCondition::new(
            range,
            false,
            XlsSortOn::Values {
                custom_list: Some("中".into()),
            },
        ));
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        let mut malformed = records[1].1.clone();
        malformed[FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN] = 1;
        malformed[FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN + 1..]
            .copy_from_slice(&0xd800u16.to_le_bytes());
        malformed[FRT_HEADER_LEN + 26..FRT_HEADER_LEN + 30].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[&malformed]).is_err());

        let mut bad_icon = records[1].1.clone();
        bad_icon[FRT_HEADER_LEN..FRT_HEADER_LEN + 2].copy_from_slice(&6u16.to_le_bytes());
        bad_icon[FRT_HEADER_LEN + 18..FRT_HEADER_LEN + 22].copy_from_slice(&99u32.to_le_bytes());
        bad_icon.truncate(FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN);
        assert!(parse_sort_data(&records[0].1, &[&bad_icon]).is_err());

        let mut reserved_bit = records[1].1.clone();
        let flags = read_u16(&reserved_bit, FRT_HEADER_LEN).unwrap() | 0x8000;
        reserved_bit[FRT_HEADER_LEN..FRT_HEADER_LEN + 2].copy_from_slice(&flags.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[&reserved_bit]).is_err());
    }

    #[test]
    fn rejects_condition_larger_than_one_continue_frt12() {
        let range = XlsSortRange::new(0, 0, 0, 0).unwrap();
        let mut value = XlsSortData::new(range, XlsSortParent::Sheet);
        value.add_condition(XlsSortCondition::new(
            range,
            false,
            XlsSortOn::Values {
                custom_list: Some("a".repeat(MAX_CONTINUE_RGB_LEN)),
            },
        ));
        assert!(value.write_biff_records(&mut Vec::new()).is_err());
    }
}
