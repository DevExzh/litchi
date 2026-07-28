//! BIFF8 worksheet tables (`FeatHdr11`, `Feature11`, and `List12`).

use super::autofilter12::{
    XlsTableAutoFilter12, parse_table_autofilter12, write_table_autofilter12,
};
use super::{XlsError, XlsResult};
use std::collections::HashSet;

pub(crate) const FEAT_HDR11_RECORD_TYPE: u16 = 0x0871;
pub(crate) const FEATURE11_RECORD_TYPE: u16 = 0x0872;
pub(crate) const CONTINUE_FRT11_RECORD_TYPE: u16 = 0x0875;
pub(crate) const LIST12_RECORD_TYPE: u16 = 0x0877;
pub(crate) const FEATURE12_RECORD_TYPE: u16 = 0x0878;
pub(crate) const AUTO_FILTER12_RECORD_TYPE: u16 = 0x087e;
const ISF_LIST: u16 = 5;
const MAX_PAYLOAD: usize = 8_224;
const MAX_CONTINUE_RGB: usize = 8_212;
const MAX_FEATURE_BYTES: usize = 1_048_576;

fn invalid(rt: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: rt,
        message: message.into(),
    }
}
fn u16_at(data: &[u8], offset: usize, rt: u16, field: &str) -> XlsResult<u16> {
    data.get(offset..offset + 2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
fn u32_at(data: &[u8], offset: usize, rt: u16, field: &str) -> XlsResult<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
fn append_range(out: &mut Vec<u8>, range: XlsListObjectRange) {
    out.extend_from_slice(&range.first_row.to_le_bytes());
    out.extend_from_slice(&range.last_row.to_le_bytes());
    out.extend_from_slice(&range.first_column.to_le_bytes());
    out.extend_from_slice(&range.last_column.to_le_bytes());
}
fn parse_range(data: &[u8], offset: usize, rt: u16) -> XlsResult<XlsListObjectRange> {
    XlsListObjectRange::try_new(
        u16_at(data, offset, rt, "rwFirst")?,
        u16_at(data, offset + 2, rt, "rwLast")?,
        u16_at(data, offset + 4, rt, "colFirst")?,
        u16_at(data, offset + 6, rt, "colLast")?,
    )
}
fn append_frt(out: &mut Vec<u8>, rt: u16, range: Option<XlsListObjectRange>) {
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&u16::from(range.is_some()).to_le_bytes());
    if let Some(range) = range {
        append_range(out, range);
    } else {
        out.extend_from_slice(&[0; 8]);
    }
}
fn validate_frt(data: &[u8], rt: u16, reference: bool) -> XlsResult<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt
        || u16_at(data, 2, rt, "frt.flags")? != u16::from(reference)
    {
        return Err(invalid(rt, "future-record header is invalid"));
    }
    if !reference && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(rt, "future-record reserved bytes must be zero"));
    }
    Ok(())
}
fn validate_frt_any(data: &[u8], rt: u16) -> XlsResult<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt {
        return Err(invalid(rt, "future-record type echo is invalid"));
    }
    let flags = u16_at(data, 2, rt, "frt.flags")?;
    if flags & 0x0002 != 0 {
        return Err(invalid(rt, "future-record alert flag must be zero"));
    }
    if flags & 0x0001 == 0 && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(
            rt,
            "future-record reference is present without fFrtRef",
        ));
    }
    Ok(())
}
fn record(rt: u16, payload: Vec<u8>) -> XlsResult<Vec<u8>> {
    let len =
        u16::try_from(payload.len()).map_err(|_| invalid(rt, "payload exceeds BIFF8 length"))?;
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&payload);
    Ok(out)
}
fn append_string(out: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if value.is_ascii() {
        out.push(0);
        out.extend_from_slice(value.as_bytes());
    } else {
        out.push(1);
        out.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
}
fn parse_string(data: &[u8], offset: usize, rt: u16, field: &str) -> XlsResult<(String, usize)> {
    let count = usize::from(u16_at(data, offset, rt, field)?);
    let flags = *data
        .get(offset + 2)
        .ok_or_else(|| invalid(rt, format!("truncated {field} flags")))?;
    if flags & !1 != 0 {
        return Err(invalid(rt, format!("{field} flags are unsupported")));
    }
    let width = if flags == 0 { 1 } else { 2 };
    let end = offset
        .checked_add(3)
        .and_then(|v| count.checked_mul(width).and_then(|n| v.checked_add(n)))
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let bytes = data
        .get(offset + 3..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?;
    let value = if width == 1 {
        bytes.iter().map(|b| char::from(*b)).collect()
    } else {
        char::decode_utf16(
            bytes
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]])),
        )
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(rt, format!("invalid UTF-16 in {field}")))?
    };
    Ok((value, end))
}
fn validate_name(value: &str, field: &str) -> XlsResult<()> {
    if !(1..=255).contains(&value.encode_utf16().count())
        || value
            .chars()
            .any(|c| c <= '\u{1f}' || matches!(c, '\u{fffe}' | '\u{ffff}'))
    {
        return Err(invalid(FEATURE11_RECORD_TYPE, format!("invalid {field}")));
    }
    Ok(())
}
fn validate_table_name(value: &str) -> XlsResult<()> {
    validate_name(value, "table name")?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first.is_alphabetic() || matches!(first, '_' | '\\'))
        || chars.any(|c| !(c.is_alphanumeric() || matches!(c, '_' | '.' | '\\')))
    {
        return Err(invalid(
            FEATURE11_RECORD_TYPE,
            "table name must use Excel identifier syntax",
        ));
    }
    Ok(())
}
fn validate_column_name(value: &str) -> XlsResult<()> {
    if !(1..=255).contains(&value.encode_utf16().count())
        || value.chars().any(|c| {
            (c < '\u{20}' && !matches!(c, '\t' | '\n' | '\r'))
                || matches!(c, '\u{fffe}' | '\u{ffff}')
        })
    {
        return Err(invalid(FEATURE11_RECORD_TYPE, "invalid column name"));
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum FormulaExtraKind {
    Array,
    Memory,
}

fn parse_list_formula_extra_end(
    data: &[u8],
    tokens: &[u8],
    mut offset: usize,
    rt: u16,
) -> XlsResult<usize> {
    let mut extras = Vec::new();
    let mut position = 0usize;
    while position < tokens.len() {
        let opcode = tokens[position];
        let base = if opcode < 0x20 {
            opcode
        } else {
            (opcode & 0x1f) | 0x20
        };
        let size = match base {
            0x03..=0x16 => 1,
            0x17 => {
                let count = usize::from(
                    *tokens
                        .get(position + 1)
                        .ok_or_else(|| invalid(rt, "truncated formula string token"))?,
                );
                let flags = *tokens
                    .get(position + 2)
                    .ok_or_else(|| invalid(rt, "truncated formula string flags"))?;
                if flags & !1 != 0 {
                    return Err(invalid(rt, "unsupported formula string flags"));
                }
                3usize
                    .checked_add(
                        count
                            .checked_mul(if flags == 0 { 1 } else { 2 })
                            .ok_or_else(|| invalid(rt, "formula string length overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "formula string length overflows"))?
            },
            0x19 => {
                let header = tokens
                    .get(position..position + 4)
                    .ok_or_else(|| invalid(rt, "truncated Attr token"))?;
                if header[1] & 0x04 != 0 {
                    4usize
                        .checked_add(
                            (usize::from(u16::from_le_bytes([header[2], header[3]])) + 1) * 2,
                        )
                        .ok_or_else(|| invalid(rt, "Attr token length overflows"))?
                } else {
                    4
                }
            },
            0x1c | 0x1d => 2,
            0x1e => 3,
            0x1f => 9,
            0x20 => {
                extras.push(FormulaExtraKind::Array);
                8
            },
            0x21 => 3,
            0x22 => 4,
            0x23 | 0x24 | 0x2a | 0x2c => 5,
            0x25 | 0x2b | 0x2d => 9,
            0x26 => {
                extras.push(FormulaExtraKind::Memory);
                7
            },
            0x27 => 7,
            0x29 => 3,
            0x39 | 0x3a | 0x3c => 7,
            0x3b | 0x3d => 11,
            _ => {
                return Err(invalid(
                    rt,
                    "invalid or forbidden token in list array formula",
                ));
            },
        };
        position = position
            .checked_add(size)
            .ok_or_else(|| invalid(rt, "formula token length overflows"))?;
        if position > tokens.len() {
            return Err(invalid(rt, "truncated formula token"));
        }
    }
    for extra in extras {
        match extra {
            FormulaExtraKind::Memory => {
                let count = usize::from(u16_at(data, offset, rt, "PtgExtraMem count")?);
                offset = offset
                    .checked_add(2)
                    .and_then(|value| value.checked_add(count.checked_mul(8)?))
                    .ok_or_else(|| invalid(rt, "PtgExtraMem length overflows"))?;
                data.get(..offset)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraMem"))?;
            },
            FormulaExtraKind::Array => {
                let dimensions = data
                    .get(offset..offset + 3)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraArray dimensions"))?;
                let count = (usize::from(dimensions[0]) + 1)
                    .checked_mul(
                        usize::from(u16::from_le_bytes([dimensions[1], dimensions[2]])) + 1,
                    )
                    .ok_or_else(|| invalid(rt, "PtgExtraArray dimensions overflow"))?;
                offset += 3;
                for _ in 0..count {
                    let kind = *data
                        .get(offset)
                        .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                    offset += 1;
                    match kind {
                        0 | 1 | 4 | 16 => {
                            offset = offset
                                .checked_add(8)
                                .ok_or_else(|| invalid(rt, "PtgExtraArray length overflows"))?;
                            data.get(..offset)
                                .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                        },
                        2 => {
                            offset = parse_string(data, offset, rt, "PtgExtraArray string")?.1;
                        },
                        _ => return Err(invalid(rt, "invalid PtgExtraArray value type")),
                    }
                }
            },
        }
    }
    Ok(offset)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct XlsListObjectId(u32);
impl XlsListObjectId {
    pub fn try_new(value: u32) -> XlsResult<Self> {
        if value == 0 {
            Err(invalid(FEATURE11_RECORD_TYPE, "table id must be nonzero"))
        } else {
            Ok(Self(value))
        }
    }
    pub const fn value(self) -> u32 {
        self.0
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct XlsListColumnId(u32);
impl XlsListColumnId {
    pub fn try_new(value: u32) -> XlsResult<Self> {
        if value == 0 {
            Err(invalid(FEATURE11_RECORD_TYPE, "column id must be nonzero"))
        } else {
            Ok(Self(value))
        }
    }
    pub const fn value(self) -> u32 {
        self.0
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsListObjectRange {
    first_row: u16,
    last_row: u16,
    first_column: u16,
    last_column: u16,
}
impl XlsListObjectRange {
    pub fn try_new(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> XlsResult<Self> {
        if first_row > last_row || first_column > last_column || last_column > 255 {
            Err(invalid(
                FEATURE11_RECORD_TYPE,
                "table range is reversed or outside BIFF8 columns",
            ))
        } else {
            Ok(Self {
                first_row,
                last_row,
                first_column,
                last_column,
            })
        }
    }
    pub const fn first_row(self) -> u16 {
        self.first_row
    }
    pub const fn last_row(self) -> u16 {
        self.last_row
    }
    pub const fn first_column(self) -> u16 {
        self.first_column
    }
    pub const fn last_column(self) -> u16 {
        self.last_column
    }
    pub const fn column_count(self) -> usize {
        (self.last_column - self.first_column + 1) as usize
    }
    pub const fn overlaps(self, other: Self) -> bool {
        self.first_row <= other.last_row
            && other.first_row <= self.last_row
            && self.first_column <= other.last_column
            && other.first_column <= self.last_column
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsListTotalAggregation {
    None,
    Average,
    Count,
    CountNumbers,
    Max,
    Min,
    Sum,
    StandardDeviation,
    Variance,
    Custom,
}
impl XlsListTotalAggregation {
    fn code(self) -> u32 {
        self as u32
    }
    fn from_code(v: u32) -> XlsResult<Self> {
        Ok(match v {
            0 => Self::None,
            1 => Self::Average,
            2 => Self::Count,
            3 => Self::CountNumbers,
            4 => Self::Max,
            5 => Self::Min,
            6 => Self::Sum,
            7 => Self::StandardDeviation,
            8 => Self::Variance,
            9 => Self::Custom,
            _ => return Err(invalid(FEATURE11_RECORD_TYPE, "invalid total aggregation")),
        })
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsListObjectColumn {
    id: XlsListColumnId,
    name: String,
    aggregation: XlsListTotalAggregation,
    total_formula: Option<Vec<u8>>,
    total_string: Option<String>,
}
impl XlsListObjectColumn {
    pub fn try_new(id: XlsListColumnId, name: impl Into<String>) -> XlsResult<Self> {
        let value = Self {
            id,
            name: name.into(),
            aggregation: XlsListTotalAggregation::None,
            total_formula: None,
            total_string: None,
        };
        validate_column_name(&value.name)?;
        Ok(value)
    }
    pub fn with_total_aggregation(mut self, value: XlsListTotalAggregation) -> XlsResult<Self> {
        self.aggregation = value;
        self.validate_totals()?;
        Ok(self)
    }
    pub fn with_total_formula_tokens(mut self, tokens: Vec<u8>) -> XlsResult<Self> {
        if tokens.is_empty() || tokens.len() > u16::MAX as usize {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total formula token length must be 1..=65535",
            ));
        }
        self.aggregation = XlsListTotalAggregation::Custom;
        self.total_formula = Some(tokens);
        self.validate_totals()?;
        Ok(self)
    }
    pub fn with_total_string(mut self, value: impl Into<String>) -> XlsResult<Self> {
        let value = value.into();
        if value.encode_utf16().count() > 32767 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string exceeds 32767 UTF-16 units",
            ));
        }
        self.aggregation = XlsListTotalAggregation::None;
        self.total_string = Some(value);
        self.validate_totals()?;
        Ok(self)
    }
    fn validate_totals(&self) -> XlsResult<()> {
        if self.total_formula.is_some() != (self.aggregation == XlsListTotalAggregation::Custom) {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "custom aggregation and total formula must occur together",
            ));
        }
        if self.total_string.is_some() && self.aggregation != XlsListTotalAggregation::None {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string requires no aggregation",
            ));
        }
        Ok(())
    }
    pub const fn id(&self) -> XlsListColumnId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn total_aggregation(&self) -> XlsListTotalAggregation {
        self.aggregation
    }
    pub fn total_formula_tokens(&self) -> Option<&[u8]> {
        self.total_formula.as_deref()
    }
    pub fn total_string(&self) -> Option<&str> {
        self.total_string.as_deref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsListObjectStyleOptions {
    name: String,
    first: bool,
    last: bool,
    row_stripes: bool,
    column_stripes: bool,
    default_style: bool,
}
impl XlsListObjectStyleOptions {
    pub fn try_new(name: impl Into<String>) -> XlsResult<Self> {
        let name = name.into();
        validate_name(&name, "table style name")?;
        Ok(Self {
            name,
            first: false,
            last: false,
            row_stripes: true,
            column_stripes: false,
            default_style: false,
        })
    }
    pub fn with_first_column(mut self, v: bool) -> Self {
        self.first = v;
        self
    }
    pub fn with_last_column(mut self, v: bool) -> Self {
        self.last = v;
        self
    }
    pub fn with_row_stripes(mut self, v: bool) -> Self {
        self.row_stripes = v;
        self
    }
    pub fn with_column_stripes(mut self, v: bool) -> Self {
        self.column_stripes = v;
        self
    }
    pub fn with_default_style(mut self, v: bool) -> Self {
        self.default_style = v;
        self
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn shows_first_column(&self) -> bool {
        self.first
    }
    pub const fn shows_last_column(&self) -> bool {
        self.last
    }
    pub const fn shows_row_stripes(&self) -> bool {
        self.row_stripes
    }
    pub const fn shows_column_stripes(&self) -> bool {
        self.column_stripes
    }
    pub const fn is_default_style(&self) -> bool {
        self.default_style
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsListObjectFeatureVersion {
    Feature11,
    Feature12,
}

/// Excel version recorded by an external-data table definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsExternalTableVersion {
    Excel2003,
    Excel2007,
}
impl XlsExternalTableVersion {
    const fn code(self) -> u32 {
        match self {
            Self::Excel2003 => 0xB,
            Self::Excel2007 => 0xC,
        }
    }
    fn from_code(value: u32) -> XlsResult<Self> {
        match value {
            0xB => Ok(Self::Excel2003),
            0xC => Ok(Self::Excel2007),
            _ => Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external table verXL must be 0xB or 0xC",
            )),
        }
    }
}

/// Inert formatting metadata for a headerless external-table column.
///
/// The DXFN12List payload is preserved without interpretation. Parsed values
/// retain the original XLUnicodeString encoding of the optional style name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsCachedDiskHeader {
    encoded: Vec<u8>,
    format_end: usize,
    style_name: Option<String>,
}

impl XlsCachedDiskHeader {
    /// Construct a cached header from an inert serialized DXFN12List payload.
    pub fn try_new(formatting: Vec<u8>) -> XlsResult<Self> {
        if formatting.len() > MAX_FEATURE_BYTES.saturating_sub(4)
            || formatting.len() > u32::MAX as usize
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "cached header formatting exceeds resource bound",
            ));
        }
        let mut encoded = Vec::with_capacity(4 + formatting.len());
        encoded.extend_from_slice(&(formatting.len() as u32).to_le_bytes());
        encoded.extend_from_slice(&formatting);
        Ok(Self {
            format_end: encoded.len(),
            encoded,
            style_name: None,
        })
    }

    fn empty() -> Self {
        Self::try_new(Vec::new()).expect("empty cached header is valid")
    }

    fn parse(encoded: Vec<u8>, has_style_name: bool, rt: u16) -> XlsResult<Self> {
        if encoded.len() > MAX_FEATURE_BYTES {
            return Err(invalid(rt, "cached header exceeds resource bound"));
        }
        let format_len = usize::try_from(u32_at(&encoded, 0, rt, "cbdxfHdrDisk")?)
            .map_err(|_| invalid(rt, "cached header format length overflows"))?;
        let format_end = 4usize
            .checked_add(format_len)
            .ok_or_else(|| invalid(rt, "cached header format length overflows"))?;
        encoded
            .get(4..format_end)
            .ok_or_else(|| invalid(rt, "truncated cached header formatting"))?;
        let style_name = if has_style_name {
            let (name, end) = parse_string(&encoded, format_end, rt, "cached header style")?;
            if end != encoded.len() {
                return Err(invalid(rt, "trailing cached header data"));
            }
            validate_name(&name, "cached header style name")?;
            Some(name)
        } else {
            if format_end != encoded.len() {
                return Err(invalid(
                    rt,
                    "cached header style data exists without fSaveStyleName",
                ));
            }
            None
        };
        Ok(Self {
            encoded,
            format_end,
            style_name,
        })
    }

    pub fn with_style_name(mut self, name: impl Into<String>) -> XlsResult<Self> {
        let name = name.into();
        validate_name(&name, "cached header style name")?;
        self.encoded.truncate(self.format_end);
        append_string(&mut self.encoded, &name);
        if self.encoded.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "cached header exceeds resource bound",
            ));
        }
        self.style_name = Some(name);
        Ok(self)
    }

    pub fn without_style_name(mut self) -> Self {
        self.encoded.truncate(self.format_end);
        self.style_name = None;
        self
    }

    /// Complete CachedDiskHeader bytes, including the format-length prefix.
    pub fn as_bytes(&self) -> &[u8] {
        &self.encoded
    }

    /// Inert DXFN12List bytes without the CachedDiskHeader length prefix.
    pub fn formatting_bytes(&self) -> &[u8] {
        &self.encoded[4..self.format_end]
    }

    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }
}

/// Inert metadata that associates one table column with a query-table field.
///
/// Opaque byte slices are retained for BIFF substructures that litchi does not
/// execute or render. `auto_filter` contains the complete Feat11FdaAutoFilter.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsExternalTableField {
    column_id: XlsListColumnId,
    source_name: String,
    query_field_id: u32,
    aggregate_format: Vec<u8>,
    insert_row_format: Vec<u8>,
    auto_filter: Vec<u8>,
    formula_extra: Vec<u8>,
    header_cache: XlsCachedDiskHeader,
    aggregate_style: u32,
    insert_row_style: u32,
    filter_hidden: bool,
    total_array_formula: bool,
    auto_create_calculated_column: bool,
}
impl XlsExternalTableField {
    pub fn try_new(
        column_id: XlsListColumnId,
        source_name: impl Into<String>,
        query_field_id: u32,
    ) -> XlsResult<Self> {
        let value = Self {
            column_id,
            source_name: source_name.into(),
            query_field_id,
            aggregate_format: Vec::new(),
            insert_row_format: Vec::new(),
            auto_filter: vec![0; 6],
            formula_extra: Vec::new(),
            header_cache: XlsCachedDiskHeader::empty(),
            aggregate_style: u32::MAX,
            insert_row_style: u32::MAX,
            filter_hidden: false,
            total_array_formula: false,
            auto_create_calculated_column: false,
        };
        value.validate()?;
        Ok(value)
    }
    fn validate(&self) -> XlsResult<()> {
        validate_name(&self.source_name, "external source field name")?;
        if self.query_field_id == 0 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external query field id must be nonzero",
            ));
        }
        for (name, bytes) in [
            ("aggregate format", self.aggregate_format.as_slice()),
            ("insert-row format", self.insert_row_format.as_slice()),
            ("AutoFilter", self.auto_filter.as_slice()),
            ("formula extra data", self.formula_extra.as_slice()),
            ("header cache", self.header_cache.as_bytes()),
        ] {
            if bytes.len() > MAX_FEATURE_BYTES {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    format!("external {name} exceeds resource bound"),
                ));
            }
        }
        if self.auto_filter.len() < 6
            || usize::try_from(u32_at(
                &self.auto_filter,
                0,
                FEATURE12_RECORD_TYPE,
                "cbAutoFilter",
            )?)
            .ok()
                != Some(self.auto_filter.len() - 6)
            || self.auto_filter.len() - 6 > 2080
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external AutoFilter size is inconsistent",
            ));
        }
        Ok(())
    }
    pub const fn column_id(&self) -> XlsListColumnId {
        self.column_id
    }
    pub fn source_name(&self) -> &str {
        &self.source_name
    }
    pub const fn query_field_id(&self) -> u32 {
        self.query_field_id
    }
    pub fn aggregate_format_bytes(&self) -> &[u8] {
        &self.aggregate_format
    }
    pub fn insert_row_format_bytes(&self) -> &[u8] {
        &self.insert_row_format
    }
    pub fn auto_filter_bytes(&self) -> &[u8] {
        &self.auto_filter
    }
    pub fn formula_extra_bytes(&self) -> &[u8] {
        &self.formula_extra
    }
    pub fn header_cache_bytes(&self) -> &[u8] {
        self.header_cache.as_bytes()
    }
    pub const fn cached_disk_header(&self) -> &XlsCachedDiskHeader {
        &self.header_cache
    }
    pub const fn aggregate_style_index(&self) -> u32 {
        self.aggregate_style
    }
    pub const fn insert_row_style_index(&self) -> u32 {
        self.insert_row_style
    }
    pub const fn is_filter_hidden(&self) -> bool {
        self.filter_hidden
    }
    pub const fn is_total_array_formula(&self) -> bool {
        self.total_array_formula
    }
    pub const fn auto_creates_calculated_column(&self) -> bool {
        self.auto_create_calculated_column
    }
    pub fn with_aggregate_format_bytes(mut self, bytes: Vec<u8>) -> XlsResult<Self> {
        self.aggregate_format = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_insert_row_format_bytes(mut self, bytes: Vec<u8>) -> XlsResult<Self> {
        self.insert_row_format = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_auto_filter_bytes(mut self, bytes: Vec<u8>) -> XlsResult<Self> {
        self.auto_filter = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_formula_extra_bytes(mut self, bytes: Vec<u8>, array: bool) -> XlsResult<Self> {
        self.formula_extra = bytes;
        self.total_array_formula = array;
        self.validate()?;
        Ok(self)
    }
    pub fn with_header_cache_bytes(mut self, bytes: Vec<u8>) -> XlsResult<Self> {
        let format_len = usize::try_from(u32_at(&bytes, 0, FEATURE12_RECORD_TYPE, "cbdxfHdrDisk")?)
            .map_err(|_| invalid(FEATURE12_RECORD_TYPE, "cached header length overflows"))?;
        let format_end = 4usize
            .checked_add(format_len)
            .ok_or_else(|| invalid(FEATURE12_RECORD_TYPE, "cached header length overflows"))?;
        let has_style_name = format_end < bytes.len();
        self.header_cache =
            XlsCachedDiskHeader::parse(bytes, has_style_name, FEATURE12_RECORD_TYPE)?;
        self.validate()?;
        Ok(self)
    }
    pub fn with_cached_disk_header(mut self, header: XlsCachedDiskHeader) -> XlsResult<Self> {
        self.header_cache = header;
        self.validate()?;
        Ok(self)
    }
}

/// Typed, non-executing metadata for a Feature12 LTEXTERNALDATA table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsExternalTableMetadata {
    version: XlsExternalTableVersion,
    build_number: u16,
    fields: Vec<XlsExternalTableField>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsWebColumnType {
    Text,
    Number,
    Boolean,
    DateTime,
    Note,
    Currency,
    Lookup,
    Choice,
    Url,
    Counter,
    MultipleChoices,
}
impl XlsWebColumnType {
    pub const ALL: &'static [Self] = &[
        Self::Text,
        Self::Number,
        Self::Boolean,
        Self::DateTime,
        Self::Note,
        Self::Currency,
        Self::Lookup,
        Self::Choice,
        Self::Url,
        Self::Counter,
        Self::MultipleChoices,
    ];
    pub const fn value(self) -> u32 {
        self as u32 + 1
    }
    fn code(self) -> u32 {
        self.value()
    }
    fn from_code(value: u32) -> XlsResult<Self> {
        Ok(match value {
            1 => Self::Text,
            2 => Self::Number,
            3 => Self::Boolean,
            4 => Self::DateTime,
            5 => Self::Note,
            6 => Self::Currency,
            7 => Self::Lookup,
            8 => Self::Choice,
            9 => Self::Url,
            10 => Self::Counter,
            11 => Self::MultipleChoices,
            _ => return Err(invalid(FEATURE11_RECORD_TYPE, "invalid Web column type")),
        })
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsWebReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}
impl XlsWebReadingOrder {
    fn code(self) -> u32 {
        self as u32
    }
    fn from_code(v: u32) -> XlsResult<Self> {
        match v {
            0 => Ok(Self::Context),
            1 => Ok(Self::LeftToRight),
            2 => Ok(Self::RightToLeft),
            _ => Err(invalid(FEATURE11_RECORD_TYPE, "invalid Web reading order")),
        }
    }
}
#[derive(Debug, Clone, PartialEq)]
pub enum XlsWebDefaultValue {
    String(String),
    Boolean(bool),
    Number(f64),
    DateTime(f64),
}
impl Eq for XlsWebDefaultValue {}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWebFieldInfo {
    locale: u32,
    decimal_places: u32,
    percent: bool,
    fixed_decimal: bool,
    date_only: bool,
    reading_order: XlsWebReadingOrder,
    rich_text: bool,
    unknown_rich_text: bool,
    alert_unknown_rich_text: bool,
    read_only: bool,
    required: bool,
    minimum_set: bool,
    maximum_set: bool,
    default_today: bool,
    allow_fill_in: bool,
    default_value: Option<XlsWebDefaultValue>,
    validation_formula: Option<String>,
    ignored_display_flags: u32,
    ignored_validation_flags: u32,
}
impl XlsWebFieldInfo {
    pub fn new(locale: u32) -> Self {
        Self {
            locale,
            decimal_places: 0,
            percent: false,
            fixed_decimal: false,
            date_only: false,
            reading_order: XlsWebReadingOrder::Context,
            rich_text: false,
            unknown_rich_text: false,
            alert_unknown_rich_text: false,
            read_only: false,
            required: false,
            minimum_set: false,
            maximum_set: false,
            default_today: false,
            allow_fill_in: false,
            default_value: None,
            validation_formula: None,
            ignored_display_flags: 0,
            ignored_validation_flags: 0,
        }
    }
    pub fn with_decimal_display(mut self, places: u32, percent: bool) -> Self {
        self.decimal_places = places;
        self.fixed_decimal = true;
        self.percent = percent;
        self
    }
    pub fn with_default_value(mut self, value: XlsWebDefaultValue) -> Self {
        self.default_value = Some(value);
        self
    }
    pub fn with_validation_formula(mut self, value: impl Into<String>) -> XlsResult<Self> {
        let value = value.into();
        if value.encode_utf16().count() > 255 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "Web validation formula exceeds 255 characters",
            ));
        }
        self.validation_formula = Some(value);
        Ok(self)
    }
    pub fn with_read_only(mut self, value: bool) -> Self {
        self.read_only = value;
        self
    }
    pub fn with_required(mut self, value: bool) -> Self {
        self.required = value;
        self
    }
    pub const fn locale(&self) -> u32 {
        self.locale
    }
    pub const fn decimal_places(&self) -> u32 {
        self.decimal_places
    }
    pub fn default_value(&self) -> Option<&XlsWebDefaultValue> {
        self.default_value.as_ref()
    }
    pub fn validation_formula(&self) -> Option<&str> {
        self.validation_formula.as_deref()
    }
    /// Undefined display-flag bits retained from a parsed WSS field.
    pub const fn ignored_display_flags(&self) -> u32 {
        self.ignored_display_flags
    }
    /// Undefined validation-flag bits retained from a parsed WSS field.
    pub const fn ignored_validation_flags(&self) -> u32 {
        self.ignored_validation_flags
    }
    fn validate(&self, kind: XlsWebColumnType) -> XlsResult<()> {
        if self.reading_order.code() > 2 {
            return Err(invalid(FEATURE11_RECORD_TYPE, "invalid Web reading order"));
        }
        if let Some(value) = &self.default_value {
            let valid = matches!(
                (kind, value),
                (
                    XlsWebColumnType::Text
                        | XlsWebColumnType::Choice
                        | XlsWebColumnType::MultipleChoices,
                    XlsWebDefaultValue::String(_)
                ) | (XlsWebColumnType::Boolean, XlsWebDefaultValue::Boolean(_))
                    | (
                        XlsWebColumnType::Number | XlsWebColumnType::Currency,
                        XlsWebDefaultValue::Number(_)
                    )
                    | (XlsWebColumnType::DateTime, XlsWebDefaultValue::DateTime(_))
            );
            if !valid {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Web default value does not match column type",
                ));
            }
            if let XlsWebDefaultValue::String(value) = value {
                if value.encode_utf16().count() > 255 {
                    return Err(invalid(
                        FEATURE11_RECORD_TYPE,
                        "Web default string exceeds 255 characters",
                    ));
                }
            }
        }
        Ok(())
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWebTableField {
    column_id: XlsListColumnId,
    source_name: String,
    data_type: XlsWebColumnType,
    info: XlsWebFieldInfo,
    calculated_formula: Option<Vec<u8>>,
    auto_filter: Vec<u8>,
    aggregate_format: Vec<u8>,
    insert_row_format: Vec<u8>,
    total_formula_extra: Vec<u8>,
    header_cache: Vec<u8>,
    ignored_flags: u32,
}
impl XlsWebTableField {
    pub fn try_new(
        column_id: XlsListColumnId,
        source_name: impl Into<String>,
        data_type: XlsWebColumnType,
        info: XlsWebFieldInfo,
    ) -> XlsResult<Self> {
        let value = Self {
            column_id,
            source_name: source_name.into(),
            data_type,
            info,
            calculated_formula: None,
            auto_filter: vec![0; 6],
            aggregate_format: Vec::new(),
            insert_row_format: Vec::new(),
            total_formula_extra: Vec::new(),
            header_cache: vec![0; 4],
            ignored_flags: 0,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn with_calculated_formula_tokens(mut self, tokens: Vec<u8>) -> XlsResult<Self> {
        if tokens.is_empty() || tokens.len() > u16::MAX as usize {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "Web calculated formula token length must be 1..=65535",
            ));
        }
        self.calculated_formula = Some(tokens);
        Ok(self)
    }
    pub const fn column_id(&self) -> XlsListColumnId {
        self.column_id
    }
    pub fn source_name(&self) -> &str {
        &self.source_name
    }
    pub const fn data_type(&self) -> XlsWebColumnType {
        self.data_type
    }
    pub const fn info(&self) -> &XlsWebFieldInfo {
        &self.info
    }
    pub fn calculated_formula_tokens(&self) -> Option<&[u8]> {
        self.calculated_formula.as_deref()
    }
    /// Undefined Feat11FieldDataItem flag bits retained from parsed input.
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
    fn validate(&self) -> XlsResult<()> {
        validate_name(&self.source_name, "Web source field name")?;
        self.info.validate(self.data_type)?;
        if self.auto_filter.len() < 6
            || usize::try_from(u32_at(
                &self.auto_filter,
                0,
                FEATURE11_RECORD_TYPE,
                "cbAutoFilter",
            )?)
            .ok()
                != Some(self.auto_filter.len() - 6)
            || self.auto_filter.len() - 6 > 2080
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid Web field AutoFilter",
            ));
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsWebEditMode {
    Normal,
    RefreshCopy,
    RefreshCache,
    RefreshCacheUndo,
    RefreshLoaded,
    RefreshTemplate,
    RefreshRefresh,
    NoInsertRequired,
    NoInsertDocumentLibrary,
    RefreshLoadDiscarded,
    RefreshLoadHashValidation,
    NoEditModeratedView,
}
impl XlsWebEditMode {
    fn code(self) -> u32 {
        self as u32
    }
    fn from_code(v: u32) -> XlsResult<Self> {
        Ok(match v {
            0 => Self::Normal,
            1 => Self::RefreshCopy,
            2 => Self::RefreshCache,
            3 => Self::RefreshCacheUndo,
            4 => Self::RefreshLoaded,
            5 => Self::RefreshTemplate,
            6 => Self::RefreshRefresh,
            7 => Self::NoInsertRequired,
            8 => Self::NoInsertDocumentLibrary,
            9 => Self::RefreshLoadDiscarded,
            10 => Self::RefreshLoadHashValidation,
            11 => Self::NoEditModeratedView,
            _ => {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "invalid Web table edit mode",
                ));
            },
        })
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWebInvalidCell {
    row_id: u32,
    column_id: XlsListColumnId,
}
impl XlsWebInvalidCell {
    pub fn new(row_id: u32, column_id: XlsListColumnId) -> Self {
        Self { row_id, column_id }
    }
    pub const fn row_id(self) -> u32 {
        self.row_id
    }
    pub const fn column_id(self) -> XlsListColumnId {
        self.column_id
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWebTableMetadata {
    version: XlsExternalTableVersion,
    build_number: u16,
    fields: Vec<XlsWebTableField>,
    edit_mode: XlsWebEditMode,
    cache_position: u32,
    cache_size: u32,
    cache_characters: u32,
    hash_parameters: [u8; 16],
    provider_name: Option<String>,
    entry_id: Option<String>,
    deleted_row_ids: Vec<u32>,
    changed_row_ids: Vec<u32>,
    invalid_cells: Vec<XlsWebInvalidCell>,
    needs_commit: bool,
    compressed_cache: bool,
    ignored_fixed_word: u16,
    ignored_flags: u32,
}
impl XlsWebTableMetadata {
    pub fn try_new(fields: Vec<XlsWebTableField>) -> XlsResult<Self> {
        let value = Self {
            version: XlsExternalTableVersion::Excel2003,
            build_number: 0,
            fields,
            edit_mode: XlsWebEditMode::Normal,
            cache_position: 0,
            cache_size: 0,
            cache_characters: 0,
            hash_parameters: [0; 16],
            provider_name: None,
            entry_id: None,
            deleted_row_ids: Vec::new(),
            changed_row_ids: Vec::new(),
            invalid_cells: Vec::new(),
            needs_commit: false,
            compressed_cache: false,
            ignored_fixed_word: 0,
            ignored_flags: 0,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn fields(&self) -> &[XlsWebTableField] {
        &self.fields
    }
    pub const fn edit_mode(&self) -> XlsWebEditMode {
        self.edit_mode
    }
    pub fn deleted_row_ids(&self) -> &[u32] {
        &self.deleted_row_ids
    }
    pub fn changed_row_ids(&self) -> &[u32] {
        &self.changed_row_ids
    }
    pub fn invalid_cells(&self) -> &[XlsWebInvalidCell] {
        &self.invalid_cells
    }
    pub const fn ignored_fixed_word(&self) -> u16 {
        self.ignored_fixed_word
    }
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
    pub fn with_deleted_row_ids(mut self, v: Vec<u32>) -> XlsResult<Self> {
        self.deleted_row_ids = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_changed_row_ids(mut self, v: Vec<u32>) -> XlsResult<Self> {
        self.changed_row_ids = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_invalid_cells(mut self, v: Vec<XlsWebInvalidCell>) -> XlsResult<Self> {
        self.invalid_cells = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_provider_name(mut self, v: impl Into<String>) -> XlsResult<Self> {
        let v = v.into();
        validate_name(&v, "Web cryptographic provider")?;
        self.provider_name = Some(v);
        Ok(self)
    }
    pub fn with_entry_id(mut self, v: impl Into<String>) -> XlsResult<Self> {
        let v = v.into();
        validate_name(&v, "Web entry id")?;
        self.entry_id = Some(v);
        Ok(self)
    }
    fn validate(&self) -> XlsResult<()> {
        if !(1..=256).contains(&self.fields.len())
            || self.deleted_row_ids.len() > u16::MAX as usize
            || self.changed_row_ids.len() > u16::MAX as usize
            || self.invalid_cells.len() > u16::MAX as usize
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "Web table source count exceeds BIFF bounds",
            ));
        }
        let mut ids = HashSet::new();
        let mut names = HashSet::new();
        for field in &self.fields {
            field.validate()?;
            if !ids.insert(field.column_id) || !names.insert(field.source_name.to_lowercase()) {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "duplicate Web source field ownership",
                ));
            }
        }
        for cell in &self.invalid_cells {
            if !ids.contains(&cell.column_id) {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "invalid Web synchronization cell column",
                ));
            }
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum XlsXmlDataType {
    Null = 0x0800,
    Schema = 0x1000,
    Attribute = 0x1001,
    AttributeGroup = 0x1002,
    Notation = 0x1003,
    IdentityConstraint = 0x1100,
    Key = 0x1101,
    KeyRef = 0x1102,
    Unique = 0x1103,
    AnyType = 0x2000,
    DataType = 0x2100,
    DataTypeAnyType = 0x2101,
    DataTypeAnyUri = 0x2102,
    DataTypeBase64Binary = 0x2103,
    DataTypeBoolean = 0x2104,
    DataTypeByte = 0x2105,
    DataTypeDate = 0x2106,
    DataTypeDateTime = 0x2107,
    DataTypeDay = 0x2108,
    DataTypeDecimal = 0x2109,
    DataTypeDouble = 0x210a,
    DataTypeDuration = 0x210b,
    DataTypeEntities = 0x210c,
    DataTypeEntity = 0x210d,
    DataTypeFloat = 0x210e,
    DataTypeHexBinary = 0x210f,
    DataTypeId = 0x2110,
    DataTypeIdRef = 0x2111,
    DataTypeIdRefs = 0x2112,
    DataTypeInt = 0x2113,
    DataTypeInteger = 0x2114,
    DataTypeLanguage = 0x2115,
    DataTypeLong = 0x2116,
    DataTypeMonth = 0x2117,
    DataTypeMonthDay = 0x2118,
    DataTypeName = 0x2119,
    DataTypeNcName = 0x211a,
    DataTypeNegativeInteger = 0x211b,
    DataTypeNmToken = 0x211c,
    DataTypeNmTokens = 0x211d,
    DataTypeNonNegativeInteger = 0x211e,
    DataTypeNonPositiveInteger = 0x211f,
    DataTypeNormalizedString = 0x2120,
    DataTypeNotation = 0x2121,
    DataTypePositiveInteger = 0x2122,
    DataTypeQName = 0x2123,
    DataTypeShort = 0x2124,
    DataTypeString = 0x2125,
    DataTypeTime = 0x2126,
    DataTypeToken = 0x2127,
    DataTypeUnsignedByte = 0x2128,
    DataTypeUnsignedInt = 0x2129,
    DataTypeUnsignedLong = 0x212a,
    DataTypeUnsignedShort = 0x212b,
    DataTypeYear = 0x212c,
    DataTypeYearMonth = 0x212d,
    DataTypeAnySimpleType = 0x21ff,
    SimpleType = 0x2200,
    ComplexType = 0x2400,
    NullType = 0x2800,
    Particle = 0x4000,
    Any = 0x4001,
    AnyAttribute = 0x4002,
    Element = 0x4003,
    Group = 0x4100,
    All = 0x4101,
    Choice = 0x4102,
    Sequence = 0x4103,
    EmptyParticle = 0x4104,
    NullAny = 0x4801,
    NullAnyAttribute = 0x4802,
    NullElement = 0x4803,
}
impl XlsXmlDataType {
    pub const ALL: &'static [Self] = &[
        Self::Null,
        Self::Schema,
        Self::Attribute,
        Self::AttributeGroup,
        Self::Notation,
        Self::IdentityConstraint,
        Self::Key,
        Self::KeyRef,
        Self::Unique,
        Self::AnyType,
        Self::DataType,
        Self::DataTypeAnyType,
        Self::DataTypeAnyUri,
        Self::DataTypeBase64Binary,
        Self::DataTypeBoolean,
        Self::DataTypeByte,
        Self::DataTypeDate,
        Self::DataTypeDateTime,
        Self::DataTypeDay,
        Self::DataTypeDecimal,
        Self::DataTypeDouble,
        Self::DataTypeDuration,
        Self::DataTypeEntities,
        Self::DataTypeEntity,
        Self::DataTypeFloat,
        Self::DataTypeHexBinary,
        Self::DataTypeId,
        Self::DataTypeIdRef,
        Self::DataTypeIdRefs,
        Self::DataTypeInt,
        Self::DataTypeInteger,
        Self::DataTypeLanguage,
        Self::DataTypeLong,
        Self::DataTypeMonth,
        Self::DataTypeMonthDay,
        Self::DataTypeName,
        Self::DataTypeNcName,
        Self::DataTypeNegativeInteger,
        Self::DataTypeNmToken,
        Self::DataTypeNmTokens,
        Self::DataTypeNonNegativeInteger,
        Self::DataTypeNonPositiveInteger,
        Self::DataTypeNormalizedString,
        Self::DataTypeNotation,
        Self::DataTypePositiveInteger,
        Self::DataTypeQName,
        Self::DataTypeShort,
        Self::DataTypeString,
        Self::DataTypeTime,
        Self::DataTypeToken,
        Self::DataTypeUnsignedByte,
        Self::DataTypeUnsignedInt,
        Self::DataTypeUnsignedLong,
        Self::DataTypeUnsignedShort,
        Self::DataTypeYear,
        Self::DataTypeYearMonth,
        Self::DataTypeAnySimpleType,
        Self::SimpleType,
        Self::ComplexType,
        Self::NullType,
        Self::Particle,
        Self::Any,
        Self::AnyAttribute,
        Self::Element,
        Self::Group,
        Self::All,
        Self::Choice,
        Self::Sequence,
        Self::EmptyParticle,
        Self::NullAny,
        Self::NullAnyAttribute,
        Self::NullElement,
    ];
    pub fn try_new(v: u32) -> XlsResult<Self> {
        Self::ALL
            .iter()
            .copied()
            .find(|value| value.value() == v)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "invalid XML column data type"))
    }
    pub const fn value(self) -> u32 {
        self as u32
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsXmlColumnMapping {
    can_be_single: bool,
    map_id: u32,
    xpath: String,
}
impl XlsXmlColumnMapping {
    pub fn try_new(map_id: u32, xpath: impl Into<String>, can_be_single: bool) -> XlsResult<Self> {
        let xpath = xpath.into();
        if map_id == 0 || xpath.encode_utf16().count() >= 32000 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid XML map id or XPath",
            ));
        }
        Ok(Self {
            can_be_single,
            map_id,
            xpath,
        })
    }
    pub const fn map_id(&self) -> u32 {
        self.map_id
    }
    pub fn xpath(&self) -> &str {
        &self.xpath
    }
    pub const fn can_be_single(&self) -> bool {
        self.can_be_single
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsXmlTableField {
    column_id: XlsListColumnId,
    source_name: String,
    data_type: XlsXmlDataType,
    mapping: Option<XlsXmlColumnMapping>,
    auto_filter: Vec<u8>,
    aggregate_format: Vec<u8>,
    insert_row_format: Vec<u8>,
    total_formula_extra: Vec<u8>,
    header_cache: Vec<u8>,
    ignored_flags: u32,
}
impl XlsXmlTableField {
    pub fn try_new(
        column_id: XlsListColumnId,
        source_name: impl Into<String>,
        data_type: XlsXmlDataType,
    ) -> XlsResult<Self> {
        let value = Self {
            column_id,
            source_name: source_name.into(),
            data_type,
            mapping: None,
            auto_filter: vec![0; 6],
            aggregate_format: Vec::new(),
            insert_row_format: Vec::new(),
            total_formula_extra: Vec::new(),
            header_cache: vec![0; 4],
            ignored_flags: 0,
        };
        validate_name(&value.source_name, "XML source field name")?;
        Ok(value)
    }
    pub fn with_mapping(mut self, v: XlsXmlColumnMapping) -> Self {
        self.mapping = Some(v);
        self
    }
    pub const fn column_id(&self) -> XlsListColumnId {
        self.column_id
    }
    pub fn source_name(&self) -> &str {
        &self.source_name
    }
    pub const fn data_type(&self) -> XlsXmlDataType {
        self.data_type
    }
    pub fn mapping(&self) -> Option<&XlsXmlColumnMapping> {
        self.mapping.as_ref()
    }
    /// Undefined Feat11FieldDataItem flag bits retained from parsed input.
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsXmlTableMetadata {
    version: XlsExternalTableVersion,
    build_number: u16,
    fields: Vec<XlsXmlTableField>,
    entry_id: Option<String>,
    single_cell: bool,
    ignored_fixed_word: u16,
    ignored_flags: u32,
    ignored_fixed_tail: [u8; 32],
}
impl XlsXmlTableMetadata {
    pub fn try_new(fields: Vec<XlsXmlTableField>) -> XlsResult<Self> {
        if !(1..=256).contains(&fields.len()) {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "XML field count must be 1..=256",
            ));
        }
        let value = Self {
            version: XlsExternalTableVersion::Excel2003,
            build_number: 0,
            fields,
            entry_id: None,
            single_cell: false,
            ignored_fixed_word: 0,
            ignored_flags: 0,
            ignored_fixed_tail: [0; 32],
        };
        value.validate()?;
        Ok(value)
    }
    pub fn fields(&self) -> &[XlsXmlTableField] {
        &self.fields
    }
    pub const fn is_single_cell(&self) -> bool {
        self.single_cell
    }
    pub fn entry_id(&self) -> Option<&str> {
        self.entry_id.as_deref()
    }
    pub const fn ignored_fixed_word(&self) -> u16 {
        self.ignored_fixed_word
    }
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
    pub const fn ignored_fixed_tail(&self) -> &[u8; 32] {
        &self.ignored_fixed_tail
    }
    pub fn with_entry_id(mut self, v: impl Into<String>) -> XlsResult<Self> {
        let v = v.into();
        validate_name(&v, "XML entry id")?;
        self.entry_id = Some(v);
        Ok(self)
    }
    pub fn with_single_cell(mut self, v: bool) -> XlsResult<Self> {
        self.single_cell = v;
        self.validate()?;
        Ok(self)
    }
    fn validate(&self) -> XlsResult<()> {
        if self.single_cell && self.fields.len() != 1 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "single-cell XML table requires one field",
            ));
        }
        let mut ids = HashSet::new();
        for field in &self.fields {
            validate_name(&field.source_name, "XML source field name")?;
            if !ids.insert(field.column_id) {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "duplicate XML source field ownership",
                ));
            }
        }
        Ok(())
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsListObjectSourceMetadata {
    Web(XlsWebTableMetadata),
    Xml(XlsXmlTableMetadata),
}

fn append_formula(out: &mut Vec<u8>, tokens: &[u8]) -> XlsResult<()> {
    let len = u16::try_from(tokens.len())
        .map_err(|_| invalid(FEATURE11_RECORD_TYPE, "formula token length exceeds 65535"))?;
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(tokens);
    Ok(())
}
fn parse_formula(data: &[u8], offset: &mut usize, rt: u16, field: &str) -> XlsResult<Vec<u8>> {
    let len = usize::from(u16_at(data, *offset, rt, field)?);
    if len == 0 {
        return Err(invalid(rt, format!("empty {field}")));
    }
    let end = (*offset)
        .checked_add(2 + len)
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let value = data
        .get(*offset + 2..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?
        .to_vec();
    *offset = end;
    Ok(value)
}
fn append_web_info(out: &mut Vec<u8>, info: &XlsWebFieldInfo) -> XlsResult<()> {
    out.extend_from_slice(&info.locale.to_le_bytes());
    out.extend_from_slice(&info.decimal_places.to_le_bytes());
    let flags1 = u32::from(info.percent)
        | (u32::from(info.fixed_decimal) << 1)
        | (u32::from(info.date_only) << 2)
        | (info.reading_order.code() << 3)
        | (u32::from(info.rich_text) << 5)
        | (u32::from(info.unknown_rich_text) << 6)
        | (u32::from(info.alert_unknown_rich_text) << 7)
        | info.ignored_display_flags;
    out.extend_from_slice(&flags1.to_le_bytes());
    let default_type = match info.default_value {
        None => 0,
        Some(XlsWebDefaultValue::String(_)) => 1,
        Some(XlsWebDefaultValue::Boolean(_)) => 2,
        Some(XlsWebDefaultValue::Number(_) | XlsWebDefaultValue::DateTime(_)) => 3,
    };
    let flags2 = u32::from(info.read_only)
        | (u32::from(info.required) << 1)
        | (u32::from(info.minimum_set) << 2)
        | (u32::from(info.maximum_set) << 3)
        | (u32::from(info.default_value.is_some()) << 4)
        | (u32::from(info.default_today) << 5)
        | (u32::from(info.validation_formula.is_some()) << 6)
        | (u32::from(info.allow_fill_in) << 7)
        | (default_type << 8)
        | info.ignored_validation_flags;
    out.extend_from_slice(&flags2.to_le_bytes());
    if let Some(value) = &info.default_value {
        match value {
            XlsWebDefaultValue::String(v) => append_string(out, v),
            XlsWebDefaultValue::Boolean(v) => out.extend_from_slice(&u32::from(*v).to_le_bytes()),
            XlsWebDefaultValue::Number(v) | XlsWebDefaultValue::DateTime(v) => {
                out.extend_from_slice(&v.to_le_bytes())
            },
        }
    }
    if let Some(v) = &info.validation_formula {
        append_string(out, v)
    }
    out.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}
fn parse_web_info(
    data: &[u8],
    offset: &mut usize,
    kind: XlsWebColumnType,
    rt: u16,
) -> XlsResult<XlsWebFieldInfo> {
    let locale = u32_at(data, *offset, rt, "Web LCID")?;
    let decimal_places = u32_at(data, *offset + 4, rt, "Web cDec")?;
    let a = u32_at(data, *offset + 8, rt, "Web display flags")?;
    let b = u32_at(data, *offset + 12, rt, "Web validation flags")?;
    let reading_order = XlsWebReadingOrder::from_code((a >> 3) & 3)?;
    let default_set = b & 0x10 != 0;
    let default_type = ((b >> 8) & 0xff) as u8;
    *offset += 16;
    let default_value = if default_set {
        Some(match (default_type, kind) {
            (
                1,
                XlsWebColumnType::Text
                | XlsWebColumnType::Choice
                | XlsWebColumnType::MultipleChoices,
            ) => {
                let (v, end) = parse_string(data, *offset, rt, "Web default string")?;
                *offset = end;
                XlsWebDefaultValue::String(v)
            },
            (2, XlsWebColumnType::Boolean) => {
                let v = u32_at(data, *offset, rt, "Web default boolean")?;
                if v > 1 {
                    return Err(invalid(rt, "invalid Web default boolean"));
                }
                *offset += 4;
                XlsWebDefaultValue::Boolean(v != 0)
            },
            (
                3,
                XlsWebColumnType::Number | XlsWebColumnType::Currency | XlsWebColumnType::DateTime,
            ) => {
                let bytes = data
                    .get(*offset..*offset + 8)
                    .ok_or_else(|| invalid(rt, "truncated Web default number"))?;
                *offset += 8;
                let v = f64::from_le_bytes(bytes.try_into().unwrap());
                if kind == XlsWebColumnType::DateTime {
                    XlsWebDefaultValue::DateTime(v)
                } else {
                    XlsWebDefaultValue::Number(v)
                }
            },
            _ => return Err(invalid(rt, "Web default type does not match column type")),
        })
    } else {
        if default_type != 0 {
            return Err(invalid(rt, "Web default type exists without a default"));
        }
        None
    };
    let validation_formula = if b & 0x40 != 0 {
        let (v, end) = parse_string(data, *offset, rt, "Web validation formula")?;
        *offset = end;
        Some(v)
    } else {
        None
    };
    if u32_at(data, *offset, rt, "Web reserved")? != 0 {
        return Err(invalid(rt, "Web field-info reserved value must be zero"));
    }
    *offset += 4;
    let value = XlsWebFieldInfo {
        locale,
        decimal_places,
        percent: a & 1 != 0,
        fixed_decimal: a & 2 != 0,
        date_only: a & 4 != 0,
        reading_order,
        rich_text: a & 0x20 != 0,
        unknown_rich_text: a & 0x40 != 0,
        alert_unknown_rich_text: a & 0x80 != 0,
        read_only: b & 1 != 0,
        required: b & 2 != 0,
        minimum_set: b & 4 != 0,
        maximum_set: b & 8 != 0,
        default_today: b & 0x20 != 0,
        allow_fill_in: b & 0x80 != 0,
        default_value,
        validation_formula,
        ignored_display_flags: a & !0xff,
        ignored_validation_flags: b & 0xffff_0000,
    };
    value.validate(kind)?;
    Ok(value)
}
impl XlsExternalTableMetadata {
    pub fn try_new(fields: Vec<XlsExternalTableField>) -> XlsResult<Self> {
        let value = Self {
            version: XlsExternalTableVersion::Excel2007,
            build_number: 0,
            fields,
        };
        value.validate()?;
        Ok(value)
    }
    fn validate(&self) -> XlsResult<()> {
        if !(1..=256).contains(&self.fields.len()) {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external field count must be 1..=256",
            ));
        }
        let mut columns = HashSet::new();
        let mut sources = HashSet::new();
        let mut queries = HashSet::new();
        for field in &self.fields {
            field.validate()?;
            if !columns.insert(field.column_id)
                || !sources.insert(field.source_name.to_lowercase())
                || !queries.insert(field.query_field_id)
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "external column, source name, or query field ownership is duplicated",
                ));
            }
        }
        Ok(())
    }
    pub const fn version(&self) -> XlsExternalTableVersion {
        self.version
    }
    pub const fn build_number(&self) -> u16 {
        self.build_number
    }
    pub fn fields(&self) -> &[XlsExternalTableField] {
        &self.fields
    }
    pub fn with_version(mut self, version: XlsExternalTableVersion) -> Self {
        self.version = version;
        self
    }
    pub fn with_build_number(mut self, build_number: u16) -> Self {
        self.build_number = build_number;
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsOpaqueListObjectFeature {
    record_type: u16,
    base_payload: Vec<u8>,
    continuation_payloads: Vec<Vec<u8>>,
}
impl XlsOpaqueListObjectFeature {
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }
    pub fn base_payload(&self) -> &[u8] {
        &self.base_payload
    }
    pub fn continuation_payloads(&self) -> &[Vec<u8>] {
        &self.continuation_payloads
    }
    pub fn total_payload_len(&self) -> usize {
        self.base_payload.len()
            + self
                .continuation_payloads
                .iter()
                .map(|v| v.len() - 12)
                .sum::<usize>()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsOpaqueListObjectFutureRecord {
    record_type: u16,
    payload: Vec<u8>,
    continuation_payloads: Vec<Vec<u8>>,
    after_list12_count: usize,
}
impl XlsOpaqueListObjectFutureRecord {
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }
    pub fn continuation_payloads(&self) -> &[Vec<u8>] {
        &self.continuation_payloads
    }
    pub const fn after_list12_count(&self) -> usize {
        self.after_list12_count
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsListObject {
    id: XlsListObjectId,
    name: String,
    range: XlsListObjectRange,
    columns: Vec<XlsListObjectColumn>,
    style: Option<XlsListObjectStyleOptions>,
    has_header: bool,
    has_totals: bool,
    autofilter: bool,
    comment: String,
    feature_version: XlsListObjectFeatureVersion,
    opaque_feature: Option<XlsOpaqueListObjectFeature>,
    opaque_future_records: Vec<XlsOpaqueListObjectFutureRecord>,
    autofilter12_criteria: Option<XlsTableAutoFilter12>,
    external_metadata: Option<XlsExternalTableMetadata>,
    source_metadata: Option<XlsListObjectSourceMetadata>,
}
impl XlsListObject {
    pub fn try_new(
        id: XlsListObjectId,
        name: impl Into<String>,
        range: XlsListObjectRange,
        columns: Vec<XlsListObjectColumn>,
        style: XlsListObjectStyleOptions,
    ) -> XlsResult<Self> {
        let feature_version = if columns
            .iter()
            .any(|c| c.total_formula.is_some() || c.total_string.is_some())
        {
            XlsListObjectFeatureVersion::Feature12
        } else {
            XlsListObjectFeatureVersion::Feature11
        };
        let value = Self {
            id,
            name: name.into(),
            range,
            columns,
            style: Some(style),
            has_header: true,
            has_totals: false,
            autofilter: true,
            comment: String::new(),
            feature_version,
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn with_header_row(mut self, v: bool) -> XlsResult<Self> {
        self.has_header = v;
        if !v {
            self.autofilter = false;
            self.feature_version = XlsListObjectFeatureVersion::Feature12;
        } else if self.opaque_feature.is_none() {
            self.feature_version = XlsListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_totals_row(mut self, v: bool) -> XlsResult<Self> {
        self.has_totals = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter(mut self, v: bool) -> XlsResult<Self> {
        self.autofilter = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter12_criteria(mut self, value: XlsTableAutoFilter12) -> XlsResult<Self> {
        self.autofilter12_criteria = Some(value);
        self.validate()?;
        Ok(self)
    }
    pub fn with_comment(mut self, v: impl Into<String>) -> XlsResult<Self> {
        self.comment = v.into();
        if self.comment.encode_utf16().count() > 255 {
            return Err(invalid(
                LIST12_RECORD_TYPE,
                "table comment exceeds 255 characters",
            ));
        }
        Ok(self)
    }
    pub fn with_external_data(mut self, metadata: XlsExternalTableMetadata) -> XlsResult<Self> {
        metadata.validate()?;
        self.external_metadata = Some(metadata);
        self.feature_version = XlsListObjectFeatureVersion::Feature12;
        self.opaque_feature = None;
        self.validate()?;
        Ok(self)
    }
    pub fn with_web_source(mut self, metadata: XlsWebTableMetadata) -> XlsResult<Self> {
        metadata.validate()?;
        self.source_metadata = Some(XlsListObjectSourceMetadata::Web(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != XlsListObjectFeatureVersion::Feature12 {
            self.feature_version = XlsListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_xml_source(mut self, metadata: XlsXmlTableMetadata) -> XlsResult<Self> {
        metadata.validate()?;
        self.source_metadata = Some(XlsListObjectSourceMetadata::Xml(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != XlsListObjectFeatureVersion::Feature12 {
            self.feature_version = XlsListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub const fn id(&self) -> XlsListObjectId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn range(&self) -> XlsListObjectRange {
        self.range
    }
    pub fn columns(&self) -> &[XlsListObjectColumn] {
        &self.columns
    }
    pub fn style(&self) -> Option<&XlsListObjectStyleOptions> {
        self.style.as_ref()
    }
    pub const fn has_header_row(&self) -> bool {
        self.has_header
    }
    pub const fn has_totals_row(&self) -> bool {
        self.has_totals
    }
    pub const fn shows_autofilter(&self) -> bool {
        self.autofilter
    }
    pub fn comment(&self) -> &str {
        &self.comment
    }
    pub const fn feature_version(&self) -> XlsListObjectFeatureVersion {
        self.feature_version
    }
    pub fn opaque_feature(&self) -> Option<&XlsOpaqueListObjectFeature> {
        self.opaque_feature.as_ref()
    }
    pub fn opaque_future_records(&self) -> &[XlsOpaqueListObjectFutureRecord] {
        &self.opaque_future_records
    }
    pub fn autofilter12_criteria(&self) -> Option<&XlsTableAutoFilter12> {
        self.autofilter12_criteria.as_ref()
    }
    pub fn external_metadata(&self) -> Option<&XlsExternalTableMetadata> {
        self.external_metadata.as_ref()
    }
    pub fn source_metadata(&self) -> Option<&XlsListObjectSourceMetadata> {
        self.source_metadata.as_ref()
    }
    pub(crate) fn validate(&self) -> XlsResult<()> {
        validate_table_name(&self.name)?;
        if self.opaque_feature.is_none()
            && (self.columns.is_empty()
                || self.columns.len() > 256
                || self.columns.len() != self.range.column_count())
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "column count must match the table range",
            ));
        }
        if self.opaque_feature.is_some()
            && self.feature_version != XlsListObjectFeatureVersion::Feature12
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "opaque table feature must be Feature12",
            ));
        }
        if let Some(metadata) = &self.external_metadata {
            metadata.validate()?;
            if self.feature_version != XlsListObjectFeatureVersion::Feature12
                || metadata.fields.len() != self.columns.len()
                || metadata
                    .fields
                    .iter()
                    .zip(&self.columns)
                    .any(|(field, column)| field.column_id != column.id)
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "external metadata must be Feature12 and owned one-for-one by table columns",
                ));
            }
            for (field, column) in metadata.fields.iter().zip(&self.columns) {
                if field.total_array_formula != column.total_formula.is_some()
                    && field.total_array_formula
                {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "array formula metadata requires a total formula",
                    ));
                }
                if !field.total_array_formula && !field.formula_extra.is_empty() {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "scalar total formula cannot carry RgbExtra",
                    ));
                }
                if self.has_header
                    && (!field.header_cache.formatting_bytes().is_empty()
                        || field.header_cache.style_name().is_some())
                {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "cached disk header requires a headerless external table",
                    ));
                }
            }
        }
        if let Some(source) = &self.source_metadata {
            if !matches!(
                self.feature_version,
                XlsListObjectFeatureVersion::Feature11 | XlsListObjectFeatureVersion::Feature12
            ) || self.external_metadata.is_some()
                || self.opaque_feature.is_some()
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Web/XML source metadata requires a typed Feature11 or Feature12",
                ));
            }
            let has_feature12_field = self
                .columns
                .iter()
                .any(|column| column.total_formula.is_some() || column.total_string.is_some());
            if self.feature_version == XlsListObjectFeatureVersion::Feature11 && has_feature12_field
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Feature11 source fields cannot load total formulas or strings",
                ));
            }
            if self.feature_version == XlsListObjectFeatureVersion::Feature12
                && self.has_header
                && !has_feature12_field
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "Feature12 Web/XML source requires a Feature12-only property",
                ));
            }
            match source {
                XlsListObjectSourceMetadata::Web(metadata) => {
                    metadata.validate()?;
                    if !self.has_header
                        || metadata.fields.len() != self.columns.len()
                        || metadata
                            .fields
                            .iter()
                            .zip(&self.columns)
                            .any(|(field, column)| field.column_id != column.id)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "Web source fields must be owned one-for-one by a headered table",
                        ));
                    }
                    for (field, column) in metadata.fields.iter().zip(&self.columns) {
                        if column.total_formula.is_some()
                            && field.total_formula_extra.len() > MAX_FEATURE_BYTES
                        {
                            return Err(invalid(
                                FEATURE11_RECORD_TYPE,
                                "Web total formula extra data exceeds resource bound",
                            ));
                        }
                    }
                },
                XlsListObjectSourceMetadata::Xml(metadata) => {
                    metadata.validate()?;
                    if metadata
                        .entry_id
                        .as_deref()
                        .is_some_and(|entry| entry != self.id.value().to_string())
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "XML entry id must equal the decimal table id",
                        ));
                    }
                    if metadata.fields.len() != self.columns.len()
                        || metadata
                            .fields
                            .iter()
                            .zip(&self.columns)
                            .any(|(field, column)| field.column_id != column.id)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "XML source fields must be owned one-for-one by table columns",
                        ));
                    }
                    if metadata.single_cell
                        && (self.has_header
                            || self.has_totals
                            || self.columns.len() != 1
                            || self.range.first_row != self.range.last_row
                            || self.range.first_column != self.range.last_column)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "single-cell XML source requires one unheadered cell",
                        ));
                    }
                    if !metadata.single_cell && !self.has_header {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "multi-cell XML source requires a header row",
                        ));
                    }
                },
            }
        }
        if self.autofilter && !self.has_header {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "AutoFilter requires a header row",
            ));
        }
        if let Some(filter) = &self.autofilter12_criteria {
            filter.validate()?;
            if !self.autofilter
                || !self.has_header
                || usize::from(filter.column_index()) >= self.range.column_count()
            {
                return Err(invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "typed AutoFilter12 criteria require an in-range column on a headered table AutoFilter",
                ));
            }
            if self
                .opaque_future_records
                .iter()
                .any(|future| future.record_type == AUTO_FILTER12_RECORD_TYPE)
            {
                return Err(invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "typed and opaque AutoFilter12 records cannot coexist",
                ));
            }
        }
        if self.has_totals && self.range.first_row == self.range.last_row {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "totals row requires a range below the header",
            ));
        }
        let mut ids = HashSet::new();
        let mut names = HashSet::new();
        for column in &self.columns {
            column.validate_totals()?;
            if !ids.insert(column.id) || !names.insert(column.name.to_lowercase()) {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "duplicate column id or name",
                ));
            }
        }
        if self.style.is_none() {
            return Err(invalid(LIST12_RECORD_TYPE, "missing table style"));
        }
        Ok(())
    }
    pub(crate) fn to_feature_record_bytes(&self) -> XlsResult<Vec<Vec<u8>>> {
        self.validate()?;
        if let Some(opaque) = &self.opaque_feature {
            let mut records = vec![record(opaque.record_type, opaque.base_payload.clone())?];
            for payload in &opaque.continuation_payloads {
                records.push(record(CONTINUE_FRT11_RECORD_TYPE, payload.clone())?);
            }
            return Ok(records);
        }
        if let Some(metadata) = &self.external_metadata {
            return self.to_external_feature_record_bytes(metadata);
        }
        if let Some(metadata) = &self.source_metadata {
            return self.to_source_feature_record_bytes(metadata);
        }
        let rt = match self.feature_version {
            XlsListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            XlsListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&0u32.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|c| c.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&[0; 4]);
        let mut flags = 0x001B_0000u32;
        if self.autofilter {
            flags |= 0x806;
        }
        if self.has_totals {
            flags |= 0x40;
        }
        feature.extend_from_slice(&flags.to_le_bytes());
        feature.extend_from_slice(&[0; 32]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(self.columns.len() as u16).to_le_bytes());
        append_string(&mut feature, &self.id.value().to_string());
        for column in &self.columns {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let cflags = u32::from(self.autofilter)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(column.total_string.is_some()) << 10);
            feature.extend_from_slice(&cflags.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, &column.name);
            append_string(&mut feature, &column.name);
            if self.autofilter {
                feature.extend_from_slice(&0u32.to_le_bytes());
                feature.extend_from_slice(&(column.id.value() as u16).to_le_bytes());
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
                feature.extend_from_slice(tokens);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first_len].to_vec())?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
    fn to_source_feature_record_bytes(
        &self,
        source: &XlsListObjectSourceMetadata,
    ) -> XlsResult<Vec<Vec<u8>>> {
        let rt = match self.feature_version {
            XlsListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            XlsListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let (lt, version, build, single, fields_len): (u32, _, _, _, _) = match source {
            XlsListObjectSourceMetadata::Web(v) => {
                (1u32, v.version, v.build_number, false, v.fields.len())
            },
            XlsListObjectSourceMetadata::Xml(v) => (
                2u32,
                v.version,
                v.build_number,
                v.single_cell,
                v.fields.len(),
            ),
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&lt.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|c| c.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&build.to_le_bytes());
        let ignored_fixed_word = match source {
            XlsListObjectSourceMetadata::Web(v) => v.ignored_fixed_word,
            XlsListObjectSourceMetadata::Xml(v) => v.ignored_fixed_word,
        };
        feature.extend_from_slice(&ignored_fixed_word.to_le_bytes());
        let mut flags = (version.code() << 16)
            | (u32::from(self.autofilter) * 0x806)
            | (u32::from(self.has_totals) * 0x40)
            | (u32::from(single) << 9)
            | 0x0040_0000;
        match source {
            XlsListObjectSourceMetadata::Web(v) => {
                flags |= u32::from(!v.deleted_row_ids.is_empty()) << 5
                    | u32::from(v.needs_commit) << 8
                    | u32::from(v.compressed_cache) << 13
                    | u32::from(v.provider_name.is_some()) << 14
                    | u32::from(!v.changed_row_ids.is_empty()) << 15
                    | u32::from(v.entry_id.is_some()) << 20
                    | u32::from(!v.invalid_cells.is_empty()) << 21
                    | v.ignored_flags
            },
            XlsListObjectSourceMetadata::Xml(v) => {
                flags |= u32::from(v.entry_id.is_some()) << 20 | v.ignored_flags
            },
        };
        feature.extend_from_slice(&flags.to_le_bytes());
        match source {
            XlsListObjectSourceMetadata::Web(v) => {
                feature.extend_from_slice(&v.cache_position.to_le_bytes());
                feature.extend_from_slice(&v.cache_size.to_le_bytes());
                feature.extend_from_slice(&v.cache_characters.to_le_bytes());
                feature.extend_from_slice(&v.edit_mode.code().to_le_bytes());
                feature.extend_from_slice(&v.hash_parameters)
            },
            XlsListObjectSourceMetadata::Xml(v) => feature.extend_from_slice(&v.ignored_fixed_tail),
        };
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(fields_len as u16).to_le_bytes());
        match source {
            XlsListObjectSourceMetadata::Web(v) => {
                if let Some(name) = &v.provider_name {
                    append_string(&mut feature, name)
                }
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
            XlsListObjectSourceMetadata::Xml(v) => {
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
        }
        for (index, column) in self.columns.iter().enumerate() {
            let (web, xml) = match source {
                XlsListObjectSourceMetadata::Web(v) => (Some(&v.fields[index]), None),
                XlsListObjectSourceMetadata::Xml(v) => (None, Some(&v.fields[index])),
            };
            let (
                source_name,
                web_type,
                xml_type,
                mapped,
                calc,
                auto_filter,
                agg_fmt,
                insert_fmt,
                total_extra,
                ignored_flags,
            ) = if let Some(v) = web {
                (
                    &v.source_name,
                    v.data_type.code(),
                    0,
                    false,
                    v.calculated_formula.as_deref(),
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            } else {
                let v = xml.unwrap();
                (
                    &v.source_name,
                    0,
                    v.data_type.value(),
                    v.mapping.is_some(),
                    None,
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            };
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&web_type.to_le_bytes());
            feature.extend_from_slice(&xml_type.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&(agg_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let ff = u32::from(self.autofilter)
                | (u32::from(mapped) << 2)
                | (u32::from(calc.is_some()) << 3)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(!total_extra.is_empty()) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | ignored_flags;
            feature.extend_from_slice(&ff.to_le_bytes());
            feature.extend_from_slice(&(insert_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, source_name);
            if !single {
                append_string(&mut feature, &column.name)
            }
            feature.extend_from_slice(agg_fmt);
            feature.extend_from_slice(insert_fmt);
            if self.autofilter {
                feature.extend_from_slice(auto_filter)
            }
            if let Some(v) = xml.and_then(|v| v.mapping.as_ref()) {
                feature.extend_from_slice(&1u16.to_le_bytes());
                feature.extend_from_slice(&(2u32 | u32::from(v.can_be_single) << 2).to_le_bytes());
                feature.extend_from_slice(&v.map_id.to_le_bytes());
                append_string(&mut feature, &v.xpath)
            }
            if let Some(tokens) = calc {
                append_formula(&mut feature, tokens)?
            }
            if let Some(tokens) = &column.total_formula {
                append_formula(&mut feature, tokens)?;
                feature.extend_from_slice(total_extra)
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value)
            }
            if let Some(v) = web {
                append_web_info(&mut feature, &v.info)?
            }
        }
        if let XlsListObjectSourceMetadata::Web(v) = source {
            if !v.deleted_row_ids.is_empty() {
                feature.extend_from_slice(&(v.deleted_row_ids.len() as u16).to_le_bytes());
                for id in &v.deleted_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.changed_row_ids.is_empty() {
                feature.extend_from_slice(&(v.changed_row_ids.len() as u16).to_le_bytes());
                for id in &v.changed_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.invalid_cells.is_empty() {
                feature.extend_from_slice(&(v.invalid_cells.len() as u16).to_le_bytes());
                for cell in &v.invalid_cells {
                    feature.extend_from_slice(&cell.row_id.to_le_bytes());
                    feature.extend_from_slice(&cell.column_id.value().to_le_bytes())
                }
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&(feature.len() as u32).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "Web/XML table feature exceeds aggregate resource bound",
            ));
        }
        let first = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first].to_vec())?];
        for chunk in payload[first..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?)
        }
        Ok(records)
    }
    fn to_external_feature_record_bytes(
        &self,
        metadata: &XlsExternalTableMetadata,
    ) -> XlsResult<Vec<Vec<u8>>> {
        let mut feature = Vec::new();
        feature.extend_from_slice(&3u32.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|column| column.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE12_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&metadata.build_number.to_le_bytes());
        feature.extend_from_slice(&0u16.to_le_bytes());
        let mut flags = metadata.version.code() << 16 | 0x0010_0000;
        if self.autofilter {
            flags |= 0x0000_0806;
        }
        if self.has_totals {
            flags |= 0x40;
        }
        feature.extend_from_slice(&flags.to_le_bytes());
        feature.extend_from_slice(&[0; 12]);
        feature.extend_from_slice(&0u32.to_le_bytes());
        feature.extend_from_slice(&[0; 16]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(self.columns.len() as u16).to_le_bytes());
        append_string(&mut feature, &self.id.value().to_string());
        for (column, field) in self.columns.iter().zip(&metadata.fields) {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&(field.aggregate_format.len() as u32).to_le_bytes());
            feature.extend_from_slice(&field.aggregate_style.to_le_bytes());
            let mut field_flags = u32::from(self.autofilter)
                | (u32::from(field.filter_hidden) << 1)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(field.total_array_formula) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | (u32::from(field.auto_create_calculated_column) << 11);
            if !self.has_header && field.header_cache.style_name().is_some() {
                field_flags |= 0x200;
            }
            feature.extend_from_slice(&field_flags.to_le_bytes());
            feature.extend_from_slice(&(field.insert_row_format.len() as u32).to_le_bytes());
            feature.extend_from_slice(&field.insert_row_style.to_le_bytes());
            append_string(&mut feature, &field.source_name);
            append_string(&mut feature, &column.name);
            feature.extend_from_slice(&field.aggregate_format);
            feature.extend_from_slice(&field.insert_row_format);
            if self.autofilter {
                feature.extend_from_slice(&field.auto_filter);
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
                feature.extend_from_slice(tokens);
                feature.extend_from_slice(&field.formula_extra);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
            feature.extend_from_slice(&field.query_field_id.to_le_bytes());
            if !self.has_header {
                feature.extend_from_slice(field.header_cache.as_bytes());
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, FEATURE12_RECORD_TYPE, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&(feature.len() as u32).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(
            FEATURE12_RECORD_TYPE,
            payload[..first_len].to_vec(),
        )?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
    pub(crate) fn to_list12_record_bytes(&self) -> XlsResult<Vec<Vec<u8>>> {
        let mut block = Vec::new();
        append_frt(&mut block, LIST12_RECORD_TYPE, None);
        block.extend_from_slice(&0u16.to_le_bytes());
        block.extend_from_slice(&self.id.value().to_le_bytes());
        for v in [0i32, -1, 0, -1, 0, -1, 0, 0, 0] {
            block.extend_from_slice(&v.to_le_bytes());
        }
        let style = self.style.as_ref().unwrap();
        let mut styled = Vec::new();
        append_frt(&mut styled, LIST12_RECORD_TYPE, None);
        styled.extend_from_slice(&1u16.to_le_bytes());
        styled.extend_from_slice(&self.id.value().to_le_bytes());
        let bits = u16::from(style.first)
            | u16::from(style.last) << 1
            | u16::from(style.row_stripes) << 2
            | u16::from(style.column_stripes) << 3
            | u16::from(style.default_style) << 6;
        styled.extend_from_slice(&bits.to_le_bytes());
        append_string(&mut styled, &style.name);
        let mut display = Vec::new();
        append_frt(&mut display, LIST12_RECORD_TYPE, None);
        display.extend_from_slice(&2u16.to_le_bytes());
        display.extend_from_slice(&self.id.value().to_le_bytes());
        append_string(&mut display, &self.name);
        append_string(&mut display, &self.comment);
        Ok(vec![
            record(LIST12_RECORD_TYPE, block)?,
            record(LIST12_RECORD_TYPE, styled)?,
            record(LIST12_RECORD_TYPE, display)?,
        ])
    }
    pub(crate) fn to_following_record_bytes(&self) -> XlsResult<Vec<Vec<u8>>> {
        let list12 = self.to_list12_record_bytes()?;
        let mut output = Vec::new();
        for (index, item) in list12.into_iter().enumerate() {
            output.push(item);
            if index == 0 {
                if let Some(filter) = &self.autofilter12_criteria {
                    output.extend(write_table_autofilter12(filter, self.range, self.id)?);
                }
            }
            for future in self
                .opaque_future_records
                .iter()
                .filter(|v| v.after_list12_count == index + 1)
            {
                output.push(record(future.record_type, future.payload.clone())?);
                for payload in &future.continuation_payloads {
                    output.push(record(
                        crate::xls::sort_data::CONTINUE_FRT12_RECORD_TYPE,
                        payload.clone(),
                    )?);
                }
            }
        }
        if self
            .opaque_future_records
            .iter()
            .any(|v| v.after_list12_count == 0 || v.after_list12_count > 3)
        {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "opaque table future-record insertion point is invalid",
            ));
        }
        Ok(output)
    }
    fn parse_opaque_feature12(pending: PendingFeature) -> XlsResult<Self> {
        let data = &pending.combined;
        if data.len() < 108 {
            return Err(invalid(FEATURE12_RECORD_TYPE, "truncated opaque Feature12"));
        }
        validate_frt(data, FEATURE12_RECORD_TYPE, true)?;
        if u16_at(data, 12, FEATURE12_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 55, FEATURE12_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 fixed fields",
            ));
        }
        let range = parse_range(data, 4, FEATURE12_RECORD_TYPE)?;
        let id = XlsListObjectId::try_new(u32_at(data, 39, FEATURE12_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, 43, FEATURE12_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, 47, FEATURE12_RECORD_TYPE, "crwTotals")?;
        if header > 1 || totals > 1 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 row flags",
            ));
        }
        let flags = u32_at(data, 63, FEATURE12_RECORD_TYPE, "flags")?;
        let (name, _) = parse_string(data, 99, FEATURE12_RECORD_TYPE, "rgbName")?;
        let opaque_feature = XlsOpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns: Vec::new(),
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: XlsListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
    fn parse_external_feature12(pending: PendingFeature) -> XlsResult<Self> {
        let data = &pending.combined;
        let rt = FEATURE12_RECORD_TYPE;
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid external table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, rt, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, rt, "reserved2")? != 0
            || u16_at(data, 19, rt, "cref2")? != 1
            || u16_at(data, 25, rt, "reserved3")? != 0
            || parse_range(data, 27, rt)? != range
        {
            return Err(invalid(rt, "invalid Feature12 fixed fields"));
        }
        let declared = usize::try_from(u32_at(data, 21, rt, "cbFeatData")?)
            .map_err(|_| invalid(rt, "cbFeatData overflows"))?;
        if declared != 0 && declared != data.len() - 35 {
            return Err(invalid(
                rt,
                "cbFeatData does not match external feature size",
            ));
        }
        let base = 35;
        if u32_at(data, base, rt, "lt")? != 3 {
            return Err(invalid(rt, "external parser requires LTEXTERNALDATA"));
        }
        let id = XlsListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, rt, "cbFSData")? != 64
            || u32_at(data, base + 44, rt, "lem")? != 0
        {
            return Err(invalid(
                rt,
                "invalid external TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let version = XlsExternalTableVersion::from_code((flags >> 16) & 0xF)?;
        if flags & 0x0020_E7A0 != 0
            || flags & 0x4 != 0 && flags & 0x2 == 0
            || flags & 0x10 != 0 && flags & 0x8 == 0
            || flags & 0x2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid external table flags"));
        }
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "external field count must match table range"));
        }
        if flags & 0x0000_4000 != 0 {
            return Err(invalid(
                rt,
                "external table cannot load a SharePoint CSP name",
            ));
        }
        if flags & 0x0010_0000 != 0 {
            offset = parse_string(data, offset, rt, "entryId")?.1;
        }
        let mut columns = Vec::with_capacity(count);
        let mut fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = XlsListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            if u32_at(data, start + 4, rt, "lfdt")? != 0
                || u32_at(data, start + 8, rt, "lfxidt")? != 0
            {
                return Err(invalid(rt, "external field data types must be zero"));
            }
            let aggregation =
                XlsListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let aggregate_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_style = u32_at(data, start + 20, rt, "istnAgg")?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert format length overflows"))?;
            let insert_row_style = u32_at(data, start + 32, rt, "istnInsertRow")?;
            if field_flags & 0x4C != 0
                || field_flags & 0x40 != 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != XlsListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != XlsListTotalAggregation::None
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || (field_flags & 1 != 0) != (flags & 2 != 0)
            {
                return Err(invalid(rt, "invalid external field flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            validate_name(&source_name, "external source field name")?;
            let (caption, after_caption) = parse_string(data, after_source, rt, "strCaption")?;
            validate_column_name(&caption)?;
            let aggregate_end = after_caption
                .checked_add(aggregate_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..aggregate_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = aggregate_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(aggregate_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_| invalid(rt, "AutoFilter length overflows"))?;
                if size > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + size)
                    .ok_or_else(|| invalid(rt, "AutoFilter length overflows"))?;
                let value = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                value
            } else {
                vec![0; 6]
            };
            let (total_formula, formula_extra) = if field_flags & 0x80 != 0 {
                let size = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if size == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let token_end = offset
                    .checked_add(2 + size)
                    .ok_or_else(|| invalid(rt, "total formula length overflows"))?;
                let tokens = data
                    .get(offset + 2..token_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = token_end;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            let query_field_id = u32_at(data, offset, rt, "qsif")?;
            if query_field_id == 0 {
                return Err(invalid(rt, "external qsif must be nonzero"));
            }
            offset += 4;
            let header_cache = if header == 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbdxfHdrDisk")?)
                    .map_err(|_| invalid(rt, "header cache length overflows"))?;
                let format_end = offset
                    .checked_add(4 + size)
                    .ok_or_else(|| invalid(rt, "header cache length overflows"))?;
                data.get(offset..format_end)
                    .ok_or_else(|| invalid(rt, "truncated header cache"))?;
                let end = if field_flags & 0x200 != 0 {
                    parse_string(data, format_end, rt, "header style name")?.1
                } else {
                    format_end
                };
                let value = XlsCachedDiskHeader::parse(
                    data[offset..end].to_vec(),
                    field_flags & 0x200 != 0,
                    rt,
                )?;
                offset = end;
                value
            } else {
                if field_flags & 0x200 != 0 {
                    return Err(invalid(rt, "header style name requires a CachedDiskHeader"));
                }
                XlsCachedDiskHeader::empty()
            };
            let column = XlsListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
            fields.push(XlsExternalTableField {
                column_id: cid,
                source_name,
                query_field_id,
                aggregate_format,
                insert_row_format,
                auto_filter,
                formula_extra,
                header_cache,
                aggregate_style,
                insert_row_style,
                filter_hidden: field_flags & 2 != 0,
                total_array_formula: field_flags & 0x100 != 0,
                auto_create_calculated_column: field_flags & 0x800 != 0,
            });
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing external Feature12 data"));
        }
        let metadata = XlsExternalTableMetadata {
            version,
            build_number: u16_at(data, base + 24, rt, "rupBuild")?,
            fields,
        };
        metadata.validate()?;
        let opaque_feature = XlsOpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: XlsListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: Some(metadata),
            source_metadata: None,
        })
    }
    fn parse_source_feature(data: &[u8], rt: u16, lt: u32) -> XlsResult<Self> {
        if !matches!(rt, FEATURE11_RECORD_TYPE | FEATURE12_RECORD_TYPE) {
            return Err(invalid(
                rt,
                "Web/XML source table requires Feature11 or Feature12",
            ));
        }
        let range = parse_range(data, 4, rt)?;
        let base = 35;
        let id = XlsListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1 || totals > 1 || u32_at(data, base + 20, rt, "cbFSData")? != 64 {
            return Err(invalid(rt, "invalid Web/XML TableFeatureType fixed fields"));
        }
        let build = u16_at(data, base + 24, rt, "rupBuild")?;
        let ignored_fixed_word = u16_at(data, base + 26, rt, "unused1")?;
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let version = XlsExternalTableVersion::from_code((flags >> 16) & 0xf)?;
        if flags & 0x0000_0480 != 0
            || flags & 4 != 0 && flags & 2 == 0
            || flags & 0x10 != 0 && flags & 8 == 0
            || flags & 2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid Web/XML table flags"));
        }
        let single = flags & 0x200 != 0;
        if single && (lt != 2 || header != 0 || totals != 0 || range.column_count() != 1) {
            return Err(invalid(rt, "invalid single-cell XML table"));
        }
        if lt == 2 && flags & 0x0020_e120 != 0 {
            return Err(invalid(rt, "Web-only flags occur on XML table"));
        }
        let mut ignored_fixed_tail = [0; 32];
        let (cache_position, cache_size, cache_characters, edit_mode, hash) = if lt == 1 {
            let mut hash = [0; 16];
            hash.copy_from_slice(
                data.get(base + 48..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated Web hash parameters"))?,
            );
            (
                u32_at(data, base + 32, rt, "cache position")?,
                u32_at(data, base + 36, rt, "cache size")?,
                u32_at(data, base + 40, rt, "cache characters")?,
                XlsWebEditMode::from_code(u32_at(data, base + 44, rt, "edit mode")?)?,
                hash,
            )
        } else {
            ignored_fixed_tail.copy_from_slice(
                data.get(base + 32..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated XML fixed tail"))?,
            );
            if ignored_fixed_tail[12..16].iter().any(|byte| *byte != 0) {
                return Err(invalid(rt, "XML edit mode must be zero"));
            }
            (0, 0, 0, XlsWebEditMode::Normal, [0; 16])
        };
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "Web/XML field count must match table range"));
        }
        let provider_name = if flags & 0x4000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "cSPName")?;
            offset = end;
            Some(v)
        } else {
            None
        };
        let entry_id = if flags & 0x0010_0000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "entryId")?;
            offset = end;
            if lt == 2 && v != id.value().to_string() {
                return Err(invalid(rt, "XML entryId does not match table id"));
            }
            Some(v)
        } else {
            None
        };
        let mut columns = Vec::with_capacity(count);
        let mut web_fields = Vec::with_capacity(count);
        let mut xml_fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = XlsListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            let web_type = u32_at(data, start + 4, rt, "lfdt")?;
            let xml_type = u32_at(data, start + 8, rt, "lfxidt")?;
            if (lt == 1 && xml_type != 0) || (lt == 2 && web_type != 0) {
                return Err(invalid(rt, "field data type does not match table source"));
            }
            let aggregation =
                XlsListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let agg_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert-row format length overflows"))?;
            if field_flags & 0x0000_0040 != 0
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != XlsListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != XlsListTotalAggregation::None
                || (field_flags & 1 != 0) != (flags & 2 != 0)
                || (lt == 1 && field_flags & 0x804 != 0)
                || (lt == 2 && field_flags & 8 != 0)
                || (rt == FEATURE11_RECORD_TYPE && field_flags & 0x480 != 0)
            {
                return Err(invalid(rt, "invalid source field condition flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            let (caption, after_caption) = if single {
                (source_name.clone(), after_source)
            } else {
                parse_string(data, after_source, rt, "strCaption")?
            };
            validate_column_name(&caption)?;
            let agg_end = after_caption
                .checked_add(agg_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..agg_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = agg_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(agg_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let n = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_| invalid(rt, "AutoFilter size overflows"))?;
                if n > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(rt, "AutoFilter size overflows"))?;
                let v = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                v
            } else {
                vec![0; 6]
            };
            let mapping = if field_flags & 4 != 0 {
                if u16_at(data, offset, rt, "iXmapMac")? != 1 {
                    return Err(invalid(rt, "XML mapped field must contain one map entry"));
                }
                let map_flags = u32_at(data, offset + 2, rt, "XMap flags")?;
                if map_flags & !6 != 0 || map_flags & 2 == 0 {
                    return Err(invalid(rt, "invalid XML map flags"));
                }
                let map_id = u32_at(data, offset + 6, rt, "XML map id")?;
                let (xpath, end) = parse_string(data, offset + 10, rt, "XPath")?;
                offset = end;
                Some(XlsXmlColumnMapping::try_new(
                    map_id,
                    xpath,
                    map_flags & 4 != 0,
                )?)
            } else {
                None
            };
            let calculated_formula = if field_flags & 8 != 0 {
                Some(parse_formula(data, &mut offset, rt, "calculated formula")?)
            } else {
                None
            };
            let (total_formula, total_extra) = if field_flags & 0x80 != 0 {
                let tokens = parse_formula(data, &mut offset, rt, "total formula")?;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (v, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(v)
            } else {
                None
            };
            let web_kind = if lt == 1 {
                Some(XlsWebColumnType::from_code(web_type)?)
            } else {
                None
            };
            let web_info = if let Some(kind) = web_kind {
                Some(parse_web_info(data, &mut offset, kind, rt)?)
            } else {
                None
            };
            if header == 0 && !single {
                let n = usize::try_from(u32_at(data, offset, rt, "cached header format size")?)
                    .map_err(|_| invalid(rt, "cached header size overflows"))?;
                let end = offset
                    .checked_add(4 + n)
                    .ok_or_else(|| invalid(rt, "cached header size overflows"))?;
                data.get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated cached header"))?;
                offset = if field_flags & 0x200 != 0 {
                    parse_string(data, end, rt, "cached header style")?.1
                } else {
                    end
                }
            } else if field_flags & 0x200 != 0 {
                return Err(invalid(rt, "cached header style lacks cached header"));
            }
            let column = XlsListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
            if let (Some(kind), Some(info)) = (web_kind, web_info) {
                web_fields.push(XlsWebTableField {
                    column_id: cid,
                    source_name,
                    data_type: kind,
                    info,
                    calculated_formula,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            } else {
                xml_fields.push(XlsXmlTableField {
                    column_id: cid,
                    source_name,
                    data_type: XlsXmlDataType::try_new(xml_type)?,
                    mapping,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            }
        }
        let source_metadata = if lt == 1 {
            let parse_ids = |data: &[u8], offset: &mut usize, label: &str| -> XlsResult<Vec<u32>> {
                let count = usize::from(u16_at(data, *offset, rt, label)?);
                *offset += 2;
                let end = (*offset)
                    .checked_add(
                        count
                            .checked_mul(4)
                            .ok_or_else(|| invalid(rt, "source id count overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "source id count overflows"))?;
                let bytes = data
                    .get(*offset..end)
                    .ok_or_else(|| invalid(rt, format!("truncated {label}")))?;
                let ids = bytes
                    .chunks_exact(4)
                    .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
                    .collect();
                *offset = end;
                Ok(ids)
            };
            let deleted_row_ids = if flags & 0x20 != 0 {
                parse_ids(data, &mut offset, "deleted row ids")?
            } else {
                Vec::new()
            };
            let changed_row_ids = if flags & 0x8000 != 0 {
                parse_ids(data, &mut offset, "changed row ids")?
            } else {
                Vec::new()
            };
            let invalid_cells = if flags & 0x0020_0000 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "invalid cell count")?);
                offset += 2;
                let mut out = Vec::with_capacity(n);
                for _ in 0..n {
                    let row = u32_at(data, offset, rt, "invalid cell row")?;
                    let column = XlsListColumnId::try_new(u32_at(
                        data,
                        offset + 4,
                        rt,
                        "invalid cell field",
                    )?)?;
                    offset += 8;
                    out.push(XlsWebInvalidCell::new(row, column))
                }
                out
            } else {
                Vec::new()
            };
            XlsListObjectSourceMetadata::Web(XlsWebTableMetadata {
                version,
                build_number: build,
                fields: web_fields,
                edit_mode,
                cache_position,
                cache_size,
                cache_characters,
                hash_parameters: hash,
                provider_name,
                entry_id,
                deleted_row_ids,
                changed_row_ids,
                invalid_cells,
                needs_commit: flags & 0x100 != 0,
                compressed_cache: flags & 0x2000 != 0,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
            })
        } else {
            XlsListObjectSourceMetadata::Xml(XlsXmlTableMetadata {
                version,
                build_number: build,
                fields: xml_fields,
                entry_id,
                single_cell: single,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
                ignored_fixed_tail,
            })
        };
        let has_feature12_field = columns
            .iter()
            .any(|column| column.total_formula.is_some() || column.total_string.is_some());
        if rt == FEATURE12_RECORD_TYPE && header != 0 && !has_feature12_field {
            return Err(invalid(
                rt,
                "Feature12 Web/XML source lacks a Feature12-only property",
            ));
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing Web/XML feature data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: if rt == FEATURE12_RECORD_TYPE {
                XlsListObjectFeatureVersion::Feature12
            } else {
                XlsListObjectFeatureVersion::Feature11
            },
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: Some(source_metadata),
        })
    }
    fn parse_feature(data: &[u8], rt: u16) -> XlsResult<Self> {
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, FEATURE11_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, FEATURE11_RECORD_TYPE, "reserved2")? != 0
            || u16_at(data, 19, FEATURE11_RECORD_TYPE, "cref2")? != 1
            || u16_at(data, 25, FEATURE11_RECORD_TYPE, "reserved3")? != 0
            || parse_range(data, 27, FEATURE11_RECORD_TYPE)? != range
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid Feature11 fixed fields",
            ));
        }
        let base = 35;
        let source_type = u32_at(data, base, FEATURE11_RECORD_TYPE, "lt")?;
        if matches!(source_type, 1 | 2) {
            return Self::parse_source_feature(data, rt, source_type);
        }
        if source_type != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported table source type",
            ));
        }
        let id =
            XlsListObjectId::try_new(u32_at(data, base + 4, FEATURE11_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, base + 8, FEATURE11_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, base + 12, FEATURE11_RECORD_TYPE, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, FEATURE11_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, FEATURE11_RECORD_TYPE, "flags")?;
        if flags & 0x0020_E320 != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported external table flags",
            ));
        }
        let (name, mut offset) = parse_string(data, base + 64, FEATURE11_RECORD_TYPE, "rgbName")?;
        let count = usize::from(u16_at(data, offset, FEATURE11_RECORD_TYPE, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) {
            return Err(invalid(FEATURE11_RECORD_TYPE, "invalid table column count"));
        }
        if flags & 0x0010_0000 != 0 {
            let (entry, next) = parse_string(data, offset, FEATURE11_RECORD_TYPE, "entryId")?;
            if entry != id.value().to_string() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "entryId mismatch"));
            }
            offset = next;
        }
        let mut columns = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid =
                XlsListColumnId::try_new(u32_at(data, start, FEATURE11_RECORD_TYPE, "idField")?)?;
            if u32_at(data, start + 4, FEATURE11_RECORD_TYPE, "lfdt")? != 0
                || u32_at(data, start + 8, FEATURE11_RECORD_TYPE, "lfxidt")? != 0
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external column data is unsupported",
                ));
            }
            let agg = u32_at(data, start + 16, FEATURE11_RECORD_TYPE, "cbFmtAgg")? as usize;
            let cflags = u32_at(data, start + 24, FEATURE11_RECORD_TYPE, "column flags")?;
            let insert = u32_at(data, start + 28, FEATURE11_RECORD_TYPE, "cbFmtInsert")? as usize;
            let aggregation = XlsListTotalAggregation::from_code(u32_at(
                data,
                start + 12,
                FEATURE11_RECORD_TYPE,
                "ilta",
            )?)?;
            // Feature11 permits calculated columns (fAutoCreateCalcCol) but not
            // Feature12-only XML/Web mappings or loaded total formulas/strings.
            let forbidden = 0x4c
                | if rt == FEATURE11_RECORD_TYPE {
                    0x580
                } else {
                    0
                };
            if cflags & forbidden != 0 || cflags & 0x100 != 0 {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external, array, or reserved column data is unsupported",
                ));
            }
            let (_, after_field) =
                parse_string(data, start + 36, FEATURE11_RECORD_TYPE, "field name")?;
            let (caption, after_caption) =
                parse_string(data, after_field, FEATURE11_RECORD_TYPE, "caption")?;
            offset = after_caption
                .checked_add(agg + insert)
                .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column size overflow"))?;
            if cflags & 1 != 0 {
                let n = u32_at(data, offset, FEATURE11_RECORD_TYPE, "AutoFilter size")? as usize;
                offset = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "AutoFilter size overflow"))?;
            }
            let total_formula = if cflags & 0x80 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if n == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let end = offset + 2 + n;
                let value = data
                    .get(offset + 2..end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = end;
                Some(value)
            } else {
                None
            };
            let total_string = if cflags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "total string")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            if cflags & 0x200 != 0 {
                offset = parse_string(data, offset, FEATURE11_RECORD_TYPE, "cached style")?.1;
            }
            if offset > data.len() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "truncated column data"));
            }
            let column = XlsListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
        }
        if offset != data.len() {
            return Err(invalid(FEATURE11_RECORD_TYPE, "trailing Feature11 data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: if rt == FEATURE12_RECORD_TYPE {
                XlsListObjectFeatureVersion::Feature12
            } else {
                XlsListObjectFeatureVersion::Feature11
            },
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
    fn apply_list12(&mut self, data: &[u8]) -> XlsResult<u16> {
        if data.len() < 18 {
            return Err(invalid(LIST12_RECORD_TYPE, "truncated List12"));
        }
        validate_frt(data, LIST12_RECORD_TYPE, false)?;
        let kind = u16_at(data, 12, LIST12_RECORD_TYPE, "lsd")?;
        if u32_at(data, 14, LIST12_RECORD_TYPE, "idList")? != self.id.value() {
            return Err(invalid(LIST12_RECORD_TYPE, "List12 id mismatch"));
        }
        match kind {
            0 => {
                if data.len() < 54 {
                    return Err(invalid(LIST12_RECORD_TYPE, "truncated block-level List12"));
                }
            },
            1 => {
                let bits = u16_at(data, 18, LIST12_RECORD_TYPE, "style flags")?;
                let (name, end) = parse_string(data, 20, LIST12_RECORD_TYPE, "style name")?;
                if end != data.len() {
                    return Err(invalid(LIST12_RECORD_TYPE, "trailing style List12 data"));
                }
                self.style = Some(XlsListObjectStyleOptions {
                    name,
                    first: bits & 1 != 0,
                    last: bits & 2 != 0,
                    row_stripes: bits & 4 != 0,
                    column_stripes: bits & 8 != 0,
                    default_style: bits & 0x40 != 0,
                });
            },
            2 => {
                let (name, next) = parse_string(data, 18, LIST12_RECORD_TYPE, "display name")?;
                let (comment, end) = parse_string(data, next, LIST12_RECORD_TYPE, "comment")?;
                if end != data.len() || (!name.is_empty() && name != self.name) {
                    return Err(invalid(LIST12_RECORD_TYPE, "inconsistent display List12"));
                }
                self.comment = comment;
            },
            _ => return Err(invalid(LIST12_RECORD_TYPE, "reserved List12 type")),
        }
        Ok(kind)
    }
}

struct PendingFeature {
    record_type: u16,
    base: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    combined: Vec<u8>,
}
struct PendingFuture {
    payload: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    after_list12_count: usize,
}
pub(crate) struct ListObjectCollector {
    header: Option<u32>,
    pending: Option<PendingFeature>,
    current: Option<XlsListObject>,
    pending_future: Option<PendingFuture>,
    kinds: HashSet<u16>,
    list12_count: usize,
    sort_continuations: usize,
    tables: Vec<XlsListObject>,
    ended: bool,
}
impl ListObjectCollector {
    pub(crate) fn new() -> Self {
        Self {
            header: None,
            pending: None,
            current: None,
            pending_future: None,
            kinds: HashSet::new(),
            list12_count: 0,
            sort_continuations: 0,
            tables: Vec::new(),
            ended: false,
        }
    }
    pub(crate) fn feed_record(&mut self, rt: u16, data: &[u8]) -> XlsResult<()> {
        if self.header.is_none()
            && matches!(
                rt,
                AUTO_FILTER12_RECORD_TYPE
                    | crate::xls::sort_data::SORT_DATA_RECORD_TYPE
                    | crate::xls::sort_data::CONTINUE_FRT12_RECORD_TYPE
            )
        {
            return Ok(());
        }
        let family = matches!(
            rt,
            FEAT_HDR11_RECORD_TYPE
                | FEATURE11_RECORD_TYPE
                | FEATURE12_RECORD_TYPE
                | CONTINUE_FRT11_RECORD_TYPE
                | LIST12_RECORD_TYPE
                | AUTO_FILTER12_RECORD_TYPE
                | crate::xls::sort_data::SORT_DATA_RECORD_TYPE
                | crate::xls::sort_data::CONTINUE_FRT12_RECORD_TYPE
        );
        if !family {
            if self.header.is_some() {
                if self.sort_continuations != 0 {
                    return Err(invalid(rt, "incomplete table SortData continuation chain"));
                }
                self.materialize()?;
                self.finish_future()?;
                self.ended = true;
            }
            return Ok(());
        }
        if rt == FEAT_HDR11_RECORD_TYPE {
            // FeatHdr11 is shared by all Feat11 feature families. Leave non-list
            // discriminators to their dedicated collectors.
            if data.len() < 14 || u16_at(data, 12, rt, "isf")? != ISF_LIST {
                if self.header.is_some() {
                    self.ended = true;
                }
                return Ok(());
            }
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            if self.header.is_some() || data.len() != 29 {
                return Err(invalid(rt, "duplicate or malformed FeatHdr11"));
            }
            validate_frt(data, rt, false)?;
            if u16_at(data, 12, rt, "isf")? != ISF_LIST
                || data[14] != 1
                || u32_at(data, 15, rt, "reserved2")? != u32::MAX
                || u32_at(data, 19, rt, "reserved3")? != u32::MAX
                || u16_at(data, 27, rt, "reserved4")? != 0
            {
                return Err(invalid(rt, "invalid FeatHdr11 fields"));
            }
            self.header = Some(u32_at(data, 23, rt, "idListNext")?);
        } else if matches!(rt, FEATURE11_RECORD_TYPE | FEATURE12_RECORD_TYPE) {
            if data.len() < 14 || u16_at(data, 12, rt, "isf")? != ISF_LIST {
                if self.header.is_some() {
                    self.ended = true;
                }
                return Ok(());
            }
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            if self.header.is_none() {
                return Err(invalid(rt, "table feature without FeatHdr11"));
            }
            if self.sort_continuations != 0 {
                return Err(invalid(
                    rt,
                    "table feature interrupts SortData continuation chain",
                ));
            }
            self.materialize()?;
            self.finish_future()?;
            self.flush()?;
            if data.len() > MAX_PAYLOAD {
                return Err(invalid(rt, "base table feature exceeds BIFF record limit"));
            }
            self.pending = Some(PendingFeature {
                record_type: rt,
                base: data.to_vec(),
                continuations: Vec::new(),
                combined: data.to_vec(),
            });
            self.kinds.clear();
            self.list12_count = 0;
        } else if rt == CONTINUE_FRT11_RECORD_TYPE {
            let pending = self
                .pending
                .as_mut()
                .ok_or_else(|| invalid(rt, "orphan ContinueFrt11"))?;
            if pending.base.len() != MAX_PAYLOAD
                || pending
                    .continuations
                    .last()
                    .is_some_and(|v| v.len() != MAX_PAYLOAD)
            {
                return Err(invalid(
                    rt,
                    "ContinueFrt11 follows a non-full feature fragment",
                ));
            }
            if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                return Err(invalid(rt, "invalid ContinueFrt11 length"));
            }
            validate_frt(data, rt, false)?;
            pending.combined.extend_from_slice(&data[12..]);
            if pending.combined.len() > MAX_FEATURE_BYTES {
                return Err(invalid(
                    rt,
                    "table feature continuation chain exceeds resource bound",
                ));
            }
            pending.continuations.push(data.to_vec());
        } else if rt == LIST12_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            let kind = self
                .current
                .as_mut()
                .ok_or_else(|| invalid(rt, "List12 without Feature11"))?
                .apply_list12(data)?;
            if !self.kinds.insert(kind) {
                return Err(invalid(rt, "duplicate List12 type"));
            }
            self.list12_count += 1;
        } else if rt == AUTO_FILTER12_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            if self.current.is_none() || self.list12_count == 0 {
                return Err(invalid(
                    rt,
                    "AutoFilter12 is not attached after a table List12",
                ));
            }
            if self.current.as_ref().is_some_and(|table| {
                table.autofilter12_criteria.is_some()
                    || table
                        .opaque_future_records
                        .iter()
                        .any(|future| future.record_type == AUTO_FILTER12_RECORD_TYPE)
            }) {
                return Err(invalid(rt, "duplicate AutoFilter12 for table"));
            }
            if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                return Err(invalid(rt, "invalid AutoFilter12 length"));
            }
            validate_frt_any(data, rt)?;
            self.pending_future = Some(PendingFuture {
                payload: data.to_vec(),
                continuations: Vec::new(),
                after_list12_count: self.list12_count,
            });
        } else if rt == crate::xls::sort_data::SORT_DATA_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            let table = self
                .current
                .as_ref()
                .ok_or_else(|| invalid(rt, "SortData without table feature"))?;
            if data.len() != 38
                || ((u16_at(data, 12, rt, "sort flags")? >> 3) & 0x7) != 1
                || u32_at(data, 34, rt, "sort parent id")? != table.id.value()
            {
                return Err(invalid(
                    rt,
                    "table SortData parent does not match Feature11/12",
                ));
            }
            self.sort_continuations = u32_at(data, 30, rt, "sort condition count")? as usize;
        } else {
            if let Some(future) = self.pending_future.as_mut() {
                if future.payload.len() < 60 {
                    return Err(invalid(
                        rt,
                        "ContinueFrt12 follows a truncated AutoFilter12 base",
                    ));
                }
                if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                    return Err(invalid(rt, "invalid ContinueFrt12 length"));
                }
                validate_frt_any(data, rt)?;
                let total = future.payload.len()
                    + future.continuations.iter().map(Vec::len).sum::<usize>()
                    + data.len();
                if total > MAX_FEATURE_BYTES {
                    return Err(invalid(
                        rt,
                        "AutoFilter12 continuation chain exceeds resource bound",
                    ));
                }
                future.continuations.push(data.to_vec());
            } else if self.sort_continuations != 0 {
                self.sort_continuations -= 1;
            } else {
                return Err(invalid(rt, "orphan ContinueFrt12 in table feature family"));
            }
        }
        Ok(())
    }
    fn materialize(&mut self) -> XlsResult<()> {
        let Some(pending) = self.pending.take() else {
            return Ok(());
        };
        let source_type = u32_at(&pending.combined, 35, pending.record_type, "lt")?;
        self.current = Some(
            if pending.record_type == FEATURE12_RECORD_TYPE && source_type == 3 {
                XlsListObject::parse_external_feature12(pending)?
            } else if pending.record_type == FEATURE12_RECORD_TYPE && source_type > 3 {
                XlsListObject::parse_opaque_feature12(pending)?
            } else {
                XlsListObject::parse_feature(&pending.combined, pending.record_type)?
            },
        );
        Ok(())
    }
    fn finish_future(&mut self) -> XlsResult<()> {
        if let Some(future) = self.pending_future.take() {
            let table = self
                .current
                .as_mut()
                .ok_or_else(|| invalid(AUTO_FILTER12_RECORD_TYPE, "detached AutoFilter12"))?;
            if let Some(filter) = parse_table_autofilter12(
                &future.payload,
                &future.continuations,
                table.range,
                table.id,
            )? {
                table.autofilter12_criteria = Some(filter);
            } else {
                table
                    .opaque_future_records
                    .push(XlsOpaqueListObjectFutureRecord {
                        record_type: AUTO_FILTER12_RECORD_TYPE,
                        payload: future.payload,
                        continuation_payloads: future.continuations,
                        after_list12_count: future.after_list12_count,
                    });
            }
        }
        Ok(())
    }
    fn flush(&mut self) -> XlsResult<()> {
        if let Some(table) = self.current.take() {
            table.validate()?;
            self.tables.push(table);
        }
        Ok(())
    }
    pub(crate) fn finish(mut self) -> XlsResult<Vec<XlsListObject>> {
        if self.sort_continuations != 0 {
            return Err(invalid(
                crate::xls::sort_data::SORT_DATA_RECORD_TYPE,
                "incomplete table SortData continuation chain",
            ));
        }
        self.materialize()?;
        self.finish_future()?;
        self.flush()?;
        let mut ids = HashSet::new();
        let mut names = HashSet::new();
        for table in &self.tables {
            if !ids.insert(table.id) || !names.insert(table.name.to_lowercase()) {
                return Err(invalid(FEATURE11_RECORD_TYPE, "duplicate table id or name"));
            }
        }
        Ok(self.tables)
    }
}
pub(crate) fn feature_header_record(tables: &[XlsListObject]) -> XlsResult<Vec<u8>> {
    let next = tables
        .iter()
        .map(|t| t.id.value())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| invalid(FEAT_HDR11_RECORD_TYPE, "next table id overflows"))?;
    let mut p = Vec::new();
    append_frt(&mut p, FEAT_HDR11_RECORD_TYPE, None);
    p.extend_from_slice(&ISF_LIST.to_le_bytes());
    p.push(1);
    p.extend_from_slice(&u32::MAX.to_le_bytes());
    p.extend_from_slice(&u32::MAX.to_le_bytes());
    p.extend_from_slice(&next.to_le_bytes());
    p.extend_from_slice(&0u16.to_le_bytes());
    record(FEAT_HDR11_RECORD_TYPE, p)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(column_count: usize, name_len: usize) -> XlsListObject {
        let columns = (0..column_count)
            .map(|index| {
                XlsListObjectColumn::try_new(
                    XlsListColumnId::try_new(index as u32 + 1).unwrap(),
                    format!("C{index}_{}", "x".repeat(name_len)),
                )
                .unwrap()
            })
            .collect();
        XlsListObject::try_new(
            XlsListObjectId::try_new(1).unwrap(),
            "TableOne",
            XlsListObjectRange::try_new(0, 2, 0, column_count as u16 - 1).unwrap(),
            columns,
            XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
        )
        .unwrap()
    }

    fn payload(record: &[u8]) -> &[u8] {
        &record[4..]
    }

    fn parse_feature_records(
        table: &XlsListObject,
        records: &[Vec<u8>],
    ) -> XlsResult<XlsListObject> {
        let mut collector = ListObjectCollector::new();
        let header = feature_header_record(std::slice::from_ref(table))?;
        collector.feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))?;
        for record in records {
            let record_type = u16::from_le_bytes(record[..2].try_into().unwrap());
            collector.feed_record(record_type, payload(record))?;
        }
        for record in table.to_list12_record_bytes()? {
            collector.feed_record(LIST12_RECORD_TYPE, payload(&record))?;
        }
        Ok(collector.finish()?.remove(0))
    }

    #[test]
    fn continue_frt11_rejects_orphans_bad_echoes_and_short_predecessors() {
        let continuation = {
            let mut value = Vec::new();
            append_frt(&mut value, CONTINUE_FRT11_RECORD_TYPE, None);
            value
        };
        assert!(
            ListObjectCollector::new()
                .feed_record(CONTINUE_FRT11_RECORD_TYPE, &continuation)
                .is_err()
        );

        let short = table(2, 3);
        let mut collector = ListObjectCollector::new();
        let header = feature_header_record(std::slice::from_ref(&short)).unwrap();
        collector
            .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
            .unwrap();
        let feature = short.to_feature_record_bytes().unwrap();
        collector
            .feed_record(FEATURE11_RECORD_TYPE, payload(&feature[0]))
            .unwrap();
        assert!(
            collector
                .feed_record(CONTINUE_FRT11_RECORD_TYPE, &continuation)
                .is_err()
        );

        let long = table(256, 220);
        let mut records = long.to_feature_record_bytes().unwrap();
        assert!(records.len() > 2);
        records[1][4] = 0;
        let mut collector = ListObjectCollector::new();
        let header = feature_header_record(std::slice::from_ref(&long)).unwrap();
        collector
            .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
            .unwrap();
        collector
            .feed_record(FEATURE11_RECORD_TYPE, payload(&records[0]))
            .unwrap();
        assert!(
            collector
                .feed_record(CONTINUE_FRT11_RECORD_TYPE, payload(&records[1]))
                .is_err()
        );
    }

    #[test]
    fn unsupported_feature12_bytes_are_retained_and_autofilter12_chain_is_strict() {
        let value = table(2, 3).with_header_row(false).unwrap();
        let header = feature_header_record(std::slice::from_ref(&value)).unwrap();
        let mut feature = value.to_feature_record_bytes().unwrap().remove(0);
        feature[4 + 35..4 + 39].copy_from_slice(&4u32.to_le_bytes());
        let list12 = value.to_list12_record_bytes().unwrap();
        let mut collector = ListObjectCollector::new();
        collector
            .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
            .unwrap();
        collector
            .feed_record(FEATURE12_RECORD_TYPE, payload(&feature))
            .unwrap();
        collector
            .feed_record(LIST12_RECORD_TYPE, payload(&list12[0]))
            .unwrap();
        let mut autofilter = Vec::new();
        append_frt(&mut autofilter, AUTO_FILTER12_RECORD_TYPE, None);
        collector
            .feed_record(AUTO_FILTER12_RECORD_TYPE, &autofilter)
            .unwrap();
        let mut continuation = Vec::new();
        append_frt(
            &mut continuation,
            crate::xls::sort_data::CONTINUE_FRT12_RECORD_TYPE,
            None,
        );
        assert!(
            collector
                .feed_record(
                    crate::xls::sort_data::CONTINUE_FRT12_RECORD_TYPE,
                    &continuation
                )
                .is_err()
        );

        let mut collector = ListObjectCollector::new();
        collector
            .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
            .unwrap();
        collector
            .feed_record(FEATURE12_RECORD_TYPE, payload(&feature))
            .unwrap();
        for record in list12 {
            collector
                .feed_record(LIST12_RECORD_TYPE, payload(&record))
                .unwrap();
        }
        let parsed = collector.finish().unwrap().remove(0);
        assert!(parsed.opaque_feature().is_some());
        assert_eq!(parsed.to_feature_record_bytes().unwrap()[0], feature);
    }

    #[test]
    fn external_feature12_is_lossless_and_hostile_versions_cardinality_and_feature11_are_rejected()
    {
        let base_table = table(2, 3);
        let metadata = XlsExternalTableMetadata::try_new(vec![
            XlsExternalTableField::try_new(base_table.columns[0].id, "SOURCE_A", 41).unwrap(),
            XlsExternalTableField::try_new(base_table.columns[1].id, "SOURCE_B", 42).unwrap(),
        ])
        .unwrap();
        let value = base_table.with_external_data(metadata).unwrap();
        let header = feature_header_record(std::slice::from_ref(&value)).unwrap();
        let feature = value.to_feature_record_bytes().unwrap();
        let list12 = value.to_list12_record_bytes().unwrap();
        let parse = |records: &[Vec<u8>]| -> XlsResult<Vec<XlsListObject>> {
            let mut collector = ListObjectCollector::new();
            collector.feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))?;
            collector.feed_record(FEATURE12_RECORD_TYPE, payload(&records[0]))?;
            for continuation in &records[1..] {
                collector.feed_record(CONTINUE_FRT11_RECORD_TYPE, payload(continuation))?;
            }
            for record in &list12 {
                collector.feed_record(LIST12_RECORD_TYPE, payload(record))?;
            }
            collector.finish()
        };
        let parsed = parse(&feature).unwrap().remove(0);
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), feature);
        assert_eq!(
            parsed.external_metadata().unwrap().fields()[1].query_field_id(),
            42
        );

        let mut bad_version = feature.clone();
        let flags = u32::from_le_bytes(bad_version[0][4 + 63..4 + 67].try_into().unwrap());
        bad_version[0][4 + 63..4 + 67]
            .copy_from_slice(&((flags & !0x000F_0000) | 0x000A_0000).to_le_bytes());
        assert!(parse(&bad_version).is_err());

        let mut bad_count = feature.clone();
        let (_, count_offset) =
            parse_string(payload(&bad_count[0]), 99, FEATURE12_RECORD_TYPE, "rgbName").unwrap();
        bad_count[0][4 + count_offset..4 + count_offset + 2].copy_from_slice(&3u16.to_le_bytes());
        assert!(parse(&bad_count).is_err());

        let ordinary = table(2, 3);
        let ordinary_header = feature_header_record(std::slice::from_ref(&ordinary)).unwrap();
        let mut feature11 = ordinary.to_feature_record_bytes().unwrap().remove(0);
        feature11[4 + 35..4 + 39].copy_from_slice(&3u32.to_le_bytes());
        let mut collector = ListObjectCollector::new();
        collector
            .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&ordinary_header))
            .unwrap();
        collector
            .feed_record(FEATURE11_RECORD_TYPE, payload(&feature11))
            .unwrap();
        assert!(collector.finish().is_err());
    }

    #[test]
    fn feature11_web_lfdt_values_and_defaults_round_trip_strictly() {
        let base = table(XlsWebColumnType::ALL.len(), 1);
        let fields = XlsWebColumnType::ALL
            .iter()
            .copied()
            .enumerate()
            .map(|(index, kind)| {
                let info = match kind {
                    XlsWebColumnType::Text
                    | XlsWebColumnType::Choice
                    | XlsWebColumnType::MultipleChoices => XlsWebFieldInfo::new(1033)
                        .with_default_value(XlsWebDefaultValue::String(format!("v{index}"))),
                    XlsWebColumnType::Boolean => XlsWebFieldInfo::new(1033)
                        .with_default_value(XlsWebDefaultValue::Boolean(true)),
                    XlsWebColumnType::Number | XlsWebColumnType::Currency => {
                        XlsWebFieldInfo::new(1033)
                            .with_default_value(XlsWebDefaultValue::Number(12.5))
                    },
                    XlsWebColumnType::DateTime => XlsWebFieldInfo::new(1033)
                        .with_default_value(XlsWebDefaultValue::DateTime(45_000.25)),
                    _ => XlsWebFieldInfo::new(1033),
                };
                XlsWebTableField::try_new(
                    base.columns[index].id,
                    format!("SOURCE_{index}"),
                    kind,
                    info,
                )
                .unwrap()
            })
            .collect();
        let value = base
            .with_web_source(XlsWebTableMetadata::try_new(fields).unwrap())
            .unwrap();
        assert_eq!(
            value.feature_version(),
            XlsListObjectFeatureVersion::Feature11
        );
        let records = value.to_feature_record_bytes().unwrap();
        let parsed = parse_feature_records(&value, &records).unwrap();
        let XlsListObjectSourceMetadata::Web(metadata) = parsed.source_metadata().unwrap() else {
            panic!("expected Web metadata")
        };
        assert_eq!(
            metadata
                .fields()
                .iter()
                .map(XlsWebTableField::data_type)
                .collect::<Vec<_>>(),
            XlsWebColumnType::ALL
        );
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

        assert!(
            XlsWebTableField::try_new(
                value.columns[0].id,
                "INVALID_DEFAULT",
                XlsWebColumnType::Note,
                XlsWebFieldInfo::new(0)
                    .with_default_value(XlsWebDefaultValue::String("x".to_string())),
            )
            .is_err()
        );
    }

    #[test]
    fn feature11_xml_lfxidt_is_exhaustive_and_preserves_ignored_storage() {
        let base = table(XlsXmlDataType::ALL.len(), 1);
        let fields = XlsXmlDataType::ALL
            .iter()
            .copied()
            .enumerate()
            .map(|(index, kind)| {
                XlsXmlTableField::try_new(base.columns[index].id, format!("XML_{index}"), kind)
                    .unwrap()
            })
            .collect();
        let value = base
            .with_xml_source(XlsXmlTableMetadata::try_new(fields).unwrap())
            .unwrap();
        let mut records = value.to_feature_record_bytes().unwrap();
        assert_eq!(records.len(), 1);
        let (_, count_offset) =
            parse_string(payload(&records[0]), 99, FEATURE11_RECORD_TYPE, "rgbName").unwrap();
        let first_field = count_offset + 2;
        records[0][4 + 35 + 26..4 + 35 + 28].copy_from_slice(&0xa55au16.to_le_bytes());
        let flags = u32::from_le_bytes(records[0][4 + 63..4 + 67].try_into().unwrap());
        records[0][4 + 63..4 + 67].copy_from_slice(&(flags | 0x8280_0001).to_le_bytes());
        records[0][4 + 35 + 32] = 0x5a;
        records[0][4 + 35 + 48] = 0xa5;
        let field_flags = u32::from_le_bytes(
            records[0][4 + first_field + 24..4 + first_field + 28]
                .try_into()
                .unwrap(),
        );
        records[0][4 + first_field + 24..4 + first_field + 28]
            .copy_from_slice(&(field_flags | 0x8000_0030).to_le_bytes());

        let parsed = parse_feature_records(&value, &records).unwrap();
        let XlsListObjectSourceMetadata::Xml(metadata) = parsed.source_metadata().unwrap() else {
            panic!("expected XML metadata")
        };
        assert_eq!(metadata.ignored_fixed_word(), 0xa55a);
        assert_eq!(metadata.ignored_flags(), 0x8280_0001);
        assert_eq!(metadata.ignored_fixed_tail()[0], 0x5a);
        assert_eq!(metadata.ignored_fixed_tail()[16], 0xa5);
        assert_eq!(metadata.fields()[0].ignored_flags(), 0x8000_0030);
        assert_eq!(
            metadata
                .fields()
                .iter()
                .map(XlsXmlTableField::data_type)
                .collect::<Vec<_>>(),
            XlsXmlDataType::ALL
        );
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

        let mut invalid = records.clone();
        invalid[0][4 + first_field + 8..4 + first_field + 12]
            .copy_from_slice(&0x212eu32.to_le_bytes());
        assert!(parse_feature_records(&value, &invalid).is_err());
        let mut wrong_source = records.clone();
        wrong_source[0][4 + first_field + 4..4 + first_field + 8]
            .copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_feature_records(&value, &wrong_source).is_err());
    }

    #[test]
    fn feature11_web_ignored_flags_round_trip_but_reserved_and_source_bits_fail() {
        let base = table(1, 1);
        let field = XlsWebTableField::try_new(
            base.columns[0].id,
            "SOURCE",
            XlsWebColumnType::Text,
            XlsWebFieldInfo::new(1033)
                .with_default_value(XlsWebDefaultValue::String("default".to_string())),
        )
        .unwrap();
        let value = base
            .with_web_source(XlsWebTableMetadata::try_new(vec![field]).unwrap())
            .unwrap();
        let canonical = value.to_feature_record_bytes().unwrap();
        let (_, count_offset) =
            parse_string(payload(&canonical[0]), 99, FEATURE11_RECORD_TYPE, "rgbName").unwrap();
        let field_offset = count_offset + 2;
        let (_, after_source) = parse_string(
            payload(&canonical[0]),
            field_offset + 36,
            FEATURE11_RECORD_TYPE,
            "source",
        )
        .unwrap();
        let (_, after_caption) = parse_string(
            payload(&canonical[0]),
            after_source,
            FEATURE11_RECORD_TYPE,
            "caption",
        )
        .unwrap();
        let web_info_offset = after_caption + 6;

        let mut ignored = canonical.clone();
        let flags = u32::from_le_bytes(ignored[0][4 + 63..4 + 67].try_into().unwrap());
        ignored[0][4 + 63..4 + 67].copy_from_slice(&(flags | 0x8280_0001).to_le_bytes());
        let field_flags = u32::from_le_bytes(
            ignored[0][4 + field_offset + 24..4 + field_offset + 28]
                .try_into()
                .unwrap(),
        );
        ignored[0][4 + field_offset + 24..4 + field_offset + 28]
            .copy_from_slice(&(field_flags | 0x8000_0030).to_le_bytes());
        let display = u32::from_le_bytes(
            ignored[0][4 + web_info_offset + 8..4 + web_info_offset + 12]
                .try_into()
                .unwrap(),
        );
        ignored[0][4 + web_info_offset + 8..4 + web_info_offset + 12]
            .copy_from_slice(&(display | 0x8000_0000).to_le_bytes());
        let validation = u32::from_le_bytes(
            ignored[0][4 + web_info_offset + 12..4 + web_info_offset + 16]
                .try_into()
                .unwrap(),
        );
        ignored[0][4 + web_info_offset + 12..4 + web_info_offset + 16]
            .copy_from_slice(&(validation | 0x4000_0000).to_le_bytes());
        let parsed = parse_feature_records(&value, &ignored).unwrap();
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), ignored);

        let mut bad_lfdt = canonical.clone();
        bad_lfdt[0][4 + field_offset + 4..4 + field_offset + 8]
            .copy_from_slice(&12u32.to_le_bytes());
        assert!(parse_feature_records(&value, &bad_lfdt).is_err());
        let mut bad_xml_type = canonical.clone();
        bad_xml_type[0][4 + field_offset + 8..4 + field_offset + 12]
            .copy_from_slice(&XlsXmlDataType::DataTypeString.value().to_le_bytes());
        assert!(parse_feature_records(&value, &bad_xml_type).is_err());
        let mut reserved = canonical.clone();
        let field_flags = u32::from_le_bytes(
            reserved[0][4 + field_offset + 24..4 + field_offset + 28]
                .try_into()
                .unwrap(),
        );
        reserved[0][4 + field_offset + 24..4 + field_offset + 28]
            .copy_from_slice(&(field_flags | 0x40).to_le_bytes());
        assert!(parse_feature_records(&value, &reserved).is_err());

        let mut unsupported_feature12 = canonical.clone();
        unsupported_feature12[0][..2].copy_from_slice(&FEATURE12_RECORD_TYPE.to_le_bytes());
        unsupported_feature12[0][4..6].copy_from_slice(&FEATURE12_RECORD_TYPE.to_le_bytes());
        assert!(parse_feature_records(&value, &unsupported_feature12).is_err());
    }

    #[test]
    fn feature12_single_cell_xml_source_round_trips() {
        let mut base = table(1, 1);
        base.range = XlsListObjectRange::try_new(0, 0, 0, 0).unwrap();
        let base = base.with_header_row(false).unwrap();
        let field =
            XlsXmlTableField::try_new(base.columns[0].id, "single", XlsXmlDataType::DataTypeString)
                .unwrap();
        let metadata = XlsXmlTableMetadata::try_new(vec![field])
            .unwrap()
            .with_single_cell(true)
            .unwrap();
        let value = base.with_xml_source(metadata).unwrap();
        let records = value.to_feature_record_bytes().unwrap();
        assert_eq!(
            u16::from_le_bytes(records[0][..2].try_into().unwrap()),
            FEATURE12_RECORD_TYPE
        );
        let parsed = parse_feature_records(&value, &records).unwrap();
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);
    }

    #[test]
    fn cached_disk_header_typed_builder_and_noncanonical_string_round_trip() {
        let mut raw = Vec::new();
        raw.extend_from_slice(&3u32.to_le_bytes());
        raw.extend_from_slice(&[0xaa, 0xbb, 0xcc]);
        raw.extend_from_slice(&6u16.to_le_bytes());
        raw.push(1);
        raw.extend("Header".encode_utf16().flat_map(u16::to_le_bytes));
        let base = table(1, 1).with_header_row(false).unwrap();
        let field = XlsExternalTableField::try_new(base.columns[0].id, "SOURCE", 7)
            .unwrap()
            .with_header_cache_bytes(raw.clone())
            .unwrap();
        assert_eq!(
            field.cached_disk_header().formatting_bytes(),
            &[0xaa, 0xbb, 0xcc]
        );
        assert_eq!(field.cached_disk_header().style_name(), Some("Header"));
        assert_eq!(field.header_cache_bytes(), raw);
        let value = base
            .clone()
            .with_external_data(XlsExternalTableMetadata::try_new(vec![field]).unwrap())
            .unwrap();
        let records = value.to_feature_record_bytes().unwrap();
        let parsed = parse_feature_records(&value, &records).unwrap();
        let field = &parsed.external_metadata().unwrap().fields()[0];
        assert_eq!(field.header_cache_bytes(), raw);
        assert_eq!(field.cached_disk_header().style_name(), Some("Header"));
        assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

        let built = XlsCachedDiskHeader::try_new(vec![1, 2])
            .unwrap()
            .with_style_name("BuiltInHeader")
            .unwrap();
        assert_eq!(built.formatting_bytes(), &[1, 2]);
        assert_eq!(built.style_name(), Some("BuiltInHeader"));
        assert_eq!(built.clone().without_style_name().style_name(), None);
    }

    #[test]
    fn cached_disk_header_presence_lengths_and_flags_are_strict() {
        assert!(XlsCachedDiskHeader::try_new(vec![0; MAX_FEATURE_BYTES]).is_err());
        assert!(
            XlsExternalTableField::try_new(XlsListColumnId::try_new(1).unwrap(), "SOURCE", 1,)
                .unwrap()
                .with_header_cache_bytes(vec![2, 0, 0, 0, 1])
                .is_err()
        );

        let header = XlsCachedDiskHeader::try_new(vec![1])
            .unwrap()
            .with_style_name("HeaderStyle")
            .unwrap();
        let headered = table(1, 1);
        let field = XlsExternalTableField::try_new(headered.columns[0].id, "SOURCE", 1)
            .unwrap()
            .with_cached_disk_header(header)
            .unwrap();
        assert!(
            headered
                .with_external_data(XlsExternalTableMetadata::try_new(vec![field]).unwrap())
                .is_err()
        );

        let base = table(1, 1).with_header_row(false).unwrap();
        let field = XlsExternalTableField::try_new(base.columns[0].id, "SOURCE", 1)
            .unwrap()
            .with_cached_disk_header(
                XlsCachedDiskHeader::try_new(vec![0x10])
                    .unwrap()
                    .with_style_name("HeaderStyle")
                    .unwrap(),
            )
            .unwrap();
        let value = base
            .clone()
            .with_external_data(XlsExternalTableMetadata::try_new(vec![field]).unwrap())
            .unwrap();
        let records = value.to_feature_record_bytes().unwrap();
        let (_, count_offset) =
            parse_string(payload(&records[0]), 99, FEATURE12_RECORD_TYPE, "rgbName").unwrap();
        let mut field_offset = count_offset + 2;
        let table_flags = u32::from_le_bytes(records[0][4 + 63..4 + 67].try_into().unwrap());
        if table_flags & 0x0010_0000 != 0 {
            field_offset = parse_string(
                payload(&records[0]),
                field_offset,
                FEATURE12_RECORD_TYPE,
                "entryId",
            )
            .unwrap()
            .1;
        }

        let mut missing_flag = records.clone();
        let flags = u32::from_le_bytes(
            missing_flag[0][4 + field_offset + 24..4 + field_offset + 28]
                .try_into()
                .unwrap(),
        );
        missing_flag[0][4 + field_offset + 24..4 + field_offset + 28]
            .copy_from_slice(&(flags & !0x200).to_le_bytes());
        assert!(parse_feature_records(&value, &missing_flag).is_err());

        let empty_field = XlsExternalTableField::try_new(base.columns[0].id, "SOURCE", 1).unwrap();
        let empty_value = base
            .with_external_data(XlsExternalTableMetadata::try_new(vec![empty_field]).unwrap())
            .unwrap();
        let mut spurious_flag = empty_value.to_feature_record_bytes().unwrap();
        let (_, count_offset) = parse_string(
            payload(&spurious_flag[0]),
            99,
            FEATURE12_RECORD_TYPE,
            "rgbName",
        )
        .unwrap();
        let mut field_offset = count_offset + 2;
        let table_flags = u32::from_le_bytes(spurious_flag[0][4 + 63..4 + 67].try_into().unwrap());
        if table_flags & 0x0010_0000 != 0 {
            field_offset = parse_string(
                payload(&spurious_flag[0]),
                field_offset,
                FEATURE12_RECORD_TYPE,
                "entryId",
            )
            .unwrap()
            .1;
        }
        let flags = u32::from_le_bytes(
            spurious_flag[0][4 + field_offset + 24..4 + field_offset + 28]
                .try_into()
                .unwrap(),
        );
        spurious_flag[0][4 + field_offset + 24..4 + field_offset + 28]
            .copy_from_slice(&(flags | 0x200).to_le_bytes());
        assert!(parse_feature_records(&empty_value, &spurious_flag).is_err());
    }
}
