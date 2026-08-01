//! Pivot table record parsing for XLS BIFF8 files.
//!
//! Parses the family of SX* records that define pivot table structures:
//!
//! - **SXVIEW** (0x00B0): View definition — the main pivot table header.
//! - **SXVD** (0x00B1): View field — describes a single field (dimension).
//! - **SXVI** (0x00B2): View item — a single item within a field.
//! - **SXDI** (0x00C5): Data item — describes a data field (value area).
//! - **SXVS** (0x00E3): View source — source type of the pivot cache.
//! - **SXPI** (0x00B6): Page item — page field entries.
//!
//! # References
//!
//! - MS-XLS sections 2.4.271–2.4.283
//! - Apache POI `org.apache.poi.hssf.record.pivottable.*`

use crate::error::{XlsError, XlsResult};
use litchi_core::binary;

/// A BIFF8 cell error stored in an `SXERROR` PivotCache item.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum PivotCacheError {
    Null = 0x00,
    DivisionByZero = 0x07,
    Value = 0x0F,
    Reference = 0x17,
    Name = 0x1D,
    Number = 0x24,
    NotAvailable = 0x2A,
}

impl PivotCacheError {
    pub const fn code(self) -> u16 {
        self as u16
    }
}

impl TryFrom<u16> for PivotCacheError {
    type Error = XlsError;

    fn try_from(value: u16) -> Result<Self, Self::Error> {
        match value {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2A => Ok(Self::NotAvailable),
            _ => Err(cache_invalid(
                0x00CB,
                format!("invalid BIFF error code 0x{value:04X}"),
            )),
        }
    }
}

/// Lossless calendar value stored by `SXDATETIME`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PivotCacheDateTime {
    year: u16,
    month: u16,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PivotCacheDateGroupUnit {
    Seconds = 1,
    Minutes = 2,
    Hours = 3,
    Days = 4,
    Months = 5,
    Quarters = 6,
    Years = 7,
}

impl TryFrom<u16> for PivotCacheDateGroupUnit {
    type Error = XlsError;
    fn try_from(value: u16) -> Result<Self, Self::Error> {
        match value {
            1 => Ok(Self::Seconds),
            2 => Ok(Self::Minutes),
            3 => Ok(Self::Hours),
            4 => Ok(Self::Days),
            5 => Ok(Self::Months),
            6 => Ok(Self::Quarters),
            7 => Ok(Self::Years),
            _ => Err(cache_invalid(
                0x00D8,
                format!("unsupported date grouping unit {value}"),
            )),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheNumericGrouping {
    pub start: f64,
    pub end: f64,
    pub step: f64,
    pub auto_start: bool,
    pub auto_end: bool,
    pub group_items: Vec<PivotCacheItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheDateGrouping {
    pub unit: PivotCacheDateGroupUnit,
    pub start: PivotCacheDateTime,
    pub end: PivotCacheDateTime,
    pub step: u16,
    pub auto_start: bool,
    pub auto_end: bool,
    pub group_items: Vec<PivotCacheItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheDiscreteGrouping {
    pub base_field_index: u16,
    pub group_items: Vec<PivotCacheItem>,
    /// One group-item ordinal for every original item in the base field.
    pub item_to_group: Vec<u16>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheGrouping {
    Numeric(PivotCacheNumericGrouping),
    Date(PivotCacheDateGrouping),
    Discrete(PivotCacheDiscreteGrouping),
}

impl PivotCacheGrouping {
    pub(crate) fn group_items(&self) -> &[PivotCacheItem] {
        match self {
            Self::Numeric(value) => &value.group_items,
            Self::Date(value) => &value.group_items,
            Self::Discrete(value) => &value.group_items,
        }
    }
}

impl PivotCacheDateTime {
    pub fn try_new(
        year: u16,
        month: u16,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
    ) -> XlsResult<Self> {
        let value = Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
        };
        value.validate()?;
        Ok(value)
    }

    fn validate(self) -> XlsResult<()> {
        let leap = self.year.is_multiple_of(4)
            && (!self.year.is_multiple_of(100) || self.year.is_multiple_of(400));
        let max_day = match self.month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 if leap || self.year == 1900 => 29,
            2 => 28,
            _ => 0,
        };
        let sentinel = self.year == 1900
            && self.month == 1
            && self.day == 0
            && self.hour == 0
            && self.minute == 0
            && self.second == 0;
        if !(1900..=9999).contains(&self.year)
            || max_day == 0
            || (!sentinel && !(1..=max_day).contains(&self.day))
            || self.hour > 23
            || self.minute > 59
            || self.second > 59
        {
            return Err(cache_invalid(
                0x00CE,
                format!(
                    "invalid PivotCache date/time {:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                    self.year, self.month, self.day, self.hour, self.minute, self.second
                ),
            ));
        }
        Ok(())
    }

    pub const fn year(self) -> u16 {
        self.year
    }
    pub const fn month(self) -> u16 {
        self.month
    }
    pub const fn day(self) -> u8 {
        self.day
    }
    pub const fn hour(self) -> u8 {
        self.hour
    }
    pub const fn minute(self) -> u8 {
        self.minute
    }
    pub const fn second(self) -> u8 {
        self.second
    }
}

/// A typed PivotCache shared or resolved row item.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheItem {
    String(String),
    Number(f64),
    Boolean(bool),
    Error(PivotCacheError),
    DateTime(PivotCacheDateTime),
    Empty,
}

impl From<String> for PivotCacheItem {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<&str> for PivotCacheItem {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl PivotCacheItem {
    pub(crate) fn validate(&self) -> XlsResult<()> {
        match self {
            Self::String(value) if value.is_empty() => Err(cache_invalid(
                0x00CD,
                "empty strings must use PivotCacheItem::Empty",
            )),
            Self::Number(value) if !value.is_finite() => Err(cache_invalid(
                0x00C9,
                "PivotCache numeric items must be finite",
            )),
            Self::DateTime(value) => value.validate(),
            _ => Ok(()),
        }
    }

    pub(crate) fn display_text(&self) -> String {
        match self {
            Self::String(value) => value.clone(),
            Self::Number(value) => value.to_string(),
            Self::Boolean(value) => if *value { "TRUE" } else { "FALSE" }.to_string(),
            Self::Error(value) => format!("#{value:?}"),
            Self::DateTime(value) => format!(
                "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                value.year, value.month, value.day, value.hour, value.minute, value.second
            ),
            Self::Empty => String::new(),
        }
    }
}

pub(crate) fn pivot_cache_data_flags(items: &[PivotCacheItem]) -> u16 {
    let (mut text, mut integer, mut double, mut date, mut blank) =
        (false, false, false, false, false);
    for item in items {
        match item {
            PivotCacheItem::String(_) | PivotCacheItem::Boolean(_) | PivotCacheItem::Error(_) => {
                text = true
            },
            PivotCacheItem::Number(value) if value.fract() == 0.0 => integer = true,
            PivotCacheItem::Number(_) => double = true,
            PivotCacheItem::DateTime(_) => date = true,
            PivotCacheItem::Empty => blank = true,
        }
    }
    if date && blank && !text && !integer && !double {
        return 0x0980;
    }
    if blank {
        text = true;
    }
    const FLAGS: [u16; 16] = [
        0x0000, 0x0480, 0x0520, 0x05A0, 0x0560, 0x05E0, 0x0520, 0x05A0, 0x0900, 0x0D80, 0x0D00,
        0x0D80, 0x0D00, 0x0D80, 0x0D00, 0x0D80,
    ];
    FLAGS[usize::from(text)
        | (usize::from(integer) << 1)
        | (usize::from(double) << 2)
        | (usize::from(date) << 3)]
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheField {
    name: String,
    flags: u16,
    group_parent: Option<u16>,
    group_base: Option<u16>,
    items: Vec<PivotCacheItem>,
    grouping: Option<PivotCacheGrouping>,
}
impl PivotCacheField {
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn flags(&self) -> u16 {
        self.flags
    }
    pub fn items(&self) -> &[PivotCacheItem] {
        &self.items
    }
    pub const fn group_parent(&self) -> Option<u16> {
        self.group_parent
    }
    pub const fn group_base(&self) -> Option<u16> {
        self.group_base
    }
    pub fn grouping(&self) -> Option<&PivotCacheGrouping> {
        self.grouping.as_ref()
    }
    pub(crate) fn replace_grouping(&mut self, grouping: Option<PivotCacheGrouping>) {
        self.group_base = match &grouping {
            Some(PivotCacheGrouping::Discrete(value)) => Some(value.base_field_index),
            _ => None,
        };
        self.grouping = grouping;
    }
    pub(crate) fn set_group_parent(&mut self, parent: Option<u16>) {
        self.group_parent = parent;
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCache {
    stream_id: u16,
    flags: u16,
    record_count: u32,
    fields: Vec<PivotCacheField>,
    rows: Vec<Vec<PivotCacheItem>>,
}
impl PivotCache {
    pub const fn stream_id(&self) -> u16 {
        self.stream_id
    }
    pub const fn flags(&self) -> u16 {
        self.flags
    }
    pub const fn record_count(&self) -> u32 {
        self.record_count
    }
    pub fn fields(&self) -> &[PivotCacheField] {
        &self.fields
    }
    pub fn rows(&self) -> &[Vec<PivotCacheItem>] {
        &self.rows
    }
    pub(crate) fn fields_mut(&mut self) -> &mut [PivotCacheField] {
        &mut self.fields
    }
}

fn cache_invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn cache_records(data: &[u8]) -> XlsResult<Vec<(u16, &[u8])>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let header = data
            .get(offset..offset + 4)
            .ok_or(XlsError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(len))
            .ok_or_else(|| cache_invalid(kind, "PivotCache record length overflow"))?;
        let body = data.get(offset + 4..end).ok_or(XlsError::InvalidLength {
            expected: end,
            found: data.len(),
        })?;
        records.push((kind, body));
        offset = end;
    }
    Ok(records)
}

fn parse_cache_string(data: &[u8], record_type: u16) -> XlsResult<(String, usize)> {
    if data.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }
    let count = usize::from(u16::from_le_bytes([data[0], data[1]]));
    let wide = data[2] & 1 != 0;
    let byte_count = count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| cache_invalid(record_type, "PivotCache string length overflow"))?;
    let end = 3usize
        .checked_add(byte_count)
        .ok_or_else(|| cache_invalid(record_type, "PivotCache string length overflow"))?;
    let chars = data.get(3..end).ok_or(XlsError::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let value = if wide {
        let units = chars
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| cache_invalid(record_type, "invalid UTF-16 PivotCache string"))?
    } else {
        chars.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, end))
}

fn parse_cache_item(record_type: u16, data: &[u8]) -> XlsResult<PivotCacheItem> {
    let item = match record_type {
        0x00C9 => {
            if data.len() != 8 {
                return Err(XlsError::InvalidLength {
                    expected: 8,
                    found: data.len(),
                });
            }
            PivotCacheItem::Number(f64::from_le_bytes(data.try_into().unwrap()))
        },
        0x00CA => {
            if data.len() != 2 {
                return Err(XlsError::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            match u16::from_le_bytes(data.try_into().unwrap()) {
                0 => PivotCacheItem::Boolean(false),
                1 => PivotCacheItem::Boolean(true),
                value => {
                    return Err(cache_invalid(
                        record_type,
                        format!("invalid SXBOOLEAN value {value}"),
                    ));
                },
            }
        },
        0x00CB => {
            if data.len() != 2 {
                return Err(XlsError::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            PivotCacheItem::Error(PivotCacheError::try_from(u16::from_le_bytes(
                data.try_into().unwrap(),
            ))?)
        },
        0x00CC => {
            if data.len() != 2 {
                return Err(XlsError::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            PivotCacheItem::Number(f64::from(i16::from_le_bytes(data.try_into().unwrap())))
        },
        0x00CD => {
            let (value, used) = parse_cache_string(data, record_type)?;
            if used != data.len() {
                return Err(cache_invalid(record_type, "trailing SXSTRING payload"));
            }
            PivotCacheItem::String(value)
        },
        0x00CE => {
            if data.len() != 8 {
                return Err(XlsError::InvalidLength {
                    expected: 8,
                    found: data.len(),
                });
            }
            PivotCacheItem::DateTime(PivotCacheDateTime::try_new(
                u16::from_le_bytes([data[0], data[1]]),
                u16::from_le_bytes([data[2], data[3]]),
                data[4],
                data[5],
                data[6],
                data[7],
            )?)
        },
        0x00CF => {
            if !data.is_empty() {
                return Err(XlsError::InvalidLength {
                    expected: 0,
                    found: data.len(),
                });
            }
            PivotCacheItem::Empty
        },
        _ => {
            return Err(cache_invalid(
                record_type,
                "unexpected PivotCache item record",
            ));
        },
    };
    item.validate()?;
    Ok(item)
}

/// Parse one `_SX_DB_CUR/nnnn` PivotCache stream.
pub fn parse_pivot_cache_stream(data: &[u8]) -> XlsResult<PivotCache> {
    let records = cache_records(data)?;
    let (sxdb_type, sxdb) = records
        .first()
        .ok_or_else(|| cache_invalid(0x00C6, "empty PivotCache stream"))?;
    if *sxdb_type != 0x00C6 || sxdb.len() < 18 {
        return Err(cache_invalid(*sxdb_type, "PivotCache must start with SXDB"));
    }
    let record_count = u32::from_le_bytes(sxdb[0..4].try_into().unwrap());
    let stream_id = u16::from_le_bytes(sxdb[4..6].try_into().unwrap());
    let cache_flags = u16::from_le_bytes(sxdb[6..8].try_into().unwrap());
    let standard_field_count = usize::from(u16::from_le_bytes(sxdb[10..12].try_into().unwrap()));
    let field_count = usize::from(u16::from_le_bytes(sxdb[12..14].try_into().unwrap()));
    if standard_field_count > field_count {
        return Err(cache_invalid(
            0x00C6,
            "standard PivotCache field count exceeds total count",
        ));
    }
    if stream_id == 0 {
        return Err(cache_invalid(
            0x00C6,
            "PivotCache stream ID must be nonzero",
        ));
    }
    let mut position = 1usize;
    if records
        .get(position)
        .is_some_and(|(kind, _)| *kind == 0x0122)
    {
        position += 1;
    }
    let mut fields = Vec::with_capacity(field_count);
    for _ in 0..field_count {
        let (kind, body) = records
            .get(position)
            .ok_or_else(|| cache_invalid(0x00C7, "missing SXFDB"))?;
        if *kind != 0x00C7 || body.len() < 17 {
            return Err(cache_invalid(*kind, "expected valid SXFDB"));
        }
        let flags = u16::from_le_bytes(body[0..2].try_into().unwrap());
        let raw_parent = u16::from_le_bytes(body[2..4].try_into().unwrap());
        let raw_base = u16::from_le_bytes(body[4..6].try_into().unwrap());
        let group_count = usize::from(u16::from_le_bytes(body[8..10].try_into().unwrap()));
        let base_count = usize::from(u16::from_le_bytes(body[10..12].try_into().unwrap()));
        let original_count = usize::from(u16::from_le_bytes(body[12..14].try_into().unwrap()));
        let (name, used) = parse_cache_string(&body[14..], 0x00C7)?;
        if 14 + used != body.len() {
            return Err(cache_invalid(0x00C7, "trailing SXFDB payload"));
        }
        position += 1;
        let (kind, body) = records
            .get(position)
            .ok_or_else(|| cache_invalid(0x01BB, "missing SXFDBTYPE"))?;
        if *kind != 0x01BB || body != &[0, 0] {
            return Err(cache_invalid(*kind, "invalid SXFDBTYPE"));
        }
        position += 1;
        let mut group_items = Vec::with_capacity(group_count);
        for _ in 0..group_count {
            let (kind, body) = records
                .get(position)
                .ok_or_else(|| cache_invalid(0x00C7, "missing PivotCache group item"))?;
            group_items.push(parse_cache_item(*kind, body)?);
            position += 1;
        }
        let grouping = if let Some((0x00D9, body)) = records.get(position) {
            if body.len() != base_count * 2 {
                return Err(cache_invalid(
                    0x00D9,
                    "SXGROUPINFO size does not match base-item count",
                ));
            }
            let item_to_group = body
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            if item_to_group
                .iter()
                .any(|index| usize::from(*index) >= group_items.len())
            {
                return Err(cache_invalid(
                    0x00D9,
                    "SXGROUPINFO group ordinal is out of range",
                ));
            }
            position += 1;
            Some(PivotCacheGrouping::Discrete(PivotCacheDiscreteGrouping {
                base_field_index: raw_base,
                group_items: group_items.clone(),
                item_to_group,
            }))
        } else if let Some((0x00D8, body)) = records.get(position) {
            if body.len() != 2 {
                return Err(XlsError::InvalidLength {
                    expected: 2,
                    found: body.len(),
                });
            }
            let group_flags = u16::from_le_bytes((*body).try_into().unwrap());
            let group_type = (group_flags >> 2) & 0xF;
            position += 1;
            let mut limits = Vec::with_capacity(3);
            for _ in 0..3 {
                let (kind, body) = records
                    .get(position)
                    .ok_or_else(|| cache_invalid(0x00D8, "missing SXNUMGROUP limit item"))?;
                limits.push(parse_cache_item(*kind, body)?);
                position += 1;
            }
            if group_type == 8 {
                let numbers = limits
                    .iter()
                    .map(|item| match item {
                        PivotCacheItem::Number(value) => Ok(*value),
                        _ => Err(cache_invalid(
                            0x00D8,
                            "numeric grouping limits must be numeric",
                        )),
                    })
                    .collect::<XlsResult<Vec<_>>>()?;
                Some(PivotCacheGrouping::Numeric(PivotCacheNumericGrouping {
                    start: numbers[0],
                    end: numbers[1],
                    step: numbers[2],
                    auto_start: group_flags & 1 != 0,
                    auto_end: group_flags & 2 != 0,
                    group_items: group_items.clone(),
                }))
            } else {
                let start = match limits[0] {
                    PivotCacheItem::DateTime(value) => value,
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping start must be SXDATETIME",
                        ));
                    },
                };
                let end = match limits[1] {
                    PivotCacheItem::DateTime(value) => value,
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping end must be SXDATETIME",
                        ));
                    },
                };
                let step = match limits[2] {
                    PivotCacheItem::Number(value)
                        if value >= 1.0 && value <= f64::from(u16::MAX) && value.fract() == 0.0 =>
                    {
                        value as u16
                    },
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping step must be a positive integer",
                        ));
                    },
                };
                Some(PivotCacheGrouping::Date(PivotCacheDateGrouping {
                    unit: PivotCacheDateGroupUnit::try_from(group_type)?,
                    start,
                    end,
                    step,
                    auto_start: group_flags & 1 != 0,
                    auto_end: group_flags & 2 != 0,
                    group_items: group_items.clone(),
                }))
            }
        } else {
            None
        };
        let mut items = Vec::with_capacity(original_count);
        for _ in 0..original_count {
            let (kind, body) = records
                .get(position)
                .ok_or_else(|| cache_invalid(0x00C7, "missing shared PivotCache item"))?;
            items.push(parse_cache_item(*kind, body)?);
            position += 1;
        }
        if (flags & 1 != 0) == (items.is_empty() && group_items.is_empty()) {
            return Err(cache_invalid(0x00C7, "SXFDB item flag/count mismatch"));
        }
        fields.push(PivotCacheField {
            name,
            flags,
            group_parent: (flags & 0x0008 != 0).then_some(raw_parent),
            group_base: matches!(grouping, Some(PivotCacheGrouping::Discrete(_)))
                .then_some(raw_base),
            items,
            grouping,
        });
    }
    let mut rows = Vec::new();
    while let Some((kind, body)) = records.get(position) {
        if *kind == 0x000A {
            position += 1;
            break;
        }
        if *kind != 0x00C8 {
            return Err(cache_invalid(*kind, "expected SXINDEXLIST or EOF"));
        }
        let mut body_offset = 0usize;
        let mut row = Vec::with_capacity(fields.len());
        position += 1;
        for field in fields.iter().take(standard_field_count) {
            if field.items.is_empty() {
                let (item_type, item_body) = records
                    .get(position)
                    .ok_or_else(|| cache_invalid(0x00C8, "missing inline PivotCache row item"))?;
                row.push(parse_cache_item(*item_type, item_body)?);
                position += 1;
            } else {
                let width = if field.flags & 0x0200 != 0 { 2 } else { 1 };
                let encoded =
                    body.get(body_offset..body_offset + width)
                        .ok_or(XlsError::InvalidLength {
                            expected: body_offset + width,
                            found: body.len(),
                        })?;
                let index = if width == 2 {
                    usize::from(u16::from_le_bytes(encoded.try_into().unwrap()))
                } else {
                    usize::from(encoded[0])
                };
                row.push(
                    field
                        .items
                        .get(index)
                        .ok_or_else(|| {
                            cache_invalid(0x00C8, "SXINDEXLIST item index out of range")
                        })?
                        .clone(),
                );
                body_offset += width;
            }
        }
        if body_offset != body.len() {
            return Err(cache_invalid(
                0x00C8,
                "SXINDEXLIST payload size does not match fields",
            ));
        }
        rows.push(row);
    }
    if position != records.len() {
        return Err(cache_invalid(
            0x000A,
            "trailing PivotCache records after EOF",
        ));
    }
    if cache_flags & 1 != 0 && rows.len() != record_count as usize {
        return Err(cache_invalid(0x00C6, "saved PivotCache row count mismatch"));
    }
    Ok(PivotCache {
        stream_id,
        flags: cache_flags,
        record_count,
        fields,
        rows,
    })
}

// ---------------------------------------------------------------------------
// Record type constants
// ---------------------------------------------------------------------------

/// SXVIEW record type.
pub const SXVIEW_TYPE: u16 = 0x00B0;
/// SXVD (View Fields) record type.
pub const SXVD_TYPE: u16 = 0x00B1;
/// SXVI (View Item) record type.
pub const SXVI_TYPE: u16 = 0x00B2;
/// SXPI (Page Item) record type.
pub const SXPI_TYPE: u16 = 0x00B6;
/// SXDI (Data Item) record type.
pub const SXDI_TYPE: u16 = 0x00C5;
/// SXVS (View Source) record type.
pub const SXVS_TYPE: u16 = 0x00E3;
pub const SXIVD_TYPE: u16 = 0x00B4;
pub const SXLI_TYPE: u16 = 0x00B5;
pub const SXEX_TYPE: u16 = 0x00F1;
pub const SXVDEX_TYPE: u16 = 0x0100;
pub const QSI_SX_TAG_TYPE: u16 = 0x0802;
pub const SXVIEWEX9_TYPE: u16 = 0x0810;
pub const SXADDL_TYPE: u16 = 0x0864;

const DATA_LAYOUT_FIELD: u16 = 0xFFFE;
const MAX_PIVOT_VIEWS_PER_SHEET: usize = 1_024;
const MAX_PIVOT_FIELDS: usize = 4_096;
const MAX_PIVOT_ITEMS: usize = 1_048_576;
const MAX_PIVOT_EXTENSION_BYTES: usize = 1_048_576;

pub(crate) const fn is_worksheet_view_record(record_type: u16) -> bool {
    matches!(
        record_type,
        SXVIEW_TYPE
            | SXVD_TYPE
            | SXVI_TYPE
            | SXIVD_TYPE
            | SXLI_TYPE
            | SXPI_TYPE
            | SXDI_TYPE
            | SXEX_TYPE
            | SXVDEX_TYPE
            | QSI_SX_TAG_TYPE
            | SXVIEWEX9_TYPE
            | SXADDL_TYPE
    )
}

// ---------------------------------------------------------------------------
// Axis constants (used by SXVD and SXDI)
// ---------------------------------------------------------------------------

/// Pivot field axis placement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAxis {
    /// No axis (hidden / unused).
    None,
    /// Row axis.
    Row,
    /// Column axis.
    Column,
    /// Page (filter) axis.
    Page,
    /// Data (values) axis.
    Data,
}

impl PivotAxis {
    fn from_u16(val: u16) -> XlsResult<Self> {
        match val {
            0x0000 => Ok(Self::None),
            0x0001 => Ok(Self::Row),
            0x0002 => Ok(Self::Column),
            0x0004 => Ok(Self::Page),
            0x0008 => Ok(Self::Data),
            _ => Err(cache_invalid(
                SXVD_TYPE,
                format!("invalid PivotTable axis 0x{val:04X}"),
            )),
        }
    }
    pub const fn code(self) -> u16 {
        match self {
            Self::None => 0,
            Self::Row => 1,
            Self::Column => 2,
            Self::Page => 4,
            Self::Data => 8,
        }
    }
}

/// Aggregation function for data items.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotFunction {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevP,
    Var,
    VarP,
    Unknown(u16),
}

impl PivotFunction {
    fn from_u16(val: u16) -> Self {
        match val {
            0x0000 => Self::Sum,
            0x0001 => Self::Count,
            0x0002 => Self::Average,
            0x0003 => Self::Max,
            0x0004 => Self::Min,
            0x0005 => Self::Product,
            0x0006 => Self::CountNums,
            0x0007 => Self::StdDev,
            0x0008 => Self::StdDevP,
            0x0009 => Self::Var,
            0x000A => Self::VarP,
            other => Self::Unknown(other),
        }
    }
    pub const fn code(self) -> u16 {
        match self {
            Self::Sum => 0,
            Self::Count => 1,
            Self::Average => 2,
            Self::Max => 3,
            Self::Min => 4,
            Self::Product => 5,
            Self::CountNums => 6,
            Self::StdDev => 7,
            Self::StdDevP => 8,
            Self::Var => 9,
            Self::VarP => 10,
            Self::Unknown(value) => value,
        }
    }
}

// ---------------------------------------------------------------------------
// SXVIEW — View Definition
// ---------------------------------------------------------------------------

/// Parsed SXVIEW record (pivot table header / definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewDef {
    /// First row of the pivot table output range.
    pub first_row: u16,
    /// Last row of the pivot table output range.
    pub last_row: u16,
    /// First column of the pivot table output range.
    pub first_col: u16,
    /// Last column of the pivot table output range.
    pub last_col: u16,
    /// First header row.
    pub first_header_row: u16,
    /// First data row (body).
    pub first_data_row: u16,
    /// First data column (body).
    pub first_data_col: u16,
    /// Zero-based index into the workbook's global PivotCache list.
    pub cache_index: u16,
    /// Number of row fields.
    pub row_field_count: u16,
    /// Number of column fields.
    pub col_field_count: u16,
    /// Number of page (filter) fields.
    pub page_field_count: u16,
    /// Number of data (value) fields.
    pub data_field_count: u16,
    /// Total number of data rows in the source.
    pub data_row_count: u16,
    /// Total number of visible data columns.
    pub data_col_count: u16,
    /// Total number of fields (dimensions).
    pub field_count: u16,
    /// Axis used for the data field header (when >1 data field).
    pub data_axis: PivotAxis,
    /// Position of data field label within the axis.
    pub data_position: u16,
    /// SXVIEW option flags.
    pub flags: u16,
    /// Built-in PivotTable auto-format index.
    pub auto_format_index: u16,
    /// Name of the pivot table.
    pub name: String,
    /// Name of the data field header (e.g. "Values").
    pub data_field_name: String,
}

/// Parse an SXVIEW record.
///
/// Layout (Apache POI `ViewDefinitionRecord`):
/// ```text
///  0  u16  rwFirst
///  2  u16  rwLast
///  4  u16  colFirst
///  6  u16  colLast
///  8  u16  rwFirstHead
/// 10  u16  rwFirstData
/// 12  u16  colFirstData
/// 14  u16  cDimRw       (row field count)
/// 16  u16  cDimCol
/// 18  u16  cDimPg
/// 20  u16  cDimData
/// 22  u16  cRw          (data row count)
/// 24  u16  cDim         (total field count)
/// 26  u16  cItm         (unused)
/// 28  u16  cITMData     (unused)
/// 30  u16  sxaxis4Data
/// 32  u16  ipos4Data
/// 34  u16  cchName      (length of name)
/// 36  u16  cchData      (length of data field name)
/// 38  var  name (XLUnicodeStringNoCch)
///     var  dataField (XLUnicodeStringNoCch)
/// ```
pub fn parse_sxview(data: &[u8]) -> XlsResult<PivotViewDef> {
    if data.len() < 44 {
        return Err(XlsError::InvalidLength {
            expected: 44,
            found: data.len(),
        });
    }

    let first_row = binary::read_u16_le_at(data, 0)?;
    let last_row = binary::read_u16_le_at(data, 2)?;
    let first_col = binary::read_u16_le_at(data, 4)?;
    let last_col = binary::read_u16_le_at(data, 6)?;
    let first_header_row = binary::read_u16_le_at(data, 8)?;
    let first_data_row = binary::read_u16_le_at(data, 10)?;
    let first_data_col = binary::read_u16_le_at(data, 12)?;
    let cache_index = binary::read_u16_le_at(data, 14)?;
    if binary::read_u16_le_at(data, 16)? != 0 {
        return Err(cache_invalid(SXVIEW_TYPE, "nonzero SXVIEW reserved field"));
    }
    let data_axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 18)?)?;
    let data_position = binary::read_u16_le_at(data, 20)?;
    let field_count = binary::read_u16_le_at(data, 22)?;
    let row_field_count = binary::read_u16_le_at(data, 24)?;
    let col_field_count = binary::read_u16_le_at(data, 26)?;
    let page_field_count = binary::read_u16_le_at(data, 28)?;
    let data_field_count = binary::read_u16_le_at(data, 30)?;
    let data_row_count = binary::read_u16_le_at(data, 32)?;
    let data_col_count = binary::read_u16_le_at(data, 34)?;
    let flags = binary::read_u16_le_at(data, 36)?;
    let auto_format_index = binary::read_u16_le_at(data, 38)?;
    let cch_name = usize::from(binary::read_u16_le_at(data, 40)?);
    let cch_data = usize::from(binary::read_u16_le_at(data, 42)?);
    if usize::from(field_count) > MAX_PIVOT_FIELDS {
        return Err(cache_invalid(
            SXVIEW_TYPE,
            "PivotTable field count exceeds resource bound",
        ));
    }
    if first_row > last_row
        || first_col > last_col
        || first_header_row < first_row
        || first_header_row > last_row
        || first_data_row < first_row
        || first_data_row > last_row
        || first_data_col < first_col
        || first_data_col > last_col
    {
        return Err(cache_invalid(
            SXVIEW_TYPE,
            "invalid or reversed PivotTable output range",
        ));
    }

    let mut offset = 44;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;
    let data_field_name = read_xl_string_no_cch(data, &mut offset, cch_data)?;
    if offset != data.len() {
        return Err(cache_invalid(SXVIEW_TYPE, "trailing SXVIEW payload"));
    }

    Ok(PivotViewDef {
        first_row,
        last_row,
        first_col,
        last_col,
        first_header_row,
        first_data_row,
        first_data_col,
        cache_index,
        row_field_count,
        col_field_count,
        page_field_count,
        data_field_count,
        data_row_count,
        data_col_count,
        field_count,
        data_axis,
        data_position,
        flags,
        auto_format_index,
        name,
        data_field_name,
    })
}

// ---------------------------------------------------------------------------
// SXVD — View Field
// ---------------------------------------------------------------------------

/// Parsed SXVD record (single pivot field definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewField {
    /// Axis this field is assigned to.
    pub axis: PivotAxis,
    /// Number of subtotals.
    pub subtotal_count: u16,
    /// Subtotal function bitmask.
    pub subtotal_flags: u16,
    /// Number of items in this field.
    pub item_count: u16,
    /// Optional field name override (empty string = use source name).
    pub name: Option<String>,
    /// SXVI records owned by this field.
    pub items: Vec<PivotViewItem>,
    /// Mandatory legacy extended-field properties.
    pub extension: Option<PivotViewFieldExtension>,
    /// Losslessly preserved future records scoped to this field.
    pub additional_extensions: Vec<PivotAdditionalExtension>,
}

/// Parse an SXVD record.
///
/// Layout:
/// ```text
///  0  u16  sxaxis   (axis)
///  2  u16  cSub     (subtotal count)
///  4  u16  grbitSub (subtotal flags)
///  6  u16  cItm     (item count)
///  8  u16  cchName  (0xFFFF = not present)
/// 10  var  name (XLUnicodeStringNoCch)  — only if cchName != 0xFFFF
/// ```
pub fn parse_sxvd(data: &[u8]) -> XlsResult<PivotViewField> {
    if data.len() < 10 {
        return Err(XlsError::InvalidLength {
            expected: 10,
            found: data.len(),
        });
    }

    let axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 0)?)?;
    let subtotal_count = binary::read_u16_le_at(data, 2)?;
    let subtotal_flags = binary::read_u16_le_at(data, 4)?;
    let item_count = binary::read_u16_le_at(data, 6)?;
    let cch_name = binary::read_u16_le_at(data, 8)?;

    let mut offset = 10;
    let name = if cch_name != 0xFFFF {
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    } else {
        None
    };

    if offset != data.len() {
        return Err(cache_invalid(SXVD_TYPE, "trailing SXVD payload"));
    }
    Ok(PivotViewField {
        axis,
        subtotal_count,
        subtotal_flags,
        item_count,
        name,
        items: Vec::with_capacity(usize::from(item_count)),
        extension: None,
        additional_extensions: Vec::new(),
    })
}

// ---------------------------------------------------------------------------
// SXVI — View Item
// ---------------------------------------------------------------------------

/// Item type within a pivot field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotItemType {
    Data,
    Default,
    Sum,
    CountA,
    Average,
    Max,
    Min,
    Product,
    Count,
    StdDev,
    StdDevP,
    Var,
    VarP,
    Grand,
    Blank,
    Unknown(u16),
}

impl PivotItemType {
    fn from_u16(val: u16) -> Self {
        match val {
            0xFE => Self::Data,
            0xFF => Self::Default,
            0x00 => Self::Sum,
            0x01 => Self::CountA,
            0x02 => Self::Average,
            0x03 => Self::Max,
            0x04 => Self::Min,
            0x05 => Self::Product,
            0x06 => Self::Count,
            0x07 => Self::StdDev,
            0x08 => Self::StdDevP,
            0x09 => Self::Var,
            0x0A => Self::VarP,
            0x0B => Self::Grand,
            0x0C => Self::Blank,
            other => Self::Unknown(other),
        }
    }

    pub const fn code(self) -> u16 {
        match self {
            Self::Data => 0xFE,
            Self::Default => 0xFF,
            Self::Sum => 0x00,
            Self::CountA => 0x01,
            Self::Average => 0x02,
            Self::Max => 0x03,
            Self::Min => 0x04,
            Self::Product => 0x05,
            Self::Count => 0x06,
            Self::StdDev => 0x07,
            Self::StdDevP => 0x08,
            Self::Var => 0x09,
            Self::VarP => 0x0A,
            Self::Grand => 0x0B,
            Self::Blank => 0x0C,
            Self::Unknown(value) => value,
        }
    }
}

/// Parsed SXVI record (pivot field item).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewItem {
    /// Item type.
    pub item_type: PivotItemType,
    /// Option flags.
    pub flags: u16,
    /// Cache index.
    pub cache_index: u16,
    /// Optional item name override.
    pub name: Option<String>,
}

/// Parse an SXVI record.
///
/// Layout:
/// ```text
///  0  u16  itmType
///  2  u16  grbitItem
///  4  u16  iCache
///  6  u16  cchName  (0xFFFF = not present)
///  8  var  name
/// ```
pub fn parse_sxvi(data: &[u8]) -> XlsResult<PivotViewItem> {
    if data.len() < 8 {
        return Err(XlsError::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }

    let item_type = PivotItemType::from_u16(binary::read_u16_le_at(data, 0)?);
    let flags = binary::read_u16_le_at(data, 2)?;
    let cache_index = binary::read_u16_le_at(data, 4)?;
    let cch_name = binary::read_u16_le_at(data, 6)?;

    let mut offset = 8;
    let name = if cch_name != 0xFFFF {
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    } else {
        None
    };

    if offset != data.len() {
        return Err(cache_invalid(SXVI_TYPE, "trailing SXVI payload"));
    }
    Ok(PivotViewItem {
        item_type,
        flags,
        cache_index,
        name,
    })
}

// ---------------------------------------------------------------------------
// SXDI — Data Item
// ---------------------------------------------------------------------------

/// Parsed SXDI record (data/value field definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotDataItem {
    /// Index of the source field in the pivot cache.
    pub source_field_index: u16,
    /// Aggregation function.
    pub function: PivotFunction,
    /// Display format flags.
    pub display_format: u16,
    /// Index into SXVD for base field (used for "show values as").
    pub base_field_index: u16,
    /// Index into SXVI for base item.
    pub base_item_index: u16,
    /// Number format index.
    pub num_format_index: u16,
    /// Optional name override.
    pub name: String,
}

/// Parse an SXDI record.
///
/// Layout (POI `DataItemRecord`):
/// ```text
///  0  u16  isxvdData   (source field index)
///  2  u16  iiftab      (aggregation function)
///  4  u16  df          (display format)
///  6  u16  isxvd       (base field index)
///  8  u16  isxvi       (base item index)
/// 10  u16  ifmt        (number format)
/// 12  u16  cchName
/// 14  var  name
/// ```
pub fn parse_sxdi(data: &[u8]) -> XlsResult<PivotDataItem> {
    if data.len() < 14 {
        return Err(XlsError::InvalidLength {
            expected: 14,
            found: data.len(),
        });
    }

    let source_field_index = binary::read_u16_le_at(data, 0)?;
    let function = PivotFunction::from_u16(binary::read_u16_le_at(data, 2)?);
    let display_format = binary::read_u16_le_at(data, 4)?;
    let base_field_index = binary::read_u16_le_at(data, 6)?;
    let base_item_index = binary::read_u16_le_at(data, 8)?;
    let num_format_index = binary::read_u16_le_at(data, 10)?;
    let cch_name = binary::read_u16_le_at(data, 12)? as usize;

    let mut offset = 14;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;
    if offset != data.len() {
        return Err(cache_invalid(SXDI_TYPE, "trailing SXDI payload"));
    }

    Ok(PivotDataItem {
        source_field_index,
        function,
        display_format,
        base_field_index,
        base_item_index,
        num_format_index,
        name,
    })
}

// ---------------------------------------------------------------------------
// SXVS — View Source
// ---------------------------------------------------------------------------

/// Pivot cache source type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotSourceType {
    /// Data from a worksheet range.
    Worksheet,
    /// Data from an external source.
    External,
    /// Consolidation ranges.
    Consolidation,
    /// Data from a named range / scenario.
    Scenario,
    /// Unknown source type.
    Unknown(u16),
}

impl PivotSourceType {
    fn from_u16(val: u16) -> Self {
        match val {
            0x0001 => Self::Worksheet,
            0x0002 => Self::External,
            0x0004 => Self::Consolidation,
            0x0010 => Self::Scenario,
            other => Self::Unknown(other),
        }
    }
}

/// Parse an SXVS record (2 bytes: source type).
pub fn parse_sxvs(data: &[u8]) -> XlsResult<PivotSourceType> {
    if data.len() != 2 {
        return Err(XlsError::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(PivotSourceType::from_u16(binary::read_u16_le_at(data, 0)?))
}

// ---------------------------------------------------------------------------
// SXPI — Page Item
// ---------------------------------------------------------------------------

/// A single page field entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageFieldEntry {
    /// Index into SXVI for the selected item.
    pub item_index: u16,
    /// Index into SXVD for the field.
    pub field_index: u16,
    /// Object ID (unused in most cases).
    pub object_id: u16,
}

/// Typed page-field selection derived from an SXPI selected-item ordinal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotPageSelection {
    Item(u16),
    All,
}

impl PageFieldEntry {
    pub const fn selection(self) -> PivotPageSelection {
        if self.item_index == 0x7FFD {
            PivotPageSelection::All
        } else {
            PivotPageSelection::Item(self.item_index)
        }
    }
}

/// Parse an SXPI record.
///
/// Each entry is 6 bytes: `(isxvi: u16, isxvd: u16, idObj: u16)`.
/// The number of entries is `data.len() / 6`.
pub fn parse_sxpi(data: &[u8]) -> XlsResult<Vec<PageFieldEntry>> {
    if !data.len().is_multiple_of(6) {
        return Err(cache_invalid(
            SXPI_TYPE,
            "SXPI length is not a multiple of six",
        ));
    }
    let entry_count = data.len() / 6;
    let mut entries = Vec::with_capacity(entry_count);

    for i in 0..entry_count {
        let offset = i * 6;
        entries.push(PageFieldEntry {
            field_index: binary::read_u16_le_at(data, offset)?,
            item_index: binary::read_u16_le_at(data, offset + 2)?,
            object_id: binary::read_u16_le_at(data, offset + 4)?,
        });
    }

    Ok(entries)
}

/// One row or column axis entry from SXIVD.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAxisField {
    Field(u16),
    DataLayout,
}

pub fn parse_sxivd(data: &[u8]) -> XlsResult<Vec<PivotAxisField>> {
    if !data.len().is_multiple_of(2) {
        return Err(cache_invalid(SXIVD_TYPE, "SXIVD length must be even"));
    }
    data.chunks_exact(2)
        .map(|bytes| match u16::from_le_bytes([bytes[0], bytes[1]]) {
            DATA_LAYOUT_FIELD => Ok(PivotAxisField::DataLayout),
            value if value != u16::MAX => Ok(PivotAxisField::Field(value)),
            _ => Err(cache_invalid(SXIVD_TYPE, "invalid SXIVD field ordinal")),
        })
        .collect()
}

/// Legacy extended properties attached to one SXVD field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewFieldExtension {
    pub flags: u32,
    pub auto_sort_data_index: Option<u16>,
    pub auto_show_data_index: Option<u16>,
    pub number_format_index: u16,
    pub subtotal_name: Option<String>,
    /// Reserved bytes retained exactly for roundtrip-aware readers.
    pub reserved: [u8; 8],
}

pub fn parse_sxvdex(data: &[u8]) -> XlsResult<PivotViewFieldExtension> {
    if data.len() < 20 {
        return Err(XlsError::InvalidLength {
            expected: 20,
            found: data.len(),
        });
    }
    let cch = binary::read_u16_le_at(data, 10)?;
    let mut offset = 20;
    let subtotal_name = if cch == u16::MAX {
        None
    } else {
        Some(read_xl_string_no_cch(data, &mut offset, usize::from(cch))?)
    };
    if offset != data.len() {
        return Err(cache_invalid(SXVDEX_TYPE, "trailing SXVDEx payload"));
    }
    Ok(PivotViewFieldExtension {
        flags: binary::read_u32_le_at(data, 0)?,
        auto_sort_data_index: match binary::read_u16_le_at(data, 4)? {
            u16::MAX => None,
            value => Some(value),
        },
        auto_show_data_index: match binary::read_u16_le_at(data, 6)? {
            u16::MAX => None,
            value => Some(value),
        },
        number_format_index: binary::read_u16_le_at(data, 8)?,
        subtotal_name,
        reserved: data[12..20].try_into().expect("fixed length checked"),
    })
}

/// One visible layout line from an SXLI row/column line array.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotLayoutLine {
    pub repeated_item_count: u16,
    pub item_type: u16,
    pub custom_name_flags: u16,
    pub item_indices: Vec<u16>,
}

fn parse_sxli(
    data: &[u8],
    expected_lines: usize,
    max_indices: usize,
) -> XlsResult<Vec<PivotLayoutLine>> {
    if expected_lines == 0 || !data.len().is_multiple_of(expected_lines) {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI byte length is inconsistent with its declared line count",
        ));
    }
    let line_size = data.len() / expected_lines;
    if line_size < 8 || !(line_size - 8).is_multiple_of(2) {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI has an invalid fixed line size",
        ));
    }
    let index_count = (line_size - 8) / 2;
    if index_count > max_indices {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI item-index count exceeds the PivotTable field count",
        ));
    }
    let mut lines = Vec::with_capacity(expected_lines);
    for line in data.chunks_exact(line_size) {
        let declared_max = usize::from(u16::from_le_bytes([line[4], line[5]]));
        if declared_max > index_count && declared_max != usize::from(u16::MAX) {
            return Err(cache_invalid(
                SXLI_TYPE,
                "SXLI declared item ordinal exceeds its line payload",
            ));
        }
        let indices = line[8..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect();
        lines.push(PivotLayoutLine {
            repeated_item_count: u16::from_le_bytes([line[0], line[1]]),
            item_type: u16::from_le_bytes([line[2], line[3]]),
            custom_name_flags: u16::from_le_bytes([line[6], line[7]]),
            item_indices: indices,
        });
    }
    Ok(lines)
}

/// Extended PivotTable view properties from SXEX.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewExtension {
    pub format_count: u16,
    pub select_count: u16,
    pub page_rows: u16,
    pub page_cols: u16,
    pub flags: u32,
    pub error_string: Option<String>,
    pub null_string: Option<String>,
    pub tag: Option<String>,
    pub page_field_style: Option<String>,
    pub table_style: Option<String>,
    pub vacate_style: Option<String>,
}

fn parse_optional_sx_string(
    data: &[u8],
    offset: &mut usize,
    cch: u16,
) -> XlsResult<Option<String>> {
    if cch == u16::MAX {
        Ok(None)
    } else {
        read_xl_string_no_cch(data, offset, usize::from(cch)).map(Some)
    }
}

pub fn parse_sxex(data: &[u8]) -> XlsResult<PivotViewExtension> {
    if data.len() < 24 {
        return Err(XlsError::InvalidLength {
            expected: 24,
            found: data.len(),
        });
    }
    let lengths = [2usize, 4, 6, 18, 20, 22]
        .map(|offset| u16::from_le_bytes([data[offset], data[offset + 1]]));
    let mut offset = 24;
    let error_string = parse_optional_sx_string(data, &mut offset, lengths[0])?;
    let null_string = parse_optional_sx_string(data, &mut offset, lengths[1])?;
    let tag = parse_optional_sx_string(data, &mut offset, lengths[2])?;
    let page_field_style = parse_optional_sx_string(data, &mut offset, lengths[3])?;
    let table_style = parse_optional_sx_string(data, &mut offset, lengths[4])?;
    let vacate_style = parse_optional_sx_string(data, &mut offset, lengths[5])?;
    if offset != data.len() {
        return Err(cache_invalid(SXEX_TYPE, "trailing SXEX payload"));
    }
    Ok(PivotViewExtension {
        format_count: binary::read_u16_le_at(data, 0)?,
        select_count: binary::read_u16_le_at(data, 8)?,
        page_rows: binary::read_u16_le_at(data, 10)?,
        page_cols: binary::read_u16_le_at(data, 12)?,
        flags: binary::read_u32_le_at(data, 14)?,
        error_string,
        null_string,
        tag,
        page_field_style,
        table_style,
        vacate_style,
    })
}

/// Query/Pivot tag metadata from QsiSxTag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotQueryTag {
    pub table_type: u16,
    pub flags: u16,
    pub options: u32,
    pub last_refresh_version: u8,
    pub minimum_refresh_version: u8,
    pub first_created_version: u8,
    pub table_name: String,
    /// Bytes following the table name, retained exactly.
    pub trailing_payload: Vec<u8>,
}

fn read_xl_unicode_string(data: &[u8], offset: &mut usize) -> XlsResult<String> {
    let cch_end = offset
        .checked_add(2)
        .ok_or_else(|| cache_invalid(QSI_SX_TAG_TYPE, "string offset overflow"))?;
    let cch_bytes = data.get(*offset..cch_end).ok_or(XlsError::InvalidLength {
        expected: cch_end,
        found: data.len(),
    })?;
    *offset = cch_end;
    read_xl_string_no_cch(
        data,
        offset,
        usize::from(u16::from_le_bytes([cch_bytes[0], cch_bytes[1]])),
    )
}

pub fn parse_qsi_sx_tag(data: &[u8]) -> XlsResult<PivotQueryTag> {
    if data.len() < 19 {
        return Err(XlsError::InvalidLength {
            expected: 19,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != QSI_SX_TAG_TYPE || data[14] != 16 {
        return Err(cache_invalid(
            QSI_SX_TAG_TYPE,
            "invalid QsiSxTag FRT header",
        ));
    }
    let mut offset = 16;
    let table_name = read_xl_unicode_string(data, &mut offset)?;
    Ok(PivotQueryTag {
        table_type: binary::read_u16_le_at(data, 4)?,
        flags: binary::read_u16_le_at(data, 6)?,
        options: binary::read_u32_le_at(data, 8)?,
        last_refresh_version: data[12],
        minimum_refresh_version: data[13],
        first_created_version: data[15],
        table_name,
        trailing_payload: data[offset..].to_vec(),
    })
}

/// Excel 9+ PivotTable layout metadata from SXVIEWEX9.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewEx9 {
    pub frt_flags: u16,
    pub report_flags: u32,
    pub view_flags: u32,
    pub auto_format_index: u16,
    pub grand_total_name: String,
}

pub fn parse_sxviewex9(data: &[u8]) -> XlsResult<PivotViewEx9> {
    if data.len() < 17 {
        return Err(XlsError::InvalidLength {
            expected: 17,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != SXVIEWEX9_TYPE {
        return Err(cache_invalid(
            SXVIEWEX9_TYPE,
            "invalid SXVIEWEX9 FRT header",
        ));
    }
    let mut offset = 14;
    let grand_total_name = read_xl_unicode_string(data, &mut offset)?;
    if offset != data.len() {
        return Err(cache_invalid(SXVIEWEX9_TYPE, "trailing SXVIEWEX9 payload"));
    }
    Ok(PivotViewEx9 {
        frt_flags: binary::read_u16_le_at(data, 2)?,
        report_flags: binary::read_u32_le_at(data, 4)?,
        view_flags: binary::read_u32_le_at(data, 8)?,
        auto_format_index: binary::read_u16_le_at(data, 12)?,
        grand_total_name,
    })
}

/// Losslessly preserved SXADDL view- or field-extension record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotAdditionalExtension {
    pub class: u8,
    pub kind: u8,
    pub reserved: u16,
    pub payload: Vec<u8>,
}

pub fn parse_sxaddl(data: &[u8]) -> XlsResult<PivotAdditionalExtension> {
    if data.len() < 6 {
        return Err(XlsError::InvalidLength {
            expected: 6,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != SXADDL_TYPE {
        return Err(cache_invalid(SXADDL_TYPE, "invalid SXADDL FRT header"));
    }
    Ok(PivotAdditionalExtension {
        reserved: binary::read_u16_le_at(data, 2)?,
        class: data[4],
        kind: data[5],
        payload: data[6..].to_vec(),
    })
}

// ---------------------------------------------------------------------------
// Aggregate: PivotTable
// ---------------------------------------------------------------------------

/// Complete pivot table definition aggregated from multiple SX* records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTable {
    /// View definition (SXVIEW).
    pub view: PivotViewDef,
    /// Source type (SXVS).
    pub source_type: PivotSourceType,
    /// Field definitions (SXVD records, in order).
    pub fields: Vec<PivotViewField>,
    /// All items across all fields (SXVI records, in order).
    pub items: Vec<PivotViewItem>,
    /// Data field definitions (SXDI records).
    pub data_items: Vec<PivotDataItem>,
    /// Page field entries (SXPI records).
    pub page_entries: Vec<PageFieldEntry>,
    /// Explicit row-axis field ordering (first SXIVD).
    pub row_fields: Vec<PivotAxisField>,
    /// Explicit column-axis field ordering (second SXIVD).
    pub column_fields: Vec<PivotAxisField>,
    /// Visible row layout lines (first SXLI).
    pub row_lines: Vec<PivotLayoutLine>,
    /// Visible column layout lines (second SXLI).
    pub column_lines: Vec<PivotLayoutLine>,
    /// Legacy view extension (SXEX).
    pub extension: Option<PivotViewExtension>,
    /// Query/Pivot producer tag.
    pub query_tag: Option<PivotQueryTag>,
    /// Excel 9+ layout extension.
    pub view_ex9: Option<PivotViewEx9>,
    /// Losslessly preserved view-scoped SXADDL records.
    pub additional_extensions: Vec<PivotAdditionalExtension>,
}

impl PivotTable {
    /// Create a new pivot table from its view definition.
    pub fn new(view: PivotViewDef) -> Self {
        Self {
            source_type: PivotSourceType::Worksheet,
            fields: Vec::with_capacity(view.field_count as usize),
            items: Vec::new(),
            data_items: Vec::with_capacity(view.data_field_count as usize),
            page_entries: Vec::with_capacity(view.page_field_count as usize),
            row_fields: Vec::with_capacity(view.row_field_count as usize),
            column_fields: Vec::with_capacity(view.col_field_count as usize),
            row_lines: Vec::with_capacity(view.data_row_count as usize),
            column_lines: Vec::with_capacity(view.data_col_count as usize),
            extension: None,
            query_tag: None,
            view_ex9: None,
            additional_extensions: Vec::new(),
            view,
        }
    }

    pub const fn cache_index(&self) -> u16 {
        self.view.cache_index
    }

    /// Returns the cache field addressed by a view-field ordinal.
    pub fn cache_field<'a>(
        &self,
        cache: &'a PivotCache,
        field_index: u16,
    ) -> Option<&'a PivotCacheField> {
        cache.fields().get(usize::from(field_index))
    }
}

struct PivotTableBuild {
    table: PivotTable,
    row_axis_seen: bool,
    column_axis_seen: bool,
    page_seen: bool,
    row_lines_seen: bool,
    column_lines_seen: bool,
    sxaddl_field_cursor: usize,
    sxaddl_field_open: bool,
    extension_bytes: usize,
}

impl PivotTableBuild {
    fn new(view: PivotViewDef) -> Self {
        Self {
            table: PivotTable::new(view),
            row_axis_seen: false,
            column_axis_seen: false,
            page_seen: false,
            row_lines_seen: false,
            column_lines_seen: false,
            sxaddl_field_cursor: 0,
            sxaddl_field_open: false,
            extension_bytes: 0,
        }
    }

    fn fields_complete(&self) -> bool {
        self.table.fields.len() == usize::from(self.table.view.field_count)
            && self.table.fields.iter().all(|field| {
                field.items.len() == usize::from(field.item_count) && field.extension.is_some()
            })
    }

    fn axes_complete(&self) -> bool {
        (self.table.view.row_field_count == 0 || self.row_axis_seen)
            && (self.table.view.col_field_count == 0 || self.column_axis_seen)
    }

    fn page_complete(&self) -> bool {
        self.table.view.page_field_count == 0 || self.page_seen
    }
    fn data_complete(&self) -> bool {
        self.table.data_items.len() == usize::from(self.table.view.data_field_count)
    }
    fn lines_complete(&self) -> bool {
        (self.table.view.data_row_count == 0 || self.row_lines_seen)
            && (self.table.view.data_col_count == 0 || self.column_lines_seen)
    }

    fn require_fields(&self, record_type: u16) -> XlsResult<()> {
        if self.fields_complete() {
            Ok(())
        } else {
            Err(cache_invalid(
                record_type,
                "record appears before all SXVD/SXVI/SXVDEx field groups",
            ))
        }
    }

    fn add_extension_bytes(&mut self, record_type: u16, count: usize) -> XlsResult<()> {
        self.extension_bytes = self
            .extension_bytes
            .checked_add(count)
            .ok_or_else(|| cache_invalid(record_type, "PivotTable extension size overflow"))?;
        if self.extension_bytes > MAX_PIVOT_EXTENSION_BYTES {
            return Err(cache_invalid(
                record_type,
                "PivotTable extensions exceed resource bound",
            ));
        }
        Ok(())
    }

    fn feed(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            SXVD_TYPE => {
                if let Some(previous) = self.table.fields.last()
                    && (previous.extension.is_none()
                        || previous.items.len() != usize::from(previous.item_count))
                {
                    return Err(cache_invalid(
                        record_type,
                        "SXVD starts before the previous field group is complete",
                    ));
                }
                if self.table.fields.len() == usize::from(self.table.view.field_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVD count exceeds SXVIEW field count",
                    ));
                }
                self.table.fields.push(parse_sxvd(data)?);
            },
            SXVI_TYPE => {
                let field = self
                    .table
                    .fields
                    .last_mut()
                    .ok_or_else(|| cache_invalid(record_type, "SXVI appears without SXVD"))?;
                if field.extension.is_some() {
                    return Err(cache_invalid(record_type, "SXVI appears after SXVDEx"));
                }
                if field.items.len() == usize::from(field.item_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVI count exceeds SXVD item count",
                    ));
                }
                let item = parse_sxvi(data)?;
                field.items.push(item.clone());
                self.table.items.push(item);
                if self.table.items.len() > MAX_PIVOT_ITEMS {
                    return Err(cache_invalid(
                        record_type,
                        "PivotTable items exceed resource bound",
                    ));
                }
            },
            SXVDEX_TYPE => {
                let extension = parse_sxvdex(data)?;
                let field = self
                    .table
                    .fields
                    .last_mut()
                    .ok_or_else(|| cache_invalid(record_type, "SXVDEx appears without SXVD"))?;
                if field.items.len() != usize::from(field.item_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVDEx appears before all declared SXVI items",
                    ));
                }
                if field.extension.replace(extension).is_some() {
                    return Err(cache_invalid(record_type, "duplicate SXVDEx"));
                }
            },
            SXIVD_TYPE => {
                self.require_fields(record_type)?;
                let fields = parse_sxivd(data)?;
                if self.table.view.row_field_count != 0 && !self.row_axis_seen {
                    if fields.len() != usize::from(self.table.view.row_field_count) {
                        return Err(cache_invalid(
                            record_type,
                            "row SXIVD count does not match SXVIEW",
                        ));
                    }
                    self.table.row_fields = fields;
                    self.row_axis_seen = true;
                } else if self.table.view.col_field_count != 0 && !self.column_axis_seen {
                    if fields.len() != usize::from(self.table.view.col_field_count) {
                        return Err(cache_invalid(
                            record_type,
                            "column SXIVD count does not match SXVIEW",
                        ));
                    }
                    self.table.column_fields = fields;
                    self.column_axis_seen = true;
                } else {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order SXIVD",
                    ));
                }
            },
            SXPI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() {
                    return Err(cache_invalid(record_type, "SXPI appears before SXIVD axes"));
                }
                if self.page_seen || self.table.view.page_field_count == 0 {
                    return Err(cache_invalid(record_type, "duplicate or unexpected SXPI"));
                }
                let entries = parse_sxpi(data)?;
                if entries.len() != usize::from(self.table.view.page_field_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXPI count does not match SXVIEW",
                    ));
                }
                self.table.page_entries = entries;
                self.page_seen = true;
            },
            SXDI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() || !self.page_complete() {
                    return Err(cache_invalid(
                        record_type,
                        "SXDI appears before axes/page fields",
                    ));
                }
                if self.table.data_items.len() == usize::from(self.table.view.data_field_count) {
                    return Err(cache_invalid(record_type, "SXDI count exceeds SXVIEW"));
                }
                self.table.data_items.push(parse_sxdi(data)?);
            },
            SXLI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() || !self.page_complete() || !self.data_complete() {
                    return Err(cache_invalid(
                        record_type,
                        "SXLI appears before axis/page/data records",
                    ));
                }
                if self.table.view.data_row_count != 0 && !self.row_lines_seen {
                    self.table.row_lines = parse_sxli(
                        data,
                        usize::from(self.table.view.data_row_count),
                        usize::from(self.table.view.field_count).saturating_add(1),
                    )?;
                    self.row_lines_seen = true;
                } else if self.table.view.data_col_count != 0 && !self.column_lines_seen {
                    self.table.column_lines = parse_sxli(
                        data,
                        usize::from(self.table.view.data_col_count),
                        usize::from(self.table.view.field_count).saturating_add(1),
                    )?;
                    self.column_lines_seen = true;
                } else {
                    return Err(cache_invalid(record_type, "duplicate or out-of-order SXLI"));
                }
            },
            SXEX_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete()
                    || !self.page_complete()
                    || !self.data_complete()
                    || !self.lines_complete()
                {
                    return Err(cache_invalid(
                        record_type,
                        "SXEX appears before the core PivotTable view is complete",
                    ));
                }
                if self.table.extension.replace(parse_sxex(data)?).is_some() {
                    return Err(cache_invalid(record_type, "duplicate SXEX"));
                }
            },
            QSI_SX_TAG_TYPE => {
                if self.table.extension.is_none() {
                    return Err(cache_invalid(record_type, "QsiSxTag appears before SXEX"));
                }
                if self.table.query_tag.is_some()
                    || self.table.view_ex9.is_some()
                    || !self.table.additional_extensions.is_empty()
                {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order QsiSxTag",
                    ));
                }
                self.add_extension_bytes(record_type, data.len())?;
                let tag = parse_qsi_sx_tag(data)?;
                if tag.table_name != self.table.view.name {
                    return Err(cache_invalid(
                        record_type,
                        "QsiSxTag table name does not match SXVIEW",
                    ));
                }
                self.table.query_tag = Some(tag);
            },
            SXVIEWEX9_TYPE => {
                if self.table.query_tag.is_none()
                    || self.table.view_ex9.is_some()
                    || !self.table.additional_extensions.is_empty()
                {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order SXVIEWEX9",
                    ));
                }
                self.add_extension_bytes(record_type, data.len())?;
                self.table.view_ex9 = Some(parse_sxviewex9(data)?);
            },
            SXADDL_TYPE => {
                if self.table.extension.is_none() {
                    return Err(cache_invalid(record_type, "SXADDL appears before SXEX"));
                }
                self.add_extension_bytes(record_type, data.len())?;
                let extension = parse_sxaddl(data)?;
                if extension.class == 0x17 {
                    if self.sxaddl_field_cursor >= self.table.fields.len() {
                        return Err(cache_invalid(
                            record_type,
                            "field SXADDL ordinal exceeds SXVD count",
                        ));
                    }
                    if !self.sxaddl_field_open && extension.kind != 0x00 {
                        return Err(cache_invalid(
                            record_type,
                            "field SXADDL group does not start with a name record",
                        ));
                    }
                    self.sxaddl_field_open = extension.kind != 0xFF;
                    self.table.fields[self.sxaddl_field_cursor]
                        .additional_extensions
                        .push(extension);
                    if !self.sxaddl_field_open {
                        self.sxaddl_field_cursor += 1;
                    }
                } else {
                    if self.sxaddl_field_open {
                        return Err(cache_invalid(
                            record_type,
                            "view SXADDL interrupts a field extension group",
                        ));
                    }
                    self.table.additional_extensions.push(extension);
                }
            },
            _ => return Err(cache_invalid(record_type, "unexpected PivotTable record")),
        }
        Ok(())
    }

    fn finish(self) -> XlsResult<PivotTable> {
        if !self.fields_complete()
            || !self.axes_complete()
            || !self.page_complete()
            || !self.data_complete()
            || !self.lines_complete()
            || self.table.extension.is_none()
            || self.sxaddl_field_open
        {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "incomplete or unterminated PivotTable worksheet-view record set",
            ));
        }
        Ok(self.table)
    }
}

/// Ordered worksheet PivotTable record collector.
pub(crate) struct PivotTableCollector {
    current: Option<PivotTableBuild>,
    completed: Vec<PivotTable>,
}

impl PivotTableCollector {
    pub(crate) fn new() -> Self {
        Self {
            current: None,
            completed: Vec::new(),
        }
    }

    fn push_current(&mut self) -> XlsResult<()> {
        let Some(build) = self.current.take() else {
            return Ok(());
        };
        let table = build.finish()?;
        if self.completed.len() == MAX_PIVOT_VIEWS_PER_SHEET {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "PivotTable view count exceeds resource bound",
            ));
        }
        if self
            .completed
            .iter()
            .any(|prior| ranges_overlap(&prior.view, &table.view))
        {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "overlapping PivotTable output ranges",
            ));
        }
        self.completed.push(table);
        Ok(())
    }

    /// Returns true when the record belongs to the PivotTable aggregate.
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<bool> {
        let pivot_record = is_worksheet_view_record(record_type);
        if record_type == SXVIEW_TYPE {
            self.push_current()?;
            self.current = Some(PivotTableBuild::new(parse_sxview(data)?));
            return Ok(true);
        }
        if pivot_record {
            if self.current.is_none()
                && record_type == QSI_SX_TAG_TYPE
                && parse_qsi_sx_tag(data)?.table_type != 1
            {
                return Ok(false);
            }
            let build = self.current.as_mut().ok_or_else(|| {
                cache_invalid(record_type, "orphan PivotTable worksheet-view record")
            })?;
            build.feed(record_type, data)?;
            return Ok(true);
        }
        if self
            .current
            .as_ref()
            .is_some_and(|build| build.table.extension.is_none())
        {
            self.push_current()?;
        }
        Ok(false)
    }

    pub(crate) fn finish(mut self) -> XlsResult<Vec<PivotTable>> {
        self.push_current()?;
        Ok(self.completed)
    }
}

fn ranges_overlap(left: &PivotViewDef, right: &PivotViewDef) -> bool {
    left.first_row <= right.last_row
        && right.first_row <= left.last_row
        && left.first_col <= right.last_col
        && right.first_col <= left.last_col
}

pub(crate) fn validate_pivot_cache_links(
    worksheets: &[crate::worksheet::XlsWorksheet],
    caches: &[PivotCache],
    cache_stream_ids: &[u16],
) -> XlsResult<()> {
    for worksheet in worksheets {
        for table in worksheet.pivot_tables() {
            let stream_id = *cache_stream_ids
                .get(usize::from(table.view.cache_index))
                .ok_or_else(|| {
                    cache_invalid(SXVIEW_TYPE, "SXVIEW global cache index is out of range")
                })?;
            let cache = caches
                .iter()
                .find(|cache| cache.stream_id() == stream_id)
                .ok_or_else(|| {
                    cache_invalid(
                        SXVIEW_TYPE,
                        "SXStreamID has no matching PivotCache storage stream",
                    )
                })?;
            if table.fields.len() != cache.fields().len() {
                return Err(cache_invalid(
                    SXVIEW_TYPE,
                    "SXVIEW field count does not match linked PivotCache",
                ));
            }
            validate_axis_fields(table, &table.row_fields, PivotAxis::Row)?;
            validate_axis_fields(table, &table.column_fields, PivotAxis::Column)?;
            for (index, field) in table.fields.iter().enumerate() {
                let cache_field = &cache.fields()[index];
                let visible_items = cache_field
                    .grouping()
                    .map(PivotCacheGrouping::group_items)
                    .unwrap_or(cache_field.items());
                for item in &field.items {
                    if item.item_type.code() == 0
                        && usize::from(item.cache_index) >= visible_items.len()
                    {
                        return Err(cache_invalid(
                            SXVI_TYPE,
                            "SXVI cache item ordinal is out of range",
                        ));
                    }
                }
                if let Some(extension) = &field.extension {
                    for ordinal in [
                        extension.auto_sort_data_index,
                        extension.auto_show_data_index,
                    ]
                    .into_iter()
                    .flatten()
                    {
                        if usize::from(ordinal) >= table.data_items.len() {
                            return Err(cache_invalid(
                                SXVDEX_TYPE,
                                "SXVDEx data-item ordinal is out of range",
                            ));
                        }
                    }
                }
            }
            for page in &table.page_entries {
                let field = table
                    .fields
                    .get(usize::from(page.field_index))
                    .ok_or_else(|| {
                        cache_invalid(SXPI_TYPE, "SXPI field ordinal is out of range")
                    })?;
                if field.axis != PivotAxis::Page
                    || !matches!(page.selection(), PivotPageSelection::All)
                        && usize::from(page.item_index) >= field.items.len()
                {
                    return Err(cache_invalid(
                        SXPI_TYPE,
                        "SXPI item ordinal or axis is invalid",
                    ));
                }
            }
            for item in &table.data_items {
                if usize::from(item.source_field_index) >= cache.fields().len() {
                    return Err(cache_invalid(
                        SXDI_TYPE,
                        "SXDI source field ordinal is out of range",
                    ));
                }
                if item.display_format != 0 {
                    let base = table
                        .fields
                        .get(usize::from(item.base_field_index))
                        .ok_or_else(|| {
                            cache_invalid(SXDI_TYPE, "SXDI base field ordinal is out of range")
                        })?;
                    if usize::from(item.base_item_index) >= base.items.len() {
                        return Err(cache_invalid(
                            SXDI_TYPE,
                            "SXDI base item ordinal is out of range",
                        ));
                    }
                }
            }
        }
    }
    Ok(())
}

fn validate_axis_fields(
    table: &PivotTable,
    fields: &[PivotAxisField],
    axis: PivotAxis,
) -> XlsResult<()> {
    let mut data_layout_seen = false;
    for entry in fields {
        match *entry {
            PivotAxisField::Field(index) => {
                let field = table.fields.get(usize::from(index)).ok_or_else(|| {
                    cache_invalid(SXIVD_TYPE, "SXIVD field ordinal is out of range")
                })?;
                if field.axis != axis {
                    return Err(cache_invalid(
                        SXIVD_TYPE,
                        "SXIVD field ordinal references the wrong axis",
                    ));
                }
            },
            PivotAxisField::DataLayout => {
                if data_layout_seen
                    || table.view.data_field_count <= 1
                    || table.view.data_axis != axis
                {
                    return Err(cache_invalid(
                        SXIVD_TYPE,
                        "invalid or duplicate SXIVD data-layout field",
                    ));
                }
                data_layout_seen = true;
            },
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// String helper
// ---------------------------------------------------------------------------

/// Read an XLUnicodeStringNoCch: 1-byte flags then `cch` chars.
fn read_xl_string_no_cch(data: &[u8], offset: &mut usize, cch: usize) -> XlsResult<String> {
    if cch == 0 {
        let end = offset
            .checked_add(1)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        data.get(*offset..end).ok_or(XlsError::InvalidLength {
            expected: end,
            found: data.len(),
        })?;
        *offset = end;
        return Ok(String::new());
    }

    if *offset >= data.len() {
        return Err(XlsError::InvalidLength {
            expected: *offset + 1,
            found: data.len(),
        });
    }

    let flags = data[*offset];
    *offset += 1;
    let is_utf16 = flags & 0x01 != 0;

    if is_utf16 {
        let byte_len = cch
            .checked_mul(2)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string size overflow"))?;
        let end = offset
            .checked_add(byte_len)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        if end > data.len() {
            return Err(XlsError::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let words: Vec<u16> = data[*offset..end]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        *offset = end;
        String::from_utf16(&words)
            .map_err(|e| XlsError::InvalidData(format!("Invalid UTF-16 in pivot string: {}", e)))
    } else {
        let end = offset
            .checked_add(cch)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        if end > data.len() {
            return Err(XlsError::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let s: String = data[*offset..end].iter().map(|&b| b as char).collect();
        *offset = end;
        Ok(s)
    }
}

#[cfg(test)]
mod worksheet_view_record_tests {
    use super::*;

    fn view_payload() -> Vec<u8> {
        let mut data = vec![0u8; 44];
        data[2..4].copy_from_slice(&1u16.to_le_bytes());
        data[6..8].copy_from_slice(&1u16.to_le_bytes());
        data[10..12].copy_from_slice(&1u16.to_le_bytes());
        data[12..14].copy_from_slice(&1u16.to_le_bytes());
        data[36..38].copy_from_slice(&0x020Bu16.to_le_bytes());
        data[38..40].copy_from_slice(&1u16.to_le_bytes());
        data[40..42].copy_from_slice(&1u16.to_le_bytes());
        data[42..44].copy_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&[0, b'P', 0, b'V']);
        data
    }

    fn sxex_payload() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        for _ in 0..3 {
            data.extend_from_slice(&u16::MAX.to_le_bytes());
        }
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0x004F_0200u32.to_le_bytes());
        for _ in 0..3 {
            data.extend_from_slice(&u16::MAX.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_current_sxview_layout_and_cache_index() {
        let mut payload = view_payload();
        payload[14..16].copy_from_slice(&7u16.to_le_bytes());
        let view = parse_sxview(&payload).unwrap();
        assert_eq!(view.cache_index, 7);
        assert_eq!(view.name, "P");
        assert_eq!(view.data_field_name, "V");
    }

    #[test]
    fn rejects_out_of_order_and_duplicate_singletons() {
        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        assert!(collector.feed_record(SXLI_TYPE, &[]).is_err());

        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        collector.feed_record(SXEX_TYPE, &sxex_payload()).unwrap();
        assert!(collector.feed_record(SXEX_TYPE, &sxex_payload()).is_err());
    }

    #[test]
    fn preserves_sxaddl_payload_exactly_and_rejects_bad_lengths() {
        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        collector.feed_record(SXEX_TYPE, &sxex_payload()).unwrap();
        let payload = [0x64, 0x08, 0, 0, 0, 2, 0xAA, 0xBB, 0xCC];
        collector.feed_record(SXADDL_TYPE, &payload).unwrap();
        let tables = collector.finish().unwrap();
        assert_eq!(
            tables[0].additional_extensions[0].payload,
            [0xAA, 0xBB, 0xCC]
        );
        assert!(parse_sxaddl(&payload[..5]).is_err());
        assert!(parse_sxivd(&[0]).is_err());
        assert!(parse_sxpi(&[0; 5]).is_err());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_sxvs() {
        let data = 0x0001u16.to_le_bytes();
        assert_eq!(parse_sxvs(&data).unwrap(), PivotSourceType::Worksheet);
    }

    #[test]
    fn test_parse_sxpi_two_entries() {
        let mut data = Vec::new();
        // Entry 1
        data.extend_from_slice(&0u16.to_le_bytes()); // isxvd
        data.extend_from_slice(&1u16.to_le_bytes()); // isxvi
        data.extend_from_slice(&0u16.to_le_bytes()); // idObj
        // Entry 2
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());

        let entries = parse_sxpi(&data).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].item_index, 1);
        assert_eq!(entries[1].field_index, 1);
    }

    #[test]
    fn test_parse_sxvd_no_name() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0001u16.to_le_bytes()); // axis = Row
        data.extend_from_slice(&0u16.to_le_bytes()); // cSub
        data.extend_from_slice(&0u16.to_le_bytes()); // grbitSub
        data.extend_from_slice(&5u16.to_le_bytes()); // cItm
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cchName = not present

        let field = parse_sxvd(&data).unwrap();
        assert_eq!(field.axis, PivotAxis::Row);
        assert_eq!(field.item_count, 5);
        assert!(field.name.is_none());
    }

    #[test]
    fn test_parse_sxvi_data_item() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x00FEu16.to_le_bytes()); // itmType = Data
        data.extend_from_slice(&0u16.to_le_bytes()); // flags
        data.extend_from_slice(&3u16.to_le_bytes()); // iCache
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // no name

        let item = parse_sxvi(&data).unwrap();
        assert_eq!(item.item_type, PivotItemType::Data);
        assert_eq!(item.cache_index, 3);
        assert!(item.name.is_none());
    }
}
