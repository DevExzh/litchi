//! Typed scalar criteria for BIFF `AutoFilter12` table records.
//!
//! This module only models the deterministic `ft = 0`, `cft = 0` form. Color,
//! font, icon, dynamic, date-grouping, worksheet, and producer-extension forms
//! remain opaque in the ListObject collector. Criteria are metadata only; they
//! are never evaluated by the reader or writer.

use super::list_object::{XlsListObjectId, XlsListObjectRange};
use super::{XlsError, XlsResult};

pub const AUTO_FILTER12_RECORD_TYPE: u16 = 0x087E;
pub(crate) const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087F;
const MAX_RECORD_PAYLOAD: usize = 8_224;
const MAX_AGGREGATE_BYTES: usize = 1_048_576;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn u16_at(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u16> {
    data.get(offset..offset + 2)
        .map(|value| u16::from_le_bytes([value[0], value[1]]))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn u32_at(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u32> {
    data.get(offset..offset + 4)
        .map(|value| u32::from_le_bytes(value.try_into().unwrap()))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn append_frt(out: &mut Vec<u8>, record_type: u16, range: XlsListObjectRange) {
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&1u16.to_le_bytes());
    out.extend_from_slice(&range.first_row().to_le_bytes());
    out.extend_from_slice(&range.last_row().to_le_bytes());
    out.extend_from_slice(&range.first_column().to_le_bytes());
    out.extend_from_slice(&range.last_column().to_le_bytes());
}

fn validate_frt(data: &[u8], record_type: u16, range: XlsListObjectRange) -> XlsResult<()> {
    if u16_at(data, 0, record_type, "frt.rt")? != record_type
        || u16_at(data, 2, record_type, "frt.flags")? != 1
        || u16_at(data, 4, record_type, "frt.rwFirst")? != range.first_row()
        || u16_at(data, 6, record_type, "frt.rwLast")? != range.last_row()
        || u16_at(data, 8, record_type, "frt.colFirst")? != range.first_column()
        || u16_at(data, 10, record_type, "frt.colLast")? != range.last_column()
    {
        return Err(invalid(
            record_type,
            "future-record range or type echo does not match the owning table",
        ));
    }
    Ok(())
}

fn record(record_type: u16, payload: Vec<u8>) -> XlsResult<Vec<u8>> {
    if payload.len() > MAX_RECORD_PAYLOAD {
        return Err(invalid(
            record_type,
            "payload exceeds the BIFF record limit",
        ));
    }
    let mut output = Vec::with_capacity(payload.len() + 4);
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    output.extend_from_slice(&payload);
    Ok(output)
}

/// Comparison operator stored in an `AFDOper` structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsAutoFilter12Operator {
    LessThan,
    Equal,
    LessThanOrEqual,
    GreaterThan,
    NotEqual,
    GreaterThanOrEqual,
}

impl XlsAutoFilter12Operator {
    const fn code(self) -> u8 {
        match self {
            Self::LessThan => 1,
            Self::Equal => 2,
            Self::LessThanOrEqual => 3,
            Self::GreaterThan => 4,
            Self::NotEqual => 5,
            Self::GreaterThanOrEqual => 6,
        }
    }

    fn from_code(value: u8) -> Option<Self> {
        Some(match value {
            1 => Self::LessThan,
            2 => Self::Equal,
            3 => Self::LessThanOrEqual,
            4 => Self::GreaterThan,
            5 => Self::NotEqual,
            6 => Self::GreaterThanOrEqual,
            _ => return None,
        })
    }
}

/// Scalar comparison value stored in an `AF12Criteria` structure.
#[derive(Debug, Clone)]
pub enum XlsAutoFilter12Value {
    Number(f64),
    String(String),
    Boolean(bool),
    Error(u8),
    Blanks,
    NonBlanks,
}

/// Icon-set discriminant stored in an `AF12CellIcon` filter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsAutoFilter12IconSet {
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

impl XlsAutoFilter12IconSet {
    const fn code(self) -> u32 {
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

    fn from_code(value: u32) -> Option<Self> {
        Some(match value {
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
            _ => return None,
        })
    }

    const fn expected_icon_bound(self) -> Option<u32> {
        match self {
            Self::NoIcon => None,
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => Some(3),
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => Some(4),
            Self::FiveArrows | Self::FiveArrowsGray | Self::FiveRating | Self::FiveQuarters => {
                Some(5)
            },
        }
    }
}

/// A validated `AF12CellIcon` selection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsAutoFilter12Icon {
    set: XlsAutoFilter12IconSet,
    index: u32,
}

/// Dynamic filter discriminator stored in `AutoFilter12.cft`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsAutoFilter12DynamicType {
    AboveAverage,
    BelowAverage,
    Tomorrow,
    Today,
    Yesterday,
    NextWeek,
    ThisWeek,
    LastWeek,
    NextMonth,
    ThisMonth,
    LastMonth,
    NextQuarter,
    ThisQuarter,
    LastQuarter,
    NextYear,
    ThisYear,
    LastYear,
    YearToDate,
    Quarter1,
    Quarter2,
    Quarter3,
    Quarter4,
    Month1,
    Month2,
    Month3,
    Month4,
    Month5,
    Month6,
    Month7,
    Month8,
    Month9,
    Month10,
    Month11,
    Month12,
}

impl XlsAutoFilter12DynamicType {
    const fn code(self) -> u32 {
        use XlsAutoFilter12DynamicType::*;
        match self {
            AboveAverage => 1,
            BelowAverage => 2,
            Tomorrow => 8,
            Today => 9,
            Yesterday => 10,
            NextWeek => 11,
            ThisWeek => 12,
            LastWeek => 13,
            NextMonth => 14,
            ThisMonth => 15,
            LastMonth => 16,
            NextQuarter => 17,
            ThisQuarter => 18,
            LastQuarter => 19,
            NextYear => 20,
            ThisYear => 21,
            LastYear => 22,
            YearToDate => 23,
            Quarter1 => 24,
            Quarter2 => 25,
            Quarter3 => 26,
            Quarter4 => 27,
            Month1 => 28,
            Month2 => 29,
            Month3 => 30,
            Month4 => 31,
            Month5 => 32,
            Month6 => 33,
            Month7 => 34,
            Month8 => 35,
            Month9 => 36,
            Month10 => 37,
            Month11 => 38,
            Month12 => 39,
        }
    }

    fn from_code(value: u32) -> Option<Self> {
        use XlsAutoFilter12DynamicType::*;
        Some(match value {
            1 => AboveAverage,
            2 => BelowAverage,
            8 => Tomorrow,
            9 => Today,
            10 => Yesterday,
            11 => NextWeek,
            12 => ThisWeek,
            13 => LastWeek,
            14 => NextMonth,
            15 => ThisMonth,
            16 => LastMonth,
            17 => NextQuarter,
            18 => ThisQuarter,
            19 => LastQuarter,
            20 => NextYear,
            21 => ThisYear,
            22 => LastYear,
            23 => YearToDate,
            24 => Quarter1,
            25 => Quarter2,
            26 => Quarter3,
            27 => Quarter4,
            28 => Month1,
            29 => Month2,
            30 => Month3,
            31 => Month4,
            32 => Month5,
            33 => Month6,
            34 => Month7,
            35 => Month8,
            36 => Month9,
            37 => Month10,
            38 => Month11,
            39 => Month12,
            _ => return None,
        })
    }
}

/// Granularity selected by an `AF12DateInfo` grouping.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsAutoFilter12DateLevel {
    Year,
    Month,
    Day,
    Hour,
    Minute,
    Second,
}

impl XlsAutoFilter12DateLevel {
    const fn code(self) -> u32 {
        match self {
            Self::Year => 0,
            Self::Month => 1,
            Self::Day => 2,
            Self::Hour => 3,
            Self::Minute => 4,
            Self::Second => 5,
        }
    }
    fn from_code(value: u32) -> Option<Self> {
        Some(match value {
            0 => Self::Year,
            1 => Self::Month,
            2 => Self::Day,
            3 => Self::Hour,
            4 => Self::Minute,
            5 => Self::Second,
            _ => return None,
        })
    }
}

/// One validated `AF12DateInfo` continuation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsAutoFilter12DateGroup {
    year: u16,
    month: u16,
    day: u32,
    hour: u16,
    minute: u16,
    second: u16,
    level: XlsAutoFilter12DateLevel,
}

impl XlsAutoFilter12DateGroup {
    pub fn try_new(
        year: u16,
        month: u16,
        day: u32,
        hour: u16,
        minute: u16,
        second: u16,
        level: XlsAutoFilter12DateLevel,
    ) -> XlsResult<Self> {
        if !(1..=12).contains(&month) || hour > 23 || minute > 59 || second > 59 {
            return Err(invalid(
                CONTINUE_FRT12_RECORD_TYPE,
                "date grouping component is out of range",
            ));
        }
        let leap =
            year.is_multiple_of(4) && (!year.is_multiple_of(100) || year.is_multiple_of(400));
        let days = match month {
            2 if leap => 29,
            2 => 28,
            4 | 6 | 9 | 11 => 30,
            _ => 31,
        };
        if day == 0 || day > days {
            return Err(invalid(
                CONTINUE_FRT12_RECORD_TYPE,
                "date grouping day is invalid for its month",
            ));
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
            level,
        })
    }
    pub const fn year(self) -> u16 {
        self.year
    }
    pub const fn month(self) -> u16 {
        self.month
    }
    pub const fn day(self) -> u32 {
        self.day
    }
    pub const fn hour(self) -> u16 {
        self.hour
    }
    pub const fn minute(self) -> u16 {
        self.minute
    }
    pub const fn second(self) -> u16 {
        self.second
    }
    pub const fn level(self) -> XlsAutoFilter12DateLevel {
        self.level
    }
}

/// Formatting property selected by a `DXFN12NoCB` filter payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsAutoFilter12FormatKind {
    CellColor,
    FontColor,
}

/// Bounded serialized `DXFN12NoCB` metadata. It is passive formatting data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsAutoFilter12DifferentialFormat(Vec<u8>);

impl XlsAutoFilter12DifferentialFormat {
    pub fn try_new(bytes: Vec<u8>) -> XlsResult<Self> {
        if bytes.is_empty() || bytes.len() > MAX_RECORD_PAYLOAD - 60 {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "DXFN12NoCB payload length is invalid",
            ));
        }
        Ok(Self(bytes))
    }
    pub fn bytes(&self) -> &[u8] {
        &self.0
    }
}

impl XlsAutoFilter12Icon {
    pub fn try_new(set: XlsAutoFilter12IconSet, index: u32) -> XlsResult<Self> {
        let valid = match set.expected_icon_bound() {
            None => index == u32::MAX,
            Some(bound) => index < bound,
        };
        if !valid {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "icon index is outside its icon-set cardinality",
            ));
        }
        Ok(Self { set, index })
    }

    pub const fn set(self) -> XlsAutoFilter12IconSet {
        self.set
    }
    pub const fn index(self) -> u32 {
        self.index
    }
}

impl PartialEq for XlsAutoFilter12Value {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Number(left), Self::Number(right)) => left.to_bits() == right.to_bits(),
            (Self::String(left), Self::String(right)) => left == right,
            (Self::Boolean(left), Self::Boolean(right)) => left == right,
            (Self::Error(left), Self::Error(right)) => left == right,
            (Self::Blanks, Self::Blanks) | (Self::NonBlanks, Self::NonBlanks) => true,
            _ => false,
        }
    }
}

impl Eq for XlsAutoFilter12Value {}

/// One inert comparison criterion for a table column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsAutoFilter12Criterion {
    operator: XlsAutoFilter12Operator,
    value: XlsAutoFilter12Value,
}

impl XlsAutoFilter12Criterion {
    pub fn try_new(
        operator: XlsAutoFilter12Operator,
        value: XlsAutoFilter12Value,
    ) -> XlsResult<Self> {
        let criterion = Self { operator, value };
        criterion.validate()?;
        Ok(criterion)
    }

    pub const fn operator(&self) -> XlsAutoFilter12Operator {
        self.operator
    }
    pub const fn value(&self) -> &XlsAutoFilter12Value {
        &self.value
    }

    fn validate(&self) -> XlsResult<()> {
        match &self.value {
            XlsAutoFilter12Value::String(value) => {
                let count = value.encode_utf16().count();
                if !(1..=255).contains(&count) {
                    return Err(invalid(
                        AUTO_FILTER12_RECORD_TYPE,
                        "criterion string must contain 1 to 255 UTF-16 code units",
                    ));
                }
            },
            XlsAutoFilter12Value::Error(value)
                if !matches!(
                    *value,
                    0x00 | 0x07 | 0x0F | 0x17 | 0x1D | 0x24 | 0x2A | 0x2B
                ) =>
            {
                return Err(invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "criterion contains an invalid BIFF error value",
                ));
            },
            _ => {},
        }
        Ok(())
    }
}

/// Typed, table-owned `AutoFilter12` metadata for one relative table column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsTableAutoFilter12 {
    column_index: u16,
    hide_arrow: bool,
    criteria: Vec<XlsAutoFilter12Criterion>,
    date_groupings: Vec<XlsAutoFilter12DateGroup>,
    dynamic_filter: Option<XlsAutoFilter12DynamicType>,
    icon: Option<XlsAutoFilter12Icon>,
    format_kind: Option<XlsAutoFilter12FormatKind>,
    differential_format: Option<XlsAutoFilter12DifferentialFormat>,
}

impl XlsTableAutoFilter12 {
    pub fn try_new(column_index: u16, criteria: Vec<XlsAutoFilter12Criterion>) -> XlsResult<Self> {
        if criteria.is_empty() {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "typed AutoFilter12 requires at least one criterion",
            ));
        }
        let value = Self {
            column_index,
            hide_arrow: false,
            criteria,
            date_groupings: Vec::new(),
            dynamic_filter: None,
            icon: None,
            format_kind: None,
            differential_format: None,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn try_new_icon(column_index: u16, icon: XlsAutoFilter12Icon) -> XlsResult<Self> {
        let value = Self {
            column_index,
            hide_arrow: false,
            criteria: Vec::new(),
            date_groupings: Vec::new(),
            dynamic_filter: None,
            icon: Some(icon),
            format_kind: None,
            differential_format: None,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn try_new_date_groups(
        column_index: u16,
        date_groupings: Vec<XlsAutoFilter12DateGroup>,
    ) -> XlsResult<Self> {
        let value = Self {
            column_index,
            hide_arrow: false,
            criteria: Vec::new(),
            date_groupings,
            dynamic_filter: None,
            icon: None,
            format_kind: None,
            differential_format: None,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn try_new_dynamic(
        column_index: u16,
        dynamic_filter: XlsAutoFilter12DynamicType,
        criteria: Vec<XlsAutoFilter12Criterion>,
    ) -> XlsResult<Self> {
        let value = Self {
            column_index,
            hide_arrow: false,
            criteria,
            date_groupings: Vec::new(),
            dynamic_filter: Some(dynamic_filter),
            icon: None,
            format_kind: None,
            differential_format: None,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn try_new_format(
        column_index: u16,
        kind: XlsAutoFilter12FormatKind,
        format: XlsAutoFilter12DifferentialFormat,
    ) -> XlsResult<Self> {
        let value = Self {
            column_index,
            hide_arrow: false,
            criteria: Vec::new(),
            date_groupings: Vec::new(),
            dynamic_filter: None,
            icon: None,
            format_kind: Some(kind),
            differential_format: Some(format),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn with_hidden_arrow(mut self, hide: bool) -> Self {
        self.hide_arrow = hide;
        self
    }

    pub const fn column_index(&self) -> u16 {
        self.column_index
    }
    pub const fn hides_arrow(&self) -> bool {
        self.hide_arrow
    }
    pub fn criteria(&self) -> &[XlsAutoFilter12Criterion] {
        &self.criteria
    }
    pub fn date_groupings(&self) -> &[XlsAutoFilter12DateGroup] {
        &self.date_groupings
    }
    pub const fn dynamic_filter(&self) -> Option<XlsAutoFilter12DynamicType> {
        self.dynamic_filter
    }
    pub const fn icon_filter(&self) -> Option<XlsAutoFilter12Icon> {
        self.icon
    }
    pub const fn format_kind(&self) -> Option<XlsAutoFilter12FormatKind> {
        self.format_kind
    }
    pub fn differential_format(&self) -> Option<&XlsAutoFilter12DifferentialFormat> {
        self.differential_format.as_ref()
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
        if self.column_index > 255 {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "AutoFilter12 column index exceeds 255",
            ));
        }
        if self.format_kind.is_some() != self.differential_format.is_some() {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "format filter kind and DXFN12NoCB payload must coexist",
            ));
        }
        let comparison = !self.criteria.is_empty() || !self.date_groupings.is_empty();
        let groups = usize::from(comparison)
            + usize::from(self.icon.is_some())
            + usize::from(self.differential_format.is_some());
        if groups != 1 || self.dynamic_filter.is_some() && !comparison {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "AutoFilter12 must contain exactly one compatible filter group",
            ));
        }
        for criterion in &self.criteria {
            criterion.validate()?;
        }
        let bytes = self
            .criteria
            .iter()
            .try_fold(60usize, |size, criterion| {
                let extra = match &criterion.value {
                    XlsAutoFilter12Value::String(value) => 23 + value.encode_utf16().count() * 2,
                    _ => 22,
                };
                size.checked_add(extra)
                    .ok_or_else(|| invalid(AUTO_FILTER12_RECORD_TYPE, "criteria size overflows"))
            })?
            .checked_add(self.date_groupings.len().checked_mul(36).ok_or_else(|| {
                invalid(AUTO_FILTER12_RECORD_TYPE, "date grouping size overflows")
            })?)
            .ok_or_else(|| invalid(AUTO_FILTER12_RECORD_TYPE, "AutoFilter12 size overflows"))?;
        if bytes > MAX_AGGREGATE_BYTES {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "AutoFilter12 criteria exceed the aggregate resource bound",
            ));
        }
        Ok(())
    }
}

fn encode_date_group(
    value: XlsAutoFilter12DateGroup,
    range: XlsListObjectRange,
) -> XlsResult<Vec<u8>> {
    XlsAutoFilter12DateGroup::try_new(
        value.year,
        value.month,
        value.day,
        value.hour,
        value.minute,
        value.second,
        value.level,
    )?;
    let mut payload = Vec::with_capacity(36);
    append_frt(&mut payload, CONTINUE_FRT12_RECORD_TYPE, range);
    payload.extend_from_slice(&value.year.to_le_bytes());
    payload.extend_from_slice(&value.month.to_le_bytes());
    payload.extend_from_slice(&value.day.to_le_bytes());
    payload.extend_from_slice(&value.hour.to_le_bytes());
    payload.extend_from_slice(&value.minute.to_le_bytes());
    payload.extend_from_slice(&value.second.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&value.level.code().to_le_bytes());
    Ok(payload)
}

fn parse_date_group(data: &[u8], range: XlsListObjectRange) -> XlsResult<XlsAutoFilter12DateGroup> {
    validate_frt(data, CONTINUE_FRT12_RECORD_TYPE, range)?;
    if data.len() != 36 || data[28..32].iter().any(|byte| *byte != 0) {
        return Err(invalid(
            CONTINUE_FRT12_RECORD_TYPE,
            "invalid AF12DateInfo length or reserved bytes",
        ));
    }
    XlsAutoFilter12DateGroup::try_new(
        u16_at(data, 12, CONTINUE_FRT12_RECORD_TYPE, "year")?,
        u16_at(data, 14, CONTINUE_FRT12_RECORD_TYPE, "month")?,
        u32_at(data, 16, CONTINUE_FRT12_RECORD_TYPE, "day")?,
        u16_at(data, 20, CONTINUE_FRT12_RECORD_TYPE, "hour")?,
        u16_at(data, 22, CONTINUE_FRT12_RECORD_TYPE, "minute")?,
        u16_at(data, 24, CONTINUE_FRT12_RECORD_TYPE, "second")?,
        XlsAutoFilter12DateLevel::from_code(u32_at(
            data,
            32,
            CONTINUE_FRT12_RECORD_TYPE,
            "nodeType",
        )?)
        .ok_or_else(|| invalid(CONTINUE_FRT12_RECORD_TYPE, "reserved AF12DateInfo nodeType"))?,
    )
}

fn encode_criterion(
    criterion: &XlsAutoFilter12Criterion,
    range: XlsListObjectRange,
) -> XlsResult<Vec<u8>> {
    criterion.validate()?;
    let mut payload = Vec::new();
    append_frt(&mut payload, CONTINUE_FRT12_RECORD_TYPE, range);
    let mut doper = [0u8; 10];
    doper[1] = criterion.operator.code();
    let trailing = match &criterion.value {
        XlsAutoFilter12Value::Number(value) => {
            doper[0] = 0x04;
            doper[2..10].copy_from_slice(&value.to_le_bytes());
            None
        },
        XlsAutoFilter12Value::String(value) => {
            let units = value.encode_utf16().collect::<Vec<_>>();
            doper[0] = 0x06;
            doper[2] = units.len() as u8;
            doper[3] = u8::from(!value.contains(['?', '*']));
            Some(units)
        },
        XlsAutoFilter12Value::Boolean(value) => {
            doper[0] = 0x08;
            doper[2] = u8::from(*value);
            None
        },
        XlsAutoFilter12Value::Error(value) => {
            doper[0] = 0x08;
            doper[2] = *value;
            doper[3] = 1;
            None
        },
        XlsAutoFilter12Value::Blanks => {
            doper[0] = 0x0C;
            None
        },
        XlsAutoFilter12Value::NonBlanks => {
            doper[0] = 0x0E;
            None
        },
    };
    payload.extend_from_slice(&doper);
    if let Some(units) = trailing {
        payload.push(1);
        payload.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
    Ok(payload)
}

pub(crate) fn write_table_autofilter12(
    filter: &XlsTableAutoFilter12,
    range: XlsListObjectRange,
    table_id: XlsListObjectId,
) -> XlsResult<Vec<Vec<u8>>> {
    filter.validate()?;
    if usize::from(filter.column_index) >= range.column_count() {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "AutoFilter12 column is outside the owning table",
        ));
    }
    let rgb_len = filter
        .differential_format
        .as_ref()
        .map_or(0, |value| value.bytes().len());
    let mut base = Vec::with_capacity(if filter.icon.is_some() {
        68
    } else {
        60 + rgb_len
    });
    append_frt(&mut base, AUTO_FILTER12_RECORD_TYPE, range);
    base.extend_from_slice(&filter.column_index.to_le_bytes());
    base.extend_from_slice(&u32::from(filter.hide_arrow).to_le_bytes());
    let ft: u32 = if filter.icon.is_some() {
        3
    } else {
        match filter.format_kind {
            Some(XlsAutoFilter12FormatKind::CellColor) => 1,
            Some(XlsAutoFilter12FormatKind::FontColor) => 2,
            None => 0,
        }
    };
    base.extend_from_slice(&ft.to_le_bytes());
    base.extend_from_slice(
        &filter
            .dynamic_filter
            .map_or(0, XlsAutoFilter12DynamicType::code)
            .to_le_bytes(),
    );
    base.extend_from_slice(&(filter.criteria.len() as u32).to_le_bytes());
    base.extend_from_slice(&(filter.date_groupings.len() as u32).to_le_bytes());
    base.extend_from_slice(&0u16.to_le_bytes()); // table ownership
    base.extend_from_slice(&0u32.to_le_bytes());
    base.extend_from_slice(&table_id.value().to_le_bytes());
    base.extend_from_slice(&[0; 16]);
    if let Some(icon) = filter.icon {
        base.extend_from_slice(&icon.set.code().to_le_bytes());
        base.extend_from_slice(&icon.index.to_le_bytes());
    } else if let Some(format) = &filter.differential_format {
        base.extend_from_slice(format.bytes());
    }
    let mut records = vec![record(AUTO_FILTER12_RECORD_TYPE, base)?];
    for criterion in &filter.criteria {
        records.push(record(
            CONTINUE_FRT12_RECORD_TYPE,
            encode_criterion(criterion, range)?,
        )?);
    }
    for grouping in &filter.date_groupings {
        records.push(record(
            CONTINUE_FRT12_RECORD_TYPE,
            encode_date_group(*grouping, range)?,
        )?);
    }
    Ok(records)
}

fn parse_criterion(data: &[u8], range: XlsListObjectRange) -> XlsResult<XlsAutoFilter12Criterion> {
    validate_frt(data, CONTINUE_FRT12_RECORD_TYPE, range)?;
    if data.len() < 22 {
        return Err(invalid(
            CONTINUE_FRT12_RECORD_TYPE,
            "truncated AF12Criteria",
        ));
    }
    let operator = XlsAutoFilter12Operator::from_code(data[13]).ok_or_else(|| {
        invalid(
            CONTINUE_FRT12_RECORD_TYPE,
            "invalid AFDOper comparison operator",
        )
    })?;
    let value = match data[12] {
        0x04 => {
            if data.len() != 22 {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "numeric criterion has trailing data",
                ));
            }
            XlsAutoFilter12Value::Number(f64::from_le_bytes(data[14..22].try_into().unwrap()))
        },
        0x06 => {
            if data[16] != 0 || data[15] > 1 {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "invalid string criterion flags",
                ));
            }
            let count = usize::from(data[14]);
            if count == 0 {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "empty string criterion",
                ));
            }
            let flags = *data
                .get(22)
                .ok_or_else(|| invalid(CONTINUE_FRT12_RECORD_TYPE, "missing criterion string"))?;
            if flags & !1 != 0 {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "unsupported criterion string flags",
                ));
            }
            let width = if flags == 0 { 1 } else { 2 };
            let end = 23usize
                .checked_add(count.checked_mul(width).ok_or_else(|| {
                    invalid(
                        CONTINUE_FRT12_RECORD_TYPE,
                        "criterion string length overflows",
                    )
                })?)
                .ok_or_else(|| {
                    invalid(
                        CONTINUE_FRT12_RECORD_TYPE,
                        "criterion string length overflows",
                    )
                })?;
            if end != data.len() {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "criterion string length mismatch",
                ));
            }
            let value = if width == 1 {
                data[23..].iter().map(|byte| char::from(*byte)).collect()
            } else {
                char::decode_utf16(
                    data[23..]
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
                )
                .collect::<Result<String, _>>()
                .map_err(|_| {
                    invalid(
                        CONTINUE_FRT12_RECORD_TYPE,
                        "criterion string is invalid UTF-16",
                    )
                })?
            };
            XlsAutoFilter12Value::String(value)
        },
        0x08 => {
            if data.len() != 22 {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "invalid Boolean/error criterion length",
                ));
            }
            match data[15] {
                0 if data[14] <= 1 => XlsAutoFilter12Value::Boolean(data[14] != 0),
                1 => XlsAutoFilter12Value::Error(data[14]),
                _ => {
                    return Err(invalid(
                        CONTINUE_FRT12_RECORD_TYPE,
                        "invalid Boolean/error criterion",
                    ));
                },
            }
        },
        0x0C | 0x0E => {
            if data.len() != 22 || data[14..22].iter().any(|byte| *byte != 0) {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "blank criterion reserved bytes are nonzero",
                ));
            }
            if data[12] == 0x0C {
                XlsAutoFilter12Value::Blanks
            } else {
                XlsAutoFilter12Value::NonBlanks
            }
        },
        0x02 => {
            return Err(invalid(
                CONTINUE_FRT12_RECORD_TYPE,
                "RK criteria are forbidden in AutoFilter12",
            ));
        },
        _ => {
            return Err(invalid(
                CONTINUE_FRT12_RECORD_TYPE,
                "unsupported AFDOper value type",
            ));
        },
    };
    XlsAutoFilter12Criterion::try_new(operator, value)
}

/// Returns `None` for a structurally owned but unsupported extension form.
pub(crate) fn parse_table_autofilter12(
    base: &[u8],
    continuations: &[Vec<u8>],
    range: XlsListObjectRange,
    table_id: XlsListObjectId,
) -> XlsResult<Option<XlsTableAutoFilter12>> {
    if !(60..=MAX_RECORD_PAYLOAD).contains(&base.len()) {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "invalid AutoFilter12 base length",
        ));
    }
    validate_frt(base, AUTO_FILTER12_RECORD_TYPE, range)?;
    let column_index = u16_at(base, 12, AUTO_FILTER12_RECORD_TYPE, "iEntry")?;
    if column_index > 255 || usize::from(column_index) >= range.column_count() {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "AutoFilter12 column is outside the owning table",
        ));
    }
    let hide_arrow = match u32_at(base, 14, AUTO_FILTER12_RECORD_TYPE, "fHideArrow")? {
        0 => false,
        1 => true,
        _ => {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "fHideArrow is not Boolean",
            ));
        },
    };
    let ft = u32_at(base, 18, AUTO_FILTER12_RECORD_TYPE, "ft")?;
    if ft > 3 {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "reserved AutoFilter12 filter type",
        ));
    }
    let cft_code = u32_at(base, 22, AUTO_FILTER12_RECORD_TYPE, "cft")?;
    let dynamic_filter = if cft_code == 0 {
        None
    } else {
        Some(
            XlsAutoFilter12DynamicType::from_code(cft_code).ok_or_else(|| {
                invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "reserved AutoFilter12 custom filter type",
                )
            })?,
        )
    };
    let criterion_count = u32_at(base, 26, AUTO_FILTER12_RECORD_TYPE, "cCriteria")? as usize;
    let date_count = u32_at(base, 30, AUTO_FILTER12_RECORD_TYPE, "cDateGroupings")? as usize;
    let ownership = u16_at(base, 34, AUTO_FILTER12_RECORD_TYPE, "ownership flags")?;
    if ownership & 0x0007 != 0
        || ownership & 0x0008 != 0
        || u32_at(base, 40, AUTO_FILTER12_RECORD_TYPE, "idList")? != table_id.value()
    {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "AutoFilter12 is not owned by the attached table",
        ));
    }
    let aggregate = base
        .len()
        .checked_add(continuations.iter().map(Vec::len).sum::<usize>())
        .ok_or_else(|| invalid(AUTO_FILTER12_RECORD_TYPE, "AutoFilter12 size overflows"))?;
    if aggregate > MAX_AGGREGATE_BYTES {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "AutoFilter12 exceeds the aggregate resource bound",
        ));
    }
    if ft == 0 {
        let expected = criterion_count.checked_add(date_count).ok_or_else(|| {
            invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "AutoFilter12 item count overflows",
            )
        })?;
        if expected == 0 || expected != continuations.len() {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "AutoFilter12 continuation count does not match its criteria/date counts",
            ));
        }
        for continuation in continuations {
            if !(12..=MAX_RECORD_PAYLOAD).contains(&continuation.len()) {
                return Err(invalid(
                    CONTINUE_FRT12_RECORD_TYPE,
                    "invalid AutoFilter12 continuation length",
                ));
            }
            validate_frt(continuation, CONTINUE_FRT12_RECORD_TYPE, range)?;
        }
    }
    if ft == 3 && dynamic_filter.is_none() {
        if base.len() != 68 || !continuations.is_empty() {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "icon AutoFilter12 must contain exactly one eight-byte AF12CellIcon payload",
            ));
        }
        let set = XlsAutoFilter12IconSet::from_code(u32_at(
            base,
            60,
            AUTO_FILTER12_RECORD_TYPE,
            "iIconSet",
        )?)
        .ok_or_else(|| {
            invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "reserved AF12CellIcon icon-set value",
            )
        })?;
        let icon = XlsAutoFilter12Icon::try_new(
            set,
            u32_at(base, 64, AUTO_FILTER12_RECORD_TYPE, "iIcon")?,
        )?;
        return Ok(Some(XlsTableAutoFilter12 {
            column_index,
            hide_arrow,
            criteria: Vec::new(),
            date_groupings: Vec::new(),
            dynamic_filter: None,
            icon: Some(icon),
            format_kind: None,
            differential_format: None,
        }));
    }
    if matches!(ft, 1 | 2) {
        if dynamic_filter.is_some() || !continuations.is_empty() || base.len() == 60 {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "format AutoFilter12 has incompatible metadata",
            ));
        }
        let format = XlsAutoFilter12DifferentialFormat::try_new(base[60..].to_vec())?;
        let format_kind = Some(if ft == 1 {
            XlsAutoFilter12FormatKind::CellColor
        } else {
            XlsAutoFilter12FormatKind::FontColor
        });
        return Ok(Some(XlsTableAutoFilter12 {
            column_index,
            hide_arrow,
            criteria: Vec::new(),
            date_groupings: Vec::new(),
            dynamic_filter: None,
            icon: None,
            format_kind,
            differential_format: Some(format),
        }));
    }
    if ft != 0 || base.len() != 60 {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "invalid AutoFilter12 filter payload",
        ));
    }
    let criteria = continuations[..criterion_count]
        .iter()
        .map(|item| parse_criterion(item, range))
        .collect::<XlsResult<Vec<_>>>()?;
    let date_groupings = continuations[criterion_count..]
        .iter()
        .map(|item| parse_date_group(item, range))
        .collect::<XlsResult<Vec<_>>>()?;
    let value = XlsTableAutoFilter12 {
        column_index,
        hide_arrow,
        criteria,
        date_groupings,
        dynamic_filter,
        icon: None,
        format_kind: None,
        differential_format: None,
    };
    value.validate()?;
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn range() -> XlsListObjectRange {
        XlsListObjectRange::try_new(1, 9, 2, 4).unwrap()
    }

    #[test]
    fn scalar_criteria_round_trip_and_reject_hostile_lengths() {
        let criteria = vec![
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::GreaterThan,
                XlsAutoFilter12Value::Number(7.5),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::Equal,
                XlsAutoFilter12Value::String("A*".into()),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::NotEqual,
                XlsAutoFilter12Value::Error(0x2A),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::Equal,
                XlsAutoFilter12Value::Boolean(true),
            )
            .unwrap(),
        ];
        let expected = XlsTableAutoFilter12::try_new(1, criteria)
            .unwrap()
            .with_hidden_arrow(true);
        let records =
            write_table_autofilter12(&expected, range(), XlsListObjectId::try_new(9).unwrap())
                .unwrap();
        let payload = |record: &[u8]| record[4..].to_vec();
        let base = payload(&records[0]);
        let continuations = records[1..]
            .iter()
            .map(|item| payload(item))
            .collect::<Vec<_>>();
        assert_eq!(
            parse_table_autofilter12(
                &base,
                &continuations,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .unwrap(),
            Some(expected.clone())
        );

        let mut wrong_owner = base.clone();
        wrong_owner[40..44].copy_from_slice(&10u32.to_le_bytes());
        assert!(
            parse_table_autofilter12(
                &wrong_owner,
                &continuations,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_err()
        );
        assert!(
            parse_table_autofilter12(
                &base,
                &continuations[..2],
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_err()
        );
        let mut wrong_range = continuations.clone();
        wrong_range[0][4..6].copy_from_slice(&0u16.to_le_bytes());
        assert!(
            parse_table_autofilter12(
                &base,
                &wrong_range,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_err()
        );

        let mut producer_padding = continuations.clone();
        producer_padding[3][16..22].copy_from_slice(&[1, 2, 3, 4, 5, 6]);
        assert_eq!(
            parse_table_autofilter12(
                &base,
                &producer_padding,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .unwrap(),
            Some(expected),
        );
    }

    #[test]
    fn icon_filter_round_trips_and_rejects_reserved_sets_indices_and_continuations() {
        assert!(
            XlsAutoFilter12Icon::try_new(XlsAutoFilter12IconSet::ThreeTrafficLights1, 3).is_err()
        );
        assert!(XlsAutoFilter12Icon::try_new(XlsAutoFilter12IconSet::NoIcon, 0).is_err());
        let expected = XlsTableAutoFilter12::try_new_icon(
            2,
            XlsAutoFilter12Icon::try_new(XlsAutoFilter12IconSet::FiveQuarters, 4).unwrap(),
        )
        .unwrap();
        let records =
            write_table_autofilter12(&expected, range(), XlsListObjectId::try_new(9).unwrap())
                .unwrap();
        assert_eq!(records.len(), 1);
        let mut base = records[0][4..].to_vec();
        assert_eq!(
            parse_table_autofilter12(&base, &[], range(), XlsListObjectId::try_new(9).unwrap())
                .unwrap(),
            Some(expected)
        );
        base[60..64].copy_from_slice(&17u32.to_le_bytes());
        assert!(
            parse_table_autofilter12(&base, &[], range(), XlsListObjectId::try_new(9).unwrap())
                .is_err()
        );
        base[60..64].copy_from_slice(&XlsAutoFilter12IconSet::FiveQuarters.code().to_le_bytes());
        base[64..68].copy_from_slice(&5u32.to_le_bytes());
        assert!(
            parse_table_autofilter12(&base, &[], range(), XlsListObjectId::try_new(9).unwrap())
                .is_err()
        );
        assert!(
            parse_table_autofilter12(
                &records[0][4..],
                &[vec![0; 12]],
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_err()
        );
    }

    #[test]
    fn dynamic_date_and_format_filters_round_trip_typed() {
        let date = XlsAutoFilter12DateGroup::try_new(
            2024,
            2,
            29,
            23,
            59,
            58,
            XlsAutoFilter12DateLevel::Second,
        )
        .unwrap();
        let expected = XlsTableAutoFilter12::try_new_date_groups(0, vec![date]).unwrap();
        let records =
            write_table_autofilter12(&expected, range(), XlsListObjectId::try_new(9).unwrap())
                .unwrap();
        let base = records[0][4..].to_vec();
        let continuations = records[1..]
            .iter()
            .map(|item| item[4..].to_vec())
            .collect::<Vec<_>>();
        assert_eq!(
            parse_table_autofilter12(
                &base,
                &continuations,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .unwrap(),
            Some(expected)
        );
        let mut unused = continuations.clone();
        unused[0][26] = 0xa5;
        assert!(
            parse_table_autofilter12(
                &base,
                &unused,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_ok()
        );
        let mut reserved = continuations;
        reserved[0][28] = 1;
        assert!(
            parse_table_autofilter12(
                &base,
                &reserved,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .is_err()
        );
        assert!(
            XlsAutoFilter12DateGroup::try_new(2023, 2, 29, 0, 0, 0, XlsAutoFilter12DateLevel::Day)
                .is_err()
        );

        let criterion = XlsAutoFilter12Criterion::try_new(
            XlsAutoFilter12Operator::GreaterThan,
            XlsAutoFilter12Value::Number(1.0),
        )
        .unwrap();
        let expected = XlsTableAutoFilter12::try_new_dynamic(
            1,
            XlsAutoFilter12DynamicType::AboveAverage,
            vec![criterion],
        )
        .unwrap();
        let records =
            write_table_autofilter12(&expected, range(), XlsListObjectId::try_new(9).unwrap())
                .unwrap();
        let base = records[0][4..].to_vec();
        let continuations = records[1..]
            .iter()
            .map(|item| item[4..].to_vec())
            .collect::<Vec<_>>();
        assert_eq!(
            parse_table_autofilter12(
                &base,
                &continuations,
                range(),
                XlsListObjectId::try_new(9).unwrap()
            )
            .unwrap(),
            Some(expected)
        );

        for kind in [
            XlsAutoFilter12FormatKind::CellColor,
            XlsAutoFilter12FormatKind::FontColor,
        ] {
            let expected = XlsTableAutoFilter12::try_new_format(
                2,
                kind,
                XlsAutoFilter12DifferentialFormat::try_new(vec![1, 2, 3, 4]).unwrap(),
            )
            .unwrap();
            let records =
                write_table_autofilter12(&expected, range(), XlsListObjectId::try_new(9).unwrap())
                    .unwrap();
            assert_eq!(
                parse_table_autofilter12(
                    &records[0][4..],
                    &[],
                    range(),
                    XlsListObjectId::try_new(9).unwrap()
                )
                .unwrap(),
                Some(expected)
            );
        }
    }
}
