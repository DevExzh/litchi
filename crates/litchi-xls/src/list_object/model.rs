//! Semantic BIFF8 table values, source metadata, and validation.

use super::codec::{append_string, parse_string, u32_at};
use super::{
    AUTO_FILTER12_RECORD_TYPE, FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, LIST12_RECORD_TYPE,
    MAX_FEATURE_BYTES, invalid,
};
use crate::Result;
use crate::autofilter12::TableAutoFilter12;
use std::collections::HashSet;

pub(super) fn validate_name(value: &str, field: &str) -> Result<()> {
    if !(1..=255).contains(&value.encode_utf16().count())
        || value
            .chars()
            .any(|c| c <= '\u{1f}' || matches!(c, '\u{fffe}' | '\u{ffff}'))
    {
        return Err(invalid(FEATURE11_RECORD_TYPE, format!("invalid {field}")));
    }
    Ok(())
}
pub(super) fn validate_table_name(value: &str) -> Result<()> {
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
pub(super) fn validate_column_name(value: &str) -> Result<()> {
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ListObjectId(u32);
impl ListObjectId {
    pub fn try_new(value: u32) -> Result<Self> {
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
pub struct ListColumnId(u32);
impl ListColumnId {
    pub fn try_new(value: u32) -> Result<Self> {
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
pub struct ListObjectRange {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u16,
    pub(super) last_column: u16,
}
impl ListObjectRange {
    pub fn try_new(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> Result<Self> {
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
pub enum ListTotalAggregation {
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
impl ListTotalAggregation {
    pub(super) fn code(self) -> u32 {
        self as u32
    }
    pub(super) fn from_code(v: u32) -> Result<Self> {
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
pub struct ListObjectColumn {
    pub(super) id: ListColumnId,
    pub(super) name: String,
    pub(super) aggregation: ListTotalAggregation,
    pub(super) total_formula: Option<Vec<u8>>,
    pub(super) total_string: Option<String>,
}
impl ListObjectColumn {
    pub fn try_new(id: ListColumnId, name: impl Into<String>) -> Result<Self> {
        let value = Self {
            id,
            name: name.into(),
            aggregation: ListTotalAggregation::None,
            total_formula: None,
            total_string: None,
        };
        validate_column_name(&value.name)?;
        Ok(value)
    }
    pub fn with_total_aggregation(mut self, value: ListTotalAggregation) -> Result<Self> {
        self.aggregation = value;
        self.validate_totals()?;
        Ok(self)
    }
    pub fn with_total_formula_tokens(mut self, tokens: Vec<u8>) -> Result<Self> {
        if tokens.is_empty() || tokens.len() > u16::MAX as usize {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total formula token length must be 1..=65535",
            ));
        }
        self.aggregation = ListTotalAggregation::Custom;
        self.total_formula = Some(tokens);
        self.validate_totals()?;
        Ok(self)
    }
    pub fn with_total_string(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.encode_utf16().count() > 32767 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string exceeds 32767 UTF-16 units",
            ));
        }
        self.aggregation = ListTotalAggregation::None;
        self.total_string = Some(value);
        self.validate_totals()?;
        Ok(self)
    }
    pub(super) fn validate_totals(&self) -> Result<()> {
        if self.total_formula.is_some() != (self.aggregation == ListTotalAggregation::Custom) {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "custom aggregation and total formula must occur together",
            ));
        }
        if self.total_string.is_some() && self.aggregation != ListTotalAggregation::None {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string requires no aggregation",
            ));
        }
        Ok(())
    }
    pub const fn id(&self) -> ListColumnId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn total_aggregation(&self) -> ListTotalAggregation {
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
pub struct ListObjectStyleOptions {
    pub(super) name: String,
    pub(super) first: bool,
    pub(super) last: bool,
    pub(super) row_stripes: bool,
    pub(super) column_stripes: bool,
    pub(super) default_style: bool,
}
impl ListObjectStyleOptions {
    pub fn try_new(name: impl Into<String>) -> Result<Self> {
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
pub enum ListObjectFeatureVersion {
    Feature11,
    Feature12,
}

/// Excel version recorded by an external-data table definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalTableVersion {
    Excel2003,
    Excel2007,
}
impl ExternalTableVersion {
    pub(super) const fn code(self) -> u32 {
        match self {
            Self::Excel2003 => 0xB,
            Self::Excel2007 => 0xC,
        }
    }
    pub(super) fn from_code(value: u32) -> Result<Self> {
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
pub struct CachedDiskHeader {
    pub(super) encoded: Vec<u8>,
    pub(super) format_end: usize,
    pub(super) style_name: Option<String>,
}

impl CachedDiskHeader {
    /// Construct a cached header from an inert serialized DXFN12List payload.
    pub fn try_new(formatting: Vec<u8>) -> Result<Self> {
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

    pub(super) fn empty() -> Self {
        Self::try_new(Vec::new()).expect("empty cached header is valid")
    }

    pub(super) fn parse(encoded: Vec<u8>, has_style_name: bool, rt: u16) -> Result<Self> {
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

    pub fn with_style_name(mut self, name: impl Into<String>) -> Result<Self> {
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
pub struct ExternalTableField {
    pub(super) column_id: ListColumnId,
    pub(super) source_name: String,
    pub(super) query_field_id: u32,
    pub(super) aggregate_format: Vec<u8>,
    pub(super) insert_row_format: Vec<u8>,
    pub(super) auto_filter: Vec<u8>,
    pub(super) formula_extra: Vec<u8>,
    pub(super) header_cache: CachedDiskHeader,
    pub(super) aggregate_style: u32,
    pub(super) insert_row_style: u32,
    pub(super) filter_hidden: bool,
    pub(super) total_array_formula: bool,
    pub(super) auto_create_calculated_column: bool,
}
impl ExternalTableField {
    pub fn try_new(
        column_id: ListColumnId,
        source_name: impl Into<String>,
        query_field_id: u32,
    ) -> Result<Self> {
        let value = Self {
            column_id,
            source_name: source_name.into(),
            query_field_id,
            aggregate_format: Vec::new(),
            insert_row_format: Vec::new(),
            auto_filter: vec![0; 6],
            formula_extra: Vec::new(),
            header_cache: CachedDiskHeader::empty(),
            aggregate_style: u32::MAX,
            insert_row_style: u32::MAX,
            filter_hidden: false,
            total_array_formula: false,
            auto_create_calculated_column: false,
        };
        value.validate()?;
        Ok(value)
    }
    pub(super) fn validate(&self) -> Result<()> {
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
    pub const fn column_id(&self) -> ListColumnId {
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
    pub const fn cached_disk_header(&self) -> &CachedDiskHeader {
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
    pub fn with_aggregate_format_bytes(mut self, bytes: Vec<u8>) -> Result<Self> {
        self.aggregate_format = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_insert_row_format_bytes(mut self, bytes: Vec<u8>) -> Result<Self> {
        self.insert_row_format = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_auto_filter_bytes(mut self, bytes: Vec<u8>) -> Result<Self> {
        self.auto_filter = bytes;
        self.validate()?;
        Ok(self)
    }
    pub fn with_formula_extra_bytes(mut self, bytes: Vec<u8>, array: bool) -> Result<Self> {
        self.formula_extra = bytes;
        self.total_array_formula = array;
        self.validate()?;
        Ok(self)
    }
    pub fn with_header_cache_bytes(mut self, bytes: Vec<u8>) -> Result<Self> {
        let format_len = usize::try_from(u32_at(&bytes, 0, FEATURE12_RECORD_TYPE, "cbdxfHdrDisk")?)
            .map_err(|_| invalid(FEATURE12_RECORD_TYPE, "cached header length overflows"))?;
        let format_end = 4usize
            .checked_add(format_len)
            .ok_or_else(|| invalid(FEATURE12_RECORD_TYPE, "cached header length overflows"))?;
        let has_style_name = format_end < bytes.len();
        self.header_cache = CachedDiskHeader::parse(bytes, has_style_name, FEATURE12_RECORD_TYPE)?;
        self.validate()?;
        Ok(self)
    }
    pub fn with_cached_disk_header(mut self, header: CachedDiskHeader) -> Result<Self> {
        self.header_cache = header;
        self.validate()?;
        Ok(self)
    }
}

/// Typed, non-executing metadata for a Feature12 LTEXTERNALDATA table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalTableMetadata {
    pub(super) version: ExternalTableVersion,
    pub(super) build_number: u16,
    pub(super) fields: Vec<ExternalTableField>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebColumnType {
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
impl WebColumnType {
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
    pub(super) fn code(self) -> u32 {
        self.value()
    }
    pub(super) fn from_code(value: u32) -> Result<Self> {
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
pub enum WebReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}
impl WebReadingOrder {
    pub(super) fn code(self) -> u32 {
        self as u32
    }
    pub(super) fn from_code(v: u32) -> Result<Self> {
        match v {
            0 => Ok(Self::Context),
            1 => Ok(Self::LeftToRight),
            2 => Ok(Self::RightToLeft),
            _ => Err(invalid(FEATURE11_RECORD_TYPE, "invalid Web reading order")),
        }
    }
}
#[derive(Debug, Clone, PartialEq)]
pub enum WebDefaultValue {
    String(String),
    Boolean(bool),
    Number(f64),
    DateTime(f64),
}
impl Eq for WebDefaultValue {}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebFieldInfo {
    pub(super) locale: u32,
    pub(super) decimal_places: u32,
    pub(super) percent: bool,
    pub(super) fixed_decimal: bool,
    pub(super) date_only: bool,
    pub(super) reading_order: WebReadingOrder,
    pub(super) rich_text: bool,
    pub(super) unknown_rich_text: bool,
    pub(super) alert_unknown_rich_text: bool,
    pub(super) read_only: bool,
    pub(super) required: bool,
    pub(super) minimum_set: bool,
    pub(super) maximum_set: bool,
    pub(super) default_today: bool,
    pub(super) allow_fill_in: bool,
    pub(super) default_value: Option<WebDefaultValue>,
    pub(super) validation_formula: Option<String>,
    pub(super) ignored_display_flags: u32,
    pub(super) ignored_validation_flags: u32,
}
impl WebFieldInfo {
    pub fn new(locale: u32) -> Self {
        Self {
            locale,
            decimal_places: 0,
            percent: false,
            fixed_decimal: false,
            date_only: false,
            reading_order: WebReadingOrder::Context,
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
    pub fn with_default_value(mut self, value: WebDefaultValue) -> Self {
        self.default_value = Some(value);
        self
    }
    pub fn with_validation_formula(mut self, value: impl Into<String>) -> Result<Self> {
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
    pub fn default_value(&self) -> Option<&WebDefaultValue> {
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
    pub(super) fn validate(&self, kind: WebColumnType) -> Result<()> {
        if self.reading_order.code() > 2 {
            return Err(invalid(FEATURE11_RECORD_TYPE, "invalid Web reading order"));
        }
        if let Some(value) = &self.default_value {
            let valid = matches!(
                (kind, value),
                (
                    WebColumnType::Text | WebColumnType::Choice | WebColumnType::MultipleChoices,
                    WebDefaultValue::String(_)
                ) | (WebColumnType::Boolean, WebDefaultValue::Boolean(_))
                    | (
                        WebColumnType::Number | WebColumnType::Currency,
                        WebDefaultValue::Number(_)
                    )
                    | (WebColumnType::DateTime, WebDefaultValue::DateTime(_))
            );
            if !valid {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Web default value does not match column type",
                ));
            }
            if let WebDefaultValue::String(value) = value
                && value.encode_utf16().count() > 255
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Web default string exceeds 255 characters",
                ));
            }
        }
        Ok(())
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebTableField {
    pub(super) column_id: ListColumnId,
    pub(super) source_name: String,
    pub(super) data_type: WebColumnType,
    pub(super) info: WebFieldInfo,
    pub(super) calculated_formula: Option<Vec<u8>>,
    pub(super) auto_filter: Vec<u8>,
    pub(super) aggregate_format: Vec<u8>,
    pub(super) insert_row_format: Vec<u8>,
    pub(super) total_formula_extra: Vec<u8>,
    pub(super) header_cache: Vec<u8>,
    pub(super) ignored_flags: u32,
}
impl WebTableField {
    pub fn try_new(
        column_id: ListColumnId,
        source_name: impl Into<String>,
        data_type: WebColumnType,
        info: WebFieldInfo,
    ) -> Result<Self> {
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
    pub fn with_calculated_formula_tokens(mut self, tokens: Vec<u8>) -> Result<Self> {
        if tokens.is_empty() || tokens.len() > u16::MAX as usize {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "Web calculated formula token length must be 1..=65535",
            ));
        }
        self.calculated_formula = Some(tokens);
        Ok(self)
    }
    pub const fn column_id(&self) -> ListColumnId {
        self.column_id
    }
    pub fn source_name(&self) -> &str {
        &self.source_name
    }
    pub const fn data_type(&self) -> WebColumnType {
        self.data_type
    }
    pub const fn info(&self) -> &WebFieldInfo {
        &self.info
    }
    pub fn calculated_formula_tokens(&self) -> Option<&[u8]> {
        self.calculated_formula.as_deref()
    }
    /// Undefined Feat11FieldDataItem flag bits retained from parsed input.
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
    pub(super) fn validate(&self) -> Result<()> {
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
pub enum WebEditMode {
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
impl WebEditMode {
    pub(super) fn code(self) -> u32 {
        self as u32
    }
    pub(super) fn from_code(v: u32) -> Result<Self> {
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
pub struct WebInvalidCell {
    pub(super) row_id: u32,
    pub(super) column_id: ListColumnId,
}
impl WebInvalidCell {
    pub fn new(row_id: u32, column_id: ListColumnId) -> Self {
        Self { row_id, column_id }
    }
    pub const fn row_id(self) -> u32 {
        self.row_id
    }
    pub const fn column_id(self) -> ListColumnId {
        self.column_id
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebTableMetadata {
    pub(super) version: ExternalTableVersion,
    pub(super) build_number: u16,
    pub(super) fields: Vec<WebTableField>,
    pub(super) edit_mode: WebEditMode,
    pub(super) cache_position: u32,
    pub(super) cache_size: u32,
    pub(super) cache_characters: u32,
    pub(super) hash_parameters: [u8; 16],
    pub(super) provider_name: Option<String>,
    pub(super) entry_id: Option<String>,
    pub(super) deleted_row_ids: Vec<u32>,
    pub(super) changed_row_ids: Vec<u32>,
    pub(super) invalid_cells: Vec<WebInvalidCell>,
    pub(super) needs_commit: bool,
    pub(super) compressed_cache: bool,
    pub(super) ignored_fixed_word: u16,
    pub(super) ignored_flags: u32,
}
impl WebTableMetadata {
    pub fn try_new(fields: Vec<WebTableField>) -> Result<Self> {
        let value = Self {
            version: ExternalTableVersion::Excel2003,
            build_number: 0,
            fields,
            edit_mode: WebEditMode::Normal,
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
    pub fn fields(&self) -> &[WebTableField] {
        &self.fields
    }
    pub const fn edit_mode(&self) -> WebEditMode {
        self.edit_mode
    }
    pub fn deleted_row_ids(&self) -> &[u32] {
        &self.deleted_row_ids
    }
    pub fn changed_row_ids(&self) -> &[u32] {
        &self.changed_row_ids
    }
    pub fn invalid_cells(&self) -> &[WebInvalidCell] {
        &self.invalid_cells
    }
    pub const fn ignored_fixed_word(&self) -> u16 {
        self.ignored_fixed_word
    }
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
    pub fn with_deleted_row_ids(mut self, v: Vec<u32>) -> Result<Self> {
        self.deleted_row_ids = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_changed_row_ids(mut self, v: Vec<u32>) -> Result<Self> {
        self.changed_row_ids = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_invalid_cells(mut self, v: Vec<WebInvalidCell>) -> Result<Self> {
        self.invalid_cells = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_provider_name(mut self, v: impl Into<String>) -> Result<Self> {
        let v = v.into();
        validate_name(&v, "Web cryptographic provider")?;
        self.provider_name = Some(v);
        Ok(self)
    }
    pub fn with_entry_id(mut self, v: impl Into<String>) -> Result<Self> {
        let v = v.into();
        validate_name(&v, "Web entry id")?;
        self.entry_id = Some(v);
        Ok(self)
    }
    pub(super) fn validate(&self) -> Result<()> {
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
pub enum XmlDataType {
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
impl XmlDataType {
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
    pub fn try_new(v: u32) -> Result<Self> {
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
pub struct XmlColumnMapping {
    pub(super) can_be_single: bool,
    pub(super) map_id: u32,
    pub(super) xpath: String,
}
impl XmlColumnMapping {
    pub fn try_new(map_id: u32, xpath: impl Into<String>, can_be_single: bool) -> Result<Self> {
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
pub struct XmlTableField {
    pub(super) column_id: ListColumnId,
    pub(super) source_name: String,
    pub(super) data_type: XmlDataType,
    pub(super) mapping: Option<XmlColumnMapping>,
    pub(super) auto_filter: Vec<u8>,
    pub(super) aggregate_format: Vec<u8>,
    pub(super) insert_row_format: Vec<u8>,
    pub(super) total_formula_extra: Vec<u8>,
    pub(super) header_cache: Vec<u8>,
    pub(super) ignored_flags: u32,
}
impl XmlTableField {
    pub fn try_new(
        column_id: ListColumnId,
        source_name: impl Into<String>,
        data_type: XmlDataType,
    ) -> Result<Self> {
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
    pub fn with_mapping(mut self, v: XmlColumnMapping) -> Self {
        self.mapping = Some(v);
        self
    }
    pub const fn column_id(&self) -> ListColumnId {
        self.column_id
    }
    pub fn source_name(&self) -> &str {
        &self.source_name
    }
    pub const fn data_type(&self) -> XmlDataType {
        self.data_type
    }
    pub fn mapping(&self) -> Option<&XmlColumnMapping> {
        self.mapping.as_ref()
    }
    /// Undefined Feat11FieldDataItem flag bits retained from parsed input.
    pub const fn ignored_flags(&self) -> u32 {
        self.ignored_flags
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlTableMetadata {
    pub(super) version: ExternalTableVersion,
    pub(super) build_number: u16,
    pub(super) fields: Vec<XmlTableField>,
    pub(super) entry_id: Option<String>,
    pub(super) single_cell: bool,
    pub(super) ignored_fixed_word: u16,
    pub(super) ignored_flags: u32,
    pub(super) ignored_fixed_tail: [u8; 32],
}
impl XmlTableMetadata {
    pub fn try_new(fields: Vec<XmlTableField>) -> Result<Self> {
        if !(1..=256).contains(&fields.len()) {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "XML field count must be 1..=256",
            ));
        }
        let value = Self {
            version: ExternalTableVersion::Excel2003,
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
    pub fn fields(&self) -> &[XmlTableField] {
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
    pub fn with_entry_id(mut self, v: impl Into<String>) -> Result<Self> {
        let v = v.into();
        validate_name(&v, "XML entry id")?;
        self.entry_id = Some(v);
        Ok(self)
    }
    pub fn with_single_cell(mut self, v: bool) -> Result<Self> {
        self.single_cell = v;
        self.validate()?;
        Ok(self)
    }
    pub(super) fn validate(&self) -> Result<()> {
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
pub enum ListObjectSourceMetadata {
    Web(WebTableMetadata),
    Xml(XmlTableMetadata),
}

impl ExternalTableMetadata {
    pub fn try_new(fields: Vec<ExternalTableField>) -> Result<Self> {
        let value = Self {
            version: ExternalTableVersion::Excel2007,
            build_number: 0,
            fields,
        };
        value.validate()?;
        Ok(value)
    }
    pub(super) fn validate(&self) -> Result<()> {
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
    pub const fn version(&self) -> ExternalTableVersion {
        self.version
    }
    pub const fn build_number(&self) -> u16 {
        self.build_number
    }
    pub fn fields(&self) -> &[ExternalTableField] {
        &self.fields
    }
    pub fn with_version(mut self, version: ExternalTableVersion) -> Self {
        self.version = version;
        self
    }
    pub fn with_build_number(mut self, build_number: u16) -> Self {
        self.build_number = build_number;
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueListObjectFeature {
    pub(super) record_type: u16,
    pub(super) base_payload: Vec<u8>,
    pub(super) continuation_payloads: Vec<Vec<u8>>,
}
impl OpaqueListObjectFeature {
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
pub struct OpaqueListObjectFutureRecord {
    pub(super) record_type: u16,
    pub(super) payload: Vec<u8>,
    pub(super) continuation_payloads: Vec<Vec<u8>>,
    pub(super) after_list12_count: usize,
}
impl OpaqueListObjectFutureRecord {
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
pub struct ListObject {
    pub(super) id: ListObjectId,
    pub(super) name: String,
    pub(super) range: ListObjectRange,
    pub(super) columns: Vec<ListObjectColumn>,
    pub(super) style: Option<ListObjectStyleOptions>,
    pub(super) has_header: bool,
    pub(super) has_totals: bool,
    pub(super) autofilter: bool,
    pub(super) comment: String,
    pub(super) feature_version: ListObjectFeatureVersion,
    pub(super) opaque_feature: Option<OpaqueListObjectFeature>,
    pub(super) opaque_future_records: Vec<OpaqueListObjectFutureRecord>,
    pub(super) autofilter12_criteria: Option<TableAutoFilter12>,
    pub(super) external_metadata: Option<ExternalTableMetadata>,
    pub(super) source_metadata: Option<ListObjectSourceMetadata>,
}
impl ListObject {
    pub fn try_new(
        id: ListObjectId,
        name: impl Into<String>,
        range: ListObjectRange,
        columns: Vec<ListObjectColumn>,
        style: ListObjectStyleOptions,
    ) -> Result<Self> {
        let feature_version = if columns
            .iter()
            .any(|c| c.total_formula.is_some() || c.total_string.is_some())
        {
            ListObjectFeatureVersion::Feature12
        } else {
            ListObjectFeatureVersion::Feature11
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
    pub fn with_header_row(mut self, v: bool) -> Result<Self> {
        self.has_header = v;
        if !v {
            self.autofilter = false;
            self.feature_version = ListObjectFeatureVersion::Feature12;
        } else if self.opaque_feature.is_none() {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_totals_row(mut self, v: bool) -> Result<Self> {
        self.has_totals = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter(mut self, v: bool) -> Result<Self> {
        self.autofilter = v;
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter12_criteria(mut self, value: TableAutoFilter12) -> Result<Self> {
        self.autofilter12_criteria = Some(value);
        self.validate()?;
        Ok(self)
    }
    pub fn with_comment(mut self, v: impl Into<String>) -> Result<Self> {
        self.comment = v.into();
        if self.comment.encode_utf16().count() > 255 {
            return Err(invalid(
                LIST12_RECORD_TYPE,
                "table comment exceeds 255 characters",
            ));
        }
        Ok(self)
    }
    pub fn with_external_data(mut self, metadata: ExternalTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.external_metadata = Some(metadata);
        self.feature_version = ListObjectFeatureVersion::Feature12;
        self.opaque_feature = None;
        self.validate()?;
        Ok(self)
    }
    pub fn with_web_source(mut self, metadata: WebTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.source_metadata = Some(ListObjectSourceMetadata::Web(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != ListObjectFeatureVersion::Feature12 {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_xml_source(mut self, metadata: XmlTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.source_metadata = Some(ListObjectSourceMetadata::Xml(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != ListObjectFeatureVersion::Feature12 {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub const fn id(&self) -> ListObjectId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn range(&self) -> ListObjectRange {
        self.range
    }
    pub fn columns(&self) -> &[ListObjectColumn] {
        &self.columns
    }
    pub fn style(&self) -> Option<&ListObjectStyleOptions> {
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
    pub const fn feature_version(&self) -> ListObjectFeatureVersion {
        self.feature_version
    }
    pub fn opaque_feature(&self) -> Option<&OpaqueListObjectFeature> {
        self.opaque_feature.as_ref()
    }
    pub fn opaque_future_records(&self) -> &[OpaqueListObjectFutureRecord] {
        &self.opaque_future_records
    }
    pub fn autofilter12_criteria(&self) -> Option<&TableAutoFilter12> {
        self.autofilter12_criteria.as_ref()
    }
    pub fn external_metadata(&self) -> Option<&ExternalTableMetadata> {
        self.external_metadata.as_ref()
    }
    pub fn source_metadata(&self) -> Option<&ListObjectSourceMetadata> {
        self.source_metadata.as_ref()
    }
    pub(crate) fn validate(&self) -> Result<()> {
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
            && self.feature_version != ListObjectFeatureVersion::Feature12
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "opaque table feature must be Feature12",
            ));
        }
        if let Some(metadata) = &self.external_metadata {
            metadata.validate()?;
            if self.feature_version != ListObjectFeatureVersion::Feature12
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
                ListObjectFeatureVersion::Feature11 | ListObjectFeatureVersion::Feature12
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
            if self.feature_version == ListObjectFeatureVersion::Feature11 && has_feature12_field {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Feature11 source fields cannot load total formulas or strings",
                ));
            }
            if self.feature_version == ListObjectFeatureVersion::Feature12
                && self.has_header
                && !has_feature12_field
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "Feature12 Web/XML source requires a Feature12-only property",
                ));
            }
            match source {
                ListObjectSourceMetadata::Web(metadata) => {
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
                ListObjectSourceMetadata::Xml(metadata) => {
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
}
