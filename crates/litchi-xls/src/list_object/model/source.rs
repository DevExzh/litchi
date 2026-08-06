//! Typed Web, XML, and external-data source metadata.

use super::super::codec::{append_string, parse_string, u32_at};
use super::super::{FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, MAX_FEATURE_BYTES, invalid};
use super::{ListColumnId, validate_name};
use crate::Result;
use crate::xml_map::{MapId, XPath};
use std::collections::HashSet;

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
    pub(in crate::list_object) const fn code(self) -> u32 {
        match self {
            Self::Excel2003 => 0xB,
            Self::Excel2007 => 0xC,
        }
    }
    pub(in crate::list_object) fn from_code(value: u32) -> Result<Self> {
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
    pub(in crate::list_object) encoded: Vec<u8>,
    pub(in crate::list_object) format_end: usize,
    pub(in crate::list_object) style_name: Option<String>,
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

    pub(in crate::list_object) fn empty() -> Self {
        Self::try_new(Vec::new()).expect("empty cached header is valid")
    }

    pub(in crate::list_object) fn parse(
        encoded: Vec<u8>,
        has_style_name: bool,
        rt: u16,
    ) -> Result<Self> {
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
    pub(in crate::list_object) column_id: ListColumnId,
    pub(in crate::list_object) source_name: String,
    pub(in crate::list_object) query_field_id: u32,
    pub(in crate::list_object) aggregate_format: Vec<u8>,
    pub(in crate::list_object) insert_row_format: Vec<u8>,
    pub(in crate::list_object) auto_filter: Vec<u8>,
    pub(in crate::list_object) formula_extra: Vec<u8>,
    pub(in crate::list_object) header_cache: CachedDiskHeader,
    pub(in crate::list_object) aggregate_style: u32,
    pub(in crate::list_object) insert_row_style: u32,
    pub(in crate::list_object) filter_hidden: bool,
    pub(in crate::list_object) total_array_formula: bool,
    pub(in crate::list_object) auto_create_calculated_column: bool,
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
    pub(in crate::list_object) fn validate(&self) -> Result<()> {
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
    pub(in crate::list_object) version: ExternalTableVersion,
    pub(in crate::list_object) build_number: u16,
    pub(in crate::list_object) fields: Vec<ExternalTableField>,
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
    pub(in crate::list_object) fn code(self) -> u32 {
        self.value()
    }
    pub(in crate::list_object) fn from_code(value: u32) -> Result<Self> {
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
    pub(in crate::list_object) fn code(self) -> u32 {
        self as u32
    }
    pub(in crate::list_object) fn from_code(v: u32) -> Result<Self> {
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
    pub(in crate::list_object) locale: u32,
    pub(in crate::list_object) decimal_places: u32,
    pub(in crate::list_object) percent: bool,
    pub(in crate::list_object) fixed_decimal: bool,
    pub(in crate::list_object) date_only: bool,
    pub(in crate::list_object) reading_order: WebReadingOrder,
    pub(in crate::list_object) rich_text: bool,
    pub(in crate::list_object) unknown_rich_text: bool,
    pub(in crate::list_object) alert_unknown_rich_text: bool,
    pub(in crate::list_object) read_only: bool,
    pub(in crate::list_object) required: bool,
    pub(in crate::list_object) minimum_set: bool,
    pub(in crate::list_object) maximum_set: bool,
    pub(in crate::list_object) default_today: bool,
    pub(in crate::list_object) allow_fill_in: bool,
    pub(in crate::list_object) default_value: Option<WebDefaultValue>,
    pub(in crate::list_object) validation_formula: Option<String>,
    pub(in crate::list_object) ignored_display_flags: u32,
    pub(in crate::list_object) ignored_validation_flags: u32,
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
    pub(in crate::list_object) fn validate(&self, kind: WebColumnType) -> Result<()> {
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
    pub(in crate::list_object) column_id: ListColumnId,
    pub(in crate::list_object) source_name: String,
    pub(in crate::list_object) data_type: WebColumnType,
    pub(in crate::list_object) info: WebFieldInfo,
    pub(in crate::list_object) calculated_formula: Option<Vec<u8>>,
    pub(in crate::list_object) auto_filter: Vec<u8>,
    pub(in crate::list_object) aggregate_format: Vec<u8>,
    pub(in crate::list_object) insert_row_format: Vec<u8>,
    pub(in crate::list_object) total_formula_extra: Vec<u8>,
    pub(in crate::list_object) header_cache: Vec<u8>,
    pub(in crate::list_object) ignored_flags: u32,
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
    pub(in crate::list_object) fn validate(&self) -> Result<()> {
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
    pub(in crate::list_object) fn code(self) -> u32 {
        self as u32
    }
    pub(in crate::list_object) fn from_code(v: u32) -> Result<Self> {
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
    pub(in crate::list_object) row_id: u32,
    pub(in crate::list_object) column_id: ListColumnId,
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
    pub(in crate::list_object) version: ExternalTableVersion,
    pub(in crate::list_object) build_number: u16,
    pub(in crate::list_object) fields: Vec<WebTableField>,
    pub(in crate::list_object) edit_mode: WebEditMode,
    pub(in crate::list_object) cache_position: u32,
    pub(in crate::list_object) cache_size: u32,
    pub(in crate::list_object) cache_characters: u32,
    pub(in crate::list_object) hash_parameters: [u8; 16],
    pub(in crate::list_object) provider_name: Option<String>,
    pub(in crate::list_object) entry_id: Option<String>,
    pub(in crate::list_object) deleted_row_ids: Vec<u32>,
    pub(in crate::list_object) changed_row_ids: Vec<u32>,
    pub(in crate::list_object) invalid_cells: Vec<WebInvalidCell>,
    pub(in crate::list_object) needs_commit: bool,
    pub(in crate::list_object) compressed_cache: bool,
    pub(in crate::list_object) ignored_fixed_word: u16,
    pub(in crate::list_object) ignored_flags: u32,
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
    pub(in crate::list_object) fn validate(&self) -> Result<()> {
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
    pub(in crate::list_object) can_be_single: bool,
    pub(in crate::list_object) map_id: MapId,
    pub(in crate::list_object) xpath: XPath,
}
impl XmlColumnMapping {
    pub fn try_new(map_id: u32, xpath: impl Into<String>, can_be_single: bool) -> Result<Self> {
        let map_id =
            MapId::new(map_id).map_err(|_| invalid(FEATURE11_RECORD_TYPE, "invalid XML map id"))?;
        let xpath = XPath::new(xpath.into())
            .map_err(|_| invalid(FEATURE11_RECORD_TYPE, "invalid XML XPath"))?;
        Ok(Self {
            can_be_single,
            map_id,
            xpath,
        })
    }
    pub const fn map_id(&self) -> u32 {
        self.map_id.get()
    }
    pub const fn map_identifier(&self) -> MapId {
        self.map_id
    }
    pub fn xpath(&self) -> &str {
        self.xpath.as_str()
    }
    pub const fn path(&self) -> &XPath {
        &self.xpath
    }
    pub const fn can_be_single(&self) -> bool {
        self.can_be_single
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlTableField {
    pub(in crate::list_object) column_id: ListColumnId,
    pub(in crate::list_object) source_name: String,
    pub(in crate::list_object) data_type: XmlDataType,
    pub(in crate::list_object) mapping: Option<XmlColumnMapping>,
    pub(in crate::list_object) auto_filter: Vec<u8>,
    pub(in crate::list_object) aggregate_format: Vec<u8>,
    pub(in crate::list_object) insert_row_format: Vec<u8>,
    pub(in crate::list_object) total_formula_extra: Vec<u8>,
    pub(in crate::list_object) header_cache: Vec<u8>,
    pub(in crate::list_object) ignored_flags: u32,
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
    pub(in crate::list_object) version: ExternalTableVersion,
    pub(in crate::list_object) build_number: u16,
    pub(in crate::list_object) fields: Vec<XmlTableField>,
    pub(in crate::list_object) entry_id: Option<String>,
    pub(in crate::list_object) single_cell: bool,
    pub(in crate::list_object) ignored_fixed_word: u16,
    pub(in crate::list_object) ignored_flags: u32,
    pub(in crate::list_object) ignored_fixed_tail: [u8; 32],
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
    pub(in crate::list_object) fn validate(&self) -> Result<()> {
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
    pub(in crate::list_object) fn validate(&self) -> Result<()> {
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
