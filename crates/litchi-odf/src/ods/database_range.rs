//! ODF spreadsheet database ranges, filters, sorting, and subtotals.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_FILTER_DEPTH: usize = 128;

trait HasLocalName {
    fn has_local_name(&self, expected: &[u8]) -> bool;
}

impl HasLocalName for BytesStart<'_> {
    fn has_local_name(&self, expected: &[u8]) -> bool {
        self.local_name().as_ref() == expected
    }
}

impl HasLocalName for BytesEnd<'_> {
    fn has_local_name(&self, expected: &[u8]) -> bool {
        self.local_name().as_ref() == expected
    }
}

/// Whether database fields are arranged in columns or rows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DatabaseOrientation {
    /// Each field occupies a column.
    Column,
    /// Each field occupies a row.
    Row,
}

impl DatabaseOrientation {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "column" => Ok(Self::Column),
            "row" => Ok(Self::Row),
            _ => Err(invalid("table:orientation", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Column => "column",
            Self::Row => "row",
        }
    }
}

/// An inert external database source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DatabaseSource {
    /// A SQL statement stored in the document. It is never executed by this crate.
    Sql {
        /// Database identifier or URI.
        database_name: String,
        /// SQL statement retained as data.
        statement: String,
        /// Whether a consumer should parse the SQL statement.
        parse_statement: Option<bool>,
    },
    /// A database table source.
    Table {
        /// Database identifier or URI.
        database_name: String,
        /// Database table name.
        table_name: String,
    },
    /// A named database query source.
    Query {
        /// Database identifier or URI.
        database_name: String,
        /// Query name.
        query_name: String,
    },
}

/// Sort order for keys and subtotal groups.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortOrder {
    /// Ascending order.
    Ascending,
    /// Descending order.
    Descending,
}

impl SortOrder {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "ascending" => Ok(Self::Ascending),
            "descending" => Ok(Self::Descending),
            _ => Err(invalid("table:order", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }
}

/// How embedded numbers in text participate in sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmbeddedNumberBehavior {
    /// Compare the entire value alphabetically.
    AlphaNumeric,
    /// Compare embedded integer runs numerically.
    Integer,
    /// Compare embedded numbers as floating-point values.
    Double,
}

impl EmbeddedNumberBehavior {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "alpha-numeric" => Ok(Self::AlphaNumeric),
            "integer" => Ok(Self::Integer),
            "double" => Ok(Self::Double),
            _ => Err(invalid("table:embedded-number-behavior", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::AlphaNumeric => "alpha-numeric",
            Self::Integer => "integer",
            Self::Double => "double",
        }
    }
}

/// One field in a database-range sort specification.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseSortKey {
    /// Zero-based field number.
    pub field_number: u64,
    /// Standard or application-defined sort data type.
    pub data_type: Option<String>,
    /// Optional explicit order.
    pub order: Option<SortOrder>,
}

impl DatabaseSortKey {
    /// Create a sort key for a zero-based field number.
    pub fn new(field_number: u64) -> Self {
        Self {
            field_number,
            data_type: None,
            order: None,
        }
    }
}

/// Sort configuration attached to a database range.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DatabaseSort {
    /// Whether styles remain bound to sorted content.
    pub bind_styles_to_content: Option<bool>,
    /// Optional destination range.
    pub target_range_address: Option<String>,
    /// Whether string comparisons are case-sensitive.
    pub case_sensitive: Option<bool>,
    /// Legacy language code.
    pub language: Option<String>,
    /// Legacy country code.
    pub country: Option<String>,
    /// ISO 15924 script code.
    pub script: Option<String>,
    /// BCP 47 language tag.
    pub rfc_language_tag: Option<String>,
    /// Application-defined collation algorithm.
    pub algorithm: Option<String>,
    /// Embedded-number comparison behavior.
    pub embedded_number_behavior: Option<EmbeddedNumberBehavior>,
    /// Ordered sort keys. ODF requires at least one.
    pub keys: Vec<DatabaseSortKey>,
}

/// Source used to obtain filter conditions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterConditionSource {
    /// Conditions are contained in the filter itself.
    SelfContained,
    /// Conditions come from another cell range.
    CellRange,
}

impl FilterConditionSource {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "self" => Ok(Self::SelfContained),
            "cell-range" => Ok(Self::CellRange),
            _ => Err(invalid("table:condition-source", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::SelfContained => "self",
            Self::CellRange => "cell-range",
        }
    }
}

/// Standard filter comparison data type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterDataType {
    /// Text comparison.
    Text,
    /// Numeric comparison.
    Number,
}

impl FilterDataType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "text" => Ok(Self::Text),
            "number" => Ok(Self::Number),
            _ => Err(invalid("table:data-type", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Number => "number",
        }
    }
}

/// A leaf filter comparison.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FilterCondition {
    /// Zero-based field number.
    pub field_number: u64,
    /// Comparison value.
    pub value: String,
    /// Standard or application-defined operator.
    pub operator: String,
    /// Optional case-sensitivity override.
    pub case_sensitive: Option<bool>,
    /// Optional comparison data type.
    pub data_type: Option<FilterDataType>,
    /// Values in a set-membership condition.
    pub set_items: Vec<String>,
}

impl FilterCondition {
    /// Create a filter condition.
    pub fn new(field_number: u64, operator: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            field_number,
            value: value.into(),
            operator: operator.into(),
            case_sensitive: None,
            data_type: None,
            set_items: Vec::new(),
        }
    }
}

/// Recursive filter expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FilterExpression {
    /// A leaf comparison.
    Condition(FilterCondition),
    /// All child expressions must match. Children may be conditions or OR groups.
    And(Vec<FilterExpression>),
    /// At least one child expression must match. Children may be conditions or AND groups.
    Or(Vec<FilterExpression>),
}

/// Filter configuration attached to a database range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseFilter {
    /// Optional destination range.
    pub target_range_address: Option<String>,
    /// Optional condition source.
    pub condition_source: Option<FilterConditionSource>,
    /// Range containing conditions when `condition_source` is `CellRange`.
    pub condition_source_range_address: Option<String>,
    /// Whether duplicate rows remain visible.
    pub display_duplicates: Option<bool>,
    /// Root filter expression.
    pub expression: FilterExpression,
}

/// Sort configuration for subtotal groups.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SubtotalSortGroups {
    /// Standard or application-defined data type.
    pub data_type: Option<String>,
    /// Optional explicit order.
    pub order: Option<SortOrder>,
}

/// A field aggregated by a subtotal rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SubtotalField {
    /// Zero-based field number.
    pub field_number: u64,
    /// Standard or application-defined aggregation function.
    pub function: String,
}

/// One subtotal grouping rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SubtotalRule {
    /// Zero-based grouping field number.
    pub group_by_field_number: u64,
    /// Fields aggregated for the group.
    pub fields: Vec<SubtotalField>,
}

/// Subtotal rules attached to a database range.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SubtotalRules {
    /// Whether styles remain bound to content.
    pub bind_styles_to_content: Option<bool>,
    /// Whether grouping comparisons are case-sensitive.
    pub case_sensitive: Option<bool>,
    /// Whether to insert page breaks when a group changes.
    pub page_breaks_on_group_change: Option<bool>,
    /// Optional group sorting.
    pub sort_groups: Option<SubtotalSortGroups>,
    /// Ordered subtotal rules.
    pub rules: Vec<SubtotalRule>,
}

/// A spreadsheet database range and its non-executing query/filter metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseRange {
    /// Optional range name.
    pub name: Option<String>,
    /// Whether this range represents the current selection.
    pub is_selection: Option<bool>,
    /// Whether styles are retained after external refresh.
    pub on_update_keep_styles: Option<bool>,
    /// Whether the range size is retained after external refresh.
    pub on_update_keep_size: Option<bool>,
    /// Whether imported data is persisted in the document.
    pub has_persistent_data: Option<bool>,
    /// Field orientation.
    pub orientation: Option<DatabaseOrientation>,
    /// Whether the first field is a header.
    pub contains_header: Option<bool>,
    /// Whether filter buttons are displayed.
    pub display_filter_buttons: Option<bool>,
    /// Required cell range occupied by the database range.
    pub target_range_address: String,
    /// Optional XML Schema refresh duration.
    pub refresh_delay: Option<String>,
    /// Optional inert external source.
    pub source: Option<DatabaseSource>,
    /// Optional filter.
    pub filter: Option<DatabaseFilter>,
    /// Optional sorting.
    pub sort: Option<DatabaseSort>,
    /// Optional subtotal rules.
    pub subtotals: Option<SubtotalRules>,
}

impl DatabaseRange {
    /// Create a database range for an ODF cell range address.
    pub fn new(target_range_address: impl Into<String>) -> Self {
        Self {
            name: None,
            is_selection: None,
            on_update_keep_styles: None,
            on_update_keep_size: None,
            has_persistent_data: None,
            orientation: None,
            contains_header: None,
            display_filter_buttons: None,
            target_range_address: target_range_address.into(),
            refresh_delay: None,
            source: None,
            filter: None,
            sort: None,
            subtotals: None,
        }
    }

    /// Validate required values and recursive schema constraints.
    pub fn validate(&self) -> Result<()> {
        if self.target_range_address.is_empty() {
            return Err(Error::InvalidFormat(
                "database range target address cannot be empty".to_string(),
            ));
        }
        if self
            .refresh_delay
            .as_deref()
            .is_some_and(|value| !is_xsd_duration(value))
        {
            return Err(invalid(
                "table:refresh-delay",
                self.refresh_delay.as_deref().expect("delay checked above"),
            ));
        }
        if let Some(filter) = &self.filter {
            validate_filter(filter)?;
        }
        if self.sort.as_ref().is_some_and(|sort| sort.keys.is_empty()) {
            return Err(Error::InvalidFormat(
                "database sort requires at least one sort key".to_string(),
            ));
        }
        Ok(())
    }
}

pub(crate) fn validate_filter(filter: &DatabaseFilter) -> Result<()> {
    validate_filter_expression(&filter.expression, 0, None)?;
    if filter.condition_source == Some(FilterConditionSource::CellRange)
        && filter.condition_source_range_address.is_none()
    {
        return Err(Error::InvalidFormat(
            "cell-range filter source requires table:condition-source-range-address".to_string(),
        ));
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FilterParent {
    And,
    Or,
}

fn validate_filter_expression(
    expression: &FilterExpression,
    depth: usize,
    parent: Option<FilterParent>,
) -> Result<()> {
    if depth > MAX_FILTER_DEPTH {
        return Err(Error::InvalidFormat(
            "filter expression exceeds the supported nesting limit".to_string(),
        ));
    }
    let (children, kind) = match expression {
        FilterExpression::Condition(_) => return Ok(()),
        FilterExpression::And(children) => (children, FilterParent::And),
        FilterExpression::Or(children) => (children, FilterParent::Or),
    };
    if children.is_empty() {
        return Err(Error::InvalidFormat(
            "filter boolean group cannot be empty".to_string(),
        ));
    }
    if parent == Some(kind) {
        return Err(Error::InvalidFormat(
            "ODF filter groups must alternate AND and OR operators".to_string(),
        ));
    }
    for child in children {
        validate_filter_expression(child, depth + 1, Some(kind))?;
    }
    Ok(())
}

pub(crate) fn parse_database_ranges(xml: &str) -> Result<Vec<DatabaseRange>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_ranges = false;
    let mut ranges = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"database-ranges") => {
                in_ranges = true;
            },
            Event::Start(ref element)
                if in_ranges && is_table(&namespace, element, b"database-range") =>
            {
                let range = parse_database_range(&mut reader, element)?;
                range.validate()?;
                ranges.push(range);
            },
            Event::Empty(ref element)
                if in_ranges && is_table(&namespace, element, b"database-range") =>
            {
                let range = database_range_from_start(&reader, element)?;
                range.validate()?;
                ranges.push(range);
            },
            Event::End(ref element) if is_table(&namespace, element, b"database-ranges") => {
                in_ranges = false;
            },
            Event::Eof => break,
            _ => {},
        }
        buf.clear();
    }
    Ok(ranges)
}

fn parse_database_range(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<DatabaseRange> {
    let mut range = database_range_from_start(reader, start)?;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-sql") =>
            {
                ensure_absent(&range.source, "database source")?;
                range.source = Some(parse_source_sql(reader, element)?);
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-table") =>
            {
                ensure_absent(&range.source, "database source")?;
                range.source = Some(parse_source_table(reader, element)?);
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-query") =>
            {
                ensure_absent(&range.source, "database source")?;
                range.source = Some(parse_source_query(reader, element)?);
            },
            Event::Start(ref element) if is_table(&namespace, element, b"filter") => {
                ensure_absent(&range.filter, "database filter")?;
                range.filter = Some(parse_filter(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"filter") => {
                return Err(Error::InvalidFormat(
                    "table:filter has no expression".to_string(),
                ));
            },
            Event::Start(ref element) if is_table(&namespace, element, b"sort") => {
                ensure_absent(&range.sort, "database sort")?;
                range.sort = Some(parse_sort(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"sort") => {
                return Err(Error::InvalidFormat(
                    "database sort requires at least one sort key".to_string(),
                ));
            },
            Event::Start(ref element) if is_table(&namespace, element, b"subtotal-rules") => {
                ensure_absent(&range.subtotals, "subtotal rules")?;
                range.subtotals = Some(parse_subtotals(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"subtotal-rules") => {
                ensure_absent(&range.subtotals, "subtotal rules")?;
                range.subtotals = Some(SubtotalRules::default());
            },
            Event::End(ref element) if is_table(&namespace, element, b"database-range") => break,
            Event::Eof => return Err(unexpected_eof("table:database-range")),
            _ => {},
        }
        buf.clear();
    }
    Ok(range)
}

fn database_range_from_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DatabaseRange> {
    let target = required_attr(reader, element, b"target-range-address")?;
    Ok(DatabaseRange {
        name: optional_attr(reader, element, b"name")?,
        is_selection: optional_bool(reader, element, b"is-selection")?,
        on_update_keep_styles: optional_bool(reader, element, b"on-update-keep-styles")?,
        on_update_keep_size: optional_bool(reader, element, b"on-update-keep-size")?,
        has_persistent_data: optional_bool(reader, element, b"has-persistent-data")?,
        orientation: optional_attr(reader, element, b"orientation")?
            .map(|value| DatabaseOrientation::parse(&value))
            .transpose()?,
        contains_header: optional_bool(reader, element, b"contains-header")?,
        display_filter_buttons: optional_bool(reader, element, b"display-filter-buttons")?,
        target_range_address: target,
        refresh_delay: optional_attr(reader, element, b"refresh-delay")?,
        source: None,
        filter: None,
        sort: None,
        subtotals: None,
    })
}

pub(crate) fn parse_source_sql(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DatabaseSource> {
    let parse_statement = optional_bool(reader, element, b"parse-sql-statement")?
        .or(optional_bool(reader, element, b"parse-sql-statements")?);
    Ok(DatabaseSource::Sql {
        database_name: required_attr(reader, element, b"database-name")?,
        statement: required_attr(reader, element, b"sql-statement")?,
        parse_statement,
    })
}

pub(crate) fn parse_source_table(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DatabaseSource> {
    Ok(DatabaseSource::Table {
        database_name: required_attr(reader, element, b"database-name")?,
        table_name: optional_attr(reader, element, b"database-table-name")?
            .or(optional_attr(reader, element, b"table-name")?)
            .ok_or_else(|| missing("table:database-table-name"))?,
    })
}

pub(crate) fn parse_source_query(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DatabaseSource> {
    Ok(DatabaseSource::Query {
        database_name: required_attr(reader, element, b"database-name")?,
        query_name: required_attr(reader, element, b"query-name")?,
    })
}

pub(crate) fn parse_filter(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<DatabaseFilter> {
    let target_range_address = optional_attr(reader, start, b"target-range-address")?;
    let condition_source = optional_attr(reader, start, b"condition-source")?
        .map(|value| FilterConditionSource::parse(&value))
        .transpose()?;
    let condition_source_range_address =
        optional_attr(reader, start, b"condition-source-range-address")?;
    let display_duplicates = optional_bool(reader, start, b"display-duplicates")?;
    let mut expression = None;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"filter-condition") => {
                ensure_absent(&expression, "filter root expression")?;
                expression = Some(FilterExpression::Condition(parse_condition(
                    reader, element,
                )?));
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"filter-condition") => {
                ensure_absent(&expression, "filter root expression")?;
                expression = Some(FilterExpression::Condition(condition_from_start(
                    reader, element,
                )?));
            },
            Event::Start(ref element) if is_table(&namespace, element, b"filter-and") => {
                ensure_absent(&expression, "filter root expression")?;
                expression = Some(parse_filter_group(reader, FilterParent::And, 1)?);
            },
            Event::Start(ref element) if is_table(&namespace, element, b"filter-or") => {
                ensure_absent(&expression, "filter root expression")?;
                expression = Some(parse_filter_group(reader, FilterParent::Or, 1)?);
            },
            Event::End(ref element) if is_table(&namespace, element, b"filter") => break,
            Event::Eof => return Err(unexpected_eof("table:filter")),
            _ => {},
        }
        buf.clear();
    }
    Ok(DatabaseFilter {
        target_range_address,
        condition_source,
        condition_source_range_address,
        display_duplicates,
        expression: expression
            .ok_or_else(|| Error::InvalidFormat("table:filter has no expression".to_string()))?,
    })
}

fn parse_filter_group(
    reader: &mut NsReader<&[u8]>,
    kind: FilterParent,
    depth: usize,
) -> Result<FilterExpression> {
    if depth > MAX_FILTER_DEPTH {
        return Err(Error::InvalidFormat(
            "filter expression exceeds the supported nesting limit".to_string(),
        ));
    }
    let mut children = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"filter-condition") => {
                children.push(FilterExpression::Condition(parse_condition(
                    reader, element,
                )?));
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"filter-condition") => {
                children.push(FilterExpression::Condition(condition_from_start(
                    reader, element,
                )?));
            },
            Event::Start(ref element)
                if kind == FilterParent::And && is_table(&namespace, element, b"filter-or") =>
            {
                children.push(parse_filter_group(reader, FilterParent::Or, depth + 1)?);
            },
            Event::Start(ref element)
                if kind == FilterParent::Or && is_table(&namespace, element, b"filter-and") =>
            {
                children.push(parse_filter_group(reader, FilterParent::And, depth + 1)?);
            },
            Event::Start(ref element)
                if is_table(&namespace, element, b"filter-and")
                    || is_table(&namespace, element, b"filter-or") =>
            {
                return Err(Error::InvalidFormat(
                    "ODF filter groups must alternate AND and OR operators".to_string(),
                ));
            },
            Event::End(ref element)
                if (kind == FilterParent::And && is_table(&namespace, element, b"filter-and"))
                    || (kind == FilterParent::Or
                        && is_table(&namespace, element, b"filter-or")) =>
            {
                break;
            },
            Event::Eof => return Err(unexpected_eof("filter boolean group")),
            _ => {},
        }
        buf.clear();
    }
    if children.is_empty() {
        return Err(Error::InvalidFormat(
            "filter boolean group cannot be empty".to_string(),
        ));
    }
    Ok(match kind {
        FilterParent::And => FilterExpression::And(children),
        FilterParent::Or => FilterExpression::Or(children),
    })
}

fn parse_condition(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<FilterCondition> {
    let mut condition = condition_from_start(reader, start)?;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Empty(ref element) | Event::Start(ref element)
                if is_table(&namespace, element, b"filter-set-item") =>
            {
                condition
                    .set_items
                    .push(required_attr(reader, element, b"value")?);
            },
            Event::End(ref element) if is_table(&namespace, element, b"filter-condition") => break,
            Event::Eof => return Err(unexpected_eof("table:filter-condition")),
            _ => {},
        }
        buf.clear();
    }
    Ok(condition)
}

fn condition_from_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<FilterCondition> {
    Ok(FilterCondition {
        field_number: required_u64(reader, element, b"field-number")?,
        value: required_attr(reader, element, b"value")?,
        operator: required_attr(reader, element, b"operator")?,
        case_sensitive: optional_bool(reader, element, b"case-sensitive")?,
        data_type: optional_attr(reader, element, b"data-type")?
            .map(|value| FilterDataType::parse(&value))
            .transpose()?,
        set_items: Vec::new(),
    })
}

fn parse_sort(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DatabaseSort> {
    let mut sort = DatabaseSort {
        bind_styles_to_content: optional_bool(reader, start, b"bind-styles-to-content")?,
        target_range_address: optional_attr(reader, start, b"target-range-address")?,
        case_sensitive: optional_bool(reader, start, b"case-sensitive")?,
        language: optional_attr(reader, start, b"language")?,
        country: optional_attr(reader, start, b"country")?,
        script: optional_attr(reader, start, b"script")?,
        rfc_language_tag: optional_attr(reader, start, b"rfc-language-tag")?,
        algorithm: optional_attr(reader, start, b"algorithm")?,
        embedded_number_behavior: optional_attr(reader, start, b"embedded-number-behavior")?
            .map(|value| EmbeddedNumberBehavior::parse(&value))
            .transpose()?,
        keys: Vec::new(),
    };
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Empty(ref element) | Event::Start(ref element)
                if is_table(&namespace, element, b"sort-by") =>
            {
                sort.keys.push(DatabaseSortKey {
                    field_number: required_u64(reader, element, b"field-number")?,
                    data_type: optional_attr(reader, element, b"data-type")?,
                    order: optional_attr(reader, element, b"order")?
                        .map(|value| SortOrder::parse(&value))
                        .transpose()?,
                });
            },
            Event::End(ref element) if is_table(&namespace, element, b"sort") => break,
            Event::Eof => return Err(unexpected_eof("table:sort")),
            _ => {},
        }
        buf.clear();
    }
    if sort.keys.is_empty() {
        return Err(Error::InvalidFormat(
            "database sort requires at least one sort key".to_string(),
        ));
    }
    Ok(sort)
}

fn parse_subtotals(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<SubtotalRules> {
    let mut subtotals = SubtotalRules {
        bind_styles_to_content: optional_bool(reader, start, b"bind-styles-to-content")?,
        case_sensitive: optional_bool(reader, start, b"case-sensitive")?,
        page_breaks_on_group_change: optional_bool(reader, start, b"page-breaks-on-group-change")?,
        sort_groups: None,
        rules: Vec::new(),
    };
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Empty(ref element) | Event::Start(ref element)
                if is_table(&namespace, element, b"sort-groups") =>
            {
                ensure_absent(&subtotals.sort_groups, "subtotal sort-groups")?;
                subtotals.sort_groups = Some(SubtotalSortGroups {
                    data_type: optional_attr(reader, element, b"data-type")?,
                    order: optional_attr(reader, element, b"order")?
                        .map(|value| SortOrder::parse(&value))
                        .transpose()?,
                });
            },
            Event::Start(ref element) if is_table(&namespace, element, b"subtotal-rule") => {
                subtotals.rules.push(parse_subtotal_rule(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"subtotal-rule") => {
                subtotals.rules.push(SubtotalRule {
                    group_by_field_number: required_u64(reader, element, b"group-by-field-number")?,
                    fields: Vec::new(),
                });
            },
            Event::End(ref element) if is_table(&namespace, element, b"subtotal-rules") => break,
            Event::Eof => return Err(unexpected_eof("table:subtotal-rules")),
            _ => {},
        }
        buf.clear();
    }
    Ok(subtotals)
}

fn parse_subtotal_rule(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<SubtotalRule> {
    let mut rule = SubtotalRule {
        group_by_field_number: required_u64(reader, start, b"group-by-field-number")?,
        fields: Vec::new(),
    };
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Empty(ref element) | Event::Start(ref element)
                if is_table(&namespace, element, b"subtotal-field") =>
            {
                rule.fields.push(SubtotalField {
                    field_number: required_u64(reader, element, b"field-number")?,
                    function: required_attr(reader, element, b"function")?,
                });
            },
            Event::End(ref element) if is_table(&namespace, element, b"subtotal-rule") => break,
            Event::Eof => return Err(unexpected_eof("table:subtotal-rule")),
            _ => {},
        }
        buf.clear();
    }
    Ok(rule)
}

pub(crate) fn write_database_ranges(output: &mut String, ranges: &[DatabaseRange]) -> Result<()> {
    if ranges.is_empty() {
        return Ok(());
    }
    output.push_str("<table:database-ranges>");
    for range in ranges {
        range.validate()?;
        write_database_range(output, range);
    }
    output.push_str("</table:database-ranges>");
    Ok(())
}

fn write_database_range(output: &mut String, range: &DatabaseRange) {
    output.push_str("<table:database-range");
    attr(output, "table:name", range.name.as_deref());
    bool_attr(output, "table:is-selection", range.is_selection);
    bool_attr(
        output,
        "table:on-update-keep-styles",
        range.on_update_keep_styles,
    );
    bool_attr(
        output,
        "table:on-update-keep-size",
        range.on_update_keep_size,
    );
    bool_attr(
        output,
        "table:has-persistent-data",
        range.has_persistent_data,
    );
    attr(
        output,
        "table:orientation",
        range.orientation.map(DatabaseOrientation::as_str),
    );
    bool_attr(output, "table:contains-header", range.contains_header);
    bool_attr(
        output,
        "table:display-filter-buttons",
        range.display_filter_buttons,
    );
    attr(
        output,
        "table:target-range-address",
        Some(&range.target_range_address),
    );
    attr(
        output,
        "table:refresh-delay",
        range.refresh_delay.as_deref(),
    );
    if range.source.is_none()
        && range.filter.is_none()
        && range.sort.is_none()
        && range.subtotals.is_none()
    {
        output.push_str("/>");
        return;
    }
    output.push('>');
    if let Some(source) = &range.source {
        write_database_source(output, source);
    }
    if let Some(filter) = &range.filter {
        write_filter(output, filter);
    }
    if let Some(sort) = &range.sort {
        write_sort(output, sort);
    }
    if let Some(subtotals) = &range.subtotals {
        write_subtotals(output, subtotals);
    }
    output.push_str("</table:database-range>");
}

pub(crate) fn write_database_source(output: &mut String, source: &DatabaseSource) {
    match source {
        DatabaseSource::Sql {
            database_name,
            statement,
            parse_statement,
        } => {
            output.push_str("<table:database-source-sql");
            attr(output, "table:database-name", Some(database_name));
            attr(output, "table:sql-statement", Some(statement));
            bool_attr(output, "table:parse-sql-statement", *parse_statement);
        },
        DatabaseSource::Table {
            database_name,
            table_name,
        } => {
            output.push_str("<table:database-source-table");
            attr(output, "table:database-name", Some(database_name));
            attr(output, "table:database-table-name", Some(table_name));
        },
        DatabaseSource::Query {
            database_name,
            query_name,
        } => {
            output.push_str("<table:database-source-query");
            attr(output, "table:database-name", Some(database_name));
            attr(output, "table:query-name", Some(query_name));
        },
    }
    output.push_str("/>");
}

pub(crate) fn write_filter(output: &mut String, filter: &DatabaseFilter) {
    output.push_str("<table:filter");
    attr(
        output,
        "table:target-range-address",
        filter.target_range_address.as_deref(),
    );
    attr(
        output,
        "table:condition-source",
        filter.condition_source.map(FilterConditionSource::as_str),
    );
    attr(
        output,
        "table:condition-source-range-address",
        filter.condition_source_range_address.as_deref(),
    );
    bool_attr(
        output,
        "table:display-duplicates",
        filter.display_duplicates,
    );
    output.push('>');
    write_filter_expression(output, &filter.expression);
    output.push_str("</table:filter>");
}

fn write_filter_expression(output: &mut String, expression: &FilterExpression) {
    match expression {
        FilterExpression::Condition(condition) => {
            output.push_str("<table:filter-condition");
            u64_attr(output, "table:field-number", condition.field_number);
            attr(output, "table:value", Some(&condition.value));
            attr(output, "table:operator", Some(&condition.operator));
            bool_attr(output, "table:case-sensitive", condition.case_sensitive);
            attr(
                output,
                "table:data-type",
                condition.data_type.map(FilterDataType::as_str),
            );
            if condition.set_items.is_empty() {
                output.push_str("/>");
            } else {
                output.push('>');
                for item in &condition.set_items {
                    output.push_str("<table:filter-set-item");
                    attr(output, "table:value", Some(item));
                    output.push_str("/>");
                }
                output.push_str("</table:filter-condition>");
            }
        },
        FilterExpression::And(children) => {
            output.push_str("<table:filter-and>");
            children
                .iter()
                .for_each(|child| write_filter_expression(output, child));
            output.push_str("</table:filter-and>");
        },
        FilterExpression::Or(children) => {
            output.push_str("<table:filter-or>");
            children
                .iter()
                .for_each(|child| write_filter_expression(output, child));
            output.push_str("</table:filter-or>");
        },
    }
}

fn write_sort(output: &mut String, sort: &DatabaseSort) {
    output.push_str("<table:sort");
    bool_attr(
        output,
        "table:bind-styles-to-content",
        sort.bind_styles_to_content,
    );
    attr(
        output,
        "table:target-range-address",
        sort.target_range_address.as_deref(),
    );
    bool_attr(output, "table:case-sensitive", sort.case_sensitive);
    attr(output, "table:language", sort.language.as_deref());
    attr(output, "table:country", sort.country.as_deref());
    attr(output, "table:script", sort.script.as_deref());
    attr(
        output,
        "table:rfc-language-tag",
        sort.rfc_language_tag.as_deref(),
    );
    attr(output, "table:algorithm", sort.algorithm.as_deref());
    attr(
        output,
        "table:embedded-number-behavior",
        sort.embedded_number_behavior
            .map(EmbeddedNumberBehavior::as_str),
    );
    output.push('>');
    for key in &sort.keys {
        output.push_str("<table:sort-by");
        u64_attr(output, "table:field-number", key.field_number);
        attr(output, "table:data-type", key.data_type.as_deref());
        attr(output, "table:order", key.order.map(SortOrder::as_str));
        output.push_str("/>");
    }
    output.push_str("</table:sort>");
}

fn write_subtotals(output: &mut String, subtotals: &SubtotalRules) {
    output.push_str("<table:subtotal-rules");
    bool_attr(
        output,
        "table:bind-styles-to-content",
        subtotals.bind_styles_to_content,
    );
    bool_attr(output, "table:case-sensitive", subtotals.case_sensitive);
    bool_attr(
        output,
        "table:page-breaks-on-group-change",
        subtotals.page_breaks_on_group_change,
    );
    if subtotals.sort_groups.is_none() && subtotals.rules.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    if let Some(groups) = &subtotals.sort_groups {
        output.push_str("<table:sort-groups");
        attr(output, "table:data-type", groups.data_type.as_deref());
        attr(output, "table:order", groups.order.map(SortOrder::as_str));
        output.push_str("/>");
    }
    for rule in &subtotals.rules {
        output.push_str("<table:subtotal-rule");
        u64_attr(
            output,
            "table:group-by-field-number",
            rule.group_by_field_number,
        );
        if rule.fields.is_empty() {
            output.push_str("/>");
            continue;
        }
        output.push('>');
        for field in &rule.fields {
            output.push_str("<table:subtotal-field");
            u64_attr(output, "table:field-number", field.field_number);
            attr(output, "table:function", Some(&field.function));
            output.push_str("/>");
        }
        output.push_str("</table:subtotal-rule>");
    }
    output.push_str("</table:subtotal-rules>");
}

fn optional_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid database-range attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == TABLE_NAMESPACE)
            && local.as_ref() == local_name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid database-range attribute value: {error}"))
                });
        }
    }
    Ok(None)
}

fn required_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    optional_attr(reader, element, local_name)?
        .ok_or_else(|| missing(&format!("table:{}", String::from_utf8_lossy(local_name))))
}

fn optional_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<bool>> {
    optional_attr(reader, element, local_name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid(
                &format!("table:{}", String::from_utf8_lossy(local_name)),
                &value,
            )),
        })
        .transpose()
}

fn required_u64(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<u64> {
    let value = required_attr(reader, element, local_name)?;
    value.parse().map_err(|_| {
        invalid(
            &format!("table:{}", String::from_utf8_lossy(local_name)),
            &value,
        )
    })
}

fn is_table(namespace: &ResolveResult<'_>, element: &impl HasLocalName, local: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE)
        && element.has_local_name(local)
}

fn ensure_absent<T>(value: &Option<T>, what: &str) -> Result<()> {
    if value.is_some() {
        Err(Error::InvalidFormat(format!("duplicate {what}")))
    } else {
        Ok(())
    }
}

fn attr(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

fn bool_attr(output: &mut String, name: &str, value: Option<bool>) {
    attr(
        output,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn u64_attr(output: &mut String, name: &str, value: u64) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&value.to_string());
    output.push('"');
}

fn is_xsd_duration(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return false;
    }
    index += 1;
    let mut any = false;
    any |= consume_integer_component(bytes, &mut index, b'Y');
    any |= consume_integer_component(bytes, &mut index, b'M');
    any |= consume_integer_component(bytes, &mut index, b'D');
    if bytes.get(index) == Some(&b'T') {
        index += 1;
        let mut any_time = false;
        any_time |= consume_integer_component(bytes, &mut index, b'H');
        any_time |= consume_integer_component(bytes, &mut index, b'M');
        any_time |= consume_seconds(bytes, &mut index);
        if !any_time {
            return false;
        }
        any = true;
    }
    any && index == bytes.len()
}

fn consume_integer_component(bytes: &[u8], index: &mut usize, suffix: u8) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end > *index && bytes.get(end) == Some(&suffix) {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn consume_seconds(bytes: &[u8], index: &mut usize) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end == *index {
        return false;
    }
    if bytes.get(end) == Some(&b'.') {
        end += 1;
        let start = end;
        while bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
        if end == start {
            return false;
        }
    }
    if bytes.get(end) == Some(&b'S') {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn invalid(attribute: &str, value: &str) -> Error {
    Error::InvalidFormat(format!("invalid {attribute} value '{value}'"))
}

fn missing(attribute: &str) -> Error {
    Error::InvalidFormat(format!("missing required {attribute}"))
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("database-range XML parsing error: {error}"))
}

fn unexpected_eof(element: &str) -> Error {
    Error::InvalidFormat(format!("unexpected end of XML inside {element}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{MutableSpreadsheet, Spreadsheet, SpreadsheetBuilder};

    #[test]
    fn parses_and_writes_complete_database_range_metadata() {
        let xml = r##"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:name="Data &amp; More" t:is-selection="false" t:on-update-keep-styles="true" t:on-update-keep-size="false" t:has-persistent-data="true" t:orientation="column" t:contains-header="true" t:display-filter-buttons="true" t:target-range-address="Sheet1.A1:Sheet1.D20" t:refresh-delay="PT5M"><t:database-source-sql t:database-name="db&amp;1" t:sql-statement="SELECT &lt;x&gt;" t:parse-sql-statement="true"/><t:filter t:target-range-address="Sheet1.F1:Sheet1.I20" t:condition-source="cell-range" t:condition-source-range-address="Sheet2.A1:Sheet2.B2" t:display-duplicates="false"><t:filter-and><t:filter-condition t:field-number="0" t:value="alpha" t:operator="=" t:case-sensitive="true" t:data-type="text"/><t:filter-or><t:filter-condition t:field-number="1" t:value="10" t:operator=">=" t:data-type="number"/><t:filter-condition t:field-number="2" t:value="" t:operator="in"><t:filter-set-item t:value="A&amp;B"/><t:filter-set-item t:value="C"/></t:filter-condition></t:filter-or></t:filter-and></t:filter><t:sort t:bind-styles-to-content="true" t:target-range-address="Sheet1.A2:Sheet1.D20" t:case-sensitive="false" t:language="en" t:country="US" t:script="Latn" t:rfc-language-tag="en-US" t:algorithm="unicode" t:embedded-number-behavior="integer"><t:sort-by t:field-number="1" t:data-type="number" t:order="descending"/></t:sort><t:subtotal-rules t:bind-styles-to-content="false" t:case-sensitive="true" t:page-breaks-on-group-change="true"><t:sort-groups t:data-type="text" t:order="ascending"/><t:subtotal-rule t:group-by-field-number="0"><t:subtotal-field t:field-number="3" t:function="sum"/></t:subtotal-rule></t:subtotal-rules></t:database-range></t:database-ranges></o:spreadsheet>"##;
        let parsed = parse_database_ranges(xml).unwrap();
        assert_eq!(parsed.len(), 1);
        let range = &parsed[0];
        assert_eq!(range.name.as_deref(), Some("Data & More"));
        assert_eq!(range.orientation, Some(DatabaseOrientation::Column));
        assert_eq!(range.refresh_delay.as_deref(), Some("PT5M"));
        assert!(matches!(range.source, Some(DatabaseSource::Sql { .. })));
        assert_eq!(range.sort.as_ref().unwrap().keys[0].field_number, 1);
        assert_eq!(
            range.subtotals.as_ref().unwrap().rules[0].fields[0].function,
            "sum"
        );

        let mut written = String::new();
        write_database_ranges(&mut written, &parsed).unwrap();
        let reparsed = parse_database_ranges(&format!(
            r#"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{written}</o:spreadsheet>"#
        ))
        .unwrap();
        assert_eq!(reparsed, parsed);
        assert!(written.contains("Data &amp; More"));
        assert!(written.contains("SELECT &lt;x&gt;"));
    }

    #[test]
    fn rejects_invalid_filter_shapes_and_required_values() {
        let same_group = FilterExpression::And(vec![FilterExpression::And(vec![
            FilterExpression::Condition(FilterCondition::new(0, "=", "x")),
        ])]);
        assert!(validate_filter_expression(&same_group, 0, None).is_err());

        let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="A1:B2"><t:sort/></t:database-range></t:database-ranges></s>"#;
        assert!(parse_database_ranges(xml).is_err());
    }

    #[test]
    fn external_database_sources_remain_inert_data() {
        let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="A1"><t:database-source-query t:database-name="file:///database.odb" t:query-name="DangerousQuery"/></t:database-range></t:database-ranges></s>"#;
        let ranges = parse_database_ranges(xml).unwrap();
        assert_eq!(
            ranges[0].source,
            Some(DatabaseSource::Query {
                database_name: "file:///database.odb".to_string(),
                query_name: "DangerousQuery".to_string(),
            })
        );
    }

    #[test]
    fn database_ranges_round_trip_through_builder_and_mutable_packages() {
        let mut range = DatabaseRange::new("Sheet1.A1:Sheet1.C20");
        range.name = Some("Sales".to_string());
        range.orientation = Some(DatabaseOrientation::Column);
        range.display_filter_buttons = Some(true);
        range.source = Some(DatabaseSource::Query {
            database_name: "file:///sales&forecast.odb".to_string(),
            query_name: "Quarter <One>".to_string(),
        });
        range.filter = Some(DatabaseFilter {
            target_range_address: None,
            condition_source: Some(FilterConditionSource::SelfContained),
            condition_source_range_address: None,
            display_duplicates: Some(false),
            expression: FilterExpression::And(vec![
                FilterExpression::Condition(FilterCondition::new(0, "=", "East")),
                FilterExpression::Or(vec![
                    FilterExpression::Condition(FilterCondition::new(1, ">", "100")),
                    FilterExpression::Condition(FilterCondition::new(2, "=", "Open")),
                ]),
            ]),
        });
        range.sort = Some(DatabaseSort {
            embedded_number_behavior: Some(EmbeddedNumberBehavior::Integer),
            keys: vec![DatabaseSortKey {
                field_number: 1,
                data_type: Some("number".to_string()),
                order: Some(SortOrder::Descending),
            }],
            ..DatabaseSort::default()
        });
        range.subtotals = Some(SubtotalRules {
            rules: vec![SubtotalRule {
                group_by_field_number: 0,
                fields: vec![SubtotalField {
                    field_number: 1,
                    function: "sum".to_string(),
                }],
            }],
            ..SubtotalRules::default()
        });

        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_database_range(range.clone()).unwrap();
        let bytes = builder.build().unwrap();
        let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
        assert_eq!(spreadsheet.database_ranges(), &[range.clone()]);

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        let FilterExpression::And(expressions) = &mut mutable.database_ranges_mut()[0]
            .filter
            .as_mut()
            .unwrap()
            .expression
        else {
            panic!("expected AND filter")
        };
        let FilterExpression::Condition(condition) = &mut expressions[0] else {
            panic!("expected filter condition")
        };
        condition.value = "West & Central".to_string();

        let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let reopened_range = &reopened.database_ranges()[0];
        let FilterExpression::And(expressions) =
            &reopened_range.filter.as_ref().unwrap().expression
        else {
            panic!("expected AND filter")
        };
        let FilterExpression::Condition(condition) = &expressions[0] else {
            panic!("expected filter condition")
        };
        assert_eq!(condition.value, "West & Central");
        assert_eq!(reopened_range.source, range.source);
        assert_eq!(reopened_range.sort, range.sort);
        assert_eq!(reopened_range.subtotals, range.subtotals);
    }
}
