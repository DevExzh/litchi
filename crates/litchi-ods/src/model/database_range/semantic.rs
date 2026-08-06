//! Typed ODF spreadsheet database-range vocabulary and ergonomic constructors.

/// Whether database fields are arranged in columns or rows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Orientation {
    /// Each field occupies a column.
    Column,
    /// Each field occupies a row.
    Row,
}

/// An inert external database source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
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
pub enum Order {
    /// Ascending order.
    Ascending,
    /// Descending order.
    Descending,
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

/// One field in a database-range sort specification.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Key {
    /// Zero-based field number.
    pub field_number: u64,
    /// Standard or application-defined sort data type.
    pub data_type: Option<String>,
    /// Optional explicit order.
    pub order: Option<Order>,
}

impl Key {
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
pub struct Sort {
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
    pub keys: Vec<Key>,
}

/// Source used to obtain filter conditions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionSource {
    /// Conditions are contained in the filter itself.
    SelfContained,
    /// Conditions come from another cell range.
    CellRange,
}

/// Standard filter comparison data type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataType {
    /// Text comparison.
    Text,
    /// Numeric comparison.
    Number,
}

/// A leaf filter comparison.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Condition {
    /// Zero-based field number.
    pub field_number: u64,
    /// Comparison value.
    pub value: String,
    /// Standard or application-defined operator.
    pub operator: String,
    /// Optional case-sensitivity override.
    pub case_sensitive: Option<bool>,
    /// Optional comparison data type.
    pub data_type: Option<DataType>,
    /// Values in a set-membership condition.
    pub set_items: Vec<String>,
}

impl Condition {
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
pub enum Expression {
    /// A leaf comparison.
    Condition(Condition),
    /// All child expressions must match. Children may be conditions or OR groups.
    And(Vec<Expression>),
    /// At least one child expression must match. Children may be conditions or AND groups.
    Or(Vec<Expression>),
}

/// Filter configuration attached to a database range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Filter {
    /// Optional destination range.
    pub target_range_address: Option<String>,
    /// Optional condition source.
    pub condition_source: Option<ConditionSource>,
    /// Range containing conditions when `condition_source` is `CellRange`.
    pub condition_source_range_address: Option<String>,
    /// Whether duplicate rows remain visible.
    pub display_duplicates: Option<bool>,
    /// Root filter expression.
    pub expression: Expression,
}

/// Sort configuration for subtotal groups.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SortGroups {
    /// Standard or application-defined data type.
    pub data_type: Option<String>,
    /// Optional explicit order.
    pub order: Option<Order>,
}

/// A field aggregated by a subtotal rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Field {
    /// Zero-based field number.
    pub field_number: u64,
    /// Standard or application-defined aggregation function.
    pub function: String,
}

/// One subtotal grouping rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rule {
    /// Zero-based grouping field number.
    pub group_by_field_number: u64,
    /// Fields aggregated for the group.
    pub fields: Vec<Field>,
}

/// Subtotal rules attached to a database range.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Rules {
    /// Whether styles remain bound to content.
    pub bind_styles_to_content: Option<bool>,
    /// Whether grouping comparisons are case-sensitive.
    pub case_sensitive: Option<bool>,
    /// Whether to insert page breaks when a group changes.
    pub page_breaks_on_group_change: Option<bool>,
    /// Optional group sorting.
    pub sort_groups: Option<SortGroups>,
    /// Ordered subtotal rules.
    pub rules: Vec<Rule>,
}

/// A spreadsheet database range and its non-executing query/filter metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
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
    pub orientation: Option<Orientation>,
    /// Whether the first field is a header.
    pub contains_header: Option<bool>,
    /// Whether filter buttons are displayed.
    pub display_filter_buttons: Option<bool>,
    /// Required cell range occupied by the database range.
    pub target_range_address: String,
    /// Optional XML Schema refresh duration.
    pub refresh_delay: Option<String>,
    /// Optional inert external source.
    pub source: Option<Source>,
    /// Optional filter.
    pub filter: Option<Filter>,
    /// Optional sorting.
    pub sort: Option<Sort>,
    /// Optional subtotal rules.
    pub subtotals: Option<Rules>,
}

impl Range {
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
}
