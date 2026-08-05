//! Typed data-pilot vocabulary and ergonomic constructors.

use crate::model::database_range::{Filter, Source as DatabaseSource};
use litchi_core::Result;

use super::invalid;

macro_rules! string_enum {
    ($(#[$meta:meta])* $vis:vis enum $name:ident { $($(#[$variant_meta:meta])* $variant:ident => $value:literal),+ $(,)? }) => {
        $(#[$meta])*
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        $vis enum $name { $($(#[$variant_meta])* $variant),+ }

        impl $name {
            pub(super) fn parse(value: &str) -> Result<Self> {
                match value {
                    $($value => Ok(Self::$variant),)+
                    _ => Err(invalid(stringify!($name), value)),
                }
            }

            pub(crate) const fn as_str(self) -> &'static str {
                match self { $(Self::$variant => $value,)+ }
            }
        }
    };
}

string_enum! {
    /// Which grand totals a data-pilot table displays.
    pub enum GrandTotal {
        None => "none", Row => "row", Column => "column", Both => "both"
    }
}

string_enum! {
    /// Orientation of a LibreOffice named grand-total extension element.
    pub enum GrandTotalOrientation {
        Row => "row", Column => "column", Both => "both"
    }
}

/// LibreOffice's inert named grand-total extension metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GrandTotalElement {
    pub orientation: GrandTotalOrientation,
    pub display: bool,
    pub display_name: Option<String>,
}

string_enum! {
    /// Placement of a data-pilot field.
    pub enum Orientation {
        Row => "row", Column => "column", Data => "data", Hidden => "hidden", Page => "page"
    }
}

string_enum! {
    /// Automatic member display direction.
    pub enum DisplayMemberMode {
        FromTop => "from-top", FromBottom => "from-bottom"
    }
}

string_enum! {
    /// Member sorting policy.
    pub enum SortMode {
        None => "none", Manual => "manual", Name => "name", Data => "data"
    }
}

string_enum! {
    /// Sort direction.
    pub enum SortOrder {
        Ascending => "ascending", Descending => "descending"
    }
}

string_enum! {
    /// Field layout policy.
    pub enum LayoutMode {
        Tabular => "tabular-layout",
        OutlineSubtotalsTop => "outline-subtotals-top",
        OutlineSubtotalsBottom => "outline-subtotals-bottom"
    }
}

string_enum! {
    /// The member used by a data-field reference.
    pub enum ReferenceMemberType {
        Named => "named", Previous => "previous", Next => "next"
    }
}

string_enum! {
    /// Calculation performed relative to another pivot member or total.
    pub enum ReferenceType {
        None => "none",
        MemberDifference => "member-difference",
        MemberPercentage => "member-percentage",
        MemberPercentageDifference => "member-percentage-difference",
        RunningTotal => "running-total",
        RowPercentage => "row-percentage",
        ColumnPercentage => "column-percentage",
        TotalPercentage => "total-percentage",
        Index => "index"
    }
}

string_enum! {
    /// Calendar unit used for automatic grouping.
    pub enum GroupBy {
        Seconds => "seconds", Minutes => "minutes", Hours => "hours", Days => "days",
        Months => "months", Quarters => "quarters", Years => "years"
    }
}

/// An inclusive automatic-grouping boundary.
#[derive(Clone, Debug, PartialEq)]
pub enum GroupBoundary {
    /// Let the spreadsheet consumer determine a numeric boundary.
    AutomaticNumber,
    /// Let the spreadsheet consumer determine a date boundary.
    AutomaticDate,
    /// Numeric boundary.
    Number(f64),
    /// ISO date or date-time boundary retained verbatim.
    Date(String),
}

/// An inert data source for a data-pilot table.
#[derive(Clone, Debug, PartialEq)]
pub enum Source {
    /// SQL, database table, or database query metadata. Litchi never executes it.
    Database(DatabaseSource),
    /// Application service metadata. Litchi never invokes the service.
    Service {
        name: String,
        source_name: String,
        object_name: String,
        user_name: Option<String>,
        password: Option<String>,
    },
    /// A spreadsheet range and optional filter.
    CellRange {
        /// Optional ODF 1.3 named-range source identifier.
        name: Option<String>,
        cell_range_address: String,
        filter: Option<Filter>,
    },
}

/// One explicitly grouped pivot member.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Group {
    pub name: String,
    pub members: Vec<String>,
}

/// Grouping configuration for a pivot field.
#[derive(Clone, Debug, PartialEq)]
pub struct Groups {
    pub source_field_name: String,
    pub start: GroupBoundary,
    pub end: GroupBoundary,
    pub step: f64,
    pub grouped_by: GroupBy,
    pub groups: Vec<Group>,
}

/// Visibility settings for one pivot member.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Member {
    pub name: String,
    pub display: Option<bool>,
    pub show_details: Option<bool>,
}

/// Automatic top/bottom member display settings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DisplayInfo {
    pub enabled: bool,
    pub data_field: String,
    pub member_count: u64,
    pub mode: DisplayMemberMode,
}

/// Sort settings for a field level.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortInfo {
    pub mode: SortMode,
    pub data_field: Option<String>,
    pub order: SortOrder,
}

/// Layout settings for a field level.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LayoutInfo {
    pub mode: LayoutMode,
    pub add_empty_lines: bool,
}

/// Member presentation and aggregation details for a pivot field.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Level {
    pub show_empty: Option<bool>,
    /// LibreOffice `calcext:repeat-item-labels`; retained but never evaluated.
    pub repeat_item_labels: Option<bool>,
    /// Standard or implementation-defined aggregation names.
    pub subtotals: Vec<String>,
    pub members: Vec<Member>,
    pub display: Option<DisplayInfo>,
    pub sort: Option<SortInfo>,
    pub layout: Option<LayoutInfo>,
}

/// Relative calculation settings for a data field.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FieldReference {
    pub field_name: String,
    pub member_type: ReferenceMemberType,
    pub member_name: Option<String>,
    pub reference_type: ReferenceType,
}

/// One field in a data-pilot table.
#[derive(Clone, Debug, PartialEq)]
pub struct Field {
    pub source_field_name: String,
    pub orientation: Orientation,
    pub selected_page: Option<String>,
    pub is_data_layout_field: Option<String>,
    /// Standard or implementation-defined aggregation name.
    pub function: Option<String>,
    pub used_hierarchy: Option<i64>,
    pub level: Option<Level>,
    pub reference: Option<FieldReference>,
    pub groups: Option<Groups>,
}

impl Field {
    /// Create a field in the requested orientation.
    pub fn new(source_field_name: impl Into<String>, orientation: Orientation) -> Self {
        Self {
            source_field_name: source_field_name.into(),
            orientation,
            selected_page: None,
            is_data_layout_field: None,
            function: None,
            used_hierarchy: None,
            level: None,
            reference: None,
            groups: None,
        }
    }
}

/// A complete ODF data-pilot (pivot-table) declaration.
#[derive(Clone, Debug, PartialEq)]
pub struct Table {
    pub name: String,
    pub application_data: Option<String>,
    pub grand_total: Option<GrandTotal>,
    pub ignore_empty_rows: Option<bool>,
    pub identify_categories: Option<bool>,
    pub target_range_address: String,
    pub buttons: Option<String>,
    pub show_filter_button: Option<bool>,
    pub drill_down_on_double_click: Option<bool>,
    /// LibreOffice named grand-total extension elements in schema position.
    pub grand_totals: Vec<GrandTotalElement>,
    pub source: Option<Source>,
    pub fields: Vec<Field>,
}

impl Table {
    /// Create a pivot declaration. At least one field must be added before writing.
    pub fn new(name: impl Into<String>, target_range_address: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            application_data: None,
            grand_total: None,
            ignore_empty_rows: None,
            identify_categories: None,
            target_range_address: target_range_address.into(),
            buttons: None,
            show_filter_button: None,
            drill_down_on_double_click: None,
            grand_totals: Vec::new(),
            source: None,
            fields: Vec::new(),
        }
    }
}
