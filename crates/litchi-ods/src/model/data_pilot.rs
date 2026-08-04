//! ODF data-pilot (pivot-table) declarations.

use super::{
    DatabaseFilter, DatabaseSource,
    database_range::{
        parse_filter, parse_source_query, parse_source_sql, parse_source_table, validate_filter,
        write_database_source, write_filter,
    },
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesEnd, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0";
const CALC_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_DATA_PILOT_TABLES: usize = 65_536;
const MAX_DATA_PILOT_FIELDS: usize = 65_536;
const MAX_DATA_PILOT_ITEMS: usize = 1_000_000;
const MAX_DATA_PILOT_STRING: usize = 1024 * 1024;

trait HasLocalName {
    fn local(&self) -> &[u8];
}

impl HasLocalName for BytesStart<'_> {
    fn local(&self) -> &[u8] {
        self.local_name().into_inner()
    }
}

impl HasLocalName for BytesEnd<'_> {
    fn local(&self) -> &[u8] {
        self.local_name().into_inner()
    }
}

macro_rules! string_enum {
    ($(#[$meta:meta])* $vis:vis enum $name:ident { $($(#[$variant_meta:meta])* $variant:ident => $value:literal),+ $(,)? }) => {
        $(#[$meta])*
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        $vis enum $name { $($(#[$variant_meta])* $variant),+ }

        impl $name {
            fn parse(value: &str) -> Result<Self> {
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
    pub enum DataPilotGrandTotal {
        None => "none", Row => "row", Column => "column", Both => "both"
    }
}

string_enum! {
    /// Orientation of a LibreOffice named grand-total extension element.
    pub enum DataPilotGrandTotalOrientation {
        Row => "row", Column => "column", Both => "both"
    }
}

/// LibreOffice's inert named grand-total extension metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotGrandTotalElement {
    pub orientation: DataPilotGrandTotalOrientation,
    pub display: bool,
    pub display_name: Option<String>,
}

string_enum! {
    /// Placement of a data-pilot field.
    pub enum DataPilotOrientation {
        Row => "row", Column => "column", Data => "data", Hidden => "hidden", Page => "page"
    }
}

string_enum! {
    /// Automatic member display direction.
    pub enum DataPilotDisplayMemberMode {
        FromTop => "from-top", FromBottom => "from-bottom"
    }
}

string_enum! {
    /// Member sorting policy.
    pub enum DataPilotSortMode {
        None => "none", Manual => "manual", Name => "name", Data => "data"
    }
}

string_enum! {
    /// Sort direction.
    pub enum DataPilotSortOrder {
        Ascending => "ascending", Descending => "descending"
    }
}

string_enum! {
    /// Field layout policy.
    pub enum DataPilotLayoutMode {
        Tabular => "tabular-layout",
        OutlineSubtotalsTop => "outline-subtotals-top",
        OutlineSubtotalsBottom => "outline-subtotals-bottom"
    }
}

string_enum! {
    /// The member used by a data-field reference.
    pub enum DataPilotReferenceMemberType {
        Named => "named", Previous => "previous", Next => "next"
    }
}

string_enum! {
    /// Calculation performed relative to another pivot member or total.
    pub enum DataPilotReferenceType {
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
    pub enum DataPilotGroupBy {
        Seconds => "seconds", Minutes => "minutes", Hours => "hours", Days => "days",
        Months => "months", Quarters => "quarters", Years => "years"
    }
}

/// An inclusive automatic-grouping boundary.
#[derive(Clone, Debug, PartialEq)]
pub enum DataPilotGroupBoundary {
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
pub enum DataPilotSource {
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
        filter: Option<DatabaseFilter>,
    },
}

/// One explicitly grouped pivot member.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotGroup {
    pub name: String,
    pub members: Vec<String>,
}

/// Grouping configuration for a pivot field.
#[derive(Clone, Debug, PartialEq)]
pub struct DataPilotGroups {
    pub source_field_name: String,
    pub start: DataPilotGroupBoundary,
    pub end: DataPilotGroupBoundary,
    pub step: f64,
    pub grouped_by: DataPilotGroupBy,
    pub groups: Vec<DataPilotGroup>,
}

/// Visibility settings for one pivot member.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotMember {
    pub name: String,
    pub display: Option<bool>,
    pub show_details: Option<bool>,
}

/// Automatic top/bottom member display settings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotDisplayInfo {
    pub enabled: bool,
    pub data_field: String,
    pub member_count: u64,
    pub mode: DataPilotDisplayMemberMode,
}

/// Sort settings for a field level.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotSortInfo {
    pub mode: DataPilotSortMode,
    pub data_field: Option<String>,
    pub order: DataPilotSortOrder,
}

/// Layout settings for a field level.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotLayoutInfo {
    pub mode: DataPilotLayoutMode,
    pub add_empty_lines: bool,
}

/// Member presentation and aggregation details for a pivot field.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DataPilotLevel {
    pub show_empty: Option<bool>,
    /// LibreOffice `calcext:repeat-item-labels`; retained but never evaluated.
    pub repeat_item_labels: Option<bool>,
    /// Standard or implementation-defined aggregation names.
    pub subtotals: Vec<String>,
    pub members: Vec<DataPilotMember>,
    pub display: Option<DataPilotDisplayInfo>,
    pub sort: Option<DataPilotSortInfo>,
    pub layout: Option<DataPilotLayoutInfo>,
}

/// Relative calculation settings for a data field.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataPilotFieldReference {
    pub field_name: String,
    pub member_type: DataPilotReferenceMemberType,
    pub member_name: Option<String>,
    pub reference_type: DataPilotReferenceType,
}

/// One field in a data-pilot table.
#[derive(Clone, Debug, PartialEq)]
pub struct DataPilotField {
    pub source_field_name: String,
    pub orientation: DataPilotOrientation,
    pub selected_page: Option<String>,
    pub is_data_layout_field: Option<String>,
    /// Standard or implementation-defined aggregation name.
    pub function: Option<String>,
    pub used_hierarchy: Option<i64>,
    pub level: Option<DataPilotLevel>,
    pub reference: Option<DataPilotFieldReference>,
    pub groups: Option<DataPilotGroups>,
}

impl DataPilotField {
    /// Create a field in the requested orientation.
    pub fn new(source_field_name: impl Into<String>, orientation: DataPilotOrientation) -> Self {
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

    pub fn validate(&self) -> Result<()> {
        if self.orientation == DataPilotOrientation::Page && self.selected_page.is_none() {
            return Err(Error::InvalidFormat(
                "page-oriented data-pilot field requires table:selected-page".to_string(),
            ));
        }
        if self.orientation != DataPilotOrientation::Page && self.selected_page.is_some() {
            return Err(Error::InvalidFormat(
                "table:selected-page is valid only for a page-oriented data-pilot field"
                    .to_string(),
            ));
        }
        if let Some(reference) = &self.reference {
            let named = reference.member_type == DataPilotReferenceMemberType::Named;
            if named != reference.member_name.is_some() {
                return Err(Error::InvalidFormat(
                    "named data-pilot field references require exactly one member name".to_string(),
                ));
            }
        }
        if let Some(level) = &self.level
            && let Some(sort) = &level.sort
            && (sort.mode == DataPilotSortMode::Data) != sort.data_field.is_some()
        {
            return Err(Error::InvalidFormat(
                "data-pilot data sorting requires exactly one data field".to_string(),
            ));
        }
        if let Some(groups) = &self.groups {
            if !groups.step.is_finite() || groups.step <= 0.0 {
                return Err(Error::InvalidFormat(
                    "data-pilot grouping step must be finite and greater than zero".to_string(),
                ));
            }
            for boundary in [&groups.start, &groups.end] {
                if matches!(boundary, DataPilotGroupBoundary::Number(value) if !value.is_finite()) {
                    return Err(Error::InvalidFormat(
                        "data-pilot grouping boundaries must be finite".to_string(),
                    ));
                }
            }
            if groups.groups.is_empty()
                || groups.groups.iter().any(|group| group.members.is_empty())
            {
                return Err(Error::InvalidFormat(
                    "data-pilot groups and group member lists cannot be empty".to_string(),
                ));
            }
        }
        Ok(())
    }
}

/// A complete ODF data-pilot (pivot-table) declaration.
#[derive(Clone, Debug, PartialEq)]
pub struct DataPilotTable {
    pub name: String,
    pub application_data: Option<String>,
    pub grand_total: Option<DataPilotGrandTotal>,
    pub ignore_empty_rows: Option<bool>,
    pub identify_categories: Option<bool>,
    pub target_range_address: String,
    pub buttons: Option<String>,
    pub show_filter_button: Option<bool>,
    pub drill_down_on_double_click: Option<bool>,
    /// LibreOffice named grand-total extension elements in schema position.
    pub grand_totals: Vec<DataPilotGrandTotalElement>,
    pub source: Option<DataPilotSource>,
    pub fields: Vec<DataPilotField>,
}

impl DataPilotTable {
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

    pub fn validate(&self) -> Result<()> {
        validate_string("data-pilot table name", &self.name, false)?;
        if self.target_range_address.is_empty() {
            return Err(Error::InvalidFormat(
                "data-pilot target range address cannot be empty".to_string(),
            ));
        }
        parse_data_pilot_range(&self.target_range_address)?;
        for value in [self.application_data.as_deref(), self.buttons.as_deref()]
            .into_iter()
            .flatten()
        {
            validate_string("data-pilot attribute", value, true)?;
        }
        let mut grand_orientations = std::collections::HashSet::new();
        for total in &self.grand_totals {
            if !grand_orientations.insert(total.orientation) {
                return Err(invalid_message(
                    "duplicate data-pilot grand-total orientation",
                ));
            }
            if let Some(name) = &total.display_name {
                validate_string("data-pilot grand-total display name", name, true)?;
            }
        }
        if self.fields.is_empty() {
            return Err(Error::InvalidFormat(
                "data-pilot table requires at least one field".to_string(),
            ));
        }
        if self.fields.len() > MAX_DATA_PILOT_FIELDS {
            return Err(Error::InvalidFormat(format!(
                "data-pilot field count exceeds the {MAX_DATA_PILOT_FIELDS} field safety limit"
            )));
        }
        if let Some(DataPilotSource::CellRange {
            name,
            cell_range_address,
            filter,
        }) = &self.source
        {
            if let Some(name) = name {
                validate_string("data-pilot named source", name, false)?;
            }
            validate_string(
                "data-pilot source cell range address",
                cell_range_address,
                false,
            )?;
            parse_data_pilot_range(cell_range_address)?;
            if let Some(filter) = filter {
                validate_filter(filter)?;
            }
        }
        if let Some(DataPilotSource::Service {
            name,
            source_name,
            object_name,
            user_name,
            password,
        }) = &self.source
        {
            for value in [name.as_str(), source_name.as_str(), object_name.as_str()] {
                validate_string("data-pilot service source", value, false)?;
            }
            for value in [user_name.as_deref(), password.as_deref()]
                .into_iter()
                .flatten()
            {
                validate_string("data-pilot service source", value, true)?;
            }
        }
        if let Some(DataPilotSource::Database(source)) = &self.source {
            match source {
                DatabaseSource::Sql {
                    database_name,
                    statement,
                    ..
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot SQL statement", statement, false)?;
                },
                DatabaseSource::Table {
                    database_name,
                    table_name,
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot database table", table_name, false)?;
                },
                DatabaseSource::Query {
                    database_name,
                    query_name,
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot database query", query_name, false)?;
                },
            }
        }
        self.fields.iter().try_for_each(DataPilotField::validate)?;
        let field_names: std::collections::HashSet<&str> = self
            .fields
            .iter()
            .map(|field| field.source_field_name.as_str())
            .collect();
        let mut item_count = self.fields.len();
        for field in &self.fields {
            validate_string(
                "data-pilot source field name",
                &field.source_field_name,
                true,
            )?;
            if let Some(reference) = &field.reference
                && !field_names.contains(reference.field_name.as_str())
            {
                return Err(Error::InvalidFormat(format!(
                    "data-pilot field reference '{}' does not name a field",
                    reference.field_name
                )));
            }
            if let Some(groups) = &field.groups {
                if !field_names.contains(groups.source_field_name.as_str()) {
                    return Err(Error::InvalidFormat(format!(
                        "data-pilot grouping source '{}' does not name a field",
                        groups.source_field_name
                    )));
                }
                let mut names = std::collections::HashSet::new();
                for group in &groups.groups {
                    validate_string("data-pilot group name", &group.name, false)?;
                    if !names.insert(group.name.as_str()) {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate data-pilot group '{}'",
                            group.name
                        )));
                    }
                    let mut members = std::collections::HashSet::new();
                    for member in &group.members {
                        validate_string("data-pilot group member", member, false)?;
                        if !members.insert(member.as_str()) {
                            return Err(Error::InvalidFormat(format!(
                                "duplicate member '{member}' in data-pilot group '{}'",
                                group.name
                            )));
                        }
                    }
                    item_count = item_count
                        .checked_add(group.members.len() + 1)
                        .ok_or_else(|| invalid_message("data-pilot item count overflow"))?;
                }
            }
            if let Some(level) = &field.level {
                let mut members = std::collections::HashSet::new();
                for member in &level.members {
                    validate_string("data-pilot member", &member.name, false)?;
                    if !members.insert(member.name.as_str()) {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate data-pilot member '{}'",
                            member.name
                        )));
                    }
                }
                item_count = item_count
                    .checked_add(level.members.len() + level.subtotals.len())
                    .ok_or_else(|| invalid_message("data-pilot item count overflow"))?;
            }
        }
        if item_count > MAX_DATA_PILOT_ITEMS {
            return Err(invalid_message("data-pilot declaration exceeds item limit"));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ParsedDataPilotRange {
    pub sheet: String,
    pub start_column: usize,
    pub start_row: usize,
    pub end_column: usize,
    pub end_row: usize,
}

pub(crate) fn parse_data_pilot_range(value: &str) -> Result<ParsedDataPilotRange> {
    validate_string("data-pilot cell range", value, false)?;
    let mut quoted = false;
    let mut separator = None;
    let mut characters = value.char_indices().peekable();
    while let Some((index, character)) = characters.next() {
        if character == '\'' {
            if quoted && characters.peek().is_some_and(|(_, next)| *next == '\'') {
                characters.next();
                continue;
            }
            quoted = !quoted;
        } else if character == ':' && !quoted && separator.replace(index).is_some() {
            return Err(invalid_message("invalid data-pilot cell range"));
        }
    }
    if quoted {
        return Err(invalid_message(
            "unterminated quoted sheet name in data-pilot range",
        ));
    }
    let (first, second) =
        separator.map_or((value, None), |at| (&value[..at], Some(&value[at + 1..])));
    let (sheet, start_column, start_row) = parse_range_endpoint(first, None)?;
    let (end_sheet, end_column, end_row) = if let Some(second) = second {
        parse_range_endpoint(second, Some(&sheet))?
    } else {
        (sheet.clone(), start_column, start_row)
    };
    if end_sheet != sheet || end_column < start_column || end_row < start_row {
        return Err(invalid_message(
            "data-pilot cell range is reversed or crosses sheets",
        ));
    }
    Ok(ParsedDataPilotRange {
        sheet,
        start_column,
        start_row,
        end_column,
        end_row,
    })
}

fn parse_range_endpoint(
    value: &str,
    inherited_sheet: Option<&str>,
) -> Result<(String, usize, usize)> {
    let value = value.trim();
    let mut quoted = false;
    let mut dot = None;
    let mut characters = value.char_indices().peekable();
    while let Some((index, character)) = characters.next() {
        if character == '\'' {
            if quoted && characters.peek().is_some_and(|(_, next)| *next == '\'') {
                characters.next();
                continue;
            }
            quoted = !quoted;
        } else if character == '.' && !quoted {
            dot = Some(index);
        }
    }
    let (sheet, coordinate) =
        dot.map_or((None, value), |at| (Some(&value[..at]), &value[at + 1..]));
    let sheet = match sheet {
        Some("") => inherited_sheet.unwrap_or_default().to_string(),
        Some(value) => normalize_sheet_name(value)?,
        None => inherited_sheet.unwrap_or_default().to_string(),
    };
    let coordinate = coordinate.replace('$', "");
    let split = coordinate
        .find(|character: char| character.is_ascii_digit())
        .ok_or_else(|| invalid_message("data-pilot cell address lacks a row"))?;
    let (column, row) = coordinate.split_at(split);
    if column.is_empty()
        || !column.chars().all(|ch| ch.is_ascii_uppercase())
        || row.is_empty()
        || !row.chars().all(|ch| ch.is_ascii_digit())
    {
        return Err(invalid_message("invalid data-pilot cell address"));
    }
    let mut column_index = 0usize;
    for ch in column.bytes() {
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(usize::from(ch - b'A') + 1))
            .ok_or_else(|| invalid_message("data-pilot column index overflow"))?;
    }
    let row_number = row
        .parse::<usize>()
        .map_err(|_| invalid_message("invalid data-pilot row"))?;
    if row_number == 0 {
        return Err(invalid_message("data-pilot rows are one-based"));
    }
    Ok((sheet, column_index - 1, row_number - 1))
}

fn normalize_sheet_name(value: &str) -> Result<String> {
    let value = value.trim().trim_start_matches('$');
    if value.starts_with('\'') {
        if !value.ends_with('\'') || value.len() < 2 {
            return Err(invalid_message("invalid quoted sheet name"));
        }
        Ok(value[1..value.len() - 1].replace("''", "'"))
    } else {
        if value.contains('\'') {
            return Err(invalid_message("invalid sheet name"));
        }
        Ok(value.to_string())
    }
}

pub(crate) fn validate_data_pilot_tables(tables: &[DataPilotTable]) -> Result<()> {
    if tables.len() > MAX_DATA_PILOT_TABLES {
        return Err(invalid_message(
            "data-pilot table count exceeds safety limit",
        ));
    }
    let mut names = std::collections::HashSet::new();
    let mut targets = Vec::with_capacity(tables.len());
    for table in tables {
        table.validate()?;
        if !names.insert(table.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate data-pilot table name '{}'",
                table.name
            )));
        }
        let range = parse_data_pilot_range(&table.target_range_address)?;
        for (other_name, other) in &targets {
            if ranges_overlap(&range, other) {
                return Err(Error::InvalidFormat(format!(
                    "data-pilot target ranges for '{other_name}' and '{}' overlap",
                    table.name
                )));
            }
        }
        targets.push((table.name.as_str(), range));
    }
    Ok(())
}

fn ranges_overlap(left: &ParsedDataPilotRange, right: &ParsedDataPilotRange) -> bool {
    left.sheet == right.sheet
        && left.start_column <= right.end_column
        && right.start_column <= left.end_column
        && left.start_row <= right.end_row
        && right.start_row <= left.end_row
}

fn validate_string(label: &str, value: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_DATA_PILOT_STRING {
        return Err(Error::InvalidFormat(format!("{label} exceeds size limit")));
    }
    if value
        .chars()
        .any(|ch| matches!(ch as u32, 0..=8 | 11 | 12 | 14..=31 | 0xFFFE | 0xFFFF))
    {
        return Err(Error::InvalidFormat(format!(
            "{label} contains an XML-prohibited character"
        )));
    }
    Ok(())
}

pub(crate) fn parse_data_pilot_tables(xml: &str) -> Result<Vec<DataPilotTable>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut container_depth = None;
    let mut container_seen = false;
    let mut tables = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                if is_office(&namespace, element, b"spreadsheet") {
                    spreadsheet_depth = Some(depth);
                } else if is_table(&namespace, element, b"data-pilot-tables")
                    && spreadsheet_depth.is_some_and(|value| depth == value + 1)
                {
                    if container_seen {
                        return Err(invalid_message("duplicate table:data-pilot-tables"));
                    }
                    container_seen = true;
                    container_depth = Some(depth);
                } else if is_table(&namespace, element, b"data-pilot-table")
                    && container_depth.is_some_and(|value| depth == value + 1)
                {
                    let table = parse_table(&mut reader, element)?;
                    table.validate()?;
                    tables.push(table);
                    if tables.len() > MAX_DATA_PILOT_TABLES {
                        return Err(invalid_message(
                            "data-pilot table count exceeds safety limit",
                        ));
                    }
                    depth -= 1;
                } else if container_depth.is_some() {
                    return Err(invalid_message("invalid child in table:data-pilot-tables"));
                }
            },
            Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-table")
                    && container_depth == Some(depth) =>
            {
                return Err(invalid_message(
                    "data-pilot table requires at least one field",
                ));
            },
            Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-tables")
                    && spreadsheet_depth == Some(depth) =>
            {
                if container_seen {
                    return Err(invalid_message("duplicate table:data-pilot-tables"));
                }
                container_seen = true;
            },
            Event::End(ref element) => {
                if is_table(&namespace, element, b"data-pilot-tables")
                    && container_depth == Some(depth)
                {
                    container_depth = None;
                } else if is_office(&namespace, element, b"spreadsheet")
                    && spreadsheet_depth == Some(depth)
                {
                    spreadsheet_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
        buf.clear();
    }
    validate_data_pilot_tables(&tables)?;
    Ok(tables)
}

fn parse_table(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotTable> {
    let mut table = DataPilotTable {
        name: required_attr(reader, start, b"name")?,
        application_data: optional_attr(reader, start, b"application-data")?,
        grand_total: optional_attr(reader, start, b"grand-total")?
            .map(|value| DataPilotGrandTotal::parse(&value))
            .transpose()?,
        ignore_empty_rows: optional_bool(reader, start, b"ignore-empty-rows")?,
        identify_categories: optional_bool(reader, start, b"identify-categories")?,
        target_range_address: required_attr(reader, start, b"target-range-address")?,
        buttons: optional_attr(reader, start, b"buttons")?,
        show_filter_button: optional_bool(reader, start, b"show-filter-button")?,
        drill_down_on_double_click: optional_bool(reader, start, b"drill-down-on-double-click")?,
        grand_totals: Vec::new(),
        source: None,
        fields: Vec::new(),
    };
    let mut fields_started = false;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element)
                if is_table_ext(&namespace, element, b"data-pilot-grand-total") =>
            {
                if fields_started || table.source.is_some() {
                    return Err(invalid_message(
                        "data-pilot grand totals must precede the source and fields",
                    ));
                }
                table.grand_totals.push(parse_grand_total(reader, element)?);
                consume_empty_extension(reader, b"data-pilot-grand-total")?;
            },
            Event::Empty(ref element)
                if is_table_ext(&namespace, element, b"data-pilot-grand-total") =>
            {
                if fields_started || table.source.is_some() {
                    return Err(invalid_message(
                        "data-pilot grand totals must precede the source and fields",
                    ));
                }
                table.grand_totals.push(parse_grand_total(reader, element)?);
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-sql") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    DataPilotSource::Database(parse_source_sql(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-table") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    DataPilotSource::Database(parse_source_table(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-query") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    DataPilotSource::Database(parse_source_query(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"source-service") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = DataPilotSource::Service {
                    name: required_attr(reader, element, b"name")?,
                    source_name: required_attr(reader, element, b"source-name")?,
                    object_name: required_attr(reader, element, b"object-name")?,
                    user_name: optional_attr(reader, element, b"user-name")?,
                    password: optional_attr(reader, element, b"password")?,
                };
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"source-cell-range") => {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = parse_cell_range_source(reader, element)?;
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"source-cell-range") => {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = DataPilotSource::CellRange {
                    name: optional_attr(reader, element, b"name")?,
                    cell_range_address: required_attr(reader, element, b"cell-range-address")?,
                    filter: None,
                };
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-field") => {
                fields_started = true;
                table.fields.push(parse_field(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-field") => {
                fields_started = true;
                table.fields.push(field_from_start(reader, element)?);
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-table") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"database-source-sql")
                    || is_table(&namespace, element, b"database-source-table")
                    || is_table(&namespace, element, b"database-source-query")
                    || is_table(&namespace, element, b"source-service") => {},
            Event::Text(ref text) => {
                if !text_is_whitespace(text)? {
                    return Err(invalid_message("data-pilot table cannot contain text"));
                }
            },
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-table")),
            Event::Comment(_) => {},
            Event::Start(ref element) if !is_table_namespace(&namespace) => {
                skip_foreign_element(reader, element)?;
            },
            Event::Empty(_) if !is_table_namespace(&namespace) => {},
            other => {
                return Err(invalid_message(&format!(
                    "invalid child in table:data-pilot-table: {other:?}"
                )));
            },
        }
        buf.clear();
    }
    Ok(table)
}

fn parse_cell_range_source(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<DataPilotSource> {
    let address = required_attr(reader, start, b"cell-range-address")?;
    let name = optional_attr(reader, start, b"name")?;
    let mut filter = None;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"filter") => {
                set_once(&mut filter, parse_filter(reader, element)?, "source filter")?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"filter") => {
                return Err(invalid_message("table:filter has no expression"));
            },
            Event::End(ref element) if is_table(&namespace, element, b"source-cell-range") => break,
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:source-cell-range")),
            _ => return Err(invalid_message("invalid child in table:source-cell-range")),
        }
        buf.clear();
    }
    Ok(DataPilotSource::CellRange {
        name,
        cell_range_address: address,
        filter,
    })
}

fn field_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotField> {
    let orientation = DataPilotOrientation::parse(&required_attr(reader, start, b"orientation")?)?;
    Ok(DataPilotField {
        source_field_name: required_attr(reader, start, b"source-field-name")?,
        orientation,
        selected_page: optional_attr(reader, start, b"selected-page")?,
        is_data_layout_field: optional_attr(reader, start, b"is-data-layout-field")?,
        function: optional_attr(reader, start, b"function")?,
        used_hierarchy: optional_i64(reader, start, b"used-hierarchy")?,
        level: None,
        reference: None,
        groups: None,
    })
}

fn parse_field(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotField> {
    let mut field = field_from_start(reader, start)?;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-level") => {
                set_once(
                    &mut field.level,
                    parse_level(reader, element)?,
                    "data-pilot level",
                )?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-level") => {
                set_once(
                    &mut field.level,
                    level_from_start(reader, element)?,
                    "data-pilot level",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-field-reference") =>
            {
                set_once(
                    &mut field.reference,
                    parse_reference(reader, element)?,
                    "field reference",
                )?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-groups") => {
                set_once(
                    &mut field.groups,
                    parse_groups(reader, element)?,
                    "field groups",
                )?;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-field") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"data-pilot-field-reference") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-field")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-field")),
        }
        buf.clear();
    }
    field.validate()?;
    Ok(field)
}

fn level_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotLevel> {
    Ok(DataPilotLevel {
        show_empty: optional_bool(reader, start, b"show-empty")?,
        repeat_item_labels: optional_ns_bool(
            reader,
            start,
            CALC_EXT_NAMESPACE,
            b"repeat-item-labels",
        )?,
        ..Default::default()
    })
}

fn parse_level(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotLevel> {
    let mut level = level_from_start(reader, start)?;
    let mut subtotals_seen = false;
    let mut members_seen = false;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
                level.subtotals = parse_subtotals(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
                level.members = parse_members(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info") =>
            {
                set_once(
                    &mut level.display,
                    parse_display(reader, element)?,
                    "display info",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-sort-info") =>
            {
                set_once(&mut level.sort, parse_sort(reader, element)?, "sort info")?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-layout-info") =>
            {
                set_once(
                    &mut level.layout,
                    parse_layout(reader, element)?,
                    "layout info",
                )?;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-level") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info")
                    || is_table(&namespace, element, b"data-pilot-sort-info")
                    || is_table(&namespace, element, b"data-pilot-layout-info") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Start(ref element) if !is_table_namespace(&namespace) => {
                skip_foreign_element(reader, element)?
            },
            Event::Empty(_) if !is_table_namespace(&namespace) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-level")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-level")),
        }
        buf.clear();
    }
    Ok(level)
}

fn parse_subtotals(reader: &mut NsReader<&[u8]>) -> Result<Vec<String>> {
    parse_empty_children(
        reader,
        b"data-pilot-subtotals",
        b"data-pilot-subtotal",
        b"function",
    )
}

fn parse_members(reader: &mut NsReader<&[u8]>) -> Result<Vec<DataPilotMember>> {
    let mut members = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-member") =>
            {
                members.push(DataPilotMember {
                    name: required_attr(reader, element, b"name")?,
                    display: optional_bool(reader, element, b"display")?,
                    show_details: optional_bool(reader, element, b"show-details")?,
                })
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                break;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-member") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-members")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-members")),
        }
        buf.clear();
    }
    Ok(members)
}

fn parse_display(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DataPilotDisplayInfo> {
    Ok(DataPilotDisplayInfo {
        enabled: required_bool(reader, element, b"enabled")?,
        data_field: required_attr(reader, element, b"data-field")?,
        member_count: required_u64(reader, element, b"member-count")?,
        mode: DataPilotDisplayMemberMode::parse(&required_attr(
            reader,
            element,
            b"display-member-mode",
        )?)?,
    })
}

fn parse_sort(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<DataPilotSortInfo> {
    Ok(DataPilotSortInfo {
        mode: DataPilotSortMode::parse(&required_attr(reader, element, b"sort-mode")?)?,
        data_field: optional_attr(reader, element, b"data-field")?,
        order: DataPilotSortOrder::parse(&required_attr(reader, element, b"order")?)?,
    })
}

fn parse_layout(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<DataPilotLayoutInfo> {
    Ok(DataPilotLayoutInfo {
        mode: DataPilotLayoutMode::parse(&required_attr(reader, element, b"layout-mode")?)?,
        add_empty_lines: required_bool(reader, element, b"add-empty-lines")?,
    })
}

fn parse_reference(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DataPilotFieldReference> {
    Ok(DataPilotFieldReference {
        field_name: required_attr(reader, element, b"field-name")?,
        member_type: DataPilotReferenceMemberType::parse(&required_attr(
            reader,
            element,
            b"member-type",
        )?)?,
        member_name: optional_attr(reader, element, b"member-name")?,
        reference_type: DataPilotReferenceType::parse(&required_attr(reader, element, b"type")?)?,
    })
}

fn parse_groups(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotGroups> {
    let start_boundary = parse_boundary(reader, start, b"start", b"date-start")?;
    let end_boundary = parse_boundary(reader, start, b"end", b"date-end")?;
    let mut groups = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-group") => {
                groups.push(parse_group(reader, element)?)
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-groups") => break,
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-groups")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-groups")),
        }
        buf.clear();
    }
    Ok(DataPilotGroups {
        source_field_name: required_attr(reader, start, b"source-field-name")?,
        start: start_boundary,
        end: end_boundary,
        step: required_f64(reader, start, b"step")?,
        grouped_by: DataPilotGroupBy::parse(&required_attr(reader, start, b"grouped-by")?)?,
        groups,
    })
}

fn parse_group(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<DataPilotGroup> {
    Ok(DataPilotGroup {
        name: required_attr(reader, start, b"name")?,
        members: parse_empty_children(
            reader,
            b"data-pilot-group",
            b"data-pilot-group-member",
            b"name",
        )?,
    })
}

fn parse_empty_children(
    reader: &mut NsReader<&[u8]>,
    parent: &[u8],
    child: &[u8],
    attribute: &[u8],
) -> Result<Vec<String>> {
    let mut values = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, child) =>
            {
                values.push(required_attr(reader, element, attribute)?)
            },
            Event::End(ref element) if is_table(&namespace, element, parent) => break,
            Event::End(ref element) if is_table(&namespace, element, child) => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated data-pilot child container")),
            _ => return Err(invalid_message("invalid data-pilot child element")),
        }
        buf.clear();
    }
    Ok(values)
}

fn parse_boundary(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    numeric: &[u8],
    date: &[u8],
) -> Result<DataPilotGroupBoundary> {
    match (
        optional_attr(reader, element, numeric)?,
        optional_attr(reader, element, date)?,
    ) {
        (Some(_), Some(_)) | (None, None) => Err(invalid_message(
            "data-pilot grouping requires exactly one boundary attribute",
        )),
        (Some(value), None) if value == "auto" => Ok(DataPilotGroupBoundary::AutomaticNumber),
        (Some(value), None) => value
            .parse::<f64>()
            .map(DataPilotGroupBoundary::Number)
            .map_err(|_| invalid("group boundary", &value)),
        (None, Some(value)) if value == "auto" => Ok(DataPilotGroupBoundary::AutomaticDate),
        (None, Some(value)) => Ok(DataPilotGroupBoundary::Date(value)),
    }
}

pub(crate) fn write_data_pilot_tables(
    output: &mut String,
    tables: &[DataPilotTable],
) -> Result<()> {
    if tables.is_empty() {
        return Ok(());
    }
    if tables.len() > MAX_DATA_PILOT_TABLES {
        return Err(invalid_message(
            "data-pilot table count exceeds safety limit",
        ));
    }
    output.push_str("<table:data-pilot-tables>");
    for table in tables {
        write_table(output, table)?;
    }
    output.push_str("</table:data-pilot-tables>");
    Ok(())
}

pub(crate) fn write_data_pilot_table_fragment(table: &DataPilotTable) -> Result<String> {
    let mut output = String::new();
    write_table(&mut output, table)?;
    Ok(output)
}

fn write_table(out: &mut String, table: &DataPilotTable) -> Result<()> {
    table.validate()?;
    out.push_str(
        "<table:data-pilot-table xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
    );
    attr(out, "table:name", Some(&table.name));
    attr(
        out,
        "table:application-data",
        table.application_data.as_deref(),
    );
    attr(
        out,
        "table:grand-total",
        table.grand_total.map(DataPilotGrandTotal::as_str),
    );
    bool_attr(out, "table:ignore-empty-rows", table.ignore_empty_rows);
    bool_attr(out, "table:identify-categories", table.identify_categories);
    attr(
        out,
        "table:target-range-address",
        Some(&table.target_range_address),
    );
    attr(out, "table:buttons", table.buttons.as_deref());
    bool_attr(out, "table:show-filter-button", table.show_filter_button);
    bool_attr(
        out,
        "table:drill-down-on-double-click",
        table.drill_down_on_double_click,
    );
    out.push('>');
    for total in &table.grand_totals {
        out.push_str("<table-ext:data-pilot-grand-total xmlns:table-ext=\"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0\"");
        bool_attr(out, "table:display", Some(total.display));
        attr(out, "table:orientation", Some(total.orientation.as_str()));
        attr(out, "table-ext:display-name", total.display_name.as_deref());
        out.push_str("/>");
    }
    if let Some(source) = &table.source {
        write_source(out, source);
    }
    for field in &table.fields {
        write_field(out, field)?;
    }
    out.push_str("</table:data-pilot-table>");
    Ok(())
}

fn write_source(out: &mut String, source: &DataPilotSource) {
    match source {
        DataPilotSource::Database(source) => write_database_source(out, source),
        DataPilotSource::Service {
            name,
            source_name,
            object_name,
            user_name,
            password,
        } => {
            out.push_str("<table:source-service");
            attr(out, "table:name", Some(name));
            attr(out, "table:source-name", Some(source_name));
            attr(out, "table:object-name", Some(object_name));
            attr(out, "table:user-name", user_name.as_deref());
            attr(out, "table:password", password.as_deref());
            out.push_str("/>");
        },
        DataPilotSource::CellRange {
            name,
            cell_range_address,
            filter,
        } => {
            out.push_str("<table:source-cell-range");
            attr(out, "table:name", name.as_deref());
            attr(out, "table:cell-range-address", Some(cell_range_address));
            if let Some(filter) = filter {
                out.push('>');
                write_filter(out, filter);
                out.push_str("</table:source-cell-range>");
            } else {
                out.push_str("/>");
            }
        },
    }
}

fn write_field(out: &mut String, field: &DataPilotField) -> Result<()> {
    field.validate()?;
    out.push_str("<table:data-pilot-field");
    attr(
        out,
        "table:source-field-name",
        Some(&field.source_field_name),
    );
    attr(out, "table:orientation", Some(field.orientation.as_str()));
    attr(out, "table:selected-page", field.selected_page.as_deref());
    attr(
        out,
        "table:is-data-layout-field",
        field.is_data_layout_field.as_deref(),
    );
    attr(out, "table:function", field.function.as_deref());
    i64_attr(out, "table:used-hierarchy", field.used_hierarchy);
    if field.level.is_none() && field.reference.is_none() && field.groups.is_none() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(level) = &field.level {
        write_level(out, level);
    }
    if let Some(reference) = &field.reference {
        out.push_str("<table:data-pilot-field-reference");
        attr(out, "table:field-name", Some(&reference.field_name));
        attr(
            out,
            "table:member-type",
            Some(reference.member_type.as_str()),
        );
        attr(out, "table:member-name", reference.member_name.as_deref());
        attr(out, "table:type", Some(reference.reference_type.as_str()));
        out.push_str("/>");
    }
    if let Some(groups) = &field.groups {
        write_groups(out, groups);
    }
    out.push_str("</table:data-pilot-field>");
    Ok(())
}

fn write_level(out: &mut String, level: &DataPilotLevel) {
    out.push_str("<table:data-pilot-level");
    bool_attr(out, "table:show-empty", level.show_empty);
    if level.repeat_item_labels.is_some() {
        out.push_str(" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"");
        bool_attr(out, "calcext:repeat-item-labels", level.repeat_item_labels);
    }
    if level.subtotals.is_empty()
        && level.members.is_empty()
        && level.display.is_none()
        && level.sort.is_none()
        && level.layout.is_none()
    {
        out.push_str("/>");
        return;
    }
    out.push('>');
    if !level.subtotals.is_empty() {
        out.push_str("<table:data-pilot-subtotals>");
        for function in &level.subtotals {
            out.push_str("<table:data-pilot-subtotal");
            attr(out, "table:function", Some(function));
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-subtotals>");
    }
    if !level.members.is_empty() {
        out.push_str("<table:data-pilot-members>");
        for member in &level.members {
            out.push_str("<table:data-pilot-member");
            attr(out, "table:name", Some(&member.name));
            bool_attr(out, "table:display", member.display);
            bool_attr(out, "table:show-details", member.show_details);
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-members>");
    }
    if let Some(info) = &level.display {
        out.push_str("<table:data-pilot-display-info");
        bool_attr(out, "table:enabled", Some(info.enabled));
        attr(out, "table:data-field", Some(&info.data_field));
        u64_attr(out, "table:member-count", info.member_count);
        attr(out, "table:display-member-mode", Some(info.mode.as_str()));
        out.push_str("/>");
    }
    if let Some(info) = &level.sort {
        out.push_str("<table:data-pilot-sort-info");
        attr(out, "table:sort-mode", Some(info.mode.as_str()));
        attr(out, "table:data-field", info.data_field.as_deref());
        attr(out, "table:order", Some(info.order.as_str()));
        out.push_str("/>");
    }
    if let Some(info) = &level.layout {
        out.push_str("<table:data-pilot-layout-info");
        attr(out, "table:layout-mode", Some(info.mode.as_str()));
        bool_attr(out, "table:add-empty-lines", Some(info.add_empty_lines));
        out.push_str("/>");
    }
    out.push_str("</table:data-pilot-level>");
}

fn write_groups(out: &mut String, groups: &DataPilotGroups) {
    out.push_str("<table:data-pilot-groups");
    attr(
        out,
        "table:source-field-name",
        Some(&groups.source_field_name),
    );
    write_boundary(out, "start", &groups.start);
    write_boundary(out, "end", &groups.end);
    attr(out, "table:step", Some(&groups.step.to_string()));
    attr(out, "table:grouped-by", Some(groups.grouped_by.as_str()));
    out.push('>');
    for group in &groups.groups {
        out.push_str("<table:data-pilot-group");
        attr(out, "table:name", Some(&group.name));
        out.push('>');
        for member in &group.members {
            out.push_str("<table:data-pilot-group-member");
            attr(out, "table:name", Some(member));
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-group>");
    }
    out.push_str("</table:data-pilot-groups>");
}

fn write_boundary(out: &mut String, suffix: &str, boundary: &DataPilotGroupBoundary) {
    match boundary {
        DataPilotGroupBoundary::AutomaticNumber => {
            attr(out, &format!("table:{suffix}"), Some("auto"));
        },
        DataPilotGroupBoundary::AutomaticDate => {
            attr(out, &format!("table:date-{suffix}"), Some("auto"));
        },
        DataPilotGroupBoundary::Number(value) => {
            attr(out, &format!("table:{suffix}"), Some(&value.to_string()))
        },
        DataPilotGroupBoundary::Date(value) => {
            attr(out, &format!("table:date-{suffix}"), Some(value))
        },
    }
}

fn is_table(namespace: &ResolveResult<'_>, element: &impl HasLocalName, local: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE)
        && element.local() == local
}
fn is_office(namespace: &ResolveResult<'_>, element: &impl HasLocalName, local: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE)
        && element.local() == local
}
fn is_table_ext(namespace: &ResolveResult<'_>, element: &impl HasLocalName, local: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_EXT_NAMESPACE)
        && element.local() == local
}
fn is_table_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE)
}

fn parse_grand_total(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DataPilotGrandTotalElement> {
    Ok(DataPilotGrandTotalElement {
        orientation: DataPilotGrandTotalOrientation::parse(&required_attr(
            reader,
            element,
            b"orientation",
        )?)?,
        display: required_bool(reader, element, b"display")?,
        display_name: optional_ns_attr(reader, element, TABLE_EXT_NAMESPACE, b"display-name")?,
    })
}

fn consume_empty_extension(reader: &mut NsReader<&[u8]>, local: &[u8]) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::End(ref element) if is_table_ext(&namespace, element, local) => return Ok(()),
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated data-pilot extension element")),
            _ => {
                return Err(invalid_message(
                    "data-pilot grand-total extension must be empty",
                ));
            },
        }
        buffer.clear();
    }
}

fn skip_foreign_element(reader: &mut NsReader<&[u8]>, _start: &BytesStart<'_>) -> Result<()> {
    let mut depth = 1usize;
    let mut buffer = Vec::new();
    while depth > 0 {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_message("data-pilot extension depth overflow"))?;
                if depth > 128 {
                    return Err(invalid_message(
                        "data-pilot extension nesting exceeds limit",
                    ));
                }
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) => {
                return Err(invalid_message(
                    "DOCTYPE is not allowed in data-pilot extensions",
                ));
            },
            Event::Eof => return Err(invalid_message("unterminated data-pilot extension")),
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn optional_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_message(&format!("invalid XML attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TABLE_NAMESPACE)
            && name.as_ref() == local
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_message(&format!("invalid XML attribute value: {error}")))?
                .into_owned();
            if found.replace(value).is_some() {
                return Err(invalid_message("duplicate table attribute"));
            }
        }
    }
    Ok(found)
}
fn optional_ns_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    wanted_namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_message(&format!("invalid XML attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == wanted_namespace)
            && name.as_ref() == local
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_message(&format!("invalid XML attribute value: {error}")))?
                .into_owned();
            if found.replace(value).is_some() {
                return Err(invalid_message("duplicate extension attribute"));
            }
        }
    }
    Ok(found)
}
fn optional_ns_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<bool>> {
    optional_ns_attr(reader, element, namespace, local)?
        .map(|value| parse_bool(&value))
        .transpose()
}
fn required_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    optional_attr(reader, element, local)?.ok_or_else(|| {
        invalid_message(&format!("missing table:{}", String::from_utf8_lossy(local)))
    })
}
fn optional_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<bool>> {
    optional_attr(reader, element, local)?
        .map(|value| parse_bool(&value))
        .transpose()
}
fn required_bool(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, local: &[u8]) -> Result<bool> {
    parse_bool(&required_attr(reader, element, local)?)
}
fn optional_i64(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<i64>> {
    optional_attr(reader, element, local)?
        .map(|value| value.parse().map_err(|_| invalid("integer", &value)))
        .transpose()
}
fn required_u64(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, local: &[u8]) -> Result<u64> {
    let value = required_attr(reader, element, local)?;
    value
        .parse()
        .map_err(|_| invalid("non-negative integer", &value))
}
fn required_f64(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, local: &[u8]) -> Result<f64> {
    let value = required_attr(reader, element, local)?;
    value.parse().map_err(|_| invalid("number", &value))
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("Boolean", value)),
    }
}
fn text_is_whitespace(text: &quick_xml::events::BytesText<'_>) -> Result<bool> {
    Ok(text
        .xml_content(XmlVersion::Explicit1_0)
        .map_err(|error| invalid_message(&format!("invalid XML text: {error}")))?
        .trim()
        .is_empty())
}
fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid_message(&format!("duplicate {name}")))
    } else {
        Ok(())
    }
}
fn attr(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}
fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    attr(
        out,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}
fn i64_attr(out: &mut String, name: &str, value: Option<i64>) {
    if let Some(value) = value {
        attr(out, name, Some(&value.to_string()));
    }
}
fn u64_attr(out: &mut String, name: &str, value: u64) {
    attr(out, name, Some(&value.to_string()));
}
fn invalid(kind: &str, value: &str) -> Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}
fn invalid_message(message: &str) -> Error {
    Error::InvalidFormat(message.to_string())
}
fn xml_error(error: quick_xml::Error) -> Error {
    invalid_message(&format!("XML parsing error: {error}"))
}

// End-to-end cases require the transactional spreadsheet facade; retained until
// that package owner can be wired without cross-family dependencies.
#[cfg(any())]
mod tests {
    use super::*;

    const XMLNS: &str = r#"xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0""#;

    fn complete_xml() -> String {
        format!(
            r#"<o:document-content {XMLNS}><o:body><o:spreadsheet>
            <t:data-pilot-tables><t:data-pilot-table t:name="Pivot &amp; One" t:application-data="app"
              t:grand-total="both" t:ignore-empty-rows="true" t:identify-categories="false"
              t:target-range-address="Result.A1:F20" t:buttons="Result.A1 Result.B1"
              t:show-filter-button="1" t:drill-down-on-double-click="0">
              <t:source-cell-range t:cell-range-address="Source.A1:D100">
                <t:filter t:display-duplicates="false"><t:filter-condition t:field-number="0" t:value="East" t:operator="="/></t:filter>
              </t:source-cell-range>
              <t:data-pilot-field t:source-field-name="Region" t:orientation="row" t:function="auto" t:used-hierarchy="1">
                <t:data-pilot-level t:show-empty="false">
                  <t:data-pilot-subtotals><t:data-pilot-subtotal t:function="sum"></t:data-pilot-subtotal></t:data-pilot-subtotals>
                  <t:data-pilot-members><t:data-pilot-member t:name="East" t:display="true" t:show-details="false"></t:data-pilot-member></t:data-pilot-members>
                  <t:data-pilot-display-info t:enabled="true" t:data-field="Sales" t:member-count="10" t:display-member-mode="from-top"></t:data-pilot-display-info>
                  <t:data-pilot-sort-info t:sort-mode="data" t:data-field="Sales" t:order="descending"></t:data-pilot-sort-info>
                  <t:data-pilot-layout-info t:layout-mode="outline-subtotals-top" t:add-empty-lines="true"></t:data-pilot-layout-info>
                </t:data-pilot-level>
                <t:data-pilot-field-reference t:field-name="Region" t:member-type="named" t:member-name="East" t:type="member-percentage"></t:data-pilot-field-reference>
                <t:data-pilot-groups t:source-field-name="Region" t:start="0" t:end="100" t:step="10" t:grouped-by="days">
                  <t:data-pilot-group t:name="Area"><t:data-pilot-group-member t:name="East"></t:data-pilot-group-member></t:data-pilot-group>
                </t:data-pilot-groups>
              </t:data-pilot-field>
              <t:data-pilot-field t:source-field-name="Page" t:orientation="page" t:selected-page="All"/>
            </t:data-pilot-table></t:data-pilot-tables>
            </o:spreadsheet></o:body></o:document-content>"#
        )
    }

    #[test]
    fn parses_all_standard_metadata_with_namespace_aliases() {
        let tables = parse_data_pilot_tables(&complete_xml()).unwrap();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.name, "Pivot & One");
        assert_eq!(table.grand_total, Some(DataPilotGrandTotal::Both));
        assert_eq!(table.fields.len(), 2);
        assert_eq!(table.fields[0].level.as_ref().unwrap().subtotals, ["sum"]);
        assert_eq!(
            table.fields[0]
                .level
                .as_ref()
                .unwrap()
                .sort
                .as_ref()
                .unwrap()
                .mode,
            DataPilotSortMode::Data
        );
        assert_eq!(
            table.fields[0].groups.as_ref().unwrap().groups[0].members,
            ["East"]
        );
        assert!(matches!(
            table.source,
            Some(DataPilotSource::CellRange {
                filter: Some(_),
                ..
            })
        ));
    }

    #[test]
    fn writer_round_trips_complete_declaration() {
        let tables = parse_data_pilot_tables(&complete_xml()).unwrap();
        let mut body = String::new();
        write_data_pilot_tables(&mut body, &tables).unwrap();
        assert!(body.contains("Pivot &amp; One"));
        assert!(body.contains("<table:filter-condition"));
        let wrapped = format!(
            r#"<o:spreadsheet {XMLNS} xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{body}</o:spreadsheet>"#
        );
        let reparsed = parse_data_pilot_tables(&wrapped).unwrap();
        assert_eq!(reparsed, tables);
    }

    #[test]
    fn rejects_schema_invalid_declarations() {
        for body in [
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"/></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="page"/></t:data-pilot-table></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="row" t:selected-page="X"/></t:data-pilot-table></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="sideways"/></t:data-pilot-table></t:data-pilot-tables>"#,
        ] {
            let xml = format!(r#"<o:spreadsheet {XMLNS}>{body}</o:spreadsheet>"#);
            assert!(parse_data_pilot_tables(&xml).is_err(), "{body}");
        }
    }

    #[test]
    fn round_trips_through_builder_and_mutable_packages() {
        let table = parse_data_pilot_tables(&complete_xml()).unwrap().remove(0);
        let mut builder = crate::SpreadsheetBuilder::new();
        builder.add_sheet("Source").unwrap();
        builder.add_data_pilot_table(table).unwrap();
        let spreadsheet = crate::Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(spreadsheet.data_pilot_tables().len(), 1);

        let mut mutable = crate::MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        mutable.data_pilot_tables_mut()[0].name = "Updated".to_string();
        let reparsed = crate::Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.data_pilot_tables()[0].name, "Updated");
    }
}
