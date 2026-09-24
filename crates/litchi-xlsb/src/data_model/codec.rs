//! Bounded codec for the XLSB workbook Data Model record family.

use std::collections::HashSet;
use std::ops::Range;

use crate::data_model::model::{
    Definition, Relationship, Table, TimeGrouping, TimeGroupingColumn, TimeGroupingContentType,
};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Limits as RawLimits, Records, Writer, kind};

const FRT_HEADER_LEN: usize = 4;
const MIN_MODEL_VERSION: u8 = 5;
const MAX_STRING_BYTES: usize = 1_048_576;
const MAX_STRING_UNITS: usize = 1_048_576;
const MAX_TABLES: usize = 65_536;
const MAX_RELATIONSHIPS: usize = 65_536;
const MAX_TIME_GROUPINGS: usize = 65_536;
const MAX_TIME_GROUPING_COLUMNS: usize = 256;
const MAX_REWRITE_BYTES: usize = 64 * 1024 * 1024;
const MAX_RECORDS: usize = 1_048_576;
/// Maximum bytes retained for one Workbook/model source stream or lexical
/// source token by this owner.
pub const HARD_MAX_PART_BYTES: usize = 512 * 1024 * 1024;
/// Maximum physical parts visited by one Data Model package admission pass.
pub const HARD_MAX_GRAPH_PARTS: usize = 1_000_000;
/// Maximum relationships visited by one Data Model package admission pass.
pub const HARD_MAX_GRAPH_RELATIONSHIPS: usize = 1_000_000;
/// Maximum aggregate metadata bytes charged before an owned package clone.
pub const HARD_MAX_METADATA_BYTES: usize = 256 * 1024 * 1024;
/// Maximum bytes admitted from the workbook External Data Connections owner.
pub const HARD_MAX_CONNECTION_BYTES: usize = 32 * 1024 * 1024;
/// Maximum typed External Data Connections retained by one snapshot.
pub const HARD_MAX_CONNECTIONS: usize = 4_096;

/// Finite limits for Data Model record inspection and rewrite.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadLimits {
    /// Raw BIFF12 record and string limits.
    pub raw: RawLimits,
    /// Maximum number of workbook table records.
    pub max_tables: usize,
    /// Maximum number of workbook relationship records.
    pub max_relationships: usize,
    /// Maximum number of time grouping records.
    pub max_time_groupings: usize,
    /// Maximum number of generated columns in one time grouping.
    pub max_time_grouping_columns: usize,
    /// Maximum rewritten workbook-stream bytes attributable to this family.
    pub max_rewrite_bytes: usize,
    /// Maximum number of BIFF12 records indexed while locating the model block.
    /// This is a workbook-record diagnostic/admission bound only; package graph
    /// and source ownership bounds are separate fields below.
    pub max_records: usize,
    /// Maximum bytes in one retained Workbook/model stream or lexical source
    /// token captured for source-bound publication.
    pub max_part_bytes: usize,
    /// Maximum physical package parts visited during package admission.
    pub max_graph_parts: usize,
    /// Maximum relationships visited during package admission.
    pub max_graph_relationships: usize,
    /// Maximum conservative charge for modeled part and relationship metadata
    /// before cloning a package or retaining a source-bound snapshot.
    ///
    /// This is not a total allocation budget. OPC source-provenance maps and
    /// other retained state are bounded separately by their OPC limits.
    pub max_metadata_bytes: usize,
    /// Maximum External Data Connections part bytes admitted for closure
    /// validation and source capture.
    pub max_connection_bytes: usize,
    /// Maximum typed External Data Connections retained for closure checks.
    pub max_connections: usize,
}

impl ReadLimits {
    /// Conservative finite limits for ordinary XLSB processing.
    pub const DEFAULT: Self = Self {
        raw: RawLimits::DEFAULT,
        max_tables: MAX_TABLES,
        max_relationships: MAX_RELATIONSHIPS,
        max_time_groupings: MAX_TIME_GROUPINGS,
        max_time_grouping_columns: MAX_TIME_GROUPING_COLUMNS,
        max_rewrite_bytes: MAX_REWRITE_BYTES,
        max_records: MAX_RECORDS,
        max_part_bytes: HARD_MAX_PART_BYTES,
        max_graph_parts: 100_000,
        max_graph_relationships: 1_000_000,
        max_metadata_bytes: 64 * 1024 * 1024,
        max_connection_bytes: 32 * 1024 * 1024,
        max_connections: 4_096,
    };

    pub(crate) fn validate(self) -> Result<Self> {
        self.raw.validate()?;
        if self.max_tables > MAX_TABLES
            || self.max_relationships > MAX_RELATIONSHIPS
            || self.max_time_groupings > MAX_TIME_GROUPINGS
            || self.max_time_grouping_columns > MAX_TIME_GROUPING_COLUMNS
            || self.max_rewrite_bytes > MAX_REWRITE_BYTES
            || self.max_records > MAX_RECORDS
            || self.max_part_bytes > HARD_MAX_PART_BYTES
            || self.max_graph_parts > HARD_MAX_GRAPH_PARTS
            || self.max_graph_relationships > HARD_MAX_GRAPH_RELATIONSHIPS
            || self.max_metadata_bytes > HARD_MAX_METADATA_BYTES
            || self.max_connection_bytes > HARD_MAX_CONNECTION_BYTES
            || self.max_connections > HARD_MAX_CONNECTIONS
        {
            return Err(invalid(
                "Data Model limit exceeds the implementation ceiling",
            ));
        }
        Ok(self)
    }
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[cfg(test)]
#[derive(Clone)]
struct WireRecord {
    kind: u16,
    range: Range<usize>,
}

#[derive(Clone, Debug)]
struct Layout {
    begin_header: [u8; FRT_HEADER_LEN],
    table_headers: Vec<[u8; FRT_HEADER_LEN]>,
    relationship_headers: Vec<[u8; FRT_HEADER_LEN]>,
    time_grouping_headers: Vec<[u8; FRT_HEADER_LEN]>,
    column_headers: Vec<Vec<[u8; FRT_HEADER_LEN]>>,
    has_tables: bool,
    has_relationships: bool,
    has_time_groupings: bool,
    opaque: bool,
}

impl Default for Layout {
    fn default() -> Self {
        Self {
            begin_header: [0; FRT_HEADER_LEN],
            table_headers: Vec::new(),
            relationship_headers: Vec::new(),
            time_grouping_headers: Vec::new(),
            column_headers: Vec::new(),
            has_tables: false,
            has_relationships: false,
            has_time_groupings: false,
            opaque: false,
        }
    }
}

#[derive(Clone, Debug)]
pub(crate) struct Parsed {
    pub(crate) definition: Option<Definition>,
    pub(crate) block_range: Option<Range<usize>>,
    pub(crate) end_book_offset: Option<usize>,
    layout: Layout,
}

/// Parse the workbook's optional `BrtBeginDataModel` block.
pub(crate) fn parse_workbook(data: &[u8], limits: ReadLimits) -> Result<Parsed> {
    let limits = limits.validate()?;
    let mut iterator = Records::try_with_limits(data, limits.raw)?;
    let mut record_count = 0_usize;
    let mut begin: Option<usize> = None;
    let mut end: Option<usize> = None;
    let mut end_book_offset = None;
    let mut book_ended = false;
    let mut definition = None;
    let mut layout = Layout::default();
    let mut tables_open = false;
    let mut relationships_open = false;
    let mut time_groupings_open = false;
    let mut seen_tables = false;
    let mut seen_relationships = false;
    let mut seen_time_groupings = false;
    let mut section = 0_u8;
    let mut current_grouping: Option<TimeGrouping> = None;
    let mut tables = Vec::new();
    let mut relationships = Vec::new();
    let mut time_groupings = Vec::new();

    while let Some(record) = iterator.next() {
        let record = record?;
        if record_count >= limits.max_records {
            return Err(limit("workbook record count"));
        }
        record_count += 1;
        let record_start = record.offset();
        let record_end = iterator.offset();
        let record_kind = record.kind().get();
        let payload = record.payload();
        if record_kind == kind::END_BOOK.get() {
            require_empty(payload, "BrtEndBook")?;
            if end_book_offset.is_some() {
                return Err(invalid("duplicate BrtEndBook record"));
            }
            end_book_offset = Some(record_start);
            book_ended = true;
        } else if book_ended && is_model_record(record_kind) {
            return Err(invalid(
                "Data Model record appears after BrtEndBook in the workbook stream",
            ));
        }
        if begin.is_none() {
            if record_kind == kind::BEGIN_DATA_MODEL.get() {
                if definition.is_some() {
                    return Err(invalid("duplicate BrtBeginDataModel record"));
                }
                let (header, min_version) = parse_begin_data_model(payload)?;
                layout.begin_header = header;
                definition = Some(Definition {
                    min_version_load: min_version,
                    tables: Vec::new(),
                    relationships: Vec::new(),
                    time_groupings: Vec::new(),
                });
                begin = Some(record_start);
            } else if is_model_record(record_kind) {
                return Err(invalid(
                    "Data Model collection record appears outside BrtBeginDataModel",
                ));
            }
            continue;
        }

        if end.is_some() {
            if record_kind == kind::BEGIN_DATA_MODEL.get() {
                return Err(invalid("multiple BrtBeginDataModel blocks"));
            }
            if is_model_record(record_kind) {
                return Err(invalid(
                    "Data Model collection record appears after BrtEndDataModel",
                ));
            }
            continue;
        }

        if record_kind == kind::END_BOOK.get() {
            return Err(invalid(
                "BrtEndBook appears before the Data Model block is closed",
            ));
        }

        match record_kind {
            value if value == kind::BEGIN_DATA_MODEL.get() => {
                return Err(invalid("nested BrtBeginDataModel record"));
            },
            value if value == kind::BEGIN_MODEL_TABLES.get() => {
                require_empty(payload, "BrtBeginModelTables")?;
                if tables_open
                    || relationships_open
                    || time_groupings_open
                    || seen_tables
                    || section != 0
                {
                    return Err(invalid("nested or overlapping Data Model table collection"));
                }
                tables_open = true;
                seen_tables = true;
                section = 1;
                layout.has_tables = true;
            },
            value if value == kind::END_MODEL_TABLES.get() => {
                require_empty(payload, "BrtEndModelTables")?;
                if !tables_open {
                    return Err(invalid("BrtEndModelTables has no matching begin"));
                }
                tables_open = false;
            },
            value if value == kind::MODEL_TABLE.get() => {
                if !tables_open {
                    return Err(invalid("BrtModelTable is outside BrtBeginModelTables"));
                }
                if tables.len() >= limits.max_tables {
                    return Err(limit("table count"));
                }
                reserve_one(&mut tables, "Data Model tables")?;
                reserve_one(&mut layout.table_headers, "Data Model table headers")?;
                let (header, table) = parse_table(payload, limits.raw)?;
                layout.table_headers.push(header);
                tables.push(table);
            },
            value if value == kind::BEGIN_MODEL_RELATIONSHIPS.get() => {
                require_empty(payload, "BrtBeginModelRelationships")?;
                if relationships_open || tables_open || time_groupings_open {
                    return Err(invalid(
                        "nested or overlapping Data Model relationship collection",
                    ));
                }
                if seen_relationships || section > 1 {
                    return Err(invalid(
                        "Data Model relationship collection is duplicated or out of order",
                    ));
                }
                relationships_open = true;
                seen_relationships = true;
                section = 2;
                layout.has_relationships = true;
            },
            value if value == kind::END_MODEL_RELATIONSHIPS.get() => {
                require_empty(payload, "BrtEndModelRelationships")?;
                if !relationships_open {
                    return Err(invalid("BrtEndModelRelationships has no matching begin"));
                }
                relationships_open = false;
            },
            value if value == kind::MODEL_RELATIONSHIP.get() => {
                if !relationships_open {
                    return Err(invalid(
                        "brtModelRelationship is outside brtBeginModelRelationships",
                    ));
                }
                if relationships.len() >= limits.max_relationships {
                    return Err(limit("relationship count"));
                }
                reserve_one(&mut relationships, "Data Model relationships")?;
                reserve_one(
                    &mut layout.relationship_headers,
                    "Data Model relationship headers",
                )?;
                let (header, relationship) = parse_relationship(payload, limits.raw)?;
                layout.relationship_headers.push(header);
                relationships.push(relationship);
            },
            value if value == kind::BEGIN_MODEL_TIME_GROUPINGS.get() => {
                require_empty(payload, "brtBeginModelTimeGroupings")?;
                if time_groupings_open || tables_open || relationships_open {
                    return Err(invalid(
                        "nested or overlapping Data Model time-grouping collection",
                    ));
                }
                if seen_time_groupings || section > 2 {
                    return Err(invalid(
                        "Data Model time-grouping collection is duplicated or out of order",
                    ));
                }
                time_groupings_open = true;
                seen_time_groupings = true;
                section = 3;
                layout.has_time_groupings = true;
            },
            value if value == kind::END_MODEL_TIME_GROUPINGS.get() => {
                require_empty(payload, "brtEndModelTimeGroupings")?;
                if !time_groupings_open || current_grouping.is_some() {
                    return Err(invalid("brtEndModelTimeGroupings has no matching begin"));
                }
                time_groupings_open = false;
            },
            value if value == kind::BEGIN_MODEL_TIME_GROUPING.get() => {
                if !time_groupings_open || current_grouping.is_some() {
                    return Err(invalid("invalid nested brtBeginModelTimeGrouping"));
                }
                if time_groupings.len() >= limits.max_time_groupings {
                    return Err(limit("time grouping count"));
                }
                reserve_one(
                    &mut layout.time_grouping_headers,
                    "Data Model grouping headers",
                )?;
                reserve_one(&mut layout.column_headers, "Data Model grouping columns")?;
                let (header, grouping) = parse_time_grouping(payload, limits.raw)?;
                layout.time_grouping_headers.push(header);
                layout.column_headers.push(Vec::new());
                current_grouping = Some(grouping);
            },
            value if value == kind::END_MODEL_TIME_GROUPING.get() => {
                require_empty(payload, "brtEndModelTimeGrouping")?;
                let grouping = current_grouping
                    .take()
                    .ok_or_else(|| invalid("brtEndModelTimeGrouping has no matching begin"))?;
                time_groupings.push(grouping);
            },
            value if value == kind::MODEL_TIME_GROUPING_CALC_COL.get() => {
                let grouping = current_grouping
                    .as_mut()
                    .ok_or_else(|| invalid("brtModelTimeGroupingCalcCol is outside a grouping"))?;
                if grouping.columns.len() >= limits.max_time_grouping_columns {
                    return Err(limit("time grouping calculated-column count"));
                }
                reserve_one(
                    &mut layout.column_headers[time_groupings.len()],
                    "Data Model grouping column headers",
                )?;
                reserve_one(&mut grouping.columns, "Data Model grouping columns")?;
                let (header, column) = parse_time_grouping_column(payload, limits.raw)?;
                let grouping_index = time_groupings.len();
                layout.column_headers[grouping_index].push(header);
                grouping.columns.push(column);
            },
            value if value == kind::END_DATA_MODEL.get() => {
                require_empty(payload, "BrtEndDataModel")?;
                if tables_open
                    || relationships_open
                    || time_groupings_open
                    || current_grouping.is_some()
                {
                    return Err(invalid("BrtEndDataModel closes an unbalanced collection"));
                }
                end = Some(record_end);
            },
            _ => {
                // A model block can contain future records. Read them and keep
                // their exact source bytes, but refuse edits that would need to
                // regenerate the block without knowing their structure.
                layout.opaque = true;
            },
        }
    }

    let Some(begin_index) = begin else {
        return Ok(Parsed {
            definition: None,
            block_range: None,
            end_book_offset,
            layout,
        });
    };
    let Some(end_offset) = end else {
        return Err(invalid("BrtBeginDataModel has no matching BrtEndDataModel"));
    };
    if begin_index >= end_offset {
        return Err(invalid("invalid Data Model record range"));
    }
    let mut value = definition.ok_or_else(|| invalid("missing Data Model definition"))?;
    value.tables = tables;
    value.relationships = relationships;
    value.time_groupings = time_groupings;
    validate_definition(&value, limits)?;
    let range = begin_index..end_offset;
    Ok(Parsed {
        definition: Some(value),
        block_range: Some(range),
        end_book_offset,
        layout,
    })
}

#[cfg(test)]
fn collect_records(data: &[u8], limits: RawLimits, max_records: usize) -> Result<Vec<WireRecord>> {
    let mut iterator = Records::try_with_limits(data, limits)?;
    let mut records = Vec::new();
    while let Some(record) = iterator.next() {
        let record = record?;
        if records.len() >= max_records {
            return Err(limit("workbook record count"));
        }
        let start = record.offset();
        let end = iterator.offset();
        records.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "Data Model record index",
            source,
        })?;
        records.push(WireRecord {
            kind: record.kind().get(),
            range: start..end,
        });
    }
    Ok(records)
}

fn parse_begin_data_model(payload: &[u8]) -> Result<([u8; FRT_HEADER_LEN], u8)> {
    if payload.len() != FRT_HEADER_LEN + 1 {
        return Err(Error::InvalidLength {
            expected: FRT_HEADER_LEN + 1,
            found: payload.len(),
        });
    }
    let header = payload[..FRT_HEADER_LEN]
        .try_into()
        .map_err(|_error| invalid("invalid Data Model FRTHeader"))?;
    validate_frt_blank(header, "BrtBeginDataModel")?;
    let version = payload[FRT_HEADER_LEN];
    if version < MIN_MODEL_VERSION {
        return Err(invalid(format!(
            "bVerLoadModelMin {version} is below the required minimum {MIN_MODEL_VERSION}"
        )));
    }
    Ok((header, version))
}

fn require_empty(payload: &[u8], name: &'static str) -> Result<()> {
    if !payload.is_empty() {
        return Err(invalid(format!(
            "{name} payload must be empty, found {} bytes",
            payload.len()
        )));
    }
    Ok(())
}

fn parse_table(payload: &[u8], limits: RawLimits) -> Result<([u8; FRT_HEADER_LEN], Table)> {
    let (header, mut cursor) = start_cursor(payload, "BrtModelTable", limits)?;
    let table = Table {
        id: cursor.read_wide_string()?,
        name: cursor.read_wide_string()?,
        connection: cursor.read_wide_string()?,
    };
    cursor.finish()?;
    Ok((header, table))
}

fn parse_relationship(
    payload: &[u8],
    limits: RawLimits,
) -> Result<([u8; FRT_HEADER_LEN], Relationship)> {
    let (header, mut cursor) = start_cursor(payload, "brtModelRelationship", limits)?;
    let relationship = Relationship {
        from_table: cursor.read_wide_string()?,
        from_column: cursor.read_wide_string()?,
        to_table: cursor.read_wide_string()?,
        to_column: cursor.read_wide_string()?,
    };
    cursor.finish()?;
    Ok((header, relationship))
}

fn parse_time_grouping(
    payload: &[u8],
    limits: RawLimits,
) -> Result<([u8; FRT_HEADER_LEN], TimeGrouping)> {
    let (header, mut cursor) = start_cursor(payload, "brtBeginModelTimeGrouping", limits)?;
    let grouping = TimeGrouping {
        table_name: cursor.read_wide_string()?,
        column_name: cursor.read_wide_string()?,
        column_id: cursor.read_wide_string()?,
        columns: Vec::new(),
    };
    cursor.finish()?;
    Ok((header, grouping))
}

fn parse_time_grouping_column(
    payload: &[u8],
    limits: RawLimits,
) -> Result<([u8; FRT_HEADER_LEN], TimeGroupingColumn)> {
    let (header, mut cursor) = start_cursor(payload, "brtModelTimeGroupingCalcCol", limits)?;
    let flags = cursor.read_u8()?;
    if flags & 0xC0 != 0 {
        return Err(invalid(
            "brtModelTimeGroupingCalcCol reserved bits must be zero",
        ));
    }
    let content_type = TimeGroupingContentType::from_wire((flags >> 1) & 0x1F)
        .ok_or_else(|| invalid("brtModelTimeGroupingCalcCol has an unknown contentType"))?;
    let column = TimeGroupingColumn {
        is_selected: flags & 1 != 0,
        content_type,
        column_name: cursor.read_wide_string()?,
        column_id: cursor.read_wide_string()?,
    };
    cursor.finish()?;
    Ok((header, column))
}

fn start_cursor<'a>(
    payload: &'a [u8],
    context: &'static str,
    limits: RawLimits,
) -> Result<([u8; FRT_HEADER_LEN], Cursor<'a>)> {
    if payload.len() < FRT_HEADER_LEN {
        return Err(Error::InvalidLength {
            expected: FRT_HEADER_LEN,
            found: payload.len(),
        });
    }
    let header = payload[..FRT_HEADER_LEN]
        .try_into()
        .map_err(|_error| invalid("invalid Data Model FRTHeader"))?;
    validate_frt_blank(header, context)?;
    Ok((
        header,
        Cursor::try_with_limits(&payload[FRT_HEADER_LEN..], context, limits)?,
    ))
}

fn validate_frt_blank(header: [u8; FRT_HEADER_LEN], context: &'static str) -> Result<()> {
    if header != [0; FRT_HEADER_LEN] {
        return Err(invalid(format!(
            "{context} FRTHeader reserved bytes must be zero"
        )));
    }
    Ok(())
}

/// Validate typed workbook Data Model metadata before publication.
pub(crate) fn validate_definition(value: &Definition, limits: ReadLimits) -> Result<()> {
    if value.min_version_load < MIN_MODEL_VERSION {
        return Err(invalid("min_version_load must be at least 5"));
    }
    if value.tables.len() > limits.max_tables {
        return Err(limit("table count"));
    }
    if value.relationships.len() > limits.max_relationships {
        return Err(limit("relationship count"));
    }
    if value.time_groupings.len() > limits.max_time_groupings {
        return Err(limit("time grouping count"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    ids.try_reserve(value.tables.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model table IDs",
            source,
        })?;
    names
        .try_reserve(value.tables.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model table names",
            source,
        })?;
    for table in &value.tables {
        bounded(&table.id, "table id")?;
        bounded(&table.name, "table name")?;
        bounded(&table.connection, "connection name")?;
        if !ids.insert(table.id.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate Data Model table id {:?}",
                table.id
            )));
        }
        if !names.insert(table.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate Data Model table name {:?}",
                table.name
            )));
        }
    }
    let mut relationship_keys = HashSet::new();
    relationship_keys
        .try_reserve(value.relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model relationship keys",
            source,
        })?;
    for relationship in &value.relationships {
        bounded(&relationship.from_table, "from table")?;
        bounded(&relationship.from_column, "from column")?;
        bounded(&relationship.to_table, "to table")?;
        bounded(&relationship.to_column, "to column")?;
        if !names.contains(&relationship.from_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown from_table {:?}",
                relationship.from_table
            )));
        }
        if !names.contains(&relationship.to_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown to_table {:?}",
                relationship.to_table
            )));
        }
        let key = (
            relationship.from_table.to_lowercase(),
            relationship.from_column.to_lowercase(),
            relationship.to_table.to_lowercase(),
            relationship.to_column.to_lowercase(),
        );
        if !relationship_keys.insert(key) {
            return Err(invalid("duplicate Data Model relationship"));
        }
    }
    let mut grouping_keys = HashSet::new();
    grouping_keys
        .try_reserve(value.time_groupings.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model time grouping keys",
            source,
        })?;
    for grouping in &value.time_groupings {
        bounded(&grouping.table_name, "time grouping table")?;
        bounded(&grouping.column_name, "time grouping column")?;
        bounded(&grouping.column_id, "time grouping column id")?;
        if !names.contains(&grouping.table_name.to_lowercase()) {
            return Err(invalid(format!(
                "time grouping references unknown table {:?}",
                grouping.table_name
            )));
        }
        if grouping.columns.len() > limits.max_time_grouping_columns {
            return Err(limit("time grouping calculated-column count"));
        }
        let grouping_key = (
            grouping.table_name.to_lowercase(),
            grouping.column_name.to_lowercase(),
            grouping.column_id.to_lowercase(),
        );
        if !grouping_keys.insert(grouping_key) {
            return Err(invalid("duplicate Data Model time grouping"));
        }
        let mut column_keys = HashSet::new();
        column_keys
            .try_reserve(grouping.columns.len())
            .map_err(|source| Error::Allocation {
                resource: "Data Model time grouping columns",
                source,
            })?;
        for column in &grouping.columns {
            bounded(&column.column_name, "calculated column name")?;
            bounded(&column.column_id, "calculated column id")?;
            if !column_keys.insert((
                column.content_type,
                column.column_name.to_lowercase(),
                column.column_id.to_lowercase(),
            )) {
                return Err(invalid("duplicate Data Model time grouping column"));
            }
        }
    }
    Ok(())
}

fn bounded(value: &str, label: &'static str) -> Result<()> {
    let bytes = value.len();
    let units = value.encode_utf16().count();
    if value.is_empty() {
        return Err(invalid(format!("{label} must not be empty")));
    }
    if bytes > MAX_STRING_BYTES || units > MAX_STRING_UNITS {
        return Err(limit(label));
    }
    Ok(())
}

fn serialize_definition(
    value: &Definition,
    layout: Option<&Layout>,
    limits: ReadLimits,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_definition(value, limits)?;
    if layout.is_some_and(|layout| layout.opaque) {
        return Err(Error::UnsupportedFeature(
            "Data Model workbook block contains opaque future records".to_string(),
        ));
    }
    let write_tables = layout.map_or(!value.tables.is_empty(), |value_layout| {
        value_layout.has_tables || !value.tables.is_empty()
    });
    let write_relationships = layout.map_or(!value.relationships.is_empty(), |value_layout| {
        value_layout.has_relationships || !value.relationships.is_empty()
    });
    let write_time_groupings = layout.map_or(!value.time_groupings.is_empty(), |value_layout| {
        value_layout.has_time_groupings || !value.time_groupings.is_empty()
    });
    let output_size = serialized_size(
        value,
        write_tables,
        write_relationships,
        write_time_groupings,
        limits,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_size)
        .map_err(|source| Error::Allocation {
            resource: "Data Model workbook serialization",
            source,
        })?;
    {
        let mut writer = Writer::try_with_limits(&mut output, limits.raw)?;
        let begin_header = layout.map_or([0; FRT_HEADER_LEN], |layout| layout.begin_header);
        let mut begin_payload = begin_header.to_vec();
        begin_payload.push(value.min_version_load);
        writer.write_record(kind::BEGIN_DATA_MODEL, &begin_payload)?;

        if write_tables {
            writer.write_record(kind::BEGIN_MODEL_TABLES, &[])?;
            for (index, table) in value.tables.iter().enumerate() {
                let header = layout
                    .and_then(|value_layout| value_layout.table_headers.get(index))
                    .copied()
                    .unwrap_or([0; FRT_HEADER_LEN]);
                let payload = write_payload(limits.raw, |writer| {
                    writer.write_all(&header)?;
                    writer.write_wide_string(&table.id)?;
                    writer.write_wide_string(&table.name)?;
                    writer.write_wide_string(&table.connection)
                })?;
                writer.write_record(kind::MODEL_TABLE, &payload)?;
            }
            writer.write_record(kind::END_MODEL_TABLES, &[])?;
        }

        if write_relationships {
            writer.write_record(kind::BEGIN_MODEL_RELATIONSHIPS, &[])?;
            for (index, relationship) in value.relationships.iter().enumerate() {
                let header = layout
                    .and_then(|value_layout| value_layout.relationship_headers.get(index))
                    .copied()
                    .unwrap_or([0; FRT_HEADER_LEN]);
                let payload = write_payload(limits.raw, |writer| {
                    writer.write_all(&header)?;
                    writer.write_wide_string(&relationship.from_table)?;
                    writer.write_wide_string(&relationship.from_column)?;
                    writer.write_wide_string(&relationship.to_table)?;
                    writer.write_wide_string(&relationship.to_column)
                })?;
                writer.write_record(kind::MODEL_RELATIONSHIP, &payload)?;
            }
            writer.write_record(kind::END_MODEL_RELATIONSHIPS, &[])?;
        }

        if write_time_groupings {
            writer.write_record(kind::BEGIN_MODEL_TIME_GROUPINGS, &[])?;
            for (grouping_index, grouping) in value.time_groupings.iter().enumerate() {
                let header = layout
                    .and_then(|value_layout| value_layout.time_grouping_headers.get(grouping_index))
                    .copied()
                    .unwrap_or([0; FRT_HEADER_LEN]);
                let payload = write_payload(limits.raw, |writer| {
                    writer.write_all(&header)?;
                    writer.write_wide_string(&grouping.table_name)?;
                    writer.write_wide_string(&grouping.column_name)?;
                    writer.write_wide_string(&grouping.column_id)
                })?;
                writer.write_record(kind::BEGIN_MODEL_TIME_GROUPING, &payload)?;
                for (column_index, column) in grouping.columns.iter().enumerate() {
                    let header = layout
                        .and_then(|value_layout| {
                            value_layout
                                .column_headers
                                .get(grouping_index)
                                .and_then(|headers| headers.get(column_index))
                        })
                        .copied()
                        .unwrap_or([0; FRT_HEADER_LEN]);
                    let flags = (column.content_type.wire() << 1) | u8::from(column.is_selected);
                    let payload = write_payload(limits.raw, |writer| {
                        writer.write_all(&header)?;
                        writer.write_u8(flags)?;
                        writer.write_wide_string(&column.column_name)?;
                        writer.write_wide_string(&column.column_id)
                    })?;
                    writer.write_record(kind::MODEL_TIME_GROUPING_CALC_COL, &payload)?;
                }
                writer.write_record(kind::END_MODEL_TIME_GROUPING, &[])?;
            }
            writer.write_record(kind::END_MODEL_TIME_GROUPINGS, &[])?;
        }
        writer.write_record(kind::END_DATA_MODEL, &[])?;
    }
    if output.len() > limits.max_rewrite_bytes {
        return Err(limit("rewritten workbook Data Model bytes"));
    }
    Ok(output)
}

fn serialized_size(
    value: &Definition,
    write_tables: bool,
    write_relationships: bool,
    write_time_groupings: bool,
    limits: ReadLimits,
) -> Result<usize> {
    let mut total = 0_usize;
    add_record_size(
        &mut total,
        kind::BEGIN_DATA_MODEL,
        FRT_HEADER_LEN + 1,
        limits,
    )?;
    if write_tables {
        add_record_size(&mut total, kind::BEGIN_MODEL_TABLES, 0, limits)?;
        for table in &value.tables {
            let payload = checked_sum([
                Ok(FRT_HEADER_LEN),
                wide_string_size(&table.id),
                wide_string_size(&table.name),
                wide_string_size(&table.connection),
            ])?;
            add_record_size(&mut total, kind::MODEL_TABLE, payload, limits)?;
        }
        add_record_size(&mut total, kind::END_MODEL_TABLES, 0, limits)?;
    }
    if write_relationships {
        add_record_size(&mut total, kind::BEGIN_MODEL_RELATIONSHIPS, 0, limits)?;
        for relationship in &value.relationships {
            let payload = checked_sum([
                Ok(FRT_HEADER_LEN),
                wide_string_size(&relationship.from_table),
                wide_string_size(&relationship.from_column),
                wide_string_size(&relationship.to_table),
                wide_string_size(&relationship.to_column),
            ])?;
            add_record_size(&mut total, kind::MODEL_RELATIONSHIP, payload, limits)?;
        }
        add_record_size(&mut total, kind::END_MODEL_RELATIONSHIPS, 0, limits)?;
    }
    if write_time_groupings {
        add_record_size(&mut total, kind::BEGIN_MODEL_TIME_GROUPINGS, 0, limits)?;
        for grouping in &value.time_groupings {
            let payload = checked_sum([
                Ok(FRT_HEADER_LEN),
                wide_string_size(&grouping.table_name),
                wide_string_size(&grouping.column_name),
                wide_string_size(&grouping.column_id),
            ])?;
            add_record_size(&mut total, kind::BEGIN_MODEL_TIME_GROUPING, payload, limits)?;
            for column in &grouping.columns {
                let payload = checked_sum([
                    Ok(FRT_HEADER_LEN + 1),
                    wide_string_size(&column.column_name),
                    wide_string_size(&column.column_id),
                ])?;
                add_record_size(
                    &mut total,
                    kind::MODEL_TIME_GROUPING_CALC_COL,
                    payload,
                    limits,
                )?;
            }
            add_record_size(&mut total, kind::END_MODEL_TIME_GROUPING, 0, limits)?;
        }
        add_record_size(&mut total, kind::END_MODEL_TIME_GROUPINGS, 0, limits)?;
    }
    add_record_size(&mut total, kind::END_DATA_MODEL, 0, limits)?;
    Ok(total)
}

fn add_record_size(
    total: &mut usize,
    kind: crate::raw::Kind,
    payload: usize,
    limits: ReadLimits,
) -> Result<()> {
    if payload > limits.raw.payload() {
        return Err(invalid(format!(
            "Data Model record payload {payload} exceeds raw limit {}",
            limits.raw.payload()
        )));
    }
    let record_size = kind_len(kind)
        .checked_add(varint_len(payload))
        .and_then(|size| size.checked_add(payload))
        .ok_or_else(|| limit("serialized workbook Data Model bytes"))?;
    *total = total
        .checked_add(record_size)
        .ok_or_else(|| limit("serialized workbook Data Model bytes"))?;
    if *total > limits.max_rewrite_bytes {
        return Err(limit("serialized workbook Data Model bytes"));
    }
    Ok(())
}

fn checked_sum<I>(sizes: I) -> Result<usize>
where
    I: IntoIterator<Item = Result<usize>>,
{
    sizes.into_iter().try_fold(0_usize, |total, size| {
        total
            .checked_add(size?)
            .ok_or_else(|| limit("serialized workbook Data Model bytes"))
    })
}

fn wide_string_size(value: &str) -> Result<usize> {
    let units = value.encode_utf16().count();
    if units > MAX_STRING_UNITS {
        return Err(limit("string units"));
    }
    4_usize
        .checked_add(
            units
                .checked_mul(2)
                .ok_or_else(|| limit("serialized workbook Data Model bytes"))?,
        )
        .ok_or_else(|| limit("serialized workbook Data Model bytes"))
}

fn kind_len(kind: crate::raw::Kind) -> usize {
    if kind.get() < 0x80 { 1 } else { 2 }
}

fn varint_len(mut value: usize) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

pub(crate) fn serialize_new_workbook(definition: &Definition) -> Result<Vec<u8>> {
    serialize_definition(definition, None, ReadLimits::DEFAULT)
}

fn write_payload(
    limits: RawLimits,
    write: impl FnOnce(&mut PayloadWriter<'_>) -> Result<()>,
) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    let mut writer = PayloadWriter {
        inner: Writer::try_with_limits(&mut payload, limits)?,
    };
    write(&mut writer)?;
    Ok(payload)
}

struct PayloadWriter<'a> {
    inner: Writer<&'a mut Vec<u8>>,
}

impl PayloadWriter<'_> {
    fn write_all(&mut self, bytes: &[u8]) -> Result<()> {
        self.inner.get_mut().extend_from_slice(bytes);
        Ok(())
    }

    fn write_wide_string(&mut self, value: &str) -> Result<()> {
        Ok(self.inner.write_wide_string(value)?)
    }

    fn write_u8(&mut self, value: u8) -> Result<()> {
        Ok(self.inner.write_u8(value)?)
    }
}

pub(crate) fn patch_workbook(
    source: &[u8],
    before: Option<&Definition>,
    after: Option<&Definition>,
    limits: ReadLimits,
) -> Result<Vec<u8>> {
    let parsed = parse_workbook(source, limits)?;
    if parsed.definition.as_ref() != before {
        return Err(Error::InvalidFormat(
            "Data Model workbook source does not match the transaction snapshot".to_string(),
        ));
    }
    if before == after {
        return Ok(source.to_vec());
    }
    if parsed.layout.opaque {
        return Err(Error::UnsupportedFeature(
            "cannot edit a Data Model workbook block containing opaque future records".to_string(),
        ));
    }
    let replacement = match after {
        Some(value) => serialize_definition(
            value,
            parsed.block_range.as_ref().map(|_| &parsed.layout),
            limits,
        )?,
        None => Vec::new(),
    };
    let mut output = Vec::new();
    if let Some(range) = parsed.block_range {
        let output_size = source
            .len()
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(replacement.len()))
            .ok_or_else(|| limit("rewritten workbook bytes"))?;
        if output_size > limits.max_rewrite_bytes {
            return Err(limit("rewritten workbook bytes"));
        }
        output
            .try_reserve(output_size)
            .map_err(|source| Error::Allocation {
                resource: "Data Model workbook rewrite",
                source,
            })?;
        output.extend_from_slice(&source[..range.start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&source[range.end..]);
    } else {
        let output_size = source
            .len()
            .checked_add(replacement.len())
            .ok_or_else(|| limit("rewritten workbook insertion"))?;
        if output_size > limits.max_rewrite_bytes {
            return Err(limit("rewritten workbook insertion"));
        }
        output
            .try_reserve(output_size)
            .map_err(|source| Error::Allocation {
                resource: "Data Model workbook insertion",
                source,
            })?;
        let insert_at = parsed
            .end_book_offset
            .ok_or_else(|| invalid("cannot insert Data Model without a BrtEndBook anchor"))?;
        output.extend_from_slice(&source[..insert_at]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&source[insert_at..]);
    }
    if output.len() > limits.max_rewrite_bytes {
        return Err(limit("rewritten workbook bytes"));
    }
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn is_model_record(value: u16) -> bool {
    value == kind::BEGIN_DATA_MODEL.get()
        || value == kind::END_DATA_MODEL.get()
        || value == kind::BEGIN_MODEL_TABLES.get()
        || value == kind::END_MODEL_TABLES.get()
        || value == kind::MODEL_TABLE.get()
        || value == kind::BEGIN_MODEL_RELATIONSHIPS.get()
        || value == kind::END_MODEL_RELATIONSHIPS.get()
        || value == kind::MODEL_RELATIONSHIP.get()
        || value == kind::BEGIN_MODEL_TIME_GROUPINGS.get()
        || value == kind::END_MODEL_TIME_GROUPINGS.get()
        || value == kind::BEGIN_MODEL_TIME_GROUPING.get()
        || value == kind::END_MODEL_TIME_GROUPING.get()
        || value == kind::MODEL_TIME_GROUPING_CALC_COL.get()
}

fn limit(name: &'static str) -> Error {
    Error::InvalidFormat(format!("Data Model {name} limit exceeded"))
}

#[cfg(test)]
#[allow(clippy::expect_used)]
mod tests {
    use super::*;
    use crate::raw::Kind;

    fn definition() -> Definition {
        Definition {
            min_version_load: 5,
            tables: vec![Table {
                id: "table-1".to_string(),
                name: "Sales".to_string(),
                connection: "Connection".to_string(),
            }],
            relationships: Vec::new(),
            time_groupings: vec![TimeGrouping {
                table_name: "Sales".to_string(),
                column_name: "Date".to_string(),
                column_id: "date-id".to_string(),
                columns: vec![TimeGroupingColumn {
                    is_selected: true,
                    content_type: TimeGroupingContentType::Years,
                    column_name: "Year".to_string(),
                    column_id: "year-id".to_string(),
                }],
            }],
        }
    }

    fn fixture() -> Vec<u8> {
        serialize_definition(&definition(), None, ReadLimits::DEFAULT).expect("fixture")
    }

    #[test]
    fn round_trip_model_records() {
        let bytes = fixture();
        let parsed = parse_workbook(&bytes, ReadLimits::DEFAULT).expect("parse");
        assert_eq!(parsed.definition, Some(definition()));
        assert!(!parsed.layout.opaque);
        assert_eq!(
            patch_workbook(
                &bytes,
                Some(&definition()),
                Some(&definition()),
                ReadLimits::DEFAULT
            )
            .expect("noop"),
            bytes
        );
    }

    #[test]
    fn reserved_time_grouping_bits_are_rejected() {
        let mut bytes = fixture();
        let records = collect_records(&bytes, RawLimits::DEFAULT, MAX_RECORDS).expect("records");
        let column = records
            .iter()
            .find(|record| record.kind == kind::MODEL_TIME_GROUPING_CALC_COL.get())
            .expect("column");
        let payload_offset = column.range.start + 2 + 1 + FRT_HEADER_LEN;
        bytes[payload_offset] |= 0x80;
        assert!(parse_workbook(&bytes, ReadLimits::DEFAULT).is_err());
    }

    #[test]
    fn frt_blank_reserved_headers_are_rejected() {
        let fixture = fixture();
        let records = collect_records(&fixture, RawLimits::DEFAULT, MAX_RECORDS).expect("records");

        let begin = records
            .iter()
            .find(|record| record.kind == kind::BEGIN_DATA_MODEL.get())
            .expect("begin");
        let mut bytes = fixture.clone();
        bytes[begin.range.start + 3] = 1;
        assert!(parse_workbook(&bytes, ReadLimits::DEFAULT).is_err());

        let table = records
            .iter()
            .find(|record| record.kind == kind::MODEL_TABLE.get())
            .expect("table");
        let mut bytes = fixture;
        bytes[table.range.start + 3] = 1;
        assert!(parse_workbook(&bytes, ReadLimits::DEFAULT).is_err());
    }

    #[test]
    fn future_record_blocks_edit_but_not_read() {
        let mut bytes = fixture();
        let records = collect_records(&bytes, RawLimits::DEFAULT, MAX_RECORDS).expect("records");
        let end = records
            .iter()
            .find(|record| record.kind == kind::END_DATA_MODEL.get())
            .expect("end");
        let mut inserted = Vec::new();
        let mut writer = Writer::new(&mut inserted);
        writer
            .write_record(Kind::new(4095).expect("kind"), &[1, 2])
            .expect("record");
        bytes.splice(end.range.start..end.range.start, inserted);
        let parsed = parse_workbook(&bytes, ReadLimits::DEFAULT).expect("read");
        assert!(parsed.layout.opaque);
        assert_eq!(
            patch_workbook(
                &bytes,
                parsed.definition.as_ref(),
                parsed.definition.as_ref(),
                ReadLimits::DEFAULT,
            )
            .expect("opaque no-op"),
            bytes
        );
        let mut changed = definition();
        changed.min_version_load = 6;
        assert!(
            patch_workbook(
                &bytes,
                parsed.definition.as_ref(),
                Some(&changed),
                ReadLimits::DEFAULT
            )
            .is_err()
        );
    }

    #[test]
    fn workbook_record_index_has_a_finite_limit() {
        let bytes = fixture();
        let mut limits = ReadLimits::DEFAULT;
        limits.max_records = 0;
        assert!(parse_workbook(&bytes, limits).is_err());
    }

    #[test]
    fn rewrite_byte_limit_is_checked_at_the_exact_boundary() {
        let bytes = fixture();
        let parsed = parse_workbook(&bytes, ReadLimits::DEFAULT).expect("parse");
        let mut changed = definition();
        changed.min_version_load = 6;

        let mut exact = ReadLimits::DEFAULT;
        exact.max_rewrite_bytes = bytes.len();
        let rewritten = patch_workbook(&bytes, parsed.definition.as_ref(), Some(&changed), exact)
            .expect("exact rewrite budget");
        assert_eq!(rewritten.len(), bytes.len());

        exact.max_rewrite_bytes = bytes.len().saturating_sub(1);
        assert!(
            patch_workbook(&bytes, parsed.definition.as_ref(), Some(&changed), exact,).is_err()
        );
    }

    #[test]
    fn model_records_after_end_book_are_rejected() {
        let fixture = fixture();
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer.write_record(kind::END_BOOK, &[]).expect("end book");
        bytes.extend_from_slice(&fixture);
        assert!(parse_workbook(&bytes, ReadLimits::DEFAULT).is_err());
    }

    #[test]
    fn end_book_inside_model_block_is_rejected() {
        let fixture = fixture();
        let records = collect_records(&fixture, RawLimits::DEFAULT, MAX_RECORDS).expect("records");
        let end = records
            .iter()
            .find(|record| record.kind == kind::END_DATA_MODEL.get())
            .expect("end");
        let mut end_book = Vec::new();
        Writer::new(&mut end_book)
            .write_record(kind::END_BOOK, &[])
            .expect("end book");
        let mut bytes = fixture;
        bytes.splice(end.range.start..end.range.start, end_book);
        assert!(parse_workbook(&bytes, ReadLimits::DEFAULT).is_err());
    }

    #[test]
    fn model_insertion_requires_an_end_book_anchor() {
        assert!(patch_workbook(&[], None, Some(&definition()), ReadLimits::DEFAULT,).is_err());
    }
}
