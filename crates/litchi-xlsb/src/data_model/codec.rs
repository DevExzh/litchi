//! Bounded codec for the XLSB workbook Data Model record family.

use std::collections::HashSet;
use std::ops::Range;

use crate::data_model::model::{
    Definition, Relationship, Table, TimeGrouping, TimeGroupingColumn, TimeGroupingContentType,
};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Encoder, Limits as RawLimits, Records, kind};

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
pub struct SemanticLimits {
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

impl SemanticLimits {
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

impl Default for SemanticLimits {
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

/// The Data Model collections, in the order a block must list them.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum ModelSection {
    /// No collection has been read yet.
    None,
    Tables,
    Relationships,
    TimeGroupings,
}

/// Where the scanner stands inside an open `BrtBeginDataModel` block.
///
/// ```text
/// Between(None)              --BrtBeginModelTables-->        Tables
/// Between(None | Tables)     --BrtBeginModelRelationships--> Relationships
/// Between(.. Relationships)  --brtBeginModelTimeGroupings--> TimeGroupings
/// Tables                     --BrtEndModelTables-->          Between(Tables)
/// Relationships              --BrtEndModelRelationships-->   Between(Relationships)
/// TimeGroupings              --brtEndModelTimeGroupings-->   Between(TimeGroupings)
/// TimeGroupings              --brtBeginModelTimeGrouping-->  TimeGrouping
/// TimeGrouping               --brtEndModelTimeGrouping-->    TimeGroupings
/// ```
///
/// Member records are read only inside their own collection, and
/// `BrtEndDataModel` is accepted only in `Between`. A record of another kind
/// never moves the position.
enum ModelPosition {
    /// No collection is open; holds the last collection that was closed.
    Between(ModelSection),
    Tables,
    Relationships,
    TimeGroupings,
    /// Inside one grouping of the time-grouping collection.
    TimeGrouping(TimeGrouping),
}

impl ModelPosition {
    /// The last collection closed, when no collection is open.
    const fn between(&self) -> Option<ModelSection> {
        match self {
            Self::Between(last) => Some(*last),
            Self::Tables | Self::Relationships | Self::TimeGroupings | Self::TimeGrouping(_) => {
                None
            },
        }
    }
}

/// Where the workbook stream stands relative to its one Data Model block.
///
/// ```text
/// Before --BrtBeginDataModel--> Open --BrtEndDataModel--> Closed
/// ```
///
/// Every Data Model record is refused in `Before` (except the begin) and in
/// `Closed`; `BrtEndBook` is refused in `Open`.
enum ModelBlockState {
    Before,
    Open {
        /// Offset of the `BrtBeginDataModel` record.
        begin: usize,
        min_version_load: u8,
        position: ModelPosition,
    },
    Closed {
        /// From the start of `BrtBeginDataModel` to the end of
        /// `BrtEndDataModel`.
        range: Range<usize>,
        min_version_load: u8,
    },
}

/// Parse the workbook's optional `BrtBeginDataModel` block.
pub(crate) fn parse_workbook(data: &[u8], limits: SemanticLimits) -> Result<Parsed> {
    let limits = limits.validate()?;
    let mut iterator = Records::bounded(data, limits.raw)?;
    let mut record_count = 0_usize;
    let mut end_book_offset = None;
    let mut state = ModelBlockState::Before;
    let mut layout = Layout::default();
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
        } else if end_book_offset.is_some() && is_model_record(record_kind) {
            return Err(invalid(
                "Data Model record appears after BrtEndBook in the workbook stream",
            ));
        }
        let (begin, min_version_load, position) = match &mut state {
            ModelBlockState::Before => {
                if record_kind == kind::BEGIN_DATA_MODEL.get() {
                    let (header, min_version_load) = parse_begin_data_model(payload)?;
                    layout.begin_header = header;
                    state = ModelBlockState::Open {
                        begin: record_start,
                        min_version_load,
                        position: ModelPosition::Between(ModelSection::None),
                    };
                } else if is_model_record(record_kind) {
                    return Err(invalid(
                        "Data Model collection record appears outside BrtBeginDataModel",
                    ));
                }
                continue;
            },
            ModelBlockState::Closed { .. } => {
                if record_kind == kind::BEGIN_DATA_MODEL.get() {
                    return Err(invalid("multiple BrtBeginDataModel blocks"));
                }
                if is_model_record(record_kind) {
                    return Err(invalid(
                        "Data Model collection record appears after BrtEndDataModel",
                    ));
                }
                continue;
            },
            ModelBlockState::Open {
                begin,
                min_version_load,
                position,
            } => (*begin, *min_version_load, position),
        };

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
                if position.between() != Some(ModelSection::None) {
                    return Err(invalid("nested or overlapping Data Model table collection"));
                }
                *position = ModelPosition::Tables;
                layout.has_tables = true;
            },
            value if value == kind::END_MODEL_TABLES.get() => {
                require_empty(payload, "BrtEndModelTables")?;
                if !matches!(position, ModelPosition::Tables) {
                    return Err(invalid("BrtEndModelTables has no matching begin"));
                }
                *position = ModelPosition::Between(ModelSection::Tables);
            },
            value if value == kind::MODEL_TABLE.get() => {
                if !matches!(position, ModelPosition::Tables) {
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
                let Some(last) = position.between() else {
                    return Err(invalid(
                        "nested or overlapping Data Model relationship collection",
                    ));
                };
                if last > ModelSection::Tables {
                    return Err(invalid(
                        "Data Model relationship collection is duplicated or out of order",
                    ));
                }
                *position = ModelPosition::Relationships;
                layout.has_relationships = true;
            },
            value if value == kind::END_MODEL_RELATIONSHIPS.get() => {
                require_empty(payload, "BrtEndModelRelationships")?;
                if !matches!(position, ModelPosition::Relationships) {
                    return Err(invalid("BrtEndModelRelationships has no matching begin"));
                }
                *position = ModelPosition::Between(ModelSection::Relationships);
            },
            value if value == kind::MODEL_RELATIONSHIP.get() => {
                if !matches!(position, ModelPosition::Relationships) {
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
                let Some(last) = position.between() else {
                    return Err(invalid(
                        "nested or overlapping Data Model time-grouping collection",
                    ));
                };
                if last > ModelSection::Relationships {
                    return Err(invalid(
                        "Data Model time-grouping collection is duplicated or out of order",
                    ));
                }
                *position = ModelPosition::TimeGroupings;
                layout.has_time_groupings = true;
            },
            value if value == kind::END_MODEL_TIME_GROUPINGS.get() => {
                require_empty(payload, "brtEndModelTimeGroupings")?;
                if !matches!(position, ModelPosition::TimeGroupings) {
                    return Err(invalid("brtEndModelTimeGroupings has no matching begin"));
                }
                *position = ModelPosition::Between(ModelSection::TimeGroupings);
            },
            value if value == kind::BEGIN_MODEL_TIME_GROUPING.get() => {
                if !matches!(position, ModelPosition::TimeGroupings) {
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
                *position = ModelPosition::TimeGrouping(grouping);
            },
            value if value == kind::END_MODEL_TIME_GROUPING.get() => {
                require_empty(payload, "brtEndModelTimeGrouping")?;
                // The position is left as written only on the accepted path;
                // the refusal returns from the scan.
                let ModelPosition::TimeGrouping(grouping) =
                    std::mem::replace(position, ModelPosition::TimeGroupings)
                else {
                    return Err(invalid("brtEndModelTimeGrouping has no matching begin"));
                };
                time_groupings.push(grouping);
            },
            value if value == kind::MODEL_TIME_GROUPING_CALC_COL.get() => {
                let ModelPosition::TimeGrouping(grouping) = position else {
                    return Err(invalid("brtModelTimeGroupingCalcCol is outside a grouping"));
                };
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
                if position.between().is_none() {
                    return Err(invalid("BrtEndDataModel closes an unbalanced collection"));
                }
                state = ModelBlockState::Closed {
                    range: begin..record_end,
                    min_version_load,
                };
            },
            _ => {
                // A model block can contain future records. Read them and keep
                // their exact source bytes, but refuse edits that would need to
                // regenerate the block without knowing their structure.
                layout.opaque = true;
            },
        }
    }

    match state {
        ModelBlockState::Before => Ok(Parsed {
            definition: None,
            block_range: None,
            end_book_offset,
            layout,
        }),
        ModelBlockState::Open { .. } => {
            Err(invalid("BrtBeginDataModel has no matching BrtEndDataModel"))
        },
        ModelBlockState::Closed {
            range,
            min_version_load,
        } => {
            if range.start >= range.end {
                return Err(invalid("invalid Data Model record range"));
            }
            let value = Definition {
                min_version_load,
                tables,
                relationships,
                time_groupings,
            };
            validate_definition(&value, limits)?;
            Ok(Parsed {
                definition: Some(value),
                block_range: Some(range),
                end_book_offset,
                layout,
            })
        },
    }
}

#[cfg(test)]
fn collect_records(data: &[u8], limits: RawLimits, max_records: usize) -> Result<Vec<WireRecord>> {
    let mut iterator = Records::bounded(data, limits)?;
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
        return Err(invalid("bVerLoadModelMin is below the required minimum"));
    }
    Ok((header, version))
}

fn require_empty(payload: &[u8], _name: &'static str) -> Result<()> {
    if !payload.is_empty() {
        return Err(invalid("payload must be empty, found bytes"));
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
        Cursor::bounded(&payload[FRT_HEADER_LEN..], context, limits)?,
    ))
}

fn validate_frt_blank(header: [u8; FRT_HEADER_LEN], _context: &'static str) -> Result<()> {
    if header != [0; FRT_HEADER_LEN] {
        return Err(invalid("FRTHeader reserved bytes must be zero"));
    }
    Ok(())
}

/// Validate typed workbook Data Model metadata before publication.
pub(crate) fn validate_definition(value: &Definition, limits: SemanticLimits) -> Result<()> {
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
            return Err(invalid("duplicate Data Model table id"));
        }
        if !names.insert(table.name.to_lowercase()) {
            return Err(invalid("duplicate Data Model table name"));
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
            return Err(invalid("relationship references unknown from_table"));
        }
        if !names.contains(&relationship.to_table.to_lowercase()) {
            return Err(invalid("relationship references unknown to_table"));
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
            return Err(invalid("time grouping references unknown table"));
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
        return Err(invalid("must not be empty"));
    }
    if bytes > MAX_STRING_BYTES || units > MAX_STRING_UNITS {
        return Err(limit(label));
    }
    Ok(())
}

fn serialize_definition(
    value: &Definition,
    layout: Option<&Layout>,
    limits: SemanticLimits,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_definition(value, limits)?;
    if layout.is_some_and(|layout| layout.opaque) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "the Data Model workbook block contains opaque future records",
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
        let mut writer = Encoder::bounded(&mut output, limits.raw)?;
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
    limits: SemanticLimits,
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
    limits: SemanticLimits,
) -> Result<()> {
    if payload > limits.raw.payload() {
        return Err(invalid("Data Model record payload exceeds raw limit"));
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
    serialize_definition(definition, None, SemanticLimits::DEFAULT)
}

fn write_payload(
    limits: RawLimits,
    write: impl FnOnce(&mut PayloadWriter<'_>) -> Result<()>,
) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    let mut writer = PayloadWriter {
        inner: Encoder::bounded(&mut payload, limits)?,
    };
    write(&mut writer)?;
    Ok(payload)
}

struct PayloadWriter<'a> {
    inner: Encoder<&'a mut Vec<u8>>,
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
    limits: SemanticLimits,
) -> Result<Vec<u8>> {
    let parsed = parse_workbook(source, limits)?;
    if parsed.definition.as_ref() != before {
        return Err(Error::PatchConflict {
            part: "data model workbook",
        });
    }
    if before == after {
        return Ok(source.to_vec());
    }
    if parsed.layout.opaque {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "the Data Model workbook block contains opaque future records",
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

fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
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

fn limit(_name: &'static str) -> Error {
    Error::InvalidFormat("Data Model limit exceeded")
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
        serialize_definition(&definition(), None, SemanticLimits::DEFAULT).expect("fixture")
    }

    #[test]
    fn round_trip_model_records() {
        let bytes = fixture();
        let parsed = parse_workbook(&bytes, SemanticLimits::DEFAULT).expect("parse");
        assert_eq!(parsed.definition, Some(definition()));
        assert!(!parsed.layout.opaque);
        assert_eq!(
            patch_workbook(
                &bytes,
                Some(&definition()),
                Some(&definition()),
                SemanticLimits::DEFAULT
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
        assert!(parse_workbook(&bytes, SemanticLimits::DEFAULT).is_err());
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
        assert!(parse_workbook(&bytes, SemanticLimits::DEFAULT).is_err());

        let table = records
            .iter()
            .find(|record| record.kind == kind::MODEL_TABLE.get())
            .expect("table");
        let mut bytes = fixture;
        bytes[table.range.start + 3] = 1;
        assert!(parse_workbook(&bytes, SemanticLimits::DEFAULT).is_err());
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
        let mut writer = Encoder::new(&mut inserted, crate::raw::Limits::DEFAULT);
        writer
            .write_record(Kind::new(4095).expect("kind"), &[1, 2])
            .expect("record");
        bytes.splice(end.range.start..end.range.start, inserted);
        let parsed = parse_workbook(&bytes, SemanticLimits::DEFAULT).expect("read");
        assert!(parsed.layout.opaque);
        assert_eq!(
            patch_workbook(
                &bytes,
                parsed.definition.as_ref(),
                parsed.definition.as_ref(),
                SemanticLimits::DEFAULT,
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
                SemanticLimits::DEFAULT
            )
            .is_err()
        );
    }

    #[test]
    fn workbook_record_index_has_a_finite_limit() {
        let bytes = fixture();
        let mut limits = SemanticLimits::DEFAULT;
        limits.max_records = 0;
        assert!(parse_workbook(&bytes, limits).is_err());
    }

    #[test]
    fn rewrite_byte_limit_is_checked_at_the_exact_boundary() {
        let bytes = fixture();
        let parsed = parse_workbook(&bytes, SemanticLimits::DEFAULT).expect("parse");
        let mut changed = definition();
        changed.min_version_load = 6;

        let mut exact = SemanticLimits::DEFAULT;
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
        let mut writer = Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT);
        writer.write_record(kind::END_BOOK, &[]).expect("end book");
        bytes.extend_from_slice(&fixture);
        assert!(parse_workbook(&bytes, SemanticLimits::DEFAULT).is_err());
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
        Encoder::new(&mut end_book, crate::raw::Limits::DEFAULT)
            .write_record(kind::END_BOOK, &[])
            .expect("end book");
        let mut bytes = fixture;
        bytes.splice(end.range.start..end.range.start, end_book);
        assert!(parse_workbook(&bytes, SemanticLimits::DEFAULT).is_err());
    }

    /// Runs the scanner over a space-separated record sequence and summarizes
    /// the outcome. Tokens: `B`/`E` begin/end the model block; `T( t u T)` the
    /// table collection with two tables; `R( r R)` the relationship
    /// collection; `G( g( c g) G)` the time-grouping collection with one
    /// grouping and one calculated column; `K` BrtEndBook; `X` an unknown
    /// record; ranges are reported in tokens.
    fn scan(sequence: &str) -> String {
        let model = Definition {
            min_version_load: 5,
            tables: vec![
                Table {
                    id: "table-1".to_string(),
                    name: "Sales".to_string(),
                    connection: "Connection".to_string(),
                },
                Table {
                    id: "table-2".to_string(),
                    name: "Dates".to_string(),
                    connection: "Connection".to_string(),
                },
            ],
            relationships: vec![Relationship {
                from_table: "Sales".to_string(),
                from_column: "Date".to_string(),
                to_table: "Dates".to_string(),
                to_column: "Date".to_string(),
            }],
            time_groupings: definition().time_groupings,
        };
        let source =
            serialize_definition(&model, None, SemanticLimits::DEFAULT).expect("source model");
        let records = collect_records(&source, RawLimits::DEFAULT, MAX_RECORDS).expect("records");
        let wire = |wanted: Kind, occurrence: usize| -> Vec<u8> {
            let record = records
                .iter()
                .filter(|record| record.kind == wanted.get())
                .nth(occurrence)
                .expect("source record");
            source[record.range.clone()].to_vec()
        };
        let synthetic = |wanted: Kind, payload: &[u8]| -> Vec<u8> {
            let mut bytes = Vec::new();
            Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT)
                .write_record(wanted, payload)
                .expect("record");
            bytes
        };
        let mut bytes = Vec::new();
        let mut starts = Vec::new();
        for token in sequence.split_whitespace() {
            starts.push(bytes.len());
            bytes.extend_from_slice(&match token {
                "B" => wire(kind::BEGIN_DATA_MODEL, 0),
                "E" => wire(kind::END_DATA_MODEL, 0),
                "T(" => wire(kind::BEGIN_MODEL_TABLES, 0),
                "t" => wire(kind::MODEL_TABLE, 0),
                "u" => wire(kind::MODEL_TABLE, 1),
                "T)" => wire(kind::END_MODEL_TABLES, 0),
                "R(" => wire(kind::BEGIN_MODEL_RELATIONSHIPS, 0),
                "r" => wire(kind::MODEL_RELATIONSHIP, 0),
                "R)" => wire(kind::END_MODEL_RELATIONSHIPS, 0),
                "G(" => wire(kind::BEGIN_MODEL_TIME_GROUPINGS, 0),
                "g(" => wire(kind::BEGIN_MODEL_TIME_GROUPING, 0),
                "c" => wire(kind::MODEL_TIME_GROUPING_CALC_COL, 0),
                "g)" => wire(kind::END_MODEL_TIME_GROUPING, 0),
                "G)" => wire(kind::END_MODEL_TIME_GROUPINGS, 0),
                "K" => synthetic(kind::END_BOOK, &[]),
                "X" => synthetic(Kind::new(4095).expect("kind"), &[1, 2]),
                other => panic!("unknown token {other}"),
            });
        }
        starts.push(bytes.len());
        let token_at = |offset: usize| {
            starts
                .iter()
                .position(|start| *start == offset)
                .expect("record boundary")
        };
        match parse_workbook(&bytes, SemanticLimits::DEFAULT) {
            Err(error) => format!("error: {error}"),
            Ok(parsed) => {
                let end_book = parsed.end_book_offset.map(token_at);
                let range = parsed
                    .block_range
                    .map(|range| token_at(range.start)..token_at(range.end));
                let layout = &parsed.layout;
                let content = parsed.definition.map(|value| {
                    (
                        value.tables.len(),
                        value.relationships.len(),
                        value
                            .time_groupings
                            .iter()
                            .map(|grouping| grouping.columns.len())
                            .collect::<Vec<_>>(),
                    )
                });
                format!(
                    "model {content:?} at {range:?}, sections {}{}{}, opaque {}, end book {end_book:?}",
                    if layout.has_tables { "T" } else { "-" },
                    if layout.has_relationships { "R" } else { "-" },
                    if layout.has_time_groupings { "G" } else { "-" },
                    layout.opaque,
                )
            },
        }
    }

    #[test]
    fn block_scanner_accepts_and_rejects_the_same_sequences() {
        let cases: &[(&str, &str)] = &[
            (
                "",
                "model None at None, sections ---, opaque false, end book None",
            ),
            (
                "X K",
                "model None at None, sections ---, opaque false, end book Some(1)",
            ),
            (
                "B E K",
                "model Some((0, 0, [])) at Some(0..2), sections ---, opaque false, end book Some(2)",
            ),
            (
                "X B E X K",
                "model Some((0, 0, [])) at Some(1..3), sections ---, opaque false, end book Some(4)",
            ),
            (
                "B T( t u T) R( r R) G( g( c g) G) E K",
                "model Some((2, 1, [1])) at Some(0..14), sections TRG, opaque false, end book Some(14)",
            ),
            (
                "B T( t T) E",
                "model Some((1, 0, [])) at Some(0..5), sections T--, opaque false, end book None",
            ),
            (
                "B T( T) E",
                "model Some((0, 0, [])) at Some(0..4), sections T--, opaque false, end book None",
            ),
            (
                "B R( R) E",
                "model Some((0, 0, [])) at Some(0..4), sections -R-, opaque false, end book None",
            ),
            (
                "B G( G) E",
                "model Some((0, 0, [])) at Some(0..4), sections --G, opaque false, end book None",
            ),
            (
                "B T( t u T) G( g( c g) G) E",
                "model Some((2, 0, [1])) at Some(0..11), sections T-G, opaque false, end book None",
            ),
            (
                "B T( t u T) R( r R) E",
                "model Some((2, 1, [])) at Some(0..9), sections TR-, opaque false, end book None",
            ),
            (
                "B R( r R) E",
                "error: Invalid format: relationship references unknown from_table",
            ),
            (
                "B G( g( g) g( c c g) G) E",
                "error: Invalid format: time grouping references unknown table",
            ),
            (
                "B X T( X t X T) X E",
                "model Some((1, 0, [])) at Some(0..9), sections T--, opaque true, end book None",
            ),
            // Collections out of order or repeated.
            (
                "B T( T) T( T) E",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B R( R) T( T) E",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B G( G) T( T) E",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B R( R) R( R) E",
                "error: Invalid format: Data Model relationship collection is duplicated or out of order",
            ),
            (
                "B G( G) R( R) E",
                "error: Invalid format: Data Model relationship collection is duplicated or out of order",
            ),
            (
                "B T( T) G( G) R( R) E",
                "error: Invalid format: Data Model relationship collection is duplicated or out of order",
            ),
            (
                "B G( G) G( G) E",
                "error: Invalid format: Data Model time-grouping collection is duplicated or out of order",
            ),
            // Collections nested in one another.
            (
                "B T( T(",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B T( R(",
                "error: Invalid format: nested or overlapping Data Model relationship collection",
            ),
            (
                "B T( G(",
                "error: Invalid format: nested or overlapping Data Model time-grouping collection",
            ),
            (
                "B R( T(",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B R( R(",
                "error: Invalid format: nested or overlapping Data Model relationship collection",
            ),
            (
                "B R( G(",
                "error: Invalid format: nested or overlapping Data Model time-grouping collection",
            ),
            (
                "B G( T(",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B G( R(",
                "error: Invalid format: nested or overlapping Data Model relationship collection",
            ),
            (
                "B G( G(",
                "error: Invalid format: nested or overlapping Data Model time-grouping collection",
            ),
            (
                "B G( g( T(",
                "error: Invalid format: nested or overlapping Data Model table collection",
            ),
            (
                "B G( g( R(",
                "error: Invalid format: nested or overlapping Data Model relationship collection",
            ),
            (
                "B G( g( G(",
                "error: Invalid format: nested or overlapping Data Model time-grouping collection",
            ),
            (
                "B G( g( g(",
                "error: Invalid format: invalid nested brtBeginModelTimeGrouping",
            ),
            // Ends without a begin.
            (
                "B T)",
                "error: Invalid format: BrtEndModelTables has no matching begin",
            ),
            (
                "B R)",
                "error: Invalid format: BrtEndModelRelationships has no matching begin",
            ),
            (
                "B G)",
                "error: Invalid format: brtEndModelTimeGroupings has no matching begin",
            ),
            (
                "B g)",
                "error: Invalid format: brtEndModelTimeGrouping has no matching begin",
            ),
            (
                "B T( T) T)",
                "error: Invalid format: BrtEndModelTables has no matching begin",
            ),
            (
                "B T( R)",
                "error: Invalid format: BrtEndModelRelationships has no matching begin",
            ),
            (
                "B R( G)",
                "error: Invalid format: brtEndModelTimeGroupings has no matching begin",
            ),
            (
                "B G( T)",
                "error: Invalid format: BrtEndModelTables has no matching begin",
            ),
            (
                "B G( g)",
                "error: Invalid format: brtEndModelTimeGrouping has no matching begin",
            ),
            (
                "B G( g( G)",
                "error: Invalid format: brtEndModelTimeGroupings has no matching begin",
            ),
            (
                "B G( g( g) g)",
                "error: Invalid format: brtEndModelTimeGrouping has no matching begin",
            ),
            // Members outside their collection.
            (
                "B t",
                "error: Invalid format: BrtModelTable is outside BrtBeginModelTables",
            ),
            (
                "B r",
                "error: Invalid format: brtModelRelationship is outside brtBeginModelRelationships",
            ),
            (
                "B g(",
                "error: Invalid format: invalid nested brtBeginModelTimeGrouping",
            ),
            (
                "B c",
                "error: Invalid format: brtModelTimeGroupingCalcCol is outside a grouping",
            ),
            (
                "B T( r",
                "error: Invalid format: brtModelRelationship is outside brtBeginModelRelationships",
            ),
            (
                "B T( g(",
                "error: Invalid format: invalid nested brtBeginModelTimeGrouping",
            ),
            (
                "B T( c",
                "error: Invalid format: brtModelTimeGroupingCalcCol is outside a grouping",
            ),
            (
                "B R( t",
                "error: Invalid format: BrtModelTable is outside BrtBeginModelTables",
            ),
            (
                "B G( t",
                "error: Invalid format: BrtModelTable is outside BrtBeginModelTables",
            ),
            (
                "B G( c",
                "error: Invalid format: brtModelTimeGroupingCalcCol is outside a grouping",
            ),
            (
                "B G( g( t",
                "error: Invalid format: BrtModelTable is outside BrtBeginModelTables",
            ),
            (
                "B G( g( r",
                "error: Invalid format: brtModelRelationship is outside brtBeginModelRelationships",
            ),
            (
                "B T( T) t",
                "error: Invalid format: BrtModelTable is outside BrtBeginModelTables",
            ),
            // The block itself.
            (
                "B",
                "error: Invalid format: BrtBeginDataModel has no matching BrtEndDataModel",
            ),
            (
                "B T( t",
                "error: Invalid format: BrtBeginDataModel has no matching BrtEndDataModel",
            ),
            (
                "B B",
                "error: Invalid format: nested BrtBeginDataModel record",
            ),
            (
                "B T( B",
                "error: Invalid format: nested BrtBeginDataModel record",
            ),
            (
                "B E B E",
                "error: Invalid format: multiple BrtBeginDataModel blocks",
            ),
            (
                "B E X B",
                "error: Invalid format: multiple BrtBeginDataModel blocks",
            ),
            (
                "B E E",
                "error: Invalid format: Data Model collection record appears after BrtEndDataModel",
            ),
            (
                "B E t",
                "error: Invalid format: Data Model collection record appears after BrtEndDataModel",
            ),
            (
                "B E T(",
                "error: Invalid format: Data Model collection record appears after BrtEndDataModel",
            ),
            (
                "E",
                "error: Invalid format: Data Model collection record appears outside BrtBeginDataModel",
            ),
            (
                "t",
                "error: Invalid format: Data Model collection record appears outside BrtBeginDataModel",
            ),
            (
                "T( T)",
                "error: Invalid format: Data Model collection record appears outside BrtBeginDataModel",
            ),
            (
                "X g(",
                "error: Invalid format: Data Model collection record appears outside BrtBeginDataModel",
            ),
            (
                "B T( E",
                "error: Invalid format: BrtEndDataModel closes an unbalanced collection",
            ),
            (
                "B R( E",
                "error: Invalid format: BrtEndDataModel closes an unbalanced collection",
            ),
            (
                "B G( E",
                "error: Invalid format: BrtEndDataModel closes an unbalanced collection",
            ),
            (
                "B G( g( E",
                "error: Invalid format: BrtEndDataModel closes an unbalanced collection",
            ),
            // BrtEndBook.
            (
                "K",
                "model None at None, sections ---, opaque false, end book Some(0)",
            ),
            ("K K", "error: Invalid format: duplicate BrtEndBook record"),
            (
                "B E K K",
                "error: Invalid format: duplicate BrtEndBook record",
            ),
            (
                "K B E",
                "error: Invalid format: Data Model record appears after BrtEndBook in the workbook stream",
            ),
            (
                "K X",
                "model None at None, sections ---, opaque false, end book Some(0)",
            ),
            (
                "B K E",
                "error: Invalid format: BrtEndBook appears before the Data Model block is closed",
            ),
            (
                "B T( K",
                "error: Invalid format: BrtEndBook appears before the Data Model block is closed",
            ),
            (
                "B E K t",
                "error: Invalid format: Data Model record appears after BrtEndBook in the workbook stream",
            ),
        ];
        let mismatches: Vec<String> = cases
            .iter()
            .filter_map(|(sequence, expected)| {
                let actual = scan(sequence);
                (actual != *expected).then(|| format!("(\"{sequence}\", \"{actual}\"),"))
            })
            .collect();
        assert!(mismatches.is_empty(), "\n{}\n", mismatches.join("\n"));
    }

    #[test]
    fn model_insertion_requires_an_end_book_anchor() {
        assert!(patch_workbook(&[], None, Some(&definition()), SemanticLimits::DEFAULT,).is_err());
    }
}
