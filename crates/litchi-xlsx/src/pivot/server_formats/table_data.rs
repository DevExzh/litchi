//! Source-bound support for the MS-XLSX `pivotTableData` extension (C444).
//!
//! The owner is intentionally scalar and source preserving.  It reads the
//! authored rows and cells of a non-worksheet PivotTable without allocating a
//! dense matrix, and it edits only existing `v` text or attributes on an
//! existing `x` child.  Relationship IDs, Part names, extension URIs, and
//! XML prefixes stay below this module's ordinary Workbook facade.

#[cfg(test)]
use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, SourceLimits};

use super::cached_unique_names::PivotCacheId;
use super::*;
use crate::Workbook;
use crate::error::{Result, invalid};

const DATA_PAYLOAD: &[u8] = b"pivotTableData";
const ROW_PAYLOAD: &[u8] = b"pivotRow";
const CELL_PAYLOAD: &[u8] = b"c";
const VALUE_PAYLOAD: &[u8] = b"v";
const EXTRA_PAYLOAD: &[u8] = b"x";
const ROW_ITEMS_PAYLOAD: &[u8] = b"rowItems";
const COL_ITEMS_PAYLOAD: &[u8] = b"colItems";
const ROW_COUNT_ATTRIBUTE: &[u8] = b"rowCount";
const COLUMN_COUNT_ATTRIBUTE: &[u8] = b"columnCount";
const COUNT_ATTRIBUTE: &[u8] = b"count";
const CACHE_ID_ATTRIBUTE: &[u8] = b"cacheId";
const ROW_INDEX_ATTRIBUTE: &[u8] = b"r";
const COLUMN_INDEX_ATTRIBUTE: &[u8] = b"i";
const TYPE_ATTRIBUTE: &[u8] = b"t";
const FORMAT_INDEX_ATTRIBUTE: &[u8] = b"in";
const BACKGROUND_COLOR_ATTRIBUTE: &[u8] = b"bc";
const FOREGROUND_COLOR_ATTRIBUTE: &[u8] = b"fc";
const ITALIC_ATTRIBUTE: &[u8] = b"i";
const UNDERLINE_ATTRIBUTE: &[u8] = b"un";
const STRIKE_ATTRIBUTE: &[u8] = b"st";
const BOLD_ATTRIBUTE: &[u8] = b"b";
const MAX_UTF16_TEXT: usize = 65_535;
const MAX_AUTHORED_ROWS: usize = 100_000;
const MAX_AUTHORED_CELLS: usize = 100_000;
const MAX_SCALAR_DIMENSION: usize = (1usize << 31) - 1;
const MAX_RETAINED_BYTES: usize = 64 * 1024 * 1024;

/// Resource ceilings for the typed `pivotTableData` owner.
///
/// The package's [`SourceLimits`] remains authoritative for physical Parts,
/// XML events, attributes, depth, namespaces, and relationship closure.  This
/// focused policy adds the semantic ceilings that cannot be expressed by an
/// OPC package profile: authored row and cell elements, logical scalar
/// dimensions, and bytes retained by the sparse typed view and its source
/// bindings.  Setters only lower the finite owner ceilings.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_authored_rows: usize,
    max_authored_cells: usize,
    max_row_count: usize,
    max_column_count: usize,
    max_retained_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

impl Limits {
    /// Construct the default bounded C444 semantic policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_authored_rows: MAX_AUTHORED_ROWS,
            max_authored_cells: MAX_AUTHORED_CELLS,
            max_row_count: MAX_SCALAR_DIMENSION,
            max_column_count: MAX_SCALAR_DIMENSION,
            max_retained_bytes: MAX_RETAINED_BYTES,
        }
    }

    /// Set a lower authored-row ceiling.
    #[must_use]
    pub const fn with_max_authored_rows(mut self, value: usize) -> Self {
        self.max_authored_rows = if value < MAX_AUTHORED_ROWS {
            value
        } else {
            MAX_AUTHORED_ROWS
        };
        self
    }

    /// Set a lower authored-cell ceiling.
    #[must_use]
    pub const fn with_max_authored_cells(mut self, value: usize) -> Self {
        self.max_authored_cells = if value < MAX_AUTHORED_CELLS {
            value
        } else {
            MAX_AUTHORED_CELLS
        };
        self
    }

    /// Set a lower logical `rowCount` ceiling.
    #[must_use]
    pub const fn with_max_row_count(mut self, value: usize) -> Self {
        self.max_row_count = if value < MAX_SCALAR_DIMENSION {
            value
        } else {
            MAX_SCALAR_DIMENSION
        };
        self
    }

    /// Set a lower logical `columnCount` ceiling.
    #[must_use]
    pub const fn with_max_column_count(mut self, value: usize) -> Self {
        self.max_column_count = if value < MAX_SCALAR_DIMENSION {
            value
        } else {
            MAX_SCALAR_DIMENSION
        };
        self
    }

    /// Set a lower aggregate retained-byte ceiling for this semantic owner.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, value: usize) -> Self {
        self.max_retained_bytes = if value < MAX_RETAINED_BYTES {
            value
        } else {
            MAX_RETAINED_BYTES
        };
        self
    }

    /// Maximum authored row elements admitted by this policy.
    #[must_use]
    pub const fn max_authored_rows(self) -> usize {
        self.max_authored_rows
    }

    /// Maximum authored cell elements admitted by this policy.
    #[must_use]
    pub const fn max_authored_cells(self) -> usize {
        self.max_authored_cells
    }

    /// Maximum logical `rowCount` admitted by this policy.
    #[must_use]
    pub const fn max_row_count(self) -> usize {
        self.max_row_count
    }

    /// Maximum logical `columnCount` admitted by this policy.
    #[must_use]
    pub const fn max_column_count(self) -> usize {
        self.max_column_count
    }

    /// Maximum bytes retained by this semantic owner.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }
}

#[derive(Clone, Copy, Debug)]
struct RetainedBudget {
    maximum: usize,
    retained: usize,
    max_authored_rows: usize,
    max_authored_cells: usize,
    max_row_count: usize,
    max_column_count: usize,
}

impl RetainedBudget {
    fn new(policy: Limits) -> Self {
        Self {
            maximum: policy.max_retained_bytes,
            retained: 0,
            max_authored_rows: policy.max_authored_rows,
            max_authored_cells: policy.max_authored_cells,
            max_row_count: policy.max_row_count,
            max_column_count: policy.max_column_count,
        }
    }

    fn charge(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        let observed = self
            .retained
            .checked_add(amount)
            .ok_or_else(|| invalid(format!("{resource} retained bytes overflow")))?;
        if observed > self.maximum {
            return Err(Error::limit(
                LimitKind::RetainedBytes,
                observed,
                self.maximum,
            ));
        }
        self.retained = observed;
        Ok(())
    }

    fn charge_count(&mut self, observed: usize, maximum: usize, kind: LimitKind) -> Result<()> {
        if observed > maximum {
            return Err(Error::limit(kind, observed, maximum));
        }
        Ok(())
    }

    fn check_row_count(&self, value: u32) -> Result<()> {
        let observed = usize::try_from(value).unwrap_or(usize::MAX);
        if observed > self.max_row_count {
            return Err(Error::limit(
                LimitKind::RowCount,
                observed,
                self.max_row_count,
            ));
        }
        Ok(())
    }

    fn check_column_count(&self, value: u32) -> Result<()> {
        let observed = usize::try_from(value).unwrap_or(usize::MAX);
        if observed > self.max_column_count {
            return Err(Error::limit(
                LimitKind::ColumnCount,
                observed,
                self.max_column_count,
            ));
        }
        Ok(())
    }
}

fn staging_retained_bytes(cells: usize) -> Result<usize> {
    let index_item = size_of::<usize>()
        .checked_add(
            size_of::<(PivotCellAddress, CellPosition)>()
                .checked_mul(2)
                .ok_or_else(|| invalid("pivotTableData staged address index overflows"))?,
        )
        .ok_or_else(|| invalid("pivotTableData staged address index overflows"))?;
    let cell_bytes = checked_bytes(
        cells,
        size_of::<CellState>(),
        "pivotTableData staged cell states",
    )?;
    let index_bytes = checked_bytes(cells, index_item, "pivotTableData staged indexes")?;
    size_of::<Vec<CellState>>()
        .checked_add(size_of::<Vec<usize>>())
        .and_then(|bytes| bytes.checked_add(size_of::<HashMap<PivotCellAddress, CellPosition>>()))
        .and_then(|bytes| bytes.checked_add(cell_bytes))
        .and_then(|bytes| bytes.checked_add(index_bytes))
        .ok_or_else(|| invalid("pivotTableData staged state bytes overflow"))
}

fn replacement_text_retained_bytes(length: usize) -> Result<usize> {
    length
        .checked_mul(3)
        .and_then(|bytes| bytes.checked_add(size_of::<Arc<str>>() + 16))
        .ok_or_else(|| invalid("pivotTableData staged text bytes overflow"))
}

/// Admit the final staged replacement ledger and any caller-owned transient
/// input before constructing its `Arc<str>`.  The transaction maintains the
/// per-position replacement ledger, avoiding a scan of every authored cell
/// per edit while still charging the complete staged state.
fn admit_staged_text(
    before: &Snapshot,
    staged_cells: usize,
    staged_text_bytes_after: usize,
    transient_bytes: usize,
) -> Result<()> {
    let staging = staging_retained_bytes(staged_cells)?;
    let observed = before
        .retained_bytes
        .checked_add(staging)
        .and_then(|bytes| bytes.checked_add(staged_text_bytes_after))
        .and_then(|bytes| bytes.checked_add(transient_bytes))
        .ok_or_else(|| invalid("pivotTableData staged retained bytes overflow"))?;
    enforce_retained_limit(observed, before.limits.max_retained_bytes)
}

#[derive(Clone, Copy, Debug)]
enum CellPosition {
    Unique(usize),
    Ambiguous,
}

fn build_cell_positions(before: &Snapshot) -> Result<HashMap<PivotCellAddress, CellPosition>> {
    let mut positions = HashMap::new();
    positions
        .try_reserve(before.cells.len())
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData cell address index",
            source,
        })?;
    for (index, record) in before.cells.iter().enumerate() {
        let (Some(row), Some(column)) = (record.state.row_index, record.state.column_index) else {
            continue;
        };
        let address = PivotCellAddress::new(row, column);
        match positions.get_mut(&address) {
            Some(position) => *position = CellPosition::Ambiguous,
            None => {
                positions.insert(address, CellPosition::Unique(index));
            },
        }
    }
    Ok(positions)
}

fn enforce_retained_limit(observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::limit(LimitKind::RetainedBytes, observed, maximum));
    }
    Ok(())
}

#[cfg(test)]
thread_local! {
    static LOCAL_SUBTREE_BOUNDARY_VISITS: Cell<usize> = const { Cell::new(0) };
}

/// A source-authored row/column coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PivotCellAddress {
    /// Authored `pivotRow@r` value.
    pub row: u32,
    /// Authored `c@i` value.
    pub column: u32,
}

impl PivotCellAddress {
    #[must_use]
    pub const fn new(row: u32, column: u32) -> Self {
        Self { row, column }
    }
}

impl From<(u32, u32)> for PivotCellAddress {
    fn from((row, column): (u32, u32)) -> Self {
        Self::new(row, column)
    }
}

/// The semantic type selected by `CT_PivotValueCell@t`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotCellType {
    Boolean,
    Number,
    Error,
    Text,
    DateTime,
    Blank,
}

/// A safe scalar value replacement.  Numeric and date-time cells remain
/// readable but deliberately have no edit variant in this first owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CellChange {
    Boolean(bool),
    Text(String),
    Error(String),
    Blank,
}

impl CellChange {
    #[must_use]
    pub fn text(value: impl Into<String>) -> Self {
        Self::Text(value.into())
    }

    #[must_use]
    pub fn error(value: impl Into<String>) -> Self {
        Self::Error(value.into())
    }
}

/// Typed view of one optional cell-extra formatting record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotValueCellExtraView {
    /// C510 server-format index from `in`, when present and proven bounded.
    pub format_index: Option<u32>,
    /// Background RGB value from `bc`.
    pub background_color: Option<u32>,
    /// Foreground RGB value from `fc`.
    pub foreground_color: Option<u32>,
    /// Explicit italic flag from `i`.
    pub italic: Option<bool>,
    /// Explicit underline flag from `un`.
    pub underline: Option<bool>,
    /// Explicit strike-through flag from `st`.
    pub strike: Option<bool>,
    /// Explicit bold flag from `b`.
    pub bold: Option<bool>,
}

/// A tri-state edit for one typed cell-extra attribute.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Attribute<T> {
    Keep,
    Set(T),
    Clear,
}

impl<T> Attribute<T> {
    #[must_use]
    pub const fn keep() -> Self {
        Self::Keep
    }

    #[must_use]
    pub const fn clear() -> Self {
        Self::Clear
    }
}

/// Explicit scalar edits for an existing `x` element.  The `in` association
/// is intentionally read-only and therefore has no field here.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellExtra {
    pub background_color: Attribute<String>,
    pub foreground_color: Attribute<String>,
    pub italic: Attribute<bool>,
    pub underline: Attribute<bool>,
    pub strike: Attribute<bool>,
    pub bold: Attribute<bool>,
}

impl CellExtra {
    #[must_use]
    pub const fn keep() -> Self {
        Self {
            background_color: Attribute::Keep,
            foreground_color: Attribute::Keep,
            italic: Attribute::Keep,
            underline: Attribute::Keep,
            strike: Attribute::Keep,
            bold: Attribute::Keep,
        }
    }
}

/// Typed view of one authored PivotValueCell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotValueCellView {
    row_index: Option<u32>,
    column_index: Option<u32>,
    kind: PivotCellType,
    value_text: Arc<str>,
    extra: Option<PivotValueCellExtraView>,
}

impl PivotValueCellView {
    #[must_use]
    pub const fn row_index(&self) -> Option<u32> {
        self.row_index
    }

    #[must_use]
    pub const fn column_index(&self) -> Option<u32> {
        self.column_index
    }

    #[must_use]
    pub const fn kind(&self) -> PivotCellType {
        self.kind
    }

    #[must_use]
    pub fn value_text(&self) -> &str {
        self.value_text.as_ref()
    }

    #[must_use]
    pub fn extra(&self) -> Option<&PivotValueCellExtraView> {
        self.extra.as_ref()
    }

    #[must_use]
    pub fn address(&self) -> Option<PivotCellAddress> {
        Some(PivotCellAddress::new(self.row_index?, self.column_index?))
    }
}

/// Typed view of one authored pivot row.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotRowView {
    row_index: Option<u32>,
    cells: Box<[PivotValueCellView]>,
}

impl PivotRowView {
    #[must_use]
    pub const fn row_index(&self) -> Option<u32> {
        self.row_index
    }

    #[must_use]
    pub fn cells(&self) -> &[PivotValueCellView] {
        &self.cells
    }
}

/// Read-only diagnostic state attached to a data view.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotTableDataDiagnostic {
    None,
    MceAmbiguous,
    CacheClosureUnresolved,
    ServerFormatIndexUnresolved,
    OwnerShapeUnresolved,
}

/// Typed source-bound `pivotTableData` view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableDataView {
    table_name: String,
    cache_id: PivotCacheId,
    row_count: u32,
    column_count: u32,
    rows: Box<[PivotRowView]>,
    diagnostic: PivotTableDataDiagnostic,
}

impl PivotTableDataView {
    #[must_use]
    pub fn table_name(&self) -> &str {
        &self.table_name
    }

    #[must_use]
    pub const fn cache_id(&self) -> PivotCacheId {
        self.cache_id
    }

    #[must_use]
    pub const fn row_count(&self) -> u32 {
        self.row_count
    }

    #[must_use]
    pub const fn column_count(&self) -> u32 {
        self.column_count
    }

    #[must_use]
    pub fn rows(&self) -> &[PivotRowView] {
        &self.rows
    }

    /// Resolve only a fully-authored, unique coordinate.
    pub fn cell(
        &self,
        address: impl Into<PivotCellAddress>,
    ) -> Result<Option<&PivotValueCellView>> {
        let address = address.into();
        let mut found = None;
        for row in &self.rows {
            if row.row_index != Some(address.row) {
                continue;
            }
            for cell in &row.cells {
                if cell.column_index == Some(address.column) {
                    if found.is_some() {
                        return Err(invalid("pivotTableData cell coordinate is ambiguous"));
                    }
                    found = Some(cell);
                }
            }
        }
        Ok(found)
    }

    #[must_use]
    pub const fn diagnostic_status(&self) -> PivotTableDataDiagnostic {
        self.diagnostic
    }

    #[must_use]
    pub const fn is_editable(&self) -> bool {
        matches!(self.diagnostic, PivotTableDataDiagnostic::None)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ExtraState {
    format_index: Option<u32>,
    background_color: Option<u32>,
    foreground_color: Option<u32>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    bold: Option<bool>,
}

impl ExtraState {
    fn view(&self) -> PivotValueCellExtraView {
        PivotValueCellExtraView {
            format_index: self.format_index,
            background_color: self.background_color,
            foreground_color: self.foreground_color,
            italic: self.italic,
            underline: self.underline,
            strike: self.strike,
            bold: self.bold,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct CellState {
    row_index: Option<u32>,
    column_index: Option<u32>,
    kind: PivotCellType,
    value_text: Arc<str>,
    extra: Option<ExtraState>,
}

impl CellState {
    fn view(&self) -> PivotValueCellView {
        PivotValueCellView {
            row_index: self.row_index,
            column_index: self.column_index,
            kind: self.kind,
            value_text: Arc::clone(&self.value_text),
            extra: self.extra.as_ref().map(ExtraState::view),
        }
    }
}

#[derive(Clone, Debug)]
struct AttributeSourceData {
    value: Range<usize>,
    whole: Range<usize>,
}

#[derive(Clone, Debug)]
struct ExtraSource {
    start_tag: Range<usize>,
    background_color: Option<AttributeSourceData>,
    foreground_color: Option<AttributeSourceData>,
    italic: Option<AttributeSourceData>,
    underline: Option<AttributeSourceData>,
    strike: Option<AttributeSourceData>,
    bold: Option<AttributeSourceData>,
}

#[derive(Clone, Debug)]
struct CellSource {
    value_element: Range<usize>,
    value_start_tag: Range<usize>,
    value_qname: Vec<u8>,
    extra: Option<ExtraSource>,
}

#[derive(Clone, Debug)]
struct CellRecord {
    source: CellSource,
    state: CellState,
}

/// Exact source-bound snapshot of one C444 owner.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: PivotTableDataView,
    limits: Limits,
    table: SourcePart,
    cache: SourcePart,
    connections: Option<SourcePart>,
    workbook: SourcePart,
    workbook_owner: Arc<Vec<u8>>,
    workbook_context: Arc<Vec<u8>>,
    table_owner: Range<usize>,
    extension_owner: Range<usize>,
    cells: Box<[CellRecord]>,
    retained_bytes: usize,
    selection: usize,
}

impl Snapshot {
    pub fn load<'a>(
        package: &OpcPackage,
        selector: impl Into<PivotTableSelector<'a>>,
    ) -> Result<Self> {
        Self::load_with_limits(package, selector, &Limits::default())
    }

    /// Load a C444 snapshot with an owner-specific semantic resource policy.
    pub(crate) fn load_with_limits<'a>(
        package: &OpcPackage,
        selector: impl Into<PivotTableSelector<'a>>,
        limits: &Limits,
    ) -> Result<Self> {
        let selector = selector.into();
        let policy = *limits;
        let graph = Graph::load_for_table_data(package)?;
        let selection = match selector {
            PivotTableSelector::Name(name) if graph.ordinary_name_collision(name) => {
                return Err(invalid(
                    "PivotTable selector name collides with an ordinary worksheet table",
                ));
            },
            selector => resolve_reference(&graph.refs, selector)?,
        };
        let reference = graph
            .refs
            .get(selection)
            .ok_or_else(|| invalid("PivotTable selector did not resolve"))?;
        let table = &reference.table;
        let cache = graph
            .caches
            .get(&table.cache_id)
            .ok_or_else(|| invalid("PivotTable cache graph is incomplete"))?;
        let mut retained = RetainedBudget::new(policy);
        let mut source_bytes = table.source.bytes.len();
        source_bytes = source_bytes
            .checked_add(cache.source.bytes.len())
            .and_then(|length| length.checked_add(graph.workbook.bytes.len()))
            .and_then(|length| length.checked_add(graph.workbook_owner.len()))
            .and_then(|length| length.checked_add(graph.workbook_context.len()))
            .ok_or_else(|| invalid("pivotTableData source bytes overflow"))?;
        if let Some(connections) = graph.connections.as_ref() {
            source_bytes = source_bytes
                .checked_add(connections.bytes.len())
                .ok_or_else(|| invalid("pivotTableData source bytes overflow"))?;
        }
        for source in [&table.source, &cache.source, &graph.workbook] {
            source_bytes = source_bytes
                .checked_add(source.retained_closure_bytes()?)
                .ok_or_else(|| invalid("pivotTableData source closure bytes overflow"))?;
        }
        if let Some(connections) = graph.connections.as_ref() {
            source_bytes = source_bytes
                .checked_add(connections.retained_closure_bytes()?)
                .ok_or_else(|| invalid("pivotTableData source closure bytes overflow"))?;
        }
        retained.charge(source_bytes, "pivotTableData source bindings")?;
        let scan = scan_xml_with_mce(
            table.source.bytes.as_slice(),
            "pivotTableDefinition",
            package.read_limits(),
        )?;
        let parsed = parse_data_payload(
            &scan,
            table.source.bytes.as_slice(),
            table.cache_id,
            package.read_limits(),
            &mut retained,
        )?;
        let closure_mce_ambiguous = graph.workbook_mce_ambiguous || cache.mce_ambiguous;
        let diagnostic = if parsed.mce_ambiguous || closure_mce_ambiguous {
            PivotTableDataDiagnostic::MceAmbiguous
        } else if graph.workbook_closure_diagnostic || cache.closure_diagnostic {
            PivotTableDataDiagnostic::CacheClosureUnresolved
        } else if parsed.diagnostic {
            PivotTableDataDiagnostic::OwnerShapeUnresolved
        } else if parsed.server_format_index_unresolved {
            PivotTableDataDiagnostic::ServerFormatIndexUnresolved
        } else {
            PivotTableDataDiagnostic::None
        };
        let parsed_row_count = parsed.row_cells.len();
        let parsed_cell_count = parsed.cells.len();
        let view_bytes =
            semantic_view_bytes(parsed_row_count, parsed_cell_count, table.name.len())?;
        retained.charge(view_bytes, "pivotTableData semantic view")?;
        let records = parsed.cells.into_boxed_slice();
        let mut rows = Vec::new();
        rows.try_reserve_exact(parsed_row_count)
            .map_err(|source| Error::Allocation {
                resource: "pivotTableData semantic rows",
                source,
            })?;
        for row in parsed.row_cells {
            let mut cells = Vec::new();
            cells
                .try_reserve_exact(row.len())
                .map_err(|source| Error::Allocation {
                    resource: "pivotTableData semantic cells",
                    source,
                })?;
            let row_index = row
                .first()
                .and_then(|index| records.get(*index))
                .and_then(|record| record.state.row_index);
            for index in &row {
                cells.push(records[*index].state.view());
            }
            rows.push(PivotRowView {
                row_index,
                cells: cells.into_boxed_slice(),
            });
        }
        let value = PivotTableDataView {
            table_name: table.name.clone(),
            cache_id: PivotCacheId(parsed.cache_id),
            row_count: parsed.row_count,
            column_count: parsed.column_count,
            rows: rows.into_boxed_slice(),
            diagnostic,
        };
        Ok(Self {
            value,
            limits: policy,
            table: table.source.clone(),
            cache: cache.source.clone(),
            connections: graph.connections.clone(),
            workbook: graph.workbook.clone(),
            workbook_owner: Arc::clone(&graph.workbook_owner),
            workbook_context: Arc::clone(&graph.workbook_context),
            table_owner: table.owner.clone(),
            extension_owner: parsed.owner,
            cells: records,
            retained_bytes: retained.retained,
            selection,
        })
    }

    #[must_use]
    pub fn data(&self) -> &PivotTableDataView {
        &self.value
    }

    #[must_use]
    pub fn table_name(&self) -> &str {
        self.value.table_name()
    }

    #[must_use]
    pub const fn cache_id(&self) -> PivotCacheId {
        self.value.cache_id()
    }

    #[must_use]
    pub fn rows(&self) -> &[PivotRowView] {
        self.value.rows()
    }

    pub fn cell(
        &self,
        address: impl Into<PivotCellAddress>,
    ) -> Result<Option<&PivotValueCellView>> {
        self.value.cell(address)
    }

    #[must_use]
    pub const fn diagnostic_status(&self) -> PivotTableDataDiagnostic {
        self.value.diagnostic_status()
    }

    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.table.bytes.as_slice()
    }

    #[must_use]
    pub fn table_part(&self) -> &PackURI {
        &self.table.name
    }

    #[must_use]
    pub fn source_owner_range(&self) -> Range<usize> {
        self.table_owner.clone()
    }

    /// The owner-specific semantic resource policy used for this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    fn same_source(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.table.same_source(&other.table)
            && self.cache.same_source(&other.cache)
            && same_optional_source(&self.connections, &other.connections)
            && self.workbook.same_source(&other.workbook)
    }

    fn same_readset(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.table.same_closure(&other.table)
            && self.cache.same_closure(&other.cache)
            && same_optional_closure(&self.connections, &other.connections)
            && self.workbook.same_closure(&other.workbook)
            && self.workbook_owner.as_slice() == other.workbook_owner.as_slice()
            && self.workbook_context.as_slice() == other.workbook_context.as_slice()
    }
}

fn same_optional_source(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_source(right),
        _ => false,
    }
}

fn same_optional_closure(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_closure(right),
        _ => false,
    }
}

fn resolve_reference(references: &[Reference], selector: PivotTableSelector<'_>) -> Result<usize> {
    match selector {
        PivotTableSelector::Position(position) => {
            if position >= references.len() {
                return Err(invalid("PivotTable selector position is out of range"));
            }
            Ok(position)
        },
        PivotTableSelector::Name(name) => {
            let mut found = None;
            for (index, reference) in references.iter().enumerate() {
                if reference.name == name {
                    if found.is_some() {
                        return Err(invalid("PivotTable selector name is ambiguous"));
                    }
                    found = Some(index);
                }
            }
            found.ok_or_else(|| invalid("PivotTable selector name did not resolve"))
        },
    }
}

#[derive(Clone, Debug)]
struct ParsedData {
    row_count: u32,
    column_count: u32,
    cache_id: u32,
    owner: Range<usize>,
    cells: Vec<CellRecord>,
    row_cells: Vec<Vec<usize>>,
    mce_ambiguous: bool,
    server_format_index_unresolved: bool,
    diagnostic: bool,
}

fn checked_bytes(count: usize, item: usize, resource: &'static str) -> Result<usize> {
    count
        .checked_mul(item)
        .ok_or_else(|| invalid(format!("{resource} retained bytes overflow")))
}

fn structural_retained_bytes(rows: usize, cells: usize) -> Result<usize> {
    let row_item = size_of::<Vec<usize>>()
        .checked_add(size_of::<PivotRowView>())
        .and_then(|size| size.checked_mul(2))
        .ok_or_else(|| invalid("pivotTableData row allocation size overflows"))?;
    let row_vectors = checked_bytes(rows, row_item, "pivotTableData row vectors")?;
    let cell_item = size_of::<CellRecord>()
        .checked_mul(2)
        .and_then(|size| size.checked_add(size_of::<usize>().saturating_mul(8)))
        .ok_or_else(|| invalid("pivotTableData cell allocation size overflows"))?;
    let cell_storage = checked_bytes(cells, cell_item, "pivotTableData cell storage")?;
    row_vectors
        .checked_add(cell_storage)
        .ok_or_else(|| invalid("pivotTableData structural retained bytes overflow"))
}

fn semantic_view_bytes(rows: usize, cells: usize, table_name_bytes: usize) -> Result<usize> {
    let row_item = size_of::<PivotRowView>()
        .checked_mul(2)
        .and_then(|size| size.checked_add(size_of::<Vec<PivotValueCellView>>()))
        .ok_or_else(|| invalid("pivotTableData view row allocation size overflows"))?;
    let row_storage = size_of::<Vec<PivotRowView>>()
        .checked_add(checked_bytes(rows, row_item, "pivotTableData view rows")?)
        .ok_or_else(|| invalid("pivotTableData view row bytes overflow"))?;
    let cell_item = size_of::<PivotValueCellView>()
        .checked_mul(2)
        .ok_or_else(|| invalid("pivotTableData view cell allocation size overflows"))?;
    let cell_storage = checked_bytes(cells, cell_item, "pivotTableData view cells")?;
    table_name_bytes
        .checked_mul(2)
        .and_then(|length| length.checked_add(size_of::<String>()))
        .and_then(|length| length.checked_add(row_storage))
        .and_then(|length| length.checked_add(cell_storage))
        .and_then(|length| length.checked_add(size_of::<PivotTableDataView>()))
        .ok_or_else(|| invalid("pivotTableData semantic view bytes overflow"))
}

fn parse_data_payload(
    scan: &XmlScan,
    source: &[u8],
    expected_cache_id: u32,
    read_limits: SourceLimits,
    retained: &mut RetainedBudget,
) -> Result<ParsedData> {
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("PivotTable Part has no root"))?;
    let owner_exts = extension_exts(scan, root, PIVOT_TABLE_DATA_URI, true)?;
    if owner_exts.is_empty() {
        return Err(invalid("PivotTable has no pivotTableData extension"));
    }
    let mut diagnostic = owner_exts.len() > 1;
    // Validate every recognized owner before selecting the first one.  A
    // duplicate is a read-only diagnostic, but an oversized duplicate must
    // still be refused before it can be ignored.
    for owner_ext in &owner_exts {
        check_fragment_limit(owner_ext, "pivotTableData owner fragment")?;
        if !validate_data_owner_ext(scan, owner_ext)? {
            diagnostic = true;
        }
    }
    let mut payload = None;
    for candidate in scan.elements.iter().filter(|candidate| {
        candidate.ns.as_ref() == EXT_NS && candidate.local.as_slice() == DATA_PAYLOAD
    }) {
        if !is_owned_payload(scan, root, candidate, PIVOT_TABLE_DATA_URI) {
            continue;
        }
        check_fragment_limit(candidate, "pivotTableData payload fragment")?;
        if payload.is_some() {
            diagnostic = true;
        } else {
            payload = Some(candidate);
        }
    }
    let payload = payload.ok_or_else(|| invalid("pivotTableData extension has no payload"))?;
    if payload.has_non_whitespace_cdata || payload.has_non_whitespace_text {
        diagnostic = true;
    }
    let row_count = required_attr(payload, ROW_COUNT_ATTRIBUTE, "pivotTableData rowCount")
        .and_then(|value| parse_u32(value, "pivotTableData rowCount"))?;
    let column_count = required_attr(
        payload,
        COLUMN_COUNT_ATTRIBUTE,
        "pivotTableData columnCount",
    )
    .and_then(|value| parse_u32(value, "pivotTableData columnCount"))?;
    let cache_id = required_attr(payload, CACHE_ID_ATTRIBUTE, "pivotTableData cacheId")
        .and_then(|value| parse_u32(value, "pivotTableData cacheId"))?;
    if cache_id != expected_cache_id {
        return Err(invalid("pivotTableData cacheId does not match PivotTable"));
    }
    for attribute in &payload.attrs {
        if !attribute.ns.is_empty()
            || !matches!(
                attribute.local.as_slice(),
                ROW_COUNT_ATTRIBUTE | COLUMN_COUNT_ATTRIBUTE | CACHE_ID_ATTRIBUTE
            )
        {
            return Err(invalid("pivotTableData has an unknown attribute"));
        }
    }
    retained.check_row_count(row_count)?;
    retained.check_column_count(column_count)?;
    validate_standard_item_count(scan, root, ROW_ITEMS_PAYLOAD, row_count)?;
    validate_standard_item_count(scan, root, COL_ITEMS_PAYLOAD, column_count)?;

    let row_count_actual = direct_children(scan, payload)
        .filter(|element| element.ns.as_ref() == EXT_NS && element.local.as_slice() == ROW_PAYLOAD)
        .count();
    if row_count_actual == 0 || row_count_actual != usize::try_from(row_count).unwrap_or(usize::MAX)
    {
        return Err(invalid(
            "pivotTableData rowCount does not equal pivotRow count",
        ));
    }
    for child in direct_children(scan, payload) {
        if child.ns != payload.ns || child.local.as_slice() != ROW_PAYLOAD {
            return Err(invalid("pivotTableData has an unexpected direct child"));
        }
    }
    retained.charge_count(
        row_count_actual,
        retained.max_authored_rows,
        LimitKind::AuthoredRows,
    )?;
    let mut authored_cells = 0usize;
    for row in direct_children(scan, payload) {
        let actual_cells = direct_children(scan, row).count();
        authored_cells = authored_cells
            .checked_add(actual_cells)
            .ok_or_else(|| invalid("pivotTableData authored cell count overflows"))?;
    }
    retained.charge_count(
        authored_cells,
        retained.max_authored_cells,
        LimitKind::AuthoredCells,
    )?;
    let structural_bytes = structural_retained_bytes(row_count_actual, authored_cells)?;
    retained.charge(structural_bytes, "pivotTableData authored indexes")?;
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(authored_cells)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData source cell records",
            source,
        })?;
    let mut row_cells = Vec::new();
    row_cells
        .try_reserve_exact(row_count_actual)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData row cell indexes",
            source,
        })?;
    let mut seen_rows = HashSet::new();
    seen_rows
        .try_reserve(row_count_actual)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData row coordinate index",
            source,
        })?;
    for row in direct_children(scan, payload) {
        validate_row_attributes(row)?;
        let row_index = optional_attr(row, ROW_INDEX_ATTRIBUTE)?
            .map(|value| parse_u32(value, "pivotRow r"))
            .transpose()?;
        if let Some(index) = row_index {
            if index >= row_count || !seen_rows.insert(index) {
                return Err(invalid(
                    "pivotTableData has an invalid or duplicate row index",
                ));
            }
        }
        let count = required_attr(row, COUNT_ATTRIBUTE, "pivotRow count")
            .and_then(|value| parse_u32(value, "pivotRow count"))?;
        if count != column_count {
            return Err(invalid("pivotRow count does not equal columnCount"));
        }
        let actual_cells = direct_children(scan, row).count();
        if actual_cells == 0 || actual_cells != usize::try_from(count).unwrap_or(usize::MAX) {
            return Err(invalid("pivotRow count does not equal c child count"));
        }
        let mut row_indexes = Vec::new();
        row_indexes
            .try_reserve_exact(actual_cells)
            .map_err(|source| Error::Allocation {
                resource: "pivotTableData authored cells",
                source,
            })?;
        let mut seen_columns = HashSet::new();
        seen_columns
            .try_reserve(actual_cells)
            .map_err(|source| Error::Allocation {
                resource: "pivotTableData column coordinate index",
                source,
            })?;
        for cell in direct_children(scan, row) {
            if cell.ns != payload.ns || cell.local.as_slice() != CELL_PAYLOAD {
                return Err(invalid("pivotRow has an unexpected direct child"));
            }
            let (record, column_index) = parse_cell(
                scan,
                source,
                cell,
                row_index,
                column_count,
                read_limits,
                retained,
            )?;
            if let Some(index) = column_index
                && !seen_columns.insert(index)
            {
                return Err(invalid("pivotRow has duplicate authored column index"));
            }
            let index = cells.len();
            cells.push(record);
            row_indexes.push(index);
        }
        row_cells.push(row_indexes);
    }
    let has_format_index = cells.iter().any(|record| {
        record
            .state
            .extra
            .as_ref()
            .and_then(|extra| extra.format_index)
            .is_some()
    });
    let server_format_index_unresolved = if has_format_index {
        let parser_bytes = estimate_server_format_parser_bytes(scan)?;
        retained.charge(parser_bytes, "pivotTableData server-format parser")?;
        match parse_payload(scan, source) {
            Ok(payload) => {
                payload.diagnostic_index_boundary
                    || payload.opaque_index_refs
                    || payload.mce_ambiguous
            },
            Err(_) => true,
        }
    } else {
        false
    };
    Ok(ParsedData {
        row_count,
        column_count,
        cache_id,
        owner: payload.start.start..payload.end,
        cells,
        row_cells,
        mce_ambiguous: payload.mce_context,
        server_format_index_unresolved,
        diagnostic,
    })
}

fn validate_data_owner_ext(scan: &XmlScan, owner_ext: &XmlElement) -> Result<bool> {
    let mut valid = true;
    if owner_ext.has_non_whitespace_cdata || owner_ext.has_non_whitespace_text {
        valid = false;
    }
    for attribute in &owner_ext.attrs {
        if !attribute.ns.is_empty() || attribute.local.as_slice() != b"uri" {
            valid = false;
        }
    }
    let mut direct_payloads = 0usize;
    for child in direct_children(scan, owner_ext) {
        if child.ns.as_ref() == EXT_NS && child.local.as_slice() == DATA_PAYLOAD {
            direct_payloads = direct_payloads
                .checked_add(1)
                .ok_or_else(|| invalid("pivotTableData payload count overflows"))?;
        } else if child.ns.as_ref() != MCE_NS || child.local.as_slice() != b"AlternateContent" {
            valid = false;
        }
    }
    if direct_payloads > 1 {
        valid = false;
    }
    Ok(valid)
}

fn validate_standard_item_count(
    scan: &XmlScan,
    root: &XmlElement,
    local: &[u8],
    expected: u32,
) -> Result<()> {
    let mut found = None;
    for element in direct_children(scan, root)
        .filter(|element| element.ns == root.ns && element.local.as_slice() == local)
    {
        if found.is_some() {
            return Err(invalid("PivotTable has duplicate item-count elements"));
        }
        found = Some(element);
    }
    let element =
        found.ok_or_else(|| invalid("PivotTable is missing a standard item-count element"))?;
    let count = required_attr(element, COUNT_ATTRIBUTE, "Pivot item count")
        .and_then(|value| parse_u32(value, "Pivot item count"))?;
    if count != expected {
        return Err(invalid("Pivot item count does not match pivotTableData"));
    }
    Ok(())
}

fn validate_row_attributes(row: &XmlElement) -> Result<()> {
    if row.has_non_whitespace_cdata || row.has_non_whitespace_text {
        return Err(invalid("pivotRow must contain only child elements"));
    }
    for attribute in &row.attrs {
        if !attribute.ns.is_empty()
            || !matches!(
                attribute.local.as_slice(),
                ROW_INDEX_ATTRIBUTE | COUNT_ATTRIBUTE
            )
        {
            return Err(invalid("pivotRow has an unknown attribute"));
        }
    }
    Ok(())
}

/// Return a parent-local preorder window.  `scan_xml_config` assigns element
/// indexes in source preorder and records balanced source ranges, so every
/// descendant of `parent` is contiguous after `parent.index` until the first
/// element whose opening tag starts at or after `parent.end`.  Keeping this
/// window borrowed avoids a per-parent index allocation while preventing the
/// callers below from rescanning unrelated rows/cells in the Part.
fn subtree_elements<'a>(scan: &'a XmlScan, parent: &XmlElement) -> &'a [XmlElement] {
    let Some(first) = parent.index.checked_add(1) else {
        return &[];
    };
    let Some(rest) = scan.elements.get(first..) else {
        return &[];
    };
    let length = rest
        .iter()
        .position(|element| {
            #[cfg(test)]
            LOCAL_SUBTREE_BOUNDARY_VISITS.with(|counter| {
                counter.set(counter.get().saturating_add(1));
            });
            element.start.start >= parent.end
        })
        .unwrap_or(rest.len());
    &rest[..length]
}

fn direct_children<'a>(
    scan: &'a XmlScan,
    parent: &XmlElement,
) -> impl Iterator<Item = &'a XmlElement> + 'a {
    let parent_index = parent.index;
    subtree_elements(scan, parent)
        .iter()
        .filter(move |element| element.parent_index == Some(parent_index))
}

fn parse_cell(
    scan: &XmlScan,
    source: &[u8],
    cell: &XmlElement,
    row_index: Option<u32>,
    column_count: u32,
    read_limits: SourceLimits,
    retained: &mut RetainedBudget,
) -> Result<(CellRecord, Option<u32>)> {
    if cell.has_non_whitespace_cdata || cell.has_non_whitespace_text {
        return Err(invalid("pivot value cell must contain only child elements"));
    }
    for attribute in &cell.attrs {
        if !attribute.ns.is_empty()
            || !matches!(
                attribute.local.as_slice(),
                COLUMN_INDEX_ATTRIBUTE | TYPE_ATTRIBUTE
            )
        {
            return Err(invalid("pivot value cell has an unknown attribute"));
        }
    }
    let column_index = optional_attr(cell, COLUMN_INDEX_ATTRIBUTE)?
        .map(|value| parse_u32(value, "pivot value cell i"))
        .transpose()?;
    if let Some(index) = column_index
        && index >= column_count
    {
        return Err(invalid("pivot value cell i exceeds columnCount"));
    }
    let kind = match optional_attr(cell, TYPE_ATTRIBUTE)?.unwrap_or("n") {
        "b" => PivotCellType::Boolean,
        "n" => PivotCellType::Number,
        "e" => PivotCellType::Error,
        "str" => PivotCellType::Text,
        "d" => PivotCellType::DateTime,
        "bl" => PivotCellType::Blank,
        _ => return Err(invalid("pivot value cell has an invalid t value")),
    };
    let mut value = None;
    let mut extra = None;
    let mut saw_extra = false;
    for child in direct_children(scan, cell) {
        if child.ns != cell.ns {
            return Err(invalid("pivot value cell has a foreign direct child"));
        }
        match child.local.as_slice() {
            VALUE_PAYLOAD => {
                if saw_extra {
                    return Err(invalid("pivot value cell children are out of order"));
                }
                if value.is_some() {
                    return Err(invalid("pivot value cell has duplicate v elements"));
                }
                value = Some(child);
            },
            EXTRA_PAYLOAD => {
                saw_extra = true;
                if extra.is_some() {
                    return Err(invalid("pivot value cell has duplicate x elements"));
                }
                extra = Some(child);
            },
            _ => return Err(invalid("pivot value cell has an unexpected direct child")),
        }
    }
    let value = value.ok_or_else(|| invalid("pivot value cell requires one v element"))?;
    if !value.attrs.is_empty() {
        return Err(invalid("pivot value cell v has an unknown attribute"));
    }
    if value.has_element_child {
        return Err(invalid("pivot value cell v must be a text leaf"));
    }
    if kind == PivotCellType::Text {
        preflight_xstring_utf16(source, value, read_limits)?;
    }
    let value_source_bytes = value_text_source_len(source, value)?;
    // `decode_element_text` first builds a bounded unescaped byte buffer and
    // then the SpreadsheetML-decoded string.  Charge those transient buffers
    // plus a conservative final `Arc<str>` copy before entering that routine;
    // later view/stage clones share the Arc instead of multiplying the text.
    let retained_text = value_source_bytes
        .checked_mul(3)
        .and_then(|length| length.checked_add(size_of::<Arc<str>>()))
        .and_then(|length| length.checked_add(16))
        .ok_or_else(|| invalid("pivotTableData retained text bytes overflow"))?;
    retained.charge(retained_text, "pivotTableData cell text")?;
    let value_text = decode_element_text(source, value, read_limits)?;
    validate_cell_value(kind, &value_text)?;
    let extra_state = extra
        .map(|element| parse_extra(element, read_limits))
        .transpose()?;
    let value_element = source_element_range(source, value)?;
    let qname_len = source_element_qname_len(source, &value.start, MAX_NAME_BYTES)?;
    retained.charge(
        qname_len
            .checked_mul(2)
            .ok_or_else(|| invalid("pivotTableData value QName bytes overflow"))?,
        "pivotTableData value QName",
    )?;
    let value_qname = source_element_qname(source, &value.start, MAX_NAME_BYTES)?;
    let value_start_tag = value_element.start..value.start.end;
    Ok((
        CellRecord {
            source: CellSource {
                value_element,
                value_start_tag,
                value_qname,
                extra: extra.map(parse_extra_source),
            },
            state: CellState {
                row_index,
                column_index,
                kind,
                value_text: Arc::from(value_text),
                extra: extra_state,
            },
        },
        column_index,
    ))
}

fn parse_extra(element: &XmlElement, limits: SourceLimits) -> Result<ExtraState> {
    let mut result = ExtraState {
        format_index: None,
        background_color: None,
        foreground_color: None,
        italic: None,
        underline: None,
        strike: None,
        bold: None,
    };
    for attribute in &element.attrs {
        if !attribute.ns.is_empty() {
            return Err(invalid("pivot value cell extra has a qualified attribute"));
        }
        let slot = match attribute.local.as_slice() {
            FORMAT_INDEX_ATTRIBUTE => &mut result.format_index,
            BACKGROUND_COLOR_ATTRIBUTE => &mut result.background_color,
            FOREGROUND_COLOR_ATTRIBUTE => &mut result.foreground_color,
            ITALIC_ATTRIBUTE | UNDERLINE_ATTRIBUTE | STRIKE_ATTRIBUTE | BOLD_ATTRIBUTE => {
                // Boolean attributes are checked in the dedicated branch below.
                continue;
            },
            _ => return Err(invalid("pivot value cell extra has an unknown attribute")),
        };
        if slot.is_some() {
            return Err(invalid("pivot value cell extra has a duplicate attribute"));
        }
        let value = match attribute.local.as_slice() {
            FORMAT_INDEX_ATTRIBUTE => parse_u32(&attribute.value, "pivot value cell extra in")?,
            BACKGROUND_COLOR_ATTRIBUTE | FOREGROUND_COLOR_ATTRIBUTE => {
                parse_hex_u32(&attribute.value, "pivot value cell extra color")?
            },
            _ => unreachable!(),
        };
        *slot = Some(value);
    }
    for (name, slot) in [
        (ITALIC_ATTRIBUTE, &mut result.italic),
        (UNDERLINE_ATTRIBUTE, &mut result.underline),
        (STRIKE_ATTRIBUTE, &mut result.strike),
        (BOLD_ATTRIBUTE, &mut result.bold),
    ] {
        if let Some(attribute) = unique_attribute(element, name)? {
            let maximum = caller_attribute_limit(limits);
            if attribute.value.len() > maximum {
                return Err(Error::limit(
                    LimitKind::AttributeBytes,
                    attribute.value.len(),
                    maximum,
                ));
            }
            *slot = Some(parse_bool(
                &attribute.value,
                "pivot value cell extra boolean",
            )?);
        }
    }
    if element.has_element_child || element.has_cdata || element.has_text {
        return Err(invalid("pivot value cell extra must be empty"));
    }
    Ok(result)
}

fn parse_extra_source(element: &XmlElement) -> ExtraSource {
    ExtraSource {
        start_tag: element.start.clone(),
        background_color: attr_source_data(element, BACKGROUND_COLOR_ATTRIBUTE),
        foreground_color: attr_source_data(element, FOREGROUND_COLOR_ATTRIBUTE),
        italic: attr_source_data(element, ITALIC_ATTRIBUTE),
        underline: attr_source_data(element, UNDERLINE_ATTRIBUTE),
        strike: attr_source_data(element, STRIKE_ATTRIBUTE),
        bold: attr_source_data(element, BOLD_ATTRIBUTE),
    }
}

fn unique_attribute<'a>(element: &'a XmlElement, name: &[u8]) -> Result<Option<&'a XmlAttribute>> {
    let mut found = None;
    for attribute in &element.attrs {
        if attribute.ns.is_empty() && attribute.local.as_slice() == name {
            if found.is_some() {
                return Err(invalid("pivot value cell extra has a duplicate attribute"));
            }
            found = Some(attribute);
        }
    }
    Ok(found)
}

fn attr_source_data(element: &XmlElement, name: &[u8]) -> Option<AttributeSourceData> {
    unique_attribute(element, name)
        .ok()
        .flatten()
        .map(|attribute| AttributeSourceData {
            value: attribute.value_range.clone(),
            whole: attribute.whole_range.clone(),
        })
}

fn required_attr<'a>(element: &'a XmlElement, name: &[u8], owner: &str) -> Result<&'a str> {
    unique_attribute(element, name)?
        .map(|attribute| attribute.value.as_str())
        .ok_or_else(|| invalid(format!("{owner} requires attribute")))
}

fn optional_attr<'a>(element: &'a XmlElement, name: &[u8]) -> Result<Option<&'a str>> {
    Ok(unique_attribute(element, name)?.map(|attribute| attribute.value.as_str()))
}

fn parse_hex_u32(value: &str, owner: &str) -> Result<u32> {
    let value = value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'));
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(format!(
            "{owner} is not an eight-digit unsigned hexadecimal integer"
        )));
    }
    u32::from_str_radix(value, 16).map_err(|_| {
        invalid(format!(
            "{owner} is not an eight-digit unsigned hexadecimal integer"
        ))
    })
}

fn validate_cell_value(kind: PivotCellType, value: &str) -> Result<()> {
    match kind {
        PivotCellType::Boolean if !matches!(value, "true" | "false") => {
            Err(invalid("pivot boolean cell must be true or false"))
        },
        PivotCellType::Error
            if !matches!(
                value,
                "#DIV/0!" | "#VALUE!" | "#NUM!" | "#N/A" | "#GETTING_DATA"
            ) =>
        {
            Err(invalid("pivot error cell has an invalid error value"))
        },
        PivotCellType::Text
            if value
                .encode_utf16()
                .take(MAX_UTF16_TEXT.saturating_add(1))
                .count()
                > MAX_UTF16_TEXT =>
        {
            Err(Error::limit(
                LimitKind::TextUnits,
                MAX_UTF16_TEXT.saturating_add(1),
                MAX_UTF16_TEXT,
            ))
        },
        PivotCellType::Blank if !value.is_empty() => {
            Err(invalid("pivot blank cell must have an empty v value"))
        },
        _ => Ok(()),
    }
}

/// Count the decoded UTF-16 units of an ST_Xstring without first materializing
/// its decoded value.  This is the admission check for the 65,535-unit text
/// facet; it also keeps a caller-sized source string from becoming a large
/// temporary `String` before the semantic limit is known.
fn preflight_xstring_utf16(
    source: &[u8],
    element: &XmlElement,
    limits: SourceLimits,
) -> Result<()> {
    let start_tag = source
        .get(element.start.clone())
        .ok_or_else(|| invalid("pivot value element start range is invalid"))?;
    if start_tag.ends_with(b"/>") {
        return Ok(());
    }
    let content = source
        .get(element.start.end..element.end)
        .ok_or_else(|| invalid("pivot value element content range is invalid"))?;
    let body_len = text_body_len(content)?;
    let maximum = caller_attribute_limit(limits);
    if body_len > maximum {
        return Err(Error::limit(LimitKind::TextBytes, body_len, maximum));
    }
    let mut counter = XStringCounter::default();
    let mut cursor = 0usize;
    while cursor < content.len() {
        if content[cursor] != b'<' {
            let end = content[cursor..]
                .iter()
                .position(|byte| *byte == b'<')
                .map_or(Ok(content.len()), |offset| {
                    cursor
                        .checked_add(offset)
                        .ok_or_else(|| invalid("pivot value cell text range overflows"))
                })?;
            let segment = std::str::from_utf8(&content[cursor..end])
                .map_err(|error| invalid(error.to_string()))?;
            feed_xml_text(segment, &mut counter)?;
            cursor = end;
            continue;
        }
        if content[cursor..].starts_with(b"<!--") {
            cursor = skip_markup(content, cursor, b"-->", 4, "comment")?;
            continue;
        }
        if content[cursor..].starts_with(b"<?") {
            cursor = skip_markup(content, cursor, b"?>", 2, "processing instruction")?;
            continue;
        }
        if content[cursor..].starts_with(b"<![CDATA[") {
            let inner_start = cursor
                .checked_add(b"<![CDATA[".len())
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            let end_offset = content[inner_start..]
                .windows(3)
                .position(|window| window == b"]]>")
                .ok_or_else(|| invalid("pivot value cell CDATA is unterminated"))?;
            let end = inner_start
                .checked_add(end_offset)
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            let segment = std::str::from_utf8(&content[inner_start..end])
                .map_err(|error| invalid(error.to_string()))?;
            feed_decoded_text(segment, &mut counter)?;
            cursor = end
                .checked_add(3)
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            continue;
        }
        // The scanner has already rejected child elements.  The remaining
        // markup is this value's closing tag.
        break;
    }
    counter.finish()
}

fn skip_markup(
    content: &[u8],
    start: usize,
    terminator: &[u8],
    prefix_len: usize,
    kind: &str,
) -> Result<usize> {
    let content_start = start
        .checked_add(prefix_len)
        .ok_or_else(|| invalid(format!("pivot value cell {kind} range overflows")))?;
    content
        .get(content_start..)
        .ok_or_else(|| invalid(format!("pivot value cell {kind} range is invalid")))?
        .windows(terminator.len())
        .position(|window| window == terminator)
        .and_then(|offset| {
            content_start
                .checked_add(offset)
                .and_then(|value| value.checked_add(terminator.len()))
        })
        .ok_or_else(|| invalid(format!("pivot value cell {kind} is unterminated")))
}

fn text_body_len(content: &[u8]) -> Result<usize> {
    let mut length = 0usize;
    let mut cursor = 0usize;
    while cursor < content.len() {
        if content[cursor] != b'<' {
            let end = content[cursor..]
                .iter()
                .position(|byte| *byte == b'<')
                .map_or(Ok(content.len()), |offset| {
                    cursor
                        .checked_add(offset)
                        .ok_or_else(|| invalid("pivot value cell text range overflows"))
                })?;
            length = length
                .checked_add(end - cursor)
                .ok_or_else(|| invalid("pivot value cell text length overflows"))?;
            cursor = end;
            continue;
        }
        if content[cursor..].starts_with(b"<!--") {
            cursor = skip_markup(content, cursor, b"-->", 4, "comment")?;
            continue;
        }
        if content[cursor..].starts_with(b"<?") {
            cursor = skip_markup(content, cursor, b"?>", 2, "processing instruction")?;
            continue;
        }
        if content[cursor..].starts_with(b"<![CDATA[") {
            let inner_start = cursor
                .checked_add(b"<![CDATA[".len())
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            let end_offset = content[inner_start..]
                .windows(3)
                .position(|window| window == b"]]>")
                .ok_or_else(|| invalid("pivot value cell CDATA is unterminated"))?;
            let end = inner_start
                .checked_add(end_offset)
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            length = length
                .checked_add(end - inner_start)
                .ok_or_else(|| invalid("pivot value cell text length overflows"))?;
            cursor = end
                .checked_add(3)
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            continue;
        }
        // The scanner has already rejected child elements.  The remaining
        // markup is this value's closing tag, which is not text content.
        break;
    }
    Ok(length)
}

fn feed_xml_text(segment: &str, counter: &mut XStringCounter) -> Result<()> {
    let bytes = segment.as_bytes();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        let Some(relative) = bytes[cursor..].iter().position(|byte| *byte == b'&') else {
            feed_decoded_text(&segment[cursor..], counter)?;
            break;
        };
        let start = cursor + relative;
        feed_decoded_text(&segment[cursor..start], counter)?;
        let end = bytes[start..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|offset| start + offset)
            .ok_or_else(|| invalid("pivot value cell XML entity is unterminated"))?;
        let entity = &segment[start + 1..end];
        let character = decode_xml_entity_char(entity)?;
        counter.push_char(character)?;
        cursor = end
            .checked_add(1)
            .ok_or_else(|| invalid("pivot value cell XML entity range overflows"))?;
    }
    Ok(())
}

fn decode_xml_entity_char(entity: &str) -> Result<char> {
    let value = if let Some(hex) = entity
        .strip_prefix("#x")
        .or_else(|| entity.strip_prefix("#X"))
    {
        u32::from_str_radix(hex, 16)
    } else if let Some(decimal) = entity.strip_prefix('#') {
        decimal.parse::<u32>()
    } else {
        return match entity {
            "amp" => Ok('&'),
            "lt" => Ok('<'),
            "gt" => Ok('>'),
            "quot" => Ok('"'),
            "apos" => Ok('\''),
            _ => Err(invalid("pivot value cell has an unknown XML entity")),
        };
    };
    let value =
        value.map_err(|_| invalid("pivot value cell has an invalid XML character reference"))?;
    char::from_u32(value)
        .ok_or_else(|| invalid("pivot value cell has an invalid XML character reference"))
}

fn feed_decoded_text(segment: &str, counter: &mut XStringCounter) -> Result<()> {
    for character in segment.chars() {
        counter.push_char(character)?;
    }
    Ok(())
}

#[derive(Default)]
struct XStringCounter {
    units: usize,
    pending: [u8; 7],
    pending_len: usize,
    high_surrogate: bool,
}

impl XStringCounter {
    fn push_char(&mut self, character: char) -> Result<()> {
        if self.pending_len != 0 {
            if !character.is_ascii() {
                self.flush_pending()?;
            } else {
                self.pending[self.pending_len] = character as u8;
                self.pending_len += 1;
                if self.pending_len == self.pending.len() {
                    if let Some((unit, _)) = raw::strings::spreadsheet_escape_at(&self.pending, 0) {
                        self.pending_len = 0;
                        self.push_escape(unit)?;
                    } else {
                        let first = self.pending[0] as char;
                        self.pending.copy_within(1.., 0);
                        self.pending_len -= 1;
                        self.push_literal(first)?;
                    }
                }
                return Ok(());
            }
        }
        if character == '_' {
            self.pending[0] = b'_';
            self.pending_len = 1;
        } else {
            self.push_literal(character)?;
        }
        Ok(())
    }

    fn flush_pending(&mut self) -> Result<()> {
        if self.high_surrogate {
            return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
        }
        let pending = self.pending;
        let length = self.pending_len;
        self.pending_len = 0;
        for byte in pending.into_iter().take(length) {
            self.push_literal(byte as char)?;
        }
        Ok(())
    }

    fn push_literal(&mut self, character: char) -> Result<()> {
        if self.high_surrogate {
            return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
        }
        self.units = self
            .units
            .checked_add(character.encode_utf16(&mut [0; 2]).len())
            .ok_or_else(|| invalid("pivot string UTF-16 length overflows"))?;
        if self.units > MAX_UTF16_TEXT {
            return Err(Error::limit(
                LimitKind::TextUnits,
                self.units,
                MAX_UTF16_TEXT,
            ));
        }
        Ok(())
    }

    fn push_escape(&mut self, unit: u16) -> Result<()> {
        if self.high_surrogate {
            if !(0xDC00..=0xDFFF).contains(&unit) {
                return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
            }
            self.high_surrogate = false;
            self.units = self
                .units
                .checked_add(1)
                .ok_or_else(|| invalid("pivot string UTF-16 length overflows"))?;
        } else if (0xD800..=0xDBFF).contains(&unit) {
            self.high_surrogate = true;
            self.units = self
                .units
                .checked_add(1)
                .ok_or_else(|| invalid("pivot string UTF-16 length overflows"))?;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid("unpaired low surrogate in SpreadsheetML escape"));
        } else {
            self.units = self
                .units
                .checked_add(1)
                .ok_or_else(|| invalid("pivot string UTF-16 length overflows"))?;
        }
        if self.units > MAX_UTF16_TEXT {
            return Err(Error::limit(
                LimitKind::TextUnits,
                self.units,
                MAX_UTF16_TEXT,
            ));
        }
        Ok(())
    }

    fn finish(mut self) -> Result<()> {
        self.flush_pending()?;
        if self.high_surrogate {
            return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
        }
        Ok(())
    }
}

fn decode_element_text(
    source: &[u8],
    element: &XmlElement,
    limits: SourceLimits,
) -> Result<String> {
    let start_tag = source
        .get(element.start.clone())
        .ok_or_else(|| invalid("pivot value element start range is invalid"))?;
    if start_tag.ends_with(b"/>") {
        return Ok(String::new());
    }
    let content_start = element.start.end;
    let content = source
        .get(content_start..element.end)
        .ok_or_else(|| invalid("pivot value element content range is invalid"))?;
    let maximum = caller_attribute_limit(limits);
    let body_len = text_body_len(content)?;
    if body_len > maximum {
        return Err(Error::limit(LimitKind::TextBytes, body_len, maximum));
    }
    let mut text = Vec::new();
    text.try_reserve_exact(body_len)
        .map_err(|source| Error::Allocation {
            resource: "pivot value cell text",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < content.len() {
        if content[cursor] != b'<' {
            let end = content[cursor..]
                .iter()
                .position(|byte| *byte == b'<')
                .map_or(Ok(content.len()), |offset| {
                    cursor
                        .checked_add(offset)
                        .ok_or_else(|| invalid("pivot value cell text range overflows"))
                })?;
            let segment = std::str::from_utf8(&content[cursor..end])
                .map_err(|error| invalid(error.to_string()))?;
            let decoded =
                quick_xml::escape::unescape(segment).map_err(|error| invalid(error.to_string()))?;
            text.extend_from_slice(decoded.as_bytes());
            cursor = end;
            continue;
        }
        if content[cursor..].starts_with(b"<!--") {
            let comment_start = cursor
                .checked_add(4)
                .ok_or_else(|| invalid("pivot value cell comment range overflows"))?;
            let end = content[comment_start..]
                .windows(3)
                .position(|window| window == b"-->")
                .and_then(|offset| {
                    comment_start
                        .checked_add(offset)
                        .and_then(|value| value.checked_add(3))
                })
                .ok_or_else(|| invalid("pivot value cell comment is unterminated"))?;
            cursor = end;
            continue;
        }
        if content[cursor..].starts_with(b"<![CDATA[") {
            let inner_start = cursor
                .checked_add(b"<![CDATA[".len())
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            let end = content[inner_start..]
                .windows(3)
                .position(|window| window == b"]]>")
                .and_then(|offset| inner_start.checked_add(offset))
                .ok_or_else(|| invalid("pivot value cell CDATA is unterminated"))?;
            let cdata = std::str::from_utf8(&content[inner_start..end])
                .map_err(|error| invalid(error.to_string()))?;
            text.extend_from_slice(cdata.as_bytes());
            cursor = end
                .checked_add(3)
                .ok_or_else(|| invalid("pivot value cell CDATA range overflows"))?;
            continue;
        }
        if content[cursor..].starts_with(b"<?") {
            let instruction_start = cursor.checked_add(2).ok_or_else(|| {
                invalid("pivot value cell processing instruction range overflows")
            })?;
            let end = content[instruction_start..]
                .windows(2)
                .position(|window| window == b"?>")
                .and_then(|offset| {
                    instruction_start
                        .checked_add(offset)
                        .and_then(|value| value.checked_add(2))
                })
                .ok_or_else(|| {
                    invalid("pivot value cell processing instruction is unterminated")
                })?;
            cursor = end;
            continue;
        }
        // The scanner has already rejected child elements.  The remaining
        // markup is this value's closing tag, which is outside its text.
        break;
    }
    let raw = std::str::from_utf8(&text).map_err(|error| invalid(error.to_string()))?;
    raw::strings::decode_spreadsheet_text(raw)
}

fn value_text_source_len(source: &[u8], element: &XmlElement) -> Result<usize> {
    let start_tag = source
        .get(element.start.clone())
        .ok_or_else(|| invalid("pivot value element start range is invalid"))?;
    if start_tag.ends_with(b"/>") {
        return Ok(0);
    }
    let content = source
        .get(element.start.end..element.end)
        .ok_or_else(|| invalid("pivot value element content range is invalid"))?;
    text_body_len(content)
}

fn source_element_qname(source: &[u8], start: &Range<usize>, maximum: usize) -> Result<Vec<u8>> {
    let end = source_element_qname_len(source, start, maximum)?;
    let bytes = source
        .get(start.clone())
        .ok_or_else(|| invalid("pivot element QName range is invalid"))?;
    let start_offset = if bytes.first() == Some(&b'<') { 1 } else { 0 };
    let mut qname = Vec::new();
    qname
        .try_reserve_exact(end)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData element QName",
            source,
        })?;
    qname.extend_from_slice(&bytes[start_offset..start_offset + end]);
    Ok(qname)
}

fn source_element_qname_len(source: &[u8], start: &Range<usize>, maximum: usize) -> Result<usize> {
    let bytes = source
        .get(start.clone())
        .ok_or_else(|| invalid("pivot element QName range is invalid"))?;
    let start_offset = if bytes.first() == Some(&b'<') { 1 } else { 0 };
    let end = bytes[start_offset..]
        .iter()
        .position(|byte| is_xml_space(*byte) || matches!(byte, b'/' | b'>'))
        .ok_or_else(|| invalid("pivot element start tag has no QName"))?;
    if end == 0 {
        return Err(invalid("pivot element start tag has an empty QName"));
    }
    if end > maximum {
        return Err(Error::limit(LimitKind::NameBytes, end, maximum));
    }
    Ok(end)
}

fn source_element_range(source: &[u8], element: &XmlElement) -> Result<Range<usize>> {
    let start = if source.get(element.start.start) == Some(&b'<') {
        element.start.start
    } else {
        element
            .start
            .start
            .checked_sub(1)
            .filter(|index| source.get(*index) == Some(&b'<'))
            .ok_or_else(|| invalid("pivot element source range has no opening delimiter"))?
    };
    if start >= element.end || element.end > source.len() {
        return Err(invalid("pivot element source range is invalid"));
    }
    Ok(start..element.end)
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// A scalar transaction over existing authored C444 cells.
#[cfg(test)]
pub(crate) struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    staged: Vec<CellState>,
    positions: HashMap<PivotCellAddress, CellPosition>,
    staged_text_ledger: Vec<usize>,
    staged_text_bytes: usize,
    selection: usize,
}

#[cfg(test)]
impl<'a> Transaction<'a> {
    pub(crate) fn new<'selector>(
        target: &'a mut OpcPackage,
        selector: impl Into<PivotTableSelector<'selector>>,
    ) -> Result<Self> {
        Self::with_limits(target, selector, &Limits::default())
    }

    /// Start a low-level transaction with an owner-specific semantic policy.
    pub(crate) fn with_limits<'selector>(
        target: &'a mut OpcPackage,
        selector: impl Into<PivotTableSelector<'selector>>,
        limits: &Limits,
    ) -> Result<Self> {
        let before = Snapshot::load_with_limits(target, selector, limits)?;
        ensure_editable(&before)?;
        let staging_bytes = staging_retained_bytes(before.cells.len())?;
        let retained_with_staging = before
            .retained_bytes
            .checked_add(staging_bytes)
            .ok_or_else(|| invalid("pivotTableData staged bytes overflow"))?;
        enforce_retained_limit(retained_with_staging, limits.max_retained_bytes)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.cells.len())
            .map_err(|source| Error::Allocation {
                resource: "pivotTableData staged cells",
                source,
            })?;
        for record in &before.cells {
            staged.push(record.state.clone());
        }
        let positions = build_cell_positions(&before)?;
        let mut staged_text_ledger = Vec::new();
        staged_text_ledger
            .try_reserve_exact(before.cells.len())
            .map_err(|source| Error::Allocation {
                resource: "pivotTableData staged text ledger",
                source,
            })?;
        staged_text_ledger.extend(std::iter::repeat_n(0, before.cells.len()));
        Ok(Self {
            selection: before.selection,
            target,
            before,
            staged,
            positions,
            staged_text_ledger,
            staged_text_bytes: 0,
        })
    }

    pub(crate) fn set_value(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: CellChange,
    ) -> Result<bool> {
        stage_set_value(
            &mut self.staged,
            &self.before,
            &self.positions,
            &mut self.staged_text_ledger,
            &mut self.staged_text_bytes,
            address.into(),
            value,
            self.target.read_limits(),
        )
    }

    /// Replace an existing text cell after validating the borrowed value.
    ///
    /// The value is checked before any staged `String` is allocated.  This is
    /// the preferred path for callers that do not already own the replacement
    /// text.
    pub(crate) fn set_text(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        stage_set_borrowed_value(
            &mut self.staged,
            &self.before,
            &self.positions,
            &mut self.staged_text_ledger,
            &mut self.staged_text_bytes,
            address.into(),
            PivotCellType::Text,
            value.as_ref(),
            self.target.read_limits(),
        )
    }

    /// Replace an existing error cell after validating the borrowed value.
    ///
    /// The value is checked before any staged `String` is allocated.  This is
    /// the preferred path for callers that do not already own the replacement
    /// error text.
    pub(crate) fn set_error(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        stage_set_borrowed_value(
            &mut self.staged,
            &self.before,
            &self.positions,
            &mut self.staged_text_ledger,
            &mut self.staged_text_bytes,
            address.into(),
            PivotCellType::Error,
            value.as_ref(),
            self.target.read_limits(),
        )
    }

    pub(crate) fn set_extra(
        &mut self,
        address: impl Into<PivotCellAddress>,
        edit: CellExtra,
    ) -> Result<bool> {
        stage_set_extra(
            &mut self.staged,
            &self.positions,
            address.into(),
            edit,
            self.target.read_limits(),
        )
    }

    #[must_use]
    pub(crate) fn is_changed(&self) -> bool {
        self.before
            .cells
            .iter()
            .zip(&self.staged)
            .any(|(before, after)| before.state != *after)
    }

    pub(crate) fn commit(self) -> Result<SourceCommit> {
        if !self.is_changed() {
            let patch = SourcePatch::new(self.before.clone(), self.before.clone());
            return Ok(SourceCommit::new(self.before, patch, false));
        }
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load_with_limits(
            self.target,
            PivotTableSelector::Position(self.selection),
            &self.before.limits,
        )?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.table_part().to_string(),
            });
        }
        let output = rewrite_data(
            &self.before,
            &self.staged,
            caller_part_limit(self.target.read_limits()),
            self.target.read_limits().max_total_part_bytes(),
            caller_attribute_limit(self.target.read_limits()),
            package_total_bytes(self.target)?,
        )?;
        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(self.before.table_part())?
            .set_blob_shared(Arc::new(output));
        validate_candidate_limits(&candidate)?;
        let after = Snapshot::load_with_limits(
            &candidate,
            PivotTableSelector::Position(self.selection),
            &self.before.limits,
        )?;
        if !after.same_readset(&self.before) || !states_equal(&after, &self.staged) {
            return Err(invalid(
                "pivotTableData publication changed semantic state or read set",
            ));
        }
        let patch = SourcePatch::new(self.before, after.clone());
        *self.target = candidate;
        Ok(SourceCommit::new(after, patch, true))
    }
}

fn ensure_editable(snapshot: &Snapshot) -> Result<()> {
    if !snapshot.value.is_editable() {
        return Err(invalid(
            "pivotTableData source is diagnostic/read-only because its owner closure is ambiguous",
        ));
    }
    Ok(())
}

fn states_equal(snapshot: &Snapshot, staged: &[CellState]) -> bool {
    snapshot.cells.len() == staged.len()
        && snapshot
            .cells
            .iter()
            .zip(staged)
            .all(|(record, state)| record.state == *state)
}

fn validate_edited_text(value: &str, limits: SourceLimits) -> Result<()> {
    let maximum = caller_attribute_limit(limits);
    if value.len() > maximum {
        return Err(Error::limit(LimitKind::TextBytes, value.len(), maximum));
    }
    let units = value
        .encode_utf16()
        .take(MAX_UTF16_TEXT.saturating_add(1))
        .count();
    if units > MAX_UTF16_TEXT {
        return Err(Error::limit(LimitKind::TextUnits, units, MAX_UTF16_TEXT));
    }
    // `ST_Xstring` serializes XML-illegal controls as SpreadsheetML
    // `_xHHHH_` escapes.  Literal controls are still rejected on source
    // ingress; edited values are checked by the bounded encoder at commit.
    Ok(())
}

fn apply_extra_edit(extra: &mut ExtraState, edit: CellExtra, limits: SourceLimits) -> Result<()> {
    apply_hex_edit(&mut extra.background_color, edit.background_color, limits)?;
    apply_hex_edit(&mut extra.foreground_color, edit.foreground_color, limits)?;
    apply_bool_edit(&mut extra.italic, edit.italic)?;
    apply_bool_edit(&mut extra.underline, edit.underline)?;
    apply_bool_edit(&mut extra.strike, edit.strike)?;
    apply_bool_edit(&mut extra.bold, edit.bold)?;
    Ok(())
}

fn apply_hex_edit(
    slot: &mut Option<u32>,
    edit: Attribute<String>,
    limits: SourceLimits,
) -> Result<()> {
    match edit {
        Attribute::Keep => {},
        Attribute::Clear => *slot = None,
        Attribute::Set(value) => {
            let maximum = caller_attribute_limit(limits);
            if value.len() > maximum {
                return Err(Error::limit(
                    LimitKind::AttributeBytes,
                    value.len(),
                    maximum,
                ));
            }
            *slot = Some(parse_hex_u32(&value, "pivotTableData color")?);
        },
    }
    Ok(())
}

fn apply_bool_edit(slot: &mut Option<bool>, edit: Attribute<bool>) -> Result<()> {
    match edit {
        Attribute::Keep => {},
        Attribute::Clear => *slot = None,
        Attribute::Set(value) => *slot = Some(value),
    }
    Ok(())
}

fn staged_cell_position(
    positions: &HashMap<PivotCellAddress, CellPosition>,
    address: PivotCellAddress,
) -> Result<usize> {
    match positions.get(&address) {
        Some(CellPosition::Unique(index)) => Ok(*index),
        Some(CellPosition::Ambiguous) => Err(invalid("pivotTableData cell address is ambiguous")),
        None => Err(invalid("pivotTableData cell address did not resolve")),
    }
}

fn stage_set_value(
    staged: &mut [CellState],
    before: &Snapshot,
    positions: &HashMap<PivotCellAddress, CellPosition>,
    staged_text_ledger: &mut [usize],
    staged_text_bytes: &mut usize,
    address: PivotCellAddress,
    value: CellChange,
    limits: SourceLimits,
) -> Result<bool> {
    let position = staged_cell_position(positions, address)?;
    let kind = staged
        .get(position)
        .map(|slot| slot.kind)
        .ok_or_else(|| invalid("pivotTableData cell address did not resolve"))?;
    let (replacement, replacement_bytes, transient_bytes) = match (kind, &value) {
        (PivotCellType::Boolean, CellChange::Boolean(value)) => {
            let replacement = if *value { "true" } else { "false" };
            (
                replacement,
                replacement_text_retained_bytes(replacement.len())?,
                0,
            )
        },
        (PivotCellType::Text, CellChange::Text(value)) => {
            validate_edited_text(value, limits)?;
            (
                value.as_str(),
                replacement_text_retained_bytes(value.len())?,
                value.capacity(),
            )
        },
        (PivotCellType::Error, CellChange::Error(value)) => {
            validate_cell_value(PivotCellType::Error, value)?;
            validate_edited_text(value, limits)?;
            (
                value.as_str(),
                replacement_text_retained_bytes(value.len())?,
                value.capacity(),
            )
        },
        (PivotCellType::Blank, CellChange::Blank) => ("", replacement_text_retained_bytes(0)?, 0),
        (PivotCellType::Number | PivotCellType::DateTime, _) => {
            return Err(Error::Unsupported {
                feature: "pivotTableData numeric/date-time scalar editing",
            });
        },
        (kind, _) => {
            return Err(invalid(format!(
                "pivotTableData value edit does not match cell kind {kind:?}"
            )));
        },
    };
    if staged[position].value_text.as_ref() == replacement {
        return Ok(false);
    }
    let previous_bytes = staged_text_ledger
        .get(position)
        .copied()
        .ok_or_else(|| invalid("pivotTableData staged text ledger is incomplete"))?;
    let updated_bytes = staged_text_bytes
        .checked_sub(previous_bytes)
        .and_then(|bytes| bytes.checked_add(replacement_bytes))
        .ok_or_else(|| invalid("pivotTableData staged text bytes overflow"))?;
    // The old staged Arc remains live until the replacement Arc has been
    // constructed.  Charge that per-position peak separately; it is removed
    // from the running ledger as soon as the assignment succeeds.
    let transient_peak_bytes = transient_bytes
        .checked_add(previous_bytes)
        .ok_or_else(|| invalid("pivotTableData staged transient bytes overflow"))?;
    admit_staged_text(before, staged.len(), updated_bytes, transient_peak_bytes)?;
    let next: Arc<str> = match value {
        CellChange::Boolean(value) => Arc::from(if value { "true" } else { "false" }),
        CellChange::Text(value) | CellChange::Error(value) => Arc::from(value),
        CellChange::Blank => Arc::from(""),
    };
    staged[position].value_text = next;
    staged_text_ledger[position] = replacement_bytes;
    *staged_text_bytes = updated_bytes;
    Ok(true)
}

fn stage_set_borrowed_value(
    staged: &mut [CellState],
    before: &Snapshot,
    positions: &HashMap<PivotCellAddress, CellPosition>,
    staged_text_ledger: &mut [usize],
    staged_text_bytes: &mut usize,
    address: PivotCellAddress,
    expected_kind: PivotCellType,
    value: &str,
    limits: SourceLimits,
) -> Result<bool> {
    let position = staged_cell_position(positions, address)?;
    let kind = staged
        .get(position)
        .map(|slot| slot.kind)
        .ok_or_else(|| invalid("pivotTableData cell address did not resolve"))?;
    if kind != expected_kind {
        return Err(invalid(format!(
            "pivotTableData value edit does not match cell kind {:?}",
            kind
        )));
    }
    if expected_kind == PivotCellType::Error {
        validate_cell_value(PivotCellType::Error, value)?;
    }
    validate_edited_text(value, limits)?;
    if staged[position].value_text.as_ref() == value {
        return Ok(false);
    }
    let replacement_bytes = replacement_text_retained_bytes(value.len())?;
    let previous_bytes = staged_text_ledger
        .get(position)
        .copied()
        .ok_or_else(|| invalid("pivotTableData staged text ledger is incomplete"))?;
    let updated_bytes = staged_text_bytes
        .checked_sub(previous_bytes)
        .and_then(|bytes| bytes.checked_add(replacement_bytes))
        .ok_or_else(|| invalid("pivotTableData staged text bytes overflow"))?;
    let transient_peak_bytes = value
        .len()
        .checked_add(previous_bytes)
        .ok_or_else(|| invalid("pivotTableData staged transient bytes overflow"))?;
    admit_staged_text(before, staged.len(), updated_bytes, transient_peak_bytes)?;
    let mut replacement = String::new();
    replacement
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData borrowed scalar value",
            source,
        })?;
    let actual_capacity = replacement.capacity();
    if actual_capacity > value.len() {
        let transient_peak_bytes = actual_capacity
            .checked_add(previous_bytes)
            .ok_or_else(|| invalid("pivotTableData staged transient bytes overflow"))?;
        admit_staged_text(before, staged.len(), updated_bytes, transient_peak_bytes)?;
    }
    replacement.push_str(value);
    staged[position].value_text = Arc::from(replacement);
    staged_text_ledger[position] = replacement_bytes;
    *staged_text_bytes = updated_bytes;
    Ok(true)
}

fn stage_set_extra(
    staged: &mut [CellState],
    positions: &HashMap<PivotCellAddress, CellPosition>,
    address: PivotCellAddress,
    edit: CellExtra,
    limits: SourceLimits,
) -> Result<bool> {
    let position = staged_cell_position(positions, address)?;
    let slot = staged
        .get_mut(position)
        .ok_or_else(|| invalid("pivotTableData cell address did not resolve"))?;
    let extra = slot
        .extra
        .as_mut()
        .ok_or_else(|| invalid("pivotTableData cannot create a missing x element"))?;
    let mut candidate = extra.clone();
    apply_extra_edit(&mut candidate, edit, limits)?;
    if *extra == candidate {
        return Ok(false);
    }
    *extra = candidate;
    Ok(true)
}

/// Exact reversible source patch for one C444 owner.
#[derive(Clone, Debug)]
pub(crate) struct SourcePatch {
    before: Snapshot,
    after: Snapshot,
}

impl SourcePatch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.table.bytes == self.after.table.bytes
    }

    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load_with_limits(
            package,
            PivotTableSelector::Position(self.before.selection),
            &self.before.limits,
        )?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.table_part().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.before.table_part())?
            .set_blob_shared(Arc::clone(&self.after.table.bytes));
        validate_candidate_limits(&candidate)?;
        let resulting = Snapshot::load_with_limits(
            &candidate,
            PivotTableSelector::Position(self.after.selection),
            &self.after.limits,
        )?;
        if !resulting.same_source(&self.after) || resulting.data() != self.after.data() {
            return Err(invalid("pivotTableData patch verification failed"));
        }
        *package = candidate;
        Ok(())
    }
}

#[derive(Clone, Debug)]
#[cfg(test)]
pub(crate) struct SourceCommit {
    snapshot: Snapshot,
    patch: SourcePatch,
    changed: bool,
}

#[cfg(test)]
impl SourceCommit {
    fn new(snapshot: Snapshot, patch: SourcePatch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub(crate) fn patch(&self) -> &SourcePatch {
        &self.patch
    }

    #[must_use]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }
}

fn package_total_bytes(package: &OpcPackage) -> Result<u64> {
    // The aggregate Part ceiling is charged against inflated bytes, so this
    // census decodes every payload (ADR 0030).
    package.try_iter_parts().try_fold(0u64, |total, part| {
        total
            .checked_add(u64::try_from(part?.blob().len()).unwrap_or(u64::MAX))
            .ok_or_else(|| invalid("PivotTable candidate Part bytes overflow"))
    })
}

fn rewrite_data(
    before: &Snapshot,
    staged: &[CellState],
    maximum_output_bytes: usize,
    maximum_total_bytes: u64,
    maximum_attribute_bytes: usize,
    current_total_bytes: u64,
) -> Result<Vec<u8>> {
    if staged.len() != before.cells.len() {
        return Err(invalid("pivotTableData source cell count changed"));
    }
    let source = before.table.bytes.as_slice();
    let owner = source
        .get(before.extension_owner.clone())
        .ok_or_else(|| invalid("pivotTableData owner range is invalid"))?;
    if owner.len() > MAX_FRAGMENT_BYTES {
        return Err(Error::limit(
            LimitKind::FragmentBytes,
            owner.len(),
            MAX_FRAGMENT_BYTES,
        ));
    }
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut output_len = source.len();
    let mut owner_len = owner.len();
    for (record, next) in before.cells.iter().zip(staged) {
        let previous = &record.state;
        if previous.value_text != next.value_text {
            if next.value_text.len() > maximum_attribute_bytes {
                return Err(Error::limit(
                    LimitKind::TextBytes,
                    next.value_text.len(),
                    maximum_attribute_bytes,
                ));
            }
            let units = next
                .value_text
                .encode_utf16()
                .take(MAX_UTF16_TEXT.saturating_add(1))
                .count();
            if units > MAX_UTF16_TEXT {
                return Err(Error::limit(LimitKind::TextUnits, units, MAX_UTF16_TEXT));
            }
            let replacement_len =
                encoded_value_element_len(source, &record.source, &next.value_text)?;
            apply_length_delta(
                &record.source.value_element,
                replacement_len,
                before.extension_owner.clone(),
                &mut output_len,
                &mut owner_len,
            )?;
        }
        if let Some(extra_source) = &record.source.extra {
            plan_extra_lengths(
                source,
                extra_source,
                previous.extra.as_ref(),
                next.extra.as_ref(),
                maximum_attribute_bytes,
                before.extension_owner.clone(),
                &mut output_len,
                &mut owner_len,
            )?;
        }
    }
    if owner_len > MAX_FRAGMENT_BYTES
        || output_len > MAX_PART_BYTES
        || output_len > maximum_output_bytes
    {
        return Err(invalid(
            "pivotTableData output exceeds the caller's Part limit",
        ));
    }
    let source_len = u64::try_from(source.len()).unwrap_or(u64::MAX);
    let output_u64 = u64::try_from(output_len).unwrap_or(u64::MAX);
    let total = current_total_bytes
        .checked_sub(source_len)
        .and_then(|value| value.checked_add(output_u64))
        .ok_or_else(|| invalid("pivotTableData aggregate output length overflows"))?;
    if total > maximum_total_bytes {
        return Err(invalid(
            "pivotTableData output exceeds aggregate Part limit",
        ));
    }
    let edit_capacity = before
        .cells
        .len()
        .checked_mul(3)
        .ok_or_else(|| invalid("pivotTableData source edit count overflows"))?;
    edits
        .try_reserve(edit_capacity)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData source edit list",
            source,
        })?;
    for (record, next) in before.cells.iter().zip(staged) {
        let previous = &record.state;
        if previous.value_text != next.value_text {
            edits.push((
                record.source.value_element.clone(),
                encode_value_element(source, &record.source, &next.value_text)?,
            ));
        }
        if let Some(extra_source) = &record.source.extra {
            append_extra_edits(
                source,
                extra_source,
                previous.extra.as_ref(),
                next.extra.as_ref(),
                &mut edits,
                maximum_attribute_bytes,
            )?;
        }
    }
    edits.sort_by_key(|edit| edit.0.start);
    for pair in edits.windows(2) {
        if pair[0].0.end > pair[1].0.start {
            return Err(invalid("pivotTableData source edit ranges overlap"));
        }
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData output",
            source,
        })?;
    let mut cursor = 0usize;
    for (range, replacement) in &edits {
        output.extend_from_slice(
            source
                .get(cursor..range.start)
                .ok_or_else(|| invalid("pivotTableData source edit range is invalid"))?,
        );
        output.extend_from_slice(replacement);
        cursor = range.end;
    }
    output.extend_from_slice(
        source
            .get(cursor..)
            .ok_or_else(|| invalid("pivotTableData source tail range is invalid"))?,
    );
    if output.len() != output_len {
        return Err(invalid("pivotTableData output preflight length disagrees"));
    }
    Ok(output)
}

fn encoded_value_element_len(
    source_bytes: &[u8],
    source: &CellSource,
    value: &str,
) -> Result<usize> {
    let text = escaped_xstring_len(value)?;
    let start_tag = source_bytes
        .get(source.value_start_tag.clone())
        .ok_or_else(|| invalid("pivotTableData value start range is invalid"))?;
    let self_closing = start_tag.ends_with(b"/>");
    let start_len = start_tag
        .len()
        .checked_sub(usize::from(self_closing))
        .ok_or_else(|| invalid("pivotTableData value start-tag length underflows"))?;
    let end_len = source
        .value_qname
        .len()
        .checked_add(3)
        .ok_or_else(|| invalid("pivotTableData value QName length overflows"))?;
    start_len
        .checked_add(text)
        .and_then(|length| length.checked_add(end_len))
        .ok_or_else(|| invalid("pivotTableData value element length overflows"))
}

fn encode_value_element(source_bytes: &[u8], source: &CellSource, value: &str) -> Result<Vec<u8>> {
    let text = try_escaped_xstring(value)?;
    let start_tag = source_bytes
        .get(source.value_start_tag.clone())
        .ok_or_else(|| invalid("pivotTableData value start range is invalid"))?;
    let self_closing = start_tag.ends_with(b"/>");
    let start_len = start_tag
        .len()
        .checked_sub(usize::from(self_closing))
        .ok_or_else(|| invalid("pivotTableData value start-tag length underflows"))?;
    let end_len = source
        .value_qname
        .len()
        .checked_add(3)
        .ok_or_else(|| invalid("pivotTableData value QName length overflows"))?;
    let length = start_len
        .checked_add(text.len())
        .and_then(|length| length.checked_add(end_len))
        .ok_or_else(|| invalid("pivotTableData value element length overflows"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "pivotTableData value element",
            source,
        })?;
    if self_closing {
        output.extend_from_slice(&start_tag[..start_tag.len() - 2]);
        output.push(b'>');
    } else {
        output.extend_from_slice(start_tag);
    }
    output.extend_from_slice(&text);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&source.value_qname);
    output.push(b'>');
    Ok(output)
}

fn apply_length_delta(
    range: &Range<usize>,
    replacement_len: usize,
    owner: Range<usize>,
    output_len: &mut usize,
    owner_len: &mut usize,
) -> Result<()> {
    if range.start < owner.start || range.end > owner.end || range.start > range.end {
        return Err(invalid("pivotTableData edit range is outside its owner"));
    }
    let old_len = range.end.saturating_sub(range.start);
    *output_len = output_len
        .checked_sub(old_len)
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| invalid("pivotTableData output length overflows"))?;
    *owner_len = owner_len
        .checked_sub(old_len)
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| invalid("pivotTableData owner length overflows"))?;
    Ok(())
}

fn plan_extra_lengths(
    source: &[u8],
    extra_source: &ExtraSource,
    previous: Option<&ExtraState>,
    next: Option<&ExtraState>,
    maximum_attribute_bytes: usize,
    owner: Range<usize>,
    output_len: &mut usize,
    owner_len: &mut usize,
) -> Result<()> {
    let (Some(previous), Some(next)) = (previous, next) else {
        return Err(invalid("pivotTableData extra source state is incomplete"));
    };
    for (old, new, attr, name) in [
        (
            previous.background_color.map(PlannedAttributeValue::Hex),
            next.background_color.map(PlannedAttributeValue::Hex),
            extra_source.background_color.as_ref(),
            BACKGROUND_COLOR_ATTRIBUTE,
        ),
        (
            previous.foreground_color.map(PlannedAttributeValue::Hex),
            next.foreground_color.map(PlannedAttributeValue::Hex),
            extra_source.foreground_color.as_ref(),
            FOREGROUND_COLOR_ATTRIBUTE,
        ),
    ] {
        plan_typed_attribute_length(
            source,
            attr,
            old,
            new,
            name,
            maximum_attribute_bytes,
            &extra_source.start_tag,
            owner.clone(),
            output_len,
            owner_len,
        )?;
    }
    for (old, new, attr, name) in [
        (
            previous.italic,
            next.italic,
            extra_source.italic.as_ref(),
            ITALIC_ATTRIBUTE,
        ),
        (
            previous.underline,
            next.underline,
            extra_source.underline.as_ref(),
            UNDERLINE_ATTRIBUTE,
        ),
        (
            previous.strike,
            next.strike,
            extra_source.strike.as_ref(),
            STRIKE_ATTRIBUTE,
        ),
        (
            previous.bold,
            next.bold,
            extra_source.bold.as_ref(),
            BOLD_ATTRIBUTE,
        ),
    ] {
        plan_typed_attribute_length(
            source,
            attr,
            old.map(PlannedAttributeValue::Bool),
            new.map(PlannedAttributeValue::Bool),
            name,
            maximum_attribute_bytes,
            &extra_source.start_tag,
            owner.clone(),
            output_len,
            owner_len,
        )?;
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PlannedAttributeValue {
    Hex(u32),
    Bool(bool),
}

fn planned_attribute_encoded_len(value: PlannedAttributeValue) -> Result<usize> {
    match value {
        // ST_UnsignedIntHex is `hexBinary` with length=4 in the local
        // SpreadsheetML schema, so every emitted value is exactly eight
        // hexadecimal digits, including leading zeroes.
        PlannedAttributeValue::Hex(_) => Ok(8),
        PlannedAttributeValue::Bool(value) => Ok(if value { 4 } else { 5 }),
    }
}

fn plan_typed_attribute_length(
    source: &[u8],
    current: Option<&AttributeSourceData>,
    previous: Option<PlannedAttributeValue>,
    next: Option<PlannedAttributeValue>,
    name: &[u8],
    maximum_attribute_bytes: usize,
    insertion_tag: &Range<usize>,
    owner: Range<usize>,
    output_len: &mut usize,
    owner_len: &mut usize,
) -> Result<()> {
    if previous == next {
        return Ok(());
    }
    if let Some(next) = next {
        let encoded = planned_attribute_encoded_len(next)?;
        let observed = name.len().checked_add(encoded).unwrap_or(usize::MAX);
        if observed > maximum_attribute_bytes {
            return Err(Error::limit(
                LimitKind::AttributeBytes,
                observed,
                maximum_attribute_bytes,
            ));
        }
    }
    let (range, replacement_len) = match (current, next) {
        (Some(current), Some(next)) => {
            (current.value.clone(), planned_attribute_encoded_len(next)?)
        },
        (Some(current), None) => (current.whole.clone(), 0),
        (None, Some(next)) => {
            let encoded = planned_attribute_encoded_len(next)?;
            let insertion = attribute_insertion(source, insertion_tag, encoded)?;
            (
                insertion..insertion,
                name.len()
                    .checked_add(encoded)
                    .and_then(|value| value.checked_add(4))
                    .ok_or_else(|| invalid("pivotTableData attribute length overflows"))?,
            )
        },
        (None, None) => return Ok(()),
    };
    apply_length_delta(&range, replacement_len, owner, output_len, owner_len)
}

fn append_extra_edits(
    source: &[u8],
    extra_source: &ExtraSource,
    previous: Option<&ExtraState>,
    next: Option<&ExtraState>,
    edits: &mut Vec<(Range<usize>, Vec<u8>)>,
    maximum_attribute_bytes: usize,
) -> Result<()> {
    let (Some(previous), Some(next)) = (previous, next) else {
        return Err(invalid("pivotTableData extra source state is incomplete"));
    };
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.background_color.as_ref(),
        previous.background_color,
        next.background_color,
        |value| format!("{value:08X}"),
        BACKGROUND_COLOR_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.foreground_color.as_ref(),
        previous.foreground_color,
        next.foreground_color,
        |value| format!("{value:08X}"),
        FOREGROUND_COLOR_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.italic.as_ref(),
        previous.italic,
        next.italic,
        |value| value.to_string(),
        ITALIC_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.underline.as_ref(),
        previous.underline,
        next.underline,
        |value| value.to_string(),
        UNDERLINE_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.strike.as_ref(),
        previous.strike,
        next.strike,
        |value| value.to_string(),
        STRIKE_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    append_typed_attr_edit(
        source,
        &extra_source.start_tag,
        extra_source.bold.as_ref(),
        previous.bold,
        next.bold,
        |value| value.to_string(),
        BOLD_ATTRIBUTE,
        edits,
        maximum_attribute_bytes,
    )?;
    Ok(())
}

fn append_typed_attr_edit<T, F>(
    source: &[u8],
    start_tag: &Range<usize>,
    current: Option<&AttributeSourceData>,
    previous: Option<T>,
    next: Option<T>,
    format: F,
    name: &[u8],
    edits: &mut Vec<(Range<usize>, Vec<u8>)>,
    maximum_attribute_bytes: usize,
) -> Result<()>
where
    T: Copy + PartialEq,
    F: Fn(T) -> String,
{
    if previous == next {
        return Ok(());
    }
    let previous = previous.map(&format);
    let next = next.map(format);
    append_attr_edit(
        source,
        start_tag,
        current,
        previous.as_deref(),
        next.as_deref(),
        name,
        edits,
        maximum_attribute_bytes,
    )
}

fn append_attr_edit(
    source: &[u8],
    start_tag: &Range<usize>,
    current: Option<&AttributeSourceData>,
    previous: Option<&str>,
    next: Option<&str>,
    name: &[u8],
    edits: &mut Vec<(Range<usize>, Vec<u8>)>,
    maximum_attribute_bytes: usize,
) -> Result<()> {
    if previous == next {
        return Ok(());
    }
    let encoded = next.map(try_escaped_xstring).transpose()?;
    if let Some(encoded) = &encoded {
        let observed = name.len().checked_add(encoded.len()).unwrap_or(usize::MAX);
        if observed > maximum_attribute_bytes {
            return Err(Error::limit(
                LimitKind::AttributeBytes,
                observed,
                maximum_attribute_bytes,
            ));
        }
    }
    let (range, replacement) = match (current, encoded) {
        (Some(current), Some(encoded)) => (current.value.clone(), encoded),
        (Some(current), None) => (current.whole.clone(), Vec::new()),
        (None, Some(encoded)) => {
            let insertion = attribute_insertion(source, start_tag, encoded.len())?;
            let mut replacement = Vec::new();
            let replacement_len = name
                .len()
                .checked_add(encoded.len())
                .and_then(|length| length.checked_add(4))
                .ok_or_else(|| invalid("pivotTableData attribute replacement overflows"))?;
            replacement
                .try_reserve_exact(replacement_len)
                .map_err(|source| Error::Allocation {
                    resource: "pivotTableData attribute replacement",
                    source,
                })?;
            replacement.push(b' ');
            replacement.extend_from_slice(name);
            replacement.extend_from_slice(b"=\"");
            replacement.extend_from_slice(&encoded);
            replacement.push(b'\"');
            (insertion..insertion, replacement)
        },
        (None, None) => return Ok(()),
    };
    edits.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "pivotTableData source edit list",
        source,
    })?;
    edits.push((range, replacement));
    Ok(())
}

fn attribute_insertion(source: &[u8], owner: &Range<usize>, _value_len: usize) -> Result<usize> {
    let end = owner.end;
    if end >= 2 && source.get(end - 2..end) == Some(b"/>") {
        return Ok(end - 2);
    }
    if end >= 1 && source.get(end - 1) == Some(&b'>') {
        return Ok(end - 1);
    }
    Err(invalid(
        "pivotTableData attribute insertion point is ambiguous",
    ))
}

fn caller_part_limit(limits: SourceLimits) -> usize {
    MAX_PART_BYTES.min(usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX))
}

fn caller_attribute_limit(limits: SourceLimits) -> usize {
    MAX_ATTRIBUTE_TEXT_BYTES.min(limits.max_xml_attribute_bytes())
}

pub(crate) fn view_workbook_with_limits<'a>(
    workbook: &Workbook,
    selector: impl Into<PivotTableSelector<'a>>,
    limits: &Limits,
) -> Result<Option<PivotTableDataView>> {
    match Snapshot::load_with_limits(workbook.pivot_package(), selector, limits) {
        Ok(snapshot) => Ok(Some(snapshot.data().clone())),
        Err(Error::Invalid(message)) if message == "PivotTable has no pivotTableData extension" => {
            Ok(None)
        },
        Err(error) => Err(error),
    }
}

crate::record::record_family!("pivotTableData");

/// A bounded dimension of `pivotTableData` reads and edits: the configured
/// [`Limits`], the source XML attribute budget and the fixed SpreadsheetML
/// text bounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Authored `pivotRow` elements ([`Limits::max_authored_rows`]).
    AuthoredRows,
    /// Authored cells across all rows ([`Limits::max_authored_cells`]).
    AuthoredCells,
    /// The declared `rowCount` ([`Limits::max_row_count`]).
    RowCount,
    /// The declared `columnCount` ([`Limits::max_column_count`]).
    ColumnCount,
    /// Bytes retained by a read or staged by an edit
    /// ([`Limits::max_retained_bytes`]).
    RetainedBytes,
    /// Bytes of one cell text value (the source XML attribute budget).
    TextBytes,
    /// UTF-16 code units of one string value (SpreadsheetML `ST_Xstring`).
    TextUnits,
    /// Bytes of one attribute value (the source XML attribute budget).
    AttributeBytes,
    /// Bytes of one element name.
    NameBytes,
    /// Bytes of the retained extension owner fragment.
    FragmentBytes,
}

impl LimitKind {
    /// A short structural description of the dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AuthoredRows => "pivotTableData authored rows",
            Self::AuthoredCells => "pivotTableData authored cells",
            Self::RowCount => "pivotTableData rowCount",
            Self::ColumnCount => "pivotTableData columnCount",
            Self::RetainedBytes => "pivotTableData retained bytes",
            Self::TextBytes => "bytes of one pivotTableData text value",
            Self::TextUnits => "UTF-16 units of one pivotTableData string",
            Self::AttributeBytes => "bytes of one pivotTableData attribute",
            Self::NameBytes => "bytes of one pivotTableData element name",
            Self::FragmentBytes => "bytes of the pivotTableData extension owner",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// The configuration of a `pivotTableData` edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Bound the edit by `limits` instead of the package's semantic limits
    /// ([`crate::SemanticLimits::pivot_table_data`]).
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    pub(crate) fn limits_or(self, default: Limits) -> Limits {
        match self.limits {
            Some(limits) => limits,
            None => default,
        }
    }
}

fn package_commit(workbook: Workbook, patch: Patch, changed: bool) -> Commit {
    Commit::new(
        crate::Package::from_workbook(workbook),
        patch,
        Diagnostics::new(changed),
    )
}

pub(crate) fn edit_workbook_with_limits<'a>(
    workbook: &Workbook,
    selector: impl Into<PivotTableSelector<'a>>,
    limits: &Limits,
) -> Result<Edit> {
    Edit::with_limits(workbook, selector, limits)
}

/// A staged scalar edit of one PivotTable's `pivotTableData` owner of a
/// [`crate::Package`], started by [`crate::Package::edit_pivot_table_data`].
/// Construction retains only the workbook owner, bounded selector and
/// semantic policy.  The source graph and staged cells are materialized by
/// the first fallible getter or edit operation; failed admission leaves the
/// edit retryable.  Dropping it discards it.
pub struct Edit {
    source: Workbook,
    selector: OwnedTableDataSelector,
    limits: Limits,
    materialized: Option<TableDataEditState>,
}

#[derive(Debug)]
struct TableDataEditState {
    before: Snapshot,
    staged: Vec<CellState>,
    positions: HashMap<PivotCellAddress, CellPosition>,
    staged_text_ledger: Vec<usize>,
    staged_text_bytes: usize,
}

#[derive(Debug)]
enum OwnedTableDataSelector {
    Name(String),
    Position(usize),
}

impl OwnedTableDataSelector {
    fn new<'a>(selector: PivotTableSelector<'a>, limits: SourceLimits) -> Result<Self> {
        match selector {
            PivotTableSelector::Position(position) => Ok(Self::Position(position)),
            PivotTableSelector::Name(name) => {
                let maximum = MAX_NAME_BYTES.min(caller_attribute_limit(limits));
                if name.len() > maximum {
                    return Err(Error::limit(LimitKind::NameBytes, name.len(), maximum));
                }
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|source| Error::Allocation {
                        resource: "pivotTableData selector",
                        source,
                    })?;
                owned.push_str(name);
                Ok(Self::Name(owned))
            },
        }
    }

    fn borrowed(&self) -> PivotTableSelector<'_> {
        match self {
            Self::Name(name) => PivotTableSelector::Name(name),
            Self::Position(position) => PivotTableSelector::Position(*position),
        }
    }
}

impl Edit {
    pub(crate) fn with_limits<'a>(
        source: &Workbook,
        selector: impl Into<PivotTableSelector<'a>>,
        limits: &Limits,
    ) -> Result<Self> {
        let selector =
            OwnedTableDataSelector::new(selector.into(), source.pivot_package().read_limits())?;
        Ok(Self {
            source: source.clone(),
            selector,
            limits: *limits,
            materialized: None,
        })
    }

    fn materialize(&mut self) -> Result<()> {
        if self.materialized.is_some() {
            return Ok(());
        }
        let before = Snapshot::load_with_limits(
            self.source.pivot_package(),
            self.selector.borrowed(),
            &self.limits,
        )?;
        ensure_editable(&before)?;
        let staging_bytes = staging_retained_bytes(before.cells.len())?;
        let retained_with_staging = before
            .retained_bytes
            .checked_add(staging_bytes)
            .ok_or_else(|| invalid("Workbook pivotTableData staged bytes overflow"))?;
        enforce_retained_limit(retained_with_staging, self.limits.max_retained_bytes)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.cells.len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook pivotTableData staged cells",
                source,
            })?;
        for record in &before.cells {
            staged.push(record.state.clone());
        }
        let positions = build_cell_positions(&before)?;
        let mut staged_text_ledger = Vec::new();
        staged_text_ledger
            .try_reserve_exact(before.cells.len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook pivotTableData staged text ledger",
                source,
            })?;
        staged_text_ledger.extend(std::iter::repeat_n(0, before.cells.len()));
        self.materialized = Some(TableDataEditState {
            before,
            staged,
            positions,
            staged_text_ledger,
            staged_text_bytes: 0,
        });
        Ok(())
    }

    fn state(&self) -> Result<&TableDataEditState> {
        self.materialized
            .as_ref()
            .ok_or_else(|| invalid("pivotTableData edit is not materialized"))
    }

    fn state_mut(&mut self) -> Result<&mut TableDataEditState> {
        self.materialized
            .as_mut()
            .ok_or_else(|| invalid("pivotTableData edit is not materialized"))
    }

    pub fn before(&mut self) -> Result<&PivotTableDataView> {
        self.materialize()?;
        Ok(&self.state()?.before.value)
    }

    pub fn set_value(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: CellChange,
    ) -> Result<bool> {
        self.materialize()?;
        let read_limits = self.source.pivot_package().read_limits();
        let state = self.state_mut()?;
        stage_set_value(
            &mut state.staged,
            &state.before,
            &state.positions,
            &mut state.staged_text_ledger,
            &mut state.staged_text_bytes,
            address.into(),
            value,
            read_limits,
        )
    }

    /// Replace an existing text cell after validating the borrowed value.
    ///
    /// The value is checked before any staged `String` is allocated.  This is
    /// the preferred path for callers that do not already own the replacement
    /// text.
    pub fn set_text(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        self.materialize()?;
        let read_limits = self.source.pivot_package().read_limits();
        let state = self.state_mut()?;
        stage_set_borrowed_value(
            &mut state.staged,
            &state.before,
            &state.positions,
            &mut state.staged_text_ledger,
            &mut state.staged_text_bytes,
            address.into(),
            PivotCellType::Text,
            value.as_ref(),
            read_limits,
        )
    }

    /// Replace an existing error cell after validating the borrowed value.
    ///
    /// The value is checked before any staged `String` is allocated.  This is
    /// the preferred path for callers that do not already own the replacement
    /// error text.
    pub fn set_error(
        &mut self,
        address: impl Into<PivotCellAddress>,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        self.materialize()?;
        let read_limits = self.source.pivot_package().read_limits();
        let state = self.state_mut()?;
        stage_set_borrowed_value(
            &mut state.staged,
            &state.before,
            &state.positions,
            &mut state.staged_text_ledger,
            &mut state.staged_text_bytes,
            address.into(),
            PivotCellType::Error,
            value.as_ref(),
            read_limits,
        )
    }

    pub fn set_blank(&mut self, address: impl Into<PivotCellAddress>) -> Result<bool> {
        self.set_value(address, CellChange::Blank)
    }

    pub fn set_extra(
        &mut self,
        address: impl Into<PivotCellAddress>,
        edit: CellExtra,
    ) -> Result<bool> {
        self.materialize()?;
        let read_limits = self.source.pivot_package().read_limits();
        let state = self.state_mut()?;
        stage_set_extra(
            &mut state.staged,
            &state.positions,
            address.into(),
            edit,
            read_limits,
        )
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        let Some(state) = self.materialized.as_mut() else {
            return;
        };
        state.staged.clear();
        state
            .staged
            .extend(state.before.cells.iter().map(|record| record.state.clone()));
        state.staged_text_ledger.fill(0);
        state.staged_text_bytes = 0;
    }

    /// Validate and publish the staged cells as a new package. An unchanged
    /// edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook and
    /// the rewrite and validation errors of the publication.
    pub fn commit(mut self) -> Result<Commit> {
        self.materialize()?;
        let state = self
            .materialized
            .take()
            .ok_or_else(|| invalid("pivotTableData edit is not materialized"))?;
        let changed = !states_equal(&state.before, &state.staged);
        let source = self.source;
        let before = state.before;
        let staged = state.staged;
        if !changed {
            let workbook = source.clone();
            let low_patch = SourcePatch::new(before.clone(), before);
            let patch = Patch::new(source, workbook.clone(), low_patch);
            return Ok(package_commit(workbook, patch, false));
        }
        if source.pivot_package().is_signed()
            || source.pivot_package().requires_signature_edit_policy()
        {
            return Err(Error::Signed);
        }
        ensure_editable(&before)?;
        let mut candidate = source.pivot_package().clone();
        let output = rewrite_data(
            &before,
            &staged,
            caller_part_limit(candidate.read_limits()),
            candidate.read_limits().max_total_part_bytes(),
            caller_attribute_limit(candidate.read_limits()),
            package_total_bytes(&candidate)?,
        )?;
        candidate
            .get_part_mut(before.table_part())?
            .set_blob_shared(Arc::new(output));
        validate_candidate_limits(&candidate)?;
        let after = Snapshot::load_with_limits(
            &candidate,
            PivotTableSelector::Position(before.selection),
            &before.limits,
        )?;
        if !after.same_readset(&before) || !states_equal(&after, &staged) {
            return Err(invalid(
                "pivotTableData Workbook publication verification failed",
            ));
        }
        let low_patch = SourcePatch::new(before, after);
        let workbook = source.adopt_published_package(candidate)?;
        let patch = Patch::new(source, workbook.clone(), low_patch);
        Ok(package_commit(workbook, patch, true))
    }
}

/// Exact reversible ordinary Workbook patch.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Workbook,
    after: Workbook,
    source_patch: SourcePatch,
}

impl Patch {
    fn new(before: Workbook, after: Workbook, source_patch: SourcePatch) -> Self {
        Self {
            before,
            after,
            source_patch,
        }
    }

    /// Typed `pivotTableData` state required before application.
    #[must_use]
    pub fn before(&self) -> &PivotTableDataView {
        &self.source_patch.before.value
    }

    /// Typed `pivotTableData` state produced by application.
    #[must_use]
    pub fn after(&self) -> &PivotTableDataView {
        &self.source_patch.after.value
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            source_patch: self.source_patch.inverse(),
        }
    }

    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source_patch.is_empty()
    }

    pub(crate) fn apply(&self, source: &crate::Package) -> Result<Commit> {
        let source = source.workbook_ref();
        source.ensure_mutation_allowed("apply_pivot_table_data")?;
        let mut candidate = source.pivot_package().clone();
        self.source_patch.apply(&mut candidate)?;
        if self.is_noop() {
            return Ok(package_commit(
                source.clone(),
                Self::new(source.clone(), source.clone(), self.source_patch.clone()),
                false,
            ));
        }
        let selection = self.source_patch.before.selection;
        let before = Snapshot::load(
            source.pivot_package(),
            PivotTableSelector::Position(selection),
        )?;
        let after = Snapshot::load(&candidate, PivotTableSelector::Position(selection))?;
        let low_patch = SourcePatch::new(before, after);
        let changed = !low_patch.is_empty();
        let workbook = source.adopt_published_package(candidate)?;
        let patch = Patch::new(source.clone(), workbook.clone(), low_patch);
        Ok(package_commit(workbook, patch, changed))
    }
}

pub fn load<'a>(
    package: &OpcPackage,
    selector: impl Into<PivotTableSelector<'a>>,
) -> Result<Snapshot> {
    Snapshot::load(package, selector)
}

/// Read a C444 owner with an explicit semantic resource policy.
#[cfg(test)]
pub(crate) fn load_with_limits<'a>(
    package: &OpcPackage,
    selector: impl Into<PivotTableSelector<'a>>,
    limits: &Limits,
) -> Result<Snapshot> {
    Snapshot::load_with_limits(package, selector, limits)
}

#[cfg(test)]
pub(crate) fn edit<'a>(
    package: &'a mut OpcPackage,
    selector: impl Into<PivotTableSelector<'a>>,
) -> Result<Transaction<'a>> {
    Transaction::new(package, selector)
}

/// Start a low-level C444 edit with an explicit semantic resource policy.
#[cfg(test)]
pub(crate) fn edit_with_limits<'a>(
    package: &'a mut OpcPackage,
    selector: impl Into<PivotTableSelector<'a>>,
    limits: &Limits,
) -> Result<Transaction<'a>> {
    Transaction::with_limits(package, selector, limits)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};

    const CORE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const EXT: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
    const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        let table_uri = PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap();
        let cache_uri = PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap();
        let workbook_xml = format!(
            r#"<workbook xmlns="{CORE}" xmlns:r="{REL}" xmlns:x15="{EXT}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches><extLst><ext uri="{PIVOT_TABLE_REFERENCES_URI}"><x15:pivotTableReferences><x15:pivotTableReference r:id="rId3"/></x15:pivotTableReferences></ext></extLst></workbook>"#
        );
        let mut workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            workbook_xml.into_bytes(),
        );
        workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        workbook.relate_to(
            "pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        workbook.relate_to("pivotTables/pivotTable1.xml", rt::PIVOT_TABLE);
        let worksheet = BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.into(),
            format!(r#"<worksheet xmlns="{CORE}"><sheetData/></worksheet>"#).into_bytes(),
        );
        let cache = BlobPart::new(
            cache_uri,
            ct::SML_PIVOT_CACHE_DEFINITION.into(),
            format!(
                r#"<pivotCacheDefinition xmlns="{CORE}" xmlns:x14="{X14}" xmlns:x15="{EXT}"><cacheSource type="external"/><extLst><ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition pivotCacheId="7"/></ext><ext uri="{PIVOT_CACHE_ID_VERSION_URI}"><x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></ext></extLst></pivotCacheDefinition>"#
            )
            .into_bytes(),
        );
        let table = BlobPart::new(
            table_uri,
            ct::SML_PIVOT_TABLE.into(),
            format!(
                r#"<pivotTableDefinition xmlns="{CORE}" xmlns:x15="{EXT}" name="Pivot" cacheId="7"><location ref="A1:B2"/><rowItems count="1"/><colItems count="2"/><extLst><ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData rowCount="1" columnCount="2" cacheId="7"><x15:pivotRow r="0" count="2"><x15:c i="0" t="str"><x15:v>old</x15:v></x15:c><x15:c i="1" t="b"><x15:v>false</x15:v><x15:x b="1"/></x15:c></x15:pivotRow></x15:pivotTableData></ext></extLst></pivotTableDefinition>"#
            )
            .into_bytes(),
        );
        let mut table = table;
        table.relate_to(
            "../pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(worksheet));
        package.add_part(Box::new(cache));
        package.add_part(Box::new(table));
        package
    }

    #[test]
    fn reads_and_rewrites_authored_cells() {
        let mut package = package();
        let snapshot = Snapshot::load(&package, "Pivot").unwrap();
        assert_eq!(snapshot.data().row_count(), 1);
        assert_eq!(snapshot.data().column_count(), 2);
        assert_eq!(snapshot.cell((0, 0)).unwrap().unwrap().value_text(), "old");
        assert_eq!(
            snapshot
                .cell((0, 1))
                .unwrap()
                .unwrap()
                .extra()
                .unwrap()
                .bold,
            Some(true)
        );
        let mut transaction = Transaction::new(&mut package, "Pivot").unwrap();
        assert!(transaction.set_text((0, 0), "new").unwrap());
        assert!(
            transaction
                .set_extra(
                    (0, 1),
                    CellExtra {
                        background_color: Attribute::Keep,
                        foreground_color: Attribute::Keep,
                        italic: Attribute::Keep,
                        underline: Attribute::Keep,
                        strike: Attribute::Keep,
                        bold: Attribute::Set(false),
                    },
                )
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        let text = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
        assert!(text.contains(">new</x15:v>"));
        assert!(text.contains(r#"<x15:x b="false"/>"#));
        let original = commit.patch().inverse();
        original.apply(&mut package).unwrap();
        assert!(
            std::str::from_utf8(Snapshot::load(&package, "Pivot").unwrap().source_xml())
                .unwrap()
                .contains(">old</x15:v>")
        );
    }

    #[test]
    fn owned_replacement_admits_old_arc_only_as_transient_peak() {
        let package = package();
        let snapshot = Snapshot::load(&package, "Pivot").unwrap();
        let first = "R".repeat(1_024);
        let required = snapshot
            .retained_bytes
            .checked_add(staging_retained_bytes(snapshot.cells.len()).unwrap())
            .and_then(|bytes| {
                bytes.checked_add(replacement_text_retained_bytes(first.len()).unwrap())
            })
            .and_then(|bytes| bytes.checked_add(first.capacity()))
            .unwrap();
        let limits = Limits::new().with_max_retained_bytes(required);
        let mut package = package;
        let mut transaction = Transaction::with_limits(&mut package, "Pivot", &limits).unwrap();
        assert!(
            transaction
                .set_value((0, 0), CellChange::text(first))
                .unwrap()
        );
        let error = transaction
            .set_value((0, 0), CellChange::text("S".repeat(1_024)))
            .unwrap_err();
        assert!(matches!(error, Error::Limit { .. }));
        assert!(transaction.commit().unwrap().changed());
    }

    #[test]
    fn borrowed_replacement_admits_old_arc_only_as_transient_peak() {
        let package = package();
        let snapshot = Snapshot::load(&package, "Pivot").unwrap();
        let first = "R".repeat(1_024);
        let required = snapshot
            .retained_bytes
            .checked_add(staging_retained_bytes(snapshot.cells.len()).unwrap())
            .and_then(|bytes| {
                bytes.checked_add(replacement_text_retained_bytes(first.len()).unwrap())
            })
            .and_then(|bytes| bytes.checked_add(first.len()))
            .unwrap();
        let limits = Limits::new().with_max_retained_bytes(required);
        let workbook = crate::Package::from_opc(package).unwrap();
        let mut transaction = workbook
            .edit_pivot_table_data_with_options("Pivot", EditOptions::default().limits(limits))
            .unwrap();
        assert!(transaction.set_text((0, 0), &first).unwrap());
        let error = transaction.set_text((0, 0), "S".repeat(1_024)).unwrap_err();
        assert!(matches!(error, Error::Limit { .. }));
        assert!(transaction.commit().unwrap().diagnostics().changed());
    }

    #[test]
    fn extra_edits_are_failure_atomic_for_low_level_and_workbook_transactions() {
        let invalid_edit = || CellExtra {
            background_color: Attribute::Set("00112233".to_owned()),
            foreground_color: Attribute::Set("not-a-color".to_owned()),
            italic: Attribute::Keep,
            underline: Attribute::Keep,
            strike: Attribute::Keep,
            bold: Attribute::Keep,
        };

        let mut low_level_package = package();
        let mut transaction = Transaction::new(&mut low_level_package, "Pivot").unwrap();
        assert!(transaction.set_extra((0, 1), invalid_edit()).is_err());
        assert!(!transaction.commit().unwrap().changed());

        let workbook = crate::Package::from_opc(package()).unwrap();
        let mut transaction = workbook.edit_pivot_table_data("Pivot").unwrap();
        assert!(transaction.set_extra((0, 1), invalid_edit()).is_err());
        assert!(!transaction.commit().unwrap().diagnostics().changed());
    }

    #[test]
    fn applied_workbook_patch_retains_the_incoming_snapshot_pair() {
        let original = crate::Package::from_opc(package()).unwrap();
        let mut edit = original.edit_pivot_table_data("Pivot").unwrap();
        edit.set_text((0, 0), "replacement").unwrap();
        let committed = edit.commit().unwrap();

        let mut incoming_package = package();
        incoming_package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/unrelated.bin").unwrap(),
            "application/octet-stream".into(),
            b"incoming-only".to_vec(),
        )));
        let incoming = crate::Package::from_opc(incoming_package).unwrap();
        let incoming_bytes = incoming.to_bytes().unwrap();
        assert_ne!(incoming_bytes, original.to_bytes().unwrap());

        let applied = committed.patch().apply(&incoming).unwrap();
        assert_eq!(
            crate::writer::to_bytes(applied.patch().before.opc_package()).unwrap(),
            incoming_bytes
        );
        assert_eq!(
            crate::writer::to_bytes(applied.patch().after.opc_package()).unwrap(),
            applied.package().to_bytes().unwrap()
        );
        let restored = applied.patch().inverse().apply(applied.package()).unwrap();
        assert_eq!(restored.package().to_bytes().unwrap(), incoming_bytes);
    }

    fn local_traversal_visits_for_authored_cells(cell_count: usize) -> usize {
        let mut xml = format!(
            r#"<pivotTableDefinition xmlns="{CORE}" xmlns:x15="{EXT}"><rowItems count="1"/><colItems count="{cell_count}"/><extLst><ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData rowCount="1" columnCount="{cell_count}" cacheId="7"><x15:pivotRow r="0" count="{cell_count}">"#
        );
        for index in 0..cell_count {
            xml.push_str(&format!(
                r#"<x15:c i="{index}" t="str"><x15:v>value</x15:v></x15:c>"#
            ));
        }
        xml.push_str(
            r#"</x15:pivotRow></x15:pivotTableData></ext></extLst></pivotTableDefinition>"#,
        );
        let source = xml.into_bytes();
        let scan = scan_xml(&source, "pivotTableDefinition", SourceLimits::default()).unwrap();
        LOCAL_SUBTREE_BOUNDARY_VISITS.with(|counter| counter.set(0));
        let mut retained = RetainedBudget::new(Limits::default());
        let parsed =
            parse_data_payload(&scan, &source, 7, SourceLimits::default(), &mut retained).unwrap();
        assert_eq!(parsed.cells.len(), cell_count);
        LOCAL_SUBTREE_BOUNDARY_VISITS.with(Cell::get)
    }

    #[test]
    fn authored_cell_traversal_is_parent_local() {
        let small = local_traversal_visits_for_authored_cells(8);
        let medium = local_traversal_visits_for_authored_cells(64);
        let large = local_traversal_visits_for_authored_cells(512);
        assert!(small > 0);
        assert!(medium > small);
        assert!(large > medium);
        assert!(medium <= small * 10);
        assert!(large <= medium * 10);
    }
}
