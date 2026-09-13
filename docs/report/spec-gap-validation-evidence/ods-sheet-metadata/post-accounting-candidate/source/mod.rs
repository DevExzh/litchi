//! Source-backed ODS worksheet metadata transactions.
//!
//! This owner is deliberately narrower than the worksheet graph.  It indexes
//! the physical spans that own consolidation, label ranges, cell-range-source,
//! and detective metadata, then changes only those spans (and the minimal
//! repeated-row/cell split required by a logical selector).

#[allow(unused_qualifications)]
pub mod detective;
mod index;

use std::{
    collections::BTreeMap, fmt, io::Write, mem::size_of, ops::Range as ByteRange, sync::Arc,
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionLimits, Position, Profile,
    Reservation, Resource, Result, SourceVersion,
};
use litchi_odf_common::core::{
    SourceContentPublicationError, SourceContentPublicationOptions, SourceContentPublicationReport,
};

pub use crate::model::consolidation::{Options, UseLabels};
pub use crate::model::detective::{
    Detective, Direction, HighlightedRange, Operation, OperationKind,
};
pub use crate::model::label_range::{Orientation, Range as LabelRange};
pub use crate::model::source::CellRange;
pub use litchi_core::Position as SheetPosition;

pub use litchi_odf_common::core::{
    SourceContentPublicationError as PublicationError,
    SourceContentPublicationOptions as PublicationOptions,
    SourceContentPublicationReport as PublicationReport,
};

pub(crate) const DEFAULT_MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const DEFAULT_MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const DEFAULT_MAX_SCRATCH_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const DEFAULT_MAX_DEPTH: usize = 1_024;
pub(crate) const DEFAULT_MAX_EVENTS: usize = 4_194_304;
pub(crate) const DEFAULT_MAX_SHEETS: usize = 65_536;
pub(crate) const DEFAULT_MAX_CELLS: usize = 4_194_304;
pub(crate) const DEFAULT_MAX_OPERATIONS: usize = 4_096;
pub(crate) const DEFAULT_MAX_NAMESPACE_BINDINGS: usize = 4_096;
pub(crate) const DEFAULT_MAX_WORK_UNITS: u64 = 16_000_000_000;

/// A finite caller profile for one metadata scan or transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_scratch_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_sheets: usize,
    max_cells: usize,
    max_operations: usize,
    max_labels: usize,
    max_detective_ranges: usize,
    max_detective_operations: usize,
    max_namespace_bindings: usize,
    max_text_bytes: usize,
    max_work_units: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: DEFAULT_MAX_INPUT_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_scratch_bytes: DEFAULT_MAX_SCRATCH_BYTES,
            max_depth: DEFAULT_MAX_DEPTH,
            max_events: DEFAULT_MAX_EVENTS,
            max_sheets: DEFAULT_MAX_SHEETS,
            max_cells: DEFAULT_MAX_CELLS,
            max_operations: DEFAULT_MAX_OPERATIONS,
            max_labels: detective::MAX_ITEMS,
            max_detective_ranges: detective::MAX_ITEMS,
            max_detective_operations: detective::MAX_ITEMS,
            max_namespace_bindings: DEFAULT_MAX_NAMESPACE_BINDINGS,
            max_text_bytes: detective::MAX_TEXT_BYTES,
            max_work_units: DEFAULT_MAX_WORK_UNITS,
        }
    }
}

impl Limits {
    /// Construct a finite profile with the core input/output/depth/event/work bounds.
    pub fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_work_units: u64,
    ) -> Result<Self> {
        let value = Self {
            max_input_bytes,
            max_output_bytes,
            max_scratch_bytes: DEFAULT_MAX_SCRATCH_BYTES.min(max_output_bytes),
            max_depth,
            max_events,
            max_work_units,
            ..Self::default()
        };
        value.validate()?;
        Ok(value)
    }

    /// Conservative finite server profile.
    #[must_use]
    pub const fn server() -> Self {
        Self {
            max_input_bytes: DEFAULT_MAX_INPUT_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_scratch_bytes: 16 * 1024 * 1024,
            max_depth: 256,
            max_events: 1_048_576,
            max_sheets: 65_536,
            max_cells: 4_194_304,
            max_operations: DEFAULT_MAX_OPERATIONS,
            max_labels: 65_536,
            max_detective_ranges: 65_536,
            max_detective_operations: 65_536,
            max_namespace_bindings: 512,
            max_text_bytes: 4 * 1024 * 1024,
            max_work_units: 1_000_000_000,
        }
    }

    /// Set owner and operation ceilings.
    pub fn with_item_limits(
        mut self,
        max_sheets: usize,
        max_cells: usize,
        max_operations: usize,
    ) -> Result<Self> {
        self.max_sheets = max_sheets;
        self.max_cells = max_cells;
        self.max_operations = max_operations;
        self.validate()?;
        Ok(self)
    }

    /// Set label and detective collection ceilings.
    pub fn with_metadata_limits(
        mut self,
        max_labels: usize,
        max_ranges: usize,
        max_detective_operations: usize,
    ) -> Result<Self> {
        self.max_labels = max_labels;
        self.max_detective_ranges = max_ranges;
        self.max_detective_operations = max_detective_operations;
        self.validate()?;
        Ok(self)
    }

    /// Set caller output and scratch ceilings.
    pub fn with_output_limits(
        mut self,
        max_output_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Result<Self> {
        self.max_output_bytes = max_output_bytes;
        self.max_scratch_bytes = max_scratch_bytes;
        self.validate()?;
        Ok(self)
    }

    /// Set scalar and namespace limits.
    pub fn with_scalar_limits(
        mut self,
        max_text_bytes: usize,
        max_namespace_bindings: usize,
    ) -> Result<Self> {
        self.max_text_bytes = max_text_bytes;
        self.max_namespace_bindings = max_namespace_bindings;
        self.validate()?;
        Ok(self)
    }

    /// Input byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }
    /// Candidate output byte ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    /// Candidate scratch ceiling.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }
    /// XML depth ceiling.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
    /// XML event ceiling.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }
    /// Sheet ceiling.
    #[must_use]
    pub const fn max_sheets(self) -> usize {
        self.max_sheets
    }
    /// Physical cell ceiling.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }
    /// Staged operation ceiling.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.max_operations
    }
    /// Label-item ceiling.
    #[must_use]
    pub const fn max_labels(self) -> usize {
        self.max_labels
    }
    /// Detective-highlight ceiling.
    #[must_use]
    pub const fn max_detective_ranges(self) -> usize {
        self.max_detective_ranges
    }
    /// Detective-operation ceiling.
    #[must_use]
    pub const fn max_detective_operations(self) -> usize {
        self.max_detective_operations
    }
    /// Namespace-context ceiling.
    #[must_use]
    pub const fn max_namespace_bindings(self) -> usize {
        self.max_namespace_bindings
    }
    /// Scalar text ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }
    /// Format-owned work ceiling.
    #[must_use]
    pub const fn max_work_units(self) -> u64 {
        self.max_work_units
    }

    pub(crate) fn validate(self) -> Result<()> {
        let checks = [
            ("input bytes", self.max_input_bytes, DEFAULT_MAX_INPUT_BYTES),
            (
                "output bytes",
                self.max_output_bytes,
                DEFAULT_MAX_OUTPUT_BYTES,
            ),
            (
                "scratch bytes",
                self.max_scratch_bytes,
                DEFAULT_MAX_SCRATCH_BYTES,
            ),
            ("XML depth", self.max_depth, DEFAULT_MAX_DEPTH),
            ("XML events", self.max_events, DEFAULT_MAX_EVENTS),
            ("sheets", self.max_sheets, DEFAULT_MAX_SHEETS),
            ("cells", self.max_cells, DEFAULT_MAX_CELLS),
            (
                "metadata operations",
                self.max_operations,
                DEFAULT_MAX_OPERATIONS,
            ),
            ("labels", self.max_labels, detective::MAX_ITEMS),
            (
                "detective ranges",
                self.max_detective_ranges,
                detective::MAX_ITEMS,
            ),
            (
                "detective operations",
                self.max_detective_operations,
                detective::MAX_ITEMS,
            ),
            (
                "namespace bindings",
                self.max_namespace_bindings,
                DEFAULT_MAX_NAMESPACE_BINDINGS,
            ),
            ("text bytes", self.max_text_bytes, detective::MAX_TEXT_BYTES),
        ];
        for (name, observed, maximum) in checks {
            if observed == 0 || observed > maximum {
                return Err(Error::InvalidFormat(format!(
                    "ODS metadata {name} limit must be between 1 and {maximum}"
                )));
            }
        }
        if self.max_work_units == 0 || self.max_work_units > DEFAULT_MAX_WORK_UNITS {
            return Err(Error::InvalidFormat(format!(
                "ODS metadata work limit must be between 1 and {DEFAULT_MAX_WORK_UNITS}"
            )));
        }
        Ok(())
    }
}

/// Exact sheet selection component used by [`CellSelector`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum SheetSelector<'a> {
    /// Match one exact worksheet name.
    Name(&'a str),
    /// Match one checked zero-based worksheet position.
    Position(Position),
}

/// Selector-first logical worksheet cell address.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellSelector<'a> {
    sheet: SheetSelector<'a>,
    row: usize,
    column: usize,
}

impl<'a> CellSelector<'a> {
    /// Select a logical cell by exact sheet name.
    #[must_use]
    pub const fn by_name(name: &'a str, row: usize, column: usize) -> Self {
        Self {
            sheet: SheetSelector::Name(name),
            row,
            column,
        }
    }
    /// Select a logical cell by checked zero-based sheet position.
    #[must_use]
    pub const fn by_position(position: Position, row: usize, column: usize) -> Self {
        Self {
            sheet: SheetSelector::Position(position),
            row,
            column,
        }
    }
    /// Sheet component.
    #[must_use]
    pub const fn sheet(self) -> SheetSelector<'a> {
        self.sheet
    }
    /// Exact sheet name when this selector is name-based.
    #[must_use]
    pub const fn sheet_name(self) -> Option<&'a str> {
        match self.sheet {
            SheetSelector::Name(value) => Some(value),
            SheetSelector::Position(_) => None,
        }
    }
    /// Sheet position when this selector is position-based.
    #[must_use]
    pub const fn sheet_position(self) -> Option<Position> {
        match self.sheet {
            SheetSelector::Name(_) => None,
            SheetSelector::Position(value) => Some(value),
        }
    }
    /// Logical zero-based row.
    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }
    /// Logical zero-based column.
    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }
}

/// Physical cell range role returned by the metadata index.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PhysicalLocation {
    row: usize,
    column: usize,
    row_repeat: usize,
    column_repeat: usize,
    merge: crate::worksheet::Merge,
}

impl PhysicalLocation {
    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }
    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }
    #[must_use]
    pub const fn row_repeat(self) -> usize {
        self.row_repeat
    }
    #[must_use]
    pub const fn column_repeat(self) -> usize {
        self.column_repeat
    }
    #[must_use]
    pub const fn merge(self) -> crate::worksheet::Merge {
        self.merge
    }
}

/// Borrowed metadata on one physical cell.
#[derive(Debug)]
pub struct CellMetadata<'a> {
    source: Option<&'a CellRange>,
    detective: Option<&'a Detective>,
    location: PhysicalLocation,
    kind: index::CellKind,
}

impl<'a> CellMetadata<'a> {
    /// Existing cell-range-source metadata.
    #[must_use]
    pub const fn range_source(&self) -> Option<&'a CellRange> {
        self.source
    }
    /// Existing detective metadata.
    #[must_use]
    pub const fn detective(&self) -> Option<&'a Detective> {
        self.detective
    }
    /// Physical run and merge role.
    #[must_use]
    pub const fn location(&self) -> PhysicalLocation {
        self.location
    }
    /// Whether this is an explicit covered-table-cell owner.
    #[must_use]
    pub const fn is_covered(&self) -> bool {
        matches!(self.kind, index::CellKind::CoveredTableCell)
    }
}

/// Cell metadata read result. Missing physical cells are returned as `None`
/// from [`Snapshot::cell_metadata`]; implicit merge-covered coordinates are
/// retained as a distinct value and never project the anchor's metadata.
#[derive(Debug)]
pub enum CellMetadataView<'a> {
    /// Metadata attached to an existing physical cell.
    Physical(CellMetadata<'a>),
    /// A coordinate covered by a merge span with no physical covered cell.
    ImplicitCovered {
        anchor: (usize, usize),
        span: (usize, usize),
    },
}

impl<'a> CellMetadataView<'a> {
    /// Existing source metadata, or `None` for an implicit covered coordinate.
    #[must_use]
    pub fn range_source(&self) -> Option<&'a CellRange> {
        match self {
            Self::Physical(value) => value.range_source(),
            Self::ImplicitCovered { .. } => None,
        }
    }
    /// Existing detective metadata, or `None` for an implicit covered coordinate.
    #[must_use]
    pub fn detective(&self) -> Option<&'a Detective> {
        match self {
            Self::Physical(value) => value.detective(),
            Self::ImplicitCovered { .. } => None,
        }
    }
    /// Physical location, when one exists.
    #[must_use]
    pub fn location(&self) -> Option<PhysicalLocation> {
        match self {
            Self::Physical(value) => Some(value.location()),
            Self::ImplicitCovered { .. } => None,
        }
    }
}

/// Presence-aware label-range view.
#[derive(Clone, Copy, Debug)]
pub struct LabelRangesView<'a> {
    present: bool,
    ranges: &'a [LabelRange],
}

impl<'a> LabelRangesView<'a> {
    /// Whether the direct `table:label-ranges` container exists.
    #[must_use]
    pub const fn is_present(self) -> bool {
        self.present
    }
    /// Alias for [`Self::is_present`].
    #[must_use]
    pub const fn present(self) -> bool {
        self.present
    }
    /// Ordered label children.
    #[must_use]
    pub const fn as_slice(self) -> &'a [LabelRange] {
        self.ranges
    }
    /// Number of label children.
    #[must_use]
    pub const fn len(self) -> usize {
        self.ranges.len()
    }
    /// Whether no label children are present.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.ranges.is_empty()
    }
    /// Iterate over ordered label children.
    pub fn iter(self) -> std::slice::Iter<'a, LabelRange> {
        self.ranges.iter()
    }
}

impl<'a> IntoIterator for LabelRangesView<'a> {
    type Item = &'a LabelRange;
    type IntoIter = std::slice::Iter<'a, LabelRange>;
    fn into_iter(self) -> Self::IntoIter {
        self.ranges.iter()
    }
}

/// Immutable owned content.xml metadata snapshot.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<str>,
    limits: Limits,
    catalog: Arc<index::Catalog>,
    _source_memory: Arc<Reservation>,
    _source_output: Option<Arc<Reservation>>,
    context: ExecutionContext,
    enforce_context_lineage: bool,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("sheets", &self.catalog.tables.len())
            .field("cells", &self.catalog.cells.len())
            .finish_non_exhaustive()
    }
}

impl Snapshot {
    /// Parse one exact content.xml source using a finite default execution profile.
    pub fn parse(source: &str) -> Result<Self> {
        let context = default_context();
        Self::parse_internal(source, Limits::default(), &context, false)
    }

    /// Parse one exact content.xml source under caller limits and context.
    pub fn parse_with_context(
        source: &str,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Self> {
        Self::parse_internal(source, limits, context, true)
    }

    fn parse_internal(
        source: &str,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        limits.validate()?;
        let catalog = index::scan(source, limits, context)?;
        let source_memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(source.len())
                    .map_err(|_| invalid("ODS metadata source size overflows u64"))?,
            )
            .map_err(map_execution)?;
        Ok(Self {
            source: Arc::from(source),
            limits,
            catalog: Arc::new(catalog),
            _source_memory: Arc::new(source_memory),
            _source_output: None,
            context: context.clone(),
            enforce_context_lineage,
        })
    }

    pub(crate) fn parse_arc_with_context(
        source: Arc<str>,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        limits.validate()?;
        let catalog = index::scan(source.as_ref(), limits, context)?;
        let source_memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(source.len())
                    .map_err(|_| invalid("ODS metadata source size overflows u64"))?,
            )
            .map_err(map_execution)?;
        Ok(Self {
            source,
            limits,
            catalog: Arc::new(catalog),
            _source_memory: Arc::new(source_memory),
            _source_output: None,
            context: context.clone(),
            enforce_context_lineage,
        })
    }

    fn parse_rendered_with_context(
        rendered: RenderedCandidate,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        limits.validate()?;
        let catalog = index::scan(rendered.xml.as_ref(), limits, context)?;
        context.check().map_err(map_execution)?;
        Ok(Self {
            source: rendered.xml,
            limits,
            catalog: Arc::new(catalog),
            _source_memory: rendered.memory,
            _source_output: Some(rendered.output),
            context: context.clone(),
            enforce_context_lineage,
        })
    }

    /// Reopen a retained target under the destination snapshot's policy.
    ///
    /// A patch target may have been accepted under another profile or
    /// cancellation token.  Reusing its retained snapshot would therefore
    /// bypass destination admission.  The exact target bytes are checked and
    /// parsed again under the destination context, with a destination-owned
    /// output reservation retained by the returned snapshot.
    fn reopen_for_destination(&self, destination: &Self) -> Result<Self> {
        self.context.check().map_err(map_execution)?;
        destination.context.check().map_err(map_execution)?;
        let target_bytes = self.source.len();
        if target_bytes > destination.limits.max_output_bytes() {
            return Err(limit(
                "output bytes",
                target_bytes,
                destination.limits.max_output_bytes(),
            ));
        }
        let output = destination
            .context
            .reserve(
                Resource::OutputBytes,
                u64::try_from(target_bytes)
                    .map_err(|_| invalid("ODS metadata target size overflows u64"))?,
            )
            .map_err(map_execution)?;
        let mut target = Self::parse_arc_with_context(
            Arc::clone(&self.source),
            destination.limits,
            &destination.context,
            destination.enforce_context_lineage,
        )?;
        target._source_output = Some(Arc::new(output));
        Ok(target)
    }

    /// Exact content.xml source bytes retained by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source
    }
    /// Alias for [`Self::source_xml`].
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.source
    }
    /// Finite limits captured by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Read the singleton consolidation declaration.
    pub fn consolidation(&self) -> Result<Option<&Options>> {
        Ok(self
            .catalog
            .consolidation
            .as_ref()
            .map(|value| &value.value))
    }
    /// Read the presence-aware label-range catalog.
    #[must_use]
    pub fn label_ranges(&self) -> LabelRangesView<'_> {
        self.catalog.labels.as_ref().map_or(
            LabelRangesView {
                present: false,
                ranges: &[],
            },
            |value| LabelRangesView {
                present: true,
                ranges: &value.ranges,
            },
        )
    }
    /// Read focused metadata on one logical cell.
    pub fn cell_metadata<'a>(
        &'a self,
        selector: CellSelector<'_>,
    ) -> Result<Option<CellMetadataView<'a>>> {
        self.catalog
            .cell_for_selector(&selector, &self.context, self.limits)
            .and_then(|physical| self.cell_metadata_for_physical(physical))
    }

    fn cell_metadata_with_budget<'a>(
        &'a self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<CellMetadataView<'a>>> {
        let physical = self
            .catalog
            .cell_for_selector_with_budget(&selector, selector_budget)?;
        self.cell_metadata_for_physical(physical)
    }

    fn cell_metadata_for_physical<'a>(
        &'a self,
        physical: index::PhysicalCell,
    ) -> Result<Option<CellMetadataView<'a>>> {
        match physical {
            index::PhysicalCell::Missing => Ok(None),
            index::PhysicalCell::ImplicitCovered {
                anchor,
                rows,
                columns,
            } => Ok(Some(CellMetadataView::ImplicitCovered {
                anchor,
                span: (rows, columns),
            })),
            index::PhysicalCell::Stored(index) => {
                let cell = self.catalog.cells.get(index).ok_or_else(|| {
                    Error::InvalidFormat("ODS metadata cell index disappeared".to_string())
                })?;
                Ok(Some(CellMetadataView::Physical(CellMetadata {
                    source: cell.source.as_ref(),
                    detective: cell.detective.as_ref(),
                    location: PhysicalLocation {
                        row: cell.row_start,
                        column: cell.column_start,
                        row_repeat: cell.row_repeat,
                        column_repeat: cell.column_repeat,
                        merge: cell.merge,
                    },
                    kind: cell.kind,
                })))
            },
        }
    }

    /// Start a non-consuming semantic metadata edit.
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            consolidation: None,
            consolidation_staged: false,
            labels: None,
            labels_staged: false,
            labels_present: self.catalog.labels.is_some(),
            cell_ops: BTreeMap::new(),
            staged_operations: 0,
            staging_memory: None,
        }
    }
}

/// One logical cell metadata mutation staged against a snapshot.
#[derive(Clone, Debug, Default)]
struct CellOp {
    source: Option<Option<CellRange>>,
    detective: Option<Option<Detective>>,
}

#[derive(Clone, Copy)]
enum CellField {
    Source,
    Detective,
}

/// Source-checked, failure-atomic metadata edit.
pub struct Edit {
    before: Snapshot,
    consolidation: Option<Options>,
    consolidation_staged: bool,
    labels: Option<Vec<LabelRange>>,
    labels_staged: bool,
    labels_present: bool,
    cell_ops: BTreeMap<CellKey, CellOp>,
    staged_operations: usize,
    staging_memory: Option<Reservation>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct CellKey {
    table_index: usize,
    row: usize,
    column: usize,
}

impl CellKey {
    fn from_selector_with_budget(
        catalog: &index::Catalog,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Self> {
        Ok(Self {
            table_index: catalog
                .table_index_for_selector_with_budget(&selector, selector_budget)?,
            row: selector.row,
            column: selector.column,
        })
    }

    fn selector(self) -> CellSelector<'static> {
        CellSelector::by_position(Position::new(self.table_index), self.row, self.column)
    }
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged_operations", &self.staged_operations)
            .field("source_bytes", &self.before.source.len())
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// Immutable source snapshot used as the edit base.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }
    /// Staged consolidation value.
    #[must_use]
    pub fn staged_consolidation(&self) -> Option<&Options> {
        if self.consolidation_staged {
            self.consolidation.as_ref()
        } else {
            match &self.before.catalog.consolidation {
                Some(value) => Some(&value.value),
                None => None,
            }
        }
    }
    /// Staged label range view.
    #[must_use]
    pub fn staged_label_ranges(&self) -> LabelRangesView<'_> {
        let ranges = if self.labels_staged {
            self.labels.as_deref().unwrap_or(&[])
        } else {
            self.before
                .catalog
                .labels
                .as_ref()
                .map_or(&[][..], |value| value.ranges.as_slice())
        };
        LabelRangesView {
            present: self.labels_present,
            ranges,
        }
    }
    /// Whether the semantic ledger is unchanged from its base.
    #[must_use]
    pub fn is_no_op(&self) -> bool {
        (!self.consolidation_staged
            || self.staged_consolidation()
                == self
                    .before
                    .catalog
                    .consolidation
                    .as_ref()
                    .map(|value| &value.value))
            && (!self.labels_staged
                || (self.labels_present == self.before.catalog.labels.is_some()
                    && self.staged_label_ranges().as_slice()
                        == self
                            .before
                            .catalog
                            .labels
                            .as_ref()
                            .map(|value| value.ranges.as_slice())
                            .unwrap_or(&[])))
            && self
                .cell_ops
                .values()
                .all(|op| op.source.is_none() && op.detective.is_none())
    }

    /// Stage singleton consolidation replacement/removal.
    pub fn set_consolidation(&mut self, value: Option<Options>) -> Result<()> {
        if let Some(value) = &value {
            value.validate()?;
            validate_options_limits(value, self.before.limits)?;
        }
        self.ensure_operation_slot()?;
        let staged_memory = value
            .as_ref()
            .map(consolidation_size_bound)
            .transpose()?
            .unwrap_or(0);
        let staged_memory = self.reserve_staging_bytes(staged_memory)?;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.consolidation_staged = true;
        self.consolidation = value;
        Ok(())
    }
    /// Stage singleton consolidation removal.
    pub fn clear_consolidation(&mut self) -> Result<()> {
        self.set_consolidation(None)
    }

    /// Add one ordered label range.
    pub fn add_label_range(&mut self, value: LabelRange) -> Result<()> {
        self.insert_label_range(self.current_labels().len(), value)
    }
    /// Insert one ordered label range.
    pub fn insert_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        value.validate()?;
        validate_label_limits(&value, self.before.limits)?;
        let current = self.current_labels();
        if index > current.len() {
            return Err(Error::InvalidFormat(format!(
                "ODS label-range index {index} is outside {} items",
                current.len()
            )));
        }
        if current.len() >= self.before.limits.max_labels {
            return Err(Error::Unsupported(
                "ODS metadata label range limit exceeded".to_string(),
            ));
        }
        self.ensure_operation_slot()?;
        let mut staged_size = labels_size_bound(std::slice::from_ref(&value))?
            .checked_add(size_of::<LabelRange>())
            .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?;
        if !self.labels_staged {
            staged_size = staged_size
                .checked_add(labels_size_bound(current)?)
                .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?;
        }
        let staged_memory = self.reserve_staging_bytes(staged_size)?;
        if self.labels_staged {
            let operation_memory = self.reserve_operation_memory()?;
            self.labels
                .get_or_insert_with(Vec::new)
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODS metadata label ranges",
                    source,
                })?;
            self.record_operation(operation_memory)?;
            self.retain_staging(staged_memory);
            self.labels
                .as_mut()
                .expect("staged label catalog materialized")
                .insert(index, value);
            self.labels_present = true;
            return Ok(());
        }
        let mut labels = current.to_vec();
        labels
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "ODS metadata label ranges",
                source,
            })?;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        labels.insert(index, value);
        self.labels = Some(labels);
        self.labels_staged = true;
        self.labels_present = true;
        Ok(())
    }
    /// Replace one ordered label range.
    pub fn replace_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        value.validate()?;
        validate_label_limits(&value, self.before.limits)?;
        let current = self.current_labels();
        let label_count = current.len();
        if !self.labels_present {
            return Err(Error::InvalidFormat(
                "ODS label-ranges container is absent".to_string(),
            ));
        }
        if index >= label_count {
            return Err(Error::InvalidFormat(format!(
                "ODS label-range index {index} is outside {label_count} items"
            )));
        }
        self.ensure_operation_slot()?;
        let mut staged_size = labels_size_bound(std::slice::from_ref(&value))?
            .checked_add(size_of::<LabelRange>())
            .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?;
        if !self.labels_staged {
            staged_size = staged_size
                .checked_add(labels_size_bound(current)?)
                .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?;
        }
        let staged_memory = self.reserve_staging_bytes(staged_size)?;
        if self.labels_staged {
            self.bump_operation()?;
            self.retain_staging(staged_memory);
            let labels = self.labels.as_mut().ok_or_else(|| {
                Error::InvalidFormat("ODS label-ranges container is absent".to_string())
            })?;
            labels[index] = value;
            return Ok(());
        }
        let mut labels = current.to_vec();
        labels[index] = value;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.labels = Some(labels);
        self.labels_staged = true;
        Ok(())
    }
    /// Remove one ordered label range.
    pub fn remove_label_range(&mut self, index: usize) -> Result<LabelRange> {
        let current = self.current_labels();
        let label_count = current.len();
        if !self.labels_present {
            return Err(Error::InvalidFormat(
                "ODS label-ranges container is absent".to_string(),
            ));
        }
        if index >= label_count {
            return Err(Error::InvalidFormat(format!(
                "ODS label-range index {index} is outside {label_count} items"
            )));
        }
        self.ensure_operation_slot()?;
        let staged_size = if self.labels_staged {
            0
        } else {
            labels_size_bound(current)?
        };
        let staged_memory = self.reserve_staging_bytes(staged_size)?;
        if self.labels_staged {
            self.bump_operation()?;
            self.retain_staging(staged_memory);
            let labels = self.labels.as_mut().ok_or_else(|| {
                Error::InvalidFormat("ODS label-ranges container is absent".to_string())
            })?;
            return Ok(labels.remove(index));
        }
        let mut labels = current.to_vec();
        let removed = labels.remove(index);
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.labels = Some(labels);
        self.labels_staged = true;
        Ok(removed)
    }
    /// Remove the label-ranges container, including an explicitly empty one.
    pub fn clear_label_ranges(&mut self) -> Result<()> {
        self.ensure_operation_slot()?;
        let staged_memory = self.reserve_staging_bytes(0)?;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.labels = None;
        self.labels_staged = true;
        self.labels_present = false;
        Ok(())
    }
    /// Ensure an explicitly present, possibly empty label-ranges container.
    pub fn ensure_label_ranges(&mut self) -> Result<()> {
        if !self.labels_present {
            self.ensure_operation_slot()?;
            let staged_memory = self.reserve_staging_bytes(0)?;
            self.bump_operation()?;
            self.retain_staging(staged_memory);
            self.labels = Some(Vec::new());
            self.labels_staged = true;
            self.labels_present = true;
            return Ok(());
        }
        if !self.labels_staged {
            self.ensure_operation_slot()?;
            let current = self.current_labels();
            let staged_memory = self.reserve_staging_bytes(labels_size_bound(current)?)?;
            let labels = current.to_vec();
            self.bump_operation()?;
            self.retain_staging(staged_memory);
            self.labels = Some(labels);
            self.labels_staged = true;
            return Ok(());
        }
        self.ensure_operation_slot()?;
        let staged_memory = self.reserve_staging_bytes(0)?;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.labels.get_or_insert_with(Vec::new);
        self.labels_present = true;
        Ok(())
    }
    /// Replace the complete ordered label catalog, retaining presence.
    pub fn replace_label_ranges(&mut self, values: Vec<LabelRange>) -> Result<()> {
        if values.len() > self.before.limits.max_labels {
            return Err(Error::Unsupported(
                "ODS metadata label range limit exceeded".to_string(),
            ));
        }
        for value in &values {
            value.validate()?;
            validate_label_limits(value, self.before.limits)?;
        }
        self.ensure_operation_slot()?;
        let staged_memory = labels_size_bound(&values)?
            .checked_add(
                values
                    .capacity()
                    .saturating_sub(values.len())
                    .checked_mul(size_of::<LabelRange>())
                    .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?,
            )
            .ok_or_else(|| invalid("ODS metadata staged label size overflows"))?;
        let staged_memory = self.reserve_staging_bytes(staged_memory)?;
        self.bump_operation()?;
        self.retain_staging(staged_memory);
        self.labels = Some(values);
        self.labels_staged = true;
        self.labels_present = true;
        Ok(())
    }

    /// Stage a cell-range-source replacement on an existing physical cell.
    pub fn set_cell_range_source(
        &mut self,
        selector: CellSelector<'_>,
        value: CellRange,
    ) -> Result<()> {
        validate_cell_range_limits(&value, self.before.limits)?;
        self.stage_cell_source(selector, Some(value))
    }
    /// Stage a cell-range-source clear on an existing physical cell.
    pub fn clear_cell_range_source(&mut self, selector: CellSelector<'_>) -> Result<()> {
        self.stage_cell_source(selector, None)
    }
    /// Stage detective replacement on an existing physical cell.
    pub fn set_detective(&mut self, selector: CellSelector<'_>, value: Detective) -> Result<()> {
        validate_detective_limits(&value, self.before.limits)?;
        if value.highlighted_ranges().len() > self.before.limits.max_detective_ranges()
            || value.operations().len() > self.before.limits.max_detective_operations()
        {
            return Err(Error::Unsupported(
                "ODS detective collection limit exceeded".to_string(),
            ));
        }
        self.stage_cell_detective(selector, Some(value))
    }
    /// Stage detective owner removal.
    pub fn clear_detective(&mut self, selector: CellSelector<'_>) -> Result<()> {
        self.stage_cell_detective(selector, None)
    }
    /// Edit one detective value in place.
    pub fn edit_detective<F>(&mut self, selector: CellSelector<'_>, update: F) -> Result<()>
    where
        F: FnOnce(&mut Detective) -> Result<()>,
    {
        let mut selector_budget = self.selector_budget();
        self.ensure_cell_target_with_budget(selector, false, &mut selector_budget)?;
        let existing = self.effective_cell_detective_with_budget(selector, &mut selector_budget)?;
        let clone_memory = existing
            .map(detective_size_bound)
            .transpose()?
            .map(|size| self.reserve_staging_bytes(size))
            .transpose()?;
        let mut value = existing.cloned().unwrap_or_else(Detective::new);
        update(&mut value)?;
        validate_detective_limits(&value, self.before.limits)?;
        if value.highlighted_ranges().len() > self.before.limits.max_detective_ranges()
            || value.operations().len() > self.before.limits.max_detective_operations()
        {
            return Err(Error::Unsupported(
                "ODS detective collection limit exceeded".to_string(),
            ));
        }
        let result =
            self.stage_cell_detective_with_budget(selector, Some(value), &mut selector_budget);
        drop(clone_memory);
        result
    }

    fn stage_cell_source(
        &mut self,
        selector: CellSelector<'_>,
        value: Option<CellRange>,
    ) -> Result<()> {
        let mut selector_budget = self.selector_budget();
        self.stage_cell_source_with_budget(selector, value, &mut selector_budget)
    }

    fn stage_cell_source_with_budget(
        &mut self,
        selector: CellSelector<'_>,
        value: Option<CellRange>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<()> {
        self.ensure_cell_target_with_budget(selector, value.is_none(), selector_budget)?;
        let baseline = self.baseline_cell_source_with_budget(selector, selector_budget)?;
        let effective = self.effective_cell_source_with_budget(selector, selector_budget)?;
        if baseline == value.as_ref() {
            self.clear_staged_cell_field(selector, CellField::Source, selector_budget)?;
            return Ok(());
        }
        if effective == value.as_ref() {
            return Ok(());
        }
        self.ensure_operation_slot()?;
        let key_memory = self.reserve_staging_bytes(cell_key_size()?)?;
        let value_memory = if let Some(value) = value.as_ref() {
            Some(self.reserve_staging_bytes(cell_range_size_bound(value)?)?)
        } else {
            None
        };
        let key =
            CellKey::from_selector_with_budget(&self.before.catalog, selector, selector_budget)?;
        self.bump_operation()?;
        self.retain_staging(key_memory);
        if let Some(value_memory) = value_memory {
            self.retain_staging(value_memory);
        }
        self.cell_ops.entry(key).or_default().source = Some(value);
        Ok(())
    }
    fn stage_cell_detective(
        &mut self,
        selector: CellSelector<'_>,
        value: Option<Detective>,
    ) -> Result<()> {
        let mut selector_budget = self.selector_budget();
        self.stage_cell_detective_with_budget(selector, value, &mut selector_budget)
    }

    fn stage_cell_detective_with_budget(
        &mut self,
        selector: CellSelector<'_>,
        value: Option<Detective>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<()> {
        self.ensure_cell_target_with_budget(selector, value.is_none(), selector_budget)?;
        let baseline = self.baseline_cell_detective_with_budget(selector, selector_budget)?;
        let effective = self.effective_cell_detective_with_budget(selector, selector_budget)?;
        if baseline == value.as_ref() {
            self.clear_staged_cell_field(selector, CellField::Detective, selector_budget)?;
            return Ok(());
        }
        if effective == value.as_ref() {
            return Ok(());
        }
        self.ensure_operation_slot()?;
        let key_memory = self.reserve_staging_bytes(cell_key_size()?)?;
        let value_memory = if let Some(value) = value.as_ref() {
            Some(self.reserve_staging_bytes(detective_size_bound(value)?)?)
        } else {
            None
        };
        let key =
            CellKey::from_selector_with_budget(&self.before.catalog, selector, selector_budget)?;
        self.bump_operation()?;
        self.retain_staging(key_memory);
        if let Some(value_memory) = value_memory {
            self.retain_staging(value_memory);
        }
        self.cell_ops.entry(key).or_default().detective = Some(value);
        Ok(())
    }

    fn selector_budget(&self) -> index::SelectorBudget {
        index::SelectorBudget::new(&self.before.context, self.before.limits)
    }

    fn cell_op_with_budget(
        &self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<&CellOp>> {
        let key =
            CellKey::from_selector_with_budget(&self.before.catalog, selector, selector_budget)?;
        Ok(self.cell_ops.get(&key))
    }

    fn baseline_cell_source_with_budget(
        &self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<&CellRange>> {
        Ok(self
            .before
            .cell_metadata_with_budget(selector, selector_budget)?
            .and_then(|view| view.range_source()))
    }

    fn baseline_cell_detective_with_budget(
        &self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<&Detective>> {
        Ok(self
            .before
            .cell_metadata_with_budget(selector, selector_budget)?
            .and_then(|view| view.detective()))
    }

    fn effective_cell_source_with_budget(
        &self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<&CellRange>> {
        if let Some(op) = self.cell_op_with_budget(selector, selector_budget)?
            && let Some(value) = &op.source
        {
            return Ok(value.as_ref());
        }
        self.baseline_cell_source_with_budget(selector, selector_budget)
    }

    fn effective_cell_detective_with_budget(
        &self,
        selector: CellSelector<'_>,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<Option<&Detective>> {
        if let Some(op) = self.cell_op_with_budget(selector, selector_budget)?
            && let Some(value) = &op.detective
        {
            return Ok(value.as_ref());
        }
        self.baseline_cell_detective_with_budget(selector, selector_budget)
    }

    fn clear_staged_cell_field(
        &mut self,
        selector: CellSelector<'_>,
        field: CellField,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<()> {
        let has_field = self
            .cell_op_with_budget(selector, selector_budget)?
            .is_some_and(|op| match field {
                CellField::Source => op.source.is_some(),
                CellField::Detective => op.detective.is_some(),
            });
        if !has_field {
            return Ok(());
        }

        // The key is recreated only after its allocation has been charged. It
        // is a temporary lookup key; the existing map node and its retained
        // value reservation remain owned by the edit until rollback/drop.
        let key_memory = self.reserve_staging_bytes(cell_key_size()?)?;
        let key =
            CellKey::from_selector_with_budget(&self.before.catalog, selector, selector_budget)?;
        if let Some(op) = self.cell_ops.get_mut(&key) {
            match field {
                CellField::Source => op.source = None,
                CellField::Detective => op.detective = None,
            }
        }
        let remove = self
            .cell_ops
            .get(&key)
            .is_some_and(|op| op.source.is_none() && op.detective.is_none());
        if remove {
            self.cell_ops.remove(&key);
        }
        drop(key_memory);
        self.release_staging_if_cell_ledger_empty();
        Ok(())
    }

    fn release_staging_if_cell_ledger_empty(&mut self) {
        if self.cell_ops.is_empty() && !self.consolidation_staged && !self.labels_staged {
            self.staging_memory = None;
        }
    }
    fn ensure_cell_target_with_budget(
        &self,
        selector: CellSelector<'_>,
        allow_missing: bool,
        selector_budget: &mut index::SelectorBudget,
    ) -> Result<()> {
        match self
            .before
            .catalog
            .cell_for_selector_with_budget(&selector, selector_budget)?
        {
            index::PhysicalCell::Missing if allow_missing => Ok(()),
            index::PhysicalCell::Missing => Err(Error::InvalidFormat("ODS metadata cell selector did not match a physical cell".to_string())),
            index::PhysicalCell::ImplicitCovered { .. } => Err(Error::Unsupported("UnsupportedImplicitCoveredCell: ODS metadata cannot materialize an implicit covered cell".to_string())),
            index::PhysicalCell::Stored(index) => {
                let cell = self.before.catalog.cells.get(index).ok_or_else(|| Error::InvalidFormat("ODS metadata cell index disappeared".to_string()))?;
                if !cell.sequence_valid || cell.unsupported_child || cell.direct_owner_count > 0 { return Err(Error::Unsupported("ODS metadata cell owner is opaque or has an ambiguous insertion anchor".to_string())); }
                Ok(())
            },
        }
    }
    fn ensure_operation_slot(&self) -> Result<()> {
        if self.staged_operations >= self.before.limits.max_operations {
            return Err(Error::Unsupported(
                "ODS metadata operation limit exceeded".to_string(),
            ));
        }
        Ok(())
    }

    fn bump_operation(&mut self) -> Result<()> {
        let staged_memory = self.reserve_operation_memory()?;
        self.record_operation(staged_memory)
    }

    fn reserve_operation_memory(&self) -> Result<Reservation> {
        const STAGED_NODE_OVERHEAD: usize = 64;
        self.ensure_operation_slot()?;
        self.reserve_staging_bytes(
            size_of::<CellKey>()
                .checked_add(size_of::<CellOp>())
                .and_then(|value| value.checked_add(STAGED_NODE_OVERHEAD))
                .ok_or_else(|| invalid("ODS metadata staged operation size overflows"))?,
        )
    }

    fn record_operation(&mut self, staged_memory: Reservation) -> Result<()> {
        self.ensure_operation_slot()?;
        self.staged_operations = self.staged_operations.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("ODS metadata operation count overflows".to_string())
        })?;
        self.retain_staging(staged_memory);
        Ok(())
    }

    fn reserve_staging_bytes(&self, amount: usize) -> Result<Reservation> {
        self.before
            .context
            .reserve(
                Resource::Memory,
                u64::try_from(amount)
                    .map_err(|_| invalid("ODS metadata staged memory size overflows"))?,
            )
            .map_err(map_execution)
    }

    fn current_labels(&self) -> &[LabelRange] {
        if self.labels_staged {
            self.labels.as_deref().unwrap_or(&[])
        } else {
            self.before
                .catalog
                .labels
                .as_ref()
                .map_or(&[][..], |value| value.ranges.as_slice())
        }
    }

    fn retain_staging(&mut self, reservation: Reservation) {
        if let Some(existing) = &mut self.staging_memory {
            if let Err(other) = existing.try_merge(reservation) {
                debug_assert!(
                    false,
                    "staged reservations must share the edit context budget"
                );
                drop(other);
            }
        } else {
            self.staging_memory = Some(reservation);
        }
    }

    /// Commit this edit without consuming its staged ledger.
    pub fn commit(&mut self, context: &ExecutionContext) -> Result<Commit> {
        context.check().map_err(map_execution)?;
        if self.before.enforce_context_lineage {
            if context.limits() != self.before.context.limits()
                || !same_budget_lineage(&self.before.context, context)?
            {
                return Err(Error::Unsupported(
                    "ODS metadata commit context does not match the snapshot execution lineage"
                        .to_string(),
                ));
            }
        }
        if self.is_no_op() {
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    source: self.before.clone(),
                    target: self.before.clone(),
                },
                changed: false,
            });
        }
        let rendered = render_candidate(&self.before, self, context)?;
        let target = Snapshot::parse_rendered_with_context(
            rendered,
            self.before.limits,
            context,
            self.before.enforce_context_lineage,
        )?;
        verify_readback(&target, self, context)?;
        Ok(Commit {
            snapshot: target.clone(),
            patch: Patch {
                source: self.before.clone(),
                target,
            },
            changed: true,
        })
    }
    /// Clear every staged operation by restoring the immutable base ledger.
    pub fn rollback(&mut self) {
        self.consolidation = None;
        self.consolidation_staged = false;
        self.labels = None;
        self.labels_staged = false;
        self.labels_present = self.before.catalog.labels.is_some();
        self.cell_ops.clear();
        self.staged_operations = 0;
        self.staging_memory = None;
    }
}

/// Accepted source-checked metadata result.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether content.xml changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
    /// Resulting immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// Exact content.xml patch authorized by one metadata snapshot.
#[derive(Clone, Debug)]
pub struct Patch {
    source: Snapshot,
    target: Snapshot,
}

impl Patch {
    /// Whether source and target are byte-identical.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.source.source.as_ref() != self.target.source.as_ref()
    }
    /// Alias for [`Self::changed`] with the opposite polarity.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.changed()
    }
    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        self.source.source_xml()
    }
    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn target_xml(&self) -> &str {
        self.target.source_xml()
    }
    /// Inverse exact-source patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }
    /// Apply to an exact source snapshot using the retained, fully reopened target.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source.as_ref() != self.source.source.as_ref() {
            return Err(Error::InvalidFormat(
                "ODS sheet metadata patch source snapshot does not match".to_string(),
            ));
        }
        snapshot.context.check().map_err(map_execution)?;
        if !self.changed() {
            return Ok(Commit {
                snapshot: snapshot.clone(),
                patch: self.clone(),
                changed: false,
            });
        }
        let target = self.target.reopen_for_destination(snapshot)?;
        Ok(Commit {
            snapshot: target.clone(),
            patch: Patch {
                source: snapshot.clone(),
                target,
            },
            changed: self.changed(),
        })
    }
}

/// Source-backed immutable metadata snapshot.
pub struct SourceMetadataSnapshot<'source> {
    owner: &'source crate::facade::SourceBackedSpreadsheet,
    inner: Snapshot,
    source_version: SourceVersion,
}

impl fmt::Debug for SourceMetadataSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceMetadataSnapshot")
            .field("source_version", &self.source_version)
            .field("source_bytes", &self.inner.source.len())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for SourceMetadataSnapshot<'source> {
    fn clone(&self) -> Self {
        Self {
            owner: self.owner,
            inner: self.inner.clone(),
            source_version: self.source_version,
        }
    }
}

impl<'source> SourceMetadataSnapshot<'source> {
    pub(crate) fn from_owner(
        owner: &'source crate::facade::SourceBackedSpreadsheet,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        owner.check_source()?;
        let source_version = owner.source_version()?;
        let source = owner.content_xml_arc()?;
        let inner =
            Snapshot::parse_arc_with_context(source, limits, context, enforce_context_lineage)?;
        owner.check_source()?;
        Ok(Self {
            owner,
            inner,
            source_version,
        })
    }
    /// Exact source version captured by this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }
    /// Borrow exact content.xml bytes.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        self.inner.source_xml()
    }
    /// Read consolidation.
    pub fn consolidation(&self) -> Result<Option<&Options>> {
        self.check_source()?;
        self.inner.consolidation()
    }
    /// Read labels.
    pub fn label_ranges(&self) -> Result<LabelRangesView<'_>> {
        self.check_source()?;
        Ok(self.inner.label_ranges())
    }
    /// Read cell metadata.
    pub fn cell_metadata(
        &self,
        selector: CellSelector<'_>,
    ) -> Result<Option<CellMetadataView<'_>>> {
        self.check_source()?;
        self.inner.cell_metadata(selector)
    }
    /// Start source-backed edit.
    pub fn edit(&self) -> Result<SourceMetadataEdit<'source>> {
        self.check_source()?;
        Ok(SourceMetadataEdit {
            before: self.clone(),
            inner: self.inner.edit(),
        })
    }
    fn check_source(&self) -> Result<()> {
        let observed = self.owner.source_version()?;
        if observed == self.source_version {
            Ok(())
        } else {
            Err(Error::SourceChanged {
                expected: self.source_version,
                observed,
            })
        }
    }
    fn is_signed(&self) -> Result<bool> {
        self.owner.sheet_metadata_signed()
    }
}

/// Source-backed semantic metadata edit. Its `commit` method borrows the edit
/// so failed validation leaves the staged ledger and source snapshot intact.
pub struct SourceMetadataEdit<'source> {
    before: SourceMetadataSnapshot<'source>,
    inner: Edit,
}

impl fmt::Debug for SourceMetadataEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceMetadataEdit")
            .field("source_version", &self.before.source_version)
            .finish_non_exhaustive()
    }
}

impl<'source> SourceMetadataEdit<'source> {
    /// Base source snapshot.
    #[must_use]
    pub const fn before(&self) -> &SourceMetadataSnapshot<'source> {
        &self.before
    }
    /// Stage consolidation replacement/removal.
    pub fn set_consolidation(&mut self, value: Option<Options>) -> Result<()> {
        self.inner.set_consolidation(value)
    }
    /// Clear consolidation.
    pub fn clear_consolidation(&mut self) -> Result<()> {
        self.inner.clear_consolidation()
    }
    /// Stage label insertion.
    pub fn add_label_range(&mut self, value: LabelRange) -> Result<()> {
        self.inner.add_label_range(value)
    }
    /// Stage label insertion at a source-order index.
    pub fn insert_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        self.inner.insert_label_range(index, value)
    }
    /// Stage label replacement.
    pub fn replace_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        self.inner.replace_label_range(index, value)
    }
    /// Stage label removal.
    pub fn remove_label_range(&mut self, index: usize) -> Result<LabelRange> {
        self.inner.remove_label_range(index)
    }
    /// Clear labels and their container.
    pub fn clear_label_ranges(&mut self) -> Result<()> {
        self.inner.clear_label_ranges()
    }
    /// Ensure labels container presence.
    pub fn ensure_label_ranges(&mut self) -> Result<()> {
        self.inner.ensure_label_ranges()
    }
    /// Set an existing cell source owner.
    pub fn set_cell_range_source(
        &mut self,
        selector: CellSelector<'_>,
        value: CellRange,
    ) -> Result<()> {
        self.inner.set_cell_range_source(selector, value)
    }
    /// Clear an existing cell source owner.
    pub fn clear_cell_range_source(&mut self, selector: CellSelector<'_>) -> Result<()> {
        self.inner.clear_cell_range_source(selector)
    }
    /// Set an existing cell detective owner.
    pub fn set_detective(&mut self, selector: CellSelector<'_>, value: Detective) -> Result<()> {
        self.inner.set_detective(selector, value)
    }
    /// Clear an existing cell detective owner.
    pub fn clear_detective(&mut self, selector: CellSelector<'_>) -> Result<()> {
        self.inner.clear_detective(selector)
    }
    /// Edit an existing cell detective owner.
    pub fn edit_detective<F>(&mut self, selector: CellSelector<'_>, update: F) -> Result<()>
    where
        F: FnOnce(&mut Detective) -> Result<()>,
    {
        self.inner.edit_detective(selector, update)
    }
    /// Commit against the retained source version without consuming this edit.
    pub fn commit(&mut self, context: &ExecutionContext) -> Result<SourceMetadataCommit<'source>> {
        self.before.check_source()?;
        let commit = self.inner.commit(context)?;
        if commit.changed() && self.before.is_signed()? {
            return Err(Error::Unsupported("signed-source refusal: changed ODS sheet metadata requires explicit unsign/resign policy".to_string()));
        }
        self.before.check_source()?;
        Ok(SourceMetadataCommit {
            snapshot: SourceMetadataSnapshot {
                owner: self.before.owner,
                inner: commit.snapshot.clone(),
                source_version: self.before.source_version,
            },
            patch: SourceMetadataPatch {
                before: self.before.clone(),
                target_inner: commit.snapshot.clone(),
            },
            changed: commit.changed(),
        })
    }
}

/// Source-backed exact content patch and accepted target snapshot.
pub struct SourceMetadataCommit<'source> {
    snapshot: SourceMetadataSnapshot<'source>,
    patch: SourceMetadataPatch<'source>,
    changed: bool,
}

impl fmt::Debug for SourceMetadataCommit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceMetadataCommit")
            .field("changed", &self.changed)
            .field("source_bytes", &self.snapshot.inner.source.len())
            .finish_non_exhaustive()
    }
}

impl<'source> SourceMetadataCommit<'source> {
    /// Candidate source-backed metadata snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceMetadataSnapshot<'source> {
        &self.snapshot
    }
    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SourceMetadataPatch<'source> {
        &self.patch
    }
    /// Whether content.xml changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
    /// Publish the complete accepted ODS package to a sequential sink.
    pub fn write_to<W: Write>(
        &self,
        writer: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        if let Err(error) = self.snapshot.check_source() {
            return Err(match error {
                Error::SourceChanged { expected, observed } => {
                    SourceContentPublicationError::SourceChanged {
                        expected,
                        observed,
                        progress:
                            litchi_odf_common::core::SourceContentPublicationProgress::Untouched,
                    }
                },
                other => SourceContentPublicationError::Core(other),
            });
        }
        if self.changed
            && self
                .snapshot
                .is_signed()
                .map_err(SourceContentPublicationError::Core)?
        {
            return Err(SourceContentPublicationError::Core(Error::Unsupported(
                "signed-source refusal: changed ODS sheet metadata cannot be published".to_string(),
            )));
        }
        self.snapshot.owner.write_sheet_metadata_content(
            writer,
            self.snapshot.source_xml().as_bytes(),
            options,
        )
    }
    /// Publish with default finite source-content options.
    pub fn write_to_default<W: Write>(
        &self,
        writer: W,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        self.write_to(writer, SourceContentPublicationOptions::new())
    }
}

/// Source-backed exact patch.
pub struct SourceMetadataPatch<'source> {
    before: SourceMetadataSnapshot<'source>,
    target_inner: Snapshot,
}

impl fmt::Debug for SourceMetadataPatch<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceMetadataPatch")
            .field("changed", &self.changed())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for SourceMetadataPatch<'source> {
    fn clone(&self) -> Self {
        Self {
            before: self.before.clone(),
            target_inner: self.target_inner.clone(),
        }
    }
}

impl<'source> SourceMetadataPatch<'source> {
    /// Exact source content.xml bytes authorized by this patch.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        self.before.source_xml()
    }
    /// Exact target content.xml bytes produced by this patch.
    #[must_use]
    pub fn target_xml(&self) -> &str {
        self.target_inner.source_xml()
    }
    /// Whether target bytes differ.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.before.source_xml() != self.target_inner.source_xml()
    }
    /// Alias for changed.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.changed()
    }
    /// Inverse source-authorized patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: SourceMetadataSnapshot {
                owner: self.before.owner,
                inner: self.target_inner.clone(),
                source_version: self.before.source_version,
            },
            target_inner: self.before.inner.clone(),
        }
    }
    /// Apply against the exact source-backed snapshot.
    pub fn apply(
        &self,
        snapshot: &SourceMetadataSnapshot<'source>,
    ) -> Result<SourceMetadataCommit<'source>> {
        snapshot.check_source()?;
        snapshot.inner.context.check().map_err(map_execution)?;
        if !std::ptr::eq(snapshot.owner, self.before.owner)
            || snapshot.source_version != self.before.source_version
            || snapshot.source_xml() != self.source_xml()
        {
            return Err(Error::InvalidFormat(
                "ODS source metadata patch source snapshot does not match".to_string(),
            ));
        }
        if self.changed() && snapshot.is_signed()? {
            return Err(Error::Unsupported(
                "signed-source refusal: changed ODS sheet metadata patch cannot be applied"
                    .to_string(),
            ));
        }
        if !self.changed() {
            return Ok(SourceMetadataCommit {
                snapshot: snapshot.clone(),
                patch: self.clone(),
                changed: false,
            });
        }
        // The accepted target was fully reopened during the originating
        // commit and is retained by this patch.  A source-backed owner may
        // still expose multiple snapshots with different destination
        // profiles, so revalidate the exact target under the destination
        // context before handing it back.  This also keeps the returned
        // reversible patch bound to that destination snapshot.
        let inner = self.target_inner.reopen_for_destination(&snapshot.inner)?;
        let patch_target = inner.clone();
        snapshot.check_source()?;
        Ok(SourceMetadataCommit {
            snapshot: SourceMetadataSnapshot {
                owner: snapshot.owner,
                inner,
                source_version: snapshot.source_version,
            },
            patch: SourceMetadataPatch {
                before: snapshot.clone(),
                target_inner: patch_target,
            },
            changed: self.changed(),
        })
    }
}

pub(crate) fn detective_limits(limits: Limits) -> Result<detective::Limits> {
    detective::Limits::new(
        limits.max_input_bytes(),
        limits.max_output_bytes(),
        limits.max_depth(),
        limits.max_events(),
        limits.max_work_units(),
    )?
    .with_item_limits(
        limits.max_cells().min(detective::MAX_ITEMS),
        limits.max_detective_ranges(),
        limits.max_detective_operations(),
    )?
    .with_scalar_limits(limits.max_text_bytes(), limits.max_namespace_bindings())
    .and_then(|value| value.with_scratch_bytes(limits.max_scratch_bytes()))
}

/// Preflight every owned buffer that is created while assembling a candidate.
///
/// OutputBytes covers the final source candidate. This reservation covers
/// the replacement list, rendered owner fragments, repeated-cell/row split
/// buffers, and the old fragments that remain live while a splice is built.
/// Keeping this reservation alive for the whole render makes a failed budget
/// check happen before any replacement string is allocated.
fn preflight_render_memory(
    before: &Snapshot,
    edit: &Edit,
    context: &ExecutionContext,
) -> Result<Reservation> {
    let replacement_capacity = edit
        .cell_ops
        .len()
        .checked_add(2)
        .ok_or_else(|| invalid("ODS metadata replacement count overflows"))?;
    let target_capacity = edit.cell_ops.len();
    let mut temporary = replacement_capacity
        .checked_mul(size_of::<Replacement>())
        .and_then(|value| {
            value.checked_add(
                target_capacity.checked_mul(
                    size_of::<CellTarget<'_>>()
                        .saturating_add(size_of::<CellVariant>())
                        .saturating_add(size_of::<RowVariant>())
                        .saturating_add(size_of::<CellFragment>()),
                )?,
            )
        })
        .ok_or_else(|| invalid("ODS metadata replacement list size overflows"))?;

    let old_consolidation = before
        .catalog
        .consolidation
        .as_ref()
        .map(|value| &value.value);
    if old_consolidation != edit.staged_consolidation()
        && let Some(value) = edit.staged_consolidation()
    {
        add_temporary(&mut temporary, consolidation_size_bound(value)?, 2)?;
    }

    let labels_changed = edit.labels_present != before.catalog.labels.is_some()
        || edit.current_labels()
            != before
                .catalog
                .labels
                .as_ref()
                .map(|value| value.ranges.as_slice())
                .unwrap_or(&[]);
    if labels_changed && edit.labels_present {
        add_temporary(&mut temporary, labels_size_bound(edit.current_labels())?, 2)?;
    }

    let mut selector_budget = index::SelectorBudget::new(context, before.limits);
    for (key, op) in &edit.cell_ops {
        render_step(context, 1)?;
        if op.source.is_none() && op.detective.is_none() {
            continue;
        }
        let selector = key.selector();
        let physical = before
            .catalog
            .cell_for_selector_with_budget(&selector, &mut selector_budget)?;
        let index::PhysicalCell::Stored(cell_index) = physical else {
            return Err(Error::Unsupported(
                "UnsupportedImplicitCoveredCell: metadata edit has no physical owner".to_string(),
            ));
        };
        let cell = before
            .catalog
            .cells
            .get(cell_index)
            .ok_or_else(|| invalid("ODS metadata cell index disappeared"))?;
        if !cell.sequence_valid || cell.unsupported_child || cell.direct_owner_count > 0 {
            return Err(Error::Unsupported(
                "ODS metadata cell owner is opaque or has an ambiguous insertion anchor"
                    .to_string(),
            ));
        }
        let span = before
            .catalog
            .spans
            .get(cell.span)
            .ok_or_else(|| invalid("ODS metadata cell span disappeared"))?;
        add_temporary(&mut temporary, span.range.len(), 5)?;
        let mut rendered = 0usize;
        if let Some(Some(value)) = &op.source {
            rendered = rendered
                .checked_add(cell_range_size_bound(value)?)
                .ok_or_else(|| invalid("ODS metadata rendered size overflows"))?;
        }
        if let Some(Some(value)) = &op.detective {
            rendered = rendered
                .checked_add(detective_size_bound(value)?)
                .ok_or_else(|| invalid("ODS metadata rendered size overflows"))?;
        }
        add_temporary(&mut temporary, rendered, 6)?;
        if op.source.is_some() && !before.source.contains("xmlns:xlink=") {
            add_temporary(&mut temporary, XLINK_NAMESPACE_ATTRIBUTE.len(), 2)?;
        }
        if cell.row_repeat > 1 {
            let row = before
                .catalog
                .rows
                .get(cell.row)
                .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
            let row_span = before
                .catalog
                .spans
                .get(row.span)
                .ok_or_else(|| invalid("ODS metadata row source span disappeared"))?;
            add_temporary(&mut temporary, row_span.range.len(), 5)?;
        }
    }

    let temporary = temporary.max(1);
    if temporary > before.limits.max_scratch_bytes() {
        return Err(limit(
            "scratch bytes",
            temporary,
            before.limits.max_scratch_bytes(),
        ));
    }
    context
        .reserve(
            Resource::Memory,
            u64::try_from(temporary)
                .map_err(|_| invalid("ODS metadata temporary memory size overflows"))?,
        )
        .map_err(map_execution)
}

const XLINK_NAMESPACE_ATTRIBUTE: &str = " xmlns:xlink=\"http://www.w3.org/1999/xlink\"";

fn add_temporary(total: &mut usize, value: usize, factor: usize) -> Result<()> {
    let amount = value
        .checked_mul(factor)
        .ok_or_else(|| invalid("ODS metadata temporary memory size overflows"))?;
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid("ODS metadata temporary memory size overflows"))?;
    Ok(())
}

fn bounded_text_size(value: &str) -> Result<usize> {
    value
        .len()
        .checked_mul(6)
        .and_then(|value| value.checked_add(64))
        .ok_or_else(|| invalid("ODS metadata rendered text size overflows"))
}

fn consolidation_size_bound(value: &Options) -> Result<usize> {
    let mut size = size_of::<Options>();
    size = size
        .checked_add(
            value
                .source_cell_range_addresses
                .capacity()
                .checked_mul(size_of::<String>())
                .ok_or_else(|| invalid("ODS consolidation size overflows"))?,
        )
        .ok_or_else(|| invalid("ODS consolidation size overflows"))?;
    size = size
        .checked_add(bounded_text_size(&value.function)?)
        .ok_or_else(|| invalid("ODS consolidation size overflows"))?;
    size = size
        .checked_add(bounded_text_size(&value.target_cell_address)?)
        .and_then(|size| size.checked_add(256))
        .ok_or_else(|| invalid("ODS consolidation size overflows"))?;
    for address in &value.source_cell_range_addresses {
        size = size
            .checked_add(bounded_text_size(address)?)
            .ok_or_else(|| invalid("ODS consolidation size overflows"))?;
    }
    size.checked_add(128)
        .ok_or_else(|| invalid("ODS consolidation size overflows"))
}

fn labels_size_bound(values: &[LabelRange]) -> Result<usize> {
    let mut size = size_of::<Vec<LabelRange>>()
        .checked_add(
            values
                .len()
                .checked_mul(size_of::<LabelRange>())
                .ok_or_else(|| invalid("ODS label-range size overflows"))?,
        )
        .and_then(|size| size.checked_add(128))
        .ok_or_else(|| invalid("ODS label-range size overflows"))?;
    for range in values {
        size = size
            .checked_add(bounded_text_size(&range.label_cell_range_address)?)
            .and_then(|value| {
                value.checked_add(bounded_text_size(&range.data_cell_range_address).ok()?)
            })
            .and_then(|value| value.checked_add(128))
            .ok_or_else(|| invalid("ODS label-range size overflows"))?;
    }
    Ok(size)
}

fn cell_range_size_bound(value: &CellRange) -> Result<usize> {
    let mut size = size_of::<CellRange>();
    size = size
        .checked_add(bounded_text_size(value.name())?)
        .ok_or_else(|| invalid("ODS cell-range-source size overflows"))?;
    size = size
        .checked_add(bounded_text_size(value.href())?)
        .and_then(|size| size.checked_add(256))
        .ok_or_else(|| invalid("ODS cell-range-source size overflows"))?;
    for text in [
        value.filter_name(),
        value.filter_options(),
        value.refresh_delay(),
    ]
    .into_iter()
    .flatten()
    {
        size = size
            .checked_add(bounded_text_size(text)?)
            .ok_or_else(|| invalid("ODS cell-range-source size overflows"))?;
    }
    Ok(size)
}

fn cell_key_size() -> Result<usize> {
    const STAGED_NODE_OVERHEAD: usize = 64;
    size_of::<CellKey>()
        .checked_add(STAGED_NODE_OVERHEAD)
        .ok_or_else(|| invalid("ODS metadata cell selector size overflows"))
}

fn detective_size_bound(value: &Detective) -> Result<usize> {
    let range_bytes = value
        .highlighted_ranges()
        .len()
        .checked_mul(size_of::<HighlightedRange>())
        .ok_or_else(|| invalid("ODS detective size overflows"))?;
    let operation_bytes = value
        .operations()
        .len()
        .checked_mul(size_of::<Operation>())
        .ok_or_else(|| invalid("ODS detective size overflows"))?;
    let mut size = size_of::<Detective>()
        .checked_add(range_bytes)
        .and_then(|size| size.checked_add(operation_bytes))
        .and_then(|size| size.checked_add(256))
        .ok_or_else(|| invalid("ODS detective size overflows"))?;
    for range in value.highlighted_ranges() {
        if let Some(address) = range.cell_range_address() {
            size = size
                .checked_add(bounded_text_size(address)?)
                .ok_or_else(|| invalid("ODS detective size overflows"))?;
        }
        size = size
            .checked_add(128)
            .ok_or_else(|| invalid("ODS detective size overflows"))?;
    }
    size.checked_add(
        value
            .operations()
            .len()
            .checked_mul(128)
            .ok_or_else(|| invalid("ODS detective size overflows"))?,
    )
    .ok_or_else(|| invalid("ODS detective size overflows"))
}

fn render_candidate(
    before: &Snapshot,
    edit: &Edit,
    context: &ExecutionContext,
) -> Result<RenderedCandidate> {
    let temporary_memory = preflight_render_memory(before, edit, context)?;
    let mut operations = Vec::<Replacement>::new();
    operations
        .try_reserve_exact(edit.cell_ops.len().saturating_add(2))
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata replacement list",
            source,
        })?;
    let catalog = &before.catalog;
    let mut selector_budget = index::SelectorBudget::new(context, before.limits);
    let old_consolidation = catalog.consolidation.as_ref().map(|value| &value.value);
    if old_consolidation != edit.staged_consolidation() {
        render_step(context, 1)?;
        let replacement = match (&catalog.consolidation, edit.staged_consolidation()) {
            (Some(record), Some(value)) => {
                if !record.canonical {
                    return Err(Error::Unsupported(
                        "ODS consolidation owner requires lexical preservation".to_string(),
                    ));
                }
                render_consolidation(
                    value,
                    qname_prefix(&catalog.spans[record.span], "consolidation"),
                )?
            },
            (Some(record), None) => {
                if !record.canonical {
                    return Err(Error::Unsupported(
                        "ODS consolidation owner requires lexical preservation".to_string(),
                    ));
                }
                String::new()
            },
            (None, Some(value)) => {
                let rendered = render_consolidation(value, table_prefix(before))?;
                render_top_level_insertion(before, "consolidation", &rendered)?
            },
            (None, None) => String::new(),
        };
        if let Some(record) = &catalog.consolidation {
            operations.push(Replacement {
                range: catalog.spans[record.span].range.clone(),
                value: replacement,
                order: 0,
            });
        } else if edit.staged_consolidation().is_some() {
            operations.push(Replacement {
                range: top_level_anchor(before, "consolidation")?,
                value: replacement,
                order: 0,
            });
        }
    }
    let labels_changed = edit.labels_present != catalog.labels.is_some()
        || edit.current_labels()
            != catalog
                .labels
                .as_ref()
                .map(|value| value.ranges.as_slice())
                .unwrap_or(&[]);
    if labels_changed {
        render_step(context, 1)?;
        if let Some(record) = &catalog.labels {
            if !record.canonical {
                return Err(Error::Unsupported(
                    "ODS label-ranges owner requires lexical preservation".to_string(),
                ));
            }
        }
        let replacement = if edit.labels_present {
            render_labels(
                edit.current_labels(),
                catalog
                    .labels
                    .as_ref()
                    .map(|value| qname_prefix(&catalog.spans[value.span], "label-ranges"))
                    .unwrap_or_else(|| table_prefix(before)),
            )?
        } else {
            String::new()
        };
        if let Some(record) = &catalog.labels {
            operations.push(Replacement {
                range: catalog.spans[record.span].range.clone(),
                value: replacement,
                order: 1,
            });
        } else if edit.labels_present {
            operations.push(Replacement {
                range: top_level_anchor(before, "label-ranges")?,
                value: replacement,
                order: 1,
            });
        }
    }
    let mut targets = Vec::<CellTarget<'_>>::new();
    targets
        .try_reserve_exact(edit.cell_ops.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata cell target list",
            source,
        })?;
    for (key, op) in &edit.cell_ops {
        render_step(context, 1)?;
        if op.source.is_none() && op.detective.is_none() {
            continue;
        }
        let selector = key.selector();
        let physical = catalog.cell_for_selector_with_budget(&selector, &mut selector_budget)?;
        let index::PhysicalCell::Stored(cell_index) = physical else {
            return Err(Error::Unsupported(
                "UnsupportedImplicitCoveredCell: metadata edit has no physical owner".to_string(),
            ));
        };
        let cell = catalog
            .cells
            .get(cell_index)
            .ok_or_else(|| invalid("ODS metadata cell index disappeared"))?;
        if !cell.sequence_valid || cell.unsupported_child || cell.direct_owner_count > 0 {
            return Err(Error::Unsupported(
                "ODS metadata cell owner is opaque or has an ambiguous insertion anchor"
                    .to_string(),
            ));
        }
        let row = catalog
            .rows
            .get(cell.row)
            .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
        let row_offset = selector
            .row()
            .checked_sub(row.logical_start)
            .ok_or_else(|| invalid("ODS metadata selected row precedes repeated row"))?;
        let column_offset = selector
            .column()
            .checked_sub(cell.column_start)
            .ok_or_else(|| invalid("ODS metadata selected column precedes repeated cell"))?;
        targets.push(CellTarget {
            op,
            cell_index,
            row_index: cell.row,
            row_offset,
            column_offset,
        });
    }
    targets.sort_by(|left, right| {
        left.row_index
            .cmp(&right.row_index)
            .then(left.row_offset.cmp(&right.row_offset))
            .then(left.cell_index.cmp(&right.cell_index))
            .then(left.column_offset.cmp(&right.column_offset))
    });
    let mut target_cursor = 0usize;
    while target_cursor < targets.len() {
        render_step(context, 1)?;
        let row_index = targets[target_cursor].row_index;
        let row = catalog
            .rows
            .get(row_index)
            .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
        let mut row_end = target_cursor + 1;
        while row_end < targets.len() && targets[row_end].row_index == row_index {
            row_end += 1;
        }
        if row.repeat > 1 {
            let row_span = catalog
                .spans
                .get(row.span)
                .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
            operations.push(Replacement {
                range: row_span.range.clone(),
                value: render_repeated_row_change(
                    before,
                    row_index,
                    &targets[target_cursor..row_end],
                    context,
                )?,
                order: 2,
            });
        } else {
            let mut cell_cursor = target_cursor;
            while cell_cursor < row_end {
                render_step(context, 1)?;
                let cell_index = targets[cell_cursor].cell_index;
                let mut cell_end = cell_cursor + 1;
                while cell_end < row_end && targets[cell_end].cell_index == cell_index {
                    cell_end += 1;
                }
                let cell = catalog
                    .cells
                    .get(cell_index)
                    .ok_or_else(|| invalid("ODS metadata cell index disappeared"))?;
                let span = catalog
                    .spans
                    .get(cell.span)
                    .ok_or_else(|| invalid("ODS metadata cell span disappeared"))?;
                operations.push(Replacement {
                    range: span.range.clone(),
                    value: render_repeated_cell_change(
                        before,
                        cell,
                        &targets[cell_cursor..cell_end],
                        context,
                    )?,
                    order: 2,
                });
                cell_cursor = cell_end;
            }
        }
        target_cursor = row_end;
    }
    if operations.is_empty() {
        let output_len = before.source.len();
        if output_len > before.limits.max_output_bytes() {
            return Err(limit(
                "output bytes",
                output_len,
                before.limits.max_output_bytes(),
            ));
        }
        context
            .consume(
                Resource::Work,
                u64::try_from(output_len)
                    .map_err(|_| invalid("ODS metadata output size overflows"))?,
            )
            .map_err(map_execution)?;
        let output = context
            .reserve(
                Resource::OutputBytes,
                u64::try_from(output_len)
                    .map_err(|_| invalid("ODS metadata output size overflows"))?,
            )
            .map_err(map_execution)?;
        let memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(output_len)
                    .map_err(|_| invalid("ODS metadata output size overflows"))?,
            )
            .map_err(map_execution)?;
        let xml = copy_source(before.source.as_ref(), before.limits.max_output_bytes())?;
        drop(temporary_memory);
        return retain_rendered_candidate(xml, output_len, memory, output, context);
    }
    operations.sort_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then(left.order.cmp(&right.order))
    });
    let mut previous_end = 0usize;
    for operation in &operations {
        if operation.range.start < previous_end {
            return Err(Error::Unsupported(
                "ODS metadata operations overlap".to_string(),
            ));
        }
        previous_end = operation.range.end;
    }
    let output_len = operations
        .iter()
        .try_fold(before.source.len(), |length, operation| {
            length
                .checked_sub(operation.range.end - operation.range.start)
                .and_then(|value| value.checked_add(operation.value.len()))
                .ok_or_else(|| invalid("ODS metadata candidate length overflows"))
        })?;
    if output_len > before.limits.max_output_bytes {
        return Err(limit(
            "output bytes",
            output_len,
            before.limits.max_output_bytes,
        ));
    }
    context
        .consume(
            Resource::Work,
            u64::try_from(output_len).map_err(|_| invalid("ODS metadata output size overflows"))?,
        )
        .map_err(map_execution)?;
    let output_budget = context
        .reserve(
            Resource::OutputBytes,
            u64::try_from(output_len).map_err(|_| invalid("ODS metadata output size overflows"))?,
        )
        .map_err(map_execution)?;
    let candidate_memory = context
        .reserve(
            Resource::Memory,
            u64::try_from(output_len).map_err(|_| invalid("ODS metadata output size overflows"))?,
        )
        .map_err(map_execution)?;
    let replacement_bytes = operations.iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value.value.len())
            .ok_or_else(|| invalid("ODS metadata replacement scratch size overflows"))
    })?;
    if replacement_bytes > before.limits.max_scratch_bytes {
        return Err(limit(
            "scratch bytes",
            replacement_bytes,
            before.limits.max_scratch_bytes,
        ));
    }
    let mut output = String::new();
    render_step(context, 1)?;
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata candidate output",
            source,
        })?;
    let mut cursor = 0usize;
    for operation in operations {
        render_step(context, 1)?;
        output.push_str(
            before
                .source
                .get(cursor..operation.range.start)
                .ok_or_else(|| invalid("ODS metadata replacement span is invalid"))?,
        );
        output.push_str(&operation.value);
        cursor = operation.range.end;
    }
    output.push_str(
        before
            .source
            .get(cursor..)
            .ok_or_else(|| invalid("ODS metadata replacement suffix is invalid"))?,
    );
    drop(temporary_memory);
    retain_rendered_candidate(output, output_len, candidate_memory, output_budget, context)
}

fn retain_rendered_candidate(
    xml: String,
    output_len: usize,
    memory: Reservation,
    output: Reservation,
    context: &ExecutionContext,
) -> Result<RenderedCandidate> {
    let arc_memory = output_len
        .checked_add(size_of::<String>())
        .ok_or_else(|| invalid("ODS metadata rendered Arc size overflows"))?;
    let arc_memory = context
        .reserve(
            Resource::Memory,
            u64::try_from(arc_memory)
                .map_err(|_| invalid("ODS metadata rendered Arc size overflows"))?,
        )
        .map_err(map_execution)?;
    let xml = Arc::from(xml);
    drop(arc_memory);
    Ok(RenderedCandidate {
        xml,
        memory: Arc::new(memory),
        output: Arc::new(output),
    })
}

struct RenderedCandidate {
    xml: Arc<str>,
    memory: Arc<Reservation>,
    output: Arc<Reservation>,
}

struct Replacement {
    range: ByteRange<usize>,
    value: String,
    order: u8,
}

struct CellTarget<'a> {
    op: &'a CellOp,
    cell_index: usize,
    row_index: usize,
    row_offset: usize,
    column_offset: usize,
}

struct CellFragment {
    range: ByteRange<usize>,
    value: String,
}

struct CellVariant {
    offset: usize,
    value: String,
}

struct RowVariant {
    offset: usize,
    value: String,
}

fn verify_readback(target: &Snapshot, edit: &Edit, context: &ExecutionContext) -> Result<()> {
    if target.consolidation()?.cloned() != edit.staged_consolidation().cloned() {
        return Err(invalid(
            "ODS sheet metadata consolidation readback differs from staged value",
        ));
    }
    if target.label_ranges().present() != edit.labels_present
        || target.label_ranges().as_slice() != edit.current_labels()
    {
        return Err(invalid(
            "ODS sheet metadata label-range readback differs from staged value",
        ));
    }
    let mut selector_budget = index::SelectorBudget::new(context, target.limits);
    for (key, op) in &edit.cell_ops {
        let view = target
            .cell_metadata_with_budget(key.selector(), &mut selector_budget)?
            .ok_or_else(|| invalid("ODS metadata target cell disappeared during readback"))?;
        if let Some(source) = &op.source {
            if view.range_source() != source.as_ref() {
                return Err(invalid(
                    "ODS cell-range-source readback differs from staged value",
                ));
            }
        }
        if let Some(detective) = &op.detective {
            if view.detective() != detective.as_ref() {
                return Err(invalid("ODS detective readback differs from staged value"));
            }
        }
    }
    Ok(())
}

fn effective_cell_op(cell: &index::CellRecord, op: &CellOp) -> CellOp {
    let mut effective = op.clone();
    if effective
        .source
        .as_ref()
        .is_some_and(|value| cell.source.as_ref() == value.as_ref())
    {
        effective.source = None;
    }
    if effective
        .detective
        .as_ref()
        .is_some_and(|value| cell.detective.as_ref() == value.as_ref())
    {
        effective.detective = None;
    }
    effective
}

fn render_repeated_cell_change(
    before: &Snapshot,
    cell: &index::CellRecord,
    targets: &[CellTarget<'_>],
    context: &ExecutionContext,
) -> Result<String> {
    let span = before
        .catalog
        .spans
        .get(cell.span)
        .ok_or_else(|| invalid("ODS metadata cell span disappeared"))?;
    let raw = before
        .source
        .get(span.range.clone())
        .ok_or_else(|| invalid("ODS metadata cell source range is invalid"))?;
    if cell.column_repeat <= 1 {
        if targets.len() > 1 {
            return Err(Error::Unsupported(
                "ODS metadata has multiple edits for one logical cell".to_string(),
            ));
        }
        let target = targets
            .first()
            .ok_or_else(|| invalid("ODS metadata cell target disappeared"))?;
        let effective = effective_cell_op(cell, target.op);
        if effective.source.is_none() && effective.detective.is_none() {
            return copy_fragment(raw, "ODS metadata unchanged cell");
        }
        return render_cell_direct(before, cell, &effective, span, raw, context);
    }

    let mut variants = Vec::<CellVariant>::new();
    variants
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata repeated-cell variants",
            source,
        })?;
    for target in targets {
        render_step(context, 1)?;
        if target.column_offset >= cell.column_repeat {
            return Err(invalid(
                "ODS metadata selected column is outside repeated cell",
            ));
        }
        let effective = effective_cell_op(cell, target.op);
        if effective.source.is_none() && effective.detective.is_none() {
            continue;
        }
        if variants
            .last()
            .is_some_and(|previous| previous.offset == target.column_offset)
        {
            return Err(Error::Unsupported(
                "ODS metadata has multiple edits for one logical repeated cell".to_string(),
            ));
        }
        variants.push(CellVariant {
            offset: target.column_offset,
            value: render_cell_direct(before, cell, &effective, span, raw, context)?,
        });
    }
    if variants.is_empty() {
        return copy_fragment(raw, "ODS metadata unchanged repeated cell");
    }

    let mut output_len = 0usize;
    let mut previous = 0usize;
    for variant in &variants {
        render_step(context, 1)?;
        let gap = variant
            .offset
            .checked_sub(previous)
            .ok_or_else(|| invalid("ODS metadata repeated-cell target order is invalid"))?;
        if gap > 0 {
            output_len = output_len
                .checked_add(replace_repeat_attr_len(
                    raw,
                    span,
                    "number-columns-repeated",
                    gap,
                )?)
                .ok_or_else(|| invalid("ODS repeated cell split size overflows"))?;
        }
        output_len = output_len
            .checked_add(replace_repeat_attr_len(
                &variant.value,
                span,
                "number-columns-repeated",
                1,
            )?)
            .ok_or_else(|| invalid("ODS repeated cell split size overflows"))?;
        previous = variant.offset + 1;
    }
    if previous < cell.column_repeat {
        output_len = output_len
            .checked_add(replace_repeat_attr_len(
                raw,
                span,
                "number-columns-repeated",
                cell.column_repeat - previous,
            )?)
            .ok_or_else(|| invalid("ODS repeated cell split size overflows"))?;
    }
    let mut output = String::new();
    render_step(context, 1)?;
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata repeated-cell split",
            source,
        })?;
    previous = 0;
    for variant in variants {
        render_step(context, 1)?;
        let gap = variant
            .offset
            .checked_sub(previous)
            .ok_or_else(|| invalid("ODS metadata repeated-cell target order is invalid"))?;
        if gap > 0 {
            output.push_str(&clone_repeat(raw, span, "number-columns-repeated", gap)?);
        }
        output.push_str(&replace_repeat_attr(
            &variant.value,
            span,
            "number-columns-repeated",
            1,
        )?);
        previous = variant.offset + 1;
    }
    if previous < cell.column_repeat {
        render_step(context, 1)?;
        output.push_str(&clone_repeat(
            raw,
            span,
            "number-columns-repeated",
            cell.column_repeat - previous,
        )?);
    }
    if output.len() != output_len {
        return Err(invalid(
            "ODS repeated-cell output disagrees with its preflight",
        ));
    }
    Ok(output)
}

fn render_repeated_row_change(
    before: &Snapshot,
    row_index: usize,
    targets: &[CellTarget<'_>],
    context: &ExecutionContext,
) -> Result<String> {
    let row = before
        .catalog
        .rows
        .get(row_index)
        .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
    let row_span = before
        .catalog
        .spans
        .get(row.span)
        .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
    let row_raw = before
        .source
        .get(row_span.range.clone())
        .ok_or_else(|| invalid("ODS metadata row source range is invalid"))?;
    let mut variants = Vec::<RowVariant>::new();
    variants
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata repeated-row variants",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < targets.len() {
        render_step(context, 1)?;
        let offset = targets[cursor].row_offset;
        if offset >= row.repeat {
            return Err(invalid("ODS metadata selected row is outside repeated row"));
        }
        let mut end = cursor + 1;
        while end < targets.len() && targets[end].row_offset == offset {
            end += 1;
        }
        let value = render_row_variant(before, row_span, row_raw, &targets[cursor..end], context)?;
        if value != row_raw {
            variants.push(RowVariant { offset, value });
        }
        cursor = end;
    }
    if variants.is_empty() {
        return copy_fragment(row_raw, "ODS metadata unchanged repeated row");
    }

    let mut output_len = 0usize;
    let mut previous = 0usize;
    for variant in &variants {
        render_step(context, 1)?;
        let gap = variant
            .offset
            .checked_sub(previous)
            .ok_or_else(|| invalid("ODS metadata repeated-row target order is invalid"))?;
        if gap > 0 {
            output_len = output_len
                .checked_add(replace_repeat_attr_len(
                    row_raw,
                    row_span,
                    "number-rows-repeated",
                    gap,
                )?)
                .ok_or_else(|| invalid("ODS repeated row split size overflows"))?;
        }
        output_len = output_len
            .checked_add(replace_repeat_attr_len(
                &variant.value,
                row_span,
                "number-rows-repeated",
                1,
            )?)
            .ok_or_else(|| invalid("ODS repeated row split size overflows"))?;
        previous = variant.offset + 1;
    }
    if previous < row.repeat {
        output_len = output_len
            .checked_add(replace_repeat_attr_len(
                row_raw,
                row_span,
                "number-rows-repeated",
                row.repeat - previous,
            )?)
            .ok_or_else(|| invalid("ODS repeated row split size overflows"))?;
    }
    let mut output = String::new();
    render_step(context, 1)?;
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata repeated-row split",
            source,
        })?;
    previous = 0;
    for variant in variants {
        render_step(context, 1)?;
        let gap = variant
            .offset
            .checked_sub(previous)
            .ok_or_else(|| invalid("ODS metadata repeated-row target order is invalid"))?;
        if gap > 0 {
            output.push_str(&clone_repeat(
                row_raw,
                row_span,
                "number-rows-repeated",
                gap,
            )?);
        }
        output.push_str(&replace_repeat_attr(
            &variant.value,
            row_span,
            "number-rows-repeated",
            1,
        )?);
        previous = variant.offset + 1;
    }
    if previous < row.repeat {
        render_step(context, 1)?;
        output.push_str(&clone_repeat(
            row_raw,
            row_span,
            "number-rows-repeated",
            row.repeat - previous,
        )?);
    }
    if output.len() != output_len {
        return Err(invalid(
            "ODS repeated-row output disagrees with its preflight",
        ));
    }
    Ok(output)
}

fn render_row_variant(
    before: &Snapshot,
    row_span: &index::Span,
    row_raw: &str,
    targets: &[CellTarget<'_>],
    context: &ExecutionContext,
) -> Result<String> {
    let mut replacements = Vec::<CellFragment>::new();
    replacements
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata repeated-row cell fragments",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < targets.len() {
        render_step(context, 1)?;
        let cell_index = targets[cursor].cell_index;
        let mut end = cursor + 1;
        while end < targets.len() && targets[end].cell_index == cell_index {
            end += 1;
        }
        let cell = before
            .catalog
            .cells
            .get(cell_index)
            .ok_or_else(|| invalid("ODS metadata cell index disappeared"))?;
        let span = before
            .catalog
            .spans
            .get(cell.span)
            .ok_or_else(|| invalid("ODS metadata cell span disappeared"))?;
        let start = span
            .range
            .start
            .checked_sub(row_span.range.start)
            .ok_or_else(|| invalid("ODS metadata cell span precedes row span"))?;
        let end_offset = span
            .range
            .end
            .checked_sub(row_span.range.start)
            .ok_or_else(|| invalid("ODS metadata cell span exceeds row span"))?;
        row_raw
            .get(start..end_offset)
            .ok_or_else(|| invalid("ODS metadata cell source range is invalid"))?;
        let value = render_repeated_cell_change(before, cell, &targets[cursor..end], context)?;
        replacements.push(CellFragment {
            range: start..end_offset,
            value,
        });
        cursor = end;
    }
    if replacements.is_empty() {
        return copy_fragment(row_raw, "ODS metadata unchanged row variant");
    }
    render_step(context, 1)?;
    let mut output = copy_fragment(row_raw, "ODS metadata row variant")?;
    for replacement in replacements.into_iter().rev() {
        render_step(context, 1)?;
        output = replace_subrange(&output, replacement.range, &replacement.value)?;
    }
    Ok(output)
}

fn render_cell_direct(
    before: &Snapshot,
    cell: &index::CellRecord,
    op: &CellOp,
    span: &index::Span,
    raw: &str,
    context: &ExecutionContext,
) -> Result<String> {
    if op.source.is_none() && op.detective.is_none() {
        return copy_fragment(raw, "ODS metadata unchanged cell");
    }
    if op.source.is_some()
        && cell
            .source_owner
            .as_ref()
            .is_some_and(|owner| owner.opaque || !owner.canonical)
    {
        return Err(Error::Unsupported(
            "ODS cell-range-source owner requires lexical preservation".to_string(),
        ));
    }
    if op.detective.is_some()
        && cell
            .detective_owner
            .as_ref()
            .is_some_and(|owner| owner.opaque || !owner.canonical)
    {
        return Err(Error::Unsupported(
            "ODS detective owner requires lexical preservation".to_string(),
        ));
    }
    render_step(context, 1)?;
    let mut direct = if span.empty {
        expand_empty_cell(before, cell, op, span, raw, context)?
    } else {
        copy_fragment(raw, "ODS metadata cell")?
    };
    if !span.empty {
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(2)
            .map_err(|source| Error::Allocation {
                resource: "ODS metadata cell replacements",
                source,
            })?;
        if let Some(value) = &op.source {
            render_step(context, 1)?;
            replacements.push(direct_owner_replacement(
                before,
                cell,
                "cell-range-source",
                value
                    .as_ref()
                    .map(|source| render_cell_range(source, qname_prefix_for_cell(span, "table"))),
                0,
            )?);
        }
        if let Some(value) = &op.detective {
            render_step(context, 1)?;
            replacements.push(direct_owner_replacement(
                before,
                cell,
                "detective",
                value
                    .as_ref()
                    .map(|value| render_detective(value, qname_prefix_for_cell(span, "table"))),
                1,
            )?);
        }
        replacements
            .sort_by(|left, right| right.0.start.cmp(&left.0.start).then(right.2.cmp(&left.2)));
        for (range, value, _) in replacements {
            render_step(context, 1)?;
            direct = replace_subrange(&direct, range, value.as_deref().unwrap_or(""))?;
        }
    }
    if op.source.as_ref().is_some_and(Option::is_some) && !before.source.contains("xmlns:xlink=") {
        direct = ensure_xlink_namespace(&direct)?;
    }
    Ok(direct)
}

fn expand_empty_cell(
    before: &Snapshot,
    cell: &index::CellRecord,
    op: &CellOp,
    span: &index::Span,
    raw: &str,
    context: &ExecutionContext,
) -> Result<String> {
    if cell.source_owner.is_some() || cell.detective_owner.is_some() {
        return Err(Error::Unsupported(
            "ODS empty metadata cell has an unexpected direct owner".to_string(),
        ));
    }
    let opening_len = opening_end(raw)?;
    let opening = raw
        .get(..opening_len)
        .ok_or_else(|| invalid("ODS empty metadata cell opening is invalid"))?;
    let opening = opening
        .strip_suffix("/>")
        .ok_or_else(|| invalid("ODS empty metadata cell is not self-closing"))?;
    let prefix = qname_prefix_for_cell(span, "table");
    let mut output = String::new();
    output.push_str(opening);
    output.push('>');
    if let Some(Some(value)) = &op.source {
        render_step(context, 1)?;
        output.push_str(&render_cell_range(value, prefix));
    }
    if let Some(Some(value)) = &op.detective {
        render_step(context, 1)?;
        output.push_str(&render_detective(value, prefix));
    }
    output.push_str("</");
    output.push_str(&span.qname);
    output.push('>');
    let _ = before;
    Ok(output)
}

fn direct_owner_replacement(
    before: &Snapshot,
    cell: &index::CellRecord,
    local: &str,
    replacement: Option<String>,
    priority: u8,
) -> Result<(ByteRange<usize>, Option<String>, u8)> {
    let span = &before.catalog.spans[cell.span];
    let owner = if local == "cell-range-source" {
        cell.source_owner.as_ref()
    } else {
        cell.detective_owner.as_ref()
    };
    match (owner, replacement) {
        (Some(owner), Some(value)) => {
            if owner.opaque || !owner.canonical {
                return Err(Error::Unsupported(format!(
                    "ODS {local} owner requires lexical preservation"
                )));
            }
            let local_start = owner
                .raw
                .start
                .checked_sub(span.range.start)
                .ok_or_else(|| invalid("ODS metadata owner span underflows"))?;
            let local_end = owner
                .raw
                .end
                .checked_sub(span.range.start)
                .ok_or_else(|| invalid("ODS metadata owner span overflows"))?;
            Ok((local_start..local_end, Some(value), priority))
        },
        (Some(owner), None) => {
            if owner.opaque || !owner.canonical {
                return Err(Error::Unsupported(format!(
                    "ODS {local} owner requires lexical preservation"
                )));
            }
            let local_start = owner
                .raw
                .start
                .checked_sub(span.range.start)
                .ok_or_else(|| invalid("ODS metadata owner span underflows"))?;
            let local_end = owner
                .raw
                .end
                .checked_sub(span.range.start)
                .ok_or_else(|| invalid("ODS metadata owner span overflows"))?;
            Ok((local_start..local_end, None, priority))
        },
        (None, Some(value)) => {
            if cell.unsupported_child || !cell.sequence_valid {
                return Err(Error::Unsupported(
                    "ODS metadata cell insertion anchor is ambiguous".to_string(),
                ));
            }
            let anchor = cell_insertion_anchor(before, cell, local)?;
            Ok((anchor..anchor, Some(value), priority))
        },
        (None, None) => Ok((0..0, None, priority)),
    }
}

fn cell_insertion_anchor(
    before: &Snapshot,
    cell: &index::CellRecord,
    local: &str,
) -> Result<usize> {
    let span = &before.catalog.spans[cell.span];
    for child in &span.children {
        let child_span = &before.catalog.spans[*child];
        let is_text = child_span.namespace.as_deref() == Some(index::TEXT_NS);
        let is_annotation = child_span.namespace.as_deref() == Some(index::OFFICE_NS)
            && child_span.local == "annotation";
        let is_detective = child_span.namespace.as_deref() == Some(index::TABLE_NS)
            && child_span.local == "detective";
        if local == "cell-range-source" {
            if is_text || is_annotation || is_detective {
                return Ok(child_span.range.start - span.range.start);
            }
        } else if is_text {
            return Ok(child_span.range.start - span.range.start);
        }
    }
    Ok(span.close_start.saturating_sub(span.range.start))
}

fn clone_repeat(raw: &str, span: &index::Span, local: &str, repeat: usize) -> Result<String> {
    replace_repeat_attr(raw, span, local, repeat)
}

fn decimal_len(value: usize) -> usize {
    value.to_string().len()
}

fn replace_repeat_attr(
    raw: &str,
    span: &index::Span,
    local: &str,
    repeat: usize,
) -> Result<String> {
    let Some((global_start, global_end)) = repeat_attribute_value_range(raw, span, local)? else {
        let mut result = String::new();
        result
            .try_reserve_exact(raw.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS repeated owner",
                source,
            })?;
        result.push_str(raw);
        return Ok(result);
    };
    let output_len = raw
        .len()
        .checked_sub(global_end - global_start)
        .and_then(|value| value.checked_add(decimal_len(repeat)))
        .ok_or_else(|| invalid("ODS repeated owner size overflows"))?;
    let mut result = String::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODS repeated owner",
            source,
        })?;
    result.push_str(
        raw.get(..global_start)
            .ok_or_else(|| invalid("ODS repeated owner prefix is invalid"))?,
    );
    fmt::Write::write_fmt(&mut result, format_args!("{repeat}"))
        .map_err(|_| invalid("ODS repeated owner number formatting failed"))?;
    result.push_str(
        raw.get(global_end..)
            .ok_or_else(|| invalid("ODS repeated owner suffix is invalid"))?,
    );
    if result.len() != output_len {
        return Err(invalid(
            "ODS repeated owner replacement disagrees with its preflight",
        ));
    }
    Ok(result)
}

fn replace_repeat_attr_len(
    raw: &str,
    span: &index::Span,
    local: &str,
    repeat: usize,
) -> Result<usize> {
    let Some((start, end)) = repeat_attribute_value_range(raw, span, local)? else {
        return Ok(raw.len());
    };
    raw.len()
        .checked_sub(end - start)
        .and_then(|value| value.checked_add(decimal_len(repeat)))
        .ok_or_else(|| invalid("ODS repeated owner size overflows"))
}

fn repeat_attribute_value_range(
    raw: &str,
    span: &index::Span,
    local: &str,
) -> Result<Option<(usize, usize)>> {
    for attr in &span.attrs {
        if attr.namespace.as_deref() != Some(index::TABLE_NS) || attr.local != local {
            continue;
        }
        let opening_len = opening_end(raw)?;
        let opening = raw
            .get(..opening_len)
            .ok_or_else(|| invalid("ODS repeated owner opening span is invalid"))?;
        let position = find_attribute_name(opening, &attr.qname)
            .ok_or_else(|| invalid("ODS repeated owner attribute span is missing"))?;
        let after_name = position + attr.qname.len();
        let equals = opening[after_name..]
            .find('=')
            .map(|value| after_name + value)
            .ok_or_else(|| invalid("ODS repeated owner attribute lacks '='"))?;
        let quote_start = opening[equals + 1..]
            .find(['\'', '\"'])
            .map(|value| equals + 1 + value)
            .ok_or_else(|| invalid("ODS repeated owner attribute lacks quote"))?;
        let quote = opening.as_bytes()[quote_start] as char;
        let quote_end = opening[quote_start + 1..]
            .find(quote)
            .map(|value| quote_start + 1 + value)
            .ok_or_else(|| invalid("ODS repeated owner attribute lacks closing quote"))?;
        return Ok(Some((quote_start + 1, quote_end)));
    }
    Ok(None)
}

fn opening_end(raw: &str) -> Result<usize> {
    let bytes = raw.as_bytes();
    let mut quote = None;
    for (index, byte) in bytes.iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => return Ok(index + 1),
            _ => {},
        }
    }
    Err(invalid("ODS metadata owner opening tag is unterminated"))
}

fn find_attribute_name(opening: &str, name: &str) -> Option<usize> {
    let bytes = opening.as_bytes();
    let name_bytes = name.as_bytes();
    let mut quote = None;
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if let Some(value) = quote {
            if byte == value {
                quote = None;
            }
            index += 1;
            continue;
        }
        match byte {
            b'\'' | b'"' => quote = Some(byte),
            _ if bytes[index..].starts_with(name_bytes)
                && (index == 0
                    || bytes[index - 1].is_ascii_whitespace()
                    || bytes[index - 1] == b'<')
                && bytes
                    .get(index + name_bytes.len())
                    .is_some_and(|value| value.is_ascii_whitespace() || *value == b'=') =>
            {
                return Some(index);
            },
            _ => {},
        }
        index += 1;
    }
    None
}

fn ensure_xlink_namespace(raw: &str) -> Result<String> {
    if raw.contains("xmlns:xlink=") {
        return Ok(raw.to_string());
    }
    let end = opening_end(raw)?;
    let opening = raw
        .get(..end)
        .ok_or_else(|| invalid("ODS metadata owner opening span is invalid"))?;
    let close = opening
        .strip_suffix('>')
        .ok_or_else(|| invalid("ODS metadata owner opening tag is invalid"))?;
    let mut output = String::new();
    output.push_str(close);
    output.push_str(" xmlns:xlink=\"http://www.w3.org/1999/xlink\">");
    output.push_str(
        raw.get(end..)
            .ok_or_else(|| invalid("ODS metadata owner suffix is invalid"))?,
    );
    Ok(output)
}

fn replace_subrange(source: &str, range: ByteRange<usize>, replacement: &str) -> Result<String> {
    if range.start > range.end || range.end > source.len() {
        return Err(invalid("ODS metadata splice range is invalid"));
    }
    let mut output = String::new();
    let length = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|value| value.checked_add(replacement.len()))
        .ok_or_else(|| invalid("ODS metadata splice length overflows"))?;
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "ODS metadata splice",
            source,
        })?;
    output.push_str(&source[..range.start]);
    output.push_str(replacement);
    output.push_str(&source[range.end..]);
    Ok(output)
}

fn render_consolidation(value: &Options, prefix: &str) -> Result<String> {
    let mut output = String::new();
    crate::model::consolidation::write_consolidation(&mut output, Some(value))?;
    Ok(rename_table_qnames(&output, prefix))
}
fn render_labels(values: &[LabelRange], prefix: &str) -> Result<String> {
    let mut output = String::new();
    if values.is_empty() {
        output.push('<');
        output.push_str(prefix);
        output.push_str(":label-ranges/>");
        return Ok(output);
    }
    crate::model::label_range::write(&mut output, values)?;
    Ok(rename_table_qnames(&output, prefix))
}
fn render_cell_range(value: &CellRange, prefix: &str) -> String {
    let mut output = String::new();
    crate::model::source::write_cell_range_source(&mut output, value);
    rename_table_qnames(&output, prefix)
}
fn render_detective(value: &Detective, prefix: &str) -> String {
    let mut output = String::new();
    crate::model::detective::write_detective(&mut output, value);
    rename_table_qnames(&output, prefix)
}
fn qname_prefix<'a>(span: &'a index::Span, local: &str) -> &'a str {
    span.qname
        .strip_suffix(&format!(":{local}"))
        .unwrap_or("table")
}
fn qname_prefix_for_cell<'a>(span: &'a index::Span, fallback: &'a str) -> &'a str {
    span.qname
        .split_once(':')
        .map(|(prefix, _)| prefix)
        .unwrap_or(fallback)
}

fn table_prefix(before: &Snapshot) -> &str {
    before
        .catalog
        .spans
        .iter()
        .find(|span| span.namespace.as_deref() == Some(index::TABLE_NS))
        .and_then(|span| span.qname.split_once(':').map(|(prefix, _)| prefix))
        .unwrap_or("table")
}

fn rename_table_qnames(source: &str, prefix: &str) -> String {
    if prefix == "table" {
        return source.to_string();
    }
    let mut output = String::with_capacity(source.len());
    let mut in_tag = false;
    let mut quote = None;
    let mut cursor = 0;
    while cursor < source.len() {
        let rest = &source[cursor..];
        let Some(character) = rest.chars().next() else {
            break;
        };
        if in_tag {
            if let Some(value) = quote {
                if character == value {
                    quote = None;
                }
                output.push(character);
                cursor += character.len_utf8();
                continue;
            }
            match character {
                '\'' | '"' => {
                    quote = Some(character);
                    output.push(character);
                    cursor += character.len_utf8();
                },
                '>' => {
                    in_tag = false;
                    output.push('>');
                    cursor += character.len_utf8();
                },
                _ if rest.starts_with("table:") => {
                    output.push_str(prefix);
                    output.push(':');
                    cursor += "table:".len();
                },
                _ => {
                    output.push(character);
                    cursor += character.len_utf8();
                },
            }
        } else if character == '<' {
            in_tag = true;
            output.push('<');
            cursor += character.len_utf8();
        } else {
            output.push(character);
            cursor += character.len_utf8();
        }
    }
    output
}

fn render_top_level_insertion(before: &Snapshot, local: &str, rendered: &str) -> Result<String> {
    let _ = (before, local);
    Ok(rendered.to_string())
}
fn top_level_anchor(before: &Snapshot, local: &str) -> Result<ByteRange<usize>> {
    let spreadsheet = before
        .catalog
        .spans
        .get(before.catalog.spreadsheet)
        .ok_or_else(|| invalid("ODS spreadsheet span disappeared"))?;
    let mut first_table = None;
    let mut first_dde = None;
    let mut last_epilogue = None;
    let mut saw_table = false;
    for child in &spreadsheet.children {
        let span = &before.catalog.spans[*child];
        if span.namespace.as_deref() == Some(index::TABLE_NS) {
            if span.local == "table" {
                first_table.get_or_insert(span.range.start);
                saw_table = true;
            } else if span.local == "dde-links" || span.local == "dde-link" {
                first_dde.get_or_insert(span.range.start);
            } else if matches!(
                span.local.as_str(),
                "named-expressions" | "database-ranges" | "data-pilot-tables"
            ) {
                last_epilogue = Some(span.range.end);
            } else if matches!(
                span.local.as_str(),
                "calculation-settings" | "content-validations" | "label-ranges"
            ) && !saw_table
            {
                // Recognized table-decls members remain in the prelude.
            } else if local == "label-ranges" && saw_table {
                break;
            } else {
                return Err(Error::Unsupported(
                    "ODS metadata insertion anchor is ambiguous around an unsupported spreadsheet child".to_string(),
                ));
            }
        } else if local == "label-ranges" && saw_table {
            break;
        } else {
            return Err(Error::Unsupported(
                "ODS metadata insertion anchor is ambiguous around an unsupported spreadsheet child".to_string(),
            ));
        }
    }
    let start = match local {
        "label-ranges" => first_table.unwrap_or(spreadsheet.close_start),
        "consolidation" => first_dde
            .or(last_epilogue)
            .unwrap_or(spreadsheet.close_start),
        _ => spreadsheet.close_start,
    };
    Ok(start..start)
}
fn copy_source(source: &str, max: usize) -> Result<String> {
    if source.len() > max {
        return Err(limit("output bytes", source.len(), max));
    }
    Ok(source.to_string())
}

fn copy_fragment(source: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|error| Error::Allocation {
            resource,
            source: error,
        })?;
    output.push_str(source);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::Unsupported(format!(
        "ODS metadata {resource} limit exceeded: {actual} > {maximum}"
    ))
}

fn validate_text_limit(value: &str, limits: Limits, name: &'static str) -> Result<()> {
    if value.len() > limits.max_text_bytes() {
        return Err(limit(name, value.len(), limits.max_text_bytes()));
    }
    Ok(())
}

fn validate_options_limits(value: &Options, limits: Limits) -> Result<()> {
    validate_text_limit(&value.function, limits, "table:function")?;
    validate_text_limit(
        &value.target_cell_address,
        limits,
        "table:target-cell-address",
    )?;
    for address in &value.source_cell_range_addresses {
        validate_text_limit(address, limits, "table:source-cell-range-addresses")?;
    }
    Ok(())
}

fn validate_label_limits(value: &LabelRange, limits: Limits) -> Result<()> {
    validate_text_limit(
        &value.label_cell_range_address,
        limits,
        "table:label-cell-range-address",
    )?;
    validate_text_limit(
        &value.data_cell_range_address,
        limits,
        "table:data-cell-range-address",
    )?;
    Ok(())
}

fn validate_cell_range_limits(value: &CellRange, limits: Limits) -> Result<()> {
    validate_text_limit(value.name(), limits, "table:name")?;
    validate_text_limit(value.href(), limits, "xlink:href")?;
    if let Some(value) = value.filter_name() {
        validate_text_limit(value, limits, "table:filter-name")?;
    }
    if let Some(value) = value.filter_options() {
        validate_text_limit(value, limits, "table:filter-options")?;
    }
    if let Some(value) = value.refresh_delay() {
        validate_text_limit(value, limits, "table:refresh-delay")?;
    }
    Ok(())
}

fn validate_detective_limits(value: &Detective, limits: Limits) -> Result<()> {
    for range in value.highlighted_ranges() {
        if let Some(address) = range.cell_range_address() {
            validate_text_limit(address, limits, "table:cell-range-address")?;
        }
    }
    Ok(())
}

fn map_execution(error: litchi_core::ExecutionError) -> Error {
    match error {
        litchi_core::ExecutionError::ResourceLimit(value) => Error::ResourceLimit(value),
        litchi_core::ExecutionError::Cancelled => {
            Error::Unsupported("ODS metadata operation cancelled".to_string())
        },
        other => Error::Unsupported(format!(
            "ODS metadata execution policy rejected operation: {other}"
        )),
    }
}

/// Compare execution-budget lineage without relying on mutable usage or
/// configured limits.  A zero-byte reservation still carries the exact
/// budget-node chain, and `Reservation::try_merge` compares that chain by
/// identity.  Both reservations also check their cancellation tokens before
/// the comparison, so a source context cancelled after parsing cannot be
/// bypassed by supplying an unrelated context with matching numeric limits.
fn same_budget_lineage(left: &ExecutionContext, right: &ExecutionContext) -> Result<bool> {
    let mut left_reservation = left.reserve(Resource::Memory, 0).map_err(map_execution)?;
    let right_reservation = right.reserve(Resource::Memory, 0).map_err(map_execution)?;
    Ok(left_reservation.try_merge(right_reservation).is_ok())
}

fn render_step(context: &ExecutionContext, units: u64) -> Result<()> {
    context.check().map_err(map_execution)?;
    context
        .consume(Resource::Work, units)
        .map_err(map_execution)
}

pub(crate) fn default_context() -> ExecutionContext {
    let budget = Budget::root(
        "ods-sheet-metadata",
        litchi_core::Limits::for_profile(Profile::TrustedBatch),
    );
    let (_source, token) = CancellationSource::pair();
    ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(
            std::num::NonZeroUsize::new(1).expect("one worker"),
            std::num::NonZeroUsize::new(1).expect("one task"),
            std::num::NonZeroU64::new(1024 * 1024).expect("one MiB"),
            0,
        )
        .expect("finite execution profile"),
    )
}

impl crate::facade::SourceBackedSpreadsheet {
    /// Capture a source-backed sheet metadata snapshot under finite defaults.
    pub fn sheet_metadata(&self) -> Result<SourceMetadataSnapshot<'_>> {
        SourceMetadataSnapshot::from_owner(self, Limits::default(), &default_context(), false)
    }
    /// Capture a source-backed sheet metadata snapshot under explicit limits and context.
    pub fn sheet_metadata_with(
        &self,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<SourceMetadataSnapshot<'_>> {
        SourceMetadataSnapshot::from_owner(self, limits, context, true)
    }
    /// Begin a source-backed metadata edit.
    pub fn edit_sheet_metadata(&self) -> Result<SourceMetadataEdit<'_>> {
        self.sheet_metadata()?.edit()
    }
    /// Apply an exact source-backed metadata patch and return its accepted commit.
    pub fn apply_sheet_metadata_patch<'a>(
        &'a self,
        patch: &SourceMetadataPatch<'a>,
    ) -> Result<SourceMetadataCommit<'a>> {
        patch.apply(&self.sheet_metadata()?)
    }
}
