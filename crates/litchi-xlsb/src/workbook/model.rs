//! Typed XLSB workbook state and its public model accessors.

use super::cache::{Stats, WorksheetCache};
use super::catalog::TabMetadata;
use crate::calc::Props;
use crate::external_link::{ExternalLinkLimits, Link};
use crate::package::Cell;
use crate::package::formula::{
    Context, ExternalBook, View, excel_name_eq, table::Definition as TableDefinition,
};
use crate::package::shared_strings::SharedString;
use crate::package::styles_table::{CellFormat, StylesTable};
use crate::sheet::{Grid, Selector, Tab, Worksheet};
use litchi_core::Position;
use litchi_opc::OpcPackage;
use std::sync::Arc;

/// Typed drawing projection selected when the workbook was opened.
///
/// The state is intentionally kept separately from `sheet_drawings`: an
/// eager workbook can legitimately have no drawing relationships yet, and a
/// later structural edit may add one.  Reparse operations must preserve the
/// caller's original projection in that case.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DrawingLoadPolicy {
    Eager,
    Skipped,
}

/// XLSB workbook implementation
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct Workbook {
    pub(crate) package: OpcPackage,
    /// Owner-scoped bounded retention for parsed worksheet snapshots.
    pub(super) worksheet_cache: WorksheetCache,
    pub(super) worksheets: Vec<Worksheet>,
    /// Worksheet-only names in public worksheet ordinal order.
    pub(super) worksheet_names: Vec<String>,
    /// Worksheet ordinal to workbook sheet-catalog position.
    pub(super) worksheet_positions: Vec<usize>,
    /// Relationship identifiers in full workbook sheet-catalog order.
    pub(super) worksheet_rel_ids: Vec<Option<String>>,
    /// The validated semantic metadata of every tab, in full workbook
    /// sheet-catalog order. The formula context retains its own full-name
    /// projection for formula resolution; native IDs and part names remain
    /// in `worksheet_rel_ids` and never enter this semantic value.
    pub(super) sheet_tabs: Vec<TabMetadata>,
    /// Primary `BrtBookView.itabCur` position in full workbook sheet order.
    pub(super) active_catalog_position: Option<usize>,
    pub(crate) formula_context: Context,
    /// Caller-selected raw BIFF12 policy retained for lazy workbook reads.
    pub(super) raw_limits: crate::raw::Limits,
    pub(super) external_link_limits: ExternalLinkLimits,
    pub(super) shared_strings: Vec<SharedString>,
    pub(super) styles: StylesTable,
    pub(super) calc: Props,
    pub(super) is_1904: bool,
    pub(super) pivot_cache_definitions: Vec<(u32, crate::package::pivot::PivotCacheDefinition)>,
    pub(super) structured_tables: Vec<(usize, crate::package::table::Table)>,
    pub(super) chart_sheets: Vec<(usize, crate::package::chartsheet::ChartSheet)>,
    pub(super) sheet_drawings: Vec<crate::package::drawing::SheetDrawing>,
    pub(super) drawing_load_policy: DrawingLoadPolicy,
    pub(super) connections: Option<crate::package::connections::Connections>,
    /// Repairs applied while the workbook-scope parts were parsed at open.
    pub(super) tolerance: crate::leniency::ToleranceReport,
    /// Opt-in repair policy applied to lazily parsed worksheet content.
    pub(super) leniency: crate::leniency::Leniency,
    /// The semantic budgets of the family reads of this projection.
    pub(super) semantic_limits: crate::SemanticLimits,
}

impl std::fmt::Debug for Workbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Workbook")
            .field("sheet_tab_count", &self.sheet_tabs.len())
            .field("worksheet_count", &self.worksheet_names.len())
            .field("active_catalog_position", &self.active_catalog_position)
            .field("shared_strings_count", &self.shared_strings.len())
            .field("cell_xfs_count", &self.styles.cell_xfs.len())
            .field("calc", &self.calc)
            .field("is_1904", &self.is_1904)
            .field("drawings_parsed", &self.drawings_parsed())
            .finish()
    }
}

impl Workbook {
    pub(crate) fn cached_worksheet(
        &self,
        index: usize,
    ) -> crate::package::error::Result<Arc<Worksheet>> {
        let catalog_position = self.catalog_position_for_worksheet(index)?;
        if let Some(worksheet) = self.worksheet_cache.get(catalog_position, true) {
            return Ok(worksheet);
        }
        let worksheet = Arc::new(self.parse_worksheet(index, catalog_position)?);
        let weight = worksheet.retained_bytes();
        Ok(self
            .worksheet_cache
            .publish(catalog_position, worksheet, weight))
    }

    /// Content-free diagnostics for the owner-scoped worksheet cache.
    #[must_use]
    pub fn semantic_cache_diagnostics(&self) -> Stats {
        self.worksheet_cache.stats()
    }

    /// Release parsed worksheet values that no grid or legacy handle pins.
    pub fn trim_semantic_cache(&self) -> usize {
        self.worksheet_cache.trim_clean()
    }

    /// Every repair applied while this workbook was opened.
    ///
    /// Covers the OPC package read that admitted the workbook (its
    /// packaging repairs come first, as [`crate::Defect::Package`] entries)
    /// and the parts parsed at open: the workbook stream (`workbook.bin`)
    /// and the styles part (`styles.bin`). Worksheets are parsed lazily, so
    /// their repairs are reported by [`Worksheet::tolerance_report`] on each
    /// materialized worksheet. An empty report means the parsed parts
    /// conformed; see [`crate::leniency`] for the repair contract.
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.tolerance
    }

    /// The [`Leniency`](crate::Leniency) this workbook was opened with.
    #[must_use]
    pub const fn leniency(&self) -> crate::leniency::Leniency {
        self.leniency
    }

    /// Clone the physical package owner while retaining this workbook's policies.
    ///
    /// Payloads retain their existing shared ownership; the workbook projection
    /// and its worksheet caches are not copied.
    #[must_use]
    pub fn to_package(&self) -> crate::Package {
        crate::Package::from_workbook(self)
    }

    /// Consume this workbook into its physical package owner without reparsing.
    ///
    /// The OPC graph, external-link limits, drawing projection, worksheet
    /// repair policy and open-time tolerance report are retained.
    ///
    /// ```
    /// # fn main() -> litchi_xlsb::package::PackageResult<()> {
    /// let bytes = litchi_xlsb::Package::new()?.to_bytes()?;
    /// let workbook = litchi_xlsb::Workbook::new(std::io::Cursor::new(bytes))?;
    /// let package = workbook.into_package();
    /// assert_eq!(package.leniency(), litchi_xlsb::Leniency::Strict);
    /// assert!(package.drawings_parsed());
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub fn into_package(self) -> crate::Package {
        crate::Package::from_validated_workbook(self)
    }

    pub(crate) fn set_projection_leniency(&mut self, leniency: crate::leniency::Leniency) {
        self.leniency = leniency;
    }

    pub(crate) fn set_semantic_limits(&mut self, limits: crate::SemanticLimits) {
        self.semantic_limits = limits;
    }

    /// The semantic budgets of this projection's family reads.
    #[must_use]
    pub const fn semantic_limits(&self) -> crate::SemanticLimits {
        self.semantic_limits
    }

    /// The BIFF12 raw policy retained by this workbook projection.
    #[must_use]
    pub const fn raw_limits_value(&self) -> crate::raw::Limits {
        self.raw_limits
    }

    /// Return the drawing projection boundary selected when this workbook was
    /// opened.  Internal publication helpers must carry this state through
    /// every reconstruction; a skipped projection is an intentional opaque
    /// boundary, not an invitation to parse drawings later.
    #[must_use]
    pub(crate) const fn drawing_load_policy(&self) -> DrawingLoadPolicy {
        self.drawing_load_policy
    }
    /// Translate a public worksheet ordinal to the full workbook sheet-catalog
    /// position used by formula and package metadata.
    /// The workbook's defined names in record order, with their scope and
    /// formula text.
    ///
    /// # Errors
    ///
    /// Returns the record and defined-name decoding errors of the workbook
    /// stream.
    pub fn defined_name_entries(
        &self,
    ) -> crate::package::error::Result<Vec<crate::named_ranges::Entry>> {
        let part = self.package.main_document_part()?;
        crate::named_ranges::read_entries(part.blob(), &self.formula_context, self.raw_limits)
    }

    /// The worksheet ordinal of a complete-catalog position, when that tab is
    /// a worksheet.
    pub(crate) fn worksheet_for_catalog_position(&self, catalog_position: usize) -> Option<usize> {
        self.worksheet_positions
            .iter()
            .position(|position| *position == catalog_position)
    }

    pub(crate) fn catalog_position_for_worksheet(
        &self,
        index: usize,
    ) -> crate::package::error::Result<usize> {
        self.worksheet_positions.get(index).copied().ok_or(
            crate::package::error::Error::InvalidFormat("Worksheet index out of bounds"),
        )
    }

    /// Every workbook tab in workbook order, including chart, dialog, and
    /// macro tabs. The worksheet-only APIs continue to expose only entries
    /// whose [`Tab::worksheet_position`] is `Some`.
    pub fn sheets(&self) -> impl ExactSizeIterator<Item = Tab<'_>> + DoubleEndedIterator + '_ {
        self.sheet_tabs
            .iter()
            .enumerate()
            .map(|(position, tab)| self.tab_at(position, tab))
    }

    /// Select any workbook tab by its semantic name or checked zero-based
    /// position in complete workbook order.
    pub fn sheet<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> crate::package::error::Result<Option<Tab<'_>>> {
        let position = match selector.into() {
            Selector::Position(position) => Some(position.get()),
            Selector::Name(name) => self
                .sheet_tabs
                .iter()
                .position(|tab| excel_name_eq(tab.name(), &name)),
        };
        Ok(position.and_then(|position| {
            self.sheet_tabs
                .get(position)
                .map(|tab| self.tab_at(position, tab))
        }))
    }

    /// Select a semantic workbook tab and return its immutable worksheet grid.
    ///
    /// A missing name or position returns `Ok(None)`. A selected chart,
    /// dialog, macro, or unbound tab returns a typed unsupported-feature
    /// error because those tabs do not own a worksheet grid.
    pub fn grid<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> crate::package::error::Result<Option<Grid>> {
        let Some(tab) = self.sheet(selector)? else {
            return Ok(None);
        };
        let Some(position) = tab.worksheet_position() else {
            return Err(crate::package::error::Error::UnsupportedFeature(
                "selected XLSB tab has no worksheet grid",
            ));
        };
        // `Tab::worksheet_position` is the worksheet-only ordinal. The
        // catalog position was already used to select the tab above, so do
        // not feed this ordinal back into the catalog-position helper when a
        // chart or other non-worksheet tab precedes it.
        self.cached_worksheet(position.get())
            .map(|worksheet| Some(Grid::from_arc(worksheet)))
    }

    fn tab_at<'a>(&'a self, position: usize, tab: &'a TabMetadata) -> Tab<'a> {
        let worksheet_position = self
            .worksheet_positions
            .binary_search(&position)
            .ok()
            .map(Position::new);
        Tab::new(
            tab.name(),
            Position::new(position),
            worksheet_position,
            tab.kind(),
            tab.state(),
        )
    }

    /// Workbook and sheet-scoped defined names in `PtgName` index order.
    pub fn defined_names(&self) -> &[String] {
        &self.formula_context.defined_names
    }

    /// Number of stored external-workbook, DDE, and OLE links.
    pub fn external_link_count(&self) -> usize {
        self.formula_context.external_books.len()
    }

    /// Return the external-link resource policy used to construct this
    /// workbook. Each eager reconstruction starts fresh counters from this
    /// plain policy.
    #[must_use]
    pub const fn external_link_limits(&self) -> ExternalLinkLimits {
        self.external_link_limits
    }

    /// Borrow one stored external link without cloning its cached values.
    pub fn external_link(&self, index: usize) -> Option<&Link> {
        self.formula_context
            .external_books
            .get(index)
            .map(ExternalBook::metadata_ref)
    }

    /// Iterate stored external links without cloning their cached values.
    pub fn external_link_iter(&self) -> impl ExactSizeIterator<Item = &Link> + DoubleEndedIterator {
        self.formula_context
            .external_books
            .iter()
            .map(ExternalBook::metadata_ref)
    }

    /// Return stored external-workbook, DDE, and OLE link metadata.
    ///
    /// The returned values are cloned data-only snapshots in workbook link
    /// order. Use [`Self::external_link_iter`] for zero-copy access to large
    /// cached matrices.
    /// Litchi never follows, opens, contacts, refreshes, evaluates, or
    /// executes any external-link target.
    pub fn external_links(&self) -> Vec<Link> {
        self.formula_context
            .external_books
            .iter()
            .map(ExternalBook::metadata)
            .collect()
    }

    /// Structured-table definitions in workbook table-ID order of discovery.
    pub fn tables(&self) -> &[TableDefinition] {
        &self.formula_context.tables
    }

    /// PivotTable views available as hosts for calculated field/item formulas.
    pub fn pivot_views(&self) -> &[View] {
        &self.formula_context.pivot_views
    }

    /// Typed PivotCache definitions paired with their workbook cache
    /// identifiers, in workbook declaration order (MS-XLSB 2.1.7.38).
    ///
    /// These are inert data snapshots: external connection identifiers,
    /// relationship identifiers, MDX expressions, and formula tokens are
    /// stored verbatim and are never dereferenced, refreshed, or evaluated.
    pub fn pivot_cache_definitions(&self) -> &[(u32, crate::package::pivot::PivotCacheDefinition)] {
        &self.pivot_cache_definitions
    }

    /// Look up a typed PivotCache definition by its workbook cache identifier.
    pub fn pivot_cache_definition(
        &self,
        cache_id: u32,
    ) -> Option<&crate::package::pivot::PivotCacheDefinition> {
        self.pivot_cache_definitions
            .iter()
            .find(|(id, _)| *id == cache_id)
            .map(|(_, definition)| definition)
    }
    /// Typed structured-table (ListObject) definitions paired with their
    /// public worksheet ordinals, in worksheet discovery order (MS-XLSB 2.1.7.51).
    ///
    /// These are inert data snapshots: relationship identifiers, external
    /// connection identifiers, differential-formatting identifiers, and
    /// formula token streams are stored verbatim and are never dereferenced,
    /// contacted, or evaluated. Named `structured_tables` because
    /// [`Workbook::tables`] already exposes the formula-context table
    /// definitions.
    pub fn structured_tables(&self) -> &[(usize, crate::package::table::Table)] {
        &self.structured_tables
    }

    /// Typed structured-table (ListObject) definitions anchored to one
    /// worksheet, selected by zero-based worksheet index.
    pub fn tables_on_sheet(&self, sheet_index: usize) -> Vec<&crate::package::table::Table> {
        self.structured_tables
            .iter()
            .filter(|(index, _)| *index == sheet_index)
            .map(|(_, table)| table)
            .collect()
    }

    /// Typed chart sheet definitions paired with their sheet indexes, in
    /// workbook sheet order (MS-XLSB 2.1.7.7).
    ///
    /// These are inert data snapshots: relationship identifiers, password
    /// verifiers, and hash data are stored verbatim and are never
    /// dereferenced, verified, or executed. The chart hosted by a chart
    /// sheet is surfaced through [`Workbook::sheet_drawing`].
    pub fn chart_sheets(&self) -> &[(usize, crate::package::chartsheet::ChartSheet)] {
        &self.chart_sheets
    }

    /// Look up the typed chart sheet anchored to one sheet, selected by
    /// zero-based sheet index; `None` for worksheets and macro sheets.
    pub fn chart_sheet(
        &self,
        sheet_index: usize,
    ) -> Option<&crate::package::chartsheet::ChartSheet> {
        self.chart_sheets
            .iter()
            .find(|(index, _)| *index == sheet_index)
            .map(|(_, chart_sheet)| chart_sheet)
    }

    /// Drawings part inventories anchored to sheets, in sheet discovery
    /// order (MS-XLSB 2.1.7.23), with referenced images resolved and charts
    /// parsed into the shared typed chart model.
    ///
    /// These are inert data snapshots. Internal image and chart parts are
    /// resolved during package loading; external targets are never fetched.
    /// The slice is empty for a workbook opened through the drawing-skipped
    /// projection, even when the underlying package contains drawing parts;
    /// use [`Self::drawings_parsed`] to distinguish that state from an eager
    /// workbook with no drawing relationships.
    pub fn sheet_drawings(&self) -> &[crate::package::drawing::SheetDrawing] {
        &self.sheet_drawings
    }

    /// Whether standard SpreadsheetDrawing parts were decoded into the typed
    /// drawing inventory.
    ///
    /// A `false` result means the workbook was opened through the explicit
    /// drawing-skipped projection; it does not mean that the OPC package has
    /// no drawing parts. Those parts remain preserved for lossless publication.
    #[must_use]
    pub const fn drawings_parsed(&self) -> bool {
        matches!(self.drawing_load_policy, DrawingLoadPolicy::Eager)
    }

    /// Look up the drawing inventory of one sheet, selected by zero-based
    /// sheet index. Returns `None` when the sheet has no Drawings part or when
    /// the workbook uses the drawing-skipped projection; use
    /// [`Self::drawings_parsed`] to distinguish those cases.
    pub fn sheet_drawing(
        &self,
        sheet_index: usize,
    ) -> Option<&crate::package::drawing::SheetDrawing> {
        self.sheet_drawings
            .iter()
            .find(|drawing| drawing.sheet_index == sheet_index)
    }

    pub fn styles(&self) -> &StylesTable {
        &self.styles
    }

    /// Unique strings loaded from `xl/sharedStrings.bin`, including rich-text
    /// and phonetic metadata when present.
    pub fn shared_strings(&self) -> &[SharedString] {
        &self.shared_strings
    }

    /// Resolve a parsed cell's style reference to its cell XF.
    pub fn style_for_cell(&self, cell: &Cell) -> Option<&CellFormat> {
        self.styles.get_cell_format(cell.style_id() as usize)
    }

    /// Validated workbook formula calculation policy.
    pub fn calc(&self) -> &Props {
        &self.calc
    }

    /// Return the primary workbook view's physical sheet-catalog position.
    #[must_use]
    pub fn active_catalog_position(&self) -> Option<usize> {
        self.active_catalog_position
    }

    /// Return the primary workbook view's public worksheet ordinal, when it
    /// names a worksheet rather than a chart, dialog, or macro sheet.
    #[must_use]
    pub fn active_worksheet_index(&self) -> Option<usize> {
        self.active_catalog_position.and_then(|catalog_position| {
            self.worksheet_positions
                .iter()
                .position(|&position| position == catalog_position)
        })
    }
}

impl litchi_core::sheet::WorkbookTrait for Workbook {
    /// The worksheet snapshot the owner's cache retains, shared rather than
    /// copied.
    type Error = crate::package::error::Error;
    type Worksheet<'a> = Arc<Worksheet>;
    type Worksheets<'a> = Worksheets<'a>;

    fn active_sheet_index(&self) -> usize {
        self.active_worksheet_index().unwrap_or(0)
    }

    fn active_worksheet(&self) -> crate::package::error::Result<Arc<Worksheet>> {
        let index = self.active_worksheet_index().ok_or(
            crate::package::error::Error::UnsupportedFeature(
                "XLSB active sheet is not a worksheet",
            ),
        )?;
        self.worksheet_by_index(index)
    }

    fn worksheet_count(&self) -> usize {
        self.worksheet_names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_index(&self, index: usize) -> crate::package::error::Result<Arc<Worksheet>> {
        self.cached_worksheet(index)
    }

    fn worksheet_by_name(&self, name: &str) -> crate::package::error::Result<Arc<Worksheet>> {
        for (i, ws_name) in self.worksheet_names.iter().enumerate() {
            if ws_name == name {
                return self.worksheet_by_index(i);
            }
        }
        Err(crate::package::error::Error::InvalidFormat("Worksheet not"))
    }

    fn worksheets(&self) -> Worksheets<'_> {
        Worksheets {
            workbook: self,
            index: 0,
        }
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904
    }
}

/// The worksheets of an eager XLSB workbook, in workbook order: the iterator
/// [`WorkbookTrait::worksheets`](litchi_core::sheet::WorkbookTrait::worksheets)
/// returns. A worksheet that fails to parse yields its error, and iteration
/// continues with the next one.
pub struct Worksheets<'a> {
    workbook: &'a Workbook,
    index: usize,
}

impl Iterator for Worksheets<'_> {
    type Item = crate::package::error::Result<Arc<Worksheet>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index < self.workbook.worksheet_names.len() {
            let result = self.workbook.cached_worksheet(self.index);
            // Continue to the next worksheet even on error.
            self.index += 1;
            Some(result)
        } else {
            None
        }
    }
}
