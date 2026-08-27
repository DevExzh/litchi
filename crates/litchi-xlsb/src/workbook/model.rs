//! Typed XLSB workbook state and its public model accessors.

use crate::calc::Props;
use crate::external_link::Link;
use crate::package::Cell;
use crate::package::formula::{Context, ExternalBook, View, table::Definition as TableDefinition};
use crate::package::shared_strings::SharedString;
use crate::package::styles_table::{CellFormat, StylesTable};
use crate::sheet::Worksheet;
use litchi_core::sheet::{
    Result as SheetResult, Worksheet as SheetTrait, WorksheetIterator as SheetIterator,
};
use litchi_opc::OpcPackage;

/// XLSB workbook implementation
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct Workbook {
    pub(crate) package: OpcPackage,
    pub(super) worksheets: Vec<Worksheet>,
    /// Worksheet-only names in public worksheet ordinal order.
    pub(super) worksheet_names: Vec<String>,
    /// Worksheet ordinal to workbook sheet-catalog position.
    pub(super) worksheet_positions: Vec<usize>,
    /// Relationship identifiers in full workbook sheet-catalog order.
    pub(super) worksheet_rel_ids: Vec<Option<String>>,
    /// Primary `BrtBookView.itabCur` position in full workbook sheet order.
    pub(super) active_catalog_position: Option<usize>,
    pub(crate) formula_context: Context,
    pub(super) shared_strings: Vec<SharedString>,
    pub(super) styles: StylesTable,
    pub(super) calc: Props,
    pub(super) is_1904: bool,
    pub(super) pivot_cache_definitions: Vec<(u32, crate::package::pivot::PivotCacheDefinition)>,
    pub(super) structured_tables: Vec<(usize, crate::package::table::Table)>,
    pub(super) chart_sheets: Vec<(usize, crate::package::chartsheet::ChartSheet)>,
    pub(super) sheet_drawings: Vec<crate::package::drawing::SheetDrawing>,
    pub(super) connections: Option<crate::package::connections::Connections>,
}

impl std::fmt::Debug for Workbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Workbook")
            .field("worksheet_names", &self.formula_context.worksheet_names)
            .field("worksheet_rel_ids", &self.worksheet_rel_ids)
            .field("shared_strings_count", &self.shared_strings.len())
            .field("cell_xfs_count", &self.styles.cell_xfs.len())
            .field("calc", &self.calc)
            .field("is_1904", &self.is_1904)
            .finish()
    }
}

impl Workbook {
    /// Translate a public worksheet ordinal to the full workbook sheet-catalog
    /// position used by formula and package metadata.
    pub(crate) fn catalog_position_for_worksheet(
        &self,
        index: usize,
    ) -> crate::package::error::Result<usize> {
        self.worksheet_positions.get(index).copied().ok_or_else(|| {
            crate::package::error::Error::InvalidFormat(format!(
                "Worksheet index {index} out of bounds"
            ))
        })
    }

    /// Workbook and sheet-scoped defined names in `PtgName` index order.
    pub fn defined_names(&self) -> &[String] {
        &self.formula_context.defined_names
    }

    /// Number of stored external-workbook, DDE, and OLE links.
    pub fn external_link_count(&self) -> usize {
        self.formula_context.external_books.len()
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
    pub fn sheet_drawings(&self) -> &[crate::package::drawing::SheetDrawing] {
        &self.sheet_drawings
    }

    /// Look up the drawing inventory of one sheet, selected by zero-based
    /// sheet index; `None` when the sheet has no Drawings part.
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
    fn active_sheet_index(&self) -> usize {
        self.active_worksheet_index().unwrap_or(0)
    }

    fn active_worksheet(&self) -> SheetResult<Box<dyn SheetTrait + '_>> {
        let index = self.active_worksheet_index().ok_or_else(|| {
            Box::new(crate::package::error::Error::UnsupportedFeature(
                "XLSB active sheet is not a worksheet".to_string(),
            ))
        })?;
        self.worksheet_by_index(index)
    }

    fn worksheet_count(&self) -> usize {
        self.worksheet_names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn SheetTrait + '_>> {
        let worksheet = self.worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn SheetTrait + '_>> {
        for (i, ws_name) in self.worksheet_names.iter().enumerate() {
            if ws_name == name {
                return self.worksheet_by_index(i);
            }
        }
        Err(Box::new(crate::package::error::Error::InvalidFormat(
            format!("Worksheet '{}' not found", name),
        )))
    }

    fn worksheets<'a>(&'a self) -> Box<dyn SheetIterator<'a> + 'a> {
        Box::new(WorksheetIterator {
            workbook: self,
            index: 0,
        })
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904
    }
}

pub struct WorksheetIterator<'a> {
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> SheetIterator<'a> for WorksheetIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn SheetTrait + 'a>>> {
        if self.index < self.workbook.worksheet_names.len() {
            match self.workbook.worksheet(self.index) {
                Ok(worksheet) => {
                    self.index += 1;
                    Some(Ok(Box::new(worksheet)))
                },
                Err(e) => {
                    self.index += 1; // Continue to next worksheet even on error
                    Some(Err(Box::new(e)))
                },
            }
        } else {
            None
        }
    }
}
