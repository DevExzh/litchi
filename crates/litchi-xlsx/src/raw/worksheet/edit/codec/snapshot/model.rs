//! Captured worksheet layout and lossless XML tag model.

use litchi_sheet::{Cell as Address, Column, Rect};

use crate::raw::worksheet::edit::model::SelectionRange;

#[derive(Debug, Clone, Copy)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct Attribute {
    pub(crate) name: Box<str>,
    pub(crate) value: Box<str>,
}

#[derive(Debug, Clone)]
pub(crate) struct Tag {
    pub(crate) name: Box<str>,
    pub(crate) attributes: Box<[Attribute]>,
}

#[derive(Debug)]
pub(crate) struct CellSlot {
    pub(crate) address: Address,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    /// `None` is the common unprefixed `<c>` form with no attributes or only
    /// one unqualified `r` attribute. The address is parsed separately, so
    /// the writer can regenerate `r` without retaining an owned tag.
    pub(crate) tag: Option<Tag>,
    pub(crate) primary: Box<[Span]>,
    pub(crate) mce_payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct RowSlot {
    pub(crate) number: u32,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) descent_attribute: Option<Box<str>>,
    pub(crate) cells: Box<[CellSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct DefaultsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) descent_attribute: Option<Box<str>>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct RootSlot {
    pub(crate) span: Span,
    pub(crate) tag: Tag,
}

#[derive(Debug)]
pub(crate) struct ColumnSlot {
    pub(crate) first: Column,
    pub(crate) last: Column,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct ColumnsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) columns: Box<[ColumnSlot]>,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct SheetData {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) rows: Box<[RowSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct DimensionTag {
    pub(crate) span: Span,
    pub(crate) tag: Tag,
    pub(crate) empty: bool,
    pub(crate) declared: Rect,
}

#[derive(Debug)]
pub(crate) struct MergeSlot {
    pub(crate) range: Rect,
    pub(crate) span: Span,
}

#[derive(Debug)]
pub(crate) struct MergeCellsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) merges: Box<[MergeSlot]>,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct SharedFormulaGroup {
    pub(crate) index: u32,
    pub(crate) reference: Box<str>,
    pub(crate) origin: Address,
    pub(crate) members: Box<[Address]>,
}

#[derive(Debug)]
pub(crate) struct Layout {
    pub(crate) root: RootSlot,
    pub(crate) defaults: Option<DefaultsSlot>,
    pub(crate) sheet_data: SheetData,
    pub(crate) columns: Option<ColumnsSlot>,
    pub(crate) dimension: Option<DimensionTag>,
    pub(crate) protected: bool,
    pub(crate) merged: Box<[SelectionRange]>,
    pub(crate) validations: Box<[SelectionRange]>,
    pub(crate) extended_validation: bool,
    pub(crate) formula_ranges: Box<[SelectionRange]>,
    pub(crate) shared_formulas: Box<[SharedFormulaGroup]>,
    pub(crate) has_shared_formulas: bool,
    pub(crate) defaults_compatibility: bool,
    pub(crate) merge_cells: Option<MergeCellsSlot>,
    pub(crate) merge_insertion: usize,
    pub(crate) merge_compatibility: bool,
}

#[derive(Debug)]
pub(crate) struct RootEffect {
    pub(crate) removed: Option<Box<str>>,
    pub(crate) appended: Vec<(Box<str>, String)>,
}

/// One source-bound cell record retained from the planning traversal.
///
/// Sixteen bytes: the resolved address and the exact source span of the whole
/// `<c>` element. Unchanged cells are copied from that span; a changed cell's
/// tag and payload spans are materialized from it at commit.
#[derive(Debug, Clone, Copy)]
pub(crate) struct CellFact {
    pub(crate) address: Address,
    pub(crate) start: u32,
    pub(crate) end: u32,
}

/// One source-bound row record retained from the planning traversal.
#[derive(Debug, Clone, Copy)]
pub(crate) struct RowFact {
    pub(crate) number: u32,
    pub(crate) start: u32,
    pub(crate) tag_end: u32,
    pub(crate) close_start: u32,
    pub(crate) end: u32,
    pub(crate) first_cell: u32,
    pub(crate) cell_count: u32,
    pub(crate) empty: bool,
}

/// The declared `<dimension>` and its span, retained so an expanding edit can
/// rewrite it without rebuilding the whole layout.
#[derive(Debug, Clone, Copy)]
pub(crate) struct DimensionFact {
    pub(crate) span: Span,
    pub(crate) empty: bool,
    pub(crate) declared: Rect,
}

/// The compact worksheet facts a value-only commit needs in place of a second
/// complete layout scan.
#[derive(Debug)]
pub(crate) struct SourceFacts {
    /// Address and length of the exact source slice these facts describe.
    ///
    /// Every snapshot that rebinds its worksheet bytes drops its facts, so
    /// this pair is a second, independent guard rather than the only one: a
    /// fact set can only be consumed against the allocation it was built from.
    pub(crate) source_ptr: usize,
    pub(crate) source_len: usize,
    pub(crate) dimension: Option<DimensionFact>,
    pub(crate) sheet_data: Span,
    pub(crate) sheet_data_tag_end: usize,
    pub(crate) sheet_data_close_start: usize,
    pub(crate) sheet_data_empty: bool,
    pub(crate) rows: Box<[RowFact]>,
    pub(crate) cells: Box<[CellFact]>,
}

impl SourceFacts {
    /// Whether these facts were captured from exactly this source.
    pub(crate) fn describes(&self, content: &[u8]) -> bool {
        self.source_ptr == content.as_ptr() as usize
            && self.source_len == content.len()
            && !self.sheet_data_empty
            && self.sheet_data.end <= content.len()
            && self.sheet_data.start <= self.sheet_data_tag_end
            && self.sheet_data_tag_end <= self.sheet_data_close_start
            && self.sheet_data_close_start <= self.sheet_data.end
    }

    /// The retained cells of one row.
    pub(crate) fn row_cells(&self, row: &RowFact) -> Option<&[CellFact]> {
        let start = usize::try_from(row.first_cell).ok()?;
        let end = start.checked_add(usize::try_from(row.cell_count).ok()?)?;
        self.cells.get(start..end)
    }
}

/// Sizes of the retained records, exposed so a test can pin them.
#[cfg(test)]
pub(crate) mod sizes {
    use super::{CellFact, CellSlot, RowFact};

    pub(crate) const CELL_FACT: usize = size_of::<CellFact>();
    pub(crate) const ROW_FACT: usize = size_of::<RowFact>();
    pub(crate) const CELL_SLOT: usize = size_of::<CellSlot>();
}
