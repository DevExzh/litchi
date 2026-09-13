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

/// Bounded source spans used by the commit-local scalar rewrite. Unlike
/// [`Layout`], this representation retains no owned tags or payload vectors.
#[derive(Debug, Clone, Copy)]
pub(crate) struct CompactSpan {
    pub(crate) start: u32,
    pub(crate) end: u32,
}

impl CompactSpan {
    pub(crate) const fn start(self) -> usize {
        self.start as usize
    }

    pub(crate) const fn end(self) -> usize {
        self.end as usize
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct CompactCellSlot {
    pub(crate) span: CompactSpan,
}

#[derive(Debug, Clone)]
pub(crate) struct CompactRowSlot {
    pub(crate) number: u32,
    pub(crate) span: CompactSpan,
    pub(crate) tag_end: u32,
    pub(crate) close_start: u32,
    pub(crate) cells: Box<[CompactCellSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug, Clone)]
pub(crate) struct CompactSheetData {
    pub(crate) span: CompactSpan,
    pub(crate) tag_end: u32,
    pub(crate) close_start: u32,
    pub(crate) rows: Box<[CompactRowSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct CompactDimensionTag {
    pub(crate) span: CompactSpan,
    pub(crate) empty: bool,
    pub(crate) declared: Rect,
}

/// Ephemeral source/store-bound facts sufficient for an existing-row scalar
/// rewrite. This value is created and dropped inside one commit.
#[derive(Debug, Clone)]
pub(crate) struct CompactLayout {
    pub(crate) source_ptr: usize,
    pub(crate) source_len: usize,
    pub(crate) entries_ptr: usize,
    pub(crate) entries_len: usize,
    pub(crate) cell_count: u32,
    pub(crate) sheet_data: CompactSheetData,
    pub(crate) dimension: Option<CompactDimensionTag>,
}

impl CompactLayout {
    pub(crate) fn matches_source(&self, source: &[u8]) -> bool {
        self.source_ptr == source.as_ptr() as usize && self.source_len == source.len()
    }

    pub(crate) fn matches_entries(&self, entries: &[crate::cell::Stored]) -> bool {
        self.entries_ptr == entries.as_ptr() as usize && self.entries_len == entries.len()
    }
}

#[derive(Debug)]
pub(crate) struct RootEffect {
    pub(crate) removed: Option<Box<str>>,
    pub(crate) appended: Vec<(Box<str>, String)>,
}
