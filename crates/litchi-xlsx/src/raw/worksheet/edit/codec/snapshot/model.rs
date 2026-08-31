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
    pub(crate) tag: Tag,
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
