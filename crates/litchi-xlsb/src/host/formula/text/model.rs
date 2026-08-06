//! Context and encoding state shared by the formula-text compiler.

use super::super::{Definition, ExternalBook, ExternalSheet, SupportingLink};
use std::cell::RefCell;

/// A defined name visible to the XLSB formula text compiler.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DefinedName {
    pub(crate) name: String,
    pub(crate) sheet_id: Option<u32>,
}

/// Workbook metadata used to compile context-dependent formula operands.
#[derive(Debug, Clone, Copy)]
pub(crate) struct CompilationContext<'a> {
    pub(crate) worksheet_names: &'a [String],
    pub(crate) defined_names: &'a [DefinedName],
    pub(crate) tables: &'a [Definition],
    pub(crate) supporting_links: &'a [SupportingLink],
    pub(crate) external_sheets: &'a [ExternalSheet],
    pub(crate) external_books: &'a [ExternalBook],
    pub(crate) sheet_ranges: &'a RefCell<Vec<(u32, u32)>>,
    pub(crate) current_sheet: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum FormulaEncoding {
    Cell,
    Shared { base_row: u32, base_col: u32 },
}
