//! Native table discovery and cell editing for Pages body attachments.

mod comments;
mod conditional_highlight;
mod formula;
mod reader;
mod semantic;
mod sort;
mod storage;
mod topology;
mod validation;

#[cfg(test)]
mod tests;

pub use conditional_highlight::TableCellConditionalHighlightInfo;
pub use formula::{
    TableFormulaAxisReference, TableFormulaBinaryOperator, TableFormulaCachedValue,
    TableFormulaCellReference, TableFormulaExpression,
};
pub use sort::TableSortRowRange;
use std::collections::HashSet;

use super::*;
use litchi_iwa_common::table::appearance::Appearance as TableAppearance;
use litchi_numbers::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;
const INLINE_TABLE_DUPLICATE_OFFSET: f32 = 0.0;

/// Strongly typed cell value shared by Pages and Numbers table storage.
pub type CellValue = litchi_iwa_common::table::cell::value::Value;
/// One mutation in a transactional Pages table-cell batch.
pub type TableCellUpdate = litchi_iwa_common::table::cell::value::Update;
pub use crate::shapes::RgbaColor as TableCellTextColor;
pub use crate::text::Background as TableCellTextBackground;
pub use crate::text::Outline as TableCellTextOutline;
pub use crate::text::ParagraphList as TableCellParagraphList;
pub use crate::text::ParagraphListBullet as TableCellParagraphListBullet;
pub use crate::text::ParagraphListBulletGeometry as TableCellParagraphListBulletGeometry;
pub use crate::text::ParagraphListIndentation as TableCellParagraphListIndentation;
pub use crate::text::ParagraphListLabelColor as TableCellParagraphListLabelColor;
pub use crate::text::ParagraphListLevel as TableCellParagraphListLevel;
pub use crate::text::ParagraphListLevelPlacement as TableCellParagraphListLevelPlacement;
pub use crate::text::ParagraphListNumberFormat as TableCellParagraphListNumberFormat;
pub use crate::text::ParagraphListNumberScale as TableCellParagraphListNumberScale;
pub use crate::text::ParagraphListNumberTiering as TableCellParagraphListNumberTiering;
pub use crate::text::ParagraphListNumbering as TableCellParagraphListNumbering;
pub use crate::text::ParagraphListPlacement as TableCellParagraphListPlacement;
pub use crate::text::ParagraphTabStops as TableCellParagraphTabStops;
pub use crate::text::Shadow as TableCellTextShadow;
pub use crate::text::TextBaselineShift as TableCellTextBaselineShift;
pub use crate::text::TextCapitalization as TableCellTextCapitalization;
pub use crate::text::TextCharacterSpacing as TableCellTextCharacterSpacing;
pub use crate::text::TextDecorations as TableCellTextDecorations;
pub use crate::text::TextFont as TableCellTextFont;
pub use crate::text::TextLigatures as TableCellTextLigatures;
pub use crate::text::TextScript as TableCellTextScript;
pub use crate::text::TextStyle as TableCellTextStyle;
pub use litchi_iwa_common::shape::stroke::Stroke;
pub use litchi_iwa_common::table::cell::BorderSide;
pub use litchi_iwa_common::table::cell::Borders;
pub use litchi_iwa_common::table::cell::layout::{
    Inset as TableCellInset, Insets as TableCellInsets, Layout as TableCellLayout,
    TextWrap as TableCellTextWrap, VerticalAlignment as TableCellVerticalAlignment,
};

pub use semantic::{Table, TableInfo};
use storage::{PagesTableGraph, body_table_graphs, clone_body_table_attachment};
use validation::decode_unique_table_model;
