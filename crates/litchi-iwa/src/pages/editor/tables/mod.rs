//! Native table discovery and cell editing for Pages body attachments.

mod appearance;
mod comments;
mod conditional_highlight;
mod formula;
mod hidden_axes;
mod layout;
mod lock;
mod semantic;
mod sort;
mod storage;
mod title;
mod topology;
mod validation;

#[cfg(test)]
mod tests;

pub use conditional_highlight::PagesTableCellConditionalHighlightInfo;
pub use formula::{
    PagesTableFormulaAxisReference, PagesTableFormulaBinaryOperator, PagesTableFormulaCachedValue,
    PagesTableFormulaCellReference, PagesTableFormulaExpression,
};
pub use hidden_axes::{PagesTableAxisIndex, PagesTableHiddenAxes};
pub use layout::{
    PagesTableDimension, PagesTableDimensionSize, PagesTablePoints,
};
pub use sort::{
    PagesTableSortColumnIndex, PagesTableSortDirection, PagesTableSortOrder,
    PagesTableSortRowRange, PagesTableSortRule, PagesTableSortScope,
};
pub use title::PagesTableTitleSettings;

use std::collections::{HashMap, HashSet};

use prost::Message;

use super::*;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::protobuf::tst::TableInfoArchive;
use crate::table_appearance::TableAppearance;
use crate::table_lock::table_lock_state_from_message;
use litchi_iwa_common::table::lock::State as TableLockState;
use litchi_numbers::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;
const INLINE_TABLE_DUPLICATE_OFFSET: f32 = 0.0;

/// Strongly typed cell value shared by Pages and Numbers table storage.
pub type PagesCellValue = litchi_numbers::cell::Value;
/// One mutation in a transactional Pages table-cell batch.
pub type PagesTableCellUpdate = litchi_numbers::cell::Update;
pub use crate::shapes::RgbaColor as PagesTableCellTextColor;
pub use crate::table_cell_border::{
    TableCellBorderSide as PagesTableCellBorderSide, TableCellBorders as PagesTableCellBorders,
};
pub use litchi_iwa_common::table::cell::layout::{
    Inset as PagesTableCellInset, Insets as PagesTableCellInsets, Layout as PagesTableCellLayout,
    TextWrap as PagesTableCellTextWrap, VerticalAlignment as PagesTableCellVerticalAlignment,
};
pub use crate::text::ParagraphIndents as PagesTableCellParagraphIndents;
pub use crate::text::ParagraphLineSpacing as PagesTableCellParagraphLineSpacing;
pub use crate::text::ParagraphList as PagesTableCellParagraphList;
pub use crate::text::ParagraphListBullet as PagesTableCellParagraphListBullet;
pub use crate::text::ParagraphListBulletGeometry as PagesTableCellParagraphListBulletGeometry;
pub use crate::text::ParagraphListIndentation as PagesTableCellParagraphListIndentation;
pub use crate::text::ParagraphListLabelColor as PagesTableCellParagraphListLabelColor;
pub use crate::text::ParagraphListLevel as PagesTableCellParagraphListLevel;
pub use crate::text::ParagraphListLevelPlacement as PagesTableCellParagraphListLevelPlacement;
pub use crate::text::ParagraphListNumberFormat as PagesTableCellParagraphListNumberFormat;
pub use crate::text::ParagraphListNumberScale as PagesTableCellParagraphListNumberScale;
pub use crate::text::ParagraphListNumberTiering as PagesTableCellParagraphListNumberTiering;
pub use crate::text::ParagraphListNumbering as PagesTableCellParagraphListNumbering;
pub use crate::text::ParagraphListPlacement as PagesTableCellParagraphListPlacement;
pub use crate::text::ParagraphSpacing as PagesTableCellParagraphSpacing;
pub use crate::text::ParagraphTabStops as PagesTableCellParagraphTabStops;
pub use crate::text::TextAlignment as PagesTableCellTextAlignment;
pub use crate::text::Background as PagesTableCellTextBackground;
pub use crate::text::TextBaselineShift as PagesTableCellTextBaselineShift;
pub use crate::text::TextCapitalization as PagesTableCellTextCapitalization;
pub use crate::text::TextCharacterSpacing as PagesTableCellTextCharacterSpacing;
pub use crate::text::TextDecorations as PagesTableCellTextDecorations;
pub use crate::text::TextFont as PagesTableCellTextFont;
pub use crate::text::TextLigatures as PagesTableCellTextLigatures;
pub use crate::text::Outline as PagesTableCellTextOutline;
pub use crate::text::TextScript as PagesTableCellTextScript;
pub use crate::text::Shadow as PagesTableCellTextShadow;
pub use crate::text::TextStyle as PagesTableCellTextStyle;

pub use semantic::{PagesTable, PagesTableInfo};
use storage::{
    PagesTableGraph, body_table_graphs, clone_body_table_attachment, remove_table_object,
};
use validation::decode_table_models;
