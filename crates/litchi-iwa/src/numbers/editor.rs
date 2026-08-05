//! Transactional semantic editing for Numbers spreadsheets.
//!
//! This module is the stable Numbers-editor facade. Typed semantic state and
//! editor operations live in [`semantic`], while package-boundary adapters live
//! in [`package`]. The existing contextual owners below remain private seams
//! behind this facade so their snapshot-edit APIs continue to resolve through
//! the same `numbers::editor` namespace.

#![allow(unused_imports)]

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

pub use litchi_numbers::table::dimension::{Dimension, Points, Size};
use litchi_iwa_common::table::cell::number_format::NumberFormat;
use litchi_numbers::cell::data_format::{
    Checkbox, Currency, Custom, DataFormat, DateTime, Duration, Fraction, Number, NumeralSystem,
    Percentage, PopUpMenu, Scientific, Slider, StarRating, Stepper, Text,
};
use prost::Message;

use super::bnc::{BncCell, CachedScalar, StoredValue};
use super::cell::{CellValue, TableCellUpdate};
use super::formula::{
    ExternalFormulaTable, ExternalPivotCategory, FormulaCachedValue, FormulaExpression,
    FormulaPivotCategoryReference, FormulaUuid, PivotFormulaKey, compile_formula,
};
use litchi_iwa_common::comment::{
    AuthorId, Comment, DrawableComment, DrawableId, DrawableInfo, DrawableReply, ListId,
    StorageId, TableCellComment, TableCellReply, Uuid,
};
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::{
    IWorkDrawableCommentEditor, advance_save_tokens_for_entries, clone_comment_storage_exact,
    current_apple_reference_date, fresh_comment_storage_uuid, insert_comment_storage,
    preferred_or_ensure_table_annotation_author, remove_generated_annotation_author_if_unused,
    update_comment_reply_reference,
};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    component_uuid_identifiers, next_object_identifier, release_package_identifier_suffix,
    remove_component_external_reference, remove_component_external_references_to_object,
    remove_component_object_uuids, remove_component_registration,
    set_package_last_object_identifier,
};
use crate::protobuf::tst::{
    self, TableDataList, TableDataListSegment, TableModelArchive, Tile, TileRowInfo,
};
use crate::protobuf::{tn, tsce, tsd, tsp, tss, tswp};
use crate::shapes::{
    DrawableGeometry, DrawableProperties, RgbaColor, ShapeTextLayout, reset_shape_text_columns,
    reset_shape_text_layout, set_shape_geometry, set_shape_properties, set_shape_text_columns,
    set_shape_text_layout, shape_geometry, shape_properties, shape_text_columns, shape_text_layout,
};
use crate::table_appearance::TableAppearance;
use litchi_iwa_common::table::cell::conditional_highlight::Rule;
use crate::table_lock::TableLockState;
use crate::text::{
    IWorkTextEditor, ParagraphBackground, ParagraphBorders, ParagraphDecimalTabCharacter,
    ParagraphDefaultTabInterval, ParagraphDropCap, ParagraphDropCapPlacement, ParagraphFlow,
    ParagraphIndents, ParagraphLineSpacing, ParagraphList, ParagraphListBullet,
    ParagraphListBulletGeometry, ParagraphListIndentation, ParagraphListLabelColor,
    ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumberFormat,
    ParagraphListNumberScale, ParagraphListNumberTiering, ParagraphListNumbering,
    ParagraphListPlacement, ParagraphSpacing, ParagraphStart, ParagraphTabStops,
    ParagraphWritingDirection, TextAlignment, TextBackground, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextComment, TextCommentBody,
    TextCommentId, TextCommentReply, TextCommentReplyBody, TextCommentReplyId, TextDecorations,
    TextFont, TextHighlight, TextHighlightId, TextHyperlink, TextHyperlinkId, TextHyperlinkTarget,
    TextLanguage, TextLanguageRun, TextLigatures, TextOutline, TextPosition, TextRange, TextScript,
    TextShadow, TextStorageInfo, TextStyle,
};
use crate::wire::{
    patch_length_delimited_field, patch_nested_fixed32_field, patch_nested_length_delimited_field,
    patch_nested_varint_field, patch_varint_field, repeated_length_delimited_payloads,
    repeated_varint_values, rewrite_repeated_length_delimited_fields,
    rewrite_repeated_varint_fields, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};
use formula_clone::{
    clone_table_formula_graph, create_empty_table_formula_graph, formula_graph_owner_uuids,
    remap_cloned_formula_owner_storage, remap_cloned_formula_storage, remove_table_formula_graph,
    table_formula_graph_is_self_contained,
};

const MAX_TABLE_UIDS: usize = 1_100_000;
const HEADER_BUCKET_ROWS: usize = 65_536;
const FORMULA_DEPENDENCY_TILE_COLUMNS: u32 = 32;
const FORMULA_DEPENDENCY_TILE_ROWS: u32 = 128;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const DOCUMENT_COMPONENT_IDENTIFIER: u64 = 1;
const DRAWABLE_DUPLICATE_OFFSET: f32 = 10.0;
const TABLE_DUPLICATE_OFFSET: f32 = DRAWABLE_DUPLICATE_OFFSET;
const EMPTY_TABLE_POSITION_OFFSET: f32 = 40.0;
const CONDITIONAL_STYLE_NO_APPLIED_RULE: u32 = 15;

/// Numbers-owned table editing vocabulary.
pub mod table {
    /// Cell-level table editing vocabulary.
    pub mod cell {
        use crate::shapes::ShapeStroke;
        use litchi_iwa_common::table::cell::BorderSide;

        /// Effective explicit borders stored for one native Numbers table cell.
        ///
        /// `None` means the table style supplies the edge, or a later native
        /// stroke run explicitly clears it.
        #[derive(Clone, Copy, Debug, Default, PartialEq)]
        pub struct Borders {
            pub left: Option<ShapeStroke>,
            pub right: Option<ShapeStroke>,
            pub top: Option<ShapeStroke>,
            pub bottom: Option<ShapeStroke>,
        }

        impl Borders {
            pub const fn get(self, side: BorderSide) -> Option<ShapeStroke> {
                match side {
                    BorderSide::Left => self.left,
                    BorderSide::Right => self.right,
                    BorderSide::Top => self.top,
                    BorderSide::Bottom => self.bottom,
                }
            }

            pub(crate) fn set(&mut self, side: BorderSide, stroke: Option<ShapeStroke>) {
                match side {
                    BorderSide::Left => self.left = stroke,
                    BorderSide::Right => self.right = stroke,
                    BorderSide::Top => self.top = stroke,
                    BorderSide::Bottom => self.bottom = stroke,
                }
            }
        }
    }
}

#[derive(Debug, Clone)]
struct NumbersTextBoxGraph {
    sheet_id: u64,
    archive_name: String,
    drawable_id: u64,
    storage_id: u64,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

#[path = "editor/package.rs"]
mod package;
#[path = "editor/semantic.rs"]
mod semantic;

mod cell_data_format;
mod cell_fill;
mod cell_layout;
mod cell_merge;
mod cell_paragraph_list;
mod cell_paragraph_style;
mod cell_style;
mod column_insert;
mod conditional_highlight;
mod date_time_fields;
mod drawable_order;
mod formula_cache;
mod formula_clone;
mod formula_dependency_shift;
mod model;
mod named_paragraph_styles;
mod row_insert;
mod sheet_audio;
mod sheet_charts;
mod sheet_delete;
mod sheet_duplicate;
mod sheet_images;
mod sheet_movies;
mod sheet_shapes;
mod storage;
mod stroke_layers;
mod table_appearance;
mod table_axis_deletion;
mod table_axis_insertion;
mod table_bootstrap;
mod table_cells;
mod table_create;
mod table_delete;
mod table_dimension;
mod table_duplicate;
mod table_formula;
mod table_headers;
mod table_hidden_axes;
mod table_lock;
mod table_move;
mod table_sort;
mod table_sparse_storage;
mod table_title;
mod table_topology;
mod text_box_create;
mod text_box_duplicate;

pub use crate::charts::Direction;
pub use cell_merge::IWorkTableCellRegion;
pub use litchi_numbers::table::title::Settings;
use model::*;
pub use semantic::*;
pub use sheet_audio::{NumbersSheetAudioInfo, NumbersSheetAudioOptions, RemovedNumbersSheetAudio};
pub use sheet_charts::{NumbersSheetChartInfo, RemovedNumbersSheetChart};
pub use sheet_images::{NumbersSheetImageInfo, NumbersSheetImageOptions, RemovedNumbersSheetImage};
pub use sheet_movies::{NumbersSheetMovieInfo, NumbersSheetMovieOptions, RemovedNumbersSheetMovie};
pub use sheet_shapes::{NumbersSheetShapeInfo, NumbersSheetShapeKind, RemovedNumbersSheetShape};
use storage::*;
pub use table_axis_deletion::{TableColumnDeletion, TableRowDeletion};
pub use table_axis_insertion::{TableColumnInsertion, TableRowInsertion};
pub(crate) use table_cells::TableCellBatch;
pub(crate) use table_duplicate::{duplicate_attached_table_graph_in_package, duplicate_table_name};
use table_duplicate::{
    register_cloned_numbers_objects, register_numbers_component_reference, table_owned_graph,
};
pub use table_headers::{NumbersTableHeaderCount, NumbersTableHeaderSettings};
pub use table_sort::{
    NumbersTableSortColumnIndex, NumbersTableSortDirection, NumbersTableSortOrder,
    NumbersTableSortRowRange, NumbersTableSortRule, NumbersTableSortScope,
};
pub(crate) use table_sort::{
    apply_table_sort_order_in_package, apply_table_sort_order_to_rows_in_package,
    clear_table_sort_order_in_package, set_table_sort_order_in_package,
    table_sort_order_in_package,
};
pub(crate) use table_title::{
    set_table_title_settings_in_package, table_title_settings_in_package,
};

pub(crate) use package::*;

#[cfg(test)]
mod tests;
