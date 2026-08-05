//! Typed Numbers editor state and semantic snapshot-edit operations.
#![allow(unused_imports)]

use super::formula_clone::remove_table_formula_graph;
use super::model::*;
use super::storage::*;
use super::table_duplicate::{
    duplicate_attached_table_graph_in_package, duplicate_table_name, table_owned_graph,
};
use super::{
    DOCUMENT_COMPONENT_IDENTIFIER, EMPTY_TABLE_POSITION_OFFSET, IWorkTableCellRegion,
    SHAPE_INFO_MESSAGE_TYPE, TABLE_DUPLICATE_OFFSET, cell_data_format, cell_fill, cell_layout,
    cell_merge, cell_paragraph_list, cell_paragraph_style, conditional_highlight, formula_cache,
    stroke_layers, table_bootstrap, table_cells, table_create, table_formula,
};

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use super::super::cell::{CellValue, TableCellUpdate};
use super::super::formula::{FormulaCachedValue, FormulaExpression, FormulaPivotCategoryReference};
use crate::archive::RawMessage;
use litchi_iwa_common::comment::{
    Comment, DrawableComment, DrawableId, DrawableInfo, DrawableReply, StorageId,
    TableCellComment, TableCellReply,
};
use crate::comments::IWorkDrawableCommentEditor;
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    component_identifier_for_entry, component_uuid_identifiers, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
    remove_component_registration,
};
use crate::protobuf::{tn, tswp};
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
    TextCapitalization, TextCharacterSpacing, TextColumns, TextComment, TextCommentBody,
    TextCommentId, TextCommentReply, TextCommentReplyBody, TextCommentReplyId, TextDecorations,
    TextFont, TextHighlight, TextHighlightId, TextHyperlink, TextHyperlinkId, TextHyperlinkTarget,
    TextLanguage, TextLanguageRun, TextLigatures, TextOutline, TextPosition, TextRange, TextScript,
    TextShadow, TextStorageInfo, TextStyle,
};
use crate::wire::{patch_length_delimited_field, patch_nested_length_delimited_field};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};

#[path = "semantic/drawables.rs"]
mod drawables;
#[path = "semantic/media.rs"]
mod media;
#[path = "semantic/model.rs"]
mod model;
#[path = "semantic/table.rs"]
mod table;
#[path = "semantic/text_box.rs"]
mod text_box;
#[path = "semantic/workbook.rs"]
mod workbook;

pub use model::*;
