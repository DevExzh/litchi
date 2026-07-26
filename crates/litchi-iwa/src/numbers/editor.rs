//! Transactional semantic editing for Numbers spreadsheets.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use super::bnc::{BncCell, StoredValue};
use super::cell::{CellValue, TableCellUpdate};
use super::formula::{
    ExternalFormulaTable, ExternalPivotCategory, FormulaCachedValue, FormulaExpression,
    FormulaPivotCategoryReference, FormulaUuid, PivotFormulaKey,
};
use super::table::{NumbersCellComment, NumbersCommentUuid};
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::{
    DrawableCommentInfo, DrawableCommentReplyInfo, IWorkDrawableCommentEditor, IWorkDrawableInfo,
    IWorkTableCellCommentInfo, IWorkTableCellCommentReplyInfo, advance_save_tokens_for_entries,
    clone_comment_storage_exact, current_apple_reference_date, fresh_comment_storage_uuid,
    insert_comment_storage, preferred_or_ensure_table_annotation_author,
    remove_generated_annotation_author_if_unused, update_comment_reply_reference,
};
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    component_uuid_identifiers, next_object_identifier, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
    remove_component_registration, set_package_last_object_identifier,
};
use crate::protobuf::tst::{
    self, TableDataList, TableDataListSegment, TableModelArchive, Tile, TileRowInfo,
};
use crate::protobuf::{tn, tsce, tsd, tsp, tswp};
use crate::registry::{Application, detect_application_from_document};
use crate::shapes::{
    DrawableGeometry, DrawableProperties, RgbaColor, ShapeTextLayout, reset_shape_text_columns,
    reset_shape_text_layout, set_shape_geometry, set_shape_properties, set_shape_text_columns,
    set_shape_text_layout, shape_geometry, shape_properties, shape_text_columns, shape_text_layout,
};
use crate::table_appearance::TableAppearance;
use crate::table_lock::TableLockState;
use crate::text::{
    IWorkTextEditor, ParagraphDropCap, ParagraphDropCapPlacement, ParagraphIndents,
    ParagraphLineSpacing, ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement,
    ParagraphSpacing, ParagraphStart, ParagraphTabStops, TextAlignment, TextBackground,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextColumns, TextComment,
    TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody, TextCommentReplyId,
    TextDecorations, TextFont, TextHighlight, TextHighlightId, TextHyperlink, TextHyperlinkId,
    TextHyperlinkTarget, TextLanguage, TextLanguageRun, TextLigatures, TextOutline, TextPosition,
    TextRange, TextScript, TextShadow, TextStorageInfo, TextStyle,
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

/// Stable identity and dimensions of a Numbers table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersTableInfo {
    pub object_id: u64,
    pub name: String,
    pub rows: usize,
    pub columns: usize,
    /// Effective alternating-row and automatic-sizing settings.
    pub appearance: TableAppearance,
    /// Interactive editing lock shown in the Arrange inspector.
    pub lock_state: TableLockState,
}

/// Stable identity and name of a sheet in workbook order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersSheetInfo {
    pub object_id: u64,
    pub index: usize,
    pub name: String,
}

/// A writable ordinary text box owned by one Numbers sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersTextBoxInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// A Numbers text box removed from a sheet with its final text state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedNumbersTextBox {
    pub text_box: NumbersTextBoxInfo,
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

/// A pivot aggregate category that can be used in a formula expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersPivotCategoryInfo {
    pub reference: FormulaPivotCategoryReference,
    pub label: Option<String>,
}

/// Address and storage identity of a comment attached to a Numbers cell.
pub type NumbersCellCommentInfo = IWorkTableCellCommentInfo;

/// A resolved direct reply in a Numbers cell-comment thread.
pub type NumbersCellCommentReplyInfo = IWorkTableCellCommentReplyInfo;

/// Mutable, transactional Numbers package editor.
///
/// Each semantic edit is applied to a cloned package and committed only after
/// all affected IWA components serialize successfully.
#[derive(Debug, Clone)]
pub struct NumbersEditor {
    package: IWorkPackage,
}

impl NumbersEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        numbers_document(&package)?;
        Ok(Self { package })
    }

    pub fn sheets(&self) -> Result<Vec<NumbersSheetInfo>> {
        let document = numbers_document(&self.package)?;
        let locations = object_locations(&self.package)?;
        document
            .sheets
            .into_iter()
            .enumerate()
            .map(|(index, reference)| {
                let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers sheet object {} is missing",
                        reference.identifier
                    ))
                })?;
                let archive = self.package.archive(archive_name)?;
                let object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers sheet object {} is missing",
                        reference.identifier
                    ))
                })?;
                let (_, sheet) = decode_sheet(object)?;
                Ok(NumbersSheetInfo {
                    object_id: reference.identifier,
                    index,
                    name: sheet.name,
                })
            })
            .collect()
    }

    pub fn tables(&self) -> Result<Vec<NumbersTableInfo>> {
        let locations = object_locations(&self.package)?;
        let mut tables = table_models(&self.package)?
            .into_iter()
            .map(|descriptor| {
                let archive_name = locations.get(&descriptor.table_info_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table drawable {} is missing",
                        descriptor.table_info_id
                    ))
                })?;
                Ok(NumbersTableInfo {
                    object_id: descriptor.object_id,
                    name: descriptor.model.table_name,
                    rows: descriptor.model.number_of_rows as usize,
                    columns: descriptor.model.number_of_columns as usize,
                    appearance: crate::table_appearance::table_appearance(
                        &self.package,
                        descriptor.object_id,
                    )?,
                    lock_state: crate::table_lock::table_lock_state_for_model(
                        &self.package,
                        archive_name,
                        descriptor.table_info_id,
                        descriptor.object_id,
                    )?,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        tables.sort_by_key(|table| table.object_id);
        Ok(tables)
    }

    /// List supported direct-comment drawables owned by one reachable sheet.
    pub fn sheet_drawables(&self, sheet_id: u64) -> Result<Vec<IWorkDrawableInfo>> {
        let owned = self.sheet_owned_drawable_ids(sheet_id)?;
        let mut drawables = IWorkDrawableCommentEditor::from_package(self.package.clone())?
            .drawables()?
            .into_iter()
            .filter(|drawable| owned.contains(&drawable.object_id))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.object_id);
        Ok(drawables)
    }

    /// Read a comment attached directly to a drawable owned by one sheet.
    pub fn sheet_drawable_comment(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<DrawableCommentInfo>> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package.clone())?.comment(drawable_object_id)
    }

    /// Create or replace a direct comment on a drawable owned by one sheet.
    pub fn set_sheet_drawable_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.set_comment(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Delete a direct comment from a drawable owned by one sheet.
    pub fn clear_sheet_drawable_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Read direct replies in a comment thread on one sheet drawable.
    pub fn sheet_drawable_comment_replies(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<DrawableCommentReplyInfo>> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package.clone())?.replies(drawable_object_id)
    }

    /// Add a reply to a direct comment on one sheet drawable.
    pub fn add_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        let reply_id = comments.add_reply(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Update a direct reply, returning its current storage identifier.
    pub fn set_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        let reply_id = comments.set_reply(drawable_object_id, reply_storage_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Remove a direct reply from a comment on one sheet drawable.
    pub fn remove_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.remove_reply(drawable_object_id, reply_storage_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// List ordinary text boxes owned by a reachable Numbers sheet.
    pub fn sheet_text_boxes(&self, sheet_id: u64) -> Result<Vec<NumbersTextBoxInfo>> {
        let (_, _, sheet) = numbers_sheet(&self.package, sheet_id)?;
        let locations = object_locations(&self.package)?;
        let text_editor = IWorkTextEditor::from_package(self.package.clone());
        let mut result = Vec::new();
        for reference in sheet.drawable_infos {
            let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} drawable {} is missing",
                    reference.identifier
                ))
            })?;
            let archive = self.package.archive(archive_name)?;
            let object = archive.object(reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} drawable {} is missing",
                    reference.identifier
                ))
            })?;
            let shape_messages = object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
                .collect::<Vec<_>>();
            if shape_messages.is_empty() {
                continue;
            }
            if shape_messages.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} has multiple shape payloads",
                    reference.identifier
                )));
            }
            let shape = tswp::ShapeInfoArchive::decode(shape_messages[0].data.as_slice())?;
            if shape.is_text_box != Some(true) {
                continue;
            }
            let graph = numbers_text_box_graph(&self.package, sheet_id, reference.identifier)?;
            let storage = text_editor.storage(graph.storage_id)?;
            result.push(NumbersTextBoxInfo {
                sheet_id,
                drawable_object_id: reference.identifier,
                storage,
            });
        }
        Ok(result)
    }

    /// Replace a UTF-16 range in an ordinary Numbers text box.
    pub fn replace_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.replace_text(graph.storage_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        numbers_text_box_graph(verified.package(), sheet_id, drawable_object_id)?;
        self.package = verified.package;
        Ok(())
    }

    /// Replace all text in an ordinary Numbers text box.
    pub fn set_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text(graph.storage_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = verified
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers text-box update lost its drawable".to_owned())
            })?;
        if updated.storage.text != replacement {
            return Err(Error::InvalidFormat(
                "Numbers text-box update failed validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }

    /// Clear an ordinary Numbers text box without deleting it.
    pub fn clear_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.set_sheet_text_box_text(sheet_id, drawable_object_id, "")
    }

    /// Read the geometry of an ordinary Numbers text box.
    pub fn sheet_text_box_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_geometry(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Update position, size, flags, and rotation on an ordinary Numbers text box.
    pub fn set_sheet_text_box_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_geometry(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers text-box geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties from an ordinary Numbers text box.
    pub fn sheet_text_box_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_properties(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Update shared drawable properties on an ordinary Numbers text box.
    pub fn set_sheet_text_box_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_properties(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers text-box properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read vertical alignment, edge insets, and autosizing for a text box.
    pub fn sheet_text_box_text_layout(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ShapeTextLayout> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_text_layout(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Replace text-frame layout while preserving text, columns, and drawing style.
    pub fn set_sheet_text_box_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        layout: ShapeTextLayout,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_text_layout(sheet_id, drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers text-box layout update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove crate-authored text-frame layout overrides.
    pub fn reset_sheet_text_box_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_layout(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the uniform column layout of an ordinary sheet text box.
    pub fn sheet_text_box_columns(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextColumns> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_text_columns(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Replace the uniform column layout of an ordinary sheet text box.
    pub fn set_sheet_text_box_columns(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        columns: &TextColumns,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let staged = set_shape_text_columns(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
            columns,
        )?;
        let verified = Self::from_package(staged)?;
        if &verified.sheet_text_box_columns(sheet_id, drawable_object_id)? != columns {
            return Err(Error::InvalidFormat(
                "Numbers text-box column update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited column layout after a crate-authored override.
    pub fn reset_sheet_text_box_columns(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_columns(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn sheet_text_box_text_style(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_style(graph.storage_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    pub fn set_sheet_text_box_text_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        style: TextStyle,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_style(graph.storage_id, style)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_style(sheet_id, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Numbers text-box character formatting update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited character formatting while preserving paragraph overrides.
    pub fn reset_sheet_text_box_text_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_style(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of a sheet-owned text box.
    pub fn sheet_text_box_text_font(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextFont> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_font(graph.storage_id)
    }

    /// Atomically set a typed font identity across a sheet-owned text box.
    pub fn set_sheet_text_box_text_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        font: TextFont,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_font(graph.storage_id, font)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore the inherited font while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_font(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every explicit language boundary in a sheet-owned text box.
    pub fn sheet_text_box_text_languages(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextLanguageRun>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_languages(graph.storage_id)
    }

    /// Read the effective language at one UTF-16 text boundary.
    pub fn sheet_text_box_text_language(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<TextLanguage> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .text_language(graph.storage_id, position)
    }

    /// Atomically create or update one text-language boundary.
    pub fn set_sheet_text_box_text_language(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_language(graph.storage_id, position, language)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Delete one nonzero language boundary so it inherits the preceding run.
    pub fn remove_sheet_text_box_text_language_boundary(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.remove_text_language_boundary(graph.storage_id, position)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Restore automatic language selection across a sheet-owned text box.
    pub fn reset_sheet_text_box_text_languages(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_languages(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every hyperlink in a sheet-owned text box.
    pub fn sheet_text_box_hyperlinks(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHyperlink>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_hyperlinks(graph.storage_id)
    }

    /// Create a hyperlink over a nonempty, unoccupied UTF-16 text range.
    pub fn add_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.add_text_hyperlink(graph.storage_id, range, target)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Update a text-box hyperlink's range and target without changing its ID.
    pub fn update_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.update_text_hyperlink(graph.storage_id, id, range, target)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Delete a text-box hyperlink and its owned smart-field object.
    pub fn remove_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.remove_text_hyperlink(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Read every plain highlight in a sheet-owned text box.
    pub fn sheet_text_box_highlights(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHighlight>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_highlights(graph.storage_id)
    }

    /// Create a plain highlight over a nonempty UTF-16 text range.
    pub fn add_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.add_text_highlight(graph.storage_id, range)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Move a plain text-box highlight without changing its ID.
    pub fn update_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.update_text_highlight(graph.storage_id, id, range)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Delete a plain text-box highlight and its empty annotation graph.
    pub fn remove_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.remove_text_highlight(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Read every ranged comment in a sheet-owned text box.
    pub fn sheet_text_box_comments(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextComment>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_comments(graph.storage_id)
    }

    /// Create a ranged comment in a sheet-owned text box.
    pub fn add_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.add_text_comment(graph.storage_id, range, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Update a text-box comment's range and body without changing its ID.
    pub fn update_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.update_text_comment(graph.storage_id, id, range, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Delete a ranged text-box comment and its owned annotation graph.
    pub fn remove_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextCommentId,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.remove_text_comment(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Read every direct reply to a sheet text-box comment in stored order.
    pub fn sheet_text_box_comment_replies(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .text_comment_replies(graph.storage_id, comment_id)
    }

    /// Append a direct reply to a sheet text-box comment.
    pub fn add_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.add_text_comment_reply(graph.storage_id, comment_id, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Update a direct sheet text-box comment reply without changing its ID.
    pub fn update_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.update_text_comment_reply(graph.storage_id, comment_id, reply_id, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Delete one direct sheet text-box comment reply and its storage.
    pub fn remove_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.remove_text_comment_reply(graph.storage_id, comment_id, reply_id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Read the canonical list preset of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_list(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphList> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_list(graph.storage_id)
    }

    /// Atomically apply a canonical list preset to a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_list(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        list: ParagraphList,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list(graph.storage_id, list)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Remove list formatting from a sheet-owned text box.
    pub fn reset_sheet_text_box_paragraph_list(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every list-level boundary in a sheet-owned text box.
    pub fn sheet_text_box_paragraph_list_levels(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_list_levels(graph.storage_id)
    }

    /// Read one paragraph's effective list nesting level.
    pub fn sheet_text_box_paragraph_list_level(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListLevel> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_level(graph.storage_id, paragraph)
    }

    /// Atomically set one paragraph's list nesting level.
    pub fn set_sheet_text_box_paragraph_list_level(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        level: ParagraphListLevel,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_level(graph.storage_id, paragraph, level)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_sheet_text_box_paragraph_list_level(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_level(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform underline and strikethrough formatting.
    pub fn sheet_text_box_text_decorations(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextDecorations> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_decorations(graph.storage_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    pub fn set_sheet_text_box_text_decorations(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        decorations: TextDecorations,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_decorations(graph.storage_id, decorations)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_decorations(sheet_id, drawable_object_id)? != decorations {
            return Err(Error::InvalidFormat(
                "Numbers text-box decoration update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited decorations while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_decorations(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_decorations(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective uniform text color of a sheet-owned text box.
    pub fn sheet_text_box_text_color(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RgbaColor> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_color(graph.storage_id)
    }

    /// Atomically set one text color across a sheet-owned text box.
    pub fn set_sheet_text_box_text_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        color: RgbaColor,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_color(graph.storage_id, color)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_color(sheet_id, drawable_object_id)? != color {
            return Err(Error::InvalidFormat(
                "Numbers text-box color update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited text color while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_color(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform capitalization from a sheet-owned text box.
    pub fn sheet_text_box_text_capitalization(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextCapitalization> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_capitalization(graph.storage_id)
    }

    /// Atomically set one capitalization mode across a sheet-owned text box.
    pub fn set_sheet_text_box_text_capitalization(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        capitalization: TextCapitalization,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_capitalization(graph.storage_id, capitalization)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_capitalization(sheet_id, drawable_object_id)?
            != capitalization
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box capitalization update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited capitalization while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_capitalization(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_capitalization(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform baseline script from a sheet-owned text box.
    pub fn sheet_text_box_text_script(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextScript> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_script(graph.storage_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    pub fn set_sheet_text_box_text_script(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        script: TextScript,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_script(graph.storage_id, script)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_script(sheet_id, drawable_object_id)? != script {
            return Err(Error::InvalidFormat(
                "Numbers text-box script update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited baseline script while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_script(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_script(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of a sheet-owned text box.
    pub fn sheet_text_box_text_baseline_shift(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextBaselineShift> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_baseline_shift(graph.storage_id)
    }

    /// Atomically set a signed custom baseline displacement.
    pub fn set_sheet_text_box_text_baseline_shift(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shift: TextBaselineShift,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_baseline_shift(graph.storage_id, shift)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_baseline_shift(sheet_id, drawable_object_id)? != shift {
            return Err(Error::InvalidFormat(
                "Numbers text-box baseline-shift update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited baseline displacement while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_baseline_shift(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_baseline_shift(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of a sheet-owned text box.
    pub fn sheet_text_box_text_character_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextCharacterSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_character_spacing(graph.storage_id)
    }

    /// Atomically set character spacing across a sheet-owned text box.
    pub fn set_sheet_text_box_text_character_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_character_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_character_spacing(sheet_id, drawable_object_id)? != spacing
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box character-spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited character spacing while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_character_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_character_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of a sheet-owned text box.
    pub fn sheet_text_box_text_ligatures(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextLigatures> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_ligatures(graph.storage_id)
    }

    /// Atomically set the ligature policy across a sheet-owned text box.
    pub fn set_sheet_text_box_text_ligatures(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        ligatures: TextLigatures,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_ligatures(graph.storage_id, ligatures)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_ligatures(sheet_id, drawable_object_id)? != ligatures {
            return Err(Error::InvalidFormat(
                "Numbers text-box ligature update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited ligatures while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_ligatures(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_ligatures(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective outline of a sheet-owned text box.
    pub fn sheet_text_box_text_outline(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextOutline> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_outline(graph.storage_id)
    }

    /// Atomically set a typed outline across a sheet-owned text box.
    pub fn set_sheet_text_box_text_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        outline: TextOutline,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_outline(graph.storage_id, outline)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_outline(sheet_id, drawable_object_id)? != outline {
            return Err(Error::InvalidFormat(
                "Numbers text-box outline update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited outline while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_outline(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective shadow of a sheet-owned text box.
    pub fn sheet_text_box_text_shadow(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextShadow> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_shadow(graph.storage_id)
    }

    /// Atomically set a typed drop shadow across a sheet-owned text box.
    pub fn set_sheet_text_box_text_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shadow: TextShadow,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_shadow(graph.storage_id, shadow)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_shadow(sheet_id, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Numbers text-box shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited shadow while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_shadow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective solid background of a sheet-owned text box.
    pub fn sheet_text_box_text_background(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextBackground> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_background(graph.storage_id)
    }

    /// Atomically set a solid background across a sheet-owned text box.
    pub fn set_sheet_text_box_text_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        background: TextBackground,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_background(graph.storage_id, background)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_background(sheet_id, drawable_object_id)? != background {
            return Err(Error::InvalidFormat(
                "Numbers text-box background update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited text background while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph alignment of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_alignment(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextAlignment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_alignment(graph.storage_id)
    }

    /// Set one paragraph alignment across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_alignment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        alignment: TextAlignment,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_alignment(graph.storage_id, alignment)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_alignment(sheet_id, drawable_object_id)? != alignment {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph-alignment update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited paragraph alignment after a private minimal override.
    pub fn reset_sheet_text_box_paragraph_alignment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_alignment(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective line spacing of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_line_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphLineSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_line_spacing(graph.storage_id)
    }

    /// Set one typed line-spacing mode across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_line_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: ParagraphLineSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_line_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_line_spacing(sheet_id, drawable_object_id)? != spacing
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph line-spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited line spacing while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_line_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_line_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_spacing(graph.storage_id)
    }

    /// Atomically set before/after paragraph spacing across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: ParagraphSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_spacing(sheet_id, drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited paragraph spacing while preserving sibling overrides.
    pub fn reset_sheet_text_box_paragraph_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right indentation of a sheet text box.
    pub fn sheet_text_box_paragraph_indents(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphIndents> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_indents(graph.storage_id)
    }

    /// Atomically set paragraph indentation across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_indents(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        indents: ParagraphIndents,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_indents(graph.storage_id, indents)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_indents(sheet_id, drawable_object_id)? != indents {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph indentation update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited indentation while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_indents(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_indents(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of a sheet text box.
    pub fn sheet_text_box_paragraph_tab_stops(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphTabStops> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_tab_stops(graph.storage_id)
    }

    /// Atomically replace every explicit ruler tab stop of a sheet text box.
    pub fn set_sheet_text_box_paragraph_tab_stops(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        stops: ParagraphTabStops,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_tab_stops(graph.storage_id, stops)?;
        let expected = text.paragraph_tab_stops(graph.storage_id)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_tab_stops(sheet_id, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph tab-stop update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited tab stops while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_tab_stops(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_tab_stops(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// List every Drop Cap in a sheet-owned text box.
    pub fn sheet_text_box_paragraph_drop_caps(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphDropCapPlacement>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_drop_caps(graph.storage_id)
    }

    /// Read the Drop Cap attached to one text-box paragraph.
    pub fn sheet_text_box_paragraph_drop_cap(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<Option<ParagraphDropCap>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_drop_cap(graph.storage_id, paragraph_start)
    }

    /// Atomically create or replace a text-box Drop Cap.
    pub fn set_sheet_text_box_paragraph_drop_cap(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
        drop_cap: ParagraphDropCap,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_drop_cap(graph.storage_id, paragraph_start, drop_cap)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_drop_cap(
            sheet_id,
            drawable_object_id,
            paragraph_start,
        )? != Some(drop_cap)
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box Drop Cap update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Atomically remove a text-box Drop Cap.
    pub fn remove_sheet_text_box_paragraph_drop_cap(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.remove_paragraph_drop_cap(graph.storage_id, paragraph_start)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Remove an ordinary sheet-owned text box and its private object graph.
    pub fn remove_sheet_text_box(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersTextBox> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let text_box = self
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers text box {drawable_object_id} lost its writable storage"
                ))
            })?;

        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &graph.archive_name,
            graph.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        staged.update_archive(&graph.archive_name, |archive| {
            for identifier in &graph.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers text-box object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &graph.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers text-box object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(
            &mut staged,
            DOCUMENT_COMPONENT_IDENTIFIER,
            &graph.uuid_object_ids,
        )?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_package(staged)?;
        if verified
            .sheet_text_boxes(sheet_id)?
            .iter()
            .any(|item| item.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersTextBox { text_box })
    }

    /// List absolute pivot categories backed by valid calculation-engine
    /// aggregate coordinates.
    pub fn pivot_categories(&self) -> Result<Vec<NumbersPivotCategoryInfo>> {
        let mut categories = formula_pivot_categories(&self.package)?
            .into_iter()
            .map(|(key, value)| NumbersPivotCategoryInfo {
                reference: FormulaPivotCategoryReference::new(
                    key.group_by_uid,
                    key.column_uid,
                    key.group_uid,
                    value.aggregate_type,
                    value.group_level,
                ),
                label: value.label,
            })
            .collect::<Vec<_>>();
        categories.sort_by(|left, right| {
            left.reference
                .group_by_uid
                .cmp(&right.reference.group_by_uid)
                .then_with(|| left.reference.column_uid.cmp(&right.reference.column_uid))
                .then_with(|| left.reference.group_level.cmp(&right.reference.group_level))
                .then_with(|| left.label.cmp(&right.label))
                .then_with(|| left.reference.group_uid.cmp(&right.reference.group_uid))
        });
        Ok(categories)
    }

    /// Set or clear a cell in a table identified by its IWA object ID.
    ///
    /// Cached results of dependent numeric/Boolean formulas are refreshed in
    /// dependency order. If an impacted formula is outside the strict local
    /// evaluator subset, the entire edit is rejected without changing the
    /// package rather than persisting a stale displayed result.
    pub fn set_cell(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        value: CellValue,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_cell_in_package(&mut staged, table_id, row, column, value)?;
        formula_cache::refresh_formula_caches_after_cell_write(&mut staged, table_id, row, column)?;
        // Exercise every serialization boundary before committing the edit.
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Set several cells in one table as one transaction.
    ///
    /// The batch must contain unique coordinates. It clones and serializes the
    /// package once, reuses one table/object lookup context for every cell, and
    /// refreshes all impacted formula caches from the final batch state in one
    /// dependency pass. The returned count equals the number of applied cells.
    pub fn set_cells(
        &mut self,
        table_id: u64,
        updates: impl IntoIterator<Item = TableCellUpdate>,
    ) -> Result<usize> {
        let batch = table_cells::TableCellBatch::collect(updates)?;
        if batch.is_empty() {
            attached_table_descriptor(&self.package, table_id)?;
            return Ok(0);
        }
        let expected = batch.len();
        let mut staged = self.package.clone();
        let applied = batch.apply_numbers(&mut staged, table_id)?;
        if applied != expected {
            return Err(Error::InvalidFormat(format!(
                "Table cell batch applied {applied} updates, expected {expected}"
            )));
        }
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(applied)
    }

    pub fn clear_cell(&mut self, table_id: u64, row: usize, column: usize) -> Result<()> {
        self.set_cell(table_id, row, column, CellValue::Empty)
    }

    /// Read the explicit data format for one zero-based table cell.
    pub fn table_cell_data_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::table_cell_data_format::TableCellDataFormat> {
        cell_data_format::cell_data_format(&self.package, table_id, row, column)
    }

    /// Create, replace, or reset one cell's typed data format transactionally.
    pub fn set_table_cell_data_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: crate::table_cell_data_format::TableCellDataFormat,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_data_format::set_cell_data_format(&mut staged, table_id, row, column, format)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_data_format(table_id, row, column)? != format {
            return Err(Error::InvalidFormat(
                "Numbers table-cell data format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read an explicit decimal-number format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_number_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<crate::table_cell_number_format::TableCellNumberFormat>> {
        cell_data_format::cell_number_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit decimal-number format transactionally.
    pub fn set_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: crate::table_cell_number_format::TableCellNumberFormat,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore iWork's automatic data format for one table cell.
    pub fn reset_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_number_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified
                .table_cell_number_format(table_id, row, column)?
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "Numbers table-cell number-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit currency format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_currency_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<crate::table_cell_data_format::TableCellCurrencyFormat>> {
        cell_data_format::cell_currency_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit currency format transactionally.
    pub fn set_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: crate::table_cell_data_format::TableCellCurrencyFormat,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Currency cell.
    pub fn reset_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_currency_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != crate::table_cell_data_format::TableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers currency-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit percentage format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_percentage_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<crate::table_cell_data_format::TableCellPercentageFormat>> {
        cell_data_format::cell_percentage_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit percentage format transactionally.
    pub fn set_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: crate::table_cell_data_format::TableCellPercentageFormat,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore iWork's automatic format from an explicit Percentage cell.
    pub fn reset_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_percentage_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != crate::table_cell_data_format::TableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers percentage-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective text layout for one zero-based table cell.
    pub fn table_cell_layout(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::table_cell_layout::TableCellLayout> {
        cell_layout::cell_layout(&self.package, table_id, row, column)
    }

    /// Create or replace local text-layout overrides for one table cell.
    pub fn set_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        layout: crate::table_cell_layout::TableCellLayout,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_layout::set_cell_layout(&mut staged, table_id, row, column, layout)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_layout(table_id, row, column)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers table-cell layout failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local text-layout overrides and restore inherited cell values.
    pub fn reset_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_layout::reset_cell_layout(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective fill for one zero-based table cell.
    pub fn table_cell_fill(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::shapes::ShapeFill> {
        cell_fill::cell_fill(&self.package, table_id, row, column)
    }

    /// Create or replace a local table-cell fill transactionally.
    pub fn set_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        fill: &crate::shapes::ShapeFill,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_fill::set_cell_fill(&mut staged, table_id, row, column, fill)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if &verified.table_cell_fill(table_id, row, column)? != fill {
            return Err(Error::InvalidFormat(
                "Numbers table-cell fill failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct fill override and restore the inherited table style.
    pub fn reset_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_fill::reset_cell_fill(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective explicit borders for one zero-based table cell.
    pub fn table_cell_borders(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::table_cell_border::TableCellBorders> {
        stroke_layers::cell_borders(&self.package, table_id, row, column)
    }

    /// Create or replace one explicit table-cell border transactionally.
    pub fn set_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
        stroke: crate::shapes::ShapeStroke,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, Some(stroke))
    }

    /// Explicitly clear one table-cell border transactionally.
    pub fn clear_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, None)
    }

    fn update_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
        stroke: Option<crate::shapes::ShapeStroke>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        stroke_layers::set_cell_border(&mut staged, table_id, row, column, side, stroke)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .table_cell_borders(table_id, row, column)?
            .get(side)
            != stroke
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell border failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// List every native merged-cell rectangle in one attached table.
    pub fn table_cell_merges(&self, table_id: u64) -> Result<Vec<IWorkTableCellRegion>> {
        cell_merge::regions_in_package(&self.package, table_id)
    }

    /// Merge one non-overlapping rectangular cell region transactionally.
    pub fn merge_cells(&mut self, table_id: u64, region: IWorkTableCellRegion) -> Result<()> {
        let mut staged = self.package.clone();
        cell_merge::merge_in_package(&mut staged, table_id, region)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if !verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell merge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove one exact merged-cell rectangle, returning whether it existed.
    pub fn unmerge_cells(&mut self, table_id: u64, region: IWorkTableCellRegion) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_merge::unmerge_in_package(&mut staged, table_id, region)?;
        if !changed {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell unmerge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read the comment attached to a writable BNC cell.
    pub fn cell_comment(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<NumbersCellCommentInfo>> {
        cell_comment_in_package(&self.package, table_id, row, column)
    }

    /// Create or replace a cell comment without changing the cell value or style.
    pub fn set_cell_comment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_cell_comment_in_package(&mut staged, table_id, row, column, text.into())?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Delete a cell comment without changing the cell value or style.
    pub fn clear_cell_comment(&mut self, table_id: u64, row: usize, column: usize) -> Result<()> {
        let mut staged = self.package.clone();
        clear_cell_comment_in_package(&mut staged, table_id, row, column)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Read the direct replies attached to a cell comment in stored order.
    pub fn cell_comment_replies(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<NumbersCellCommentReplyInfo>> {
        cell_comment_replies_in_package(&self.package, table_id, row, column)
    }

    /// Append a direct reply to an existing cell comment.
    pub fn add_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<u64> {
        let mut staged = self.package.clone();
        let reply_id =
            add_cell_comment_reply_in_package(&mut staged, table_id, row, column, text.into())?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Replace one direct reply and return its new copy-on-write object ID.
    pub fn set_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        let mut staged = self.package.clone();
        let reply_id = set_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
            text.into(),
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Remove one direct reply from an existing cell comment.
    pub fn remove_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        remove_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Set a cell to a formula expression.
    ///
    /// The expression is compiled to Numbers' native postfix AST and interned
    /// in the table's formula list. Local and cross-table cells, rectangles,
    /// and whole-row/column references are mirrored into CalculationEngine
    /// dependency records in lockstep with the formula table. Unsupported volatile, lazy,
    /// remote-data, and spill expressions fail before the package is changed.
    pub fn set_formula(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        expression: FormulaExpression,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        table_formula::set_attached_table_formula(
            &mut staged,
            table_id,
            row,
            column,
            expression,
            None,
        )?;
        self.package = staged;
        Ok(())
    }

    /// Set a formula together with the value displayed before the next recalculation.
    pub fn set_formula_with_cached_value(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        expression: FormulaExpression,
        cached_value: FormulaCachedValue,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        table_formula::set_attached_table_formula(
            &mut staged,
            table_id,
            row,
            column,
            expression,
            Some(cached_value),
        )?;
        self.package = staged;
        Ok(())
    }

    pub fn rename_sheet(&mut self, sheet_id: u64, name: &str) -> Result<()> {
        validate_name(name, "sheet")?;
        if !numbers_document(&self.package)?
            .sheets
            .iter()
            .any(|reference| reference.identifier == sheet_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not in the workbook"
            )));
        }
        let locations = object_locations(&self.package)?;
        let archive_name = locations
            .get(&sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?
            .to_owned();
        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            let (message_index, _) = decode_sheet(object)?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let data = if message_type == 3 {
                patch_nested_length_delimited_field(original, &[1, 1], true, Some(name.as_bytes()))?
            } else {
                patch_length_delimited_field(original, 1, true, Some(name.as_bytes()))?
            };
            let verified_name = if message_type == 3 {
                tn::FormBasedSheetArchive::decode(data.as_slice())?
                    .super_
                    .name
            } else {
                tn::SheetArchive::decode(data.as_slice())?.name
            };
            if verified_name != name {
                return Err(Error::InvalidFormat(
                    "Numbers sheet-name wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified
            .sheets()?
            .iter()
            .find(|sheet| sheet.object_id == sheet_id)
            .map(|sheet| sheet.name.as_str())
            != Some(name)
        {
            return Err(Error::InvalidFormat(
                "Numbers sheet rename failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    pub fn rename_table(&mut self, table_id: u64, name: &str) -> Result<()> {
        if !self
            .tables()?
            .iter()
            .any(|table| table.object_id == table_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers table object {table_id} is not attached to a workbook sheet"
            )));
        }
        let mut staged = self.package.clone();
        rename_attached_table_in_package(&mut staged, table_id, name)?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified
            .tables()?
            .iter()
            .find(|table| table.object_id == table_id)
            .map(|table| table.name.as_str())
            != Some(name)
        {
            return Err(Error::InvalidFormat(
                "Numbers table rename failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Resize a table while preserving existing cells and stable row/column UIDs.
    ///
    /// Growth creates blank trailing rows or columns. Shrinkage is accepted only
    /// when the removed trailing region contains no stored cells; this prevents
    /// silently orphaning strings, formulas, rich text, comments, or styles.
    pub fn resize_table(&mut self, table_id: u64, rows: usize, columns: usize) -> Result<()> {
        let mut staged = self.package.clone();
        resize_attached_table_in_package(&mut staged, table_id, rows, columns)?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let resized = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers resized table disappeared".to_owned()))?;
        if (resized.rows, resized.columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Numbers table resize failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Unlink and remove a table model from its owning sheet.
    ///
    /// Private storage, formula dependency owners, UUID registrations, and
    /// now-empty component members are removed. Shared storage and styles are
    /// retained. Deletion is rejected while another table has a formula edge
    /// targeting this table.
    pub fn remove_table(&mut self, table_id: u64) -> Result<NumbersTableInfo> {
        let table = self
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let owner = find_table_owner(&self.package, table_id)?;
        let locations = object_locations(&self.package)?;
        let descriptors = table_models(&self.package)?;
        let descriptor = descriptors
            .iter()
            .find(|descriptor| descriptor.object_id == table_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
            })?;
        let owned_graph = table_owned_graph(&self.package, &locations, &descriptor.model)?;
        let mut shared_owned_ids = HashSet::new();
        for other in descriptors
            .iter()
            .filter(|candidate| candidate.object_id != table_id)
        {
            shared_owned_ids
                .extend(table_owned_graph(&self.package, &locations, &other.model)?.into_keys());
        }
        let private_owned_ids = owned_graph
            .into_keys()
            .filter(|identifier| !shared_owned_ids.contains(identifier))
            .collect::<Vec<_>>();
        let sheet_archive = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
        let mut staged = self.package.clone();
        let mut removed_identifiers = remove_table_formula_graph(&mut staged, owner.table_info_id)?;
        staged.update_archive(sheet_archive, |archive| {
            let object = archive.object_mut(owner.sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let previous = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let current = previous
                .iter()
                .copied()
                .filter(|identifier| *identifier != owner.table_info_id)
                .collect::<Vec<_>>();
            if current.len() + 1 != previous.len() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {} does not reference table info {} exactly once",
                    owner.sheet_id, owner.table_info_id
                )));
            }
            replace_sheet_drawable_references(object, message_index, &previous, &current)?;
            object.archive_info.message_infos[message_index]
                .object_references
                .retain(|&identifier| identifier != owner.table_info_id);
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                field
                    .object_references
                    .retain(|&identifier| identifier != owner.table_info_id);
            }
            Ok(())
        })?;
        let info_archive = locations.get(&owner.table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {} is missing",
                owner.table_info_id
            ))
        })?;
        let model_archive = locations.get(&table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
        })?;
        let private_owned_locations = private_owned_ids
            .iter()
            .map(|identifier| {
                locations
                    .get(identifier)
                    .map(|entry| (entry.as_str(), *identifier))
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers table storage object {identifier} is missing"
                        ))
                    })
            })
            .collect::<Result<Vec<_>>>()?;
        let mut affected_components = HashMap::<String, u64>::new();
        for (entry, identifier) in std::iter::once((info_archive.as_str(), owner.table_info_id))
            .chain(std::iter::once((model_archive.as_str(), table_id)))
            .chain(private_owned_locations)
        {
            let Some(component) = component_identifier_for_entry(&staged, entry)? else {
                continue;
            };
            affected_components.insert(entry.to_owned(), component);
            remove_component_external_references_to_object(&mut staged, component, identifier)?;
            if component_uuid_identifiers(&staged, component)?
                .is_some_and(|identifiers| identifiers.contains(&identifier))
            {
                remove_component_object_uuids(&mut staged, component, &[identifier])?;
            }
        }
        let dedicated_component = format!("Index/Tables/Table-{}.iwa", owner.table_info_id);
        if info_archive == model_archive && info_archive == &dedicated_component {
            staged.remove_entry(info_archive).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table component {info_archive} is missing"))
            })?;
        } else {
            remove_object_or_empty_entry(&mut staged, &locations, owner.table_info_id)?;
            remove_object_or_empty_entry(&mut staged, &locations, table_id)?;
        }
        for identifier in &private_owned_ids {
            remove_object_or_empty_entry(&mut staged, &locations, *identifier)?;
        }
        for (entry, component) in affected_components {
            if !staged.contains_entry(&entry) {
                remove_component_registration(&mut staged, component)?;
            }
        }
        removed_identifiers.extend([owner.table_info_id, table_id]);
        removed_identifiers.extend(private_owned_ids);
        release_package_identifier_suffix(&mut staged, &removed_identifiers)?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified
            .tables()?
            .iter()
            .any(|candidate| candidate.object_id == table_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers table deletion failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(table)
    }

    /// Move a sheet to another zero-based workbook position.
    pub fn move_sheet(&mut self, from: usize, to: usize) -> Result<()> {
        let sheets = self.sheets()?;
        if from >= sheets.len() || to >= sheets.len() {
            return Err(Error::ParseError(format!(
                "Numbers sheet move {from} -> {to} is out of range for {} sheets",
                sheets.len()
            )));
        }
        if from == to {
            return Ok(());
        }
        let moved_id = sheets[from].object_id;
        let mut staged = self.package.clone();
        update_numbers_document(&mut staged, |document| {
            let reference = document.sheets.remove(from);
            document.sheets.insert(to, reference);
            Ok(())
        })?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheets()?.get(to).map(|sheet| sheet.object_id) != Some(moved_id) {
            return Err(Error::InvalidFormat(
                "Numbers sheet move failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Append an empty sheet to the workbook and return its allocated object ID.
    pub fn add_empty_sheet(&mut self, name: &str) -> Result<NumbersSheetInfo> {
        validate_name(name, "sheet")?;
        let locations = object_locations(&self.package)?;
        let identifier = locations
            .keys()
            .copied()
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let mut staged = self.package.clone();
        staged.update_archive("Index/Document.iwa", |archive| {
            archive.insert_object(crate::archive::ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 2,
                    data: tn::SheetArchive {
                        name: name.to_owned(),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })?;
        update_numbers_document(&mut staged, |document| {
            document.sheets.push(crate::protobuf::tsp::Reference {
                identifier,
                ..Default::default()
            });
            Ok(())
        })?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheets()?
            .into_iter()
            .find(|sheet| sheet.object_id == identifier)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers sheet creation failed validation".to_owned())
            })?;
        self.package = staged;
        Ok(created)
    }

    /// Add an independent empty native table to an existing sheet.
    ///
    /// An attached table supplies structural templates when one exists. If the
    /// workbook is table-less, the first table is built from its native theme
    /// preset instead. Cell stores, data lists, row/column UIDs, headers, stroke
    /// state, and the CalculationEngine owner are allocated independently;
    /// workbook styles are shared intentionally.
    #[allow(deprecated)]
    pub fn add_empty_table(
        &mut self,
        sheet_id: u64,
        name: &str,
        rows: usize,
        columns: usize,
    ) -> Result<NumbersTableInfo> {
        let sheets = self.sheets()?;
        if !sheets.iter().any(|sheet| sheet.object_id == sheet_id) {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not in the workbook"
            )));
        }

        let descriptors = table_models(&self.package)?;
        let mut staged = self.package.clone();
        let graph = if let Some(template) = descriptors.first() {
            let template_owner = find_table_owner(&self.package, template.object_id)?;
            table_create::create_empty_table_graph(
                &mut staged,
                template_owner.table_info_id,
                template.object_id,
                template_owner.sheet_id,
                sheet_id,
                name,
                rows,
                columns,
                (template_owner.sheet_id == sheet_id).then_some(EMPTY_TABLE_POSITION_OFFSET),
            )?
        } else {
            table_bootstrap::bootstrap_empty_table_graph(
                &mut staged,
                sheet_id,
                name,
                rows,
                columns,
            )?
        };
        let new_info_id = graph.info_object_id;
        let new_model_id = graph.model_object_id;
        let locations = object_locations(&staged)?;
        let sheet_archive_name = locations
            .get(&sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
        staged.update_archive(sheet_archive_name, |archive| {
            let object = archive.object_mut(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let existing_drawables = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let existing_drawable_set = existing_drawables.iter().copied().collect::<HashSet<_>>();
            let mut current_drawables = existing_drawables.clone();
            current_drawables.push(new_info_id);
            replace_sheet_drawable_references(
                object,
                message_index,
                &existing_drawables,
                &current_drawables,
            )?;
            let references =
                &mut object.archive_info.message_infos[message_index].object_references;
            if !references.contains(&new_info_id) {
                references.push(new_info_id);
            }
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|identifier| existing_drawable_set.contains(identifier))
                    && !field.object_references.contains(&new_info_id)
                {
                    field.object_references.push(new_info_id);
                }
            }
            Ok(())
        })?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == new_model_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table creation failed validation".to_owned())
            })?;
        if (created.rows, created.columns, created.name.as_str()) != (rows, columns, name) {
            return Err(Error::InvalidFormat(
                "Numbers table creation produced unexpected properties".to_owned(),
            ));
        }
        self.package = staged;
        Ok(created)
    }

    /// Duplicate a populated table on its owning sheet.
    ///
    /// Cell tiles, headers, data lists, UID maps, stroke state, formulas, and
    /// CalculationEngine dependency owners are cloned independently. Workbook
    /// styles and referenced rich-text/comment payloads retain their native
    /// copy-on-write sharing.
    #[allow(deprecated)]
    pub fn duplicate_table(&mut self, table_id: u64) -> Result<NumbersTableInfo> {
        let descriptors = table_models(&self.package)?;
        let source = descriptors
            .iter()
            .find(|descriptor| descriptor.object_id == table_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers table {table_id} not found")))?;
        let owner = find_table_owner(&self.package, table_id)?;
        let existing_names = descriptors
            .iter()
            .filter_map(|descriptor| {
                find_table_owner(&self.package, descriptor.object_id)
                    .ok()
                    .filter(|candidate| candidate.sheet_id == owner.sheet_id)
                    .map(|_| descriptor.model.table_name.as_str())
            })
            .collect::<HashSet<_>>();
        let name = duplicate_table_name(&source.model.table_name, &existing_names)?;
        let source_package = &self.package;
        let locations = object_locations(source_package)?;
        let sheet_archive_name = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
        let mut staged = source_package.clone();
        let cloned = duplicate_attached_table_graph_in_package(
            source_package,
            &mut staged,
            owner.table_info_id,
            table_id,
            &name,
            TABLE_DUPLICATE_OFFSET,
        )?;
        staged.update_archive(sheet_archive_name, |archive| {
            let object = archive.object_mut(owner.sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let previous = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let mut current = previous.clone();
            current.push(cloned.info_object_id);
            replace_sheet_drawable_references(object, message_index, &previous, &current)?;
            let info = &mut object.archive_info.message_infos[message_index];
            info.object_references.push(cloned.info_object_id);
            for field in &mut info.field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|identifier| previous.contains(identifier))
                {
                    field.object_references.push(cloned.info_object_id);
                }
            }
            Ok(())
        })?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == cloned.model_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table duplication failed validation".to_owned())
            })?;
        if (created.name.as_str(), created.rows, created.columns)
            != (
                name.as_str(),
                source.model.number_of_rows as usize,
                source.model.number_of_columns as usize,
            )
        {
            return Err(Error::InvalidFormat(
                "Numbers table duplicate has unexpected properties".to_owned(),
            ));
        }
        self.package = staged;
        Ok(created)
    }

    /// Remove a sheet from the workbook, retaining unreachable drawable data.
    ///
    /// The final sheet cannot be removed. Retaining detached drawable objects is
    /// deliberate: styles and calculation objects can be shared across sheets.
    pub fn remove_sheet(&mut self, sheet_id: u64) -> Result<NumbersSheetInfo> {
        let sheets = self.sheets()?;
        if sheets.len() <= 1 {
            return Err(Error::ParseError(
                "Cannot remove the final Numbers sheet".to_owned(),
            ));
        }
        let removed = sheets
            .iter()
            .find(|sheet| sheet.object_id == sheet_id)
            .cloned()
            .ok_or_else(|| Error::ParseError(format!("Numbers sheet {sheet_id} not found")))?;
        let locations = object_locations(&self.package)?;
        let mut staged = self.package.clone();
        update_numbers_document(&mut staged, |document| {
            let old_len = document.sheets.len();
            document
                .sheets
                .retain(|reference| reference.identifier != sheet_id);
            if document.sheets.len() + 1 != old_len {
                return Err(Error::InvalidFormat(format!(
                    "Numbers root does not reference sheet {sheet_id} exactly once"
                )));
            }
            Ok(())
        })?;
        remove_object_or_empty_entry(&mut staged, &locations, sheet_id)?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers sheet deletion failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(removed)
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    fn sheet_owned_drawable_ids(&self, sheet_id: u64) -> Result<HashSet<u64>> {
        if !self
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not reachable"
            )));
        }
        Ok(numbers_sheet_drawable_owners(&self.package)?
            .into_iter()
            .filter_map(|(drawable_id, owner_id)| (owner_id == sheet_id).then_some(drawable_id))
            .collect())
    }

    fn require_sheet_drawable(&self, sheet_id: u64, drawable_object_id: u64) -> Result<()> {
        if !self
            .sheet_owned_drawable_ids(sheet_id)?
            .contains(&drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "drawable object {drawable_object_id} is not owned by Numbers sheet {sheet_id}"
            )));
        }
        if !self
            .sheet_drawables(sheet_id)?
            .iter()
            .any(|drawable| drawable.object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers sheet drawable {drawable_object_id} has no supported direct drawable payload"
            )));
        }
        Ok(())
    }

    /// List metadata-backed media reachable from this spreadsheet package.
    pub fn media_assets(&self) -> Result<Vec<EmbeddedMediaAsset>> {
        reachable_embedded_assets(&self.package, [1])
    }

    /// List media reachable from one sheet and its drawable object graph.
    pub fn sheet_media_assets(&self, sheet_id: u64) -> Result<Vec<EmbeddedMediaAsset>> {
        if !self
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not reachable"
            )));
        }
        reachable_embedded_assets(&self.package, [sheet_id])
    }

    pub fn extract_media(&self, data_identifier: u64) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Numbers object graph"
            )));
        }
        IWorkMediaEditor::from_package(self.package.clone())?.extract(data_identifier)
    }

    /// Replace a referenced materialized asset without changing its data identifier.
    pub fn replace_media(&mut self, data_identifier: u64, replacement: &[u8]) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Numbers object graph"
            )));
        }
        let mut media = IWorkMediaEditor::from_package(self.package.clone())?;
        let old = media.replace(data_identifier, replacement)?;
        let staged = media.into_package();
        Self::from_package(staged.clone())?;
        self.package = staged;
        Ok(old)
    }

    pub fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

pub(crate) fn set_table_cell_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    model::set_attached_cell_in_package(package, table_id, row, column, value)?;
    formula_cache::refresh_formula_caches_after_cell_write(package, table_id, row, column)?;
    Ok(())
}

pub(crate) fn table_cell_borders_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<crate::table_cell_border::TableCellBorders> {
    stroke_layers::cell_borders(package, table_id, row, column)
}

pub(crate) fn table_cell_fill_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<crate::shapes::ShapeFill> {
    cell_fill::cell_fill(package, table_id, row, column)
}

pub(crate) fn table_cell_layout_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<crate::table_cell_layout::TableCellLayout> {
    cell_layout::cell_layout(package, table_id, row, column)
}

pub(crate) fn table_cell_number_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<crate::table_cell_number_format::TableCellNumberFormat>> {
    cell_data_format::cell_number_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: crate::table_cell_number_format::TableCellNumberFormat,
) -> Result<()> {
    cell_data_format::set_cell_number_format(package, table_id, row, column, format)
}

pub(crate) fn reset_table_cell_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_number_format(package, table_id, row, column)
}

pub(crate) fn table_cell_currency_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<crate::table_cell_data_format::TableCellCurrencyFormat>> {
    cell_data_format::cell_currency_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_currency_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_currency_format(package, table_id, row, column)
}

pub(crate) fn table_cell_data_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<crate::table_cell_data_format::TableCellDataFormat> {
    cell_data_format::cell_data_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_data_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: crate::table_cell_data_format::TableCellDataFormat,
) -> Result<()> {
    cell_data_format::set_cell_data_format(package, table_id, row, column, format)
}

pub(crate) fn table_cell_percentage_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<crate::table_cell_data_format::TableCellPercentageFormat>> {
    cell_data_format::cell_percentage_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_percentage_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_percentage_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_layout_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    layout: crate::table_cell_layout::TableCellLayout,
) -> Result<()> {
    cell_layout::set_cell_layout(package, table_id, row, column, layout)
}

pub(crate) fn reset_table_cell_layout_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_layout::reset_cell_layout(package, table_id, row, column)
}

pub(crate) fn set_table_cell_fill_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    fill: &crate::shapes::ShapeFill,
) -> Result<()> {
    cell_fill::set_cell_fill(package, table_id, row, column, fill)
}

pub(crate) fn reset_table_cell_fill_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_fill::reset_cell_fill(package, table_id, row, column)
}

pub(crate) fn set_table_cell_border_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    side: crate::table_cell_border::TableCellBorderSide,
    stroke: Option<crate::shapes::ShapeStroke>,
) -> Result<()> {
    stroke_layers::set_cell_border(package, table_id, row, column, side, stroke)
}

pub(crate) fn table_cell_merges_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Vec<IWorkTableCellRegion>> {
    cell_merge::regions_in_package(package, table_id)
}

pub(crate) fn merge_table_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: IWorkTableCellRegion,
) -> Result<()> {
    cell_merge::merge_in_package(package, table_id, region)
}

pub(crate) fn unmerge_table_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: IWorkTableCellRegion,
) -> Result<bool> {
    cell_merge::unmerge_in_package(package, table_id, region)
}

pub(crate) fn table_cell_comment_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumbersCellCommentInfo>> {
    model::attached_cell_comment_in_package(package, table_id, row, column)
}

pub(crate) fn set_table_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<()> {
    model::set_attached_cell_comment_in_package(package, table_id, row, column, text)
}

pub(crate) fn clear_table_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    model::clear_attached_cell_comment_in_package(package, table_id, row, column)
}

pub(crate) fn table_cell_comment_replies_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<NumbersCellCommentReplyInfo>> {
    model::attached_cell_comment_replies_in_package(package, table_id, row, column)
}

pub(crate) fn add_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<u64> {
    model::add_attached_cell_comment_reply_in_package(package, table_id, row, column, text)
}

pub(crate) fn set_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    model::set_attached_cell_comment_reply_in_package(
        package,
        table_id,
        row,
        column,
        reply_storage_object_id,
        text,
    )
}

pub(crate) fn remove_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<()> {
    model::remove_attached_cell_comment_reply_in_package(
        package,
        table_id,
        row,
        column,
        reply_storage_object_id,
    )
}

pub(crate) fn set_table_formula_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expression: FormulaExpression,
    cached_value: FormulaCachedValue,
) -> Result<()> {
    table_formula::set_attached_table_formula(
        package,
        table_id,
        row,
        column,
        expression,
        Some(cached_value),
    )
}

pub(crate) fn rename_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    name: &str,
) -> Result<()> {
    model::rename_attached_table_in_package(package, table_id, name)
}

pub(crate) fn resize_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    rows: usize,
    columns: usize,
) -> Result<()> {
    model::resize_attached_table_in_package(package, table_id, rows, columns)
}

pub(crate) fn table_dimensions_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<(usize, usize)> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    Ok((
        descriptor.model.number_of_rows as usize,
        descriptor.model.number_of_columns as usize,
    ))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TableTopologyMutation {
    InsertRow(TableRowInsertion),
    InsertColumn(TableColumnInsertion),
    RemoveRow(TableRowDeletion),
    RemoveColumn(TableColumnDeletion),
}

impl TableTopologyMutation {
    pub(crate) fn apply(self, package: &mut IWorkPackage, table_id: u64) -> Result<(usize, usize)> {
        let dimensions = table_dimensions_in_package(package, table_id)?;
        match self {
            Self::InsertRow(row) => Ok((
                insert_table_row_in_package(package, table_id, row)?,
                dimensions.1,
            )),
            Self::InsertColumn(column) => Ok((
                dimensions.0,
                insert_table_column_in_package(package, table_id, column)?,
            )),
            Self::RemoveRow(row) => remove_table_row_in_package(package, table_id, row),
            Self::RemoveColumn(column) => remove_table_column_in_package(package, table_id, column),
        }
    }
}

pub(crate) fn insert_table_row_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    insertion: TableRowInsertion,
) -> Result<usize> {
    row_insert::insert_attached_table_row(package, table_id, insertion)
}

pub(crate) fn insert_table_column_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    insertion: TableColumnInsertion,
) -> Result<usize> {
    column_insert::insert_attached_table_column(package, table_id, insertion)
}

pub(crate) fn remove_table_row_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: TableRowDeletion,
) -> Result<(usize, usize)> {
    table_delete::remove_attached_table_row(package, table_id, deletion)
}

pub(crate) fn remove_table_column_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: TableColumnDeletion,
) -> Result<(usize, usize)> {
    table_delete::remove_attached_table_column(package, table_id, deletion)
}

pub(crate) fn set_table_dimension_size_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    dimension: NumbersTableDimension,
    size: NumbersTableDimensionSize,
) -> Result<()> {
    table_dimension::set_attached_table_dimension_size(package, table_id, dimension, size)
}

pub(crate) fn table_dimension_size_in_package(
    package: &IWorkPackage,
    table_id: u64,
    dimension: NumbersTableDimension,
) -> Result<NumbersTableDimensionSize> {
    table_dimension::read_attached_table_dimension_size(package, table_id, dimension)
}

pub(crate) fn table_size_points_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<(f32, f32)> {
    table_dimension::attached_table_size_points(package, table_id)
}

pub(crate) fn table_header_settings_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<NumbersTableHeaderSettings> {
    table_headers::read_attached_table_header_settings(package, table_id)
}

pub(crate) fn set_table_header_settings_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    settings: NumbersTableHeaderSettings,
) -> Result<()> {
    table_headers::set_attached_table_header_settings(package, table_id, settings)
}

pub(crate) fn table_owned_object_ids_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Vec<u64>> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    let locations = object_locations(package)?;
    Ok(table_owned_graph(package, &locations, &descriptor.model)?
        .into_keys()
        .collect())
}

pub(crate) fn remove_table_formula_graph_in_package(
    package: &mut IWorkPackage,
    table_context_ids: &[u64],
) -> Result<Vec<u64>> {
    formula_clone::remove_table_formula_graph_for_contexts(package, table_context_ids)
}

pub(crate) fn create_empty_table_graph_in_package(
    package: &mut IWorkPackage,
    template_info_id: u64,
    template_model_id: u64,
    parent_id: u64,
    name: &str,
    rows: usize,
    columns: usize,
) -> Result<(u64, u64)> {
    let graph = table_create::create_empty_table_graph(
        package,
        template_info_id,
        template_model_id,
        parent_id,
        parent_id,
        name,
        rows,
        columns,
        None,
    )?;
    Ok((graph.info_object_id, graph.model_object_id))
}

mod cell_data_format;
mod cell_fill;
mod cell_layout;
mod cell_merge;
mod cell_style;
mod column_insert;
mod date_time_fields;
mod drawable_order;
mod formula_cache;
mod formula_clone;
mod formula_dependency_shift;
mod model;
mod row_insert;
mod sheet_audio;
mod sheet_charts;
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

pub use crate::charts::ChartSeriesDirection;
pub use cell_merge::IWorkTableCellRegion;
use model::*;
pub use sheet_audio::{NumbersSheetAudioInfo, NumbersSheetAudioOptions, RemovedNumbersSheetAudio};
pub use sheet_charts::{NumbersSheetChartInfo, RemovedNumbersSheetChart};
pub use sheet_images::{NumbersSheetImageInfo, NumbersSheetImageOptions, RemovedNumbersSheetImage};
pub use sheet_movies::{NumbersSheetMovieInfo, NumbersSheetMovieOptions, RemovedNumbersSheetMovie};
pub use sheet_shapes::{NumbersSheetShapeInfo, NumbersSheetShapeKind, RemovedNumbersSheetShape};
use storage::*;
pub use table_axis_deletion::{TableColumnDeletion, TableRowDeletion};
pub use table_axis_insertion::{TableColumnInsertion, TableRowInsertion};
pub(crate) use table_cells::TableCellBatch;
pub use table_dimension::{NumbersTableDimension, NumbersTableDimensionSize, NumbersTablePoints};
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
pub use table_title::NumbersTableTitleSettings;
pub(crate) use table_title::{
    set_table_title_settings_in_package, table_title_settings_in_package,
};
#[cfg(test)]
mod tests;
