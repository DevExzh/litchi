//! Transactional semantic editing for Numbers spreadsheets.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use super::bnc::{BncCell, StoredValue};
use super::cell::CellValue;
use super::formula::{
    ExternalFormulaTable, ExternalPivotCategory, FormulaExpression, FormulaPivotCategoryReference,
    FormulaUuid, PivotFormulaKey,
};
use super::table::{NumbersCellComment, NumbersCommentUuid};
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::{
    DrawableCommentInfo, DrawableCommentReplyInfo, IWorkDrawableCommentEditor, IWorkDrawableInfo,
    advance_save_tokens_for_entries, clone_comment_storage_exact, current_apple_reference_date,
    fresh_comment_storage_uuid, insert_comment_storage, preferred_annotation_author,
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
use crate::text::{
    IWorkTextEditor, ParagraphDropCap, ParagraphDropCapPlacement, ParagraphIndents,
    ParagraphLineSpacing, ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement,
    ParagraphSpacing, ParagraphStart, ParagraphTabStops, TextAlignment, TextBackground,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextColumns, TextDecorations,
    TextFont, TextLigatures, TextOutline, TextScript, TextShadow, TextStorageInfo, TextStyle,
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
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const DOCUMENT_COMPONENT_IDENTIFIER: u64 = 1;
const TEXT_BOX_DUPLICATE_OFFSET: f32 = 10.0;
const TABLE_DUPLICATE_OFFSET: f32 = 10.0;

/// Stable identity and dimensions of a Numbers table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersTableInfo {
    pub object_id: u64,
    pub name: String,
    pub rows: usize,
    pub columns: usize,
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
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersCellCommentInfo {
    pub table_id: u64,
    pub row: usize,
    pub column: usize,
    pub list_identifier: u32,
    pub storage_object_id: u64,
    pub comment: NumbersCellComment,
}

/// A resolved direct reply in a Numbers cell-comment thread.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersCellCommentReplyInfo {
    pub table_id: u64,
    pub row: usize,
    pub column: usize,
    pub root_storage_object_id: u64,
    pub storage_object_id: u64,
    pub comment: NumbersCellComment,
}

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
        let mut tables = table_models(&self.package)?
            .into_iter()
            .map(|descriptor| NumbersTableInfo {
                object_id: descriptor.object_id,
                name: descriptor.model.table_name,
                rows: descriptor.model.number_of_rows as usize,
                columns: descriptor.model.number_of_columns as usize,
            })
            .collect::<Vec<_>>();
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
    pub fn set_cell(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        value: CellValue,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_cell_in_package(&mut staged, table_id, row, column, value)?;
        // Exercise every serialization boundary before committing the edit.
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    pub fn clear_cell(&mut self, table_id: u64, row: usize, column: usize) -> Result<()> {
        self.set_cell(table_id, row, column, CellValue::Empty)
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
        let descriptors = table_models(&staged)?;
        let descriptor = descriptors
            .iter()
            .find(|table| table.object_id == table_id)
            .cloned()
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let external_tables = formula_external_tables(&staged, &descriptors)?;
        let pivot_categories = formula_pivot_categories(&staged)?;
        let compiled = expression.compile(
            row,
            column,
            descriptor.model.number_of_rows as usize,
            descriptor.model.number_of_columns as usize,
            &external_tables,
            &pivot_categories,
        )?;
        let formula = compiled.archive;

        let locations = object_locations(&staged)?;
        let (old_formula, old_formula_error) = {
            let location = locate_cell(&staged, table_id, row, column)?;
            let cell = read_tile_cell(
                &staged,
                &location.tile_archive,
                location.tile_id,
                location.tile_row,
                column,
            )?
            .as_deref()
            .map(BncCell::parse)
            .transpose()?;
            let formula = cell.as_ref().and_then(|cell| match cell.stored_value() {
                StoredValue::Formula(identifier) => Some(identifier),
                _ => None,
            });
            let error = cell.as_ref().and_then(BncCell::formula_error_identifier);
            (formula, error)
        };
        if old_formula.is_some()
            && let Some(identifier) = old_formula_error
        {
            decrement_formula_error_table(&mut staged, &locations, &descriptor.model, identifier)?;
        }
        if let Some(identifier) = old_formula {
            // Formula-to-formula replacement keeps the app's cached result and
            // only swaps the formula reference. Numbers can then display the
            // prior value until its calculation engine refreshes the cell.
            decrement_formula_table(
                &mut staged,
                &locations,
                descriptor.model.base_data_store.formula_table.identifier,
                identifier,
            )?;
            update_formula_dependencies(
                &mut staged,
                descriptor.table_info_id,
                row,
                column,
                false,
                &[],
                &[],
            )?;
        } else {
            // Reuse the primitive mutation path to validate the target and
            // release any string reference owned by the previous cell.
            set_cell_in_package(&mut staged, table_id, row, column, CellValue::Number(0.0))?;
        }

        let formula_id = insert_formula_table(
            &mut staged,
            &locations,
            descriptor.model.base_data_store.formula_table.identifier,
            formula.clone(),
        )?;
        set_encoded_cell_value(
            &mut staged,
            table_id,
            row,
            column,
            EncodedValue::Formula(formula_id),
        )?;
        update_formula_dependencies(
            &mut staged,
            descriptor.table_info_id,
            row,
            column,
            true,
            &compiled.local_precedents,
            &compiled.external_precedents,
        )?;

        // Reparse the complete ZIP and verify both sides of the reference
        // before committing the staged package.
        let verified = IWorkPackage::from_bytes(&staged.to_bytes()?)?;
        verify_formula_link(&verified, table_id, row, column, formula_id, &formula)?;
        verify_formula_dependency(
            &verified,
            descriptor.table_info_id,
            row,
            column,
            &compiled.local_precedents,
            &compiled.external_precedents,
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
        validate_name(name, "table")?;
        if !self
            .tables()?
            .iter()
            .any(|table| table.object_id == table_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers table object {table_id} is not attached to a workbook sheet"
            )));
        }
        let locations = object_locations(&self.package)?;
        let archive_name = locations
            .get(&table_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers table {table_id} is missing")))?
            .to_owned();
        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(table_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table {table_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    (message.type_ == 6000 || message.type_ == 6001)
                        && TableModelArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {table_id} has no Numbers table model payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let data = patch_length_delimited_field(original, 8, true, Some(name.as_bytes()))?;
            let verified = TableModelArchive::decode(data.as_slice())?;
            if verified.table_name != name {
                return Err(Error::InvalidFormat(
                    "Numbers table-name wire patch failed validation".to_owned(),
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
        let (rows_u32, columns_u32) = validate_table_dimensions(rows, columns)?;
        let descriptor = table_models(&self.package)?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let old_rows = descriptor.model.number_of_rows as usize;
        let old_columns = descriptor.model.number_of_columns as usize;
        if (rows, columns) == (old_rows, old_columns) {
            return Ok(());
        }

        let locations = object_locations(&self.package)?;
        let mut staged = self.package.clone();
        validate_and_trim_tiles(&mut staged, &locations, &descriptor.model, rows, columns)?;
        resize_header_buckets(
            &mut staged,
            &locations,
            &descriptor.model,
            rows_u32,
            columns_u32,
        )?;
        if let Some(reference) = &descriptor.model.base_column_row_uids {
            resize_uid_map(
                &mut staged,
                &locations,
                reference.identifier,
                old_rows,
                rows,
                old_columns,
                columns,
            )?;
        }
        if let Some(reference) = &descriptor.model.stroke_sidecar {
            resize_stroke_sidecar(
                &mut staged,
                &locations,
                reference.identifier,
                rows_u32,
                columns_u32,
            )?;
        }
        let table_archive = locations.get(&table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
        })?;
        staged.update_archive(table_archive, |archive| {
            let object = archive.object_mut(table_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    (message.type_ == 6000 || message.type_ == 6001)
                        && TableModelArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {table_id} has no Numbers table model payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let mut data = patch_varint_field(original, 6, true, Some(u64::from(rows_u32)))?;
            data = patch_varint_field(&data, 7, true, Some(u64::from(columns_u32)))?;
            let verified = TableModelArchive::decode(data.as_slice())?;
            if (verified.number_of_rows, verified.number_of_columns) != (rows_u32, columns_u32) {
                return Err(Error::InvalidFormat(
                    "Numbers table-dimension wire patch failed validation".to_owned(),
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

    /// Add an empty table to an existing sheet using another table as a style template.
    ///
    /// Cell stores, data lists, row/column UIDs, headers, and stroke state are
    /// allocated independently in their corresponding native components. A
    /// fresh CalculationEngine owner makes the table immediately formula-ready.
    /// Workbook styles are shared intentionally.
    #[allow(deprecated)]
    pub fn add_empty_table(
        &mut self,
        sheet_id: u64,
        name: &str,
        rows: usize,
        columns: usize,
    ) -> Result<NumbersTableInfo> {
        validate_name(name, "table")?;
        let (rows_u32, columns_u32) = validate_table_dimensions(rows, columns)?;
        let sheets = self.sheets()?;
        if !sheets.iter().any(|sheet| sheet.object_id == sheet_id) {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not in the workbook"
            )));
        }

        let descriptors = table_models(&self.package)?;
        let template = descriptors.first().ok_or_else(|| {
            Error::ParseError(
                "Adding a Numbers table requires an existing table style template".to_owned(),
            )
        })?;
        let template_owner = find_table_owner(&self.package, template.object_id)?;
        let locations = object_locations(&self.package)?;
        let template_info_archive = locations
            .get(&template_owner.table_info_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers table info is missing".to_owned()))?;
        let template_info_component = self.package.archive(template_info_archive)?;
        let template_info_object = template_info_component
            .object(template_owner.table_info_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers table info is missing".to_owned()))?;
        let (info_message_index, mut table_info) = decode_table_info(template_info_object)?;
        let template_model_archive = locations
            .get(&template.object_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers table model is missing".to_owned()))?;
        let template_model_component = self.package.archive(template_model_archive)?;
        let template_model_object = template_model_component
            .object(template.object_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers table model is missing".to_owned()))?;
        let model_message_index = find_table_model_message(template_model_object)?;

        let mut next_identifier = next_object_identifier(&self.package)?;
        let new_info_id = take_identifier(&mut next_identifier)?;
        let new_model_id = take_identifier(&mut next_identifier)?;
        let owned_kinds = table_owned_objects(&template.model)?;
        let mut remap = HashMap::with_capacity(owned_kinds.len() + 2);
        remap.insert(template_owner.table_info_id, new_info_id);
        remap.insert(template.object_id, new_model_id);
        for &identifier in owned_kinds.keys() {
            remap.insert(identifier, take_identifier(&mut next_identifier)?);
        }

        let existing_table_ids = descriptors
            .iter()
            .map(|descriptor| descriptor.model.table_id.as_str())
            .collect::<HashSet<_>>();
        let table_uuid = allocate_table_uuid(new_model_id, &existing_table_ids);
        let mut model = template.model.clone();
        prepare_empty_table_model(&mut model, &remap, &table_uuid, name, rows_u32, columns_u32)?;

        table_info.super_.parent = Some(crate::protobuf::tsp::Reference {
            identifier: sheet_id,
            ..Default::default()
        });
        if template_owner.sheet_id == sheet_id
            && let Some(position) = table_info
                .super_
                .geometry
                .as_mut()
                .and_then(|geometry| geometry.position.as_mut())
        {
            position.x += 40.0;
            position.y += 40.0;
        }
        table_info.super_.comment = None;
        table_info.super_.pencil_annotations.clear();
        table_info.super_.title = None;
        table_info.super_.caption = None;
        table_info.table_model = crate::protobuf::tsp::Reference {
            identifier: new_model_id,
            ..Default::default()
        };
        table_info.editing_state = None;
        table_info.summary_model = None;
        table_info.category_order = None;
        table_info.view_column_row_uids = None;
        table_info.pivot_data_model = None;
        table_info.pivot_order = None;

        let mut info_remap = remap.clone();
        info_remap.insert(template_owner.sheet_id, sheet_id);
        let mut objects = Vec::with_capacity(owned_kinds.len() + 2);
        objects.push((
            template_info_archive.clone(),
            clone_single_payload_object(
                template_info_object,
                new_info_id,
                info_message_index,
                table_info.encode_to_vec(),
                vec![sheet_id, new_model_id],
                &info_remap,
                false,
            )?,
        ));
        objects.push((
            template_model_archive.clone(),
            clone_single_payload_object(
                template_model_object,
                new_model_id,
                model_message_index,
                model.encode_to_vec(),
                table_model_references(&model),
                &remap,
                false,
            )?,
        ));
        for (&source_id, &kind) in &owned_kinds {
            let archive_name = locations.get(&source_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table storage object {source_id} is missing"
                ))
            })?;
            let source_archive = self.package.archive(archive_name)?;
            let source = source_archive.object(source_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table storage object {source_id} is missing"
                ))
            })?;
            objects.push((
                archive_name.clone(),
                clone_empty_table_storage(
                    source,
                    remap[&source_id],
                    kind,
                    rows_u32,
                    columns_u32,
                    new_model_id,
                )?,
            ));
        }

        let mut staged = self.package.clone();
        for (archive_name, object) in objects {
            staged.update_archive(&archive_name, |archive| archive.insert_object(object))?;
        }
        register_cloned_numbers_objects(&mut staged, &self.package, &locations, &remap)?;
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
        register_numbers_component_reference(
            &mut staged,
            sheet_archive_name,
            template_info_archive,
            new_info_id,
        )?;
        let table_last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
            Error::InvalidFormat("Numbers table creation allocated no identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, table_last_identifier)?;
        if create_empty_table_formula_graph(&mut staged, new_info_id, &table_uuid)?.is_some() {
            register_numbers_component_reference(
                &mut staged,
                "Index/CalculationEngine.iwa",
                template_info_archive,
                new_info_id,
            )?;
        }

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
        let locations = object_locations(&self.package)?;
        let info_archive_name = locations.get(&owner.table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {} is missing",
                owner.table_info_id
            ))
        })?;
        let info_archive = self.package.archive(info_archive_name)?;
        let info_object = info_archive.object(owner.table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {} is missing",
                owner.table_info_id
            ))
        })?;
        let (info_message_index, source_info) = decode_table_info(info_object)?;
        let model_archive_name = locations.get(&table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
        })?;
        let model_archive = self.package.archive(model_archive_name)?;
        let model_object = model_archive.object(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
        })?;
        let model_message_index = find_table_model_message(model_object)?;

        let graph = table_owned_graph(&self.package, &locations, &source.model)?;
        let mut next_identifier = next_object_identifier(&self.package)?;
        let new_info_id = take_identifier(&mut next_identifier)?;
        let new_model_id = take_identifier(&mut next_identifier)?;
        let mut remap = HashMap::with_capacity(graph.len() + 2);
        remap.insert(owner.table_info_id, new_info_id);
        remap.insert(table_id, new_model_id);
        for &identifier in graph.keys() {
            remap.insert(identifier, take_identifier(&mut next_identifier)?);
        }

        let existing_table_ids = descriptors
            .iter()
            .map(|descriptor| descriptor.model.table_id.as_str())
            .collect::<HashSet<_>>();
        let table_uuid = allocate_table_uuid(new_model_id, &existing_table_ids);
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

        let model_data = duplicate_table_model_wire(
            model_object.messages[model_message_index].data.as_slice(),
            &source.model,
            &remap,
            &table_uuid,
            &name,
        )?;
        let mut objects = Vec::with_capacity(graph.len() + 2);
        objects.push((
            model_archive_name.clone(),
            clone_numbers_object_metadata(
                model_object,
                new_model_id,
                vec![RawMessage {
                    type_: model_object.messages[model_message_index].type_,
                    data: model_data,
                }],
                &remap,
            )?,
        ));

        let info_data = duplicate_table_info_wire(
            info_object.messages[info_message_index].data.as_slice(),
            &source_info,
            &remap,
            TABLE_DUPLICATE_OFFSET,
        )?;
        objects.push((
            info_archive_name.clone(),
            clone_numbers_object_metadata(
                info_object,
                new_info_id,
                vec![RawMessage {
                    type_: info_object.messages[info_message_index].type_,
                    data: info_data,
                }],
                &remap,
            )?,
        ));

        for &source_id in graph.keys() {
            let archive_name = locations.get(&source_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table storage object {source_id} is missing"
                ))
            })?;
            let archive = self.package.archive(archive_name)?;
            let source_object = archive.object(source_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table storage object {source_id} is missing"
                ))
            })?;
            let mut cloned = clone_table_storage_object(source_object, &remap)?;
            remap_cloned_formula_storage(&mut cloned, &source.model.table_id, &table_uuid)?;
            objects.push((archive_name.clone(), cloned));
        }

        let mut staged = self.package.clone();
        for (archive_name, object) in objects {
            staged.update_archive(&archive_name, |archive| archive.insert_object(object))?;
        }
        register_cloned_numbers_objects(&mut staged, &self.package, &locations, &remap)?;
        if let Some((source_owner_uuid, new_owner_uuid)) = formula_graph_owner_uuids(
            &staged,
            owner.table_info_id,
            &source.model.table_id,
            &table_uuid,
        )? {
            for &source_id in graph.keys() {
                let archive_name = locations.get(&source_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table storage object {source_id} is missing"
                    ))
                })?;
                let cloned_id = remap[&source_id];
                staged.update_archive(archive_name, |archive| {
                    let object = archive.object_mut(cloned_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers cloned table storage object {cloned_id} is missing"
                        ))
                    })?;
                    remap_cloned_formula_owner_storage(object, &source_owner_uuid, &new_owner_uuid)
                })?;
            }
        }
        let sheet_archive_name = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
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
            current.push(new_info_id);
            replace_sheet_drawable_references(object, message_index, &previous, &current)?;
            let info = &mut object.archive_info.message_infos[message_index];
            info.object_references.push(new_info_id);
            for field in &mut info.field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|identifier| previous.contains(identifier))
                {
                    field.object_references.push(new_info_id);
                }
            }
            Ok(())
        })?;
        register_numbers_component_reference(
            &mut staged,
            sheet_archive_name,
            info_archive_name,
            new_info_id,
        )?;
        let table_last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
            Error::InvalidFormat("Numbers table clone allocated no identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, table_last_identifier)?;
        clone_table_formula_graph(
            &mut staged,
            owner.table_info_id,
            new_info_id,
            &source.model.table_id,
            &table_uuid,
        )?;
        register_numbers_component_reference(
            &mut staged,
            "Index/CalculationEngine.iwa",
            info_archive_name,
            new_info_id,
        )?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == new_model_id)
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

mod column_insert;
mod formula_clone;
mod formula_dependency_shift;
mod model;
mod row_insert;
mod sheet_audio;
mod sheet_duplicate;
mod sheet_images;
mod sheet_movies;
mod sheet_shapes;
mod storage;
mod table_delete;
mod table_dimension;
mod table_duplicate;
mod table_headers;
mod table_move;
mod table_title;
mod table_topology;
mod text_box_create;
mod text_box_duplicate;

use model::*;
pub use sheet_audio::{NumbersSheetAudioInfo, NumbersSheetAudioOptions, RemovedNumbersSheetAudio};
pub use sheet_images::{NumbersSheetImageInfo, RemovedNumbersSheetImage};
pub use sheet_movies::{NumbersSheetMovieInfo, NumbersSheetMovieOptions, RemovedNumbersSheetMovie};
pub use sheet_shapes::{NumbersSheetShapeInfo, NumbersSheetShapeKind, RemovedNumbersSheetShape};
use storage::*;
pub use table_dimension::{NumbersTableDimension, NumbersTableDimensionSize, NumbersTablePoints};
use table_duplicate::*;
pub use table_headers::{NumbersTableHeaderCount, NumbersTableHeaderSettings};
pub use table_title::NumbersTableTitleSettings;
#[cfg(test)]
mod tests;
