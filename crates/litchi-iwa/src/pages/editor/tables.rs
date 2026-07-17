//! Native table discovery and cell editing for Pages body attachments.

use std::collections::{HashMap, HashSet};

use prost::Message;

use super::*;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::protobuf::tst::{TableInfoArchive, TableModelArchive};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;

/// Strongly typed cell value shared by Pages and Numbers table storage.
pub type PagesCellValue = crate::numbers::CellValue;

/// Stable identity and dimensions of one native table attached to the Pages body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesTableInfo {
    /// Object identifier of the body-owned table drawable.
    pub drawable_object_id: u64,
    /// Object identifier accepted by cell-editing APIs.
    pub model_object_id: u64,
    /// UTF-16 body position of the object-replacement character.
    pub anchor_character_index: usize,
    /// Table name stored in the native table model.
    pub name: String,
    /// Number of addressable rows.
    pub rows: usize,
    /// Number of addressable columns.
    pub columns: usize,
}

/// Materialized values from one Pages table.
#[derive(Debug, Clone)]
pub struct PagesTable {
    /// Stable identity and dimensions of this table.
    pub info: PagesTableInfo,
    /// Materialized non-empty cells indexed by `(row, column)`.
    pub cells: HashMap<(usize, usize), PagesCellValue>,
}

impl PagesTable {
    /// Borrow a materialized cell value, or return `None` for an empty cell.
    pub fn get_cell(&self, row: usize, column: usize) -> Option<&PagesCellValue> {
        self.cells.get(&(row, column))
    }
}

#[derive(Debug, Clone)]
struct PagesTableGraph {
    info: PagesTableInfo,
}

impl PagesEditor {
    /// List native tables anchored in the main body in document order.
    pub fn tables(&self) -> Result<Vec<PagesTableInfo>> {
        Ok(body_table_graphs(self)?
            .into_iter()
            .map(|graph| graph.info)
            .collect())
    }

    /// Read all materialized cell values from one reachable body table.
    pub fn table(&self, model_object_id: u64) -> Result<PagesTable> {
        let info = self
            .tables()?
            .into_iter()
            .find(|table| table.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table model {model_object_id} is not attached to the body"
                ))
            })?;
        let bytes = self.package().to_bytes()?;
        let bundle = Bundle::from_bytes(&bytes)?;
        let index = ObjectIndex::from_bundle(&bundle)?;
        let object = index
            .resolve_object(&bundle, model_object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Pages table model {model_object_id} is missing"))
            })?;
        let table = TableDataExtractor::new(&bundle, &index)
            .extract_table_from_object(&object)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages object {model_object_id} has no native table model"
                ))
            })?;
        Ok(PagesTable {
            info,
            cells: table.cells,
        })
    }

    /// Set or clear one cell in a reachable body table transactionally.
    pub fn set_table_cell(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        value: PagesCellValue,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            value,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Clear one cell in a reachable body table.
    pub fn clear_table_cell(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_table_cell(model_object_id, row, column, PagesCellValue::Empty)
    }

    /// Rename a reachable body table transactionally.
    pub fn rename_table(&mut self, model_object_id: u64, name: &str) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::rename_table_in_package(&mut staged, model_object_id, name)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let renamed = verified.require_body_table(model_object_id)?;
        if renamed.info.name != name {
            return Err(Error::InvalidFormat(
                "Pages table rename failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Resize a reachable body table while preserving existing cells and UIDs.
    ///
    /// Growth creates blank trailing rows or columns. Shrinkage is rejected if
    /// any removed row or column contains stored cell data.
    pub fn resize_table(
        &mut self,
        model_object_id: u64,
        rows: usize,
        columns: usize,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::resize_table_in_package(
            &mut staged,
            model_object_id,
            rows,
            columns,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let resized = verified.require_body_table(model_object_id)?;
        if (resized.info.rows, resized.info.columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Pages table resize failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    fn require_body_table(&self, model_object_id: u64) -> Result<PagesTableGraph> {
        body_table_graphs(self)?
            .into_iter()
            .find(|graph| graph.info.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table model {model_object_id} is not attached to the body"
                ))
            })
    }
}

fn body_table_graphs(editor: &PagesEditor) -> Result<Vec<PagesTableGraph>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    let mut seen_drawables = HashSet::new();
    let mut seen_models = HashSet::new();
    let mut result = Vec::new();

    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        let Some(attachment) = decode_optional_typed_package_object::<DrawableAttachmentArchive>(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )?
        else {
            continue;
        };
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        let archive_name = find_object_archive(editor.package(), drawable.identifier)?;
        let archive = editor.package().archive(&archive_name)?;
        let object = archive.object(drawable.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table drawable {} is missing",
                drawable.identifier
            ))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} repeats its table-info payload",
                drawable.identifier
            )));
        };
        let table_info = TableInfoArchive::decode(message.data.as_slice())?;
        if table_info.super_.parent.map(|parent| parent.identifier) != Some(editor.body_storage_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} is not owned by the body",
                drawable.identifier
            )));
        }
        if body_units.get(entry.character_index as usize) != Some(&OBJECT_REPLACEMENT_CHARACTER) {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} has no object-replacement character",
                drawable.identifier
            )));
        }
        let model_id = table_info.table_model.identifier;
        let model_archive_name = find_object_archive(editor.package(), model_id)?;
        let model_archive = editor.package().archive(&model_archive_name)?;
        let model_object = model_archive.object(model_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages table model {model_id} is missing"))
        })?;
        let models = model_object
            .messages
            .iter()
            .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
            .filter_map(|message| TableModelArchive::decode(message.data.as_slice()).ok())
            .collect::<Vec<_>>();
        let [model] = models.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages table model {model_id} must contain exactly one table-model payload"
            )));
        };
        if !seen_drawables.insert(drawable.identifier) || !seen_models.insert(model_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} or model {model_id} is attached more than once",
                drawable.identifier
            )));
        }
        result.push(PagesTableGraph {
            info: PagesTableInfo {
                drawable_object_id: drawable.identifier,
                model_object_id: model_id,
                anchor_character_index: entry.character_index as usize,
                name: model.table_name.clone(),
                rows: model.number_of_rows as usize,
                columns: model.number_of_columns as usize,
            },
        });
    }
    result.sort_by_key(|graph| graph.info.anchor_character_index);
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;

    #[test]
    fn source_built_table_roundtrips_cell_updates() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Report 🙂\n")
            .body_table("Results", 3, 2)
            .build()
            .unwrap();
        let tables = editor.tables().unwrap();
        assert_eq!(tables.len(), 1);
        let info = &tables[0];
        assert_eq!(info.name, "Results");
        assert_eq!((info.rows, info.columns), (3, 2));
        assert_eq!(
            info.anchor_character_index,
            "Report 🙂\n".encode_utf16().count()
        );
        let model_id = info.model_object_id;
        assert!(editor.table(model_id).unwrap().cells.is_empty());

        editor
            .set_table_cell(model_id, 0, 0, PagesCellValue::Text("Header".to_owned()))
            .unwrap();
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(42.5))
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let table = reopened.table(model_id).unwrap();
        assert_eq!(
            table.get_cell(0, 0),
            Some(&PagesCellValue::Text("Header".to_owned()))
        );
        assert_eq!(table.get_cell(1, 1), Some(&PagesCellValue::Number(42.5)));
        reopened.clear_table_cell(model_id, 0, 0).unwrap();
        assert!(reopened.table(model_id).unwrap().get_cell(0, 0).is_none());
    }

    #[test]
    fn out_of_bounds_cell_update_is_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Bounded", 2, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell(model_id, 2, 0, PagesCellValue::Boolean(true))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn source_built_table_roundtrips_rename_and_resize() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Original", 3, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Text("kept".to_owned()))
            .unwrap();

        editor.rename_table(model_id, "Renamed").unwrap();
        editor.resize_table(model_id, 5, 4).unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let info = reopened.tables().unwrap().remove(0);
        assert_eq!(info.name, "Renamed");
        assert_eq!((info.rows, info.columns), (5, 4));
        assert_eq!(
            reopened.table(model_id).unwrap().get_cell(1, 1),
            Some(&PagesCellValue::Text("kept".to_owned()))
        );

        reopened.resize_table(model_id, 2, 2).unwrap();
        let info = reopened.tables().unwrap().remove(0);
        assert_eq!((info.rows, info.columns), (2, 2));
    }

    #[test]
    fn table_rename_and_occupied_shrink_are_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Protected", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_cell(model_id, 2, 2, PagesCellValue::Number(7.0))
            .unwrap();

        let before = editor.to_bytes().unwrap();
        assert!(editor.rename_table(model_id, "").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(editor.resize_table(model_id, 2, 2).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
