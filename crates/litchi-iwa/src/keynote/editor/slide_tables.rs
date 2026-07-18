//! Native table CRUD for Keynote slides.

use std::collections::{HashMap, HashSet};

use super::*;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

mod formula;
mod graph;
mod storage;

pub use formula::{
    KeynoteTableFormulaAxisReference, KeynoteTableFormulaBinaryOperator,
    KeynoteTableFormulaCachedValue, KeynoteTableFormulaCellReference,
    KeynoteTableFormulaExpression,
};
use graph::{require_table_model, slide_table_graph, table_template};
use storage::{remove_objects, set_table_geometry_in_package, set_uniform_table_dimensions};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_GEOMETRY_FLAGS: u32 = 3;
const TABLE_ANGLE_DEGREES: f32 = 0.0;

/// Strongly typed value stored in a Keynote table cell.
pub type KeynoteTableCellValue = crate::numbers::CellValue;
/// A validated non-zero native header or footer count.
pub type KeynoteTableHeaderCount = crate::numbers::NumbersTableHeaderCount;
/// Lossless header/footer configuration shared by native iWork tables.
pub type KeynoteTableHeaderSettings = crate::numbers::NumbersTableHeaderSettings;
/// One row or column addressed by zero-based index.
pub type KeynoteTableDimension = crate::numbers::NumbersTableDimension;
/// A validated positive point measurement for a table axis.
pub type KeynoteTablePoints = crate::numbers::NumbersTablePoints;
/// Either a table style's default axis size or an explicit point override.
pub type KeynoteTableDimensionSize = crate::numbers::NumbersTableDimensionSize;

/// Stable identity, dimensions, and geometry of one slide-owned table.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideTableInfo {
    pub slide_index: usize,
    pub slide_id: u64,
    pub drawable_object_id: u64,
    pub model_object_id: u64,
    pub name: String,
    pub rows: usize,
    pub columns: usize,
    pub geometry: DrawableGeometry,
}

/// Materialized non-empty cells from one Keynote table.
#[derive(Debug, Clone)]
pub struct KeynoteSlideTable {
    pub info: KeynoteSlideTableInfo,
    pub cells: HashMap<(usize, usize), KeynoteTableCellValue>,
}

impl KeynoteSlideTable {
    pub fn get_cell(&self, row: usize, column: usize) -> Option<&KeynoteTableCellValue> {
        self.cells.get(&(row, column))
    }
}

/// Result of removing a slide table and its private storage graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideTable {
    pub table: KeynoteSlideTableInfo,
}

impl KeynoteEditor {
    /// List native tables owned directly by one slide in z-order.
    pub fn slide_tables(&self, slide_index: usize) -> Result<Vec<KeynoteSlideTableInfo>> {
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_create::text_box_context(&graph, slide_index)?;
        let mut tables = Vec::new();
        for reference in &context.slide.drawables_z_order {
            let Some(messages) = graph.objects.get(&reference.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {} drawable {} is missing",
                    context.slide_id, reference.identifier
                )));
            };
            if messages
                .iter()
                .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            {
                tables.push(slide_table_graph(self, slide_index, reference.identifier)?.info);
            }
        }
        Ok(tables)
    }

    /// Read all materialized cell values from one reachable slide table.
    pub fn slide_table(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<KeynoteSlideTable> {
        let info = self
            .slide_tables(slide_index)?
            .into_iter()
            .find(|table| table.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote table model {model_object_id} is not owned by slide {slide_index}"
                ))
            })?;
        let bytes = self.package().to_bytes()?;
        let bundle = Bundle::from_bytes(&bytes)?;
        let index = ObjectIndex::from_bundle(&bundle)?;
        let object = index
            .resolve_object(&bundle, model_object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote table model {model_object_id} is missing"))
            })?;
        let table = TableDataExtractor::new(&bundle, &index)
            .extract_table_from_object(&object)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote object {model_object_id} has no native table model"
                ))
            })?;
        Ok(KeynoteSlideTable {
            info,
            cells: table.cells,
        })
    }

    /// Add an independently editable native table directly to a slide.
    pub fn add_slide_table(
        &mut self,
        slide_index: usize,
        name: &str,
        rows: usize,
        columns: usize,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<KeynoteSlideTableInfo> {
        let geometry = DrawableGeometry {
            position: Some(position),
            size: Some(size),
            flags: Some(TABLE_GEOMETRY_FLAGS),
            angle: Some(TABLE_ANGLE_DEGREES),
        }
        .validate()?;
        let object_graph = ObjectGraph::read(self.package())?;
        let context = text_box_create::text_box_context(&object_graph, slide_index)?;
        let slide_archive = object_graph.archive_name(context.slide_id)?.to_owned();
        let (template_info_id, template_model_id) = table_template(self.package())?;

        let mut staged = self.package().clone();
        let (info_id, model_id) = crate::numbers::editor::create_empty_table_graph_in_package(
            &mut staged,
            template_info_id,
            template_model_id,
            context.slide_id,
            name,
            rows,
            columns,
        )?;
        set_table_geometry_in_package(&mut staged, info_id, geometry)?;
        set_uniform_table_dimensions(&mut staged, model_id, rows, columns, size)?;
        patch_slide_drawable_references(
            &mut staged,
            &slide_archive,
            context.slide_id,
            None,
            Some(info_id),
        )?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = slide_table_graph(&verified, slide_index, info_id)?;
        if created.info.model_object_id != model_id
            || created.info.name != name
            || (created.info.rows, created.info.columns) != (rows, columns)
            || created.info.geometry != geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote table creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Set or clear one cell in a reachable slide table transactionally.
    pub fn set_slide_table_cell(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        value: KeynoteTableCellValue,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            value,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Clear one cell in a reachable slide table.
    pub fn clear_slide_table_cell(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_slide_table_cell(
            slide_index,
            model_object_id,
            row,
            column,
            KeynoteTableCellValue::Empty,
        )
    }

    /// Rename a reachable slide table.
    pub fn rename_slide_table(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        name: &str,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::rename_table_in_package(&mut staged, model_object_id, name)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if require_table_model(&verified, slide_index, model_object_id)?.name != name {
            return Err(Error::InvalidFormat(
                "Keynote table rename failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Resize a reachable slide table while preserving retained cells.
    pub fn resize_slide_table(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        rows: usize,
        columns: usize,
    ) -> Result<()> {
        let source = require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::resize_table_in_package(
            &mut staged,
            model_object_id,
            rows,
            columns,
        )?;
        if let Some(size) = source.geometry.size {
            set_uniform_table_dimensions(&mut staged, model_object_id, rows, columns, size)?;
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let resized = require_table_model(&verified, slide_index, model_object_id)?;
        if (resized.rows, resized.columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Keynote table resize failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read a slide table's lossless header and footer configuration.
    pub fn slide_table_header_settings(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<KeynoteTableHeaderSettings> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_header_settings_in_package(self.package(), model_object_id)
    }

    /// Replace a slide table's header and footer configuration transactionally.
    pub fn set_slide_table_header_settings(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        settings: KeynoteTableHeaderSettings,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_header_settings_in_package(
            &mut staged,
            model_object_id,
            settings,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_table_header_settings(slide_index, model_object_id)? != settings {
            return Err(Error::InvalidFormat(
                "Keynote table header settings failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one row-height or column-width override.
    pub fn slide_table_dimension_size(
        &self,
        slide_index: usize,
        model_object_id: u64,
        dimension: KeynoteTableDimension,
    ) -> Result<KeynoteTableDimensionSize> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_dimension_size_in_package(
            self.package(),
            model_object_id,
            dimension,
        )
    }

    /// Set or clear one row-height or column-width override transactionally.
    ///
    /// The drawable bounds are updated to the sum of the effective native row
    /// heights and column widths so Keynote's selection box remains exact.
    pub fn set_slide_table_dimension_size(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        dimension: KeynoteTableDimension,
        size: KeynoteTableDimensionSize,
    ) -> Result<()> {
        let source = require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_dimension_size_in_package(
            &mut staged,
            model_object_id,
            dimension,
            size,
        )?;
        let (width, height) =
            crate::numbers::editor::table_size_points_in_package(&staged, model_object_id)?;
        let geometry = DrawableGeometry {
            size: Some(DrawableSize { width, height }),
            ..source.geometry
        };
        set_table_geometry_in_package(&mut staged, source.drawable_object_id, geometry)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_table_dimension_size(slide_index, model_object_id, dimension)? != size
            || require_table_model(&verified, slide_index, model_object_id)?.geometry != geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote table dimension update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one row-height override.
    pub fn slide_table_row_height(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
    ) -> Result<KeynoteTableDimensionSize> {
        self.slide_table_dimension_size(
            slide_index,
            model_object_id,
            KeynoteTableDimension::Row(row),
        )
    }

    /// Set or clear one row-height override.
    pub fn set_slide_table_row_height(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        size: KeynoteTableDimensionSize,
    ) -> Result<()> {
        self.set_slide_table_dimension_size(
            slide_index,
            model_object_id,
            KeynoteTableDimension::Row(row),
            size,
        )
    }

    /// Read one column-width override.
    pub fn slide_table_column_width(
        &self,
        slide_index: usize,
        model_object_id: u64,
        column: usize,
    ) -> Result<KeynoteTableDimensionSize> {
        self.slide_table_dimension_size(
            slide_index,
            model_object_id,
            KeynoteTableDimension::Column(column),
        )
    }

    /// Set or clear one column-width override.
    pub fn set_slide_table_column_width(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        column: usize,
        size: KeynoteTableDimensionSize,
    ) -> Result<()> {
        self.set_slide_table_dimension_size(
            slide_index,
            model_object_id,
            KeynoteTableDimension::Column(column),
            size,
        )
    }

    /// Update one slide table's position, size, flags, and rotation.
    pub fn set_slide_table_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let geometry = geometry.validate()?;
        let source = slide_table_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_table_geometry_in_package(&mut staged, drawable_object_id, geometry)?;
        if let Some(size) = geometry.size {
            set_uniform_table_dimensions(
                &mut staged,
                source.info.model_object_id,
                source.info.rows,
                source.info.columns,
                size,
            )?;
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if slide_table_graph(&verified, slide_index, drawable_object_id)?
            .info
            .geometry
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote table geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a slide table and every private storage/formula object it owns.
    pub fn remove_slide_table(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideTable> {
        let source = slide_table_graph(self, slide_index, drawable_object_id)?;
        let mut object_ids = crate::numbers::editor::table_owned_object_ids_in_package(
            self.package(),
            source.info.model_object_id,
        )?;
        object_ids.extend([drawable_object_id, source.info.model_object_id]);
        let mut seen = HashSet::with_capacity(object_ids.len());
        object_ids.retain(|identifier| seen.insert(*identifier));

        let mut staged = self.package().clone();
        patch_slide_drawable_references(
            &mut staged,
            &source.slide_archive,
            source.info.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        remove_component_external_references_to_object(
            &mut staged,
            source.slide_component_id,
            drawable_object_id,
        )?;
        let formula_ids = crate::numbers::editor::remove_table_formula_graph_in_package(
            &mut staged,
            &object_ids,
        )?;
        remove_objects(&mut staged, &object_ids)?;
        let mut released = object_ids;
        released.extend(formula_ids);
        release_package_identifier_suffix(&mut staged, &released)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_tables(slide_index)?
            .iter()
            .any(|table| table.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote table deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteSlideTable { table: source.info })
    }
}

#[cfg(test)]
mod tests;
