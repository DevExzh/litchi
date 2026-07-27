//! Native table CRUD for Keynote slides.

use std::collections::{HashMap, HashSet};

use super::*;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use crate::table_appearance::TableAppearance;
use crate::table_lock::TableLockState;

mod appearance;
mod comments;
mod formula;
mod graph;
mod hidden_axes;
mod lock;
mod sort;
mod storage;
mod title;
mod topology;

pub use comments::{
    KeynoteTableCellComment, KeynoteTableCellCommentInfo, KeynoteTableCellCommentReplyInfo,
};
pub use formula::{
    KeynoteTableFormulaAxisReference, KeynoteTableFormulaBinaryOperator,
    KeynoteTableFormulaCachedValue, KeynoteTableFormulaCellReference,
    KeynoteTableFormulaExpression,
};
use graph::{require_table_model, slide_table_graph, table_template};
pub use hidden_axes::{KeynoteTableAxisIndex, KeynoteTableHiddenAxes};
pub use sort::{
    KeynoteTableSortColumnIndex, KeynoteTableSortDirection, KeynoteTableSortOrder,
    KeynoteTableSortRowRange, KeynoteTableSortRule, KeynoteTableSortScope,
};
use storage::{remove_objects, set_table_geometry_in_package, set_uniform_table_dimensions};
pub use title::KeynoteTableTitleSettings;

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_GEOMETRY_FLAGS: u32 = 3;
const TABLE_ANGLE_DEGREES: f32 = 0.0;

/// Strongly typed value stored in a Keynote table cell.
pub type KeynoteTableCellValue = crate::numbers::CellValue;
/// One mutation in a transactional Keynote table-cell batch.
pub type KeynoteTableCellUpdate = crate::numbers::TableCellUpdate;
/// Section-relative row deletion shared by native iWork tables.
pub type KeynoteTableRowDeletion = crate::numbers::TableRowDeletion;
/// Section-relative column deletion shared by native iWork tables.
pub type KeynoteTableColumnDeletion = crate::numbers::TableColumnDeletion;
/// Section-relative row insertion shared by native iWork tables.
pub type KeynoteTableRowInsertion = crate::numbers::TableRowInsertion;
/// Section-relative column insertion shared by native iWork tables.
pub type KeynoteTableColumnInsertion = crate::numbers::TableColumnInsertion;
/// A validated native merged-cell rectangle.
pub type KeynoteTableCellRegion = crate::numbers::editor::IWorkTableCellRegion;
pub use crate::table_cell_border::{
    TableCellBorderSide as KeynoteTableCellBorderSide, TableCellBorders as KeynoteTableCellBorders,
};
pub use crate::table_cell_data_format::{
    TableCellCheckboxFormat as KeynoteTableCellCheckboxFormat,
    TableCellCurrencyFormat as KeynoteTableCellCurrencyFormat,
    TableCellCustomFormat as KeynoteTableCellCustomFormat,
    TableCellDataFormat as KeynoteTableCellDataFormat,
    TableCellDateTimeFormat as KeynoteTableCellDateTimeFormat,
    TableCellDurationFormat as KeynoteTableCellDurationFormat,
    TableCellDurationStyle as KeynoteTableCellDurationStyle,
    TableCellDurationUnit as KeynoteTableCellDurationUnit,
    TableCellDurationUnitRange as KeynoteTableCellDurationUnitRange,
    TableCellDurationUnits as KeynoteTableCellDurationUnits,
    TableCellFractionFormat as KeynoteTableCellFractionFormat,
    TableCellNumeralSystemFormat as KeynoteTableCellNumeralSystemFormat,
    TableCellPercentageFormat as KeynoteTableCellPercentageFormat,
    TableCellPopUpMenuFormat as KeynoteTableCellPopUpMenuFormat,
    TableCellPopUpMenuInitialSelection as KeynoteTableCellPopUpMenuInitialSelection,
    TableCellPopUpMenuItem as KeynoteTableCellPopUpMenuItem,
    TableCellScientificFormat as KeynoteTableCellScientificFormat,
    TableCellSliderDisplayFormat as KeynoteTableCellSliderDisplayFormat,
    TableCellSliderFormat as KeynoteTableCellSliderFormat,
    TableCellSliderRange as KeynoteTableCellSliderRange,
    TableCellStarRatingFormat as KeynoteTableCellStarRatingFormat,
    TableCellStepperDisplayFormat as KeynoteTableCellStepperDisplayFormat,
    TableCellStepperFormat as KeynoteTableCellStepperFormat,
    TableCellStepperRange as KeynoteTableCellStepperRange,
    TableCellTextFormat as KeynoteTableCellTextFormat,
};
pub use crate::table_cell_layout::{
    TableCellInset as KeynoteTableCellInset, TableCellInsets as KeynoteTableCellInsets,
    TableCellLayout as KeynoteTableCellLayout, TableCellTextWrap as KeynoteTableCellTextWrap,
    TableCellVerticalAlignment as KeynoteTableCellVerticalAlignment,
};
pub use crate::table_cell_number_format::{
    TableCellDecimalPlaces as KeynoteTableCellDecimalPlaces,
    TableCellFixedDecimalPlaces as KeynoteTableCellFixedDecimalPlaces,
    TableCellNegativeNumberStyle as KeynoteTableCellNegativeNumberStyle,
    TableCellNumberFormat as KeynoteTableCellNumberFormat,
    TableCellThousandsSeparator as KeynoteTableCellThousandsSeparator,
};
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
    /// Effective alternating-row and automatic-sizing settings.
    pub appearance: TableAppearance,
    /// Interactive editing lock shown in the Arrange inspector.
    pub lock_state: TableLockState,
}

/// Materialized non-empty cells from one Keynote table.
#[derive(Debug, Clone)]
pub struct KeynoteSlideTable {
    pub info: KeynoteSlideTableInfo,
    pub cells: HashMap<(usize, usize), KeynoteTableCellValue>,
    /// Comments indexed independently from cell values by `(row, column)`.
    pub comments: HashMap<(usize, usize), KeynoteTableCellComment>,
    /// Native merged-cell rectangles in formula-store order.
    pub merges: Vec<KeynoteTableCellRegion>,
}

impl KeynoteSlideTable {
    pub fn get_cell(&self, row: usize, column: usize) -> Option<&KeynoteTableCellValue> {
        self.cells.get(&(row, column))
    }

    /// Borrow the comment attached to a materialized cell, if any.
    pub fn get_comment(&self, row: usize, column: usize) -> Option<&KeynoteTableCellComment> {
        self.comments.get(&(row, column))
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
            comments: table.comments,
            merges: crate::numbers::editor::table_cell_merges_in_package(
                self.package(),
                model_object_id,
            )?,
        })
    }

    /// List native merged-cell rectangles in one slide-owned table.
    pub fn slide_table_cell_merges(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<Vec<KeynoteTableCellRegion>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_merges_in_package(self.package(), model_object_id)
    }

    /// Merge one non-overlapping slide-table rectangle transactionally.
    pub fn merge_slide_table_cells(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        region: KeynoteTableCellRegion,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::merge_table_cells_in_package(&mut staged, model_object_id, region)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if !verified
            .slide_table_cell_merges(slide_index, model_object_id)?
            .contains(&region)
        {
            return Err(Error::InvalidFormat(
                "Keynote table-cell merge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove one exact slide-table merge, returning whether it existed.
    pub fn unmerge_slide_table_cells(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        region: KeynoteTableCellRegion,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::unmerge_table_cells_in_package(
            &mut staged,
            model_object_id,
            region,
        )?;
        if !changed {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_cell_merges(slide_index, model_object_id)?
            .contains(&region)
        {
            return Err(Error::InvalidFormat(
                "Keynote table-cell unmerge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
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

    /// Duplicate a populated table on one slide with independent storage.
    ///
    /// The clone receives fresh object identifiers, table UUID, and
    /// CalculationEngine owner state. It is appended to the slide's native
    /// drawable lists and offset by ten points so both tables remain directly
    /// selectable in Keynote.
    pub fn duplicate_slide_table(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<KeynoteSlideTableInfo> {
        let source = slide_table_graph(self, slide_index, drawable_object_id)?;
        let tables = self.slide_tables(slide_index)?;
        let existing_names = tables
            .iter()
            .map(|table| table.name.as_str())
            .collect::<HashSet<_>>();
        let name =
            crate::numbers::editor::duplicate_table_name(&source.info.name, &existing_names)?;
        let source_rows = source.info.rows;
        let source_columns = source.info.columns;
        let mut expected_geometry = source.info.geometry;
        if let Some(position) = expected_geometry.position.as_mut() {
            position.x += TABLE_DUPLICATE_OFFSET;
            position.y += TABLE_DUPLICATE_OFFSET;
        }

        let package = self.package();
        let mut staged = package.clone();
        let cloned = crate::numbers::editor::duplicate_attached_table_graph_in_package(
            package,
            &mut staged,
            source.info.drawable_object_id,
            source.info.model_object_id,
            &name,
            TABLE_DUPLICATE_OFFSET,
        )?;
        patch_slide_drawable_references(
            &mut staged,
            &source.slide_archive,
            source.info.slide_id,
            None,
            Some(cloned.info_object_id),
        )?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = slide_table_graph(&verified, slide_index, cloned.info_object_id)?;
        if created.info.model_object_id != cloned.model_object_id
            || created.info.name != name
            || (created.info.rows, created.info.columns) != (source_rows, source_columns)
            || created.info.geometry != expected_geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote table duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Set or clear one cell in a reachable slide table transactionally.
    ///
    /// Supported dependent formula caches are refreshed before commit;
    /// unsupported impacted formulas reject the edit instead of remaining
    /// visibly stale.
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

    /// Set several slide-table cells with one package clone and dependency pass.
    ///
    /// Coordinates must be unique. Any invalid value, coordinate, or impacted
    /// formula rejects the complete batch without changing the editor.
    pub fn set_slide_table_cells(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        updates: impl IntoIterator<Item = KeynoteTableCellUpdate>,
    ) -> Result<usize> {
        require_table_model(self, slide_index, model_object_id)?;
        let batch = crate::numbers::editor::TableCellBatch::collect(updates)?;
        if batch.is_empty() {
            return Ok(0);
        }
        let expected = batch.len();
        let mut staged = self.package().clone();
        let applied = batch.apply_attached(&mut staged, model_object_id)?;
        if applied != expected {
            return Err(Error::InvalidFormat(format!(
                "Keynote table-cell batch applied {applied} updates, expected {expected}"
            )));
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        *self = verified;
        Ok(applied)
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

    /// Read the explicit typed data format for one slide-table cell.
    pub fn slide_table_cell_data_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<KeynoteTableCellDataFormat> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_data_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create, replace, or reset one slide-table cell's data format.
    pub fn set_slide_table_cell_data_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellDataFormat,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_data_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            &format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
            != format
        {
            return Err(Error::InvalidFormat(
                "Keynote table-cell data format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read an explicit decimal-number format for one slide-table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn slide_table_cell_number_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellNumberFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_number_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit decimal-number format transactionally.
    pub fn set_slide_table_cell_number_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellNumberFormat,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_number_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified.slide_table_cell_number_format(slide_index, model_object_id, row, column)?
            != Some(format)
        {
            return Err(Error::InvalidFormat(
                "Keynote table-cell number format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore iWork's automatic data format for one slide-table cell.
    pub fn reset_slide_table_cell_number_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_number_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified
                .slide_table_cell_number_format(slide_index, model_object_id, row, column)?
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "Keynote table-cell number-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Text format for one slide-table cell.
    pub fn slide_table_cell_text_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellTextFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_text_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Text format transactionally.
    pub fn set_slide_table_cell_text_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            KeynoteTableCellTextFormat.into(),
        )
    }

    /// Restore Automatic from an explicit Text slide-table cell.
    pub fn reset_slide_table_cell_text_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Text-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read a named custom format for one slide-table cell.
    pub fn slide_table_cell_custom_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellCustomFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_custom_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a named custom slide-table format transactionally.
    pub fn set_slide_table_cell_custom_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellCustomFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from a named custom slide-table format.
    pub fn reset_slide_table_cell_custom_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_custom_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Custom-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit currency format for one slide-table cell.
    pub fn slide_table_cell_currency_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellCurrencyFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_currency_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit currency format transactionally.
    pub fn set_slide_table_cell_currency_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellCurrencyFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Currency slide-table cell.
    pub fn reset_slide_table_cell_currency_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_currency_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote currency-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit percentage format for one slide-table cell.
    pub fn slide_table_cell_percentage_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellPercentageFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_percentage_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit percentage format transactionally.
    pub fn set_slide_table_cell_percentage_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellPercentageFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Percentage slide-table cell.
    pub fn reset_slide_table_cell_percentage_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_percentage_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote percentage-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit scientific-notation format for one slide-table cell.
    pub fn slide_table_cell_scientific_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellScientificFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_scientific_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit scientific-notation format transactionally.
    pub fn set_slide_table_cell_scientific_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellScientificFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Scientific slide-table cell.
    pub fn reset_slide_table_cell_scientific_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_scientific_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote scientific-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit mixed-fraction format for one slide-table cell.
    pub fn slide_table_cell_fraction_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellFractionFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_fraction_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit mixed-fraction format transactionally.
    pub fn set_slide_table_cell_fraction_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellFractionFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Fraction slide-table cell.
    pub fn reset_slide_table_cell_fraction_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_fraction_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote fraction-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit positional numeral-system format for one slide-table cell.
    pub fn slide_table_cell_numeral_system_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellNumeralSystemFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_numeral_system_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit positional numeral-system format transactionally.
    pub fn set_slide_table_cell_numeral_system_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellNumeralSystemFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Numeral System slide-table cell.
    pub fn reset_slide_table_cell_numeral_system_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_numeral_system_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote numeral-system reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Date & Time format for one slide-table cell.
    pub fn slide_table_cell_date_time_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellDateTimeFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_date_time_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Date & Time format transactionally.
    pub fn set_slide_table_cell_date_time_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellDateTimeFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Date & Time slide-table cell.
    pub fn reset_slide_table_cell_date_time_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_date_time_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Date & Time reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Duration format for one slide-table cell.
    pub fn slide_table_cell_duration_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellDurationFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_duration_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Duration format transactionally.
    pub fn set_slide_table_cell_duration_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellDurationFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Duration slide-table cell.
    pub fn reset_slide_table_cell_duration_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_duration_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Duration reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Checkbox format for one slide-table cell.
    pub fn slide_table_cell_checkbox_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellCheckboxFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_checkbox_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Checkbox format transactionally.
    pub fn set_slide_table_cell_checkbox_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellCheckboxFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Checkbox slide-table cell.
    pub fn reset_slide_table_cell_checkbox_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_checkbox_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Checkbox reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Star Rating format for one slide-table cell.
    pub fn slide_table_cell_star_rating_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellStarRatingFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_star_rating_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native five-star rating transactionally.
    pub fn set_slide_table_cell_star_rating_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellStarRatingFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Star Rating slide-table cell.
    pub fn reset_slide_table_cell_star_rating_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_star_rating_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Star Rating reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Slider format for one slide-table cell.
    pub fn slide_table_cell_slider_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellSliderFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_slider_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Slider format transactionally.
    pub fn set_slide_table_cell_slider_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellSliderFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Slider slide-table cell.
    pub fn reset_slide_table_cell_slider_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_slider_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Slider reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Stepper format for one slide-table cell.
    pub fn slide_table_cell_stepper_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellStepperFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_stepper_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Stepper format transactionally.
    pub fn set_slide_table_cell_stepper_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellStepperFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Stepper slide-table cell.
    pub fn reset_slide_table_cell_stepper_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_stepper_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Stepper reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Pop-Up Menu format for one slide-table cell.
    pub fn slide_table_cell_pop_up_menu_format(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellPopUpMenuFormat>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_pop_up_menu_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Pop-Up Menu format transactionally.
    pub fn set_slide_table_cell_pop_up_menu_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: KeynoteTableCellPopUpMenuFormat,
    ) -> Result<()> {
        self.set_slide_table_cell_data_format(
            slide_index,
            model_object_id,
            row,
            column,
            format.into(),
        )
    }

    /// Restore Automatic from an explicit Pop-Up Menu slide-table cell.
    pub fn reset_slide_table_cell_pop_up_menu_format(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_pop_up_menu_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            if verified.slide_table_cell_data_format(slide_index, model_object_id, row, column)?
                != KeynoteTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Keynote Pop-Up Menu reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective text layout for one reachable slide-table cell.
    pub fn slide_table_cell_layout(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<KeynoteTableCellLayout> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_layout_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace local text-layout overrides for one slide-table cell.
    pub fn set_slide_table_cell_layout(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        layout: KeynoteTableCellLayout,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_layout_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            layout,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified.slide_table_cell_layout(slide_index, model_object_id, row, column)? != layout {
            return Err(Error::InvalidFormat(
                "Keynote table-cell layout failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local text-layout overrides and restore inherited cell values.
    pub fn reset_slide_table_cell_layout(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_layout_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective fill for one reachable slide-table cell.
    pub fn slide_table_cell_fill(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::shapes::ShapeFill> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_fill_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace one local slide-table cell fill.
    pub fn set_slide_table_cell_fill(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        fill: &crate::shapes::ShapeFill,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_fill_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            fill,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if &verified.slide_table_cell_fill(slide_index, model_object_id, row, column)? != fill {
            return Err(Error::InvalidFormat(
                "Keynote table-cell fill failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct fill override and restore the inherited table style.
    pub fn reset_slide_table_cell_fill(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_fill_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            require_table_model(&verified, slide_index, model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective explicit borders for one reachable slide-table cell.
    pub fn slide_table_cell_borders(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<KeynoteTableCellBorders> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_borders_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace one explicit slide-table cell border.
    pub fn set_slide_table_cell_border(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: KeynoteTableCellBorderSide,
        stroke: crate::shapes::ShapeStroke,
    ) -> Result<()> {
        self.update_slide_table_cell_border(
            slide_index,
            model_object_id,
            row,
            column,
            side,
            Some(stroke),
        )
    }

    /// Explicitly clear one slide-table cell border.
    pub fn clear_slide_table_cell_border(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: KeynoteTableCellBorderSide,
    ) -> Result<()> {
        self.update_slide_table_cell_border(slide_index, model_object_id, row, column, side, None)
    }

    fn update_slide_table_cell_border(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: KeynoteTableCellBorderSide,
        stroke: Option<crate::shapes::ShapeStroke>,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_border_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            side,
            stroke,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_cell_borders(slide_index, model_object_id, row, column)?
            .get(side)
            != stroke
        {
            return Err(Error::InvalidFormat(
                "Keynote table-cell border failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
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
