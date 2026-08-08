use super::super::model::validate_list_object_style;
use super::super::*;
use crate::error::{Error, Result};
use crate::{ListObject, MapInfo};

impl Writer {
    /// Install or replace the inert root-level XML-map catalog.
    ///
    /// Every existing mapped list column is resolved against the candidate
    /// before publication. Schemas, bindings, and XPath values are serialized
    /// as metadata only; no referenced resource is opened or refreshed.
    pub fn put_xml_map(&mut self, map_info: MapInfo) -> Result<Option<MapInfo>> {
        crate::xml_map::validate_info(&map_info)?;
        crate::xml_map::validate_list_objects(
            Some(&map_info),
            self.worksheets
                .iter()
                .flat_map(|worksheet| worksheet.list_objects.iter()),
        )?;
        Ok(self.xml_map.replace(map_info))
    }

    /// Remove the XML-map catalog when no mapped list column references it.
    pub fn remove_xml_map(&mut self) -> Result<Option<MapInfo>> {
        crate::xml_map::validate_list_objects(
            None,
            self.worksheets
                .iter()
                .flat_map(|worksheet| worksheet.list_objects.iter()),
        )?;
        Ok(self.xml_map.take())
    }

    /// Return the XML-map catalog configured for the next write.
    pub fn xml_map(&self) -> Option<&MapInfo> {
        self.xml_map.as_ref()
    }

    /// Installs a complete custom table-style family.
    ///
    /// Validation happens before assignment, so an error leaves the current
    /// writer configuration unchanged.
    pub fn set_custom_table_styles(&mut self, styles: CustomTableStyles) -> Result<()> {
        styles.validate(&self.fmt)?;
        self.custom_table_styles = Some(styles);
        Ok(())
    }

    /// Removes caller-defined table styles and restores the default write path.
    pub fn clear_custom_table_styles(&mut self) {
        self.custom_table_styles = None;
    }

    /// Adds a legacy BIFF8 worksheet table and writes its header captions.
    pub fn add_list_object(&mut self, sheet: usize, table: ListObject) -> Result<()> {
        table.validate()?;
        crate::xml_map::validate_list_objects(self.xml_map.as_ref(), std::iter::once(&table))?;
        let style = table.style().ok_or_else(|| {
            Error::InvalidData("validated table is missing its style".to_string())
        })?;
        validate_list_object_style(style.name(), self.custom_table_styles.as_ref())?;
        if self
            .worksheets
            .iter()
            .flat_map(|worksheet| &worksheet.list_objects)
            .any(|existing| {
                existing.id() == table.id() || existing.name().eq_ignore_ascii_case(table.name())
            })
        {
            return Err(Error::InvalidData(
                "table identifier or name collides within the workbook".to_string(),
            ));
        }
        if self
            .defined_names
            .iter()
            .any(|name| name.name.eq_ignore_ascii_case(table.name()))
            || self
                .defined_name_records
                .iter()
                .any(|(name, _)| name.name.eq_ignore_ascii_case(table.name()))
        {
            return Err(Error::InvalidData(
                "table name collides with a workbook defined name".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet
            .list_objects
            .iter()
            .any(|existing| existing.range().overlaps(table.range()))
        {
            return Err(Error::InvalidData(
                "table ranges overlap within the worksheet".to_string(),
            ));
        }
        if worksheet.auto_filter.is_some_and(|filter| {
            u32::from(table.range().first_row()) <= filter.last_row
                && filter.first_row <= u32::from(table.range().last_row())
                && table.range().first_column() <= filter.last_col
                && filter.first_col <= table.range().last_column()
        }) {
            return Err(Error::InvalidData(
                "table range overlaps the worksheet AutoFilter".to_string(),
            ));
        }
        let mut header_cells = Vec::new();
        for (offset, column) in table
            .columns()
            .iter()
            .enumerate()
            .filter(|_| table.has_header_row())
        {
            let key = (
                u32::from(table.range().first_row()),
                table.range().first_column() + offset as u16,
            );
            if let Some(cell) = worksheet.cells.get(&key)
                && !matches!(&cell.value, CellValue::String(value) if value == column.name())
            {
                return Err(Error::InvalidData(
                    "table header collides with a different cell value".to_string(),
                ));
            } else if !worksheet.cells.contains_key(&key) {
                header_cells.push(WritableCell::new(
                    CellPos::try_new(key.0, key.1)?,
                    CellValue::String(column.name().to_string()),
                    0,
                    None,
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.include_list_object_range(table.range());
        for cell in header_cells {
            worksheet.add_cell(cell);
        }
        worksheet.list_objects.push(table);
        Ok(())
    }

    pub fn clear_list_objects(&mut self, sheet: usize) -> Result<()> {
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .list_objects
            .clear();
        Ok(())
    }
}
