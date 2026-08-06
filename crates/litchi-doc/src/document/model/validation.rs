//! Structural table validation and materialization.
//!
//! The semantic facade delegates table grouping here so the document model has
//! one canonical implementation for nesting, row-end, and cell-end invariants.
//! The helpers remain `pub(super)` because they are implementation seams, not
//! additional public API.

use super::state::Document;
use crate::package::Result;
use crate::paragraph::Paragraph;
use crate::table::Table;

impl Document {
    /// Extract tables from a list of paragraphs at a specific nesting level.
    pub(super) fn extract_tables_from_paragraphs(
        &self,
        paragraphs: &[Paragraph],
        level: i32,
    ) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        let mut i = 0;

        while i < paragraphs.len() {
            let para = &paragraphs[i];
            let props = para.properties();

            // A table starts at the requested nesting level and continues
            // through nested table paragraphs until its level is left.
            if props.in_table && props.table_nesting_level == level {
                let mut table_paras = Vec::new();

                while i < paragraphs.len() {
                    let current_para = &paragraphs[i];
                    let current_props = current_para.properties();

                    if !current_props.in_table || current_props.table_nesting_level < level {
                        break;
                    }

                    table_paras.push(current_para.clone());
                    i += 1;
                }

                let rows = self.extract_rows_from_table_paragraphs(&table_paras, level)?;

                if !rows.is_empty() {
                    let properties = rows.first().and_then(|row| row.properties()).cloned();
                    if let Some(properties) = properties {
                        tables.push(Table::with_properties(rows, properties));
                    } else {
                        tables.push(Table::new(rows));
                    }
                }
            } else {
                i += 1;
            }
        }

        Ok(tables)
    }

    /// Convert table paragraphs into rows while enforcing nesting boundaries
    /// and retaining the source TAP metadata used by snapshot edits.
    pub(super) fn extract_rows_from_table_paragraphs(
        &self,
        table_paras: &[Paragraph],
        level: i32,
    ) -> Result<Vec<crate::table::Row>> {
        use crate::table::Row;

        let mut rows = Vec::new();
        let mut current_row_paras = Vec::new();

        for para in table_paras {
            let props = para.properties();

            if props.table_nesting_level > level {
                continue;
            }

            current_row_paras.push(para.clone());

            if props.is_table_row_end && props.table_nesting_level == level {
                let cells = self.extract_cells_from_row_paragraphs(
                    &current_row_paras,
                    props.table_properties.as_ref(),
                )?;

                if !cells.is_empty() {
                    rows.push(Row::with_metadata(
                        cells,
                        props.table_properties.clone(),
                        para.table_formatting_revision().cloned(),
                        props.table_properties_preserved_for_revision,
                    ));
                }

                current_row_paras.clear();
            }
        }

        // Preserve an incomplete final row instead of silently dropping its
        // paragraphs; malformed snapshots remain inspectable and editable.
        if !current_row_paras.is_empty() {
            let last = current_row_paras
                .last()
                .expect("non-empty row paragraph collection");
            let cells = self.extract_cells_from_row_paragraphs(
                &current_row_paras,
                last.properties().table_properties.as_ref(),
            )?;
            if !cells.is_empty() {
                rows.push(Row::with_metadata(
                    cells,
                    last.properties().table_properties.clone(),
                    last.table_formatting_revision().cloned(),
                    last.properties().table_properties_preserved_for_revision,
                ));
            }
        }

        crate::table::apply_table_cell_styles(&mut rows);
        Ok(rows)
    }

    /// Split row paragraphs at cell markers while preserving cell properties.
    pub(super) fn extract_cells_from_row_paragraphs(
        &self,
        row_paras: &[Paragraph],
        table_properties: Option<&crate::parts::tap::TableProperties>,
    ) -> Result<Vec<crate::table::Cell>> {
        use crate::table::Cell;

        let mut cells = Vec::new();
        let mut cell_paragraphs = Vec::new();
        for para in row_paras {
            let props = para.properties();

            // The row-end marker carries no cell content.
            if props.is_table_row_end {
                continue;
            }

            cell_paragraphs.push(para.clone());
            if props.is_table_cell_end {
                let properties = table_properties
                    .and_then(|tap| tap.cell_properties.get(cells.len()))
                    .cloned();
                cells.push(Cell::with_properties(
                    std::mem::take(&mut cell_paragraphs),
                    properties,
                ));
            }
        }

        if !cell_paragraphs.is_empty() {
            let properties = table_properties
                .and_then(|tap| tap.cell_properties.get(cells.len()))
                .cloned();
            cells.push(Cell::with_properties(cell_paragraphs, properties));
        }

        if cells.is_empty() && !row_paras.is_empty() {
            cells.push(Cell::new(String::new()));
        }

        Ok(cells)
    }
}
