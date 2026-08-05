//! Semantic row assembly and bounded blank-cell deferral.

use super::{
    Cell, CellBuilder, CellTextContent, Error, MAX_EXPANDED_CELLS_PER_ROW, Result, TableVisibility,
};

/// Builder for constructing a semantic [`Row`] during content traversal.
pub(crate) struct RowBuilder {
    cells: Vec<Cell>,
    repeated: usize,
    style_name: Option<String>,
    default_cell_style_name: Option<String>,
    visibility: TableVisibility,
    /// Number of attribute-free filler cells read but not yet materialised.
    deferred_blank_cells: usize,
    /// The filler cell to clone when the deferred run has to be materialised.
    deferred_blank_cell: Option<Cell>,
}

impl RowBuilder {
    #[cfg(test)]
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            repeated: 1,
            style_name: None,
            default_cell_style_name: None,
            visibility: TableVisibility::Visible,
            deferred_blank_cells: 0,
            deferred_blank_cell: None,
        }
    }

    pub fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
    }

    pub(crate) fn add_repeated_cells(
        &mut self,
        builder: &CellBuilder,
        text: &str,
        rich_text: Option<&CellTextContent>,
    ) -> Result<()> {
        // Producers pad every row out to the full sheet width with attribute-free
        // `<table:table-cell/>` fillers. Defer those instead of materialising them:
        // an interior run is still expanded when real content follows, but a
        // trailing run is dropped by `build`, which is what makes ordinary
        // spreadsheets fit inside the expansion safety limits at all.
        if builder.is_blank(text, rich_text) {
            self.deferred_blank_cells = self
                .deferred_blank_cells
                .checked_add(builder.repeated)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "table cell repetition overflows address space".to_string(),
                    )
                })?;
            if self.deferred_blank_cell.is_none() {
                self.deferred_blank_cell = Some(builder.build(text, rich_text));
            }
            return Ok(());
        }
        self.flush_deferred_blank_cells()?;
        let expanded = self
            .cells
            .len()
            .checked_add(builder.repeated)
            .ok_or_else(|| {
                Error::InvalidFormat("table cell repetition overflows address space".to_string())
            })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        for _ in 0..builder.repeated {
            self.add_cell(builder.build(text, rich_text));
        }
        Ok(())
    }

    /// Materialise the deferred blank run because real content follows it, so
    /// the column index of that content stays correct.
    fn flush_deferred_blank_cells(&mut self) -> Result<()> {
        let deferred = std::mem::take(&mut self.deferred_blank_cells);
        let Some(template) = self.deferred_blank_cell.take() else {
            return Ok(());
        };
        if deferred == 0 {
            return Ok(());
        }
        let expanded = self.cells.len().checked_add(deferred).ok_or_else(|| {
            Error::InvalidFormat("table cell repetition overflows address space".to_string())
        })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        self.cells.reserve(deferred);
        for _ in 0..deferred {
            self.add_cell(template.clone());
        }
        Ok(())
    }

    pub(crate) fn repeated(&self) -> usize {
        self.repeated
    }

    pub(crate) fn from_parts(
        repeated: usize,
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        visibility: TableVisibility,
    ) -> Self {
        Self {
            cells: Vec::new(),
            repeated,
            style_name,
            default_cell_style_name,
            visibility,
            deferred_blank_cells: 0,
            deferred_blank_cell: None,
        }
    }

    pub(crate) fn build(mut self) -> Row {
        // Row index will be set by the parent SheetBuilder
        // For now, set to 0 and update cells
        for cell in &mut self.cells {
            cell.row = 0; // Will be updated by parent
        }

        Row {
            cells: self.cells,
            index: 0, // Will be set by parent
            style_name: self.style_name,
            default_cell_style_name: self.default_cell_style_name,
            visibility: self.visibility,
        }
    }
}
