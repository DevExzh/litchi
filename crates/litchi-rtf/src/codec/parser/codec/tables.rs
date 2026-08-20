use super::{
    Cow, Destination, DrawingStoryCapture, MAX_LOGICAL_TABLE_ROWS, MAX_LOGICAL_TABLES,
    NestedTableBuilder, Parser, RtfError, RtfResult, resolve_preferred_width, resolve_row_geometry,
};

impl<'a> Parser<'a> {
    /// Start a table if not already started.
    pub(super) fn start_table_if_needed(&mut self) {
        if self.current_table.is_none() {
            self.current_table = Some(super::super::super::table::Table::new());
        }
        if self.current_row.is_none() {
            self.current_row = Some(super::super::super::table::Row::new());
        }
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "the crate nesting-depth constant is defined within the u8 range"
    )]
    pub(super) fn ensure_nested_builder(
        &mut self,
        level: u8,
    ) -> RtfResult<&mut NestedTableBuilder<'a>> {
        if !(2..=crate::MAX_TABLE_NESTING_DEPTH as u8).contains(&level) {
            return Err(RtfError::MalformedDocument(
                "RTF nested-table level is outside 2..=32".to_string(),
            ));
        }
        let index = usize::from(level - 2);
        if self.nested_table_builders.len() < index {
            return Err(RtfError::MalformedDocument(
                "RTF nested-table level transition skips a parent level".to_string(),
            ));
        }
        if self.nested_table_builders.len() == index {
            if level == 2 {
                self.start_table_if_needed();
            }
            self.nested_table_builders
                .push(NestedTableBuilder::new(level));
        }
        self.nested_table_builders.get_mut(index).ok_or_else(|| {
            RtfError::ParserError("RTF nested-table builder state is missing".to_string())
        })
    }

    pub(super) fn append_table_text(&mut self, text: &[u8], raw_level: u8) -> RtfResult<()> {
        let level = if raw_level >= 2 { raw_level } else { 1 };
        // Ordinary table text has no revision side effects. Avoid cloning the
        // complete formatting/table state for every text token in that common
        // path; retain the full snapshot only when revision metadata needs it.
        if self.current_state()?.revision_type.is_none() {
            self.drain_nested_to(level)?;
            if level == 1 {
                self.current_cell_text.extend_from_slice(text);
            } else {
                self.ensure_nested_builder(level)?
                    .cell_text
                    .extend_from_slice(text);
            }
            // Keep the established validation order: the cell buffer is
            // updated before the parser validates the transport slice.
            std::str::from_utf8(text).map_err(|_err| {
                RtfError::MalformedDocument("invalid UTF-8 in table revision".to_string())
            })?;
            return Ok(());
        }

        let state = self.prepare_revision_event()?;
        self.drain_nested_to(level)?;
        let start = if level == 1 {
            self.current_cell_text.len()
        } else {
            self.ensure_nested_builder(level)?.cell_text.len()
        };
        if state.revision_type == Some(super::super::super::annotation::RevisionType::Deletion) {
            let decoded = std::str::from_utf8(text).map_err(|_err| {
                RtfError::MalformedDocument("invalid UTF-8 in table revision".to_string())
            })?;
            return self.append_revision_text(&state, decoded, start, start);
        }
        if level == 1 {
            self.current_cell_text.extend_from_slice(text);
        } else {
            self.ensure_nested_builder(level)?
                .cell_text
                .extend_from_slice(text);
        }
        let end = start.checked_add(text.len()).ok_or_else(|| {
            RtfError::MalformedDocument("RTF table-cell text length overflow".to_string())
        })?;
        let decoded = std::str::from_utf8(text).map_err(|_err| {
            RtfError::MalformedDocument("invalid UTF-8 in table revision".to_string())
        })?;
        self.append_revision_text(&state, decoded, start, end)
    }

    pub(super) fn drain_nested_to(&mut self, parent_level: u8) -> RtfResult<()> {
        while self
            .nested_table_builders
            .last()
            .is_some_and(|builder| builder.level > parent_level)
        {
            let builder = self.nested_table_builders.pop().ok_or_else(|| {
                RtfError::ParserError("RTF nested-table builder state is missing".to_string())
            })?;
            if !builder.cell_text.is_empty()
                || !builder.cell_nested.is_empty()
                || !builder.cell_drawings.drawing_order.is_empty()
                || !builder.cell_story_events.is_empty()
                || builder.row.cell_count() > 0
            {
                return Err(RtfError::MalformedDocument(
                    "RTF nested-table level ended before nestcell/nestrow".to_string(),
                ));
            }
            if builder.table.row_count() == 0 {
                return Err(RtfError::MalformedDocument(
                    "RTF nested table has no completed rows".to_string(),
                ));
            }
            builder
                .table
                .validate_merges()
                .map_err(RtfError::MalformedDocument)?;
            if self.logical_table_count >= MAX_LOGICAL_TABLES {
                return Err(RtfError::MalformedDocument(
                    "RTF document exceeds 4096 logical tables".to_string(),
                ));
            }
            self.logical_table_count += 1;
            let entry = crate::CellNestedTable {
                text_offset: if parent_level == 1 {
                    self.current_cell_text.len()
                } else {
                    self.nested_table_builders
                        .last()
                        .map_or(0, |parent| parent.cell_text.len())
                },
                table: builder.table,
            };
            if parent_level == 1 {
                self.current_cell_story_events
                    .push(crate::CellStoryEvent::NestedTable(
                        self.current_cell_nested.len(),
                    ));
                self.current_cell_nested.push(entry);
            } else {
                let parent = self.nested_table_builders.last_mut().ok_or_else(|| {
                    RtfError::MalformedDocument("RTF nested table lacks a parent table".to_string())
                })?;
                if parent.level != parent_level {
                    return Err(RtfError::MalformedDocument(
                        "RTF nested-table parent level mismatch".to_string(),
                    ));
                }
                parent
                    .cell_story_events
                    .push(crate::CellStoryEvent::NestedTable(parent.cell_nested.len()));
                parent.cell_nested.push(entry);
            }
        }
        Ok(())
    }

    pub(super) fn finalize_nested_cell(&mut self, level: u8) -> RtfResult<()> {
        self.drain_nested_to(level)?;
        self.close_revision_at_cell_boundary(level)?;
        let arena = self.arena;
        let builder = self.ensure_nested_builder(level)?;
        if builder.row.cell_count() >= crate::MAX_TABLE_CELLS_PER_ROW {
            return Err(RtfError::MalformedDocument(
                "RTF table row exceeds 4096 cells".to_string(),
            ));
        }
        let text = std::str::from_utf8(&builder.cell_text).map_err(|_err| {
            RtfError::MalformedDocument("invalid UTF-8 in nested table cell".to_string())
        })?;
        let mut cell = crate::Cell::new(Cow::Borrowed(arena.alloc_str(text)));
        cell.nested_tables_mut().append(&mut builder.cell_nested);
        let drawings = std::mem::take(&mut builder.cell_drawings);
        let events = std::mem::take(&mut builder.cell_story_events);
        cell.set_story_content(
            drawings.shapes,
            drawings.shape_groups,
            drawings.drawing_order,
            events,
        )?;
        builder.row.add_cell(cell);
        builder.cell_text.clear();
        Ok(())
    }

    pub(super) fn finalize_nested_row(&mut self, level: u8) -> RtfResult<()> {
        self.drain_nested_to(level)?;
        let state = self.current_state()?.clone();
        let geometry = resolve_row_geometry(&state)?;
        let cell_defaults = crate::TableRowCellDefaults {
            borders: state.table_default_borders.clone(),
            padding: state.table_default_padding.clone(),
            spacing: state.table_default_spacing.clone(),
            preferred_cell_width: resolve_preferred_width(
                state.table_default_width_unit,
                state.table_default_width_value,
            )?,
        };
        let builder = self.ensure_nested_builder(level)?;
        if !builder.cell_text.is_empty()
            || !builder.cell_nested.is_empty()
            || !builder.cell_drawings.drawing_order.is_empty()
            || !builder.cell_story_events.is_empty()
        {
            return Err(RtfError::MalformedDocument(
                "RTF nestrow encountered an unterminated nested cell".to_string(),
            ));
        }
        if builder.row.cell_count() == 0 {
            return Err(RtfError::MalformedDocument(
                "RTF nestrow has no nestcell".to_string(),
            ));
        }
        if !state.cell_boundaries.is_empty()
            && state.cell_boundaries.len() != builder.row.cell_count()
        {
            return Err(RtfError::MalformedDocument(
                "RTF nested row cell boundaries do not match nestcell count".to_string(),
            ));
        }
        for (index, cell) in builder.row.cells_mut().iter_mut().enumerate() {
            if let Some((padding, spacing)) = state.cell_distances.get(index) {
                cell.set_padding(padding.clone());
                cell.set_spacing(spacing.clone());
            }
            if let Some(layout) = state.cell_layouts.get(index) {
                cell.set_layout(*layout);
            }
            if let Some(merge) = state.cell_merges.get(index) {
                cell.set_merge(*merge);
            }
            if let Some(revision) = state.cell_revisions.get(index) {
                cell.set_revision(*revision);
            }
            cell.set_right_boundary(state.cell_boundaries.get(index).copied());
            cell.set_preferred_width(state.cell_widths.get(index).copied().flatten());
            if let Some((borders, shading)) = state.cell_decorations.get(index) {
                cell.set_borders(borders.clone());
                cell.set_shading(*shading);
            }
        }
        builder.row.set_table_style(state.table_style);
        builder.row.set_table_rsid(state.table_rsid);
        builder.row.set_direction(state.table_row_direction);
        builder.row.set_layout(state.table_row_layout);
        builder.row.set_padding(state.table_row_padding.clone());
        builder.row.set_spacing(state.table_row_spacing.clone());
        builder.row.set_cell_defaults(cell_defaults);
        builder
            .row
            .set_positioning(state.table_row_positioning.clone());
        builder.row.set_borders(state.table_row_borders.clone());
        builder.row.set_shading(state.table_row_shading);
        builder.row.set_geometry(geometry);
        builder
            .row
            .set_autoformat_flags(state.table_autoformat_flags);
        builder.row.set_banding(state.table_row_banding);
        builder.row.set_revision(state.table_row_revision);
        if builder.table.row_count() >= MAX_LOGICAL_TABLE_ROWS {
            return Err(RtfError::MalformedDocument(
                "RTF logical table exceeds 65536 rows".to_string(),
            ));
        }
        if builder
            .table
            .rows()
            .first()
            .is_some_and(|first| first.positioning() != builder.row.positioning())
        {
            return Err(RtfError::MalformedDocument("RTF positioned-table properties must be identical for all rows in one logical table".to_string()));
        }
        builder.table.add_row(std::mem::take(&mut builder.row));
        Ok(())
    }

    /// Finalize the current cell and add it to the current row.
    pub(super) fn finalize_cell(&mut self, explicit: bool) -> RtfResult<()> {
        self.drain_nested_to(1)?;
        self.close_revision_at_cell_boundary(1)?;
        if explicit
            || !self.current_cell_text.is_empty()
            || !self.current_cell_nested.is_empty()
            || !self.current_cell_drawings.drawing_order.is_empty()
            || !self.current_cell_story_events.is_empty()
        {
            if self
                .current_row
                .as_ref()
                .map_or(0, crate::content::table::Row::cell_count)
                >= crate::MAX_TABLE_CELLS_PER_ROW
            {
                return Err(RtfError::MalformedDocument(
                    "RTF table row exceeds 4096 cells".to_string(),
                ));
            }
            // Convert cell text to string
            if let Ok(text_str) = std::str::from_utf8(&self.current_cell_text) {
                let allocated = self.arena.alloc_str(text_str);
                let index = self
                    .current_row
                    .as_ref()
                    .map_or(0, crate::content::table::Row::cell_count);
                let (padding, spacing) = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_distances.get(index))
                    .cloned()
                    .unwrap_or_default();
                let layout = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_layouts.get(index))
                    .copied()
                    .unwrap_or_default();
                let merge = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_merges.get(index))
                    .copied()
                    .unwrap_or_default();
                let boundary = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_boundaries.get(index))
                    .copied();
                let width = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_widths.get(index))
                    .copied()
                    .flatten();
                let (borders, shading) = self
                    .current_state()
                    .ok()
                    .and_then(|state| state.cell_decorations.get(index))
                    .cloned()
                    .unwrap_or_default();
                let mut cell = super::super::super::table::Cell::with_distances(
                    Cow::Borrowed(allocated),
                    padding,
                    spacing,
                );
                cell.set_layout(layout);
                cell.set_merge(merge);
                cell.set_right_boundary(boundary);
                cell.set_preferred_width(width);
                cell.set_borders(borders);
                cell.set_shading(shading);
                cell.nested_tables_mut()
                    .append(&mut self.current_cell_nested);
                let drawings = std::mem::take(&mut self.current_cell_drawings);
                let events = std::mem::take(&mut self.current_cell_story_events);
                cell.set_story_content(
                    drawings.shapes,
                    drawings.shape_groups,
                    drawings.drawing_order,
                    events,
                )?;

                // Add cell to current row
                if let Some(row) = &mut self.current_row {
                    row.add_cell(cell);
                }
            }
        }
        self.current_cell_text.clear();
        self.current_cell_drawings = DrawingStoryCapture::default();
        self.current_cell_story_events.clear();
        Ok(())
    }

    /// Finalize the current row and add it to the current table.
    pub(super) fn finalize_row(&mut self) -> RtfResult<()> {
        // Finalize any pending cell
        self.finalize_cell(false)?;

        // Add row to table
        if let (Some(table), Some(row)) = (&mut self.current_table, self.current_row.take())
            && row.cell_count() > 0
        {
            if table.row_count() >= MAX_LOGICAL_TABLE_ROWS {
                return Err(RtfError::MalformedDocument(
                    "RTF logical table exceeds 65536 rows".to_string(),
                ));
            }
            if table
                .rows()
                .first()
                .is_some_and(|first| first.positioning() != row.positioning())
            {
                return Err(RtfError::MalformedDocument("RTF positioned-table properties must be identical for all rows in one logical table".to_string()));
            }
            table.add_row(row);
        }

        // Start a new row for next cells
        self.current_row = Some(super::super::super::table::Row::new());
        Ok(())
    }

    /// Finalize the current table and add it to the tables list.
    pub(super) fn finalize_table(&mut self) -> RtfResult<()> {
        self.drain_nested_to(1)?;
        // Finalize any pending row
        if self.current_row.is_some() {
            self.finalize_row()?;
        }

        // Add table to tables list
        if let Some(table) = self.current_table.take()
            && table.row_count() > 0
        {
            table
                .validate_merges()
                .map_err(RtfError::MalformedDocument)?;
            if self.logical_table_count >= MAX_LOGICAL_TABLES {
                return Err(RtfError::MalformedDocument(
                    "RTF document exceeds 4096 logical tables".to_string(),
                ));
            }
            self.logical_table_count += 1;
            self.tables.push(table);
        }
        Ok(())
    }

    pub(super) fn finalize_table_before_non_table_body_content(
        &mut self,
        meaningful: bool,
    ) -> RtfResult<bool> {
        if meaningful
            && self
                .current_table
                .as_ref()
                .is_some_and(|table| table.row_count() > 0)
            && self.current_state().is_ok_and(|state| {
                state.destination == Destination::DocumentBody && !state.in_table
            })
        {
            self.finalize_table()?;
            return Ok(true);
        }
        Ok(false)
    }
}
