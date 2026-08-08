//! RTF table output.

use super::super::*;

impl<W: Write> RtfWriter<W> {
    /// Write a table
    pub(in super::super) fn write_table(
        &mut self,
        table: &Table,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        if let Some(first) = table.rows().first()
            && table
                .rows()
                .iter()
                .skip(1)
                .any(|row| row.positioning() != first.positioning())
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF positioned-table properties must be identical for all rows in one logical table",
            ));
        }
        for row in table.rows() {
            self.write_table_row(
                row,
                table.direction(),
                fields,
                navigation_entries,
                revisions,
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn validate_table_tree(
        table: &Table,
        depth: usize,
        count: &mut usize,
    ) -> io::Result<()> {
        if depth > crate::MAX_TABLE_NESTING_DEPTH {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table nesting exceeds 32 levels",
            ));
        }
        *count = count.checked_add(1).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF logical table count overflow",
            )
        })?;
        if *count > 4096 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF document exceeds 4096 logical tables",
            ));
        }
        if table.row_count() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF logical table exceeds 65536 rows",
            ));
        }
        table
            .validate_merges()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        if let Some(first) = table.rows().first()
            && table
                .rows()
                .iter()
                .skip(1)
                .any(|row| row.positioning() != first.positioning())
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF positioned-table properties must be identical for all rows in one logical table",
            ));
        }
        for row in table.rows() {
            row.geometry()
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if row.cell_count() > crate::MAX_TABLE_CELLS_PER_ROW {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF table row exceeds 4096 cells",
                ));
            }
            for cell in row.cells() {
                cell.validate_drawings().map_err(|error| {
                    io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                })?;
                if let Some(width) = cell.preferred_width() {
                    width.validate().map_err(|error| {
                        io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                    })?;
                }
                let mut previous = 0;
                for nested in cell.nested_tables() {
                    if nested.text_offset < previous
                        || nested.text_offset > cell.text().len()
                        || !cell.text().is_char_boundary(nested.text_offset)
                    {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "invalid nested-table text insertion offset",
                        ));
                    }
                    previous = nested.text_offset;
                    Self::validate_table_tree(&nested.table, depth + 1, count)?;
                }
            }
        }
        Ok(())
    }

    /// Write a table row
    pub(in super::super) fn write_table_row(
        &mut self,
        row: &Row,
        table_direction: Option<TextDirection>,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        // Row defaults
        self.write_control_word("trowd", None)?;
        if let Some(table_style) = row.table_style() {
            self.write_control_word("ts", Some(i32::from(table_style)))?;
        }
        if let Some(table_rsid) = row.table_rsid() {
            self.write_control_word("tblrsid", Some(table_rsid as i32))?;
        }
        self.write_revision_metadata("trauth", "trdate", row.revision())?;

        if let Some(direction) = table_direction {
            self.write_control_word(
                "taprtl",
                (direction == TextDirection::LeftToRight).then_some(0),
            )?;
        }
        self.write_table_row_banding(row)?;
        self.write_table_row_layout(row)?;
        self.write_table_row_geometry(row.geometry())?;
        self.write_table_row_borders(row.borders())?;
        self.write_table_shading("tr", row.shading())?;
        self.write_table_distances("trpadd", "trpaddf", row.padding())?;
        self.write_table_distances("trspd", "trspdf", row.spacing())?;
        self.write_table_row_cell_defaults(row.cell_defaults())?;
        self.write_floating_table_position(row.positioning())?;

        // Cell boundaries
        let cell_width = 2880; // Default cell width (2 inches)
        for (i, cell) in row.cells().iter().enumerate() {
            self.write_table_preferred_width("clftsWidth", "clwWidth", cell.preferred_width())?;
            self.write_table_cell_merge(cell)?;
            self.write_table_cell_revision(cell)?;
            self.write_table_cell_layout(cell)?;
            self.write_table_cell_borders(cell.borders())?;
            self.write_table_shading("cl", cell.shading())?;
            self.write_table_distances("clpad", "clpadf", cell.padding())?;
            self.write_table_distances("clspd", "clspdf", cell.spacing())?;
            let boundary = cell
                .right_boundary()
                .unwrap_or(cell_width * ((i + 1) as i32));
            self.write_control_word("cellx", Some(boundary))?;
        }

        // Write cells
        for cell in row.cells() {
            self.write_str("{")?;
            self.write_control_word("intbl", None)?;
            self.write_str(" ")?;
            self.write_cell_content(cell, 1, fields, navigation_entries, revisions)?;
            self.write_control_word("cell", None)?;
            self.write_str("}")?;
        }

        // Row end
        self.write_control_word("row", None)?;
        self.write_str("\n")?;

        Ok(())
    }

    pub(in super::super) fn write_cell_content(
        &mut self,
        cell: &crate::Cell<'_>,
        depth: usize,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        let field_owner_depth = u8::try_from(depth).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table nesting depth cannot be represented",
            )
        })?;
        let mut offset = 0usize;
        for event in cell.story_events() {
            let position = match *event {
                crate::CellStoryEvent::NestedTable(index) => {
                    cell.nested_tables()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .text_offset
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    cell.shapes()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    cell.shape_groups()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                crate::CellStoryEvent::Field(field) => field.position,
                crate::CellStoryEvent::PageBreak(page_break) => page_break.position,
                crate::CellStoryEvent::ColumnBreak(column_break) => column_break.position,
                crate::CellStoryEvent::NavigationEntry(reference)
                | crate::CellStoryEvent::RevisionStart(reference)
                | crate::CellStoryEvent::RevisionEnd(reference)
                | crate::CellStoryEvent::RevisionDeletion(reference) => reference.position,
            };
            let fragment = cell.text().get(offset..position).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF table-cell event splits or leaves its story text",
                )
            })?;
            self.write_text(fragment)?;
            match *event {
                crate::CellStoryEvent::NestedTable(index) => {
                    let nested = cell
                        .nested_tables()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    let nested_depth = depth.checked_add(1).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF table nesting depth overflow",
                        )
                    })?;
                    self.write_nested_table(
                        &nested.table,
                        nested_depth,
                        fields,
                        navigation_entries,
                        revisions,
                    )?;
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    let shape = cell
                        .shapes()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    self.write_root_shape(shape)?
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    let group = cell
                        .shape_groups()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    self.write_shape_group(group, true)?
                },
                crate::CellStoryEvent::Field(reference) => {
                    let field = fields
                        .get(reference.field_index)
                        .filter(|field| {
                            field.owner == crate::FieldOwner::TableCell(field_owner_depth)
                                && field.position == reference.position
                                && field.range_end == reference.position
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell story has an invalid generic-field owner or reference",
                            )
                        })?;
                    self.write_field_with_fields(field, fields, 0)?;
                },
                crate::CellStoryEvent::PageBreak(_) => self.write_str("\\page ")?,
                crate::CellStoryEvent::ColumnBreak(_) => self.write_str("\\column ")?,
                crate::CellStoryEvent::NavigationEntry(reference) => {
                    let entry = navigation_entries
                        .get(reference.index)
                        .filter(|entry| entry.position() == reference.position)
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell navigation reference is invalid",
                            )
                        })?;
                    self.write_navigation_entry(entry)?;
                },
                crate::CellStoryEvent::RevisionStart(reference) => {
                    let revision = revisions
                        .get(reference.index)
                        .filter(|revision| {
                            revision.revision_type == RevisionType::Insertion
                                && revision.position == reference.position
                                && cell.text().get(revision.position..revision.range_end)
                                    == Some(revision.content.as_ref())
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell insertion revision reference is invalid",
                            )
                        })?;
                    self.write_revision_start(revision)?;
                },
                crate::CellStoryEvent::RevisionEnd(reference) => {
                    revisions
                        .get(reference.index)
                        .filter(|revision| {
                            revision.revision_type == RevisionType::Insertion
                                && revision.range_end == reference.position
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell revision end reference is invalid",
                            )
                        })?;
                    self.write_str("}")?;
                },
                crate::CellStoryEvent::RevisionDeletion(reference) => {
                    let revision = revisions
                        .get(reference.index)
                        .filter(|revision| {
                            revision.revision_type == RevisionType::Deletion
                                && revision.position == reference.position
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell deletion revision reference is invalid",
                            )
                        })?;
                    self.write_revision(revision)?;
                },
            }
            offset = position;
        }
        let remainder = cell.text().get(offset..).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table-cell event leaves its story text",
            )
        })?;
        self.write_text(remainder)
    }

    pub(in super::super) fn write_nested_table(
        &mut self,
        table: &Table,
        depth: usize,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        let depth_value = i32::try_from(depth).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table nesting depth cannot be represented",
            )
        })?;
        for row in table.rows() {
            for cell in row.cells() {
                self.write_str("{")?;
                self.write_control_word("intbl", None)?;
                self.write_control_word("itap", Some(depth_value))?;
                self.write_str(" ")?;
                self.write_cell_content(cell, depth, fields, navigation_entries, revisions)?;
                self.write_control_word("nestcell", None)?;
                self.write_str("}")?;
            }
            self.write_str("{\\*")?;
            self.write_control_word("nesttableprops", None)?;
            self.write_control_word("itap", Some(depth_value))?;
            self.write_control_word("trowd", None)?;
            if let Some(table_style) = row.table_style() {
                self.write_control_word("ts", Some(i32::from(table_style)))?;
            }
            if let Some(table_rsid) = row.table_rsid() {
                self.write_control_word("tblrsid", Some(table_rsid as i32))?;
            }
            self.write_revision_metadata("trauth", "trdate", row.revision())?;
            if let Some(direction) = table.direction() {
                self.write_control_word(
                    "taprtl",
                    (direction == TextDirection::LeftToRight).then_some(0),
                )?;
            }
            self.write_table_row_banding(row)?;
            self.write_table_row_layout(row)?;
            self.write_table_row_geometry(row.geometry())?;
            self.write_table_row_borders(row.borders())?;
            self.write_table_shading("tr", row.shading())?;
            self.write_table_distances("trpadd", "trpaddf", row.padding())?;
            self.write_table_distances("trspd", "trspdf", row.spacing())?;
            self.write_table_row_cell_defaults(row.cell_defaults())?;
            self.write_floating_table_position(row.positioning())?;
            for (index, cell) in row.cells().iter().enumerate() {
                self.write_table_preferred_width("clftsWidth", "clwWidth", cell.preferred_width())?;
                self.write_table_cell_merge(cell)?;
                self.write_table_cell_revision(cell)?;
                self.write_table_cell_layout(cell)?;
                self.write_table_cell_borders(cell.borders())?;
                self.write_table_shading("cl", cell.shading())?;
                self.write_table_distances("clpad", "clpadf", cell.padding())?;
                self.write_table_distances("clspd", "clspdf", cell.spacing())?;
                self.write_control_word(
                    "cellx",
                    Some(cell.right_boundary().unwrap_or(2880 * ((index + 1) as i32))),
                )?;
            }
            self.write_control_word("nestrow", None)?;
            self.write_str("}")?;
            self.write_str("{")?;
            self.write_control_word("nonesttables", None)?;
            self.write_control_word("par", None)?;
            self.write_str("}")?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_preferred_width(
        &mut self,
        unit_control: &str,
        value_control: &str,
        width: Option<crate::TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                crate::TablePreferredWidthUnit::Null => 0,
                crate::TablePreferredWidthUnit::Auto => 1,
                crate::TablePreferredWidthUnit::Percent => 2,
                crate::TablePreferredWidthUnit::Twips => 3,
            }),
        )?;
        if let Some(value) = width.value() {
            self.write_control_word(value_control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_invisible_width(
        &mut self,
        unit_control: &str,
        value_control: &str,
        width: Option<crate::TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                crate::TablePreferredWidthUnit::Null => 0,
                crate::TablePreferredWidthUnit::Auto => 1,
                crate::TablePreferredWidthUnit::Percent => 2,
                crate::TablePreferredWidthUnit::Twips => 3,
            }),
        )?;
        if let Some(value) = width.value().filter(|value| *value != 0) {
            self.write_control_word(value_control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_row_geometry(
        &mut self,
        geometry: crate::TableRowGeometry,
    ) -> io::Result<()> {
        geometry
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(value) = geometry.half_gap_twips() {
            self.write_control_word("trgaph", Some(i32::from(value)))?;
        }
        if let Some(value) = geometry.left_edge_twips() {
            self.write_control_word("trleft", Some(value))?;
        }
        match geometry.height() {
            crate::TableRowHeight::Automatic => {},
            crate::TableRowHeight::Minimum(value) => {
                self.write_control_word("trrh", Some(i32::from(value)))?
            },
            crate::TableRowHeight::Exact(value) => {
                self.write_control_word("trrh", Some(-i32::from(value)))?
            },
        }
        self.write_table_preferred_width("trftsWidth", "trwWidth", geometry.preferred_width())?;
        self.write_table_invisible_width(
            "trftsWidthB",
            "trwWidthB",
            geometry.leading_invisible_width(),
        )?;
        self.write_table_invisible_width(
            "trftsWidthA",
            "trwWidthA",
            geometry.trailing_invisible_width(),
        )?;
        if geometry.auto_fit() {
            self.write_control_word("trautofit", Some(1))?;
        }
        if let Some(indent) = geometry.indent() {
            self.write_control_word("tblind", Some(indent.value()))?;
            self.write_control_word(
                "tblindtype",
                Some(match indent.unit() {
                    crate::TableIndentUnit::Auto => 0,
                    crate::TableIndentUnit::Twips => 1,
                    crate::TableIndentUnit::Nil => 2,
                    crate::TableIndentUnit::Percent => 3,
                }),
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_row_layout(&mut self, row: &Row<'_>) -> io::Result<()> {
        if let Some(alignment) = row.layout().alignment {
            self.write_control_word(
                match alignment {
                    crate::TableRowAlignment::Left => "trql",
                    crate::TableRowAlignment::Center => "trqc",
                    crate::TableRowAlignment::Right => "trqr",
                },
                None,
            )?;
        }
        if let Some(direction) = row.direction() {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrrow",
                    TextDirection::RightToLeft => "rtlrow",
                },
                None,
            )?;
        }
        if row.layout().header {
            self.write_control_word("trhdr", None)?;
        }
        if row.layout().keep_together {
            self.write_control_word("trkeep", None)?;
        }
        if row.layout().keep_with_following {
            self.write_control_word("trkeepfollow", None)?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_row_banding(&mut self, row: &Row<'_>) -> io::Result<()> {
        let banding = row.banding();
        if let Some(value) = banding.row_index {
            self.write_control_word("irow", Some(i32::from(value)))?;
        }
        if let Some(value) = banding.band_index {
            self.write_control_word(
                "irowband",
                Some(match value {
                    crate::TableRowBandIndex::Header => -1,
                    crate::TableRowBandIndex::Row(value) => i32::from(value),
                }),
            )?;
        }
        let flags = row.autoformat_flags();
        for (flag, word) in [
            (crate::TableAutoformatFlag::Border, "tbllkborder"),
            (crate::TableAutoformatFlag::Shading, "tbllkshading"),
            (crate::TableAutoformatFlag::Font, "tbllkfont"),
            (crate::TableAutoformatFlag::Color, "tbllkcolor"),
            (crate::TableAutoformatFlag::BestFit, "tbllkbestfit"),
            (crate::TableAutoformatFlag::HeaderRows, "tbllkhdrrows"),
            (crate::TableAutoformatFlag::LastRow, "tbllklastrow"),
            (crate::TableAutoformatFlag::HeaderColumns, "tbllkhdrcols"),
            (crate::TableAutoformatFlag::LastColumn, "tbllklastcol"),
            (crate::TableAutoformatFlag::NoRowBanding, "tbllknorowband"),
            (
                crate::TableAutoformatFlag::NoColumnBanding,
                "tbllknocolband",
            ),
        ] {
            if flags.contains(flag) {
                self.write_control_word(word, None)?;
            }
        }
        if banding.last_row {
            self.write_control_word("lastrow", None)?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_cell_layout(
        &mut self,
        cell: &crate::Cell<'_>,
    ) -> io::Result<()> {
        let layout = cell.layout();
        if let Some(alignment) = layout.vertical_alignment {
            self.write_control_word(
                match alignment {
                    crate::TableCellVerticalAlignment::Top => "clvertalt",
                    crate::TableCellVerticalAlignment::Center => "clvertalc",
                    crate::TableCellVerticalAlignment::Bottom => "clvertalb",
                },
                None,
            )?;
        }
        if let Some(flow) = layout.text_flow {
            self.write_control_word(
                match flow {
                    crate::TableCellTextFlow::LeftToRightTopToBottom => "cltxlrtb",
                    crate::TableCellTextFlow::RightToLeftTopToBottom => "cltxtbrl",
                    crate::TableCellTextFlow::LeftToRightBottomToTop => "cltxbtlr",
                    crate::TableCellTextFlow::LeftToRightTopToBottomVertical => "cltxlrtbv",
                    crate::TableCellTextFlow::TopToBottomRightToLeftVertical => "cltxtbrlv",
                },
                None,
            )?;
        }
        if layout.fit_text {
            self.write_control_word("clFitText", None)?;
        }
        if layout.no_wrap {
            self.write_control_word("clNoWrap", None)?;
        }
        if layout.hide_mark {
            self.write_control_word("clhidemark", None)?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_cell_merge(
        &mut self,
        cell: &crate::Cell<'_>,
    ) -> io::Result<()> {
        let merge = cell.merge();
        if let Some(role) = merge.horizontal {
            self.write_control_word(
                match role {
                    crate::TableCellMergeRole::First => "clmgf",
                    crate::TableCellMergeRole::Continuation => "clmrg",
                },
                None,
            )?;
        }
        if let Some(role) = merge.vertical {
            self.write_control_word(
                match role {
                    crate::TableCellMergeRole::First => "clvmgf",
                    crate::TableCellMergeRole::Continuation => "clvmrg",
                },
                None,
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_cell_revision(
        &mut self,
        cell: &crate::Cell<'_>,
    ) -> io::Result<()> {
        if let Some(revision) = cell.revision() {
            self.write_control_word(revision.kind.control_word(), None)?;
            self.write_revision_metadata(
                revision.kind.author_control_word(),
                revision.kind.date_control_word(),
                revision.metadata,
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_border(
        &mut self,
        selector: &str,
        border: &Border,
    ) -> io::Result<()> {
        border
            .validate_table()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(selector, None)?;
        self.write_control_word(border.style.control_word(), None)?;
        if border.style == BorderStyle::None {
            return Ok(());
        }
        self.write_control_word("brdrw", Some(border.width))?;
        self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
        self.write_control_word("brsp", Some(border.space))?;
        if border.shadow {
            self.write_control_word("brdrsh", None)?;
        }
        if border.frame {
            self.write_control_word("brdrframe", None)?;
        }
        Ok(())
    }
    pub(in super::super) fn write_table_row_borders(
        &mut self,
        borders: &crate::TableRowBorders,
    ) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (selector, border) in [
            ("trbrdrt", borders.top),
            ("trbrdrl", borders.left),
            ("trbrdrb", borders.bottom),
            ("trbrdrr", borders.right),
            ("trbrdrh", borders.horizontal),
            ("trbrdrv", borders.vertical),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        Ok(())
    }
    pub(in super::super) fn write_table_cell_borders(
        &mut self,
        borders: &crate::TableCellBorders,
    ) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (selector, border) in [
            ("clbrdrt", borders.top),
            ("clbrdrl", borders.left),
            ("clbrdrb", borders.bottom),
            ("clbrdrr", borders.right),
            ("cldglu", borders.upper_left_to_lower_right),
            ("cldgll", borders.upper_right_to_lower_left),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        Ok(())
    }
    pub(in super::super) fn write_table_row_cell_defaults(
        &mut self,
        defaults: &crate::TableRowCellDefaults,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let borders = &defaults.borders;
        for (selector, border) in [
            ("tsbrdrt", borders.top),
            ("tsbrdrl", borders.left),
            ("tsbrdrb", borders.bottom),
            ("tsbrdrr", borders.right),
            ("tsbrdrh", borders.horizontal_inside),
            ("tsbrdrv", borders.vertical_inside),
            ("tsbrdrdgl", borders.diagonal_upper_left_to_lower_right),
            ("tsbrdrdg", borders.diagonal_upper_right_to_lower_left),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        self.write_table_distances("tscellpadd", "tscellpaddf", &defaults.padding)?;
        self.write_table_distances("tscellspc", "tscellspcf", &defaults.spacing)?;
        self.write_table_preferred_width(
            "tscellwidthfts",
            "tscellwidth",
            defaults.preferred_cell_width,
        )?;
        Ok(())
    }

    pub(in super::super) fn write_table_shading(
        &mut self,
        prefix: &str,
        shading: crate::TableShading,
    ) -> io::Result<()> {
        shading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(index) = shading.pattern_index {
            if prefix != "tr" {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "trpat is only valid for row shading",
                ));
            }
            self.write_control_word("trpat", Some(i32::from(index)))?;
        }
        if let Some(pattern) = shading.pattern {
            let suffix = match pattern {
                crate::ShadingPattern::Horizontal => "bghoriz",
                crate::ShadingPattern::Vertical => "bgvert",
                crate::ShadingPattern::ForwardDiagonal => "bgfdiag",
                crate::ShadingPattern::BackwardDiagonal => "bgbdiag",
                crate::ShadingPattern::Cross => "bgcross",
                crate::ShadingPattern::DiagonalCross => "bgdcross",
                crate::ShadingPattern::DarkHorizontal => "bgdkhor",
                crate::ShadingPattern::DarkVertical => "bgdkvert",
                crate::ShadingPattern::DarkForwardDiagonal => "bgdkfdiag",
                crate::ShadingPattern::DarkBackwardDiagonal => "bgdkbdiag",
                crate::ShadingPattern::DarkCross => "bgdkcross",
                crate::ShadingPattern::DarkDiagonalCross => "bgdkdcross",
                _ => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "invalid explicit RTF table shading pattern",
                    ));
                },
            };
            self.write_control_word(&format!("{prefix}{suffix}"), None)?;
        }
        if let Some(color) = shading.foreground_color {
            self.write_control_word(&format!("{prefix}cfpat"), Some(i32::from(color)))?;
        }
        if let Some(color) = shading.background_color {
            self.write_control_word(&format!("{prefix}cbpat"), Some(i32::from(color)))?;
        }
        if let Some(amount) = shading.amount {
            self.write_control_word(&format!("{prefix}shdng"), Some(i32::from(amount)))?;
        }
        if let Some(amount) = shading.raw_amount {
            if prefix != "cl" {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "raw table shading is only valid for cells",
                ));
            }
            self.write_control_word("clshdngraw", Some(i32::from(amount)))?;
        }
        if shading.raw_nil {
            if prefix != "cl" {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "raw-nil table shading is only valid for cells",
                ));
            }
            self.write_control_word("clshdrawnil", None)?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_distances(
        &mut self,
        value_prefix: &str,
        unit_prefix: &str,
        distances: &TableEdgeDistances,
    ) -> io::Result<()> {
        distances
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (suffix, side) in [
            ("l", distances.left),
            ("r", distances.right),
            ("t", distances.top),
            ("b", distances.bottom),
        ] {
            if let Some(value) = side.value {
                self.write_control_word(
                    &format!("{value_prefix}{suffix}"),
                    Some(i32::from(value)),
                )?;
            }
            if let Some(unit) = side.unit {
                self.write_control_word(
                    &format!("{unit_prefix}{suffix}"),
                    Some(match unit {
                        TableDistanceUnit::Null => 0,
                        TableDistanceUnit::Twips => 3,
                    }),
                )?;
            }
        }
        Ok(())
    }

    pub(in super::super) fn write_floating_table_position(
        &mut self,
        position: &crate::FloatingTablePosition,
    ) -> io::Result<()> {
        position
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(reference) = position.horizontal_reference {
            self.write_control_word(
                match reference {
                    crate::TableHorizontalReference::Column => "tphcol",
                    crate::TableHorizontalReference::Margin => "tphmrg",
                    crate::TableHorizontalReference::Page => "tphpg",
                },
                None,
            )?
        }
        if let Some(value) = position.horizontal_position {
            let (word, param) = match value {
                crate::TableHorizontalPosition::Offset(value) => ("tposx", Some(value)),
                crate::TableHorizontalPosition::NegativeOffset(value) => ("tposnegx", Some(value)),
                crate::TableHorizontalPosition::Center => ("tposxc", None),
                crate::TableHorizontalPosition::Inside => ("tposxi", None),
                crate::TableHorizontalPosition::Left => ("tposxl", None),
                crate::TableHorizontalPosition::Outside => ("tposxo", None),
                crate::TableHorizontalPosition::Right => ("tposxr", None),
            };
            self.write_control_word(word, param)?
        }
        if let Some(reference) = position.vertical_reference {
            self.write_control_word(
                match reference {
                    crate::TableVerticalReference::Margin => "tpvmrg",
                    crate::TableVerticalReference::Paragraph => "tpvpara",
                    crate::TableVerticalReference::Page => "tpvpg",
                },
                None,
            )?
        }
        if let Some(value) = position.vertical_position {
            let (word, param) = match value {
                crate::TableVerticalPosition::Offset(value) => ("tposy", Some(value)),
                crate::TableVerticalPosition::NegativeOffset(value) => ("tposnegy", Some(value)),
                crate::TableVerticalPosition::Bottom => ("tposyb", None),
                crate::TableVerticalPosition::Center => ("tposyc", None),
                crate::TableVerticalPosition::Inline => ("tposyil", None),
                crate::TableVerticalPosition::Inside => ("tposyin", None),
                crate::TableVerticalPosition::Outside => ("tposyout", None),
                crate::TableVerticalPosition::Top => ("tposyt", None),
            };
            self.write_control_word(word, param)?
        }
        for (word, value) in [
            ("tdfrmtxtLeft", position.wrap_distances.left),
            ("tdfrmtxtRight", position.wrap_distances.right),
            ("tdfrmtxtTop", position.wrap_distances.top),
            ("tdfrmtxtBottom", position.wrap_distances.bottom),
        ] {
            if let Some(value) = value {
                self.write_control_word(word, Some(i32::from(value)))?
            }
        }
        if position.no_overlap {
            self.write_control_word("tabsnoovrlp", None)?
        }
        Ok(())
    }
}
