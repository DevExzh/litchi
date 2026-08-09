//! RTF table output.

#![allow(
    clippy::shadow_reuse,
    reason = "serialization helpers deliberately rebind a working value as the output is assembled"
)]
use super::super::{
    Border, BorderStyle, Cell, CellStoryEvent, Field, FieldOwner, FloatingTablePosition,
    MAX_TABLE_CELLS_PER_ROW, MAX_TABLE_NESTING_DEPTH, NavigationEntry, Revision, RevisionType, Row,
    RtfWriter, ShadingPattern, StoryDrawing, Table, TableAutoformatFlag, TableCellBorders,
    TableCellMergeRole, TableCellTextFlow, TableCellVerticalAlignment, TableDistanceUnit,
    TableEdgeDistances, TableHorizontalPosition, TableHorizontalReference, TableIndentUnit,
    TablePreferredWidth, TablePreferredWidthUnit, TableRowAlignment, TableRowBandIndex,
    TableRowBorders, TableRowCellDefaults, TableRowGeometry, TableRowHeight, TableShading,
    TableVerticalPosition, TableVerticalReference, TextDirection, Write, invalid_story_reference,
    io,
};

impl<W: Write> RtfWriter<W> {
    /// Write a table
    pub(in super::super) fn write_table(
        &mut self,
        table: &Table<'_>,
        fields: &[Field<'_>],
        navigation_entries: &[NavigationEntry<'_>],
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
        table: &Table<'_>,
        depth: usize,
        count: &mut usize,
    ) -> io::Result<()> {
        if depth > MAX_TABLE_NESTING_DEPTH {
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
            if row.cell_count() > MAX_TABLE_CELLS_PER_ROW {
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
        row: &Row<'_>,
        table_direction: Option<TextDirection>,
        fields: &[Field<'_>],
        navigation_entries: &[NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        // Row defaults
        self.write_control_word("trowd", None)?;
        if let Some(table_style) = row.table_style() {
            self.write_control_word("ts", Some(i32::from(table_style)))?;
        }
        if let Some(table_rsid) = row.table_rsid() {
            self.write_control_word("tblrsid", Some(table_rsid.cast_signed()))?;
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
            let cell_index = i32::try_from(i + 1).map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF table cell index exceeds the i32 range",
                )
            })?;
            let boundary = cell.right_boundary().unwrap_or(cell_width * cell_index);
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
        cell: &Cell<'_>,
        depth: usize,
        fields: &[Field<'_>],
        navigation_entries: &[NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        let field_owner_depth = u8::try_from(depth).map_err(|_err| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table nesting depth cannot be represented",
            )
        })?;
        let mut offset = 0usize;
        for event in cell.story_events() {
            let position = match *event {
                CellStoryEvent::NestedTable(index) => {
                    cell.nested_tables()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .text_offset
                },
                CellStoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    cell.shapes()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    cell.shape_groups()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                CellStoryEvent::Field(field) => field.position,
                CellStoryEvent::PageBreak(page_break) => page_break.position,
                CellStoryEvent::ColumnBreak(column_break) => column_break.position,
                CellStoryEvent::NavigationEntry(reference)
                | CellStoryEvent::RevisionStart(reference)
                | CellStoryEvent::RevisionEnd(reference)
                | CellStoryEvent::RevisionDeletion(reference) => reference.position,
            };
            let fragment = cell.text().get(offset..position).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF table-cell event splits or leaves its story text",
                )
            })?;
            self.write_text(fragment)?;
            match *event {
                CellStoryEvent::NestedTable(index) => {
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
                CellStoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    let shape = cell
                        .shapes()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    self.write_root_shape(shape)?;
                },
                CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    let group = cell
                        .shape_groups()
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    self.write_shape_group(group, true)?;
                },
                CellStoryEvent::Field(reference) => {
                    let field = fields
                        .get(reference.field_index)
                        .filter(|field| {
                            field.owner == FieldOwner::TableCell(field_owner_depth)
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
                CellStoryEvent::PageBreak(_) => self.write_str("\\page ")?,
                CellStoryEvent::ColumnBreak(_) => self.write_str("\\column ")?,
                CellStoryEvent::NavigationEntry(reference) => {
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
                CellStoryEvent::RevisionStart(reference) => {
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
                CellStoryEvent::RevisionEnd(reference) => {
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
                CellStoryEvent::RevisionDeletion(reference) => {
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
        table: &Table<'_>,
        depth: usize,
        fields: &[Field<'_>],
        navigation_entries: &[NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        let depth_value = i32::try_from(depth).map_err(|_err| {
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
                self.write_control_word("tblrsid", Some(table_rsid.cast_signed()))?;
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
                let cell_index = i32::try_from(index + 1).map_err(|_err| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF table cell index exceeds the i32 range",
                    )
                })?;
                self.write_control_word(
                    "cellx",
                    Some(cell.right_boundary().unwrap_or(2880 * cell_index)),
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
        width: Option<TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                TablePreferredWidthUnit::Null => 0,
                TablePreferredWidthUnit::Auto => 1,
                TablePreferredWidthUnit::Percent => 2,
                TablePreferredWidthUnit::Twips => 3,
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
        width: Option<TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                TablePreferredWidthUnit::Null => 0,
                TablePreferredWidthUnit::Auto => 1,
                TablePreferredWidthUnit::Percent => 2,
                TablePreferredWidthUnit::Twips => 3,
            }),
        )?;
        if let Some(value) = width.value().filter(|value| *value != 0) {
            self.write_control_word(value_control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_row_geometry(
        &mut self,
        geometry: TableRowGeometry,
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
            TableRowHeight::Automatic => {},
            TableRowHeight::Minimum(value) => {
                self.write_control_word("trrh", Some(i32::from(value)))?;
            },
            TableRowHeight::Exact(value) => {
                self.write_control_word("trrh", Some(-i32::from(value)))?;
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
                    TableIndentUnit::Auto => 0,
                    TableIndentUnit::Twips => 1,
                    TableIndentUnit::Nil => 2,
                    TableIndentUnit::Percent => 3,
                }),
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_row_layout(&mut self, row: &Row<'_>) -> io::Result<()> {
        if let Some(alignment) = row.layout().alignment {
            self.write_control_word(
                match alignment {
                    TableRowAlignment::Left => "trql",
                    TableRowAlignment::Center => "trqc",
                    TableRowAlignment::Right => "trqr",
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
                    TableRowBandIndex::Header => -1,
                    TableRowBandIndex::Row(value) => i32::from(value),
                }),
            )?;
        }
        let flags = row.autoformat_flags();
        for (flag, word) in [
            (TableAutoformatFlag::Border, "tbllkborder"),
            (TableAutoformatFlag::Shading, "tbllkshading"),
            (TableAutoformatFlag::Font, "tbllkfont"),
            (TableAutoformatFlag::Color, "tbllkcolor"),
            (TableAutoformatFlag::BestFit, "tbllkbestfit"),
            (TableAutoformatFlag::HeaderRows, "tbllkhdrrows"),
            (TableAutoformatFlag::LastRow, "tbllklastrow"),
            (TableAutoformatFlag::HeaderColumns, "tbllkhdrcols"),
            (TableAutoformatFlag::LastColumn, "tbllklastcol"),
            (TableAutoformatFlag::NoRowBanding, "tbllknorowband"),
            (TableAutoformatFlag::NoColumnBanding, "tbllknocolband"),
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

    pub(in super::super) fn write_table_cell_layout(&mut self, cell: &Cell<'_>) -> io::Result<()> {
        let layout = cell.layout();
        if let Some(alignment) = layout.vertical_alignment {
            self.write_control_word(
                match alignment {
                    TableCellVerticalAlignment::Top => "clvertalt",
                    TableCellVerticalAlignment::Center => "clvertalc",
                    TableCellVerticalAlignment::Bottom => "clvertalb",
                },
                None,
            )?;
        }
        if let Some(flow) = layout.text_flow {
            self.write_control_word(
                match flow {
                    TableCellTextFlow::LeftToRightTopToBottom => "cltxlrtb",
                    TableCellTextFlow::RightToLeftTopToBottom => "cltxtbrl",
                    TableCellTextFlow::LeftToRightBottomToTop => "cltxbtlr",
                    TableCellTextFlow::LeftToRightTopToBottomVertical => "cltxlrtbv",
                    TableCellTextFlow::TopToBottomRightToLeftVertical => "cltxtbrlv",
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

    pub(in super::super) fn write_table_cell_merge(&mut self, cell: &Cell<'_>) -> io::Result<()> {
        let merge = cell.merge();
        if let Some(role) = merge.horizontal {
            self.write_control_word(
                match role {
                    TableCellMergeRole::First => "clmgf",
                    TableCellMergeRole::Continuation => "clmrg",
                },
                None,
            )?;
        }
        if let Some(role) = merge.vertical {
            self.write_control_word(
                match role {
                    TableCellMergeRole::First => "clvmgf",
                    TableCellMergeRole::Continuation => "clvmrg",
                },
                None,
            )?;
        }
        Ok(())
    }

    pub(in super::super) fn write_table_cell_revision(
        &mut self,
        cell: &Cell<'_>,
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
        borders: &TableRowBorders,
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
        borders: &TableCellBorders,
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
        defaults: &TableRowCellDefaults,
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

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(in super::super) fn write_table_shading(
        &mut self,
        prefix: &str,
        shading: TableShading,
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
                ShadingPattern::Horizontal => "bghoriz",
                ShadingPattern::Vertical => "bgvert",
                ShadingPattern::ForwardDiagonal => "bgfdiag",
                ShadingPattern::BackwardDiagonal => "bgbdiag",
                ShadingPattern::Cross => "bgcross",
                ShadingPattern::DiagonalCross => "bgdcross",
                ShadingPattern::DarkHorizontal => "bgdkhor",
                ShadingPattern::DarkVertical => "bgdkvert",
                ShadingPattern::DarkForwardDiagonal => "bgdkfdiag",
                ShadingPattern::DarkBackwardDiagonal => "bgdkbdiag",
                ShadingPattern::DarkCross => "bgdkcross",
                ShadingPattern::DarkDiagonalCross => "bgdkdcross",
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
        position: &FloatingTablePosition,
    ) -> io::Result<()> {
        position
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(reference) = position.horizontal_reference {
            self.write_control_word(
                match reference {
                    TableHorizontalReference::Column => "tphcol",
                    TableHorizontalReference::Margin => "tphmrg",
                    TableHorizontalReference::Page => "tphpg",
                },
                None,
            )?;
        }
        if let Some(value) = position.horizontal_position {
            let (word, param) = match value {
                TableHorizontalPosition::Offset(value) => ("tposx", Some(value)),
                TableHorizontalPosition::NegativeOffset(value) => ("tposnegx", Some(value)),
                TableHorizontalPosition::Center => ("tposxc", None),
                TableHorizontalPosition::Inside => ("tposxi", None),
                TableHorizontalPosition::Left => ("tposxl", None),
                TableHorizontalPosition::Outside => ("tposxo", None),
                TableHorizontalPosition::Right => ("tposxr", None),
            };
            self.write_control_word(word, param)?;
        }
        if let Some(reference) = position.vertical_reference {
            self.write_control_word(
                match reference {
                    TableVerticalReference::Margin => "tpvmrg",
                    TableVerticalReference::Paragraph => "tpvpara",
                    TableVerticalReference::Page => "tpvpg",
                },
                None,
            )?;
        }
        if let Some(value) = position.vertical_position {
            let (word, param) = match value {
                TableVerticalPosition::Offset(value) => ("tposy", Some(value)),
                TableVerticalPosition::NegativeOffset(value) => ("tposnegy", Some(value)),
                TableVerticalPosition::Bottom => ("tposyb", None),
                TableVerticalPosition::Center => ("tposyc", None),
                TableVerticalPosition::Inline => ("tposyil", None),
                TableVerticalPosition::Inside => ("tposyin", None),
                TableVerticalPosition::Outside => ("tposyout", None),
                TableVerticalPosition::Top => ("tposyt", None),
            };
            self.write_control_word(word, param)?;
        }
        for (word, value) in [
            ("tdfrmtxtLeft", position.wrap_distances.left),
            ("tdfrmtxtRight", position.wrap_distances.right),
            ("tdfrmtxtTop", position.wrap_distances.top),
            ("tdfrmtxtBottom", position.wrap_distances.bottom),
        ] {
            if let Some(value) = value {
                self.write_control_word(word, Some(i32::from(value)))?;
            }
        }
        if position.no_overlap {
            self.write_control_word("tabsnoovrlp", None)?;
        }
        Ok(())
    }
}
