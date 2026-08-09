//! RTF wire primitives and story wrappers.

use super::super::{
    ASCII_CONTROL_LIMIT, DrawingStoryTextMode, EndnoteRestart, Field, FieldOwner, FieldType,
    FootnoteRestart, HeaderFooter, HeaderFooterType, Note, NoteNumberingStyle, PageBorders,
    PageOrientation, Revision, RevisionType, RtfWriter, Section, SectionBreakType,
    SectionFootnotePlacement, SectionLineNumberRestart, SectionNoteOptions, SectionRendering,
    Shape, ShapeGroup, StoryDrawing, StoryEvent, TextDirection, VerticalAlignment, Write, field,
    invalid_story_reference, io,
};

impl<W: Write> RtfWriter<W> {
    /// Write a control word
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_control_word(&mut self, word: &str, param: Option<i32>) -> io::Result<()> {
        self.write_str("\\")?;
        self.write_str(word)?;
        if let Some(p) = param {
            write!(self.writer, "{p}")?;
        }
        Ok(())
    }

    /// Write plain text (with proper escaping)
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_text(&mut self, text: &str) -> io::Result<()> {
        for ch in text.chars() {
            match ch {
                '\\' => self.write_str("\\\\")?,
                '{' => self.write_str("\\{")?,
                '}' => self.write_str("\\}")?,
                // The trailing space delimits the control word. Without it the
                // following character is absorbed into the word itself (`\partwo`)
                // or misread as its numeric parameter (`\par2`), silently
                // destroying the text that follows the break. RTF always consumes
                // a single delimiting space, so it never reappears as content.
                '\n' => self.write_str("\\par ")?,
                '\t' => self.write_str("\\tab ")?,
                // RTF special characters with dedicated control words keep their
                // source spelling instead of a generic \u escape. The trailing
                // space delimits the control word exactly like \par and \tab.
                '\u{2014}' => self.write_str("\\emdash ")?,
                '\u{2013}' => self.write_str("\\endash ")?,
                '\u{2003}' => self.write_str("\\emspace ")?,
                '\u{2002}' => self.write_str("\\enspace ")?,
                '\u{2005}' => self.write_str("\\qmspace ")?,
                '\u{2022}' => self.write_str("\\bullet ")?,
                '\u{2018}' => self.write_str("\\lquote ")?,
                '\u{2019}' => self.write_str("\\rquote ")?,
                '\u{201C}' => self.write_str("\\ldblquote ")?,
                '\u{201D}' => self.write_str("\\rdblquote ")?,
                '\u{200E}' => self.write_str("\\ltrmark ")?,
                '\u{200F}' => self.write_str("\\rtlmark ")?,
                '\u{200D}' => self.write_str("\\zwj ")?,
                '\u{200C}' => self.write_str("\\zwnj ")?,
                '\u{200B}' => self.write_str("\\zwbo ")?,
                '\u{FEFF}' => self.write_str("\\zwnbo ")?,
                // Readers discard raw carriage returns and other bare control
                // bytes as line-ending noise, so emit them as hex escapes to keep
                // them part of the document text.
                c if (c as u32) < ASCII_CONTROL_LIMIT => {
                    write!(self.writer, "\\'{:02x}", c as u32)?;
                },
                c if c.is_ascii() => {
                    write!(self.writer, "{c}")?;
                },
                c => {
                    // Write Unicode character
                    let code = c as i32;
                    self.write_control_word("u", Some(code))?;
                    // Fallback character
                    self.write_str("?")?;
                },
            }
        }
        Ok(())
    }

    /// Write a string
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_str(&mut self, s: &str) -> io::Result<()> {
        self.writer.write_all(s.as_bytes())
    }

    /// Flush the writer
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }

    /// Write a header or footer
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_header_footer(&mut self, hf: &HeaderFooter<'_>) -> io::Result<()> {
        self.write_header_footer_with_fields(hf, &[])
    }

    pub(in super::super) fn write_header_footer_with_fields(
        &mut self,
        hf: &HeaderFooter<'_>,
        fields: &[Field<'_>],
    ) -> io::Result<()> {
        enum EventKind<'b, 'a> {
            Shape(&'b Shape<'a>),
            Group(&'b ShapeGroup<'a>),
            Field(&'b Field<'a>),
            PageBreak,
        }
        struct Event<'b, 'a> {
            offset: usize,
            kind: EventKind<'b, 'a>,
        }

        self.write_str("{")?;

        // Write header/footer type control word
        match hf.header_type {
            HeaderFooterType::Header => self.write_control_word("header", None)?,
            HeaderFooterType::HeaderFirst => self.write_control_word("headerf", None)?,
            HeaderFooterType::HeaderLeft => self.write_control_word("headerl", None)?,
            HeaderFooterType::HeaderRight => self.write_control_word("headerr", None)?,
            HeaderFooterType::Footer => self.write_control_word("footer", None)?,
            HeaderFooterType::FooterFirst => self.write_control_word("footerf", None)?,
            HeaderFooterType::FooterLeft => self.write_control_word("footerl", None)?,
            HeaderFooterType::FooterRight => self.write_control_word("footerr", None)?,
        }

        let story = hf.text();
        field::validate_story_events(
            &story,
            &hf.shapes,
            &hf.shape_groups,
            &hf.drawing_order,
            &hf.story_events,
            "header/footer",
        )
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let owner = match hf.header_type {
            HeaderFooterType::Header
            | HeaderFooterType::HeaderFirst
            | HeaderFooterType::HeaderLeft
            | HeaderFooterType::HeaderRight => FieldOwner::Header,
            HeaderFooterType::Footer
            | HeaderFooterType::FooterFirst
            | HeaderFooterType::FooterLeft
            | HeaderFooterType::FooterRight => FieldOwner::Footer,
        };
        let mut events = Vec::new();
        events.try_reserve(hf.story_events.len()).map_err(|_err| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF header/footer event table exceeds available memory",
            )
        })?;
        for story_event in &hf.story_events {
            match *story_event {
                StoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    let shape = hf.shapes.get(index).ok_or_else(invalid_story_reference)?;
                    events.push(Event {
                        offset: shape.position,
                        kind: EventKind::Shape(shape),
                    });
                },
                StoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    let group = hf
                        .shape_groups
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    events.push(Event {
                        offset: group.position,
                        kind: EventKind::Group(group),
                    });
                },
                StoryEvent::Field(reference) => events.push(Event {
                    offset: reference.position,
                    kind: EventKind::Field(fields.get(reference.field_index).filter(|field| field.owner == owner && field.position == reference.position && field.range_end == reference.position).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF header/footer story has an invalid generic-field owner or reference"))?),
                }),
                StoryEvent::PageBreak(page_break) => events.push(Event {
                    offset: page_break.position,
                    kind: EventKind::PageBreak,
                }),
            }
        }
        /* Keep paragraph formatting while splitting its text around story events. */
        let mut next_event = 0usize;
        let mut story_offset = 0usize;

        for para in &hf.paragraphs {
            self.write_formatting(&para.formatting)?;
            self.write_paragraph_properties(&para.paragraph)?;
            self.write_str(" ")?;
            let text = para.text.as_ref();
            let end = story_offset.checked_add(text.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF header/footer text size overflow",
                )
            })?;
            let mut local = 0usize;
            while let Some(event) = events.get(next_event).filter(|event| event.offset <= end) {
                let split = event.offset.checked_sub(story_offset).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF header/footer events are out of story order",
                    )
                })?;
                let fragment = text.get(local..split).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF header/footer event splits or leaves its paragraph text",
                    )
                })?;
                self.write_text(fragment)?;
                match event.kind {
                    EventKind::Shape(shape) => self.write_root_shape(shape)?,
                    EventKind::Group(group) => self.write_shape_group(group, true)?,
                    EventKind::Field(field) => self.write_field_with_fields(field, fields, 0)?,
                    EventKind::PageBreak => self.write_str("\\page ")?,
                }
                local = split;
                next_event += 1;
            }
            let remainder = text.get(local..).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF header/footer event leaves its paragraph text",
                )
            })?;
            self.write_text(remainder)?;
            self.write_control_word("par", None)?;
            story_offset = end.checked_add(1).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF header/footer text size overflow",
                )
            })?;
        }
        while let Some(event) = events.get(next_event) {
            if event.offset != story.len() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF header/footer event position is unreachable",
                ));
            }
            match event.kind {
                EventKind::Shape(shape) => self.write_root_shape(shape)?,
                EventKind::Group(group) => self.write_shape_group(group, true)?,
                EventKind::Field(field) => self.write_field_with_fields(field, fields, 0)?,
                EventKind::PageBreak => self.write_str("\\page ")?,
            }
            next_event += 1;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write a footnote or endnote
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_note(&mut self, note: &Note<'_>) -> io::Result<()> {
        self.write_note_with_fields(note, &[])
    }

    pub(in super::super) fn write_note_with_fields(
        &mut self,
        note: &Note<'_>,
        fields: &[Field<'_>],
    ) -> io::Result<()> {
        note.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;

        // Write note type control word
        if note.is_footnote {
            self.write_control_word("footnote", None)?;
        } else {
            self.write_control_word("endnote", None)?;
        }

        // Write reference number/marker
        if !note.reference.is_empty()
            && let Ok(num) = note.reference.parse::<i32>()
        {
            self.write_control_word("chftn", Some(num))?;
        }

        // Write note content
        self.write_str(" {")?;
        self.write_formatting(&note.formatting)?;
        self.write_field_story(
            note.content.as_ref(),
            &note.shapes,
            &note.shape_groups,
            &note.drawing_order,
            &note.story_events,
            fields,
            if note.is_footnote {
                FieldOwner::Footnote
            } else {
                FieldOwner::Endnote
            },
            DrawingStoryTextMode::Note,
            0,
        )?;
        self.write_str("}")?;

        self.write_str("}")?;
        Ok(())
    }

    /// Write a hyperlink field
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_hyperlink(&mut self, url: &str, display_text: &str) -> io::Result<()> {
        let instruction = format!("HYPERLINK {}", field::quoted_field_operand(url));
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    /// Write an internal bookmark hyperlink without exposing raw field syntax.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_internal_hyperlink(
        &mut self,
        bookmark: &str,
        display_text: &str,
    ) -> io::Result<()> {
        let instruction = format!("HYPERLINK \\l {}", field::quoted_field_operand(bookmark));
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    pub(in super::super) fn write_hyperlink_instruction(
        &mut self,
        instruction: &str,
        display_text: &str,
    ) -> io::Result<()> {
        self.write_str("{\\field")?;
        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(instruction)?;
        self.write_str("}}")?;

        // Field result (display text)
        self.write_str("{\\fldrslt{")?;
        self.write_control_word("ul", None)?; // Underline hyperlinks by default
        self.write_control_word("cf", Some(1))?; // Blue color for hyperlinks
        self.write_text(display_text)?;
        self.write_str("}}}")?;

        Ok(())
    }

    /// Write a field (generic)
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_field(&mut self, field: &Field<'_>) -> io::Result<()> {
        self.write_field_with_fields(field, &[], 0)
    }

    /// Write a caller-provided legacy `EQ` expression as an inert RTF field.
    ///
    /// The expression is escaped for the field instruction and emitted with
    /// the empty cached-result group conventionally used for `EQ`. It is never
    /// parsed, calculated, formatted, or rendered by this library.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_equation(&mut self, expression: &str) -> io::Result<()> {
        let field = Field::new_equation(expression)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_field(&field)
    }

    pub(in super::super) fn write_field_with_fields(
        &mut self,
        field: &Field<'_>,
        fields: &[Field<'_>],
        depth: usize,
    ) -> io::Result<()> {
        if depth > 64 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF nested generic fields exceed 64 levels",
            ));
        }
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\field")?;
        if field.status.dirty {
            self.write_str("\\flddirty")?;
        }
        if field.status.edited {
            self.write_str("\\fldedit")?;
        }
        if field.status.locked {
            self.write_str("\\fldlock")?;
        }
        if field.status.private {
            self.write_str("\\fldpriv")?;
        }

        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(field.instruction.as_ref())?;
        self.write_str("}}")?;

        // Field result
        if field.field_type == FieldType::Equation
            && field.result.is_empty()
            && field.result_events.is_empty()
        {
            // RTF 1.9.1 examples write a null fldrslt group for EQ fields.
            self.write_str("{\\fldrslt}")?;
        } else if !field.result.is_empty() || !field.result_events.is_empty() {
            self.write_str("{\\fldrslt{")?;
            self.write_field_story(
                field.result.as_ref(),
                &field.shapes,
                &field.shape_groups,
                &field.drawing_order,
                &field.result_events,
                fields,
                FieldOwner::FieldResult,
                DrawingStoryTextMode::ShapeText,
                depth,
            )?;
            self.write_str("}}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write a revision mark (track changes)
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_revision(&mut self, revision: &Revision<'_>) -> io::Result<()> {
        revision
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_revision_start(revision)?;
        self.write_text(revision.content.as_ref())?;
        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn write_revision_start(
        &mut self,
        revision: &Revision<'_>,
    ) -> io::Result<()> {
        self.write_str("{")?;
        let (kind, author, date) = match revision.revision_type {
            RevisionType::Insertion => ("revised", "revauth", "revdttm"),
            RevisionType::Deletion => ("deleted", "revauthdel", "revdttmdel"),
            RevisionType::FormatChange | RevisionType::MovedFrom | RevisionType::MovedTo => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "this RTF revision kind has no lossless scoped-run representation",
                ));
            },
        };
        self.write_control_word(kind, None)?;
        self.write_control_word(author, Some(revision.id))?;
        if let Some(date_value) = revision.date.as_deref() {
            let packed = date_value.parse::<i32>().map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision dates must contain the packed signed DTTM value",
                )
            })?;
            self.write_control_word(date, Some(packed))?;
        }
        self.write_str(" ")?;
        Ok(())
    }

    /// Write a section with headers and footers
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_section(&mut self, section: &Section<'_>) -> io::Result<()> {
        self.write_section_with_fields(section, &[])
    }

    pub(in super::super) fn write_section_with_fields(
        &mut self,
        section: &Section<'_>,
        fields: &[Field<'_>],
    ) -> io::Result<()> {
        // Write section properties
        self.write_control_word("sectd", None)?;
        if let Some(section_style) = section.properties.section_style {
            self.write_control_word("ds", Some(i32::from(section_style)))?;
        }
        if let Some(section_rsid) = section.properties.section_rsid {
            self.write_control_word("sectrsid", Some(section_rsid.cast_signed()))?;
        }
        self.write_revision_metadata("srauth", "srdate", section.properties.revision)?;
        if section.properties.title_page {
            self.write_control_word("titlepg", None)?;
        }
        self.write_section_note_options(&section.properties.note_options)?;
        self.write_page_borders(&section.properties.page_borders)?;

        if let Some(direction) = section.properties.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrsect",
                    TextDirection::RightToLeft => "rtlsect",
                },
                None,
            )?;
        }

        match section.properties.break_type {
            SectionBreakType::Continuous => self.write_control_word("sbknone", None)?,
            SectionBreakType::Column => self.write_control_word("sbkcol", None)?,
            SectionBreakType::Page => self.write_control_word("sbkpage", None)?,
            SectionBreakType::EvenPage => self.write_control_word("sbkeven", None)?,
            SectionBreakType::OddPage => self.write_control_word("sbkodd", None)?,
        }

        // Page size
        self.write_control_word("pgwsxn", Some(section.properties.page_width))?;
        self.write_control_word("pghsxn", Some(section.properties.page_height))?;

        // Margins
        self.write_control_word("marglsxn", Some(section.properties.margin_left))?;
        self.write_control_word("margrsxn", Some(section.properties.margin_right))?;
        self.write_control_word("margtsxn", Some(section.properties.margin_top))?;
        self.write_control_word("margbsxn", Some(section.properties.margin_bottom))?;
        self.write_control_word("guttersxn", Some(section.properties.margin_gutter))?;

        // Paper-source bins
        if let Some(first) = section.properties.paper_source.first {
            self.write_control_word("binfsxn", Some(i32::from(first)))?;
        }
        if let Some(other) = section.properties.paper_source.other {
            self.write_control_word("binsxn", Some(i32::from(other)))?;
        }

        // Header/footer distance
        self.write_control_word("headery", Some(section.properties.header_distance))?;
        self.write_control_word("footery", Some(section.properties.footer_distance))?;

        if section.properties.orientation == PageOrientation::Landscape {
            self.write_control_word("lndscpsxn", None)?;
        }
        if let Some(rendering) = section.properties.rendering {
            self.write_control_word(
                match rendering {
                    SectionRendering::Horizontal => "horzsect",
                    SectionRendering::Vertical => "vertsect",
                },
                None,
            )?;
        }
        section
            .properties
            .columns
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word("cols", Some(i32::from(section.properties.columns.count)))?;
        if !section.properties.balance_columns {
            self.write_control_word("nocolbal", None)?;
        }
        if section.properties.columns.separator {
            self.write_control_word("linebetcol", None)?;
        }
        self.write_control_word("colsx", Some(section.properties.columns.default_spacing))?;
        for (index, column) in section.properties.columns.explicit.iter().enumerate() {
            let column_number = i32::try_from(index + 1).map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF column count exceeds the i32 range",
                )
            })?;
            self.write_control_word("colno", Some(column_number))?;
            self.write_control_word("colw", Some(column.width))?;
            if let Some(space) = column.space_after {
                self.write_control_word("colsr", Some(space))?;
            }
        }
        self.write_control_word("pgnstarts", Some(section.properties.page_number_start))?;
        self.write_control_word(section.properties.page_number_format.control_word(), None)?;
        if let Some(restart) = section.properties.page_number_restart {
            self.write_control_word(restart.control_word(), None)?;
        }
        if let Some(offset_x) = section.properties.page_number_offset_x {
            self.write_control_word("pgnx", Some(offset_x))?;
        }
        if let Some(offset_y) = section.properties.page_number_offset_y {
            self.write_control_word("pgny", Some(offset_y))?;
        }
        section
            .properties
            .page_number_heading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(level) = section.properties.page_number_heading.level {
            self.write_control_word("pgnhn", Some(i32::from(level)))?;
        }
        if let Some(separator) = section.properties.page_number_heading.separator {
            self.write_control_word(separator.control_word(), None)?;
        }
        section
            .properties
            .document_grid
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(line_grid) = section.properties.document_grid.line_grid {
            self.write_control_word("sectlinegrid", Some(line_grid))?;
        }
        if let Some(grid_type) = section.properties.document_grid.grid_type {
            self.write_control_word(grid_type.control_word(), None)?;
        }
        self.write_control_word(
            match section.properties.vertical_alignment {
                VerticalAlignment::Top => "vertalt",
                VerticalAlignment::Center => "vertalc",
                VerticalAlignment::Justify => "vertalj",
                VerticalAlignment::Bottom => "vertalb",
            },
            None,
        )?;
        section
            .properties
            .line_numbering
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(increment) = section.properties.line_numbering.increment {
            self.write_control_word("linemod", Some(i32::from(increment)))?;
        }
        if let Some(distance) = section.properties.line_numbering.distance {
            self.write_control_word("linex", Some(distance))?;
        }
        if let Some(start) = section.properties.line_numbering.start {
            self.write_control_word("linestarts", Some(start.cast_signed()))?;
        }
        if let Some(restart) = section.properties.line_numbering.restart {
            self.write_control_word(
                match restart {
                    SectionLineNumberRestart::Section => "linerestart",
                    SectionLineNumberRestart::Page => "lineppage",
                    SectionLineNumberRestart::Continuous => "linecont",
                },
                None,
            )?;
        }

        // Write all headers and footers for this section
        for hf in &section.headers_footers {
            self.write_header_footer_with_fields(hf, fields)?;
        }

        Ok(())
    }

    /// Write canonical section page-border options and edges.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_page_borders(&mut self, borders: &PageBorders) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if borders.is_empty() {
            return Ok(());
        }
        if borders.option_value() != 0 {
            self.write_control_word("pgbrdropt", Some(borders.option_value()))?;
        }
        if borders.surround_header {
            self.write_control_word("pgbrdrhead", None)?;
        }
        if borders.surround_footer {
            self.write_control_word("pgbrdrfoot", None)?;
        }
        if borders.snap_to_text_borders {
            self.write_control_word("pgbrdrsnap", None)?;
        }
        for (control, side) in [
            ("pgbrdrt", borders.top),
            ("pgbrdrl", borders.left),
            ("pgbrdrb", borders.bottom),
            ("pgbrdrr", borders.right),
        ] {
            let Some(border) = side else {
                continue;
            };
            self.write_control_word(control, None)?;
            if let Some(art) = border.art {
                self.write_control_word("brdrart", Some(i32::from(art)))?;
            } else {
                self.write_control_word(border.style.control_word(), None)?;
            }
            self.write_control_word("brdrw", Some(i32::from(border.width)))?;
            self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
            self.write_control_word("brsp", Some(i32::from(border.space)))?;
            if border.shadow {
                self.write_control_word("brdrsh", None)?;
            }
            if border.frame {
                self.write_control_word("brdrframe", None)?;
            }
        }
        Ok(())
    }

    /// Write explicit section-level footnote and endnote overrides.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_section_note_options(&mut self, options: &SectionNoteOptions) -> io::Result<()> {
        options
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if options.endnote_here {
            self.write_control_word("endnhere", None)?;
        }
        if let Some(value) = options.footnote_placement {
            self.write_control_word(
                match value {
                    SectionFootnotePlacement::BeneathText => "sftntj",
                    SectionFootnotePlacement::BottomOfPage => "sftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_start {
            self.write_control_word("sftnstart", Some(value))?;
        }
        if let Some(value) = options.footnote_restart {
            self.write_control_word(
                match value {
                    FootnoteRestart::Continuous => "sftnrstcont",
                    FootnoteRestart::EachSection => "sftnrestart",
                    FootnoteRestart::EachPage => "sftnrstpg",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_numbering {
            self.write_control_word(Self::section_note_numbering_control(value, false), None)?;
        }
        if let Some(value) = options.endnote_start {
            self.write_control_word("saftnstart", Some(value))?;
        }
        if let Some(value) = options.endnote_restart {
            self.write_control_word(
                match value {
                    EndnoteRestart::Continuous => "saftnrstcont",
                    EndnoteRestart::EachSection => "saftnrestart",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_numbering {
            self.write_control_word(Self::section_note_numbering_control(value, true), None)?;
        }
        Ok(())
    }

    pub(in super::super) fn section_note_numbering_control(
        style: NoteNumberingStyle,
        endnote: bool,
    ) -> &'static str {
        match (endnote, style) {
            (false, NoteNumberingStyle::Arabic) => "sftnnar",
            (false, NoteNumberingStyle::LowercaseLetter) => "sftnnalc",
            (false, NoteNumberingStyle::UppercaseLetter) => "sftnnauc",
            (false, NoteNumberingStyle::LowercaseRoman) => "sftnnrlc",
            (false, NoteNumberingStyle::UppercaseRoman) => "sftnnruc",
            (false, NoteNumberingStyle::Chicago) => "sftnnchi",
            (false, NoteNumberingStyle::KoreanChosung) => "sftnnchosung",
            (false, NoteNumberingStyle::Circle) => "sftnncnum",
            (false, NoteNumberingStyle::KanjiDigitless) => "sftnndbnum",
            (false, NoteNumberingStyle::KanjiWithDigit) => "sftnndbnumd",
            (false, NoteNumberingStyle::KanjiThree) => "sftnndbnumt",
            (false, NoteNumberingStyle::KanjiFour) => "sftnndbnumk",
            (false, NoteNumberingStyle::DoubleByte) => "sftnndbar",
            (false, NoteNumberingStyle::KoreanGanada) => "sftnnganada",
            (false, NoteNumberingStyle::ChineseOne) => "sftnngbnum",
            (false, NoteNumberingStyle::ChineseTwo) => "sftnngbnumd",
            (false, NoteNumberingStyle::ChineseThree) => "sftnngbnuml",
            (false, NoteNumberingStyle::ChineseFour) => "sftnngbnumk",
            (false, NoteNumberingStyle::ZodiacOne) => "sftnnzodiac",
            (false, NoteNumberingStyle::ZodiacTwo) => "sftnnzodiacd",
            (false, NoteNumberingStyle::ZodiacThree) => "sftnnzodiacl",
            (true, NoteNumberingStyle::Arabic) => "saftnnar",
            (true, NoteNumberingStyle::LowercaseLetter) => "saftnnalc",
            (true, NoteNumberingStyle::UppercaseLetter) => "saftnnauc",
            (true, NoteNumberingStyle::LowercaseRoman) => "saftnnrlc",
            (true, NoteNumberingStyle::UppercaseRoman) => "saftnnruc",
            (true, NoteNumberingStyle::Chicago) => "saftnnchi",
            (true, NoteNumberingStyle::KoreanChosung) => "saftnnchosung",
            (true, NoteNumberingStyle::Circle) => "saftnncnum",
            (true, NoteNumberingStyle::KanjiDigitless) => "saftnndbnum",
            (true, NoteNumberingStyle::KanjiWithDigit) => "saftnndbnumd",
            (true, NoteNumberingStyle::KanjiThree) => "saftnndbnumt",
            (true, NoteNumberingStyle::KanjiFour) => "saftnndbnumk",
            (true, NoteNumberingStyle::DoubleByte) => "saftnndbar",
            (true, NoteNumberingStyle::KoreanGanada) => "saftnnganada",
            (true, NoteNumberingStyle::ChineseOne) => "saftnngbnum",
            (true, NoteNumberingStyle::ChineseTwo) => "saftnngbnumd",
            (true, NoteNumberingStyle::ChineseThree) => "saftnngbnuml",
            (true, NoteNumberingStyle::ChineseFour) => "saftnngbnumk",
            (true, NoteNumberingStyle::ZodiacOne) => "saftnnzodiac",
            (true, NoteNumberingStyle::ZodiacTwo) => "saftnnzodiacd",
            (true, NoteNumberingStyle::ZodiacThree) => "saftnnzodiacl",
        }
    }
}
