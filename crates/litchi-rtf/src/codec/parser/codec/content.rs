use super::{
    ControlWord, Cow, Destination, FontCharset, MAX_REVISIONS, ParsedBodyStoryEvent, Parser,
    RtfEncoding, RtfError, RtfResult, SmallVec, State, StyleBlock, Token, append_transport_bytes,
    control_symbol_text, parser_classification_error, require_parameterless,
};

impl<'a> Parser<'a> {
    /// Handle a control word encountered while parsing generic group content.
    ///
    /// Split out of [`Parser::parse_content`] and marked `#[inline(never)]` for
    /// the same reason as [`Parser::dispatch_group_destination`]: the dispatch
    /// table is large, and leaving it inline would charge its stack frame to
    /// every level of group nesting on the recursive path.
    #[inline(never)]
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn dispatch_content_control(
        &mut self,
        control: &ControlWord<'a>,
        text_buffer: &mut SmallVec<[u8; 256]>,
    ) -> RtfResult<()> {
        if matches!(control, ControlWord::Unknown(_, _)) {
            if !text_buffer.is_empty() {
                self.flush_text_buffer(text_buffer)?;
            }
            let token = self.pos;
            self.pos += 1;
            self.preserve_unknown_control(token)?;
            return Ok(());
        }
        if matches!(control, ControlWord::HtmlRtf(_)) {
            if !text_buffer.is_empty() {
                self.flush_text_buffer(text_buffer)?;
            }
            let token = self.pos;
            self.pos += 1;
            self.preserve_unknown_control(token)?;
            return Ok(());
        }
        match control {
            ControlWord::Par | ControlWord::Line => {
                let structural_table_boundary =
                    self.finalize_table_before_non_table_body_content(true)?;
                self.pos += 1;
                // Paragraph break - flush current text
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                if !structural_table_boundary {
                    if self.current_state().is_ok_and(|state| {
                        state.destination == Destination::DocumentBody
                            && state.revision_type
                                != Some(super::super::super::annotation::RevisionType::Deletion)
                    }) {
                        crate::error::try_reserve_one(
                            &mut self.body_boundaries,
                            "body text boundaries",
                        )?;
                        let kind = if matches!(control, ControlWord::Par) {
                            crate::text::Break::Paragraph
                        } else {
                            crate::text::Break::Line
                        };
                        self.body_boundaries
                            .push(crate::story::Boundary::new(self.body_text_len, kind));
                    }
                    text_buffer.push(b'\n');
                }
                if matches!(control, ControlWord::Par) {
                    let state = self.current_state_mut()?;
                    state.paragraph_content_started = false;
                    state.paragraph_numbering_declared = false;
                }
            },
            ControlWord::Page(param) => {
                require_parameterless(*param, "page")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_body_page_break()?;
                self.pos += 1;
            },
            ControlWord::Column(param) => {
                require_parameterless(*param, "column")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_body_column_break()?;
                self.pos += 1;
            },
            ControlWord::EditableRegionStart(param) => {
                require_parameterless(*param, "ebcstart")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_editable_region_boundary(true)?;
                self.pos += 1;
            },
            ControlWord::EditableRegionEnd(param) => {
                require_parameterless(*param, "ebcend")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_editable_region_boundary(false)?;
                self.pos += 1;
            },
            ControlWord::SoftPageBreak(param) => {
                require_parameterless(*param, "softpage")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_soft_break(crate::SoftBreakKind::Page)?;
                self.pos += 1;
            },
            ControlWord::SoftColumnBreak(param) => {
                require_parameterless(*param, "softcol")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_soft_break(crate::SoftBreakKind::Column)?;
                self.pos += 1;
            },
            ControlWord::SoftLineBreak(param) => {
                require_parameterless(*param, "softline")?;
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.record_soft_break(crate::SoftBreakKind::Line)?;
                self.pos += 1;
            },
            ControlWord::SoftLineHeight(param) => {
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                let height = param.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF softlheight requires a numeric parameter".to_string(),
                    )
                })?;
                if !(-32768..=32767).contains(&height) {
                    return Err(RtfError::MalformedDocument(
                        "RTF softlheight is outside -32768..=32767 twips".to_string(),
                    ));
                }
                self.record_soft_break(crate::SoftBreakKind::LineHeight(height))?;
                self.pos += 1;
            },
            ControlWord::Section
            | ControlWord::Ansi
            | ControlWord::AnsiCodePage(_)
            | ControlWord::Mac
            | ControlWord::Pc
            | ControlWord::Pca
            | ControlWord::FontNumber(_)
            | ControlWord::Plain
            | ControlWord::FormProtection(_)
            | ControlWord::AnnotationProtection(_)
            | ControlWord::RevisionProtection(_)
            | ControlWord::ReadOnlyProtection(_)
            | ControlWord::AllProtection(_)
            | ControlWord::EnforceProtection(_)
            | ControlWord::ProtectionLevel(_)
            | ControlWord::ColorBackground(_) => {
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.pos += 1;
                self.apply_control_word(control)?;
            },
            ControlWord::LegacyParagraphNumbering(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF pn control must be the first control in its own destination group"
                        .to_string(),
                ));
            },
            ControlWord::Tab => {
                self.finalize_table_before_non_table_body_content(true)?;
                self.pos += 1;
                text_buffer.push(b'\t');
            },
            ControlWord::Unicode(code) => {
                self.finalize_table_before_non_table_body_content(true)?;
                // Handle Unicode character with potential fallback
                if self
                    .states
                    .last()
                    .is_some_and(|state| state.destination == Destination::DocumentBody)
                {
                    self.section_note_options_closed = true;
                    self.root_section_format_run = false;
                }
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                self.parse_unicode_sequence(*code)?;
            },
            ControlWord::NonBreakingSpace
            | ControlWord::OptionalHyphen
            | ControlWord::NonBreakingHyphen
            | ControlWord::EmDash
            | ControlWord::EnDash
            | ControlWord::EmSpace
            | ControlWord::EnSpace
            | ControlWord::QuarterEmSpace
            | ControlWord::Bullet
            | ControlWord::LeftSingleQuote
            | ControlWord::RightSingleQuote
            | ControlWord::LeftDoubleQuote
            | ControlWord::RightDoubleQuote
            | ControlWord::LeftToRightMark
            | ControlWord::RightToLeftMark
            | ControlWord::ZeroWidthJoiner
            | ControlWord::ZeroWidthNonJoiner
            | ControlWord::ZeroWidthBreakOpportunity
            | ControlWord::ZeroWidthNoBreakOpportunity => {
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                let text = control_symbol_text(control).ok_or_else(|| {
                    RtfError::MalformedDocument("missing RTF control-symbol text".to_string())
                })?;
                self.pos += 1;
                self.append_semantic_text(text)?;
            },
            ControlWord::CurrentDate
            | ControlWord::CurrentDateLong
            | ControlWord::CurrentDateAbbreviated
            | ControlWord::CurrentTime => {
                // These stamps expand to the current date or time when a
                // renderer lays out the document. The parser has no clock or
                // calendar facility, so they contribute no extracted text.
                self.pos += 1;
            },
            ControlWord::AnnotationMark => {
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                if self.pending_annotation_mark {
                    return Err(RtfError::MalformedDocument(
                        "duplicate pending RTF annotation marker".to_string(),
                    ));
                }
                self.pending_annotation_mark = true;
                self.pos += 1;
            },
            ControlWord::Revised(_)
            | ControlWord::Deleted(_)
            | ControlWord::RevisionAuthor(_)
            | ControlWord::DeletedRevisionAuthor(_)
            | ControlWord::RevisionDate(_)
            | ControlWord::DeletedRevisionDate(_) => {
                if !text_buffer.is_empty() {
                    self.flush_text_buffer(text_buffer)?;
                }
                let starting = match control {
                    ControlWord::Revised(true) => {
                        Some(super::super::super::annotation::RevisionType::Insertion)
                    },
                    ControlWord::Deleted(true) => {
                        Some(super::super::super::annotation::RevisionType::Deletion)
                    },
                    _ => None,
                };
                if matches!(control, ControlWord::Revised(false))
                    && let Some(id) = self.current_state()?.revision_event_id
                {
                    self.record_revision_end(id)?;
                }
                let in_table = self.current_state()?.in_table
                    || self.current_state()?.table_nesting_level >= 2;
                let event_id = if let Some(kind) = starting {
                    let id = self.revision_event_indices.len();
                    self.revision_event_indices.push(None);
                    if !in_table {
                        let event = match kind {
                            super::super::super::annotation::RevisionType::Insertion => {
                                ParsedBodyStoryEvent::RevisionStart(id)
                            },
                            super::super::super::annotation::RevisionType::Deletion => {
                                ParsedBodyStoryEvent::RevisionDeletion(id)
                            },
                            super::super::super::annotation::RevisionType::FormatChange
                            | super::super::super::annotation::RevisionType::MovedFrom
                            | super::super::super::annotation::RevisionType::MovedTo => {
                                return Err(parser_classification_error());
                            },
                        };
                        self.body_story_events.push(event);
                    }
                    Some(id)
                } else {
                    None
                };
                self.pos += 1;
                self.apply_control_word(control)?;
                if let Some(id) = event_id {
                    self.current_state_mut()?.revision_event_id = Some(id);
                }
            },
            _ => {
                self.pos += 1;
                // Apply formatting changes
                self.apply_control_word(control)?;
            },
        }

        Ok(())
    }

    /// Parse group content (text and control words).
    pub(super) fn parse_content(&mut self) -> RtfResult<()> {
        let mut text_buffer = SmallVec::<[u8; 256]>::new();

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    // Flush any buffered text
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.pos += 1;
                    return Ok(());
                },
                Token::OpenBrace => {
                    // Flush text before entering nested group
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.current_state_mut()?.character_border_active = false;
                    self.parse_group()?;
                },
                Token::Control(control) => {
                    self.dispatch_content_control(control, &mut text_buffer)?;
                },
                Token::Text(text) => {
                    self.pos += 1;
                    self.current_state_mut()?.character_border_active = false;
                    // Skip empty text tokens
                    if text.is_empty() {
                        continue;
                    }
                    self.finalize_table_before_non_table_body_content(!text.trim().is_empty())?;
                    if self
                        .states
                        .last()
                        .is_some_and(|state| state.destination == Destination::DocumentBody)
                        && !text.trim().is_empty()
                    {
                        self.note_options_closed = true;
                        self.section_note_options_closed = true;
                        self.root_section_format_run = false;
                    }
                    // Check if we're in a table
                    if self.current_state().is_ok_and(|s| {
                        s.destination == Destination::DocumentBody
                            && (s.in_table || s.table_nesting_level >= 2)
                    }) {
                        let state = self.current_state()?.clone();
                        let encoding = self.effective_text_encoding(&state)?;
                        let mut bytes = SmallVec::<[u8; 64]>::new();
                        append_transport_bytes(&mut bytes, text)?;
                        self.append_table_text(
                            encoding.decode(&bytes).as_bytes(),
                            state.table_nesting_level,
                        )?;
                    } else if self
                        .current_state()
                        .is_ok_and(|s| s.destination == Destination::DocumentBody)
                    {
                        append_transport_bytes(&mut text_buffer, text)?;
                    }
                },
                Token::Binary(_) => {
                    if self.current_state()?.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision text cannot contain binary data".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
            }
        }

        Err(RtfError::UnexpectedEof)
    }

    /// Flush text buffer to a style block.
    pub(super) fn flush_text_buffer(&mut self, buffer: &mut SmallVec<[u8; 256]>) -> RtfResult<()> {
        if buffer.is_empty() {
            return Ok(());
        }

        let state = self.current_state()?.clone();

        // Only create blocks for text in the document body
        // Skip text from font tables, color tables, stylesheets, etc.
        if state.destination == Destination::DocumentBody {
            let decoded_str = self.effective_text_encoding(&state)?.decode(buffer);

            // Allocate in arena and create block
            let text = self.arena.alloc_str(&decoded_str);
            let start = self.body_text_len;
            if state.revision_type == Some(super::super::super::annotation::RevisionType::Deletion)
            {
                self.append_revision_text(&state, text, start, start)?;
                buffer.clear();
                return Ok(());
            }
            let block = StyleBlock::new(Cow::Borrowed(text), state.formatting, state.paragraph);
            self.body_text_len = self.body_text_len.checked_add(text.len()).ok_or_else(|| {
                RtfError::MalformedDocument("RTF body text length overflow".to_string())
            })?;
            self.blocks.push(block);
            if !decoded_str.trim().is_empty() {
                self.current_state_mut()?.paragraph_content_started = true;
            }
            self.append_revision_text(&state, text, start, self.body_text_len)?;
        }

        buffer.clear();
        Ok(())
    }

    pub(super) fn effective_text_encoding(&self, state: &State) -> RtfResult<RtfEncoding> {
        let fonts = self.font_table.borrow();
        let Some(font) = fonts.get(state.formatting.font_ref) else {
            return Ok(state.encoding);
        };
        if let Some(page) = font.code_page {
            return Ok(RtfEncoding::from_font_page(page));
        }
        let Some(charset) = font.charset else {
            return Ok(state.encoding);
        };
        if charset == FontCharset::Default {
            return Ok(state.encoding);
        }
        charset
            .page()
            .map(RtfEncoding::from_font_page)
            .ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "unsupported RTF font charset {} for font {}",
                    charset.id(),
                    state.formatting.font_ref
                ))
            })
    }

    pub(super) fn decode_transport_text(&self, text: &str) -> RtfResult<String> {
        let mut bytes = SmallVec::<[u8; 64]>::new();
        append_transport_bytes(&mut bytes, text)?;
        Ok(self.current_state()?.encoding.decode(&bytes).into_owned())
    }

    pub(super) fn append_semantic_text(&mut self, text: &str) -> RtfResult<()> {
        self.finalize_table_before_non_table_body_content(!text.is_empty())?;
        let state = self.prepare_revision_event()?;
        if state.destination != Destination::DocumentBody {
            return Ok(());
        }
        if state.in_table || state.table_nesting_level >= 2 {
            self.append_table_text(text.as_bytes(), state.table_nesting_level)?;
            return Ok(());
        }
        let arena_text = self.arena.alloc_str(text);
        let start = self.body_text_len;
        if state.revision_type == Some(super::super::super::annotation::RevisionType::Deletion) {
            return self.append_revision_text(&state, arena_text, start, start);
        }
        self.body_text_len = self
            .body_text_len
            .checked_add(arena_text.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF body text length overflow".to_string())
            })?;
        self.blocks.push(StyleBlock::new(
            Cow::Borrowed(arena_text),
            state.formatting,
            state.paragraph,
        ));
        self.current_state_mut()?.paragraph_content_started = true;
        self.append_revision_text(&state, arena_text, start, self.body_text_len)
    }

    /// Record a nonrequired (soft) break marker in the body story.
    pub(super) fn record_soft_break(&mut self, kind: crate::SoftBreakKind) -> RtfResult<()> {
        let state = self.current_state()?;
        if state.destination != Destination::DocumentBody || state.in_table {
            return Err(RtfError::MalformedDocument(
                "RTF soft-break controls are supported only in the main body story".to_string(),
            ));
        }
        self.note_options_closed = true;
        self.section_note_options_closed = true;
        self.root_section_format_run = false;
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::SoftBreak(crate::SoftBreak::new(kind, self.body_text_len)),
        ));
        Ok(())
    }

    pub(super) fn record_body_page_break(&mut self) -> RtfResult<()> {
        let state = self.current_state()?.clone();
        if state.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF page is not permitted in this destination".to_string(),
            ));
        }
        if state.table_nesting_level >= 2 {
            let builder = self.ensure_nested_builder(state.table_nesting_level)?;
            builder
                .cell_story_events
                .push(crate::CellStoryEvent::PageBreak(crate::PageBreak::new(
                    builder.cell_text.len(),
                )));
        } else if state.in_table {
            self.current_cell_story_events
                .push(crate::CellStoryEvent::PageBreak(crate::PageBreak::new(
                    self.current_cell_text.len(),
                )));
        } else {
            self.note_options_closed = true;
            self.section_note_options_closed = true;
            self.root_section_format_run = false;
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::PageBreak(crate::PageBreak::new(self.body_text_len)),
            ));
        }
        Ok(())
    }

    pub(super) fn record_body_column_break(&mut self) -> RtfResult<()> {
        let state = self.current_state()?.clone();
        if state.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF column is not permitted in this destination".to_string(),
            ));
        }
        if state.table_nesting_level >= 2 {
            let builder = self.ensure_nested_builder(state.table_nesting_level)?;
            builder
                .cell_story_events
                .push(crate::CellStoryEvent::ColumnBreak(crate::ColumnBreak::new(
                    builder.cell_text.len(),
                )));
        } else if state.in_table {
            self.current_cell_story_events
                .push(crate::CellStoryEvent::ColumnBreak(crate::ColumnBreak::new(
                    self.current_cell_text.len(),
                )));
        } else {
            self.note_options_closed = true;
            self.section_note_options_closed = true;
            self.root_section_format_run = false;
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::ColumnBreak(crate::ColumnBreak::new(self.body_text_len)),
            ));
        }
        Ok(())
    }

    pub(super) fn append_revision_text(
        &mut self,
        state: &State,
        text: &str,
        start: usize,
        end: usize,
    ) -> RtfResult<()> {
        let Some(revision_type) = state.revision_type else {
            return Ok(());
        };
        let id = state.revision_author_id.ok_or_else(|| {
            RtfError::MalformedDocument("RTF revision text is missing an author index".to_string())
        })?;
        let index = usize::try_from(id).map_err(|_err| {
            RtfError::MalformedDocument("RTF revision author index cannot be negative".to_string())
        })?;
        let author = self.revision_authors.get(index).ok_or_else(|| {
            RtfError::MalformedDocument("RTF revision author index is outside revtbl".to_string())
        })?;
        let author_name = author.name.clone();
        let date = state
            .revision_date
            .map(|value| Cow::Owned(value.to_string()));

        self.revision_text_bytes = self
            .revision_text_bytes
            .checked_add(text.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF aggregate revision text size overflow".to_string())
            })?;
        if self.revision_text_bytes > super::super::super::annotation::MAX_REVISION_TEXT_TOTAL_BYTES
        {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision text exceeds the safety limit".to_string(),
            ));
        }

        let previous_event_revision = state
            .revision_event_id
            .and_then(|event_id| self.revision_event_indices.get(event_id).copied().flatten());
        let event_continues_previous = previous_event_revision
            .is_some_and(|last_index| Some(last_index) == self.revisions.len().checked_sub(1));
        if event_continues_previous
            && let Some(previous) = self.revisions.last_mut()
            && previous.revision_type == revision_type
            && previous.id == id
            && previous.author == author_name
            && previous.date == date
            && previous.range_end == start
            && (revision_type != super::super::super::annotation::RevisionType::Deletion
                || previous.position == start)
        {
            if previous.content.len().saturating_add(text.len())
                > super::super::super::annotation::MAX_REVISION_TEXT_BYTES
            {
                return Err(RtfError::MalformedDocument(
                    "RTF revision text exceeds the safety limit".to_string(),
                ));
            }
            previous.content.to_mut().push_str(text);
            previous.range_end = end;
            if let Some(event_id) = state.revision_event_id {
                *self
                    .revision_event_indices
                    .get_mut(event_id)
                    .ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF revision event state references a missing slot".to_string(),
                        )
                    })? = Some(self.revisions.len() - 1);
            }
            return Ok(());
        }
        if self.revisions.len() >= MAX_REVISIONS {
            return Err(RtfError::MalformedDocument(
                "RTF revision count exceeds the safety limit".to_string(),
            ));
        }
        let revision = super::super::super::annotation::Revision {
            revision_type,
            author: author_name,
            date,
            id,
            content: Cow::Owned(text.to_string()),
            position: start,
            range_end: end,
        };
        revision.validate()?;
        self.revisions.push(revision);
        if let Some(event_id) = state.revision_event_id {
            *self
                .revision_event_indices
                .get_mut(event_id)
                .ok_or_else(|| {
                    RtfError::ParserError(
                        "RTF revision event state references a missing slot".to_string(),
                    )
                })? = Some(self.revisions.len() - 1);
        }
        if (state.in_table || state.table_nesting_level >= 2) && previous_event_revision.is_none() {
            let revision_index = self.revisions.len() - 1;
            let event = match revision_type {
                super::super::super::annotation::RevisionType::Insertion => {
                    crate::CellStoryEvent::RevisionStart(crate::CellStoryReference {
                        index: revision_index,
                        position: start,
                    })
                },
                super::super::super::annotation::RevisionType::Deletion => {
                    crate::CellStoryEvent::RevisionDeletion(crate::CellStoryReference {
                        index: revision_index,
                        position: start,
                    })
                },
                super::super::super::annotation::RevisionType::FormatChange
                | super::super::super::annotation::RevisionType::MovedFrom
                | super::super::super::annotation::RevisionType::MovedTo => {
                    return Err(parser_classification_error());
                },
            };
            self.push_cell_story_event(state.table_nesting_level, event)?;
        }
        Ok(())
    }

    pub(super) fn prepare_revision_event(&mut self) -> RtfResult<State> {
        let mut state = self.current_state()?.clone();
        if let Some(kind) = state.revision_type
            && state.revision_event_id.is_none()
        {
            let id = self.revision_event_indices.len();
            self.revision_event_indices.push(None);
            if !state.in_table && state.table_nesting_level < 2 {
                self.body_story_events.push(match kind {
                    super::super::super::annotation::RevisionType::Insertion => {
                        ParsedBodyStoryEvent::RevisionStart(id)
                    },
                    super::super::super::annotation::RevisionType::Deletion => {
                        ParsedBodyStoryEvent::RevisionDeletion(id)
                    },
                    super::super::super::annotation::RevisionType::FormatChange
                    | super::super::super::annotation::RevisionType::MovedFrom
                    | super::super::super::annotation::RevisionType::MovedTo => {
                        return Err(parser_classification_error());
                    },
                });
            }
            state.revision_event_id = Some(id);
            self.current_state_mut()?.revision_event_id = Some(id);
        }
        Ok(state)
    }

    pub(super) fn current_story_position(&mut self) -> RtfResult<usize> {
        let state = self.current_state()?.clone();
        if state.table_nesting_level >= 2 {
            Ok(self
                .ensure_nested_builder(state.table_nesting_level)?
                .cell_text
                .len())
        } else if state.in_table {
            Ok(self.current_cell_text.len())
        } else {
            Ok(self.body_text_len)
        }
    }

    pub(super) fn push_cell_story_event(
        &mut self,
        raw_level: u8,
        event: crate::CellStoryEvent,
    ) -> RtfResult<()> {
        if raw_level >= 2 {
            self.ensure_nested_builder(raw_level)?
                .cell_story_events
                .push(event);
        } else {
            self.current_cell_story_events.push(event);
        }
        Ok(())
    }

    pub(super) fn record_revision_end(&mut self, id: usize) -> RtfResult<()> {
        let state = self.current_state()?.clone();
        if state.in_table || state.table_nesting_level >= 2 {
            let index = self
                .revision_event_indices
                .get(id)
                .copied()
                .flatten()
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF revision event has no tracked text".to_string(),
                    )
                })?;
            let position = self.current_story_position()?;
            self.push_cell_story_event(
                state.table_nesting_level,
                crate::CellStoryEvent::RevisionEnd(crate::CellStoryReference { index, position }),
            )
        } else {
            self.body_story_events
                .push(ParsedBodyStoryEvent::RevisionEnd(id));
            Ok(())
        }
    }

    pub(super) fn close_revision_at_cell_boundary(&mut self, level: u8) -> RtfResult<()> {
        let state = self.current_state()?.clone();
        if state.revision_type == Some(super::super::super::annotation::RevisionType::Insertion)
            && let Some(id) = state.revision_event_id
            && self
                .revision_event_indices
                .get(id)
                .is_some_and(Option::is_some)
        {
            self.record_revision_end(id)?;
        }
        if (state.in_table || state.table_nesting_level >= level) && state.revision_type.is_some() {
            self.current_state_mut()?.revision_event_id = None;
        }
        Ok(())
    }
}
