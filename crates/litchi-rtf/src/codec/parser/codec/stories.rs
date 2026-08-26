#![allow(
    clippy::shadow_same,
    clippy::shadow_unrelated,
    reason = "decoding steps deliberately rebind a working value as it is refined through the parse pipeline"
)]
use super::{
    ControlWord, Cow, Formatting, MAX_STORY_GROUP_DEPTH, MAX_TEXT_INTERMEDIATE_BYTES,
    ParsedBodyStoryEvent, Parser, RtfEncoding, RtfError, RtfResult, SmallVec, State, Token,
    control_symbol_text, is_section_control, require_parameterless,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct StoryBaseline {
    ordinary: crate::CharacterBaseline,
    associated: Option<crate::AssociatedCharacterBaseline>,
}

fn effective_story_baseline(formatting: Formatting) -> StoryBaseline {
    let ordinary = match formatting.character_positioning.baseline {
        crate::CharacterBaseline::Normal if formatting.superscript => {
            crate::CharacterBaseline::Superscript
        },
        crate::CharacterBaseline::Normal if formatting.subscript => {
            crate::CharacterBaseline::Subscript
        },
        baseline => baseline,
    };
    StoryBaseline {
        ordinary,
        associated: formatting.associated.baseline,
    }
}

fn observe_story_baseline(
    observed: &mut Option<StoryBaseline>,
    formatting: Formatting,
    owner: &'static str,
) -> RtfResult<()> {
    let baseline = effective_story_baseline(formatting);
    if observed.is_some_and(|existing| existing != baseline) {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {owner} has mixed character baseline formatting that cannot be represented losslessly"
        )));
    }
    *observed = Some(baseline);
    Ok(())
}

fn retain_story_baseline(
    mut formatting: Formatting,
    observed: Option<StoryBaseline>,
) -> Formatting {
    if let Some(baseline) = observed {
        formatting.character_positioning.baseline = baseline.ordinary;
        formatting.superscript = matches!(baseline.ordinary, crate::CharacterBaseline::Superscript);
        formatting.subscript = matches!(baseline.ordinary, crate::CharacterBaseline::Subscript);
        formatting.associated.baseline = baseline.associated;
    }
    formatting
}

impl<'a> Parser<'a> {
    /// Parse header or footer content.
    pub(super) fn parse_header_footer_content(&mut self) -> RtfResult<()> {
        let hf_type = self
            .current_hf_type
            .ok_or_else(|| RtfError::MalformedDocument("Header/footer type not set".to_string()))?;

        let mut hf = super::super::super::section::HeaderFooter::new(hf_type);
        self.current_hf_shapes.clear();
        self.current_hf_shape_groups.clear();
        self.current_hf_drawing_order.clear();
        self.current_hf_story_events.clear();
        self.current_hf_story_offset = 0;
        let mut text_buffer = SmallVec::<[u8; 256]>::new();
        let mut paragraph_baseline = None;
        let mut pending_paragraph_break = false;
        let default_state = State::default();
        let mut inert_section_format = false;

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::super::super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                retain_story_baseline(state.formatting, paragraph_baseline),
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    inert_section_format = false;
                    self.parse_header_footer_group(
                        &mut hf,
                        &mut text_buffer,
                        &mut pending_paragraph_break,
                        &mut paragraph_baseline,
                    )?;
                },
                Token::Control(control @ (ControlWord::Par | ControlWord::Line)) => {
                    let paragraph_break = matches!(control, ControlWord::Par);
                    inert_section_format = false;
                    self.pos += 1;
                    if pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                    }
                    let text = std::str::from_utf8(&text_buffer).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid UTF-8 in header/footer story: {error}"
                        ))
                    })?;
                    let state = self.current_state().ok().unwrap_or(&default_state);
                    let text_alloc = self.arena.alloc_str(text);
                    hf.add_paragraph(super::super::super::section::HeaderFooterParagraph::new(
                        Cow::Borrowed(text_alloc),
                        retain_story_baseline(state.formatting, paragraph_baseline),
                        state.paragraph,
                    ));
                    text_buffer.clear();
                    if paragraph_break {
                        paragraph_baseline = None;
                    }
                    pending_paragraph_break = true;
                },
                Token::Control(ControlWord::Page(param)) => {
                    require_parameterless(*param, "page")?;
                    self.pos += 1;
                    if pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        pending_paragraph_break = false;
                    }
                    self.current_hf_story_events
                        .push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                            self.current_hf_story_offset,
                        )));
                },
                Token::Control(ControlWord::Tab) => {
                    inert_section_format = false;
                    self.pos += 1;
                    if pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        pending_paragraph_break = false;
                    }
                    observe_story_baseline(
                        &mut paragraph_baseline,
                        self.current_state()
                            .ok()
                            .unwrap_or(&default_state)
                            .formatting,
                        "header/footer paragraph",
                    )?;
                    text_buffer.push(b'\t');
                    self.current_hf_story_offset = self.current_hf_story_offset.saturating_add(1);
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    inert_section_format = false;
                    if pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        pending_paragraph_break = false;
                    }
                    let decoded =
                        self.parse_destination_unicode_sequence_for_current_font(*code)?;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            &mut paragraph_baseline,
                            self.current_state()
                                .ok()
                                .unwrap_or(&default_state)
                                .formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(decoded.as_bytes());
                    self.current_hf_story_offset =
                        self.current_hf_story_offset.saturating_add(decoded.len());
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    inert_section_format = false;
                    self.pos += 1;
                    if pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        pending_paragraph_break = false;
                    }
                    let value = control_symbol_text(control).unwrap_or_default();
                    if !value.is_empty() {
                        observe_story_baseline(
                            &mut paragraph_baseline,
                            self.current_state()
                                .ok()
                                .unwrap_or(&default_state)
                                .formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(value.as_bytes());
                    self.current_hf_story_offset = self
                        .current_hf_story_offset
                        .saturating_add(control_symbol_text(control).unwrap_or_default().len());
                },
                Token::Control(ControlWord::SectionDefault) => {
                    self.pos += 1;
                    inert_section_format = true;
                },
                Token::Control(ControlWord::SectionBreak) if inert_section_format => {
                    self.pos += 1;
                    inert_section_format = false;
                },
                Token::Control(control) if inert_section_format && is_section_control(control) => {
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control(token)?;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text_for_current_font(text)?;
                    self.pos += 1;
                    if !decoded.is_empty() {
                        inert_section_format = false;
                        if pending_paragraph_break {
                            self.current_hf_story_offset =
                                self.current_hf_story_offset.saturating_add(1);
                            pending_paragraph_break = false;
                        }
                        observe_story_baseline(
                            &mut paragraph_baseline,
                            self.current_state()
                                .ok()
                                .unwrap_or(&default_state)
                                .formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(decoded.as_bytes());
                    self.current_hf_story_offset =
                        self.current_hf_story_offset.saturating_add(decoded.len());
                },
                Token::Binary(_) => {
                    self.pos += 1;
                },
            }
        }

        while hf.text().len() < self.current_hf_story_offset {
            hf.add_paragraph(super::super::super::section::HeaderFooterParagraph::new(
                Cow::Borrowed(""),
                default_state.formatting,
                default_state.paragraph,
            ));
        }
        hf.shapes = std::mem::take(&mut self.current_hf_shapes);
        hf.shape_groups = std::mem::take(&mut self.current_hf_shape_groups);
        hf.drawing_order = std::mem::take(&mut self.current_hf_drawing_order);
        hf.story_events = std::mem::take(&mut self.current_hf_story_events);
        crate::field::validate_story_events(
            &hf.text(),
            &hf.shapes,
            &hf.shape_groups,
            &hf.drawing_order,
            &hf.story_events,
            "header/footer",
        )?;

        // Headers and footers attach to the section currently being defined.
        if !self.section_properties_active {
            self.begin_section()?;
        }
        self.sections
            .last_mut()
            .ok_or_else(|| RtfError::MalformedDocument("no active RTF section".to_string()))?
            .add_header_footer(hf);

        self.current_hf_type = None;
        Ok(())
    }

    pub(super) fn parse_header_footer_group(
        &mut self,
        hf: &mut super::super::super::section::HeaderFooter<'a>,
        text_buffer: &mut SmallVec<[u8; 256]>,
        pending_paragraph_break: &mut bool,
        paragraph_baseline: &mut Option<StoryBaseline>,
    ) -> RtfResult<()> {
        self.reject_non_body_custom_xml_markup_group()?;
        if self.is_root_drawing_group() {
            if *pending_paragraph_break {
                self.current_hf_story_offset = self.current_hf_story_offset.saturating_add(1);
                *pending_paragraph_break = false;
            }
            return self.parse_group();
        }
        if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::Field))
        ) {
            if *pending_paragraph_break {
                self.current_hf_story_offset = self.current_hf_story_offset.saturating_add(1);
                *pending_paragraph_break = false;
            }
            self.pos += 1;
            self.parse_field()?;
            return self.skip_until_close_brace();
        }
        if matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::LegacyDrawingObject),
            ])
        ) {
            return self.parse_group();
        }
        if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
            return self.preserve_unknown_destination();
        }
        let state = self.current_state()?.clone();
        self.states.push(state);
        self.pos += 1;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    return Ok(());
                },
                Some(Token::OpenBrace) => {
                    self.parse_header_footer_group(
                        hf,
                        text_buffer,
                        pending_paragraph_break,
                        paragraph_baseline,
                    )?;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    self.pos += 1;
                    if *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        *pending_paragraph_break = false;
                    }
                    observe_story_baseline(
                        paragraph_baseline,
                        self.current_state()?.formatting,
                        "header/footer paragraph",
                    )?;
                    text_buffer.push(b'\t');
                    self.current_hf_story_offset += 1;
                },
                Some(Token::Control(control @ (ControlWord::Par | ControlWord::Line))) => {
                    let paragraph_break = matches!(control, ControlWord::Par);
                    self.pos += 1;
                    if *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                    }
                    let text = std::str::from_utf8(text_buffer).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid UTF-8 in header/footer story: {error}"
                        ))
                    })?;
                    let state = self.current_state()?.clone();
                    let text_alloc = self.arena.alloc_str(text);
                    hf.add_paragraph(super::super::super::section::HeaderFooterParagraph::new(
                        Cow::Borrowed(text_alloc),
                        retain_story_baseline(state.formatting, *paragraph_baseline),
                        state.paragraph,
                    ));
                    text_buffer.clear();
                    if paragraph_break {
                        *paragraph_baseline = None;
                    }
                    *pending_paragraph_break = true;
                },
                Some(Token::Control(ControlWord::Page(param))) => {
                    require_parameterless(*param, "page")?;
                    self.pos += 1;
                    if *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        *pending_paragraph_break = false;
                    }
                    self.current_hf_story_events
                        .push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                            self.current_hf_story_offset,
                        )));
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    if *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        *pending_paragraph_break = false;
                    }
                    let code = *code;
                    let decoded = self.parse_destination_unicode_sequence_for_current_font(code)?;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            paragraph_baseline,
                            self.current_state()?.formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(decoded.as_bytes());
                    self.current_hf_story_offset += decoded.len();
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    if *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        *pending_paragraph_break = false;
                    }
                    let value = control_symbol_text(control).unwrap_or_default();
                    if !value.is_empty() {
                        observe_story_baseline(
                            paragraph_baseline,
                            self.current_state()?.formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(value.as_bytes());
                    self.current_hf_story_offset += value.len();
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unknown(_, _))) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control(token)?;
                },
                Some(Token::Control(control)) => {
                    let control = *control;
                    self.pos += 1;
                    self.apply_control_word(&control)?;
                },
                Some(Token::Text(text)) => {
                    let decoded = self.decode_transport_text_for_current_font(text)?;
                    self.pos += 1;
                    if !decoded.is_empty() && *pending_paragraph_break {
                        self.current_hf_story_offset =
                            self.current_hf_story_offset.saturating_add(1);
                        *pending_paragraph_break = false;
                    }
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            paragraph_baseline,
                            self.current_state()?.formatting,
                            "header/footer paragraph",
                        )?;
                    }
                    text_buffer.extend_from_slice(decoded.as_bytes());
                    self.current_hf_story_offset += decoded.len();
                },
                Some(Token::Binary(_)) => {
                    self.states.pop();
                    return Err(RtfError::MalformedDocument(
                        "RTF header/footer story cannot contain direct binary data".to_string(),
                    ));
                },
                None => {
                    self.states.pop();
                    return Err(RtfError::UnexpectedEof);
                },
            }
        }
    }

    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
    )]
    pub(super) fn parse_destination_unicode_sequence(
        &mut self,
        first_code: i32,
    ) -> RtfResult<String> {
        let encoding = self.current_state()?.encoding;
        self.parse_destination_unicode_sequence_with_encoding(first_code, encoding)
    }

    pub(super) fn parse_destination_unicode_sequence_for_current_font(
        &mut self,
        first_code: i32,
    ) -> RtfResult<String> {
        let encoding = {
            let state = self.current_state()?;
            self.effective_text_encoding(state)?
        };
        self.parse_destination_unicode_sequence_with_encoding(first_code, encoding)
    }

    fn parse_destination_unicode_sequence_with_encoding(
        &mut self,
        first_code: i32,
        encoding: RtfEncoding,
    ) -> RtfResult<String> {
        let skip_count = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        Self::push_bounded_unicode_code(&mut utf16, first_code)?;
        self.pos += 1;
        while let Some(Token::Control(ControlWord::Unicode(code))) = self.tokens.get(self.pos) {
            Self::push_bounded_unicode_code(&mut utf16, *code)?;
            self.pos += 1;
        }

        let mut fallback_skip = skip_count.saturating_mul(utf16.len());
        let mut remainder = String::new();
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::Text(text)) => {
                    let count = text.chars().count();
                    if count <= fallback_skip {
                        fallback_skip -= count;
                    } else {
                        let tail = text.chars().skip(fallback_skip);
                        let additional = tail
                            .clone()
                            .map(char::len_utf8)
                            .try_fold(0usize, |total, length| total.checked_add(length))
                            .ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF Unicode fallback size overflow".to_string(),
                                )
                            })?;
                        let observed =
                            remainder.len().checked_add(additional).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF Unicode fallback size overflow".to_string(),
                                )
                            })?;
                        if observed > MAX_TEXT_INTERMEDIATE_BYTES {
                            return Err(RtfError::LimitExceeded {
                                resource: "RTF Unicode fallback text",
                                observed,
                                limit: MAX_TEXT_INTERMEDIATE_BYTES,
                            });
                        }
                        remainder.try_reserve_exact(additional).map_err(|_error| {
                            RtfError::AllocationFailed {
                                resource: "RTF Unicode fallback text",
                                requested: observed,
                            }
                        })?;
                        remainder.extend(tail);
                        fallback_skip = 0;
                    }
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) | None => break,
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
                Some(
                    Token::OpenBrace | Token::CloseBrace | Token::Binary(_) | Token::Control(_),
                ) => {
                    // Only plain text and the finite control-symbol
                    // vocabulary can be a Unicode fallback character.  Leave
                    // active controls for the owning story parser instead of
                    // consuming them as fallback bytes.
                    break;
                },
            }
        }
        let mut decoded = Self::decode_bounded_utf16(&utf16, "RTF destination Unicode")?;
        let fallback = self.decode_transport_text_strict_with_encoding(&remainder, encoding)?;
        let observed = decoded.len().checked_add(fallback.len()).ok_or_else(|| {
            RtfError::MalformedDocument("RTF Unicode fallback size overflow".to_string())
        })?;
        if observed > MAX_TEXT_INTERMEDIATE_BYTES {
            return Err(RtfError::LimitExceeded {
                resource: "RTF Unicode fallback text",
                observed,
                limit: MAX_TEXT_INTERMEDIATE_BYTES,
            });
        }
        decoded
            .try_reserve_exact(fallback.len())
            .map_err(|_error| RtfError::AllocationFailed {
                resource: "RTF Unicode fallback text",
                requested: observed,
            })?;
        decoded.push_str(&fallback);
        Ok(decoded)
    }

    /// Parse footnote or endnote content.
    pub(super) fn parse_note(&mut self, is_footnote: bool) -> RtfResult<()> {
        if self.notes.len() >= super::super::super::section::MAX_NOTES {
            return Err(RtfError::MalformedDocument(
                "RTF note count exceeds the safety limit".to_string(),
            ));
        }
        self.current_note_buffer.clear();
        self.current_note_shapes.clear();
        self.current_note_shape_groups.clear();
        self.current_note_drawing_order.clear();
        self.current_note_story_events.clear();
        let mut note_baseline = None;
        let mut reference = String::from(if is_footnote { "1" } else { "i" });

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    self.parse_note_group(&mut note_baseline)?;
                },
                Token::Control(ControlWord::FootnoteNumber(n)) => {
                    self.pos += 1;
                    reference = n.to_string();
                },
                Token::Control(ControlWord::Tab) => {
                    self.pos += 1;
                    observe_story_baseline(
                        &mut note_baseline,
                        self.current_state()?.formatting,
                        "note",
                    )?;
                    self.current_note_buffer.push(b'\t');
                },
                Token::Control(ControlWord::Par | ControlWord::Line) => {
                    self.pos += 1;
                    self.current_note_buffer.push(b'\n');
                },
                Token::Control(ControlWord::Page(param)) => {
                    require_parameterless(*param, "page")?;
                    self.pos += 1;
                    self.current_note_story_events
                        .push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                            self.current_note_buffer.len(),
                        )));
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    let decoded =
                        self.parse_destination_unicode_sequence_for_current_font(*code)?;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            &mut note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer
                        .extend_from_slice(decoded.as_bytes());
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    self.pos += 1;
                    let value = control_symbol_text(control).unwrap_or_default();
                    if !value.is_empty() {
                        observe_story_baseline(
                            &mut note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer.extend_from_slice(value.as_bytes());
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control(token)?;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text_for_current_font(text)?;
                    self.pos += 1;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            &mut note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer
                        .extend_from_slice(decoded.as_bytes());
                },
                Token::Binary(_) => {
                    self.pos += 1;
                },
            }
            if self.current_note_buffer.len() > super::super::super::section::MAX_NOTE_BODY_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF note body exceeds the safety limit".to_string(),
                ));
            }
        }

        let content = std::str::from_utf8(&self.current_note_buffer)
            .map_err(|error| RtfError::InvalidUnicode(format!("invalid note Unicode: {error}")))?;
        let content_alloc = self.arena.alloc_str(content);
        let mut note = if is_footnote {
            super::super::super::section::Note::footnote(
                Cow::Owned(reference),
                Cow::Borrowed(content_alloc),
            )
        } else {
            super::super::super::section::Note::endnote(
                Cow::Owned(reference),
                Cow::Borrowed(content_alloc),
            )
        };
        note.position = self.body_text_len;
        note.shapes = std::mem::take(&mut self.current_note_shapes);
        note.shape_groups = std::mem::take(&mut self.current_note_shape_groups);
        note.drawing_order = std::mem::take(&mut self.current_note_drawing_order);
        note.story_events = std::mem::take(&mut self.current_note_story_events);

        if let Ok(state) = self.current_state() {
            note.formatting = retain_story_baseline(state.formatting, note_baseline);
        }

        let aggregate = note.text_bytes().and_then(|initial| {
            self.notes.iter().try_fold(initial, |size, existing| {
                size.checked_add(existing.text_bytes()?)
            })
        });
        if aggregate
            .is_none_or(|size| size > super::super::super::section::MAX_NOTE_TEXT_TOTAL_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF note aggregate text exceeds the safety limit".to_string(),
            ));
        }
        note.validate()?;
        let index = self.notes.len();
        self.notes.push(note);
        self.body_story_events
            .push(ParsedBodyStoryEvent::Resolved(crate::BodyStoryEvent::Note(
                index,
            )));

        Ok(())
    }

    pub(super) fn parse_note_group(
        &mut self,
        note_baseline: &mut Option<StoryBaseline>,
    ) -> RtfResult<()> {
        self.reject_non_body_custom_xml_markup_group()?;
        if self.is_root_drawing_group() {
            return self.parse_group();
        }
        if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::Field))
        ) {
            self.pos += 1;
            self.parse_field()?;
            return self.skip_until_close_brace();
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF note nested-group parser is not at an opening brace".to_string(),
            ));
        }
        if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
            return self.preserve_unknown_destination();
        }
        if self.states.len() >= MAX_STORY_GROUP_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF note-story group nesting exceeds the safety limit".to_string(),
            ));
        }
        let state = self.current_state()?.clone();
        self.states.push(state);
        self.pos += 1;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_note_group(note_baseline)?,
                Some(Token::Control(ControlWord::Tab)) => {
                    self.pos += 1;
                    observe_story_baseline(
                        note_baseline,
                        self.current_state()?.formatting,
                        "note",
                    )?;
                    self.current_note_buffer.push(b'\t');
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                    self.pos += 1;
                    self.current_note_buffer.push(b'\n');
                },
                Some(Token::Control(ControlWord::Page(param))) => {
                    require_parameterless(*param, "page")?;
                    self.pos += 1;
                    self.current_note_story_events
                        .push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                            self.current_note_buffer.len(),
                        )));
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    let code = *code;
                    let decoded = self.parse_destination_unicode_sequence_for_current_font(code)?;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer
                        .extend_from_slice(decoded.as_bytes());
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    let value = control_symbol_text(control).unwrap_or_default();
                    if !value.is_empty() {
                        observe_story_baseline(
                            note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer.extend_from_slice(value.as_bytes());
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Footnote | ControlWord::Endnote)) => {
                    self.states.pop();
                    return Err(RtfError::MalformedDocument(
                        "RTF note story cannot contain a nested note destination".to_string(),
                    ));
                },
                Some(Token::Control(ControlWord::Unknown(_, _))) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control(token)?;
                },
                Some(Token::Control(control)) => {
                    let control = *control;
                    self.pos += 1;
                    self.apply_control_word(&control)?;
                },
                Some(Token::Text(text)) => {
                    let decoded = self.decode_transport_text_for_current_font(text)?;
                    self.pos += 1;
                    if !decoded.is_empty() {
                        observe_story_baseline(
                            note_baseline,
                            self.current_state()?.formatting,
                            "note",
                        )?;
                    }
                    self.current_note_buffer
                        .extend_from_slice(decoded.as_bytes());
                },
                Some(Token::Binary(_)) => {
                    self.states.pop();
                    return Err(RtfError::MalformedDocument(
                        "RTF note story cannot contain direct binary data".to_string(),
                    ));
                },
                None => {
                    self.states.pop();
                    return Err(RtfError::UnexpectedEof);
                },
            }
            if self.current_note_buffer.len() > super::super::super::section::MAX_NOTE_BODY_BYTES {
                self.states.pop();
                return Err(RtfError::MalformedDocument(
                    "RTF note body exceeds the safety limit".to_string(),
                ));
            }
        }
    }
}
