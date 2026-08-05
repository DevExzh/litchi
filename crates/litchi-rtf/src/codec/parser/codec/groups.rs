use super::*;

impl<'a> Parser<'a> {
    /// Parse a group (content between braces).
    pub(super) fn parse_group(&mut self) -> RtfResult<()> {
        self.expect_token(Token::OpenBrace)?;

        // Group parsing recurses, so refuse pathological nesting before it can
        // exhaust the call stack. Reporting a typed error keeps a hostile file
        // recoverable instead of aborting the process.
        if self.states.len() >= MAX_GROUP_NESTING_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF group nesting depth exceeds the safety limit".to_string(),
            ));
        }

        if self.states.len() == 2 {
            self.root_section_format_run = false;
        }
        let starts_visible_section_format = self
            .states
            .last()
            .is_some_and(|state| state.destination == Destination::DocumentBody)
            && matches!(
                self.tokens.get(self.pos),
                Some(Token::Control(ControlWord::SectionDefault))
            );

        // Push new state (inherit from parent)
        if let Some(current) = self.states.last() {
            self.states.push(current.clone());
        } else {
            self.states.push(State::default());
        }
        if starts_visible_section_format {
            self.current_state_mut()?.visible_section_format = true;
        }

        let nested_destination = match (self.tokens.get(self.pos), self.tokens.get(self.pos + 1)) {
            (Some(Token::Control(ControlWord::NestedTableProperties(param))), _) => {
                Some((true, *param, 1))
            },
            (Some(Token::Control(ControlWord::NoNestedTables(param))), _) => {
                Some((false, *param, 1))
            },
            (
                Some(Token::Control(ControlWord::IgnorableDestination)),
                Some(Token::Control(ControlWord::NestedTableProperties(param))),
            ) => Some((true, *param, 2)),
            (
                Some(Token::Control(ControlWord::IgnorableDestination)),
                Some(Token::Control(ControlWord::NoNestedTables(param))),
            ) => Some((false, *param, 2)),
            _ => None,
        };
        if let Some((properties, param, consumed)) = nested_destination {
            require_parameterless(
                param,
                if properties {
                    "nesttableprops"
                } else {
                    "nonesttables"
                },
            )?;
            self.pos += consumed;
            if properties {
                self.current_state_mut()?.destination = Destination::NestedTableProperties;
                self.parse_content()?;
            } else {
                self.current_state_mut()?.destination = Destination::Other;
                self.skip_until_close_brace()?;
            }
            self.states.pop();
            return Ok(());
        }

        if self.current_state()?.revision_type.is_some()
            && matches!(
                self.tokens.get(self.pos),
                Some(Token::Control(
                    ControlWord::IgnorableDestination
                        | ControlWord::UserProperties
                        | ControlWord::IndexEntry
                        | ControlWord::TableOfContentsEntry
                        | ControlWord::TableOfContentsEntryNoPage
                        | ControlWord::FontTable
                        | ControlWord::ColorTable
                        | ControlWord::StyleSheet
                        | ControlWord::ListTable
                        | ControlWord::ListOverrideTable
                        | ControlWord::RevisionTable
                        | ControlWord::Info
                        | ControlWord::Shape(_)
                        | ControlWord::ShapeGroup(_)
                        | ControlWord::Picture
                        | ControlWord::Object
                        | ControlWord::Result
                        | ControlWord::Field
                        | ControlWord::Header
                        | ControlWord::HeaderFirst
                        | ControlWord::HeaderLeft
                        | ControlWord::HeaderRight
                        | ControlWord::Footer
                        | ControlWord::FooterFirst
                        | ControlWord::FooterLeft
                        | ControlWord::FooterRight
                        | ControlWord::Footnote
                        | ControlWord::Endnote
                ))
            )
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision text cannot contain active or external destinations".to_string(),
            ));
        }

        // Check if this is a special group (header, destination, etc.)
        if self.dispatch_group_destination()? {
            return Ok(());
        }

        // Parse group content. A scoped revision marker without any inert text
        // is an orphan rather than an empty tracked change.
        let revision_text_bytes_before = self.revision_text_bytes;
        self.parse_content()?;
        if self.current_state()?.revision_type.is_some()
            && self.revision_text_bytes == revision_text_bytes_before
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision marker has no tracked text".to_string(),
            ));
        }
        if self.current_state()?.revision_type
            == Some(super::super::super::annotation::RevisionType::Insertion)
            && self.current_state()?.revision_event_id
                != self
                    .states
                    .get(self.states.len().saturating_sub(2))
                    .and_then(|state| state.revision_event_id)
            && let Some(id) = self.current_state()?.revision_event_id
        {
            self.record_revision_end(id)?;
        }

        Self::validate_drop_cap_state(self.current_state()?, "paragraph group")?;

        // Pop state
        self.states.pop();

        Ok(())
    }

    pub(super) fn parse_unicode_alternate_group(&mut self) -> RtfResult<()> {
        self.pos += 1; // upr
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF upr lacks its ANSI fallback group".to_string(),
            ));
        }
        // A Unicode-aware reader must ignore the first (ANSI) representation.
        self.skip_group()?;
        if !matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::UnicodeAlternateDestination),
            ])
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF upr lacks its starred ud destination".to_string(),
            ));
        }
        self.pos += 3;
        self.unicode_alternate_depth = self
            .unicode_alternate_depth
            .checked_add(1)
            .ok_or_else(|| RtfError::MalformedDocument("RTF upr nesting overflow".to_string()))?;
        if self.unicode_alternate_depth > 8 {
            return Err(RtfError::MalformedDocument(
                "RTF upr nesting exceeds the safety limit".to_string(),
            ));
        }
        let parsed = self.parse_content();
        self.unicode_alternate_depth -= 1;
        parsed?;
        self.expect_token(Token::CloseBrace)?; // outer upr group
        Ok(())
    }

    pub(super) fn parse_note_separator_destination(
        &mut self,
        kind: crate::NoteSeparatorKind,
    ) -> RtfResult<crate::NoteSeparator<'a>> {
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF note separators must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and destination
        self.current_note_separator_active = true;
        self.current_note_separator_elements.clear();
        self.current_note_separator_drawings = DrawingStoryCapture::default();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        self.parse_note_separator_elements(&mut unicode_skip, 0)?;
        self.current_note_separator_active = false;
        let drawings = std::mem::take(&mut self.current_note_separator_drawings);
        let separator = crate::NoteSeparator {
            kind,
            elements: std::mem::take(&mut self.current_note_separator_elements),
            shapes: drawings.shapes,
            shape_groups: drawings.shape_groups,
        };
        separator.validate()?;
        Ok(separator)
    }

    pub(super) fn parse_note_separator_elements(
        &mut self,
        unicode_skip: &mut i32,
        depth: usize,
    ) -> RtfResult<()> {
        if depth > 16 {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator nesting exceeds the safety limit".to_string(),
            ));
        }
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => {
                    if self.is_root_drawing_group() {
                        self.parse_group()?;
                        continue;
                    }
                    let direct = self.tokens.get(self.pos + 1);
                    let starred = self.tokens.get(self.pos + 2);
                    if matches!(
                        direct,
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::Shape(_)
                                | ControlWord::ShapeGroup(_)
                                | ControlWord::Footnote
                                | ControlWord::Endnote
                        ))
                    ) || (matches!(
                        direct,
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) && matches!(
                        starred,
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::Shape(_)
                                | ControlWord::ShapeGroup(_)
                        ))
                    )) {
                        return Err(RtfError::MalformedDocument(
                            "RTF note separator cannot contain fields, objects, pictures, or active destinations".to_string(),
                        ));
                    }
                    self.pos += 1;
                    self.parse_note_separator_elements(unicode_skip, depth + 1)?;
                    continue;
                },
                Some(Token::Text(text)) => {
                    let decoded = self.decode_transport_text(text)?;
                    self.push_note_separator_text(decoded);
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, *unicode_skip)?;
                    self.push_note_separator_text(decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    *unicode_skip = (*value).max(0)
                },
                Some(Token::Control(ControlWord::NoteSeparatorCharacter)) => self
                    .current_note_separator_elements
                    .push(crate::NoteSeparatorElement::SeparatorMark),
                Some(Token::Control(ControlWord::NoteContinuationSeparatorCharacter)) => self
                    .current_note_separator_elements
                    .push(crate::NoteSeparatorElement::ContinuationSeparatorMark),
                Some(Token::Control(ControlWord::Par)) => {
                    self.current_note_separator_elements
                        .push(crate::NoteSeparatorElement::ParagraphBreak);
                    self.current_note_separator_drawings.story_offset += 1;
                },
                Some(Token::Control(ControlWord::Line)) => {
                    self.current_note_separator_elements
                        .push(crate::NoteSeparatorElement::LineBreak);
                    self.current_note_separator_drawings.story_offset += 1;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    self.push_note_separator_text("\t".to_string())
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => self
                    .push_note_separator_text(
                        control_symbol_text(control).unwrap_or_default().to_string(),
                    ),
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note separator cannot contain binary data".to_string(),
                    ));
                },
                Some(Token::Control(_)) => {}, // formatting is inert
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if self.current_note_separator_elements.len()
                > crate::note_separator::MAX_NOTE_SEPARATOR_ELEMENTS
            {
                return Err(RtfError::MalformedDocument(
                    "RTF note separator contains too many elements".to_string(),
                ));
            }
            let text_bytes = self
                .current_note_separator_elements
                .iter()
                .map(|element| match element {
                    crate::NoteSeparatorElement::Text(text) => text.len(),
                    _ => 0,
                })
                .sum::<usize>();
            if text_bytes > crate::note_separator::MAX_NOTE_SEPARATOR_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF note-separator text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn push_note_separator_text(&mut self, text: String) {
        if text.is_empty() {
            return;
        }
        self.current_note_separator_drawings.story_offset += text.len();
        if let Some(crate::NoteSeparatorElement::Text(existing)) =
            self.current_note_separator_elements.last_mut()
        {
            existing.to_mut().push_str(&text);
        } else {
            self.current_note_separator_elements
                .push(crate::NoteSeparatorElement::Text(Cow::Owned(text)));
        }
    }
}
