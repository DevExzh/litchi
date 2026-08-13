use super::{
    ControlWord, Cow, Destination, DrawingStoryCapture, InfoTextField, InfoTimeField,
    MAX_GROUP_NESTING_DEPTH, MAX_INFO_TEXT_BYTES, Parser, RtfError, RtfResult, SmallVec, State,
    Token, control_symbol_text,
};
use std::mem::size_of;

impl<'a> Parser<'a> {
    /// Retain only a bounded, content-free marker when a focused destination
    /// parser encounters syntax it cannot safely interpret.
    pub(super) fn mark_unknown_syntax(&mut self) -> RtfResult<()> {
        if self.unknown_syntax_markers >= self.limits.max_opaque_nodes() {
            return Err(RtfError::LimitExceeded {
                resource: "unknown syntax markers",
                observed: self.unknown_syntax_markers.saturating_add(1),
                limit: self.limits.max_opaque_nodes(),
            });
        }
        self.unknown_syntax_markers += 1;
        Ok(())
    }

    pub(super) fn parse_info_text(&mut self, field: InfoTextField) -> RtfResult<()> {
        let duplicate = match field {
            InfoTextField::Title => self.info.title.is_some(),
            InfoTextField::Subject => self.info.subject.is_some(),
            InfoTextField::Author => self.info.author.is_some(),
            InfoTextField::Manager => self.info.manager.is_some(),
            InfoTextField::Company => self.info.company.is_some(),
            InfoTextField::Operator => self.info.operator.is_some(),
            InfoTextField::Category => self.info.category.is_some(),
            InfoTextField::Keywords => self.info.keywords.is_some(),
            InfoTextField::Comment => self.info.comment.is_some(),
            InfoTextField::DocumentComment => self.info.document_comment.is_some(),
            InfoTextField::HyperlinkBase => self.info.hyperlink_base.is_some(),
        };
        if duplicate {
            return Err(RtfError::MalformedDocument(
                "RTF info text destination occurs more than once".to_string(),
            ));
        }
        self.pos += 1; // destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF info nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    depth += 1;
                    self.pos += 1;
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    value.push_str(&self.decode_transport_text(&remainder)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        #[allow(
                            clippy::cast_possible_truncation,
                            clippy::cast_sign_loss,
                            reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
                        )]
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("Invalid info Unicode: {error}"))
                    })?);
                    fallback_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned()
                        as usize
                        * utf16.len();
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = *count;
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(_) => self.pos += 1,
                None => break,
            }
            if value.len() > MAX_INFO_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF info text exceeds the metadata safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let allocated = self.arena.alloc_str(value.trim_end_matches(['\r', '\n']));
        let text_value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTextField::Title => self.info.title = text_value,
            InfoTextField::Subject => self.info.subject = text_value,
            InfoTextField::Author => self.info.author = text_value,
            InfoTextField::Manager => self.info.manager = text_value,
            InfoTextField::Company => self.info.company = text_value,
            InfoTextField::Operator => self.info.operator = text_value,
            InfoTextField::Category => self.info.category = text_value,
            InfoTextField::Keywords => self.info.keywords = text_value,
            InfoTextField::Comment => self.info.comment = text_value,
            InfoTextField::DocumentComment => self.info.document_comment = text_value,
            InfoTextField::HyperlinkBase => self.info.hyperlink_base = text_value,
        }
        Ok(())
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_info_time(&mut self, field: InfoTimeField) -> RtfResult<()> {
        let duplicate = match field {
            InfoTimeField::Creation => self.info.creation_timestamp.is_some(),
            InfoTimeField::Revision => self.info.revision_timestamp.is_some(),
            InfoTimeField::Print => self.info.print_timestamp.is_some(),
            InfoTimeField::Backup => self.info.backup_timestamp.is_some(),
        };
        if duplicate {
            return Err(RtfError::MalformedDocument(
                "RTF info timestamp destination occurs more than once".to_string(),
            ));
        }
        self.pos += 1; // time destination
        let mut timestamp = crate::RtfTimestamp::default();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF info nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    depth += 1;
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Control(control)) => match control {
                    ControlWord::Year(value) => timestamp.year = Some(*value),
                    ControlWord::Month(value) => timestamp.month = Some(*value),
                    ControlWord::Day(value) => timestamp.day = Some(*value),
                    ControlWord::Hour(value) => timestamp.hour = Some(*value),
                    ControlWord::Minute(value) => timestamp.minute = Some(*value),
                    ControlWord::Second(value) => timestamp.second = Some(*value),
                    _ => {},
                },
                _ => {},
            }
            self.pos += 1;
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let serialized = timestamp.legacy_string();
        let allocated = self.arena.alloc_str(&serialized);
        let value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTimeField::Creation => {
                self.info.creation_time = value;
                self.info.creation_timestamp = Some(timestamp);
            },
            InfoTimeField::Revision => {
                self.info.revision_time = value;
                self.info.revision_timestamp = Some(timestamp);
            },
            InfoTimeField::Print => {
                self.info.print_time = value;
                self.info.print_timestamp = Some(timestamp);
            },
            InfoTimeField::Backup => {
                self.info.backup_time = value;
                self.info.backup_timestamp = Some(timestamp);
            },
        }
        Ok(())
    }

    pub(super) fn set_info_number(slot: &mut Option<u32>, value: i32, name: &str) -> RtfResult<()> {
        if slot.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF info numeric control {name} occurs more than once"
            )));
        }
        *slot = Some(u32::try_from(value).map_err(|_err| {
            RtfError::MalformedDocument(format!(
                "RTF info numeric control {name} cannot be negative"
            ))
        })?);
        Ok(())
    }

    pub(super) fn parse_info_password(&mut self) -> RtfResult<()> {
        if self.info.protection.password_hash.is_some() {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF protection password hash".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and password destination
        let value = self.parse_inert_text_group_contents(
            crate::info::PROTECTION_PASSWORD_HASH_BYTES,
            "protection password hash",
        )?;
        self.info.protection.password_hash = Some(Cow::Owned(value));
        self.info.protection.validate()
    }

    pub(super) fn ensure_protection_scope(&self) -> RtfResult<()> {
        if self.states.len() != 2 || self.body_text_len != 0 {
            return Err(RtfError::MalformedDocument(
                "RTF document protection controls must occur in the root header".to_string(),
            ));
        }
        Ok(())
    }

    pub(super) fn set_protection_flag(
        slot: &mut Option<bool>,
        value: Option<i32>,
        name: &str,
    ) -> RtfResult<()> {
        let flag = value.unwrap_or(1);
        Self::set_required_protection_flag(slot, flag, name)
    }

    pub(super) fn set_required_protection_flag(
        slot: &mut Option<bool>,
        value: i32,
        name: &str,
    ) -> RtfResult<()> {
        if slot.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF {name} control"
            )));
        }
        *slot = Some(match value {
            0 => false,
            1 => true,
            _ => {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} parameter must be 0 or 1"
                )));
            },
        });
        Ok(())
    }

    pub(super) fn skip_open_info_group(&mut self) -> RtfResult<()> {
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF skipped info nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    depth += 1;
                },
                Some(Token::CloseBrace) => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }
        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Skip tokens until closing brace.
    pub(super) fn skip_until_close_brace(&mut self) -> RtfResult<()> {
        let mut depth = 1;

        while depth > 0 {
            let Some(token) = self.tokens.get(self.pos) else {
                break;
            };
            match token {
                Token::OpenBrace => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF skipped nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    depth += 1;
                },
                Token::CloseBrace => depth -= 1,
                Token::Control(_) | Token::Text(_) | Token::Binary(_) => {},
            }
            self.pos += 1;
        }

        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Skip an entire group starting from the `OpenBrace` token.
    pub(super) fn skip_group(&mut self) -> RtfResult<()> {
        // Must be positioned at OpenBrace
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Ok(());
        }

        self.pos += 1; // Skip the OpenBrace
        let mut depth = 1;

        while depth > 0 {
            let Some(token) = self.tokens.get(self.pos) else {
                break;
            };
            match token {
                Token::OpenBrace => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF skipped group nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    depth += 1;
                },
                Token::CloseBrace => depth -= 1,
                Token::Control(_) | Token::Text(_) | Token::Binary(_) => {},
            }
            self.pos += 1;
        }

        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    pub(super) fn preserve_unknown_destination(&mut self) -> RtfResult<()> {
        let start = self.pos.checked_sub(1).ok_or_else(|| {
            RtfError::ParserError("RTF destination has no opening group".to_string())
        })?;
        self.skip_until_close_brace()?;
        self.preserve_opaque_tokens(start, self.pos, crate::opaque::Kind::Destination, None)
    }

    pub(super) fn preserve_unknown_control(&mut self, token: usize) -> RtfResult<()> {
        self.preserve_opaque_tokens(
            token,
            token.saturating_add(1),
            crate::opaque::Kind::ControlWord,
            None,
        )
    }

    pub(super) fn preserve_unknown_destination_in(
        &mut self,
        context: crate::opaque::Context,
    ) -> RtfResult<()> {
        let start = self.pos.checked_sub(1).ok_or_else(|| {
            RtfError::ParserError("RTF destination has no opening group".to_string())
        })?;
        self.skip_until_close_brace()?;
        self.preserve_opaque_tokens(
            start,
            self.pos,
            crate::opaque::Kind::Destination,
            Some(context),
        )
    }

    pub(super) fn preserve_unknown_control_in(
        &mut self,
        token: usize,
        context: crate::opaque::Context,
    ) -> RtfResult<()> {
        self.preserve_opaque_tokens(
            token,
            token.saturating_add(1),
            crate::opaque::Kind::ControlWord,
            Some(context),
        )
    }

    fn preserve_opaque_tokens(
        &mut self,
        start_token: usize,
        end_token: usize,
        kind: crate::opaque::Kind,
        forced_context: Option<crate::opaque::Context>,
    ) -> RtfResult<()> {
        let (Some(source), Some(spans)) = (self.source, self.token_spans) else {
            return Ok(());
        };
        let start = spans
            .get(start_token)
            .map(|span| span.start)
            .ok_or_else(|| {
                RtfError::ParserError("RTF opaque source start is unavailable".to_string())
            })?;
        let end_index = end_token
            .checked_sub(1)
            .ok_or_else(|| RtfError::ParserError("RTF opaque source range is empty".to_string()))?;
        let end = spans.get(end_index).map(|span| span.end).ok_or_else(|| {
            RtfError::ParserError("RTF opaque source end is unavailable".to_string())
        })?;
        let fragment = source.get(start..end).ok_or_else(|| {
            RtfError::ParserError("RTF opaque source range is invalid".to_string())
        })?;
        let observed = fragment.chars().count();
        if observed > self.limits.max_opaque_node_bytes() {
            return Err(RtfError::LimitExceeded {
                resource: "opaque node bytes",
                observed,
                limit: self.limits.max_opaque_node_bytes(),
            });
        }
        let node_count = self.opaque_nodes.len().saturating_add(1);
        if node_count > self.limits.max_opaque_nodes() {
            return Err(RtfError::LimitExceeded {
                resource: "opaque nodes",
                observed: node_count,
                limit: self.limits.max_opaque_nodes(),
            });
        }
        let total =
            self.opaque_bytes
                .checked_add(observed)
                .ok_or_else(|| RtfError::LimitExceeded {
                    resource: "opaque bytes",
                    observed: usize::MAX,
                    limit: self.limits.max_total_opaque_bytes(),
                })?;
        if total > self.limits.max_total_opaque_bytes() {
            return Err(RtfError::LimitExceeded {
                resource: "opaque bytes",
                observed: total,
                limit: self.limits.max_total_opaque_bytes(),
            });
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve(observed)
            .map_err(|_err| RtfError::AllocationFailed {
                resource: "opaque node bytes",
                requested: observed,
            })?;
        for character in fragment.chars() {
            let byte = u8::try_from(u32::from(character)).map_err(|_err| {
                RtfError::InvalidUnicode(
                    "RTF opaque transport contains a non-Latin-1 scalar".to_string(),
                )
            })?;
            bytes.push(byte);
        }
        let state = self.current_state()?;
        let anchor = if forced_context.is_none()
            && state.destination == Destination::DocumentBody
            && !state.in_table
            && state.table_nesting_level < 2
        {
            crate::opaque::Anchor::Body(self.body_text_len)
        } else {
            let context = if let Some(context) = forced_context {
                context
            } else if state.in_table
                || state.table_nesting_level >= 2
                || state.destination == Destination::NestedTableProperties
            {
                crate::opaque::Context::Table
            } else {
                match state.destination {
                    Destination::FontTable
                    | Destination::ColorTable
                    | Destination::StyleSheet
                    | Destination::Info => crate::opaque::Context::Metadata,
                    Destination::Header | Destination::Footer => {
                        crate::opaque::Context::HeaderFooter
                    },
                    Destination::FieldInstruction | Destination::FieldResult => {
                        crate::opaque::Context::Field
                    },
                    Destination::Footnote | Destination::Endnote => crate::opaque::Context::Note,
                    Destination::Picture | Destination::Result => crate::opaque::Context::Drawing,
                    Destination::Revision => crate::opaque::Context::Review,
                    Destination::DocumentBody
                    | Destination::NestedTableProperties
                    | Destination::Other => crate::opaque::Context::Other,
                }
            };
            crate::opaque::Anchor::Structural {
                context,
                token: start_token,
                depth: self.states.len(),
            }
        };
        self.opaque_nodes
            .try_reserve(1)
            .map_err(|_err| RtfError::AllocationFailed {
                resource: "opaque nodes",
                requested: node_count.saturating_mul(size_of::<crate::opaque::Node>()),
            })?;
        self.opaque_nodes
            .push(crate::opaque::Node::new(kind, anchor, bytes));
        self.opaque_bytes = total;
        Ok(())
    }

    /// Expect a specific token.
    pub(super) fn expect_token(&mut self, expected: &Token<'_>) -> RtfResult<()> {
        let actual = self.tokens.get(self.pos).ok_or(RtfError::UnexpectedEof)?;
        if actual != expected {
            return Err(RtfError::ParserError(format!(
                "Expected {expected:?}, found {actual:?}"
            )));
        }

        self.pos += 1;
        Ok(())
    }

    /// Get current state (mutable).
    pub(super) fn current_state_mut(&mut self) -> RtfResult<&mut State> {
        self.states
            .last_mut()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    /// Get current state (immutable).
    pub(super) fn current_state(&self) -> RtfResult<&State> {
        self.states
            .last()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    pub(super) fn current_field_drawing_capture(&self) -> RtfResult<&DrawingStoryCapture<'a>> {
        self.field_drawing_captures.last().ok_or_else(|| {
            RtfError::ParserError("No field-result drawing capture available".to_string())
        })
    }

    pub(super) fn current_field_drawing_capture_mut(
        &mut self,
    ) -> RtfResult<&mut DrawingStoryCapture<'a>> {
        self.field_drawing_captures.last_mut().ok_or_else(|| {
            RtfError::ParserError("No field-result drawing capture available".to_string())
        })
    }
}
