use super::*;

impl<'a> Parser<'a> {
    /// Whether the group starting at `self.pos` is a custom XML markup
    /// destination (`\xmlopen`, `\xmlclose`, or starred
    /// `\xmlattrname`/`\xmlattrvalue`).
    pub(super) fn is_custom_xml_markup_group(&self) -> bool {
        match self.tokens.get(self.pos + 1) {
            Some(Token::Control(ControlWord::XmlOpen | ControlWord::XmlClose)) => true,
            Some(Token::Control(ControlWord::IgnorableDestination)) => matches!(
                self.tokens.get(self.pos + 2),
                Some(Token::Control(
                    ControlWord::XmlAttributeName | ControlWord::XmlAttributeValue
                ))
            ),
            _ => false,
        }
    }

    /// Reject custom XML markup destinations in non-body stories.
    ///
    /// Custom XML markup is modeled only for the main body story; inside
    /// every other text story (notes, headers/footers, shape text, field
    /// stories) the destinations are rejected rather than silently dropped.
    ///
    /// Expects `self.pos` at the group's opening brace.
    pub(super) fn reject_non_body_custom_xml_markup_group(&self) -> RtfResult<()> {
        if self.is_custom_xml_markup_group() {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML markup destinations are supported only in the main body story"
                    .to_string(),
            ));
        }
        Ok(())
    }

    /// Parse a `\*\protstart`/`\*\protend` destination group.
    ///
    /// The destination text is the opaque hexadecimal identifier pairing the
    /// markers (RTF 1.9.1 Word 2003 document protection); matching open and
    /// close markers are paired by identifier like bookmarks, and unclosed
    /// markers extend to the end of the body story.
    pub(super) fn parse_protection_range_destination(&mut self) -> RtfResult<()> {
        self.pos += 1; // ignorable-destination marker
        let is_start = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::ProtectionRangeStart)) => true,
            Some(Token::Control(ControlWord::ProtectionRangeEnd)) => false,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF protection-range destination".to_string(),
                ));
            },
        };
        self.pos += 1;
        let mut id = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut id,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        "protection-range identifier",
                    )? {
                        return Err(RtfError::MalformedDocument(
                            "RTF protection-range destination contains grouped, binary, or active data"
                                .to_string(),
                        ));
                    }
                    if id.len() > crate::protection_range::MAX_PROTECTION_RANGE_ID_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF protection-range identifier exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        let id = id.trim_end_matches(['\r', '\n']).to_string();
        crate::ProtectionRange::new(Cow::Borrowed(id.as_str()), 0, Cow::Borrowed(""))?;

        if is_start {
            if self.next_protection_range_order >= crate::protection_range::MAX_PROTECTION_RANGES {
                return Err(RtfError::MalformedDocument(
                    "RTF protection-range count exceeds the safety limit".to_string(),
                ));
            }
            let range = OpenProtectionRange {
                id: id.clone(),
                position: self.body_text_len,
                order: self.next_protection_range_order,
            };
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::ProtectionRangeStart(self.next_protection_range_order),
            ));
            self.next_protection_range_order += 1;
            self.open_protection_ranges
                .entry(id)
                .or_default()
                .push(range);
        } else if let Some(open) = self.open_protection_ranges.get_mut(&id).and_then(Vec::pop) {
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::ProtectionRangeEnd(open.order),
            ));
            self.protection_range_spans.push(ProtectionRangeSpan {
                range: open,
                end: self.body_text_len,
            });
        }
        Ok(())
    }

    pub(super) fn finalize_protection_ranges(&mut self) -> RtfResult<()> {
        for ranges in self.open_protection_ranges.values_mut() {
            for range in ranges.drain(..) {
                self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                    crate::BodyStoryEvent::ProtectionRangeEnd(range.order),
                ));
                self.protection_range_spans.push(ProtectionRangeSpan {
                    range,
                    end: self.body_text_len,
                });
            }
        }
        self.protection_range_spans
            .sort_unstable_by_key(|span| span.range.order);
        if self.protection_range_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.protection_range_spans.drain(..) {
            let content = body.get(span.range.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "protection range does not align to body text".to_string(),
                )
            })?;
            self.protection_ranges.push(crate::ProtectionRange::new(
                Cow::Owned(span.range.id),
                span.range.position,
                Cow::Owned(content.to_string()),
            )?);
        }
        Ok(())
    }

    /// Record an `\ebcstart`/`\ebcend` editable-region boundary mark.
    ///
    /// The marks carry no identifier and therefore pair positionally: each
    /// `\ebcend` closes the innermost open `\ebcstart`.
    pub(super) fn record_editable_region_boundary(&mut self, is_start: bool) -> RtfResult<()> {
        let state = self.current_state()?;
        if state.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF editable-region marks are supported only in the main body story".to_string(),
            ));
        }
        if is_start {
            if self.next_editable_region_order >= crate::editable_region::MAX_EDITABLE_REGIONS {
                return Err(RtfError::MalformedDocument(
                    "RTF editable-region count exceeds the safety limit".to_string(),
                ));
            }
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::EditableRegionStart(self.next_editable_region_order),
            ));
            self.open_editable_regions.push(OpenEditableRegion {
                position: self.body_text_len,
                order: self.next_editable_region_order,
            });
            self.next_editable_region_order += 1;
            return Ok(());
        }
        let Some(open) = self.open_editable_regions.pop() else {
            return Err(RtfError::MalformedDocument(
                "RTF ebcend has no matching ebcstart".to_string(),
            ));
        };
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::EditableRegionEnd(open.order),
        ));
        self.editable_region_spans.push(EditableRegionSpan {
            position: open.position,
            order: open.order,
            end: self.body_text_len,
        });
        Ok(())
    }

    pub(super) fn finalize_editable_regions(&mut self) -> RtfResult<()> {
        if !self.open_editable_regions.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF ebcstart has no matching ebcend".to_string(),
            ));
        }
        self.editable_region_spans
            .sort_unstable_by_key(|span| span.order);
        if self.editable_region_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.editable_region_spans.drain(..) {
            let content = body.get(span.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "editable region does not align to body text".to_string(),
                )
            })?;
            self.editable_regions.push(crate::EditableRegion::new(
                span.position,
                Cow::Owned(content.to_string()),
            )?);
        }
        Ok(())
    }

    /// Parse a starred `\*\fchars`/`\*\lchars` kinsoku destination group.
    ///
    /// Expects `self.pos` at the ignorable-destination marker and consumes
    /// tokens through the group's closing brace.
    pub(super) fn parse_kinsoku_destination(&mut self, following: bool) -> RtfResult<Cow<'a, str>> {
        self.pos += 2; // ignorable marker and destination control word
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    // Word writes {\ucN\uN } encoding-switch groups inside the
                    // kinsoku destinations of CJK documents.
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::UnicodeSkip(_)))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF kinsoku destination contains an unsupported group".to_string(),
                        ));
                    }
                    self.pos += 1;
                    loop {
                        match self.tokens.get(self.pos) {
                            Some(Token::CloseBrace) => {
                                self.pos += 1;
                                break;
                            },
                            None => return Err(RtfError::UnexpectedEof),
                            _ => {
                                if !self.consume_destination_text_token(
                                    &mut value,
                                    &mut unicode_skip,
                                    &mut fallback_skip,
                                    "kinsoku character set",
                                )? {
                                    return Err(RtfError::MalformedDocument(
                                        "RTF kinsoku encoding-switch group contains grouped, binary, or active data"
                                            .to_string(),
                                    ));
                                }
                            },
                        }
                    }
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        "kinsoku character set",
                    )? {
                        return Err(RtfError::MalformedDocument(
                            "RTF kinsoku destination contains grouped, binary, or active data"
                                .to_string(),
                        ));
                    }
                    if value.len() > crate::kinsoku::MAX_KINSOKU_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF kinsoku character set exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        let value: String = value
            .chars()
            .filter(|c| !matches!(c, '\r' | '\n'))
            .collect();
        crate::DocumentKinsoku::validate_characters(
            if following { "following" } else { "leading" },
            &value,
        )?;
        Ok(Cow::Owned(value))
    }

    pub(super) fn parse_ignorable_text_destination(&mut self) -> RtfResult<String> {
        self.pos += 2; // ignorable marker and destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    value.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid Unicode annotation metadata: {error}"
                        ))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0) as usize;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                _ => {},
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation destination exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }
}
