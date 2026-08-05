use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_bookmark_destination(&mut self) -> RtfResult<()> {
        self.pos += 1; // ignorable-destination marker
        let is_start = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::BookmarkStart)) => true,
            Some(Token::Control(ControlWord::BookmarkEnd)) => false,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid bookmark destination".into(),
                ));
            },
        };
        self.pos += 1;

        let mut name = String::new();
        let mut first_column = None;
        let mut last_column = None;
        let mut is_public = false;
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
                    name.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::BookmarkFirstColumn(value))) => {
                    first_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkLastColumn(value))) => {
                    last_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkPublic)) => is_public = true,
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    name.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("invalid Unicode bookmark name: {error}"))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0) as usize;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default());
                },
                _ => {},
            }
            self.pos += 1;
            if name.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark name exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let name = name.trim_end_matches(['\r', '\n']).to_string();
        if name.is_empty() {
            return Ok(());
        }

        if is_start {
            if self.next_bookmark_order >= MAX_BOOKMARKS {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark count exceeds the safety limit".to_string(),
                ));
            }
            let bookmark = OpenBookmark {
                name: name.clone(),
                position: self.body_text_len,
                first_column,
                last_column,
                is_public,
                order: self.next_bookmark_order,
            };
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::BookmarkStart(self.next_bookmark_order),
            ));
            self.next_bookmark_order += 1;
            self.open_bookmarks.entry(name).or_default().push(bookmark);
        } else if let Some(open) = self.open_bookmarks.get_mut(&name).and_then(Vec::pop) {
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::BookmarkEnd(open.order),
            ));
            self.bookmark_spans.push(BookmarkSpan {
                bookmark: open,
                end: self.body_text_len,
            });
        }
        Ok(())
    }

    pub(super) fn finalize_bookmarks(&mut self) -> RtfResult<()> {
        for bookmarks in self.open_bookmarks.values_mut() {
            for bookmark in bookmarks.drain(..) {
                self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                    crate::BodyStoryEvent::BookmarkEnd(bookmark.order),
                ));
                self.bookmark_spans.push(BookmarkSpan {
                    bookmark,
                    end: self.body_text_len,
                });
            }
        }
        self.bookmark_spans
            .sort_unstable_by_key(|span| span.bookmark.order);
        if self.bookmark_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.bookmark_spans.drain(..) {
            let content = body.get(span.bookmark.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument("bookmark does not align to body text".to_string())
            })?;
            self.bookmarks.add(super::super::super::bookmark::Bookmark {
                name: Cow::Owned(span.bookmark.name),
                position: span.bookmark.position,
                content: Cow::Owned(content.to_string()),
                first_column: span.bookmark.first_column,
                last_column: span.bookmark.last_column,
                is_public: span.bookmark.is_public,
            });
        }
        Ok(())
    }

    /// Consume one text-like token inside a text-carrying destination.
    ///
    /// Returns `Ok(false)` when the current token is not destination text
    /// (plain text, `\uN` runs, `\ucN`, or control symbols) and must be
    /// handled by the caller.
    pub(super) fn consume_destination_text_token(
        &mut self,
        value: &mut String,
        unicode_skip: &mut usize,
        fallback_skip: &mut usize,
        context: &str,
    ) -> RtfResult<bool> {
        match self.tokens.get(self.pos).cloned() {
            Some(Token::Text(text)) => {
                let skipped = (*fallback_skip).min(text.chars().count());
                *fallback_skip -= skipped;
                let remainder: String = text.chars().skip(skipped).collect();
                value.push_str(&self.decode_transport_text(&remainder)?);
                self.pos += 1;
                Ok(true)
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
                        "invalid Unicode RTF custom XML {context}: {error}"
                    ))
                })?);
                *fallback_skip = unicode_skip.saturating_mul(utf16.len());
                Ok(true)
            },
            Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                *unicode_skip = count.max(0) as usize;
                self.pos += 1;
                Ok(true)
            },
            Some(Token::Control(control)) if control_symbol_text(&control).is_some() => {
                value.push_str(control_symbol_text(&control).unwrap_or_default());
                self.pos += 1;
                Ok(true)
            },
            _ => Ok(false),
        }
    }

    /// Collect the plain-text payload of a custom XML destination group.
    ///
    /// Consumes tokens through the group's closing brace; any grouped,
    /// binary, or active control content is rejected.
    pub(super) fn collect_custom_xml_destination_text(
        &mut self,
        context: &str,
        max_bytes: usize,
    ) -> RtfResult<String> {
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(value.trim_end_matches(['\r', '\n']).to_string());
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        context,
                    )? {
                        return Err(RtfError::MalformedDocument(format!(
                            "RTF custom XML {context} destination contains grouped, binary, or active data"
                        )));
                    }
                    if value.len() > max_bytes {
                        return Err(RtfError::MalformedDocument(format!(
                            "RTF custom XML {context} exceeds the safety limit"
                        )));
                    }
                },
            }
        }
    }

    /// Pair one custom XML attribute name/value with the tag being built.
    pub(super) fn push_custom_xml_attribute(
        attributes: &mut Vec<(String, String)>,
        pending: &mut Option<String>,
        is_name: bool,
        text: String,
    ) -> RtfResult<()> {
        if is_name {
            if pending.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML attribute name has no value".to_string(),
                ));
            }
            if text.is_empty() {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML attribute name cannot be empty".to_string(),
                ));
            }
            if text.len() > crate::custom_xml::MAX_CUSTOM_XML_ATTRIBUTE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML attribute name exceeds the safety limit".to_string(),
                ));
            }
            *pending = Some(text);
            return Ok(());
        }
        let name = pending.take().ok_or_else(|| {
            RtfError::MalformedDocument("RTF custom XML attribute value has no name".to_string())
        })?;
        if attributes.len() >= crate::custom_xml::MAX_CUSTOM_XML_ATTRIBUTES_PER_TAG {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute count exceeds the safety limit".to_string(),
            ));
        }
        if attributes.iter().any(|(existing, _)| *existing == name) {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute names must be unique within a tag".to_string(),
            ));
        }
        attributes.push((name, text));
        Ok(())
    }

    /// Parse an `\xmlopen` or `\xmlclose` destination group.
    ///
    /// The destination text is the tag name (RTF 1.9.1 custom XML markup).
    /// `\xmlopen` may additionally select a namespace with `\xmlnsN` and may
    /// carry nested starred `\xmlattrname`/`\xmlattrvalue` groups.
    pub(super) fn parse_custom_xml_tag_destination(&mut self) -> RtfResult<()> {
        if self
            .states
            .iter()
            .any(|state| !matches!(state.destination, Destination::DocumentBody))
        {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML markup destinations are supported only in the main body story"
                    .to_string(),
            ));
        }
        let is_open = matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::XmlOpen))
        );
        self.pos += 1;
        let mut name = String::new();
        let mut namespace = None;
        let mut attributes: Vec<(String, String)> = Vec::new();
        let mut pending: Option<String> = None;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos).cloned() {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Control(ControlWord::XmlNamespace(value))) if is_open => {
                    if namespace.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF custom XML tag selects multiple namespaces".to_string(),
                        ));
                    }
                    if value <= 0 {
                        return Err(RtfError::MalformedDocument(
                            "RTF custom XML namespace references must be in 1..=2147483647"
                                .to_string(),
                        ));
                    }
                    let id = value as u32;
                    if !self.xml_namespaces.iter().any(|entry| entry.id == id) {
                        return Err(RtfError::MalformedDocument(
                            "RTF custom XML tag references an unknown XML namespace".to_string(),
                        ));
                    }
                    namespace = Some(id);
                    self.pos += 1;
                },
                Some(Token::OpenBrace)
                    if is_open
                        && matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::IgnorableDestination))
                        )
                        && matches!(
                            self.tokens.get(self.pos + 2),
                            Some(Token::Control(
                                ControlWord::XmlAttributeName | ControlWord::XmlAttributeValue
                            ))
                        ) =>
                {
                    let is_attribute_name = matches!(
                        self.tokens.get(self.pos + 2),
                        Some(Token::Control(ControlWord::XmlAttributeName))
                    );
                    self.pos += 3;
                    let text = self.collect_custom_xml_destination_text(
                        "attribute",
                        crate::custom_xml::MAX_CUSTOM_XML_ATTRIBUTE_VALUE_BYTES,
                    )?;
                    Self::push_custom_xml_attribute(
                        &mut attributes,
                        &mut pending,
                        is_attribute_name,
                        text,
                    )?;
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    let context = if is_open {
                        "tag name"
                    } else {
                        "close tag name"
                    };
                    if !self.consume_destination_text_token(
                        &mut name,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        context,
                    )? {
                        return Err(RtfError::MalformedDocument(format!(
                            "RTF custom XML {context} destination contains grouped, binary, or active data"
                        )));
                    }
                    if name.len() > crate::custom_xml::MAX_CUSTOM_XML_NAME_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF custom XML tag name exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        let name = name.trim_end_matches(['\r', '\n']).to_string();
        if name.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML tag name cannot be empty".to_string(),
            ));
        }
        if pending.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute name has no value".to_string(),
            ));
        }
        if self.pending_custom_xml_attribute.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute name has no value".to_string(),
            ));
        }

        if is_open {
            if self.open_custom_xml_tags.len() >= crate::custom_xml::MAX_CUSTOM_XML_DEPTH {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML nesting depth exceeds the safety limit".to_string(),
                ));
            }
            if self.next_custom_xml_order >= crate::custom_xml::MAX_CUSTOM_XML_TAGS {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML tag count exceeds the safety limit".to_string(),
                ));
            }
            self.custom_xml_text_bytes = self
                .custom_xml_text_bytes
                .saturating_add(name.len())
                .saturating_add(
                    attributes
                        .iter()
                        .map(|(name, value)| name.len().saturating_add(value.len()))
                        .sum::<usize>(),
                );
            if self.custom_xml_text_bytes > crate::custom_xml::MAX_CUSTOM_XML_TOTAL_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML aggregate text exceeds the safety limit".to_string(),
                ));
            }
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::CustomXmlOpen(self.next_custom_xml_order),
            ));
            self.open_custom_xml_tags.push(OpenCustomXmlTag {
                name,
                namespace,
                attributes,
                position: self.body_text_len,
                order: self.next_custom_xml_order,
            });
            self.next_custom_xml_order += 1;
            return Ok(());
        }

        let open = self.open_custom_xml_tags.pop().ok_or_else(|| {
            RtfError::MalformedDocument("RTF custom XML close has no matching open".to_string())
        })?;
        if open.name != name {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML close does not match the innermost open tag".to_string(),
            ));
        }
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::CustomXmlClose(open.order),
        ));
        self.custom_xml_spans.push(CustomXmlSpan {
            tag: open,
            end: self.body_text_len,
        });
        Ok(())
    }

    /// Parse a starred sibling `\xmlattrname`/`\xmlattrvalue` destination.
    pub(super) fn parse_custom_xml_attribute_destination(&mut self) -> RtfResult<()> {
        if self
            .states
            .iter()
            .any(|state| !matches!(state.destination, Destination::DocumentBody))
        {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML markup destinations are supported only in the main body story"
                    .to_string(),
            ));
        }
        let is_name = matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::XmlAttributeName))
        );
        self.pos += 2; // ignorable marker and destination control word
        let text = self.collect_custom_xml_destination_text(
            "attribute",
            crate::custom_xml::MAX_CUSTOM_XML_ATTRIBUTE_VALUE_BYTES,
        )?;
        let tag = self.open_custom_xml_tags.last_mut().ok_or_else(|| {
            RtfError::MalformedDocument("RTF custom XML attribute has no open tag".to_string())
        })?;
        Self::push_custom_xml_attribute(
            &mut tag.attributes,
            &mut self.pending_custom_xml_attribute,
            is_name,
            text,
        )
    }

    pub(super) fn finalize_custom_xml_tags(&mut self) -> RtfResult<()> {
        if self.pending_custom_xml_attribute.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute name has no value".to_string(),
            ));
        }
        if let Some(open) = self.open_custom_xml_tags.last() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF custom XML tag '{}' is not closed",
                open.name
            )));
        }
        self.custom_xml_spans
            .sort_unstable_by_key(|span| span.tag.order);
        if self.custom_xml_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.custom_xml_spans.drain(..) {
            let content = body.get(span.tag.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "custom XML tag does not align to body text".to_string(),
                )
            })?;
            let attributes = span
                .tag
                .attributes
                .into_iter()
                .map(|(name, value)| {
                    crate::CustomXmlAttribute::new(Cow::Owned(name), Cow::Owned(value))
                })
                .collect::<RtfResult<Vec<_>>>()?;
            self.custom_xml_tags.push(crate::CustomXmlTag::new(
                Cow::Owned(span.tag.name),
                span.tag.namespace,
                attributes,
                span.tag.position,
                Cow::Owned(content.to_string()),
            )?);
        }
        Ok(())
    }
}
