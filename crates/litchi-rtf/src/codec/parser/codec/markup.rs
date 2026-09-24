use super::{
    BookmarkSpan, ControlWord, Cow, CustomXmlSpan, Destination, MAX_BOOKMARK_NAME_BYTES,
    MAX_BOOKMARKS, OpenBookmark, OpenCustomXmlTag, ParsedBodyStoryEvent, Parser, RtfError,
    RtfResult, SmallVec, Token, control_symbol_text,
};

fn decode_move_hex(value: &str) -> RtfResult<(u16, u32)> {
    if value.len() != 12 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(RtfError::MalformedDocument(
            "RTF move-bookmark payload must contain exactly 6 hexadecimal bytes".to_string(),
        ));
    }
    let mut bytes = [0_u8; 6];
    for (index, chunk) in value.as_bytes().chunks_exact(2).enumerate() {
        let high = char::from(chunk.first().copied().ok_or_else(|| {
            RtfError::MalformedDocument("invalid RTF move-bookmark payload".to_string())
        })?)
        .to_digit(16)
        .ok_or_else(|| {
            RtfError::MalformedDocument("invalid RTF move-bookmark payload".to_string())
        })?;
        let low = char::from(chunk.get(1).copied().ok_or_else(|| {
            RtfError::MalformedDocument("invalid RTF move-bookmark payload".to_string())
        })?)
        .to_digit(16)
        .ok_or_else(|| {
            RtfError::MalformedDocument("invalid RTF move-bookmark payload".to_string())
        })?;
        *bytes.get_mut(index).ok_or_else(|| {
            RtfError::MalformedDocument("invalid RTF move-bookmark payload".to_string())
        })? = ((high << 4) | low) as u8;
    }
    Ok((
        u16::from_le_bytes([bytes[0], bytes[1]]),
        u32::from_le_bytes([bytes[2], bytes[3], bytes[4], bytes[5]]),
    ))
}

fn require_body_markup(parser: &Parser<'_>, vocabulary: &str) -> RtfResult<()> {
    let state = parser.current_state()?;
    if state.destination != Destination::DocumentBody
        || state.in_table
        || state.table_nesting_level >= 2
        || parser.states.len() < 3
    {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {vocabulary} destinations are supported only in the main body story"
        )));
    }
    Ok(())
}

impl Parser<'_> {
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
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
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
                        #[allow(
                            clippy::cast_possible_truncation,
                            clippy::cast_sign_loss,
                            reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
                        )]
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
                    unicode_skip = (*value).max(0).cast_unsigned() as usize;
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
        let bookmark_name = name.trim_end_matches(['\r', '\n']).to_string();
        if bookmark_name.is_empty() {
            return Ok(());
        }

        if is_start {
            if self.next_bookmark_order >= MAX_BOOKMARKS {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark count exceeds the safety limit".to_string(),
                ));
            }
            let bookmark = OpenBookmark {
                name: bookmark_name.clone(),
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
            self.open_bookmarks
                .entry(bookmark_name)
                .or_default()
                .push(bookmark);
        } else if let Some(open) = self
            .open_bookmarks
            .get_mut(&bookmark_name)
            .and_then(Vec::pop)
        {
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

    /// Parse a starred RTF 1.9.1 SmartTag opening destination in the main body.
    /// The p.153 grammar requires `\xmlnsN` before `\factoidname` and each
    /// attribute group requires `\xmlattrnsN` before its name/value pair.
    pub(super) fn parse_smart_tag_open_destination(&mut self) -> RtfResult<()> {
        require_body_markup(self, "SmartTag")?;
        self.pos += 2; // ignorable marker and xmlopen
        let mut namespace = None;
        let mut name = None;
        let mut attributes = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Text(text)) if text.chars().all(char::is_whitespace) => {
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::XmlNamespace(value))) => {
                    if namespace.is_some() || name.is_some() || !attributes.is_empty() || *value < 0
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlopen requires one leading non-negative xmlnsN"
                                .to_string(),
                        ));
                    }
                    namespace = Some((*value).cast_unsigned());
                    self.pos += 1;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::FactoidName))
                    ) =>
                {
                    if name.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlopen has duplicate factoidname".to_string(),
                        ));
                    }
                    if namespace.is_none() || !attributes.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlopen requires xmlnsN before factoidname".to_string(),
                        ));
                    }
                    self.pos += 2;
                    let value = self.collect_custom_xml_destination_text(
                        "factoid name",
                        crate::smart_tag::MAX_SMART_TAG_NAME_BYTES,
                    )?;
                    name = Some(value.trim().to_string());
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::XmlAttributeGroup))
                    ) =>
                {
                    if namespace.is_none() || name.is_none() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattr requires xmlnsN and factoidname first"
                                .to_string(),
                        ));
                    }
                    if attributes.len() >= crate::smart_tag::MAX_SMART_TAG_ATTRIBUTES_PER_TAG {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag attribute count exceeds the safety limit".to_string(),
                        ));
                    }
                    attributes.push(self.parse_smart_tag_attribute_group()?);
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF SmartTag xmlopen contains unsupported content".to_string(),
                    ));
                },
            }
        }
        let name = name.ok_or_else(|| {
            RtfError::MalformedDocument("RTF SmartTag xmlopen lacks factoidname".to_string())
        })?;
        if namespace.is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag xmlopen lacks required xmlnsN".to_string(),
            ));
        }
        if self.next_smart_tag_order >= crate::smart_tag::MAX_SMART_TAGS {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag count exceeds the safety limit".to_string(),
            ));
        }
        if self.open_smart_tags.len() >= crate::smart_tag::MAX_SMART_TAG_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag nesting depth exceeds the safety limit".to_string(),
            ));
        }
        let order = self.next_smart_tag_order;
        self.next_smart_tag_order = self.next_smart_tag_order.checked_add(1).ok_or_else(|| {
            RtfError::MalformedDocument("RTF SmartTag order overflow".to_string())
        })?;
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::SmartTagOpen(order),
        ));
        self.open_smart_tags.push(super::OpenSmartTag {
            name,
            namespace,
            attributes,
            position: self.body_text_len,
            order,
        });
        Ok(())
    }

    fn parse_smart_tag_attribute_group(&mut self) -> RtfResult<crate::SmartTagAttribute<'static>> {
        self.pos += 2; // opening brace and xmlattr
        while matches!(
            self.tokens.get(self.pos),
            Some(Token::Text(text)) if text.chars().all(char::is_whitespace)
        ) {
            self.pos += 1;
        }
        let namespace = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::XmlAttributeNamespace(value))) => {
                if *value < 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF SmartTag attribute namespace cannot be negative".to_string(),
                    ));
                }
                let namespace = Some((*value).cast_unsigned());
                self.pos += 1;
                namespace
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF SmartTag xmlattr requires xmlattrnsN".to_string(),
                ));
            },
        };
        let mut name = None;
        let mut value = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Text(text)) if text.chars().all(char::is_whitespace) => {
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::XmlAttributeName)) => {
                    if name.is_some() || value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattrname must precede one xmlattrvalue".to_string(),
                        ));
                    }
                    self.pos += 1;
                    name = Some(self.collect_smart_tag_direct_attribute_name()?);
                },
                Some(Token::Control(ControlWord::XmlAttributeValue)) => {
                    if name.is_none() || value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattrvalue must follow xmlattrname exactly once"
                                .to_string(),
                        ));
                    }
                    self.pos += 1;
                    value = Some(self.collect_smart_tag_direct_attribute_value(
                        "attribute value",
                        crate::smart_tag::MAX_SMART_TAG_ATTRIBUTE_VALUE_BYTES,
                    )?);
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::XmlAttributeName))
                    ) =>
                {
                    if name.is_some() || value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattrname must precede one xmlattrvalue".to_string(),
                        ));
                    }
                    self.pos += 2;
                    name = Some(self.collect_custom_xml_destination_text(
                        "attribute name",
                        crate::smart_tag::MAX_SMART_TAG_ATTRIBUTE_NAME_BYTES,
                    )?);
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::XmlAttributeValue))
                    ) =>
                {
                    if name.is_none() || value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattrvalue must follow xmlattrname exactly once"
                                .to_string(),
                        ));
                    }
                    self.pos += 2;
                    value = Some(self.collect_custom_xml_destination_text(
                        "attribute value",
                        crate::smart_tag::MAX_SMART_TAG_ATTRIBUTE_VALUE_BYTES,
                    )?);
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF SmartTag xmlattr contains unsupported content".to_string(),
                    ));
                },
            }
        }
        crate::SmartTagAttribute::new(
            namespace,
            Cow::Owned(name.ok_or_else(|| {
                RtfError::MalformedDocument("RTF SmartTag xmlattr lacks xmlattrname".to_string())
            })?),
            Cow::Owned(value.ok_or_else(|| {
                RtfError::MalformedDocument("RTF SmartTag xmlattr lacks xmlattrvalue".to_string())
            })?),
        )
    }

    /// Collect direct `#PCDATA` after `\\xmlattrname` until the normative
    /// `\\xmlattrvalue` control word.  Nested destination groups remain
    /// accepted for compatibility with the existing writer and fixtures.
    fn collect_smart_tag_direct_attribute_name(&mut self) -> RtfResult<String> {
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::XmlAttributeValue)) => {
                    return Ok(value.trim_end_matches(['\r', '\n']).to_string());
                },
                Some(Token::CloseBrace) | None => {
                    return Err(RtfError::MalformedDocument(
                        "RTF SmartTag xmlattrname must be followed by xmlattrvalue".to_string(),
                    ));
                },
                _ => {
                    if !self.consume_destination_text_token_bounded(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        crate::smart_tag::MAX_SMART_TAG_ATTRIBUTE_NAME_BYTES,
                        "SmartTag attribute name",
                    )? {
                        return Err(RtfError::MalformedDocument(
                            "RTF SmartTag xmlattrname contains unsupported content".to_string(),
                        ));
                    }
                },
            }
        }
    }

    /// Collect direct `#PCDATA` after `\\xmlattrvalue` until the enclosing
    /// `xmlattr` group's closing brace.  The brace belongs to the outer
    /// group, so leave it for `parse_smart_tag_attribute_group` to consume.
    fn collect_smart_tag_direct_attribute_value(
        &mut self,
        context: &str,
        max_bytes: usize,
    ) -> RtfResult<String> {
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    return Ok(value.trim_end_matches(['\r', '\n']).to_string());
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token_bounded(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        max_bytes,
                        context,
                    )? {
                        return Err(RtfError::MalformedDocument(format!(
                            "RTF SmartTag {context} contains grouped, binary, or active data"
                        )));
                    }
                },
            }
        }
    }

    /// Parse a starred RTF 1.9.1 SmartTag closing destination.
    pub(super) fn parse_smart_tag_close_destination(&mut self) -> RtfResult<()> {
        require_body_markup(self, "SmartTag")?;
        self.pos += 2; // ignorable marker and xmlclose
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF SmartTag xmlclose contains text".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag xmlclose is malformed".to_string(),
            ));
        }
        self.pos += 1;
        let open = self.open_smart_tags.pop().ok_or_else(|| {
            RtfError::MalformedDocument("RTF SmartTag xmlclose has no open tag".to_string())
        })?;
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::SmartTagClose(open.order),
        ));
        self.smart_tag_spans.push(super::SmartTagSpan {
            tag: open,
            end: self.body_text_len,
        });
        Ok(())
    }

    /// Parse one starred main-body move-bookmark start or end destination.
    pub(super) fn parse_move_bookmark_destination(&mut self) -> RtfResult<()> {
        require_body_markup(self, "move-bookmark")?;
        let kind = match self.tokens.get(self.pos + 1) {
            Some(Token::Control(ControlWord::MoveFromStart | ControlWord::MoveFromEnd)) => {
                crate::MoveBookmarkKind::From
            },
            Some(Token::Control(ControlWord::MoveToStart | ControlWord::MoveToEnd)) => {
                crate::MoveBookmarkKind::To
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF move-bookmark destination".to_string(),
                ));
            },
        };
        let is_start = matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(
                ControlWord::MoveFromStart | ControlWord::MoveToStart
            ))
        );
        self.pos += 2;
        let value = self.collect_custom_xml_destination_text(
            "move-bookmark",
            crate::move_bookmark::MAX_MOVE_BOOKMARK_TAG_BYTES + 13,
        )?;
        let mut pieces = value.split_ascii_whitespace();
        let tag = pieces.next().unwrap_or_default().to_string();
        let payload = pieces.next();
        if pieces.next().is_some() || tag.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark destination has invalid fields".to_string(),
            ));
        }
        if tag.len() > crate::move_bookmark::MAX_MOVE_BOOKMARK_TAG_BYTES
            || !tag.bytes().all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark tag must be alphanumeric and at most 20 bytes".to_string(),
            ));
        }
        if !is_start && payload.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark end must contain only its tag".to_string(),
            ));
        }
        let key = (kind, tag.clone());
        if is_start {
            if self
                .open_move_bookmarks
                .get(&key)
                .is_some_and(|opens| !opens.is_empty())
                || self.completed_move_bookmarks.contains(&key)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF move-bookmark tag is duplicated within one location".to_string(),
                ));
            }
            let payload = payload.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF move-bookmark start lacks author/date payload".to_string(),
                )
            })?;
            let (author, date) = decode_move_hex(payload)?;
            let order = self.next_move_bookmark_order;
            if order >= crate::move_bookmark::MAX_MOVE_BOOKMARKS {
                return Err(RtfError::MalformedDocument(
                    "RTF move-bookmark count exceeds the safety limit".to_string(),
                ));
            }
            self.next_move_bookmark_order = self
                .next_move_bookmark_order
                .checked_add(1)
                .ok_or_else(|| {
                    RtfError::MalformedDocument("RTF move-bookmark order overflow".to_string())
                })?;
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::MoveBookmarkStart(order),
            ));
            self.open_move_bookmarks
                .entry(key)
                .or_default()
                .push(super::OpenMoveBookmark {
                    kind,
                    tag,
                    author,
                    date,
                    position: self.body_text_len,
                    order,
                });
        } else if let Some(open) = self.open_move_bookmarks.get_mut(&key).and_then(Vec::pop) {
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::MoveBookmarkEnd(open.order),
            ));
            self.move_bookmark_spans.push(super::MoveBookmarkSpan {
                bookmark: open,
                end: self.body_text_len,
            });
            self.completed_move_bookmarks.insert(key);
        } else {
            self.unmatched_move_bookmarks = true;
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
        self.consume_destination_text_token_bounded(
            value,
            unicode_skip,
            fallback_skip,
            usize::MAX,
            context,
        )
    }

    fn consume_destination_text_token_bounded(
        &mut self,
        value: &mut String,
        unicode_skip: &mut usize,
        fallback_skip: &mut usize,
        max_bytes: usize,
        context: &str,
    ) -> RtfResult<bool> {
        match self.tokens.get(self.pos).cloned() {
            Some(Token::Text(text)) => {
                let skipped = (*fallback_skip).min(text.chars().count());
                *fallback_skip -= skipped;
                let remainder = text
                    .char_indices()
                    .nth(skipped)
                    .and_then(|(index, _)| text.get(index..))
                    .unwrap_or_default();
                let decoded = self.decode_transport_text(remainder)?;
                Self::ensure_destination_text_capacity(
                    value.len(),
                    decoded.len(),
                    max_bytes,
                    context,
                )?;
                value.push_str(&decoded);
                self.pos += 1;
                Ok(true)
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
                let decoded = String::from_utf16(&utf16).map_err(|error| {
                    RtfError::InvalidUnicode(format!(
                        "invalid Unicode RTF custom XML {context}: {error}"
                    ))
                })?;
                Self::ensure_destination_text_capacity(
                    value.len(),
                    decoded.len(),
                    max_bytes,
                    context,
                )?;
                value.push_str(&decoded);
                *fallback_skip = unicode_skip.saturating_mul(utf16.len());
                Ok(true)
            },
            Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                *unicode_skip = count.max(0).cast_unsigned() as usize;
                self.pos += 1;
                Ok(true)
            },
            Some(Token::Control(control)) if control_symbol_text(&control).is_some() => {
                let text = control_symbol_text(&control).unwrap_or_default();
                Self::ensure_destination_text_capacity(
                    value.len(),
                    text.len(),
                    max_bytes,
                    context,
                )?;
                value.push_str(text);
                self.pos += 1;
                Ok(true)
            },
            _ => Ok(false),
        }
    }

    fn ensure_destination_text_capacity(
        current: usize,
        incoming: usize,
        max_bytes: usize,
        context: &str,
    ) -> RtfResult<()> {
        let observed = current.checked_add(incoming).ok_or_else(|| {
            RtfError::MalformedDocument(format!(
                "RTF {context} destination exceeds the safety limit"
            ))
        })?;
        if observed > max_bytes {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {context} destination exceeds the safety limit"
            )));
        }
        Ok(())
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
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(value.trim_end_matches(['\r', '\n']).to_string());
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token_bounded(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        max_bytes,
                        context,
                    )? {
                        return Err(RtfError::MalformedDocument(format!(
                            "RTF custom XML {context} destination contains grouped, binary, or active data"
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
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
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
                    let id = value.cast_unsigned();
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
        let tag_name = name.trim_end_matches(['\r', '\n']).to_string();
        if tag_name.is_empty() {
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
                .saturating_add(tag_name.len())
                .saturating_add(
                    attributes
                        .iter()
                        .map(|(attr_name, attr_value)| {
                            attr_name.len().saturating_add(attr_value.len())
                        })
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
                name: tag_name,
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
        if open.name != tag_name {
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

    /// Materialize complete SmartTag ranges after all main-body text has been
    /// decoded.  Unclosed markers are rejected; the source destinations are
    /// otherwise retained only as bounded inert metadata.
    pub(super) fn finalize_smart_tags(&mut self) -> RtfResult<()> {
        if let Some(open) = self.open_smart_tags.last() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF SmartTag '{}' is not closed",
                open.name
            )));
        }
        if self.smart_tag_spans.is_empty() {
            return Ok(());
        }
        if self.smart_tag_spans.len() > crate::smart_tag::MAX_SMART_TAGS {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag count exceeds the safety limit".to_string(),
            ));
        }

        self.smart_tag_spans
            .sort_unstable_by_key(|span| span.tag.order);
        let total_bytes = self
            .smart_tag_spans
            .iter()
            .try_fold(0usize, |total, span| {
                let content_len = span.end.checked_sub(span.tag.position).ok_or_else(|| {
                    RtfError::MalformedDocument("SmartTag does not align to body text".to_string())
                })?;
                let total = total.checked_add(span.tag.name.len()).ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF SmartTag aggregate text exceeds the safety limit".to_string(),
                    )
                })?;
                let total = span
                    .tag
                    .attributes
                    .iter()
                    .try_fold(total, |total, attribute| {
                        total
                            .checked_add(attribute.name.len())
                            .and_then(|total| total.checked_add(attribute.value.len()))
                            .ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF SmartTag aggregate text exceeds the safety limit"
                                        .to_string(),
                                )
                            })
                    })?;
                total.checked_add(content_len).ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF SmartTag aggregate text exceeds the safety limit".to_string(),
                    )
                })
            })?;
        if total_bytes > crate::smart_tag::MAX_SMART_TAG_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag aggregate text exceeds the safety limit".to_string(),
            ));
        }
        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        let mut order_to_index = vec![None; self.next_smart_tag_order];
        for span in self.smart_tag_spans.drain(..) {
            let content = body.get(span.tag.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument("SmartTag does not align to body text".to_string())
            })?;
            let index = self.smart_tags.len();
            let Some(mapped) = order_to_index.get_mut(span.tag.order) else {
                return Err(RtfError::MalformedDocument(
                    "SmartTag source order exceeds the parser bound".to_string(),
                ));
            };
            *mapped = Some(index);
            self.smart_tags.push(crate::SmartTag::new(
                Cow::Owned(span.tag.name),
                span.tag.namespace,
                span.tag.attributes,
                span.tag.position,
                Cow::Owned(content.to_string()),
            )?);
        }

        // Normally open order and materialized index are identical.  Keep the
        // explicit mapping so event references remain correct if parsing ever
        // filters a valid range before finalization.
        for event in &mut self.body_story_events {
            if let ParsedBodyStoryEvent::Resolved(event) = event {
                match event {
                    crate::BodyStoryEvent::SmartTagOpen(order)
                    | crate::BodyStoryEvent::SmartTagClose(order) => {
                        if let Some(Some(index)) = order_to_index.get(*order) {
                            *order = *index;
                        }
                    },
                    _ => {},
                }
            }
        }
        Ok(())
    }

    /// Materialize complete tracked-move start/end ranges.  An unmatched
    /// start or end is deliberately ignored for the semantic model, matching
    /// the RTF 1.9.1 compatibility rule.  Equal tags across Move From and
    /// Move To ranges remain a shared source identifier; this bounded model
    /// does not construct a cross-location pair or apply its fallback. Changed
    /// publication is refused for a source containing an unmatched marker
    /// because that marker is not represented by an opaque node and cannot be
    /// emitted safely; immutable no-op writes retain the original source
    /// bytes.
    pub(super) fn finalize_move_bookmarks(&mut self) -> RtfResult<()> {
        self.unmatched_move_bookmarks |= self
            .open_move_bookmarks
            .values()
            .any(|spans| !spans.is_empty());
        if self.move_bookmark_spans.len() > crate::move_bookmark::MAX_MOVE_BOOKMARKS {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark count exceeds the safety limit".to_string(),
            ));
        }
        self.move_bookmark_spans
            .sort_unstable_by_key(|span| span.bookmark.order);
        let total_bytes = self
            .move_bookmark_spans
            .iter()
            .try_fold(0usize, |total, span| {
                let content_len =
                    span.end
                        .checked_sub(span.bookmark.position)
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "move-bookmark does not align to body text".to_string(),
                            )
                        })?;
                total
                    .checked_add(span.bookmark.tag.len())
                    .and_then(|total| total.checked_add(content_len))
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF move-bookmark aggregate text exceeds the safety limit".to_string(),
                        )
                    })
            })?;
        if total_bytes > crate::move_bookmark::MAX_MOVE_BOOKMARK_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark aggregate text exceeds the safety limit".to_string(),
            ));
        }
        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        let mut order_to_index = vec![None; self.next_move_bookmark_order];
        for span in self.move_bookmark_spans.drain(..) {
            let content = body.get(span.bookmark.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument("move-bookmark does not align to body text".to_string())
            })?;
            let index = self.move_bookmarks.len();
            let Some(mapped) = order_to_index.get_mut(span.bookmark.order) else {
                return Err(RtfError::MalformedDocument(
                    "move-bookmark source order exceeds the parser bound".to_string(),
                ));
            };
            *mapped = Some(index);
            self.move_bookmarks.push(crate::MoveBookmark::new(
                span.bookmark.kind,
                Cow::Owned(span.bookmark.tag),
                span.bookmark.author,
                span.bookmark.date,
                span.bookmark.position,
                Cow::Owned(content.to_string()),
            )?);
        }

        // Discard unmatched move starts and translate the temporary source
        // orders used while parsing into stable public vector indices.
        self.body_story_events.retain_mut(|event| {
            let Some(ParsedBodyStoryEvent::Resolved(event)) = Some(event) else {
                return true;
            };
            match event {
                crate::BodyStoryEvent::MoveBookmarkStart(order)
                | crate::BodyStoryEvent::MoveBookmarkEnd(order) => {
                    if let Some(Some(index)) = order_to_index.get(*order) {
                        *order = *index;
                        true
                    } else {
                        false
                    }
                },
                _ => true,
            }
        });
        Ok(())
    }
}
