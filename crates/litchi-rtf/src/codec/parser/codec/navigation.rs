use super::{
    ControlWord, Cow, Destination, IndexEntry, IndexPageReference, MAX_NAVIGATION_ENTRIES,
    MAX_NAVIGATION_ENTRY_DEPTH, MAX_NAVIGATION_ENTRY_TEXT_BYTES,
    MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES, NavigationEntry, ParsedBodyStoryEvent, Parser, RtfError,
    RtfResult, SmallVec, TableOfContentsEntry, Token, control_symbol_text,
};

impl<'a> Parser<'a> {
    pub(super) fn parse_navigation_entry_destination(&mut self) -> RtfResult<()> {
        if self.navigation_entries.len() >= MAX_NAVIGATION_ENTRIES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry count limit exceeded".to_string(),
            ));
        }
        let state = self.prepare_revision_event()?;
        let entry = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::IndexEntry)) => self.parse_index_entry()?,
            Some(Token::Control(ControlWord::TableOfContentsEntry)) => {
                self.parse_table_of_contents_entry(false)?
            },
            Some(Token::Control(ControlWord::TableOfContentsEntryNoPage)) => {
                self.parse_table_of_contents_entry(true)?
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF navigation-entry destination".to_string(),
                ));
            },
        };
        entry.validate()?;
        let added = entry.text_bytes().ok_or_else(|| {
            RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
        })?;
        let position = entry.position();
        self.navigation_entry_text_bytes = self
            .navigation_entry_text_bytes
            .checked_add(added)
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
            })?;
        if self.navigation_entry_text_bytes > MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry aggregate text limit exceeded".to_string(),
            ));
        }
        let index = self.navigation_entries.len();
        self.navigation_entries.push(entry);
        if state.in_table || state.table_nesting_level >= 2 {
            self.push_cell_story_event(
                state.table_nesting_level,
                crate::CellStoryEvent::NavigationEntry(crate::CellStoryReference {
                    index,
                    position,
                }),
            )?;
        } else {
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::NavigationEntry(index),
            ));
        }
        Ok(())
    }

    pub(super) fn parse_generated_list_marker(
        &mut self,
        kind: crate::GeneratedListMarkerKind,
    ) -> RtfResult<()> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF generated list marker may occur only in the visible document body".to_string(),
            ));
        }
        if self.generated_list_markers.len()
            >= crate::generated_list_marker::MAX_GENERATED_LIST_MARKERS
        {
            return Err(RtfError::MalformedDocument(
                "RTF generated list-marker count exceeds the safety limit".to_string(),
            ));
        }

        self.pos += 1;
        let mut depth = 0usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut text = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    let marker = crate::GeneratedListMarker {
                        kind,
                        text: Cow::Borrowed(self.arena.alloc_str(&text)),
                        position: self.body_text_len,
                    };
                    marker.validate()?;
                    if self.generated_list_markers.last().is_some_and(|previous| {
                        previous.position == marker.position && previous.kind == marker.kind
                    }) {
                        return Err(RtfError::MalformedDocument(
                            "RTF contains duplicate generated list markers at one body position"
                                .to_string(),
                        ));
                    }
                    self.generated_list_marker_text_bytes = self
                        .generated_list_marker_text_bytes
                        .checked_add(marker.text.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF generated list-marker text size overflow".to_string(),
                            )
                        })?;
                    if self.generated_list_marker_text_bytes
                        > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF generated list-marker text exceeds the aggregate safety limit"
                                .to_string(),
                        ));
                    }
                    let index = self.generated_list_markers.len();
                    self.generated_list_markers.push(marker);
                    self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                        crate::BodyStoryEvent::GeneratedListMarker(index),
                    ));
                    return Ok(());
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::Shape(_)
                                | ControlWord::ShapeGroup(_)
                                | ControlWord::FormField
                                | ControlWord::DataField
                        ))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF generated list marker contains an active nested destination"
                                .to_string(),
                        ));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF generated list-marker nesting depth overflow".to_string(),
                        )
                    })?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    text.push_str(&self.parse_style_unicode(*code, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::Control(
                    ControlWord::Field
                    | ControlWord::Object
                    | ControlWord::Picture
                    | ControlWord::Shape(_)
                    | ControlWord::ShapeGroup(_)
                    | ControlWord::FormField
                    | ControlWord::DataField
                    | ControlWord::Par
                    | ControlWord::Line,
                )) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generated list marker contains active or structural content"
                            .to_string(),
                    ));
                },
                Some(Token::Control(_)) => self.pos += 1,
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generated list marker cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF generated list marker exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_index_entry(&mut self) -> RtfResult<NavigationEntry<'a>> {
        let position = self.current_story_position()?;
        self.pos += 1; // \xe
        let mut text = String::new();
        let mut index_id = None;
        let mut bold_page_number = false;
        let mut italic_page_number = false;
        let mut page_reference = IndexPageReference::CurrentPage;
        let mut yomi = None;
        let mut saw_yomi = false;
        let mut saw_text = false;
        let mut saw_reference = false;

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                        (Some(Token::Control(ControlWord::IndexReplacementText)), _) => {
                            if !saw_text || saw_reference {
                                return Err(RtfError::MalformedDocument(
                                "RTF index entry has a misplaced or duplicate txe/rxe destination".to_string(),
                            ));
                            }
                            page_reference = IndexPageReference::ReplacementText(Cow::Owned(
                                self.parse_navigation_subdestination(false)?,
                            ));
                            saw_reference = true;
                        },
                        (Some(Token::Control(ControlWord::IndexBookmarkRange)), _) => {
                            if !saw_text || saw_reference {
                                return Err(RtfError::MalformedDocument(
                                "RTF index entry has a misplaced or duplicate txe/rxe destination".to_string(),
                            ));
                            }
                            page_reference = IndexPageReference::BookmarkRange(Cow::Owned(
                                self.parse_navigation_subdestination(false)?,
                            ));
                            saw_reference = true;
                        },
                        (
                            Some(Token::Control(ControlWord::IgnorableDestination)),
                            Some(Token::Control(ControlWord::IndexPronunciation)),
                        ) => {
                            if !saw_yomi || yomi.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF pxe pronunciation requires one preceding yxe".to_string(),
                                ));
                            }
                            yomi = Some(Cow::Owned(self.parse_navigation_subdestination(true)?));
                        },
                        (Some(Token::Control(ControlWord::IndexYomi)), _) => {
                            if !saw_text || saw_yomi || yomi.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF index entry has a misplaced or duplicate yxe group"
                                        .to_string(),
                                ));
                            }
                            yomi = Some(Cow::Owned(self.parse_index_yomi_group()?));
                            saw_yomi = true;
                        },
                        _ => {
                            if saw_reference || yomi.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF index entry text must precede its subdestinations"
                                        .to_string(),
                                ));
                            }
                            self.parse_navigation_text_group(&mut text, true, 1)?;
                            saw_text = !text.is_empty();
                        },
                    }
                },
                Some(Token::Control(ControlWord::IndexIdentifier(value))) => {
                    if saw_text || index_id.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF xef must occur once before index text".to_string(),
                        ));
                    }
                    let identifier = value.ok_or_else(|| {
                        RtfError::MalformedDocument("RTF xef requires a parameter".to_string())
                    })?;
                    index_id = Some(u8::try_from(identifier).map_err(|_err| {
                        RtfError::MalformedDocument("RTF xef parameter is out of range".to_string())
                    })?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexBold(value))) => {
                    if saw_text || bold_page_number {
                        return Err(RtfError::MalformedDocument(
                            "RTF bxe must occur once before index text".to_string(),
                        ));
                    }
                    bold_page_number = *value;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexItalic(value))) => {
                    if saw_text || italic_page_number {
                        return Err(RtfError::MalformedDocument(
                            "RTF ixe must occur once before index text".to_string(),
                        ));
                    }
                    italic_page_number = *value;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexYomi)) => {
                    if !saw_text || saw_yomi {
                        return Err(RtfError::MalformedDocument(
                            "RTF yxe must occur once after index text".to_string(),
                        ));
                    }
                    saw_yomi = true;
                    self.pos += 1;
                },
                Some(_) => {
                    if saw_reference || yomi.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF index text cannot follow a subdestination".to_string(),
                        ));
                    }
                    self.parse_navigation_text_token(&mut text, true, 1)?;
                    saw_text = !text.is_empty();
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        if saw_yomi != yomi.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF yxe and pxe pronunciation controls must occur together".to_string(),
            ));
        }
        let mut entry = IndexEntry::new(position, Cow::Owned(text))?;
        entry.index_id = index_id;
        entry.bold_page_number = bold_page_number;
        entry.italic_page_number = italic_page_number;
        entry.page_reference = page_reference;
        entry.yomi = yomi;
        entry.validate()?;
        Ok(NavigationEntry::Index(entry))
    }

    pub(super) fn parse_index_yomi_group(&mut self) -> RtfResult<String> {
        self.pos += 2; // group open and \yxe
        let state = self.current_state()?.clone();
        self.states.push(state);
        if !matches!(
            (self.tokens.get(self.pos), self.tokens.get(self.pos + 1)),
            (
                Some(Token::OpenBrace),
                Some(Token::Control(ControlWord::IgnorableDestination))
            )
        ) || !matches!(
            self.tokens.get(self.pos + 2),
            Some(Token::Control(ControlWord::IndexPronunciation))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF yxe group must contain one immediate starred pxe destination".to_string(),
            ));
        }
        let value = self.parse_navigation_subdestination(true)?;
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF yxe group must contain only its pxe destination".to_string(),
            ));
        }
        self.pos += 1;
        self.states.pop();
        Ok(value)
    }

    pub(super) fn parse_table_of_contents_entry(
        &mut self,
        suppress_page_number: bool,
    ) -> RtfResult<NavigationEntry<'a>> {
        let position = self.current_story_position()?;
        self.pos += 1; // \tc or \tcn
        let mut text = String::new();
        let mut table_id = b'C';
        let mut level = 1u8;
        let mut saw_table = false;
        let mut saw_level = false;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Control(ControlWord::TableOfContentsTable(value))) => {
                    if !text.is_empty() || saw_table {
                        return Err(RtfError::MalformedDocument(
                            "RTF tcf must occur once before TOC-entry text".to_string(),
                        ));
                    }
                    table_id = u8::try_from(*value).map_err(|_err| {
                        RtfError::MalformedDocument("RTF tcf parameter is out of range".to_string())
                    })?;
                    saw_table = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::TableOfContentsLevel(value))) => {
                    if !text.is_empty() || saw_level {
                        return Err(RtfError::MalformedDocument(
                            "RTF tcl must occur once before TOC-entry text".to_string(),
                        ));
                    }
                    level = u8::try_from(*value).map_err(|_err| {
                        RtfError::MalformedDocument("RTF tcl parameter is out of range".to_string())
                    })?;
                    saw_level = true;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => self.parse_navigation_text_group(&mut text, true, 1)?,
                Some(_) => self.parse_navigation_text_token(&mut text, true, 1)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        let mut entry = TableOfContentsEntry::new(position, Cow::Owned(text))?;
        entry.table_id = table_id;
        entry.level = level;
        entry.suppress_page_number = suppress_page_number;
        entry.validate()?;
        Ok(NavigationEntry::TableOfContents(entry))
    }

    pub(super) fn parse_navigation_subdestination(&mut self, starred: bool) -> RtfResult<String> {
        self.pos += if starred { 3 } else { 2 }; // group, optional star, destination
        let state = self.current_state()?.clone();
        self.states.push(state);
        let mut value = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    break;
                },
                Some(Token::OpenBrace) => {
                    self.parse_navigation_text_group(&mut value, false, 1)?;
                },
                Some(_) => self.parse_navigation_text_token(&mut value, false, 1)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        if value.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF navigation subdestination text cannot be empty".to_string(),
            ));
        }
        Ok(value)
    }

    pub(super) fn parse_navigation_text_group(
        &mut self,
        output: &mut String,
        visible: bool,
        depth: usize,
    ) -> RtfResult<()> {
        if depth > MAX_NAVIGATION_ENTRY_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry nesting limit exceeded".to_string(),
            ));
        }
        self.pos += 1; // group open
        let state = self.current_state()?.clone();
        self.states.push(state);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    return Ok(());
                },
                Some(Token::OpenBrace) => {
                    self.parse_navigation_text_group(output, visible, depth + 1)?;
                },
                Some(_) => self.parse_navigation_text_token(output, visible, depth)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_navigation_text_token(
        &mut self,
        output: &mut String,
        visible: bool,
        _depth: usize,
    ) -> RtfResult<()> {
        let decoded = match self.tokens.get(self.pos) {
            Some(Token::Text(text)) => {
                let decoded = self.decode_transport_text(text)?;
                self.pos += 1;
                Some(decoded)
            },
            Some(Token::Control(ControlWord::Unicode(code))) => {
                Some(self.parse_navigation_unicode_sequence(*code)?)
            },
            Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                self.pos += 1;
                Some("\n".to_string())
            },
            Some(Token::Control(ControlWord::Tab)) => {
                self.pos += 1;
                Some("\t".to_string())
            },
            Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                let decoded = control_symbol_text(control).unwrap_or_default().to_string();
                self.pos += 1;
                Some(decoded)
            },
            Some(Token::Control(control)) => {
                if Self::forbidden_navigation_control(control) {
                    return Err(RtfError::MalformedDocument(
                        "RTF navigation entries cannot contain active or nested destinations"
                            .to_string(),
                    ));
                }
                let control_word = *control;
                self.pos += 1;
                self.apply_control_word(&control_word)?;
                None
            },
            Some(Token::Binary(_)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF navigation entries cannot contain binary data".to_string(),
                ));
            },
            Some(Token::OpenBrace | Token::CloseBrace) => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF navigation-entry group structure".to_string(),
                ));
            },
            None => return Err(RtfError::UnexpectedEof),
        };
        if let Some(decoded_text) = decoded {
            let new_len = output
                .len()
                .checked_add(decoded_text.len())
                .ok_or_else(|| {
                    RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
                })?;
            if new_len > MAX_NAVIGATION_ENTRY_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF navigation-entry text limit exceeded".to_string(),
                ));
            }
            output.push_str(&decoded_text);
            if visible && !self.current_state()?.formatting.hidden {
                self.append_semantic_text(&decoded_text)?;
            }
        }
        Ok(())
    }

    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
    )]
    pub(super) fn parse_navigation_unicode_sequence(
        &mut self,
        first_code: i32,
    ) -> RtfResult<String> {
        let skip_count = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        let mut code = first_code;
        let mut remainder = String::new();
        loop {
            utf16.push(code as u16);
            self.pos += 1;
            let mut fallback_skip = skip_count;
            while fallback_skip > 0 {
                match self.tokens.get(self.pos) {
                    Some(Token::Text(text)) => {
                        let count = text.chars().count();
                        if count <= fallback_skip {
                            fallback_skip -= count;
                        } else {
                            remainder.extend(text.chars().skip(fallback_skip));
                            fallback_skip = 0;
                        }
                        self.pos += 1;
                    },
                    Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                        fallback_skip -= 1;
                        self.pos += 1;
                    },
                    _ => break,
                }
            }
            if !remainder.is_empty() {
                break;
            }
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::Unicode(next))) => code = *next,
                _ => break,
            }
        }
        let mut decoded = String::from_utf16(&utf16).map_err(|error| {
            RtfError::InvalidUnicode(format!("invalid navigation-entry Unicode: {error}"))
        })?;
        decoded.push_str(&self.decode_transport_text(&remainder)?);
        Ok(decoded)
    }

    pub(super) fn forbidden_navigation_control(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::IgnorableDestination
                | ControlWord::Field
                | ControlWord::FieldInstruction
                | ControlWord::FieldResult
                | ControlWord::Object
                | ControlWord::Result
                | ControlWord::Picture
                | ControlWord::Shape(_)
                | ControlWord::ShapeGroup(_)
                | ControlWord::DocumentVariable
                | ControlWord::UserProperties
                | ControlWord::Annotation
                | ControlWord::Footnote
                | ControlWord::Endnote
                | ControlWord::Header
                | ControlWord::HeaderFirst
                | ControlWord::HeaderLeft
                | ControlWord::HeaderRight
                | ControlWord::Footer
                | ControlWord::FooterFirst
                | ControlWord::FooterLeft
                | ControlWord::FooterRight
                | ControlWord::FontTable
                | ControlWord::ColorTable
                | ControlWord::StyleSheet
                | ControlWord::ListTable
                | ControlWord::ListOverrideTable
                | ControlWord::RevisionTable
                | ControlWord::IndexEntry
                | ControlWord::IndexIdentifier(_)
                | ControlWord::IndexBold(_)
                | ControlWord::IndexItalic(_)
                | ControlWord::IndexReplacementText
                | ControlWord::IndexBookmarkRange
                | ControlWord::IndexYomi
                | ControlWord::IndexPronunciation
                | ControlWord::TableOfContentsEntry
                | ControlWord::TableOfContentsEntryNoPage
                | ControlWord::TableOfContentsTable(_)
                | ControlWord::TableOfContentsLevel(_)
        )
    }
}
