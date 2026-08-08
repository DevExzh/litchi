use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_file_table(&mut self) -> RtfResult<crate::FileTable<'a>> {
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF filetbl must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::FileTable))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF filetbl destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut table = crate::FileTable::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    table.validate()?;
                    return Ok(table);
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::FileEntry))
                    ) =>
                {
                    let entry = self.parse_file_table_entry()?;
                    table.add(entry)?;
                    continue;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF filetbl cannot contain fields, objects, or unknown destinations"
                            .to_string(),
                    ));
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF filetbl".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_file_table_entry(&mut self) -> RtfResult<crate::FileTableEntry<'a>> {
        self.pos += 2; // opening brace and file
        let mut id = None;
        let mut relative = None;
        let mut operating_system = None;
        let mut valid_on = crate::FileSystemValidity::default();
        let mut location = crate::FileLocation::Local;
        let mut name = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let trimmed = name.trim_end_matches(['\r', '\n', ' ']);
                    let name = trimmed
                        .strip_suffix(';')
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF file-table name lacks its semicolon terminator".to_string(),
                            )
                        })?
                        .trim();
                    let mut entry = crate::FileTableEntry::new(
                        id.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF file entry lacks fid".to_string())
                        })?,
                        Cow::Owned(name.to_string()),
                    );
                    entry.relative_path_level = relative;
                    entry.operating_system = operating_system;
                    entry.valid_on = valid_on;
                    entry.location = location;
                    entry.validate()?;
                    return Ok(entry);
                },
                Some(Token::OpenBrace) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF file entry cannot contain fields, objects, nested destinations, or binary data".to_string(),
                    ));
                },
                Some(Token::Text(text)) => name.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0)
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default())
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::FileId(value) => {
                        if !seen.insert("fid") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fid".to_string(),
                            ));
                        }
                        id = Some(u32::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF fid".to_string())
                        })?);
                    },
                    ControlWord::FileRelative(value) => {
                        if !seen.insert("frelative") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF frelative".to_string(),
                            ));
                        }
                        relative = Some(u8::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF frelative".to_string())
                        })?);
                    },
                    ControlWord::FileOperatingSystem(value) => {
                        if !seen.insert("fosnum") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fosnum".to_string(),
                            ));
                        }
                        operating_system = Some(u8::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF fosnum".to_string())
                        })?);
                    },
                    ControlWord::FileValidMac => {
                        if !seen.insert("fvalidmac") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fvalidmac".to_string(),
                            ));
                        }
                        valid_on.mac = true;
                    },
                    ControlWord::FileValidDos => {
                        if !seen.insert("fvaliddos") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fvaliddos".to_string(),
                            ));
                        }
                        valid_on.dos = true;
                    },
                    ControlWord::FileValidNtfs => {
                        if !seen.insert("fvalidntfs") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fvalidntfs".to_string(),
                            ));
                        }
                        valid_on.ntfs = true;
                    },
                    ControlWord::FileValidHpfs => {
                        if !seen.insert("fvalidhpfs") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fvalidhpfs".to_string(),
                            ));
                        }
                        valid_on.hpfs = true;
                    },
                    ControlWord::FileNetwork => {
                        if !seen.insert("location") {
                            return Err(RtfError::MalformedDocument(
                                "conflicting RTF file locations".to_string(),
                            ));
                        }
                        location = crate::FileLocation::Network;
                    },
                    ControlWord::FileNonFileSystem => {
                        if !seen.insert("location") {
                            return Err(RtfError::MalformedDocument(
                                "conflicting RTF file locations".to_string(),
                            ));
                        }
                        location = crate::FileLocation::NonFileSystem;
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "unsupported control in RTF file entry".to_string(),
                        ));
                    },
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if name.len() > crate::file_table::MAX_FILE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF file-table name exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse font table.
    pub(super) fn parse_font_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \fonttbl

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    self.font_table.borrow().validate()?;
                    return Ok(());
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::Unknown(_, _))
                        ])
                    ) =>
                {
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                },
                Token::OpenBrace => {
                    self.parse_font_entry()?;
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control_in(token, crate::opaque::Context::Metadata)?;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse a single font table entry.
    pub(super) fn parse_font_entry(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip {

        let mut font_num = None;
        let mut font_family = FontFamily::Nil;
        let mut charset = None;
        let mut pitch = crate::FontPitch::Default;
        let mut code_page = None;
        let mut theme = None;
        let mut bidi = false;
        let mut alternate_name = None;
        let mut non_tagged_name = None;
        let mut panose = None;
        let mut embedded = None;
        let mut name = DeferredText::default();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut seen = std::collections::HashSet::new();

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::FontAlternateName)
                        ])
                    ) =>
                {
                    if !seen.insert("falt") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF falt destination".to_string(),
                        ));
                    }
                    alternate_name =
                        Some(self.parse_font_name_destination(ControlWord::FontAlternateName)?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::FontNonTaggedName)
                        ])
                    ) =>
                {
                    if !seen.insert("fname") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF fname destination".to_string(),
                        ));
                    }
                    non_tagged_name =
                        Some(self.parse_font_name_destination(ControlWord::FontNonTaggedName)?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::FontPanose)
                        ])
                    ) =>
                {
                    if !seen.insert("panose") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF panose destination".to_string(),
                        ));
                    }
                    panose = Some(self.parse_font_panose_destination()?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::FontEmbedded)
                        ])
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::FontEmbedded))
                    ) =>
                {
                    if !seen.insert("fontemb") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF fontemb destination".to_string(),
                        ));
                    }
                    embedded = Some(self.parse_font_embedded_destination()?);
                },
                Token::OpenBrace => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::Field | ControlWord::Object))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF font entry cannot contain fields or objects".to_string(),
                        ));
                    }
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                },
                Token::Control(ControlWord::FontNumber(n)) => {
                    if !seen.insert("font-number") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font ID".to_string(),
                        ));
                    }
                    font_num = Some(FontRef::try_from(*n).map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF font ID".to_string())
                    })?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontFamily(family)) => {
                    if !seen.insert("family") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font family".to_string(),
                        ));
                    }
                    font_family = match *family {
                        "roman" => FontFamily::Roman,
                        "swiss" => FontFamily::Swiss,
                        "modern" => FontFamily::Modern,
                        "script" => FontFamily::Script,
                        "decor" => FontFamily::Decor,
                        "tech" => FontFamily::Tech,
                        _ => FontFamily::Nil,
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontCharset(cs)) => {
                    if !seen.insert("charset") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font charset".to_string(),
                        ));
                    }
                    let charset_id = u8::try_from(*cs).map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF font charset".to_string())
                    })?;
                    charset = Some(FontCharset::new(charset_id).ok_or_else(|| {
                        RtfError::MalformedDocument(format!(
                            "unsupported RTF font charset {charset_id}"
                        ))
                    })?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontPitch(value)) => {
                    if !seen.insert("pitch") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font pitch".to_string(),
                        ));
                    }
                    pitch = match *value {
                        0 => crate::FontPitch::Default,
                        1 => crate::FontPitch::Fixed,
                        2 => crate::FontPitch::Variable,
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "invalid RTF font pitch".to_string(),
                            ));
                        },
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontCodePage(value)) => {
                    if !seen.insert("code-page") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font code page".to_string(),
                        ));
                    }
                    let page = u32::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF font code page".to_string())
                    })?;
                    code_page = Some(FontPage::new(page).ok_or_else(|| {
                        RtfError::MalformedDocument(format!(
                            "unsupported RTF font code page {page}"
                        ))
                    })?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontTheme(word)) => {
                    if !seen.insert("theme") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF font theme selector".to_string(),
                        ));
                    }
                    theme = super::super::super::types::FontTheme::from_control_word(word);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontBidi(parameter)) => {
                    if !seen.insert("bidi") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF fbidi font property".to_string(),
                        ));
                    }
                    if parameter.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF fbidi font property must not have a numeric parameter".to_string(),
                        ));
                    }
                    bidi = true;
                    self.pos += 1;
                },
                Token::Text(text) => {
                    name.push_transport(text)?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(first)) => {
                    self.append_deferred_unicode(&mut name, *first, unicode_skip)?;
                },
                Token::Control(ControlWord::UnicodeSkip(value)) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    name.push_unicode(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control_in(token, crate::opaque::Context::Metadata)?;
                },
                _ => {
                    self.pos += 1;
                },
            }
            if name.source_len() > 4_096 {
                return Err(RtfError::MalformedDocument(
                    "RTF font name exceeds the safety limit".to_string(),
                ));
            }
        }

        let font_num = font_num
            .ok_or_else(|| RtfError::MalformedDocument("RTF font entry lacks an ID".to_string()))?;
        let header_encoding = self.current_state()?.encoding;
        let encoding = match (code_page, charset) {
            (Some(page), _) => RtfEncoding::from_font_page(page),
            (None, Some(FontCharset::Default) | None) => header_encoding,
            (None, Some(charset)) => match charset.page() {
                Some(page) => RtfEncoding::from_font_page(page),
                None => {
                    let has_non_ascii = name.has_non_ascii_transport()
                        || alternate_name
                            .as_ref()
                            .is_some_and(DeferredText::has_non_ascii_transport)
                        || non_tagged_name
                            .as_ref()
                            .is_some_and(DeferredText::has_non_ascii_transport);
                    if has_non_ascii {
                        return Err(RtfError::MalformedDocument(format!(
                            "unsupported RTF font charset {} for non-ASCII font metadata",
                            charset.id()
                        )));
                    }
                    // ASCII transport and explicit Unicode escapes are invariant
                    // across the unavailable legacy charset and the header page.
                    header_encoding
                },
            },
        };
        let decode_name = |value: DeferredText, context: &str| -> RtfResult<String> {
            let value = value.decode(encoding, context)?;
            Ok(value
                .trim()
                .strip_suffix(';')
                .unwrap_or(value.trim())
                .trim()
                .to_string())
        };
        let name = decode_name(name, "font name")?;
        let mut font = Font::new(Cow::Owned(name), font_family);
        font.charset = charset;
        font.alternate_name = alternate_name
            .map(|value| decode_name(value, "alternate font name"))
            .transpose()?
            .map(Cow::Owned);
        font.non_tagged_name = non_tagged_name
            .map(|value| decode_name(value, "non-tagged font name"))
            .transpose()?
            .map(Cow::Owned);
        font.panose = panose;
        font.pitch = pitch;
        font.code_page = code_page;
        font.embedded = embedded;
        font.theme = theme;
        font.bidi = bidi;
        font.validate()?;
        if let Some(existing) = self.font_table.borrow().get(font_num) {
            if existing == &font {
                return Ok(());
            }
            return Err(RtfError::MalformedDocument(
                "conflicting duplicate RTF font ID".to_string(),
            ));
        }
        if self
            .font_table
            .borrow_mut()
            .insert(font_num, font)?
            .is_some()
        {
            return Err(RtfError::MalformedDocument(
                "conflicting duplicate RTF font ID".to_string(),
            ));
        }

        Ok(())
    }

    pub(super) fn parse_font_name_destination(
        &mut self,
        expected: ControlWord<'a>,
    ) -> RtfResult<DeferredText> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || !matches!(
                self.tokens.get(self.pos + 1),
                Some(Token::Control(ControlWord::IgnorableDestination))
            )
            || self.tokens.get(self.pos + 2) != Some(&Token::Control(expected))
        {
            return Err(RtfError::MalformedDocument(
                "invalid RTF font-name destination".to_string(),
            ));
        }
        self.pos += 3;
        let mut value = DeferredText::default();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if value.source_len() > 4_096 {
                        return Err(RtfError::MalformedDocument(
                            "oversized RTF alternate font name".to_string(),
                        ));
                    }
                    return Ok(value);
                },
                Some(Token::Text(text)) => value.push_transport(text)?,
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    self.append_deferred_unicode(&mut value, *first, unicode_skip)?;
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0)
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_unicode(control_symbol_text(control).unwrap_or_default())
                },
                Some(Token::OpenBrace) | Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF font-name destination contains non-text content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if value.source_len() > 4_096 {
                return Err(RtfError::MalformedDocument(
                    "RTF alternate font name exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_font_panose_destination(&mut self) -> RtfResult<[u8; 10]> {
        self.pos += 3; // opening brace, ignorable marker, panose
        let mut digits = String::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let compact: String = digits.chars().filter(|ch| !ch.is_whitespace()).collect();
                    if compact.len() != 20 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit())
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF panose must contain exactly ten hexadecimal bytes".to_string(),
                        ));
                    }
                    let mut panose = [0u8; 10];
                    for (byte, pair) in panose.iter_mut().zip(compact.as_bytes().chunks_exact(2)) {
                        let &[high, low] = pair else {
                            return Err(RtfError::MalformedDocument(
                                "invalid RTF panose payload".to_string(),
                            ));
                        };
                        let high = Self::hex_nibble(high).ok_or_else(|| {
                            RtfError::MalformedDocument("invalid RTF panose payload".to_string())
                        })?;
                        let low = Self::hex_nibble(low).ok_or_else(|| {
                            RtfError::MalformedDocument("invalid RTF panose payload".to_string())
                        })?;
                        *byte = (high << 4) | low;
                    }
                    return Ok(panose);
                },
                Some(Token::Text(text)) => digits.push_str(text),
                Some(Token::OpenBrace) | Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF panose contains non-hexadecimal content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if digits.len() > 64 {
                return Err(RtfError::MalformedDocument(
                    "RTF panose payload exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse the inert `fontemb` destination of a font-table entry.
    pub(super) fn parse_font_embedded_destination(
        &mut self,
    ) -> RtfResult<crate::EmbeddedFont<'static>> {
        self.pos += 1; // opening brace
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if self.tokens.get(self.pos) != Some(&Token::Control(ControlWord::FontEmbedded)) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF fontemb destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut embedded = crate::EmbeddedFont::default();
        let mut format_seen = false;
        let mut file_seen = false;
        let mut data = Vec::new();
        let mut high_nibble: Option<u8> = None;
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF fontemb contains an odd number of hexadecimal digits".to_string(),
                        ));
                    }
                    if !data.is_empty() {
                        embedded.data = Some(data);
                    }
                    embedded.validate()?;
                    return Ok(embedded);
                },
                Token::Control(ControlWord::FontEmbeddedType(kind)) => {
                    if format_seen {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF embedded font format".to_string(),
                        ));
                    }
                    format_seen = true;
                    embedded.format = match *kind {
                        "truetype" => crate::EmbeddedFontFormat::TrueType,
                        _ => crate::EmbeddedFontFormat::Nil,
                    };
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    let is_font_file = matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::FontFile)
                        ])
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::FontFile))
                    );
                    if is_font_file {
                        if file_seen {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF fontfile destination".to_string(),
                            ));
                        }
                        file_seen = true;
                        let (file_name, file_code_page) = self.parse_font_file_destination()?;
                        embedded.file_name = Some(Cow::Owned(file_name));
                        embedded.file_code_page = file_code_page;
                    } else {
                        self.pos += 1;
                        self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                    }
                },
                Token::Text(text) => {
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF fontemb contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            data.push((high << 4) | nibble);
                            if data.len() > crate::EmbeddedFont::MAX_DATA_BYTES {
                                return Err(RtfError::MalformedDocument(
                                    "RTF embedded font data exceeds the safety limit".to_string(),
                                ));
                            }
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Token::Binary(bytes) => {
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF fontemb binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    data.extend_from_slice(bytes);
                    if data.len() > crate::EmbeddedFont::MAX_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF embedded font data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse the nested `fontfile` destination of a `fontemb` group.
    pub(super) fn parse_font_file_destination(&mut self) -> RtfResult<(String, Option<FontPage>)> {
        self.pos += 1; // opening brace
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if self.tokens.get(self.pos) != Some(&Token::Control(ControlWord::FontFile)) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF fontfile destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut name = DeferredText::default();
        let mut code_page = None;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let encoding = code_page
                        .map(RtfEncoding::from_font_page)
                        .unwrap_or(self.current_state()?.encoding);
                    let name = name
                        .decode(encoding, "embedded font file name")?
                        .trim()
                        .to_string();
                    if name.is_empty() || name.len() > crate::EmbeddedFont::MAX_FILE_NAME_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "invalid or oversized RTF embedded font file name".to_string(),
                        ));
                    }
                    return Ok((name, code_page));
                },
                Some(Token::Control(ControlWord::FontCodePage(value))) => {
                    if code_page.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF fontfile code page".to_string(),
                        ));
                    }
                    let page = u32::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF fontfile code page".to_string())
                    })?;
                    code_page = Some(FontPage::new(page).ok_or_else(|| {
                        RtfError::MalformedDocument(format!(
                            "unsupported RTF fontfile code page {page}"
                        ))
                    })?);
                },
                Some(Token::Text(text)) => name.push_transport(text)?,
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    self.append_deferred_unicode(&mut name, *first, unicode_skip)?;
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_unicode(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::OpenBrace) | Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF fontfile destination contains unsupported content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if name.source_len() > crate::EmbeddedFont::MAX_FILE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF embedded font file name exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse color table.
    pub(super) fn parse_color_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \colortbl

        let mut current_red = 0;
        let mut current_green = 0;
        let mut current_blue = 0;
        let mut has_component = false;

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    // Retain a lenient unterminated final RGB entry, but do not
                    // invent a trailing black entry after the normal `;`.
                    if has_component {
                        let color = Color::new(current_red, current_green, current_blue);
                        self.color_table.borrow_mut().add(color);
                    }
                    return Ok(());
                },
                Token::Control(ControlWord::Red(r)) => {
                    current_red = (*r).clamp(0, 255) as u8;
                    has_component = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Green(g)) => {
                    current_green = (*g).clamp(0, 255) as u8;
                    has_component = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Blue(b)) => {
                    current_blue = (*b).clamp(0, 255) as u8;
                    has_component = true;
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control_in(token, crate::opaque::Context::Metadata)?;
                },
                Token::Text(text) if text.trim() == ";" => {
                    // An empty entry is the RTF automatic/default color.
                    if has_component {
                        let color = Color::new(current_red, current_green, current_blue);
                        self.color_table.borrow_mut().add(color);
                    } else {
                        self.color_table.borrow_mut().add_automatic();
                    }
                    current_red = 0;
                    current_green = 0;
                    current_blue = 0;
                    has_component = false;
                    self.pos += 1;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse the standard RTF `info` destination.
    pub(super) fn parse_info(&mut self) -> RtfResult<()> {
        if self.saw_info_group {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple info groups".to_string(),
            ));
        }
        self.saw_info_group = true;
        self.pos += 1; // `info`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    let control = self.tokens.get(self.pos).cloned();
                    match control {
                        Some(Token::Control(ControlWord::Title)) => {
                            self.parse_info_text(InfoTextField::Title)?;
                        },
                        Some(Token::Control(ControlWord::Subject)) => {
                            self.parse_info_text(InfoTextField::Subject)?;
                        },
                        Some(Token::Control(ControlWord::Author)) => {
                            self.parse_info_text(InfoTextField::Author)?;
                        },
                        Some(Token::Control(ControlWord::Manager)) => {
                            self.parse_info_text(InfoTextField::Manager)?;
                        },
                        Some(Token::Control(ControlWord::Company)) => {
                            self.parse_info_text(InfoTextField::Company)?;
                        },
                        Some(Token::Control(ControlWord::Operator)) => {
                            self.parse_info_text(InfoTextField::Operator)?;
                        },
                        Some(Token::Control(ControlWord::Category)) => {
                            self.parse_info_text(InfoTextField::Category)?;
                        },
                        Some(Token::Control(ControlWord::Keywords)) => {
                            self.parse_info_text(InfoTextField::Keywords)?;
                        },
                        Some(Token::Control(ControlWord::Comment)) => {
                            self.parse_info_text(InfoTextField::Comment)?;
                        },
                        Some(Token::Control(ControlWord::DocComment)) => {
                            self.parse_info_text(InfoTextField::DocumentComment)?;
                        },
                        Some(Token::Control(ControlWord::HyperlinkBase)) => {
                            self.parse_info_text(InfoTextField::HyperlinkBase)?;
                        },
                        Some(Token::Control(ControlWord::CreationTime)) => {
                            self.parse_info_time(InfoTimeField::Creation)?;
                        },
                        Some(Token::Control(ControlWord::RevisionTime)) => {
                            self.parse_info_time(InfoTimeField::Revision)?;
                        },
                        Some(Token::Control(ControlWord::PrintTime)) => {
                            self.parse_info_time(InfoTimeField::Print)?;
                        },
                        Some(Token::Control(ControlWord::BackupTime)) => {
                            self.parse_info_time(InfoTimeField::Backup)?;
                        },
                        Some(Token::Control(ControlWord::IgnorableDestination))
                            if matches!(
                                self.tokens.get(self.pos + 1),
                                Some(Token::Control(ControlWord::Password))
                            ) =>
                        {
                            self.parse_info_password()?;
                        },
                        Some(Token::Control(ControlWord::Password)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF password hash destination must be starred".to_string(),
                            ));
                        },
                        _ => self.skip_open_info_group()?,
                    }
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::Control(control)) => {
                    match control {
                        ControlWord::InfoVersion(value) => {
                            Self::set_info_number(&mut self.info.version, *value, "version")?
                        },
                        ControlWord::InfoRevision(value) => {
                            Self::set_info_number(&mut self.info.revision, *value, "vern")?
                        },
                        ControlWord::EditingTime(value) => {
                            Self::set_info_number(&mut self.info.editing_time, *value, "edmins")?
                        },
                        ControlWord::NumberOfPages(value) => {
                            Self::set_info_number(&mut self.info.pages, *value, "nofpages")?
                        },
                        ControlWord::NumberOfWords(value) => {
                            Self::set_info_number(&mut self.info.words, *value, "nofwords")?
                        },
                        ControlWord::NumberOfCharacters(value) => {
                            Self::set_info_number(&mut self.info.characters, *value, "nofchars")?;
                        },
                        ControlWord::NumberOfCharactersWithSpaces(value) => {
                            Self::set_info_number(
                                &mut self.info.characters_with_spaces,
                                *value,
                                "nofcharsws",
                            )?;
                        },
                        ControlWord::DocumentId(value) => {
                            Self::set_info_number(&mut self.info.id, *value, "id")?
                        },
                        _ => {},
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1
                },
                Some(Token::Text(_) | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF info group contains active text or binary data".to_string(),
                    ));
                },
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }
}
