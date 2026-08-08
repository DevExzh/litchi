use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_latent_styles(&mut self) -> RtfResult<crate::LatentStyles<'a>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LatentStyles))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF latentstyles destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut max_style_index = None;
        let mut locked_default = None;
        let mut semi_hidden_default = None;
        let mut unhide_when_used_default = None;
        let mut quick_format_default = None;
        let mut priority_default = None;
        let mut exceptions = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let styles = crate::LatentStyles {
                        max_style_index: max_style_index.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF latentstyles is missing lsdstimax".to_string(),
                            )
                        })?,
                        locked_default,
                        semi_hidden_default,
                        unhide_when_used_default,
                        quick_format_default,
                        priority_default,
                        exceptions: exceptions.unwrap_or_default(),
                    };
                    styles.validate()?;
                    return Ok(styles);
                },
                Some(Token::OpenBrace) => {
                    if exceptions.is_some()
                        || !matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::LatentStyleExceptions))
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latentstyles contains a duplicate or active nested destination"
                                .to_string(),
                        ));
                    }
                    exceptions = Some(self.parse_latent_style_exceptions()?);
                },
                Some(Token::Control(control)) => {
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    concat!("duplicate RTF latent-style ", $name).to_string(),
                                ));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::LatentStyleMax(value) => {
                            let value = u32::try_from(*value).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF lsdstimax cannot be negative".to_string(),
                                )
                            })?;
                            if value > crate::latent_style::MAX_LATENT_STYLE_INDEX {
                                return Err(RtfError::MalformedDocument(
                                    "RTF lsdstimax exceeds 65535".to_string(),
                                ));
                            }
                            set_once!(max_style_index, value, "lsdstimax");
                        },
                        ControlWord::LatentStyleLockedDefault(value) => set_once!(
                            locked_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdlockeddef"
                        ),
                        ControlWord::LatentStyleSemiHiddenDefault(value) => set_once!(
                            semi_hidden_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdsemihiddendef"
                        ),
                        ControlWord::LatentStyleUnhideUsedDefault(value) => set_once!(
                            unhide_when_used_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdunhideuseddef"
                        ),
                        ControlWord::LatentStyleQuickFormatDefault(value) => set_once!(
                            quick_format_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdqformatdef"
                        ),
                        ControlWord::LatentStylePriorityDefault(value) => set_once!(
                            priority_default,
                            Self::parse_latent_style_priority(*value)?,
                            "lsdprioritydef"
                        ),
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF latentstyles contains an unsupported control".to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Binary(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latentstyles contains orphan text or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_latent_style_exceptions(
        &mut self,
    ) -> RtfResult<Vec<crate::LatentStyleException<'a>>> {
        self.expect_token(Token::OpenBrace)?;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LatentStyleExceptions))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF lsdlockedexcept destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut entries = Vec::new();
        let mut builder = LatentStyleExceptionBuilder::default();
        let mut name = String::new();
        let mut text_bytes = 0usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if !name.trim().is_empty()
                        || builder.locked.is_some()
                        || builder.semi_hidden.is_some()
                        || builder.unhide_when_used.is_some()
                        || builder.quick_format.is_some()
                        || builder.priority.is_some()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latent-style exception is missing its terminating semicolon"
                                .to_string(),
                        ));
                    }
                    return Ok(entries);
                },
                Some(Token::Control(control)) => {
                    if matches!(
                        control,
                        ControlWord::LatentStyleLocked(_)
                            | ControlWord::LatentStyleSemiHidden(_)
                            | ControlWord::LatentStyleUnhideUsed(_)
                            | ControlWord::LatentStyleQuickFormat(_)
                            | ControlWord::LatentStylePriority(_)
                    ) && !name.trim().is_empty()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latent-style properties must precede the style name".to_string(),
                        ));
                    }
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    concat!("duplicate RTF latent-style exception ", $name)
                                        .to_string(),
                                ));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::LatentStyleLocked(value) => set_once!(
                            builder.locked,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdlocked"
                        ),
                        ControlWord::LatentStyleSemiHidden(value) => set_once!(
                            builder.semi_hidden,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdsemihidden"
                        ),
                        ControlWord::LatentStyleUnhideUsed(value) => set_once!(
                            builder.unhide_when_used,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdunhideused"
                        ),
                        ControlWord::LatentStyleQuickFormat(value) => set_once!(
                            builder.quick_format,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdqformat"
                        ),
                        ControlWord::LatentStylePriority(value) => set_once!(
                            builder.priority,
                            Self::parse_latent_style_priority(*value)?,
                            "lsdpriority"
                        ),
                        ControlWord::Unicode(first) => {
                            name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                            self.drain_latent_style_exception_names(
                                &mut name,
                                &mut builder,
                                &mut entries,
                                &mut text_bytes,
                            )?;
                            if name.len() > crate::latent_style::MAX_LATENT_STYLE_NAME_BYTES {
                                return Err(RtfError::MalformedDocument(
                                    "RTF latent-style exception name exceeds the safety limit"
                                        .to_string(),
                                ));
                            }
                            continue;
                        },
                        ControlWord::UnicodeSkip(value) => {
                            unicode_skip = (*value).max(0);
                            self.pos += 1;
                            continue;
                        },
                        control if control_symbol_text(control).is_some() => {
                            name.push_str(control_symbol_text(control).unwrap_or_default());
                            self.pos += 1;
                            continue;
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF latent-style exception contains an unsupported control"
                                    .to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    name.push_str(&self.decode_transport_text(text)?);
                    self.pos += 1;
                    self.drain_latent_style_exception_names(
                        &mut name,
                        &mut builder,
                        &mut entries,
                        &mut text_bytes,
                    )?;
                },
                Some(Token::OpenBrace | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latent-style exceptions cannot contain nesting or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if name.len() > crate::latent_style::MAX_LATENT_STYLE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style exception name exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn drain_latent_style_exception_names(
        &self,
        name: &mut String,
        builder: &mut LatentStyleExceptionBuilder,
        entries: &mut Vec<crate::LatentStyleException<'a>>,
        text_bytes: &mut usize,
    ) -> RtfResult<()> {
        let Some(last_separator) = name.rfind(';') else {
            return Ok(());
        };
        let completed = name.get(..=last_separator).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF latent-style parser found an invalid text boundary".to_string(),
            )
        })?;
        let completed = completed.strip_suffix(';').ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF latent-style parser lost an exception delimiter".to_string(),
            )
        })?;
        for entry_name in completed.split(';') {
            let entry_name = entry_name.trim();
            let candidate = crate::LatentStyleException {
                name: Cow::Borrowed(entry_name),
                locked: builder.locked,
                semi_hidden: builder.semi_hidden,
                unhide_when_used: builder.unhide_when_used,
                quick_format: builder.quick_format,
                priority: builder.priority,
            };
            candidate.validate()?;
            if entries.len() >= crate::latent_style::MAX_LATENT_STYLE_EXCEPTIONS {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style exception count exceeds the safety limit".to_string(),
                ));
            }
            let next_text_bytes = text_bytes.checked_add(entry_name.len()).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF latent-style aggregate text size overflow".to_string(),
                )
            })?;
            if next_text_bytes > crate::latent_style::MAX_LATENT_STYLE_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style text exceeds the safety limit".to_string(),
                ));
            }
            crate::error::try_reserve_one(entries, "latent-style exceptions")?;
            entries.push(crate::LatentStyleException {
                name: Cow::Borrowed(self.arena.alloc_str(entry_name)),
                locked: builder.locked.take(),
                semi_hidden: builder.semi_hidden.take(),
                unhide_when_used: builder.unhide_when_used.take(),
                quick_format: builder.quick_format.take(),
                priority: builder.priority.take(),
            });
            *text_bytes = next_text_bytes;
        }
        drop(name.drain(..=last_separator));
        Ok(())
    }

    pub(super) fn parse_latent_style_bool(value: i32) -> RtfResult<bool> {
        match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(RtfError::MalformedDocument(
                "RTF latent-style Boolean values must be 0 or 1".to_string(),
            )),
        }
    }

    pub(super) fn parse_latent_style_priority(value: i32) -> RtfResult<u8> {
        u8::try_from(value)
            .ok()
            .filter(|priority| *priority <= 99)
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF latent-style priority must be in 0..=99".to_string(),
                )
            })
    }

    pub(super) fn parse_theme_hex_destination(
        &mut self,
        expected: ControlWord<'a>,
        limit: usize,
    ) -> RtfResult<Vec<u8>> {
        self.pos += 1; // ignorable-destination marker
        let matches_expected = matches!(
            (&expected, self.tokens.get(self.pos)),
            (
                ControlWord::ThemeData,
                Some(Token::Control(ControlWord::ThemeData))
            ) | (
                ControlWord::ColorSchemeMapping,
                Some(Token::Control(ControlWord::ColorSchemeMapping))
            )
        );
        if !matches_expected {
            return Err(RtfError::MalformedDocument(
                "invalid RTF theme destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut data = Vec::new();
        let mut high = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF theme payload has an odd hexadecimal digit count".to_string(),
                        ));
                    }
                    if data.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF theme payload cannot be empty".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Some(Token::Text(text)) => {
                    for byte in text.as_bytes() {
                        if byte.is_ascii_whitespace() {
                            continue;
                        }
                        let nibble = match byte {
                            b'0'..=b'9' => byte - b'0',
                            b'a'..=b'f' => byte - b'a' + 10,
                            b'A'..=b'F' => byte - b'A' + 10,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF theme payload contains a non-hexadecimal character"
                                        .to_string(),
                                ));
                            },
                        };
                        if let Some(first) = high.take() {
                            data.push(first << 4 | nibble);
                        } else {
                            high = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF theme payload cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > limit {
                return Err(RtfError::MalformedDocument(
                    "RTF theme payload exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_xml_namespace_table(&mut self) -> RtfResult<()> {
        if self.saw_xml_namespace_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple XML namespace tables".to_string(),
            ));
        }
        self.saw_xml_namespace_table = true;
        self.pos += 1; // xmlnstbl
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_xml_namespace_entry()?,
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Binary(_)) | Some(Token::Control(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace table contains ungrouped, active, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_window_caption_destination(
        &mut self,
        starred: bool,
    ) -> RtfResult<crate::DocumentWindowCaption<'a>> {
        if starred {
            self.pos += 1;
        }
        if self.tokens.get(self.pos) != Some(&Token::Control(ControlWord::WindowCaption)) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF window caption destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let value = self.arena.alloc_str(&value);
                    return crate::DocumentWindowCaption::new(Cow::Borrowed(value));
                },
                Some(Token::Text(text)) => {
                    value.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(skip))) => {
                    unicode_skip = (*skip).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF window caption contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > crate::MAX_WINDOW_CAPTION_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF window caption exceeds the resource limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_xsl_transform_destination(
        &mut self,
    ) -> RtfResult<crate::DocumentXslTransform<'a>> {
        self.pos += 1;
        if self.tokens.get(self.pos) != Some(&Token::Control(ControlWord::XslTransform)) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF XSL transform destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut location = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let location = self.arena.alloc_str(&location);
                    return crate::DocumentXslTransform::new(Cow::Borrowed(location));
                },
                Some(Token::Text(text)) => {
                    location.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    location.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(skip))) => {
                    unicode_skip = (*skip).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    location.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XSL transform contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if location.len() > crate::MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF XSL transform location exceeds the resource limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_style_list_filter_destination(
        &mut self,
        parameter: Option<i32>,
    ) -> RtfResult<crate::DocumentStyleListFilter> {
        self.pos += 1;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::StyleListFilter(_)))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF style-list filter destination".to_string(),
            ));
        }
        if parameter.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF style-list filter must be a delimited four-digit hexadecimal string"
                    .to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if value.len() != 4 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                        return Err(RtfError::MalformedDocument(
                            "RTF style-list filter must contain exactly four hexadecimal digits"
                                .to_string(),
                        ));
                    }
                    let bits = u16::from_str_radix(&value, 16).map_err(|_| {
                        RtfError::MalformedDocument(
                            "invalid RTF style-list filter hexadecimal value".to_string(),
                        )
                    })?;
                    return Ok(crate::DocumentStyleListFilter::from_parsed_bits(bits));
                },
                Some(Token::Text(text)) => {
                    value.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF style-list filter contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > 4 {
                return Err(RtfError::MalformedDocument(
                    "RTF style-list filter exceeds its four-digit resource bound".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_legacy_write_reservation_destination(
        &mut self,
    ) -> RtfResult<crate::LegacyWriteReservation<'a>> {
        self.pos += 1;
        match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::WriteReservation(None))) => {},
            Some(Token::Control(ControlWord::WriteReservation(Some(_)))) => {
                return Err(RtfError::MalformedDocument(
                    "RTF writereservation must not have a numeric parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF legacy write-reservation destination".to_string(),
                ));
            },
        }
        self.pos += 1;
        let mut data = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let data = self.arena.alloc_str(&data);
                    return crate::LegacyWriteReservation::new(Cow::Borrowed(data));
                },
                Some(Token::Text(text)) => {
                    data.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    data.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(skip))) => {
                    unicode_skip = (*skip).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    data.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy write reservation contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > crate::MAX_WRITE_RESERVATION_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF legacy write-reservation payload exceeds the resource limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_write_reservation_hash_destination(
        &mut self,
    ) -> RtfResult<crate::WriteReservationHash<'a>> {
        self.pos += 1;
        match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::WriteReservationHash(None))) => {},
            Some(Token::Control(ControlWord::WriteReservationHash(Some(_)))) => {
                return Err(RtfError::MalformedDocument(
                    "RTF writereservhash must not have a numeric parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF write-reservation hash destination".to_string(),
                ));
            },
        }
        self.pos += 1;
        let mut encoded = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if encoded.is_empty()
                        || !encoded.len().is_multiple_of(2)
                        || !encoded.bytes().all(|byte| byte.is_ascii_hexdigit())
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF write-reservation hash must contain an even number of hexadecimal digits"
                                .to_string(),
                        ));
                    }
                    let mut data = Vec::with_capacity(encoded.len() / 2);
                    for pair in encoded.as_bytes().chunks_exact(2) {
                        let pair = std::str::from_utf8(pair).map_err(|_| {
                            RtfError::MalformedDocument(
                                "invalid RTF write-reservation hash encoding".to_string(),
                            )
                        })?;
                        data.push(u8::from_str_radix(pair, 16).map_err(|_| {
                            RtfError::MalformedDocument(
                                "invalid RTF write-reservation hash encoding".to_string(),
                            )
                        })?);
                    }
                    return crate::WriteReservationHash::new(Cow::Owned(data));
                },
                Some(Token::Text(text)) => {
                    encoded.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF write-reservation hash contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if encoded.len() > crate::MAX_WRITE_RESERVATION_BYTES * 2 {
                return Err(RtfError::MalformedDocument(
                    "RTF write-reservation hash exceeds the resource limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_external_reference_destination(
        &mut self,
        expected: ControlWord<'a>,
    ) -> RtfResult<Cow<'a, str>> {
        self.pos += 1;
        if self.tokens.get(self.pos) != Some(&Token::Control(expected)) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF external document reference destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let value = self.arena.alloc_str(value.trim());
                    if value.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF external document reference name cannot be empty".to_string(),
                        ));
                    }
                    return Ok(Cow::Borrowed(value));
                },
                Some(Token::Text(text)) => {
                    value.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(skip))) => {
                    unicode_skip = (*skip).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF external document reference contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > crate::MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF external document reference exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_protection_user_table(&mut self) -> RtfResult<()> {
        if self.saw_protection_user_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple protection-user tables".to_string(),
            ));
        }
        self.saw_protection_user_table = true;
        self.pos += 1; // protusertbl
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if self.protection_users.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF protection-user table cannot be empty".to_string(),
                        ));
                    }
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_protection_user_entry()?,
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Binary(_)) | Some(Token::Control(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF protection-user table contains ungrouped, active, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_protection_user_entry(&mut self) -> RtfResult<()> {
        if self.protection_users.len() >= crate::MAX_PROTECTION_USERS {
            return Err(RtfError::MalformedDocument(
                "RTF protection-user count exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(Token::OpenBrace)?;
        let mut name = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let user = crate::ProtectionUser::new(Cow::Borrowed(
                        self.arena.alloc_str(name.trim()),
                    ))?;
                    self.protection_user_text_bytes = self
                        .protection_user_text_bytes
                        .checked_add(user.name.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF protection-user aggregate size overflow".to_string(),
                            )
                        })?;
                    if self.protection_user_text_bytes > crate::MAX_PROTECTION_USER_TOTAL_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF protection-user aggregate text exceeds the safety limit"
                                .to_string(),
                        ));
                    }
                    self.protection_users.push(user);
                    return Ok(());
                },
                Some(Token::Text(text)) => {
                    name.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF protection-user entry contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if name.len() > crate::MAX_PROTECTION_USER_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF protection username exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_xml_namespace_entry(&mut self) -> RtfResult<()> {
        if self.xml_namespaces.len() >= crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace count exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(Token::OpenBrace)?;
        let id = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::XmlNamespace(value))) => {
                let value = u32::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF XML namespace ID must be a positive signed integer".to_string(),
                    )
                })?;
                if value == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace ID must be a positive signed integer".to_string(),
                    ));
                }
                value
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace entry is missing xmlnsN".to_string(),
                ));
            },
        };
        if self.xml_namespaces.iter().any(|entry| entry.id == id) {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace IDs must be unique".to_string(),
            ));
        }
        self.pos += 1;
        let mut namespace = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let namespace = namespace.trim();
                    let entry = crate::XmlNamespace::new(
                        id,
                        Cow::Borrowed(self.arena.alloc_str(namespace)),
                    )?;
                    self.xml_namespace_text_bytes = self
                        .xml_namespace_text_bytes
                        .checked_add(entry.namespace.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF XML namespace aggregate size overflow".to_string(),
                            )
                        })?;
                    if self.xml_namespace_text_bytes
                        > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF XML namespace aggregate text exceeds the safety limit".to_string(),
                        ));
                    }
                    self.xml_namespaces.push(entry);
                    return Ok(());
                },
                Some(Token::Text(text)) => {
                    namespace.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    namespace.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    namespace.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace entry contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if namespace.len() > crate::xml_namespace::MAX_XML_NAMESPACE_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace identifier exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_generator_destination(
        &mut self,
    ) -> RtfResult<crate::DocumentGenerator<'a>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::Generator))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF generator destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let value = value.trim();
                    let value = value.strip_suffix(';').unwrap_or(value).trim_end();
                    let value = self.arena.alloc_str(value);
                    return crate::DocumentGenerator::new(Cow::Borrowed(value));
                },
                Some(Token::Text(text)) => {
                    value.push_str(&self.decode_transport_text(text)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generator destination contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > crate::generator::MAX_GENERATOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF generator value exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_revision_author_group(&mut self) -> RtfResult<String> {
        self.pos += 1; // opening brace
        let mut author = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(author
                        .trim_end_matches(['\r', '\n', ' '])
                        .strip_suffix(';')
                        .unwrap_or(author.trim_end_matches(['\r', '\n', ' ']))
                        .trim()
                        .to_string());
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision author contains a nested destination".to_string(),
                    ));
                },
                Some(Token::Text(text)) => author.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    author.push_str(&decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    author.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision author contains a non-text control or binary data"
                            .to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if author.len() > MAX_REVISION_AUTHOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF revision author exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn push_direct_revision_authors(&mut self, text: &mut String) -> RtfResult<()> {
        let Some(last_separator) = text.rfind(';') else {
            return Ok(());
        };
        let completed = text.get(..=last_separator).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF revision-author parser found an invalid text boundary".to_string(),
            )
        })?;
        let completed = completed.strip_suffix(';').ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF revision-author parser lost an author delimiter".to_string(),
            )
        })?;
        for author in completed.split(';') {
            self.push_revision_author(author.trim().to_string())?;
        }
        drop(text.drain(..=last_separator));
        Ok(())
    }

    pub(super) fn push_revision_author(&mut self, author: String) -> RtfResult<()> {
        if self.revision_authors.len() >= MAX_REVISION_AUTHORS {
            return Err(RtfError::MalformedDocument(
                "RTF revision author count exceeds the safety limit".to_string(),
            ));
        }
        let revision_author_text_bytes = self
            .revision_author_text_bytes
            .checked_add(author.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aggregate revision-author size overflow".to_string(),
                )
            })?;
        if revision_author_text_bytes
            > super::super::super::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES
        {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision-author text exceeds the safety limit".to_string(),
            ));
        }
        crate::error::try_reserve_one(&mut self.revision_authors, "revision authors")?;
        let author = super::super::super::annotation::RevisionAuthor::new(Cow::Borrowed(
            self.arena.alloc_str(&author),
        ))?;
        author.validate()?;
        self.revision_authors.push(author);
        self.revision_author_text_bytes = revision_author_text_bytes;
        Ok(())
    }

    pub(super) fn parse_list_table(&mut self) -> RtfResult<()> {
        if self.saw_list_table {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple list tables".to_string(),
            ));
        }
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF list table must occur in the root header".to_string(),
            ));
        }
        self.saw_list_table = true;
        self.pos += 1; // `listtable`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::List))
                    ) =>
                {
                    self.parse_list_definition()?;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ListPicture(None)),
                        ])
                    ) =>
                {
                    if self.list_table.picture_bullet_count != 0 {
                        return Err(RtfError::MalformedDocument(
                            "RTF list table contains multiple listpicture destinations".to_string(),
                        ));
                    }
                    let indices = self.parse_list_picture_destination()?;
                    self.list_table
                        .set_picture_bullet_picture_indices(indices)?;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ListPicture(Some(_))),
                        ])
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF listpicture does not accept a parameter".to_string(),
                    ));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListPicture(_)))
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF listpicture destination must be starred".to_string(),
                    ));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapePicture(_)),
                        ])
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF shppict must be nested in listpicture".to_string(),
                    ));
                },
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.list_table.validate()?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_list_picture_destination(&mut self) -> RtfResult<Vec<Option<usize>>> {
        self.pos += 3; // opening brace, \*, and \listpicture
        let mut indices = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::Text(text)) if text.chars().all(char::is_whitespace) => self.pos += 1,
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapePicture(None)),
                        ])
                    ) =>
                {
                    if indices.len() >= 65_536 {
                        return Err(RtfError::MalformedDocument(
                            "RTF list-picture record count exceeds the safety limit".to_string(),
                        ));
                    }
                    indices.push(Some(self.parse_list_shape_picture_destination()?));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapePicture(Some(_))),
                        ])
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF shppict does not accept a parameter".to_string(),
                    ));
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if indices.is_empty() {
                        indices.push(None);
                    }
                    return Ok(indices);
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF listpicture contains content outside a shppict record".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_list_shape_picture_destination(&mut self) -> RtfResult<usize> {
        self.pos += 3; // opening brace, \*, and \shppict
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF shppict contains text outside pict".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([Token::OpenBrace, Token::Control(ControlWord::Picture)])
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF list shppict must contain exactly one pict destination".to_string(),
            ));
        }
        self.pos += 1; // pict opening brace
        let index = self.pictures.len();
        self.parse_picture()?;
        if self.pictures.len() != index + 1 {
            return Err(RtfError::MalformedDocument(
                "RTF list picture payload cannot be empty".to_string(),
            ));
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF list pict destination is not closed".to_string(),
            ));
        }
        self.pos += 1;
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF shppict contains trailing text".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF list shppict must contain exactly one pict destination".to_string(),
            ));
        }
        self.pos += 1;
        Ok(index)
    }

    pub(super) fn parse_list_definition(&mut self) -> RtfResult<()> {
        self.pos += 2; // opening brace and `list`
        let mut list = super::super::super::list::List::new(0);
        list.simple = false;
        let mut has_id = false;
        let mut has_template_id = false;
        let mut closed = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListLevel))
                    ) =>
                {
                    if list.levels.len() >= MAX_LIST_LEVELS {
                        return Err(RtfError::MalformedDocument(
                            "RTF list exceeds the nine-level specification limit".to_string(),
                        ));
                    }
                    let level = self.parse_list_level(list.levels.len() as u8)?;
                    list.add_level(level);
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListName))
                    ) =>
                {
                    let name = self.parse_list_text_group(true, false)?;
                    list.name = Cow::Borrowed(self.arena.alloc_str(&name));
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ListStyleName),
                        ])
                    ) =>
                {
                    let name = self.parse_list_text_group(true, false)?;
                    list.style_name = Cow::Borrowed(self.arena.alloc_str(&name));
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ListTemplateId(value) => {
                        list.template_id = *value;
                        has_template_id = true;
                    },
                    ControlWord::ListSimple(value) => list.simple = *value,
                    ControlWord::ListHybrid(value) => list.hybrid = *value,
                    ControlWord::ListId(value) => {
                        list.id = *value;
                        has_id = true;
                    },
                    ControlWord::StylePriority(value) => list.style_priority = Some(*value),
                    _ => {},
                },
                Some(_) => {},
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if list.simple && (list.hybrid || list.levels.len() > 1) {
            return Err(RtfError::MalformedDocument(
                "invalid simple RTF list definition".to_string(),
            ));
        }
        if !has_template_id {
            list.template_id = list.id;
        }
        if has_id {
            if self.list_table.lists().len() >= MAX_LISTS {
                return Err(RtfError::MalformedDocument(
                    "RTF list count exceeds the safety limit".to_string(),
                ));
            }
            self.list_table.add(list);
        }
        Ok(())
    }

    pub(super) fn parse_list_level(
        &mut self,
        level_index: u8,
    ) -> RtfResult<super::super::super::list::ListLevel<'a>> {
        self.pos += 2; // opening brace and `listlevel`
        let mut level = super::super::super::list::ListLevel::new(level_index);
        let mut explicit_indent = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListNumberText))
                    ) =>
                {
                    let text = self.parse_list_text_group(false, true)?;
                    level.number_text = Cow::Borrowed(self.arena.alloc_str(&text));
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListLevelNumbers))
                    ) =>
                {
                    let positions = self.parse_list_text_group(false, false)?;
                    level.number_positions = Cow::Borrowed(self.arena.alloc_str(&positions));
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(level);
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ListLevelType(value) => {
                        level.level_type = Self::list_level_type(*value);
                    },
                    ControlWord::ListLevelJustification(value) => {
                        level.justification = match value {
                            1 => super::super::super::list::ListJustification::Center,
                            2 => super::super::super::list::ListJustification::Right,
                            _ => super::super::super::list::ListJustification::Left,
                        };
                    },
                    ControlWord::ListLevelFollow(value) => {
                        level.follow = match value {
                            1 => super::super::super::list::ListFollow::Space,
                            2 => super::super::super::list::ListFollow::Nothing,
                            _ => super::super::super::list::ListFollow::Tab,
                        };
                        level.follow_previous = *value != 0;
                    },
                    ControlWord::ListLevelStartAt(value) => level.start_at = *value,
                    ControlWord::ListLevelSpace(value) => level.space = *value,
                    ControlWord::ListLevelIndent(value) => {
                        level.indent = *value;
                        explicit_indent = true;
                    },
                    ControlWord::FontNumber(value) => {
                        level.font_ref = u16::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument(
                                "RTF list font reference is outside the supported range"
                                    .to_string(),
                            )
                        })?;
                    },
                    ControlWord::LeftIndent(value) => {
                        level.left_indent = Some(*value);
                        if !explicit_indent {
                            level.indent = *value;
                        }
                    },
                    ControlWord::FirstLineIndent(value) => level.first_line_indent = Some(*value),
                    ControlWord::TabPosition(value) => {
                        let value = value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF list-level tx control requires a numeric parameter"
                                    .to_string(),
                            )
                        })?;
                        if level.tabs.len() >= MAX_LIST_TABS {
                            return Err(RtfError::MalformedDocument(
                                "RTF list level has too many tabs".to_string(),
                            ));
                        }
                        level.tabs.push(value);
                    },
                    ControlWord::ListLevelTentative => level.tentative = true,
                    ControlWord::ListLevelLegal(value) => level.legal_format = *value,
                    ControlWord::ListLevelNoRestart(value) => level.no_restart = *value,
                    ControlWord::ListLevelOld(value) => level.legacy = *value,
                    ControlWord::ListLevelPrevious(value) => level.include_previous = *value,
                    ControlWord::ListLevelPreviousSpace(value) => {
                        level.include_previous_space = *value
                    },
                    ControlWord::ListLevelTemplateId(value) => level.template_id = Some(*value),
                    ControlWord::ListLevelPicture(value) => {
                        level.picture_index = Some(u32::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument(
                                "RTF list picture index cannot be negative".to_string(),
                            )
                        })?);
                    },
                    _ => {},
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_list_text_group(
        &mut self,
        is_name: bool,
        strip_length: bool,
    ) -> RtfResult<String> {
        self.pos += if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            3
        } else {
            2
        };
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let trimmed = value.trim_end_matches(['\r', '\n', ' ']);
                    let trimmed = trimmed.strip_suffix(';').unwrap_or(trimmed);
                    if is_name {
                        return Ok(trimmed.trim().to_string());
                    }
                    if strip_length {
                        let mut chars = trimmed.chars();
                        if chars
                            .next()
                            .is_some_and(|ch| u32::from(ch) <= u8::MAX.into())
                        {
                            return Ok(chars.collect());
                        }
                    }
                    return Ok(trimmed.to_string());
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    value.push_str(&decoded);
                    if value.len() > MAX_LIST_TEXT_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF list text exceeds the safety limit".to_string(),
                        ));
                    }
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
            if value.len() > MAX_LIST_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF list text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn list_level_type(value: i32) -> super::super::super::list::ListLevelType {
        match value {
            0 => super::super::super::list::ListLevelType::Decimal,
            1 => super::super::super::list::ListLevelType::UpperRoman,
            2 => super::super::super::list::ListLevelType::LowerRoman,
            3 => super::super::super::list::ListLevelType::UpperLetter,
            4 => super::super::super::list::ListLevelType::LowerLetter,
            5 => super::super::super::list::ListLevelType::Ordinal,
            6 => super::super::super::list::ListLevelType::CardinalText,
            7 => super::super::super::list::ListLevelType::OrdinalText,
            23 => super::super::super::list::ListLevelType::Bullet,
            255 => super::super::super::list::ListLevelType::None,
            other => super::super::super::list::ListLevelType::Other(other),
        }
    }

    pub(super) fn parse_list_override_table(&mut self) -> RtfResult<()> {
        if self.saw_list_override_table {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple list override tables".to_string(),
            ));
        }
        if !self.saw_list_table
            || self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF list override table must follow listtable in the root header".to_string(),
            ));
        }
        self.saw_list_override_table = true;
        self.pos += 1; // `listoverridetable`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListOverride))
                    ) =>
                {
                    self.parse_list_override()?;
                },
                Some(Token::OpenBrace) => self.skip_group()?,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.list_override_table.validate(&self.list_table)?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_list_override(&mut self) -> RtfResult<()> {
        self.pos += 2; // opening brace and `listoverride`
        let mut list_id = None;
        let mut index = None;
        let mut level_count = None;
        let mut start_at = None;
        let mut override_levels = Vec::new();
        let mut closed = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListOverrideLevel))
                    ) =>
                {
                    self.pos += 2;
                    let mut has_start_override = false;
                    let mut has_format_override = false;
                    let mut level_start_at = None;
                    let override_index = u8::try_from(override_levels.len()).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF list override has too many levels".to_string(),
                        )
                    })?;
                    while self.pos < self.tokens.len() {
                        match self.tokens.get(self.pos) {
                            Some(Token::CloseBrace) => {
                                self.pos += 1;
                                break;
                            },
                            Some(Token::Control(ControlWord::ListOverrideStartAt(value))) => {
                                has_start_override = *value;
                            },
                            Some(Token::Control(ControlWord::ListOverrideFormat(value))) => {
                                has_format_override = *value;
                            },
                            Some(Token::Control(ControlWord::ListLevelStartAt(value)))
                                if has_start_override =>
                            {
                                level_start_at = Some(*value);
                                start_at = Some(*value);
                            },
                            Some(Token::OpenBrace) => {
                                self.skip_group()?;
                                continue;
                            },
                            Some(_) => {},
                            None => return Err(RtfError::UnexpectedEof),
                        }
                        self.pos += 1;
                    }
                    let level_start = if has_start_override {
                        level_start_at
                    } else {
                        None
                    };
                    override_levels.push(super::super::super::list::ListOverrideLevel {
                        level: override_index,
                        start_at: level_start,
                        format_override: has_format_override,
                    });
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Some(Token::Control(ControlWord::ListId(value))) => list_id = Some(*value),
                Some(Token::Control(ControlWord::ListOverrideIndex(value))) => index = Some(*value),
                Some(Token::Control(ControlWord::ListOverrideCount(value))) => {
                    level_count = Some(u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF list override count is outside the supported range".to_string(),
                        )
                    })?);
                },
                Some(_) => {},
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if let (Some(index), Some(list_id)) = (index, list_id) {
            if self.list_override_table.overrides().len() >= MAX_LISTS {
                return Err(RtfError::MalformedDocument(
                    "RTF list override count exceeds the safety limit".to_string(),
                ));
            }
            let mut entry = super::super::super::list::ListOverride::new(index, list_id);
            entry.level_count_override = level_count;
            entry.start_at_override = start_at;
            entry.levels = override_levels;
            self.list_override_table.add(entry);
        }
        Ok(())
    }

    /// Parse the standard RTF stylesheet destination.
    pub(super) fn default_character_property_key(
        control: &ControlWord<'_>,
    ) -> Option<&'static str> {
        Some(match control {
            ControlWord::CharacterStyle(_) => "character-style",
            ControlWord::FontNumber(_) => "font",
            ControlWord::FontSize(_) => "font-size",
            ControlWord::AssociatedFontNumber(_) => "associated-font",
            ControlWord::AssociatedFontSize(_) => "associated-font-size",
            ControlWord::AssociatedLanguage(_) => "associated-language",
            ControlWord::AssociatedBold(_) => "associated-bold",
            ControlWord::AssociatedAllCaps(_) => "associated-caps",
            ControlWord::AssociatedColor(_) => "associated-color",
            ControlWord::AssociatedBaselineDown(_) | ControlWord::AssociatedBaselineUp(_) => {
                "associated-baseline"
            },
            ControlWord::AssociatedExpansion(_) => "associated-expansion",
            ControlWord::AssociatedItalic(_) => "associated-italic",
            ControlWord::AssociatedOutline(_) => "associated-outline",
            ControlWord::AssociatedSmallCaps(_) => "associated-small-caps",
            ControlWord::AssociatedShadow(_) => "associated-shadow",
            ControlWord::AssociatedStrike(_) => "associated-strike",
            ControlWord::AssociatedUnderline(_)
            | ControlWord::AssociatedUnderlineDotted(_)
            | ControlWord::AssociatedUnderlineDouble(_)
            | ControlWord::AssociatedUnderlineNone(_)
            | ControlWord::AssociatedUnderlineWords(_) => "associated-underline",
            ControlWord::ColorForeground(_) => "color",
            ControlWord::ColorBackground(_) => "background-color",
            ControlWord::Highlight(_) => "highlight",
            ControlWord::Bold(_) => "bold",
            ControlWord::Italic(_) => "italic",
            ControlWord::Underline(_)
            | ControlWord::UnderlineNone
            | ControlWord::UnderlineDouble
            | ControlWord::UnderlineDotted
            | ControlWord::UnderlineDashed
            | ControlWord::UnderlineDashDot
            | ControlWord::UnderlineDashDotDot
            | ControlWord::UnderlineWords
            | ControlWord::UnderlineThick
            | ControlWord::UnderlineWave => "underline",
            ControlWord::Strike(_) => "strike",
            ControlWord::DoubleStrike(_) => "double-strike",
            ControlWord::Superscript(_)
            | ControlWord::Subscript(_)
            | ControlWord::NoSuperSub
            | ControlWord::BaselineUp(_)
            | ControlWord::BaselineDown(_) => "baseline",
            ControlWord::SmallCaps(_) => "small-caps",
            ControlWord::AllCaps(_) => "caps",
            ControlWord::Hidden(_) => "hidden",
            ControlWord::Outline(_) => "outline",
            ControlWord::Shadow(_) => "shadow",
            ControlWord::Emboss(_) => "emboss",
            ControlWord::Imprint(_) => "imprint",
            ControlWord::CharSpacing(_) | ControlWord::CharSpacingTwips(_) => "expansion",
            ControlWord::CharScale(_) => "scale",
            ControlWord::Kerning(_) => "kerning",
            ControlWord::Language(_) => "language",
            ControlWord::LanguageEastAsian(_) => "language-east-asian",
            ControlWord::LanguageNoProof(_) => "language-no-proof",
            ControlWord::LanguageEastAsianNoProof(_) => "language-east-asian-no-proof",
            ControlWord::NoProof(_) => "no-proof",
            ControlWord::LeftToRightCharacter | ControlWord::RightToLeftCharacter => "direction",
            ControlWord::FontComplexScript(_) => "complex-script",
            ControlWord::CharacterGrid(_) => "character-grid",
            ControlWord::AnimatedText(_) => "animated-text",
            ControlWord::EmphasisMark(..) => "emphasis-mark",
            _ => return None,
        })
    }

    pub(super) fn default_paragraph_property_key(
        control: &ControlWord<'_>,
    ) -> Option<&'static str> {
        Some(match control {
            ControlWord::ParagraphStyle(_) => "paragraph-style",
            ControlWord::LeftAlign
            | ControlWord::RightAlign
            | ControlWord::Center
            | ControlWord::Justify => "alignment",
            ControlWord::LeftToRightParagraph | ControlWord::RightToLeftParagraph => "direction",
            ControlWord::SpaceBefore(_) => "space-before",
            ControlWord::SpaceAfter(_) => "space-after",
            ControlWord::SpaceBetween(_) => "line-space",
            ControlWord::LineMultiple(_) => "line-multiple",
            ControlWord::SpaceBeforeAuto(_) => "space-before-auto",
            ControlWord::SpaceAfterAuto(_) => "space-after-auto",
            ControlWord::ListSpaceBefore(_) => "list-space-before",
            ControlWord::ListSpaceAfter(_) => "list-space-after",
            ControlWord::NoSnapLineGrid(_) => "snap-grid",
            ControlWord::ContextualSpacing(_) => "contextual-space",
            ControlWord::LeftIndent(_) => "left-indent",
            ControlWord::RightIndent(_) => "right-indent",
            ControlWord::FirstLineIndent(_) => "first-indent",
            ControlWord::LogicalLeftIndent(_) => "logical-left",
            ControlWord::LogicalRightIndent(_) => "logical-right",
            ControlWord::CharacterFirstLineIndent(_) => "char-first",
            ControlWord::CharacterLeftIndent(_) => "char-left",
            ControlWord::CharacterRightIndent(_) => "char-right",
            ControlWord::MirrorIndents(_) => "mirror-indent",
            ControlWord::KeepTogether => "keep",
            ControlWord::KeepNext => "keep-next",
            ControlWord::SideBySide(_) => "side-by-side",
            ControlWord::PageBreakBefore => "page-break",
            ControlWord::WidowControl | ControlWord::NoWidowControl(_) => "widow",
            ControlWord::DropCapLines(_) => "drop-cap-lines",
            ControlWord::DropCapType(_) => "drop-cap-type",
            ControlWord::ParagraphHyphenation(_) => "hyphenation",
            ControlWord::AutoSpaceAlphabetic(_) => "auto-alpha",
            ControlWord::AutoSpaceNumbers(_) => "auto-number",
            ControlWord::AdjustRightIndent(_) => "adjust-right",
            ControlWord::WrapDefault(_)
            | ControlWord::NoCharacterWrap(_)
            | ControlWord::NoWordWrap(_)
            | ControlWord::NoOverflow(_) => "wrapping",
            ControlWord::FontAlignAuto(_)
            | ControlWord::FontAlignHanging(_)
            | ControlWord::FontAlignCenter(_)
            | ControlWord::FontAlignRoman(_)
            | ControlWord::FontAlignVariable(_)
            | ControlWord::FontAlignFixed(_) => "font-alignment",
            ControlWord::ListOverrideIndex(_) => "list-override",
            ControlWord::ListLevelIndex(_) => "list-level",
            ControlWord::Shading(_) => "shading",
            ControlWord::ForegroundPattern(_) => "shading-foreground",
            ControlWord::BackgroundPattern(_) => "shading-background",
            _ => return None,
        })
    }

    pub(super) fn parse_default_formatting_destination(
        &mut self,
        kind: crate::DefaultFormattingDestination,
    ) -> RtfResult<()> {
        if self.states.len() != 3
            || self.section_note_options_closed
            || self
                .states
                .get(1)
                .is_none_or(|state| state.destination != Destination::DocumentBody)
        {
            return Err(RtfError::MalformedDocument(
                "RTF default-formatting destination must occur once in the root document header"
                    .to_string(),
            ));
        }
        if match kind {
            crate::DefaultFormattingDestination::Character => {
                self.default_formatting.character().is_some()
            },
            crate::DefaultFormattingDestination::Paragraph => {
                self.default_formatting.paragraph().is_some()
            },
        } {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF default-formatting destination".to_string(),
            ));
        }
        self.pos += 1;
        let expected = match kind {
            crate::DefaultFormattingDestination::Character => {
                ControlWord::DefaultCharacterProperties(None)
            },
            crate::DefaultFormattingDestination::Paragraph => {
                ControlWord::DefaultParagraphProperties(None)
            },
        };
        match (self.tokens.get(self.pos), expected) {
            (
                Some(Token::Control(ControlWord::DefaultCharacterProperties(parameter))),
                ControlWord::DefaultCharacterProperties(_),
            )
            | (
                Some(Token::Control(ControlWord::DefaultParagraphProperties(parameter))),
                ControlWord::DefaultParagraphProperties(_),
            ) => require_parameterless(
                *parameter,
                match kind {
                    crate::DefaultFormattingDestination::Character => "defchp",
                    crate::DefaultFormattingDestination::Paragraph => "defpap",
                },
            )?,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF default-formatting destination".to_string(),
                ));
            },
        }
        self.pos += 1;
        let mut state = State {
            unicode_skip: self.current_state()?.unicode_skip,
            ..State::default()
        };
        let mut seen = std::collections::HashSet::new();
        let mut script = None;
        let mut low = None;
        let mut high = None;
        let mut double = None;
        let mut itap = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    if script.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF defchp script selector lacks its af value".to_string(),
                        ));
                    }
                    if state.pending_tab_alignment.is_some() || state.pending_tab_leader.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF defpap has an incomplete tab definition".to_string(),
                        ));
                    }
                    Self::validate_drop_cap_state(&state, "defpap")?;
                    self.pos += 1;
                    match kind {
                        crate::DefaultFormattingDestination::Character => self
                            .default_formatting
                            .set_character(crate::DefaultCharacterProperties {
                                formatting: state.formatting,
                                low_ansi_font: low,
                                high_ansi_font: high,
                                double_byte_font: double,
                            }),
                        crate::DefaultFormattingDestination::Paragraph => self
                            .default_formatting
                            .set_paragraph(crate::DefaultParagraphProperties {
                                paragraph: state.paragraph,
                                table_nesting_level: itap,
                            }),
                    }
                    self.default_formatting.validate()?;
                    return Ok(());
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF default-formatting destination cannot contain nested content"
                            .to_string(),
                    ));
                },
                Some(Token::Binary(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF default-formatting destination contains active content".to_string(),
                    ));
                },
                Some(Token::Control(control)) => match kind {
                    crate::DefaultFormattingDestination::Character => {
                        if matches!(
                            control,
                            ControlWord::LowAnsiCharacter(_)
                                | ControlWord::HighAnsiCharacter(_)
                                | ControlWord::DoubleByteCharacter(_)
                        ) {
                            if script.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF defchp script selector lacks its af value".to_string(),
                                ));
                            }
                            let (parameter, key, value) = match control {
                                ControlWord::LowAnsiCharacter(v) => (v, "loch", 0u8),
                                ControlWord::HighAnsiCharacter(v) => (v, "hich", 1),
                                ControlWord::DoubleByteCharacter(v) => (v, "dbch", 2),
                                _ => return Err(parser_classification_error()),
                            };
                            require_parameterless(*parameter, key)?;
                            if !seen.insert(key) {
                                return Err(RtfError::MalformedDocument(format!(
                                    "duplicate RTF {key} selector in defchp"
                                )));
                            }
                            script = Some(value);
                        } else if let ControlWord::AssociatedFontNumber(parameter) = control
                            && script.is_some()
                        {
                            let value = associated_font_ref(*parameter)?;
                            match script.take().ok_or_else(|| {
                                RtfError::ParserError(
                                    "RTF defchp script selector state was lost".to_string(),
                                )
                            })? {
                                0 => low = Some(value),
                                1 => high = Some(value),
                                _ => double = Some(value),
                            }
                        } else if Self::apply_character_decoration_control(&mut state, control)? {
                        } else if let Some(key) = Self::default_character_property_key(control) {
                            if !seen.insert(key) {
                                return Err(RtfError::MalformedDocument(format!(
                                    "duplicate RTF {key} property in defchp"
                                )));
                            }
                            match control {
                                ControlWord::FontNumber(value) => {
                                    u16::try_from(*value).map_err(|_| {
                                        RtfError::MalformedDocument(
                                            "RTF defchp font value must be in 0..=65535"
                                                .to_string(),
                                        )
                                    })?;
                                },
                                ControlWord::FontSize(value)
                                    if !(1..=i32::from(u16::MAX)).contains(value) =>
                                {
                                    return Err(RtfError::MalformedDocument(
                                        "RTF defchp fs value must be in 1..=65535".to_string(),
                                    ));
                                },
                                ControlWord::ColorForeground(value)
                                | ControlWord::Highlight(value) => {
                                    u16::try_from(*value).map_err(|_| {
                                        RtfError::MalformedDocument(
                                            "RTF defchp color value must be in 0..=65535"
                                                .to_string(),
                                        )
                                    })?;
                                },
                                ControlWord::ColorBackground(value) => {
                                    Self::required_character_value(*value, "cb", u16::MAX)?;
                                },
                                ControlWord::Language(value)
                                | ControlWord::LanguageEastAsian(value)
                                | ControlWord::LanguageNoProof(value)
                                | ControlWord::LanguageEastAsianNoProof(value) => {
                                    crate::LanguageId::from_rtf(*value)?;
                                },
                                _ => {},
                            }
                            Self::apply_style_property(&mut state, control)?
                        } else {
                            return Err(RtfError::MalformedDocument(
                                "unsupported control in RTF defchp destination".to_string(),
                            ));
                        }
                    },
                    crate::DefaultFormattingDestination::Paragraph => {
                        if Self::apply_paragraph_tab_control(&mut state, control)? {
                        } else if let ControlWord::TableNestingLevel(parameter) = control {
                            if !seen.insert("itap") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF itap property in defpap".to_string(),
                                ));
                            }
                            let value = parameter.ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF defpap itap requires a numeric parameter".to_string(),
                                )
                            })?;
                            let value = u8::try_from(value).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF defpap itap value must be in 0..=32".to_string(),
                                )
                            })?;
                            if value > 32 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF defpap itap value must be in 0..=32".to_string(),
                                ));
                            }
                            itap = Some(value);
                        } else if let Some(key) = Self::default_paragraph_property_key(control) {
                            if !seen.insert(key) {
                                return Err(RtfError::MalformedDocument(format!(
                                    "duplicate RTF {key} property in defpap"
                                )));
                            }
                            if matches!(control,ControlWord::LeftIndent(value)|ControlWord::RightIndent(value)|ControlWord::FirstLineIndent(value)|ControlWord::SpaceBefore(value)|ControlWord::SpaceAfter(value)|ControlWord::SpaceBetween(value) if value.unsigned_abs()>10_000_000)
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF defpap layout value exceeds the safety limit".to_string(),
                                ));
                            }
                            match control {
                                ControlWord::NoWidowControl(parameter) => {
                                    require_parameterless(*parameter, "nowidctlpar")?;
                                    state.paragraph.widow_control = false
                                },
                                _ => Self::apply_style_property(&mut state, control)?,
                            }
                        } else {
                            return Err(RtfError::MalformedDocument(
                                "unsupported control in RTF defpap destination".to_string(),
                            ));
                        }
                    },
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1
        }
    }

    pub(super) fn parse_stylesheet(&mut self) -> RtfResult<()> {
        if self.saw_stylesheet {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple stylesheet destinations".to_string(),
            ));
        }
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF stylesheet must occur in the root document header".to_string(),
            ));
        }
        self.saw_stylesheet = true;
        self.pos += 1; // `stylesheet`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => self.parse_style_entry()?,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.stylesheet.validate()?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_style_entry(&mut self) -> RtfResult<()> {
        self.pos += 1; // opening brace
        let starred = matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        );
        if starred {
            self.pos += 1;
        }
        let mut style_type = None;
        let mut id = None;
        let inherited_unicode_skip = self.current_state()?.unicode_skip;
        let mut state = State {
            unicode_skip: inherited_unicode_skip,
            ..State::default()
        };
        let mut name = String::new();
        let mut name_complete = false;
        let mut based_on = None;
        let mut next_style = None;
        let mut linked_style = None;
        let mut additive = false;
        let mut auto_update = false;
        let mut hidden = false;
        let mut locked = false;
        let mut semi_hidden = false;
        let mut unhide_when_used = false;
        let mut quick_format = false;
        let mut priority = None;
        let mut revision_id = None;
        let mut personal = false;
        let mut compose = false;
        let mut reply = false;
        let mut table_conditional = crate::TableStyleConditionalFormatting::default();
        let mut seen_metadata = std::collections::HashSet::new();
        let mut saw_content_before_selector = false;
        macro_rules! set_style_once {
            ($key:literal, $target:expr, $value:expr) => {{
                if !seen_metadata.insert($key) {
                    return Err(RtfError::MalformedDocument(format!(
                        "duplicate RTF style metadata control: {}",
                        $key
                    )));
                }
                $target = $value;
            }};
        }

        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Control(ControlWord::Page(_))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF page is not permitted in a stylesheet".to_string(),
                    ));
                },
                Some(Token::OpenBrace) => {
                    // Nested extension groups do not form part of the style name.
                    self.pos += 1;
                    self.preserve_unknown_destination_in(crate::opaque::Context::Metadata)?;
                    continue;
                },
                Some(Token::Text(text)) if !name_complete => {
                    let decoded = self.decode_transport_text(text)?;
                    if style_type.is_none() && name.is_empty() && decoded.trim().is_empty() {
                        self.pos += 1;
                        continue;
                    }
                    saw_content_before_selector = true;
                    Self::append_style_name(&mut name, &decoded, &mut name_complete);
                },
                Some(Token::Control(ControlWord::Unicode(first))) if !name_complete => {
                    saw_content_before_selector = true;
                    let decoded = self.parse_style_unicode(*first, state.unicode_skip)?;
                    Self::append_style_name(&mut name, &decoded, &mut name_complete);
                    if name.len() > MAX_STYLE_NAME_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF style name exceeds the safety limit".to_string(),
                        ));
                    }
                    continue;
                },
                Some(Token::Control(control)) => match control {
                    control if !name_complete && control_symbol_text(control).is_some() => {
                        Self::append_style_name(
                            &mut name,
                            control_symbol_text(control).unwrap_or_default(),
                            &mut name_complete,
                        );
                    },
                    ControlWord::ParagraphStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::super::super::stylesheet::StyleType::Paragraph);
                        id = Some(paragraph_style_reference(*value)?);
                    },
                    ControlWord::CharacterStyle(value) if style_type.is_none() => {
                        if saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::super::super::stylesheet::StyleType::Character);
                        id = Some(character_style_reference(*value)?);
                    },
                    ControlWord::CharacterStyle(value) => {
                        state.formatting.character_style = Some(character_style_reference(*value)?);
                    },
                    ControlWord::SectionStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::super::super::stylesheet::StyleType::Section);
                        id = Some(section_style_reference(*value)?);
                    },
                    ControlWord::TableStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::super::super::stylesheet::StyleType::Table);
                        id = Some(table_style_reference(*value)?);
                    },
                    control @ (ControlWord::TableStyleRowDefaults(_)
                    | ControlWord::TableStyleFirstRow(_)
                    | ControlWord::TableStyleLastRow(_)
                    | ControlWord::TableStyleFirstColumn(_)
                    | ControlWord::TableStyleLastColumn(_)
                    | ControlWord::TableStyleBandHorizontalOdd(_)
                    | ControlWord::TableStyleBandHorizontalEven(_)
                    | ControlWord::TableStyleBandVerticalOdd(_)
                    | ControlWord::TableStyleBandVerticalEven(_)
                    | ControlWord::TableStyleBandSizeHorizontal(_)
                    | ControlWord::TableStyleBandSizeVertical(_)) => {
                        if style_type != Some(super::super::super::stylesheet::StyleType::Table) {
                            return Err(RtfError::MalformedDocument(
                                "RTF table-style conditional controls may occur only in table style definitions"
                                    .to_string(),
                            ));
                        }
                        match control {
                            ControlWord::TableStyleRowDefaults(param) => {
                                require_parameterless(*param, "tsrowd")?;
                                set_style_once!(
                                    "tsrowd",
                                    table_conditional.row_defaults_marker,
                                    true
                                );
                            },
                            ControlWord::TableStyleFirstRow(param) => {
                                require_parameterless(*param, "tscfirstrow")?;
                                set_style_once!("tscfirstrow", table_conditional.first_row, true);
                            },
                            ControlWord::TableStyleLastRow(param) => {
                                require_parameterless(*param, "tsclastrow")?;
                                set_style_once!("tsclastrow", table_conditional.last_row, true);
                            },
                            ControlWord::TableStyleFirstColumn(param) => {
                                require_parameterless(*param, "tscfirstcol")?;
                                set_style_once!(
                                    "tscfirstcol",
                                    table_conditional.first_column,
                                    true
                                );
                            },
                            ControlWord::TableStyleLastColumn(param) => {
                                require_parameterless(*param, "tsclastcol")?;
                                set_style_once!("tsclastcol", table_conditional.last_column, true);
                            },
                            ControlWord::TableStyleBandHorizontalOdd(param) => {
                                require_parameterless(*param, "tscbandhorzodd")?;
                                set_style_once!(
                                    "tscbandhorzodd",
                                    table_conditional.band_horizontal_odd,
                                    true
                                );
                            },
                            ControlWord::TableStyleBandHorizontalEven(param) => {
                                require_parameterless(*param, "tscbandhorzeven")?;
                                set_style_once!(
                                    "tscbandhorzeven",
                                    table_conditional.band_horizontal_even,
                                    true
                                );
                            },
                            ControlWord::TableStyleBandVerticalOdd(param) => {
                                require_parameterless(*param, "tscbandvertodd")?;
                                set_style_once!(
                                    "tscbandvertodd",
                                    table_conditional.band_vertical_odd,
                                    true
                                );
                            },
                            ControlWord::TableStyleBandVerticalEven(param) => {
                                require_parameterless(*param, "tscbandverteven")?;
                                set_style_once!(
                                    "tscbandverteven",
                                    table_conditional.band_vertical_even,
                                    true
                                );
                            },
                            ControlWord::TableStyleBandSizeHorizontal(param) => {
                                let value =
                                    Self::required_character_value(*param, "tscbandsh", u16::MAX)?;
                                set_style_once!(
                                    "tscbandsh",
                                    table_conditional.horizontal_band_size,
                                    Some(value)
                                );
                            },
                            ControlWord::TableStyleBandSizeVertical(param) => {
                                let value =
                                    Self::required_character_value(*param, "tscbandsv", u16::MAX)?;
                                set_style_once!(
                                    "tscbandsv",
                                    table_conditional.vertical_band_size,
                                    Some(value)
                                );
                            },
                            _ => return Err(parser_classification_error()),
                        }
                    },
                    ControlWord::StyleBasedOn(value) => {
                        if !seen_metadata.insert("sbasedon") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF sbasedon".to_string(),
                            ));
                        }
                        based_on = Some(Self::style_id(*value, "based-on style")?);
                    },
                    ControlWord::StyleNext(value) => {
                        if !seen_metadata.insert("snext") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF snext".to_string(),
                            ));
                        }
                        next_style = Some(Self::style_id(*value, "next style")?);
                    },
                    ControlWord::StyleLink(value) => {
                        if !seen_metadata.insert("slink") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF slink".to_string(),
                            ));
                        }
                        linked_style = Some(Self::style_id(*value, "linked style")?);
                    },
                    ControlWord::StyleAdditive(value) => {
                        set_style_once!("additive", additive, *value)
                    },
                    ControlWord::StyleAutoUpdate(value) => {
                        set_style_once!("sautoupd", auto_update, *value)
                    },
                    ControlWord::StyleHidden(value) => set_style_once!("shidden", hidden, *value),
                    ControlWord::StyleLocked(value) => set_style_once!("slocked", locked, *value),
                    ControlWord::StyleSemiHidden(value) => {
                        set_style_once!("ssemihidden", semi_hidden, *value)
                    },
                    ControlWord::StyleUnhideWhenUsed(value) => {
                        set_style_once!("sunhideused", unhide_when_used, *value)
                    },
                    ControlWord::StyleQuickFormat(value) => {
                        set_style_once!("sqformat", quick_format, *value)
                    },
                    ControlWord::StylePriority(value) => {
                        set_style_once!("spriority", priority, Some(*value))
                    },
                    ControlWord::StyleRevisionId(value) => {
                        if !seen_metadata.insert("styrsid") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF styrsid".to_string(),
                            ));
                        }
                        revision_id = Some(*value);
                    },
                    ControlWord::StylePersonal(value) => {
                        set_style_once!("spersonal", personal, *value)
                    },
                    ControlWord::StyleCompose(value) => {
                        set_style_once!("scompose", compose, *value)
                    },
                    ControlWord::StyleReply(value) => set_style_once!("sreply", reply, *value),
                    ControlWord::UnicodeSkip(value) => state.unicode_skip = (*value).max(0),
                    _ => {
                        saw_content_before_selector = style_type.is_none();
                        Self::apply_style_property(&mut state, control)?;
                    },
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
            if name.len() > MAX_STYLE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF style name exceeds the safety limit".to_string(),
                ));
            }
        }

        if style_type.is_none() && !starred && name_complete {
            style_type = Some(super::super::super::stylesheet::StyleType::Paragraph);
            id = Some(0);
        }
        let (Some(style_type), Some(id)) = (style_type, id) else {
            // Unknown starred extension groups are permitted inside a stylesheet.
            return Ok(());
        };
        if !name_complete {
            return Err(RtfError::MalformedDocument(
                "RTF style name must end with a semicolon".to_string(),
            ));
        }
        if style_type != super::super::super::stylesheet::StyleType::Paragraph && !starred {
            return Err(RtfError::MalformedDocument(
                "RTF non-paragraph style entries must be starred".to_string(),
            ));
        }
        if self.stylesheet.styles().len() >= MAX_STYLES {
            return Err(RtfError::MalformedDocument(
                "RTF style count exceeds the safety limit".to_string(),
            ));
        }
        let name = name.trim().to_string();
        let allocated = self.arena.alloc_str(&name);
        let mut style = match style_type {
            super::super::super::stylesheet::StyleType::Paragraph => {
                super::super::super::stylesheet::Style::paragraph(id, Cow::Borrowed(allocated))
            },
            super::super::super::stylesheet::StyleType::Character => {
                super::super::super::stylesheet::Style::character(id, Cow::Borrowed(allocated))
            },
            super::super::super::stylesheet::StyleType::Section => {
                super::super::super::stylesheet::Style::section(id, Cow::Borrowed(allocated))
            },
            super::super::super::stylesheet::StyleType::Table => {
                super::super::super::stylesheet::Style::table(id, Cow::Borrowed(allocated))
            },
        };
        style.based_on = based_on;
        style.next_style = next_style;
        style.linked_style = linked_style;
        style.formatting = state.formatting;
        if style_type == super::super::super::stylesheet::StyleType::Paragraph {
            Self::validate_drop_cap_state(&state, "paragraph style")?;
            style.paragraph = Some(state.paragraph);
        }
        style.hidden = hidden;
        style.additive = additive;
        style.auto_update = auto_update;
        style.locked = locked;
        style.semi_hidden = semi_hidden;
        style.unhide_when_used = unhide_when_used;
        style.quick_format = quick_format;
        style.priority = priority;
        style.revision_id = revision_id;
        style.personal = personal;
        style.compose = compose;
        style.reply = reply;
        style.table_conditional = table_conditional;
        self.stylesheet.add(style);
        Ok(())
    }

    pub(super) fn style_id(value: i32, field: &str) -> RtfResult<u16> {
        u16::try_from(value).map_err(|_| {
            RtfError::MalformedDocument(format!("RTF {field} ID is outside the supported range"))
        })
    }

    pub(super) fn append_style_name(name: &mut String, text: &str, complete: &mut bool) {
        if let Some((prefix, _)) = text.split_once(';') {
            name.push_str(prefix);
            *complete = true;
        } else {
            name.push_str(text);
        }
    }

    pub(super) fn parse_unicode_with_remainder(
        &mut self,
        first_code: i32,
        unicode_skip: i32,
    ) -> RtfResult<(String, String)> {
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        utf16.push(first_code as u16);
        self.pos += 1;
        while let Some(Token::Control(ControlWord::Unicode(code))) = self.tokens.get(self.pos) {
            utf16.push(*code as u16);
            self.pos += 1;
        }

        let mut fallback_skip = (unicode_skip.max(0) as usize).saturating_mul(utf16.len());
        let mut remainder = String::new();
        while fallback_skip > 0 && self.pos < self.tokens.len() {
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
                Some(Token::Control(ControlWord::Unicode(_))) => break,
                Some(_) => {
                    fallback_skip -= 1;
                    self.pos += 1;
                },
                None => break,
            }
        }
        let decoded = String::from_utf16(&utf16)
            .map_err(|error| RtfError::InvalidUnicode(format!("invalid style name: {error}")))?;
        Ok((decoded, remainder))
    }

    pub(super) fn parse_style_unicode(
        &mut self,
        first_code: i32,
        unicode_skip: i32,
    ) -> RtfResult<String> {
        let (mut decoded, remainder) =
            self.parse_unicode_with_remainder(first_code, unicode_skip)?;
        decoded.push_str(&self.decode_transport_text(&remainder)?);
        Ok(decoded)
    }

    pub(super) fn append_deferred_unicode(
        &mut self,
        target: &mut DeferredText,
        first_code: i32,
        unicode_skip: i32,
    ) -> RtfResult<()> {
        let (decoded, remainder) = self.parse_unicode_with_remainder(first_code, unicode_skip)?;
        target.push_unicode(&decoded);
        target.push_transport(&remainder)
    }

    pub(super) fn required_character_value(
        value: Option<i32>,
        control: &str,
        maximum: u16,
    ) -> RtfResult<u16> {
        let value = value.ok_or_else(|| {
            RtfError::MalformedDocument(format!("RTF {control} requires a numeric parameter"))
        })?;
        let value = u16::try_from(value).map_err(|_| {
            RtfError::MalformedDocument(format!("RTF {control} value must be in 0..={maximum}"))
        })?;
        if value > maximum {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {control} value must be in 0..={maximum}"
            )));
        }
        Ok(value)
    }

    pub(super) fn table_border_style(control: &ControlWord<'_>) -> Option<crate::BorderStyle> {
        use crate::BorderStyle as Style;
        Some(match control {
            ControlWord::BorderNone => Style::None,
            ControlWord::BorderHairline => Style::Hairline,
            ControlWord::BorderSingle => Style::Single,
            ControlWord::BorderThick => Style::Thick,
            ControlWord::BorderDotted => Style::Dotted,
            ControlWord::BorderDashed => Style::Dashed,
            ControlWord::BorderDashSmall => Style::DashSmallGap,
            ControlWord::BorderDotDash => Style::DotDash,
            ControlWord::BorderDotDotDash => Style::DotDotDash,
            ControlWord::BorderDouble => Style::Double,
            ControlWord::BorderTriple => Style::Triple,
            ControlWord::BorderThinThickSmall => Style::ThinThickSmall,
            ControlWord::BorderThickThinSmall => Style::ThickThinSmall,
            ControlWord::BorderThinThickThinSmall => Style::ThinThickThinSmall,
            ControlWord::BorderThinThickMedium => Style::ThinThickMedium,
            ControlWord::BorderThickThinMedium => Style::ThickThinMedium,
            ControlWord::BorderThinThickThinMedium => Style::ThinThickThinMedium,
            ControlWord::BorderThinThickLarge => Style::ThinThickLarge,
            ControlWord::BorderThickThinLarge => Style::ThickThinLarge,
            ControlWord::BorderThinThickThinLarge => Style::ThinThickThinLarge,
            ControlWord::BorderWave => Style::Wavy,
            ControlWord::BorderWavyDouble => Style::WavyDouble,
            ControlWord::BorderStriped => Style::Striped,
            ControlWord::BorderEmbossed => Style::Embossed,
            ControlWord::BorderEngraved => Style::Engraved,
            ControlWord::BorderOutset => Style::Outset,
            ControlWord::BorderInset => Style::Inset,
            _ => return None,
        })
    }

    pub(super) fn apply_table_decoration_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        const STYLE: u8 = 1;
        const WIDTH: u8 = 2;
        const COLOR: u8 = 4;
        const SPACE: u8 = 8;
        const SHADOW: u8 = 16;
        const FRAME: u8 = 32;
        if let ControlWord::TableBorder(target, param) = control {
            require_parameterless(*param, "table border selector")?;
            state.active_table_border = Some(*target);
            state.active_table_border_seen = 0;
            let slot = match target {
                crate::table::TableBorderTarget::Row(side) => {
                    state.table_row_borders.side_mut(*side)
                },
                crate::table::TableBorderTarget::Cell(side) => {
                    state.pending_cell_borders.side_mut(*side)
                },
                crate::table::TableBorderTarget::StyleDefault(side) => {
                    state.table_default_borders.side_mut(*side)
                },
            };
            *slot = Some(crate::Border::default());
            return Ok(true);
        }
        let shading_control = matches!(
            control,
            ControlWord::TableShadingAmount(..)
                | ControlWord::TableShadingRawAmount(..)
                | ControlWord::TableShadingRawNil(..)
                | ControlWord::TableShadingForeground(..)
                | ControlWord::TableShadingBackground(..)
                | ControlWord::TableShadingPattern(..)
                | ControlWord::TableRowShadingPatternIndex(..)
        );
        if shading_control {
            state.active_table_border = None;
            state.active_table_border_seen = 0;
            let scope = match control {
                ControlWord::TableShadingAmount(scope, _)
                | ControlWord::TableShadingRawAmount(scope, _)
                | ControlWord::TableShadingRawNil(scope, _)
                | ControlWord::TableShadingForeground(scope, _)
                | ControlWord::TableShadingBackground(scope, _)
                | ControlWord::TableShadingPattern(scope, _, _) => *scope,
                ControlWord::TableRowShadingPatternIndex(_) => crate::TableDistanceScope::Row,
                _ => return Err(parser_classification_error()),
            };
            let (shading, seen) = match scope {
                crate::TableDistanceScope::Row => (
                    &mut state.table_row_shading,
                    &mut state.table_row_shading_seen,
                ),
                crate::TableDistanceScope::Cell => (
                    &mut state.pending_cell_shading,
                    &mut state.pending_cell_shading_seen,
                ),
            };
            let bit = match control {
                ControlWord::TableShadingAmount(..) => 1,
                ControlWord::TableShadingRawAmount(..) => 16,
                ControlWord::TableShadingRawNil(..) => 32,
                ControlWord::TableShadingForeground(..) => 2,
                ControlWord::TableShadingBackground(..) => 4,
                ControlWord::TableShadingPattern(..)
                | ControlWord::TableRowShadingPatternIndex(..) => 8,
                _ => return Err(parser_classification_error()),
            };
            if *seen & bit != 0 {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF table shading component".to_string(),
                ));
            }
            *seen |= bit;
            match control {
                ControlWord::TableShadingAmount(_, value) => {
                    shading.amount = Some(required_table_value(*value, "table shading", 10_000)?)
                },
                ControlWord::TableShadingRawAmount(_, value) => {
                    shading.raw_amount =
                        Some(required_table_value(*value, "raw table shading", 10_000)?)
                },
                ControlWord::TableShadingRawNil(_, value) => {
                    require_parameterless(*value, "clshdrawnil")?;
                    shading.raw_nil = true;
                },
                ControlWord::TableShadingForeground(_, value) => {
                    shading.foreground_color = Some(required_table_value(
                        *value,
                        "table shading foreground color",
                        u16::MAX,
                    )?)
                },
                ControlWord::TableShadingBackground(_, value) => {
                    shading.background_color = Some(required_table_value(
                        *value,
                        "table shading background color",
                        u16::MAX,
                    )?)
                },
                ControlWord::TableShadingPattern(_, pattern, param) => {
                    require_parameterless(*param, "table shading pattern")?;
                    shading.pattern = Some(*pattern)
                },
                ControlWord::TableRowShadingPatternIndex(value) => {
                    shading.pattern_index = Some(required_table_value(*value, "trpat", u16::MAX)?)
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(true);
        }
        let Some(target) = state.active_table_border else {
            return Ok(false);
        };
        let (component, name) = if Self::table_border_style(control).is_some() {
            (STYLE, "style")
        } else {
            match control {
                ControlWord::BorderWidth(_) => (WIDTH, "width"),
                ControlWord::BorderColor(_) => (COLOR, "color"),
                ControlWord::BorderSpace(_) => (SPACE, "spacing"),
                ControlWord::BorderShadow => (SHADOW, "shadow"),
                ControlWord::BorderFrame => (FRAME, "frame"),
                _ => {
                    state.active_table_border = None;
                    state.active_table_border_seen = 0;
                    return Ok(false);
                },
            }
        };
        if state.active_table_border_seen & component != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF table-border {name}"
            )));
        }
        if component != STYLE && state.active_table_border_seen & STYLE == 0 {
            return Err(RtfError::MalformedDocument(format!(
                "RTF table-border {name} precedes its style"
            )));
        }
        state.active_table_border_seen |= component;
        let border = match target {
            crate::table::TableBorderTarget::Row(side) => state.table_row_borders.side_mut(side),
            crate::table::TableBorderTarget::Cell(side) => {
                state.pending_cell_borders.side_mut(side)
            },
            crate::table::TableBorderTarget::StyleDefault(side) => {
                state.table_default_borders.side_mut(side)
            },
        }
        .as_mut()
        .ok_or_else(|| {
            RtfError::ParserError("RTF active table-border state is missing".to_string())
        })?;
        if let Some(style) = Self::table_border_style(control) {
            border.style = style
        } else {
            match control {
                ControlWord::BorderWidth(value) => {
                    border.width = i32::from(required_table_value(*value, "brdrw", 75)?)
                },
                ControlWord::BorderColor(value) => {
                    border.color_ref = required_table_value(*value, "brdrcf", u16::MAX)?
                },
                ControlWord::BorderSpace(value) => {
                    border.space = i32::from(required_table_value(
                        *value,
                        "brsp",
                        crate::MAX_TABLE_DISTANCE_TWIPS as u16,
                    )?)
                },
                ControlWord::BorderShadow => border.shadow = true,
                ControlWord::BorderFrame => border.frame = true,
                _ => return Err(parser_classification_error()),
            }
        }
        Ok(true)
    }

    pub(super) fn character_border_style(
        control: &ControlWord<'_>,
    ) -> Option<crate::CharacterBorderStyle> {
        use crate::CharacterBorderStyle as Style;
        Some(match control {
            ControlWord::BorderNone => Style::None,
            ControlWord::BorderHairline => Style::Hairline,
            ControlWord::BorderSingle => Style::Single,
            ControlWord::BorderThick => Style::Thick,
            ControlWord::BorderDotted => Style::Dotted,
            ControlWord::BorderDashed => Style::Dashed,
            ControlWord::BorderDashSmall => Style::DashSmallGap,
            ControlWord::BorderDotDash => Style::DotDash,
            ControlWord::BorderDotDotDash => Style::DotDotDash,
            ControlWord::BorderDouble => Style::Double,
            ControlWord::BorderTriple => Style::Triple,
            ControlWord::BorderThinThickSmall => Style::ThinThickSmallGap,
            ControlWord::BorderThickThinSmall => Style::ThickThinSmallGap,
            ControlWord::BorderThinThickThinSmall => Style::ThinThickThinSmallGap,
            ControlWord::BorderThinThickMedium => Style::ThinThickMediumGap,
            ControlWord::BorderThickThinMedium => Style::ThickThinMediumGap,
            ControlWord::BorderThinThickThinMedium => Style::ThinThickThinMediumGap,
            ControlWord::BorderThinThickLarge => Style::ThinThickLargeGap,
            ControlWord::BorderThickThinLarge => Style::ThickThinLargeGap,
            ControlWord::BorderThinThickThinLarge => Style::ThinThickThinLargeGap,
            ControlWord::BorderWave => Style::Wavy,
            ControlWord::BorderWavyDouble => Style::DoubleWavy,
            ControlWord::BorderStriped => Style::Striped,
            ControlWord::BorderEmbossed => Style::Embossed,
            ControlWord::BorderEngraved => Style::Engraved,
            ControlWord::BorderOutset => Style::Outset,
            ControlWord::BorderInset => Style::Inset,
            _ => return None,
        })
    }

    pub(super) fn apply_character_decoration_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        const STYLE: u8 = 1;
        const WIDTH: u8 = 2;
        const COLOR: u8 = 4;
        const SPACE: u8 = 8;
        const SHADOW: u8 = 16;
        const FRAME: u8 = 32;

        match control {
            ControlWord::CharacterBorder(parameter) => {
                if parameter.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF chbrdr must not have a numeric parameter".to_string(),
                    ));
                }
                state.formatting.character_border = Some(crate::CharacterBorder::default());
                state.character_border_active = true;
                state.character_border_seen = 0;
                return Ok(true);
            },
            ControlWord::CharacterShading(value) => {
                state.character_border_active = false;
                let amount = Self::required_character_value(*value, "chshdng", 10_000)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .amount = amount;
                return Ok(true);
            },
            ControlWord::CharacterForegroundPattern(value) => {
                state.character_border_active = false;
                let color = Self::required_character_value(*value, "chcfpat", u16::MAX)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .foreground_color = color;
                return Ok(true);
            },
            ControlWord::CharacterBackgroundPattern(value) => {
                state.character_border_active = false;
                let color = Self::required_character_value(*value, "chcbpat", u16::MAX)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .background_color = color;
                return Ok(true);
            },
            _ => {},
        }

        if !state.character_border_active {
            return Ok(false);
        }
        let (component, duplicate_name) = if Self::character_border_style(control).is_some() {
            (STYLE, "style")
        } else {
            match control {
                ControlWord::BorderWidth(_) => (WIDTH, "width"),
                ControlWord::BorderColor(_) => (COLOR, "color"),
                ControlWord::BorderSpace(_) => (SPACE, "space"),
                ControlWord::BorderShadow => (SHADOW, "shadow"),
                ControlWord::BorderFrame => (FRAME, "frame"),
                _ => {
                    state.character_border_active = false;
                    return Ok(false);
                },
            }
        };
        if state.character_border_seen & component != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF character-border {duplicate_name}"
            )));
        }
        state.character_border_seen |= component;
        let border = state.formatting.character_border.as_mut().ok_or_else(|| {
            RtfError::ParserError("RTF active character-border state is missing".to_string())
        })?;
        if let Some(style) = Self::character_border_style(control) {
            border.style = style;
        } else {
            match control {
                ControlWord::BorderWidth(value) => {
                    border.width = Self::required_character_value(*value, "brdrw", 75)?;
                },
                ControlWord::BorderColor(value) => {
                    border.color_ref = Self::required_character_value(*value, "brdrcf", u16::MAX)?;
                },
                ControlWord::BorderSpace(value) => {
                    border.space = Self::required_character_value(*value, "brsp", u16::MAX)?;
                },
                ControlWord::BorderShadow => border.shadow = true,
                ControlWord::BorderFrame => border.frame = true,
                _ => return Err(parser_classification_error()),
            }
        }
        Ok(true)
    }

    pub(super) fn apply_paragraph_border_side(
        state: &mut State,
        side: ParagraphBorderSide,
        apply: impl Fn(&mut crate::Border) -> RtfResult<()>,
    ) -> RtfResult<()> {
        let borders = &mut state.paragraph.borders;
        match side {
            ParagraphBorderSide::Top => apply(&mut borders.top),
            ParagraphBorderSide::Bottom => apply(&mut borders.bottom),
            ParagraphBorderSide::Left => apply(&mut borders.left),
            ParagraphBorderSide::Right => apply(&mut borders.right),
            ParagraphBorderSide::Bar => apply(&mut borders.bar),
            ParagraphBorderSide::Between => apply(&mut borders.between),
            ParagraphBorderSide::Box => {
                apply(&mut borders.top)?;
                apply(&mut borders.bottom)?;
                apply(&mut borders.left)?;
                apply(&mut borders.right)
            },
        }
    }

    /// Apply a paragraph border segment or component control.
    ///
    /// Segment controls (`\brdrt`, `\brdrb`, `\brdrl`, `\brdrr`, `\brdrbar`,
    /// `\brdrbtw`, `\box`) select the segment that subsequent style, width,
    /// color, spacing, shadow, and frame controls apply to (RTF 1.9.1
    /// paragraph borders). `\box` is normalized onto all four sides; the
    /// `\brdrbar` and `\brdrbtw` segments are retained separately.
    pub(super) fn apply_paragraph_border_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        let segment = match control {
            ControlWord::BorderTop => Some(ParagraphBorderSide::Top),
            ControlWord::BorderBottom => Some(ParagraphBorderSide::Bottom),
            ControlWord::BorderLeft => Some(ParagraphBorderSide::Left),
            ControlWord::BorderRight => Some(ParagraphBorderSide::Right),
            ControlWord::BorderBar => Some(ParagraphBorderSide::Bar),
            ControlWord::BorderBetween => Some(ParagraphBorderSide::Between),
            ControlWord::BorderBox => Some(ParagraphBorderSide::Box),
            _ => None,
        };
        if let Some(segment) = segment {
            let bit = segment.bit();
            if state.paragraph_border_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF paragraph border segment".to_string(),
                ));
            }
            state.paragraph_border_seen |= bit;
            state.paragraph_border_side = Some(segment);
            return Ok(true);
        }

        let side = state.paragraph_border_side;
        if let Some(style) = Self::table_border_style(control) {
            let side = side.ok_or_else(|| {
                RtfError::MalformedDocument("RTF paragraph border style has no segment".to_string())
            })?;
            Self::apply_paragraph_border_side(state, side, |border| {
                if border.style != crate::BorderStyle::None && style != crate::BorderStyle::None {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF paragraph border style".to_string(),
                    ));
                }
                border.style = style;
                Ok(())
            })?;
            return Ok(true);
        }
        let require_side = |state: &State| {
            state.paragraph_border_side.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF paragraph border component has no segment".to_string(),
                )
            })
        };
        match control {
            ControlWord::BorderWidth(value) => {
                let width = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF brdrw requires a numeric parameter".to_string(),
                    )
                })?;
                let side = require_side(state)?;
                Self::apply_paragraph_border_side(state, side, |border| {
                    border.width = width;
                    Ok(())
                })?;
                Ok(true)
            },
            ControlWord::BorderColor(value) => {
                let color = Self::required_character_value(*value, "brdrcf", u16::MAX)?;
                let side = require_side(state)?;
                Self::apply_paragraph_border_side(state, side, |border| {
                    border.color_ref = color;
                    Ok(())
                })?;
                Ok(true)
            },
            ControlWord::BorderSpace(value) => {
                let space = Self::required_character_value(*value, "brsp", u16::MAX)?;
                let side = require_side(state)?;
                Self::apply_paragraph_border_side(state, side, |border| {
                    border.space = i32::from(space);
                    Ok(())
                })?;
                Ok(true)
            },
            ControlWord::BorderShadow => {
                let side = require_side(state)?;
                Self::apply_paragraph_border_side(state, side, |border| {
                    border.shadow = true;
                    Ok(())
                })?;
                Ok(true)
            },
            ControlWord::BorderFrame => {
                let side = require_side(state)?;
                Self::apply_paragraph_border_side(state, side, |border| {
                    border.frame = true;
                    Ok(())
                })?;
                Ok(true)
            },
            _ => Ok(false),
        }
    }

    pub(super) fn apply_paragraph_shading_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        match control {
            ControlWord::Shading(value) => {
                state
                    .paragraph
                    .shading
                    .set_amount(Some(Self::required_character_value(
                        *value, "shading", 10_000,
                    )?))?
            },
            ControlWord::ForegroundPattern(value) => {
                state
                    .paragraph
                    .shading
                    .set_foreground_color(Some(Self::required_character_value(
                        *value,
                        "cfpat",
                        u16::MAX,
                    )?))
            },
            ControlWord::BackgroundPattern(value) => {
                state
                    .paragraph
                    .shading
                    .set_background_color(Some(Self::required_character_value(
                        *value,
                        "cbpat",
                        u16::MAX,
                    )?))
            },
            _ => return Ok(false),
        }
        Ok(true)
    }

    pub(super) fn apply_style_property(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<()> {
        if Self::apply_paragraph_tab_control(state, control)? {
            return Ok(());
        }
        if Self::apply_table_decoration_control(state, control)? {
            return Ok(());
        }
        if Self::apply_character_decoration_control(state, control)? {
            return Ok(());
        }
        if Self::apply_paragraph_border_control(state, control)? {
            return Ok(());
        }
        if Self::apply_paragraph_shading_control(state, control)? {
            return Ok(());
        }
        if Self::apply_drop_cap_control(state, control)? {
            return Ok(());
        }
        if apply_associated_character_control(&mut state.formatting.associated, control)? {
            return Ok(());
        }
        match control {
            ControlWord::CharacterStyle(value) => {
                state.formatting.character_style = Some(character_style_reference(*value)?);
            },
            ControlWord::FontNumber(value) => state.formatting.font_ref = *value as FontRef,
            ControlWord::FontSize(value) => {
                if let Some(size) = NonZeroU16::new((*value).clamp(1, i32::from(u16::MAX)) as u16) {
                    state.formatting.font_size = size;
                }
            },
            ControlWord::ColorForeground(value) => {
                state.formatting.color_ref = *value as ColorRef;
            },
            ControlWord::ColorBackground(value) => {
                state.formatting.background_color =
                    Some(Self::required_character_value(*value, "cb", u16::MAX)?);
            },
            ControlWord::Highlight(value) => {
                state.formatting.highlight_color = Some(*value as ColorRef);
            },
            ControlWord::Bold(value) => state.formatting.bold = *value,
            ControlWord::Italic(value) => state.formatting.italic = *value,
            ControlWord::Underline(value) => {
                state.formatting.underline = if *value {
                    UnderlineStyle::Single
                } else {
                    UnderlineStyle::None
                };
            },
            ControlWord::UnderlineNone => state.formatting.underline = UnderlineStyle::None,
            ControlWord::UnderlineDouble => state.formatting.underline = UnderlineStyle::Double,
            ControlWord::UnderlineDotted => state.formatting.underline = UnderlineStyle::Dotted,
            ControlWord::UnderlineDashed => state.formatting.underline = UnderlineStyle::Dashed,
            ControlWord::UnderlineDashDot => state.formatting.underline = UnderlineStyle::DashDot,
            ControlWord::UnderlineDashDotDot => {
                state.formatting.underline = UnderlineStyle::DashDotDot;
            },
            ControlWord::UnderlineWords => state.formatting.underline = UnderlineStyle::Words,
            ControlWord::UnderlineThick => state.formatting.underline = UnderlineStyle::Thick,
            ControlWord::UnderlineWave => state.formatting.underline = UnderlineStyle::Wave,
            ControlWord::UnderlineHairline => state.formatting.underline = UnderlineStyle::Hairline,
            ControlWord::UnderlineThickDotted => {
                state.formatting.underline = UnderlineStyle::ThickDotted
            },
            ControlWord::UnderlineThickDashed => {
                state.formatting.underline = UnderlineStyle::ThickDashed
            },
            ControlWord::UnderlineThickDashDot => {
                state.formatting.underline = UnderlineStyle::ThickDashDot
            },
            ControlWord::UnderlineThickDashDotDot => {
                state.formatting.underline = UnderlineStyle::ThickDashDotDot
            },
            ControlWord::UnderlineThickLongDash => {
                state.formatting.underline = UnderlineStyle::ThickLongDash
            },
            ControlWord::UnderlineLongDash => state.formatting.underline = UnderlineStyle::LongDash,
            ControlWord::UnderlineHeavyWave => {
                state.formatting.underline = UnderlineStyle::HeavyWave
            },
            ControlWord::UnderlineDoubleWave => {
                state.formatting.underline = UnderlineStyle::DoubleWave
            },
            ControlWord::UnderlineColor(value) => {
                state.formatting.underline_color = Some(u16::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF underline color is outside the supported range".to_string(),
                    )
                })?);
            },
            ControlWord::Strike(value) => state.formatting.strike = *value,
            ControlWord::DoubleStrike(value) => state.formatting.double_strike = *value,
            ControlWord::Superscript(value) => {
                state.formatting.superscript = *value;
                if *value {
                    state.formatting.subscript = false;
                }
                state
                    .formatting
                    .character_positioning
                    .set_superscript(*value);
            },
            ControlWord::Subscript(value) => {
                state.formatting.subscript = *value;
                if *value {
                    state.formatting.superscript = false;
                }
                state.formatting.character_positioning.set_subscript(*value);
            },
            ControlWord::NoSuperSub => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.clear_baseline();
            },
            ControlWord::BaselineUp(value) => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.set_raised(*value)?;
            },
            ControlWord::BaselineDown(value) => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.set_lowered(*value)?;
            },
            ControlWord::SmallCaps(value) => state.formatting.smallcaps = *value,
            ControlWord::AllCaps(value) => state.formatting.all_caps = *value,
            ControlWord::Hidden(value) => state.formatting.hidden = *value,
            ControlWord::Outline(value) => state.formatting.outline = *value,
            ControlWord::Shadow(value) => state.formatting.shadow = *value,
            ControlWord::Emboss(value) => state.formatting.emboss = *value,
            ControlWord::Imprint(value) => state.formatting.imprint = *value,
            ControlWord::CharSpacing(value) => {
                state
                    .formatting
                    .character_positioning
                    .set_quarter_point_expansion(*value)?;
                state.formatting.char_spacing = *value;
            },
            ControlWord::CharSpacingTwips(value) => {
                state
                    .formatting
                    .character_positioning
                    .set_twip_expansion(*value)?;
                state.formatting.char_spacing = *value;
            },
            ControlWord::CharScale(value) => {
                state.formatting.character_positioning.set_scale(*value)?;
                state.formatting.char_scale = *value;
            },
            ControlWord::Kerning(value) => {
                state.formatting.character_positioning.set_kerning(*value)?;
                state.formatting.kerning = *value;
            },
            ControlWord::Language(value) => {
                state.formatting.language = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageEastAsian(value) => {
                state.formatting.east_asian_language = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageNoProof(value) => {
                state.formatting.language_no_proof = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageEastAsianNoProof(value) => {
                state.formatting.east_asian_language_no_proof =
                    crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::NoProof(value) => state.formatting.no_proof = *value,
            ControlWord::LeftToRightCharacter => {
                state.formatting.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftCharacter => {
                state.formatting.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::LowAnsiCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "loch",
                    crate::CharacterType::LowAnsi,
                )?);
            },
            ControlWord::HighAnsiCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "hich",
                    crate::CharacterType::HighAnsi,
                )?);
            },
            ControlWord::DoubleByteCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "dbch",
                    crate::CharacterType::DoubleByte,
                )?);
            },
            ControlWord::FontComplexScript(value) => {
                state.formatting.complex_script = Some(complex_script_selector(*value)?);
            },
            ControlWord::CharacterGrid(value) => {
                state.formatting.character_grid = Some(character_grid(*value)?);
            },
            ControlWord::AnimatedText(value) => {
                state.formatting.animated_text = animated_text(*value)?;
            },
            ControlWord::FitText(value) => {
                state.formatting.fit_text = fit_text(*value)?;
            },
            ControlWord::EmphasisMark(mark, value) => {
                state.formatting.emphasis_mark = emphasis_mark(*mark, *value)?;
            },
            ControlWord::Plain => {
                state.formatting = Formatting::default();
                state.character_border_active = false;
                state.character_border_seen = 0;
            },
            ControlWord::ParagraphStyle(value) => {
                state.paragraph.paragraph_style = Some(paragraph_style_reference(*value)?);
            },
            ControlWord::ParagraphRsid(value) => {
                state.paragraph.paragraph_rsid = Some(*value as u32);
            },
            ControlWord::ParagraphRevisionAuthor(value) => {
                state.paragraph.revision.author = Some(nonnegative_author_index(*value, "prauth")?);
            },
            ControlWord::ParagraphRevisionDate(value) => {
                state.paragraph.revision.date = Some(*value);
            },
            ControlWord::OutlineLevel(value) => {
                let level = u8::try_from(*value)
                    .ok()
                    .filter(|level| *level <= 9)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF outline level must be between 0 and 9".to_string(),
                        )
                    })?;
                state.paragraph.outline_level = Some(level);
            },
            ControlWord::LeftAlign => state.paragraph.alignment = Alignment::Left,
            ControlWord::RightAlign => state.paragraph.alignment = Alignment::Right,
            ControlWord::Center => state.paragraph.alignment = Alignment::Center,
            ControlWord::Justify => state.paragraph.alignment = Alignment::Justify,
            ControlWord::LeftToRightParagraph => {
                state.paragraph.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftParagraph => {
                state.paragraph.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::Pard => {
                state.paragraph = Paragraph::default();
                state.paragraph_border_side = None;
                state.paragraph_border_seen = 0;
                state.drop_cap_kind = None;
                state.drop_cap_lines = None;
                state.pending_tab_alignment = None;
                state.pending_tab_leader = None;
            },
            ControlWord::SpaceBefore(value) => state.paragraph.spacing.before = *value,
            ControlWord::SpaceAfter(value) => state.paragraph.spacing.after = *value,
            ControlWord::SpaceBetween(value) => state.paragraph.spacing.line = *value,
            ControlWord::LineMultiple(value) => state.paragraph.spacing.line_multiple = *value,
            ControlWord::SpaceBeforeAuto(value) => {
                state.paragraph.spacing_policy.automatic_before =
                    required_paragraph_bool(*value, "sbauto")?
            },
            ControlWord::SpaceAfterAuto(value) => {
                state.paragraph.spacing_policy.automatic_after =
                    required_paragraph_bool(*value, "saauto")?
            },
            ControlWord::ListSpaceBefore(value) => {
                state.paragraph.spacing_policy.list_before =
                    Some(required_list_spacing(*value, "lisb")?)
            },
            ControlWord::ListSpaceAfter(value) => {
                state.paragraph.spacing_policy.list_after =
                    Some(required_list_spacing(*value, "lisa")?)
            },
            ControlWord::NoSnapLineGrid(value) => {
                strict_paragraph_selector(*value, "nosnaplinegrid")?;
                state.paragraph.spacing_policy.snap_to_line_grid = false;
            },
            ControlWord::ContextualSpacing(value) => {
                strict_paragraph_selector(*value, "contextualspace")?;
                state.paragraph.spacing_policy.contextual_spacing = true;
            },
            ControlWord::LeftIndent(value) => state.paragraph.indentation.left = *value,
            ControlWord::RightIndent(value) => state.paragraph.indentation.right = *value,
            ControlWord::FirstLineIndent(value) => {
                state.paragraph.indentation.first_line = *value;
            },
            ControlWord::LogicalLeftIndent(v) => {
                state.paragraph.logical_indentation.start =
                    Some(required_paragraph_indent(*v, "lin")?)
            },
            ControlWord::LogicalRightIndent(v) => {
                state.paragraph.logical_indentation.end =
                    Some(required_paragraph_indent(*v, "rin")?)
            },
            ControlWord::CharacterFirstLineIndent(v) => {
                state
                    .paragraph
                    .logical_indentation
                    .first_line_character_units = Some(required_paragraph_indent(*v, "cufi")?)
            },
            ControlWord::CharacterLeftIndent(v) => {
                state.paragraph.logical_indentation.left_character_units =
                    Some(required_paragraph_indent(*v, "culi")?)
            },
            ControlWord::CharacterRightIndent(v) => {
                state.paragraph.logical_indentation.right_character_units =
                    Some(required_paragraph_indent(*v, "curi")?)
            },
            ControlWord::MirrorIndents(v) => {
                strict_paragraph_selector(*v, "indmirror")?;
                state.paragraph.logical_indentation.mirrored = true;
            },
            ControlWord::KeepTogether => state.paragraph.keep_together = true,
            ControlWord::KeepNext => state.paragraph.keep_next = true,
            ControlWord::SideBySide(value) => state.paragraph.side_by_side = *value,
            ControlWord::PageBreakBefore => state.paragraph.page_break_before = true,
            ControlWord::WidowControl => state.paragraph.widow_control = true,
            ControlWord::ParagraphNoLineNumbering(param) => {
                require_parameterless(*param, "noline")?;
                state.paragraph.no_line_numbering = true;
            },
            ControlWord::ParagraphNoAutoTabIndent(param) => {
                require_parameterless(*param, "notabind")?;
                state.paragraph.no_auto_tab_indent = true;
            },
            ControlWord::ParagraphHyphenation(value) => {
                state.paragraph.line_breaking.automatic_hyphenation =
                    strict_paragraph_toggle(*value, "hyphpar")?
            },
            ControlWord::AutoSpaceAlphabetic(value) => {
                state.paragraph.line_breaking.auto_space_alphabetic =
                    strict_paragraph_toggle(*value, "aspalpha")?
            },
            ControlWord::AutoSpaceNumbers(value) => {
                state.paragraph.line_breaking.auto_space_numbers =
                    strict_paragraph_toggle(*value, "aspnum")?
            },
            ControlWord::AdjustRightIndent(value) => {
                state.paragraph.line_breaking.adjust_right_indent =
                    strict_paragraph_toggle(*value, "adjustright")?
            },
            ControlWord::WrapDefault(value) => {
                strict_paragraph_selector(*value, "wrapdefault")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::Default;
            },
            ControlWord::NoCharacterWrap(value) => {
                strict_paragraph_selector(*value, "nocwrap")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoCharacterWrap;
            },
            ControlWord::NoWordWrap(value) => {
                strict_paragraph_selector(*value, "nowwrap")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoWordWrap;
            },
            ControlWord::NoOverflow(value) => {
                strict_paragraph_selector(*value, "nooverflow")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoOverflow;
            },
            ControlWord::FontAlignAuto(value) => {
                strict_paragraph_selector(*value, "faauto")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Auto;
            },
            ControlWord::FontAlignHanging(value) => {
                strict_paragraph_selector(*value, "fahang")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Hanging;
            },
            ControlWord::FontAlignCenter(value) => {
                strict_paragraph_selector(*value, "facenter")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Center;
            },
            ControlWord::FontAlignRoman(value) => {
                strict_paragraph_selector(*value, "faroman")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Roman;
            },
            ControlWord::FontAlignVariable(value) => {
                strict_paragraph_selector(*value, "favar")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Variable;
            },
            ControlWord::FontAlignFixed(value) => {
                strict_paragraph_selector(*value, "fafixed")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Fixed;
            },
            ControlWord::ListOverrideIndex(value) => {
                state.paragraph.list_override = Some(*value);
            },
            ControlWord::ListLevelIndex(value) => {
                if let Ok(level @ 0..=8) = u8::try_from(*value) {
                    state.paragraph.list_level = Some(level);
                }
            },
            _ => {},
        }
        Ok(())
    }

    pub(super) fn apply_drop_cap_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        match control {
            ControlWord::DropCapLines(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF dropcapli requires a numeric parameter".to_string(),
                    )
                })?;
                let lines = u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument(format!(
                        "RTF dropcapli must be in 1..={}",
                        crate::MAX_PARAGRAPH_DROP_CAP_LINES
                    ))
                })?;
                if !(1..=crate::MAX_PARAGRAPH_DROP_CAP_LINES).contains(&lines) {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF dropcapli must be in 1..={}",
                        crate::MAX_PARAGRAPH_DROP_CAP_LINES
                    )));
                }
                state.drop_cap_lines = Some(lines as u8);
            },
            ControlWord::DropCapType(value) => {
                state.drop_cap_kind = Some(match value {
                    Some(1) => crate::ParagraphDropCapKind::InText,
                    Some(2) => crate::ParagraphDropCapKind::Margin,
                    Some(_) => {
                        return Err(RtfError::MalformedDocument(
                            "RTF dropcapt accepts only 1 or 2".to_string(),
                        ));
                    },
                    None => {
                        return Err(RtfError::MalformedDocument(
                            "RTF dropcapt requires a numeric parameter".to_string(),
                        ));
                    },
                });
            },
            _ => return Ok(false),
        }
        if let (Some(kind), Some(lines)) = (state.drop_cap_kind, state.drop_cap_lines) {
            state.paragraph.drop_cap = Some(crate::ParagraphDropCap::new(kind, u16::from(lines))?);
        }
        Ok(true)
    }

    pub(super) fn validate_drop_cap_state(state: &State, context: &str) -> RtfResult<()> {
        if state.drop_cap_kind.is_some() != state.drop_cap_lines.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {context} has incomplete drop-cap properties"
            )));
        }
        if let Some(drop_cap) = state.paragraph.drop_cap {
            drop_cap.validate()?;
        }
        Ok(())
    }

    pub(super) fn apply_paragraph_tab_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        use super::super::super::border::{TabAlignment, TabLeader, TabStop};

        fn require_flag(parameter: Option<i32>, name: &str) -> RtfResult<()> {
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} tab selector does not accept a numeric parameter"
                )));
            }
            Ok(())
        }

        fn select_alignment(
            state: &mut State,
            parameter: Option<i32>,
            name: &str,
            alignment: TabAlignment,
        ) -> RtfResult<()> {
            require_flag(parameter, name)?;
            if state.pending_tab_alignment.is_some() || state.pending_tab_leader.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF tab alignment must occur once and before its leader".to_string(),
                ));
            }
            state.pending_tab_alignment = Some(alignment);
            Ok(())
        }

        fn select_leader(
            state: &mut State,
            parameter: Option<i32>,
            name: &str,
            leader: TabLeader,
        ) -> RtfResult<()> {
            require_flag(parameter, name)?;
            if state.pending_tab_leader.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF tab definition contains multiple leader selectors".to_string(),
                ));
            }
            state.pending_tab_leader = Some(leader);
            Ok(())
        }

        fn append(state: &mut State, position: Option<i32>, bar: bool) -> RtfResult<()> {
            let position = position.ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "RTF {} control requires a numeric parameter",
                    if bar { "tb" } else { "tx" }
                ))
            })?;
            if bar && state.pending_tab_alignment.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF bar tab cannot have a tab-alignment selector".to_string(),
                ));
            }
            let tab = TabStop {
                position,
                alignment: if bar {
                    TabAlignment::Bar
                } else {
                    state.pending_tab_alignment.unwrap_or(TabAlignment::Left)
                },
                leader: state.pending_tab_leader.unwrap_or(TabLeader::None),
            };
            state.paragraph.tab_stops.push(tab).map_err(|_| {
                RtfError::MalformedDocument(
                    "RTF paragraph exceeds the 64-tab safety limit".to_string(),
                )
            })?;
            state.pending_tab_alignment = None;
            state.pending_tab_leader = None;
            Ok(())
        }

        match control {
            ControlWord::TabLeft(parameter) => {
                select_alignment(state, *parameter, "tql", TabAlignment::Left)?;
            },
            ControlWord::TabRight(parameter) => {
                select_alignment(state, *parameter, "tqr", TabAlignment::Right)?;
            },
            ControlWord::TabCenter(parameter) => {
                select_alignment(state, *parameter, "tqc", TabAlignment::Center)?;
            },
            ControlWord::TabDecimal(parameter) => {
                select_alignment(state, *parameter, "tqdec", TabAlignment::Decimal)?;
            },
            ControlWord::TabLeaderDot(parameter) => {
                select_leader(state, *parameter, "tldot", TabLeader::Dot)?;
            },
            ControlWord::TabLeaderMiddleDot(parameter) => {
                select_leader(state, *parameter, "tlmdot", TabLeader::MiddleDot)?;
            },
            ControlWord::TabLeaderHyphen(parameter) => {
                select_leader(state, *parameter, "tlhyph", TabLeader::Hyphen)?;
            },
            ControlWord::TabLeaderUnderscore(parameter) => {
                select_leader(state, *parameter, "tlul", TabLeader::Underscore)?;
            },
            ControlWord::TabLeaderThick(parameter) => {
                select_leader(state, *parameter, "tlth", TabLeader::ThickLine)?;
            },
            ControlWord::TabLeaderEqual(parameter) => {
                select_leader(state, *parameter, "tleq", TabLeader::Equal)?;
            },
            ControlWord::TabPosition(position) => append(state, *position, false)?,
            ControlWord::TabBar(position) => append(state, *position, true)?,
            _ => return Ok(false),
        }
        Ok(true)
    }
}
