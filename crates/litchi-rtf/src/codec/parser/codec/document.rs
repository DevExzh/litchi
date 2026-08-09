use super::{
    ControlWord, Cow, MAX_REVISION_AUTHOR_BYTES, Parser, RtfError, RtfResult, Token,
    control_symbol_text, duplicate_mail_merge, nonnegative_mail_merge, set_mail_merge_text,
};

impl<'a> Parser<'a> {
    pub(super) fn parse_revision_table(&mut self) -> RtfResult<()> {
        if self.saw_revision_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple revision-author tables".to_string(),
            ));
        }
        self.saw_revision_table = true;
        self.pos += 1; // `revtbl`
        let mut direct_text = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    self.push_direct_revision_authors(&mut direct_text)?;
                    let author = self.parse_revision_author_group()?;
                    self.push_revision_author(&author)?;
                    continue;
                },
                Some(Token::Control(ControlWord::Page(_))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF page is not permitted in a stylesheet".to_string(),
                    ));
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.push_direct_revision_authors(&mut direct_text)?;
                    if !direct_text.trim().is_empty() {
                        self.push_revision_author(direct_text.trim())?;
                    }
                    return Ok(());
                },
                Some(Token::Text(text)) => direct_text.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    direct_text.push_str(&decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    direct_text.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(_) | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-author table contains a non-text control or binary data"
                            .to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if direct_text.len() > MAX_REVISION_AUTHOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF revision author exceeds the safety limit".to_string(),
                ));
            }
            self.push_direct_revision_authors(&mut direct_text)?;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_revision_save_table(&mut self) -> RtfResult<()> {
        if self.saw_revision_save_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple revision-save tables".to_string(),
            ));
        }
        self.saw_revision_save_table = true;
        self.pos += 1; // rsidtbl
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::Control(ControlWord::RevisionSaveId(value))) => {
                    let save_id = u32::try_from(*value).map_err(|_err| {
                        RtfError::MalformedDocument(
                            "RTF revision-save IDs must be positive signed integers".to_string(),
                        )
                    })?;
                    if save_id == 0 {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save IDs must be positive signed integers".to_string(),
                        ));
                    }
                    if self.revision_save_ids.len() >= crate::revision_save::MAX_REVISION_SAVE_IDS {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save ID count exceeds the safety limit".to_string(),
                        ));
                    }
                    if self.revision_save_ids.contains(&save_id) {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save IDs must be unique".to_string(),
                        ));
                    }
                    self.revision_save_ids.push(save_id);
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_) | Token::Control(_) | Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-save table contains text, nesting, binary data, or an unsupported control"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_mail_merge_destination(&mut self) -> RtfResult<crate::MailMerge<'a>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::MailMerge))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF mailmerge destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut merge = crate::MailMerge::default();
        let mut saw_link_to_query = false;

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    merge.validate()?;
                    return Ok(merge);
                },
                Some(Token::OpenBrace) => {
                    let control = self.mail_merge_child_control()?;
                    match control {
                        ControlWord::MailMergeConnectString => {
                            if merge.connect_string.is_some() {
                                return Err(duplicate_mail_merge("connection string"));
                            }
                            merge.connect_string = Some(self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeConnectString,
                                2,
                            )?);
                        },
                        ControlWord::MailMergeConnectStringData => {
                            if merge.connect_string_data.is_some() {
                                return Err(duplicate_mail_merge("connection-string data"));
                            }
                            merge.connect_string_data =
                                Some(self.parse_mail_merge_text_destination(
                                    &ControlWord::MailMergeConnectStringData,
                                    2,
                                )?);
                        },
                        ControlWord::MailMergeQuery => {
                            if merge.query.is_some() {
                                return Err(duplicate_mail_merge("query"));
                            }
                            merge.query = Some(self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeQuery,
                                2,
                            )?);
                        },
                        ControlWord::MailMergeDataSource => {
                            if merge.data_source.is_some() {
                                return Err(duplicate_mail_merge("data source"));
                            }
                            merge.data_source = Some(self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeDataSource,
                                2,
                            )?);
                        },
                        ControlWord::MailMergeHeaderSource => {
                            if merge.header_source.is_some() {
                                return Err(duplicate_mail_merge("header source"));
                            }
                            merge.header_source = Some(self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeHeaderSource,
                                2,
                            )?);
                        },
                        ControlWord::MailMergeDataSourceObject => {
                            if merge.data_source_object.is_some() {
                                return Err(duplicate_mail_merge("data-source object"));
                            }
                            merge.data_source_object =
                                Some(self.parse_mail_merge_data_source_object(2)?);
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "unsupported nested RTF mail-merge destination".to_string(),
                            ));
                        },
                    }
                },
                Some(Token::Control(ControlWord::MailMergeLinkToQuery(value))) => {
                    if saw_link_to_query {
                        return Err(duplicate_mail_merge("link-to-query flag"));
                    }
                    saw_link_to_query = true;
                    merge.link_to_query = *value;
                    self.pos += 1;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => self.pos += 1,
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF mailmerge destination contains active or misplaced data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn mail_merge_child_control(&self) -> RtfResult<ControlWord<'a>> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "expected nested RTF mail-merge group".to_string(),
            ));
        }
        match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
            (
                Some(Token::Control(ControlWord::IgnorableDestination)),
                Some(Token::Control(control)),
            ) => Ok(*control),
            _ => Err(RtfError::MalformedDocument(
                "nested RTF mail-merge destinations must be starred".to_string(),
            )),
        }
    }

    pub(super) fn parse_mail_merge_text_destination(
        &mut self,
        expected: &ControlWord<'_>,
        depth: usize,
    ) -> RtfResult<Cow<'a, str>> {
        if depth > crate::MAX_MAIL_MERGE_NESTING_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge nesting depth exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(&Token::OpenBrace)?;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge text destination must be starred".to_string(),
            ));
        }
        self.pos += 1;
        if !matches!(self.tokens.get(self.pos), Some(Token::Control(control)) if control == expected)
        {
            return Err(RtfError::MalformedDocument(
                "unexpected RTF mail-merge text destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(Cow::Owned(value));
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
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF mail-merge text contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > crate::MAX_MAIL_MERGE_STRING_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF mail-merge string exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_mail_merge_data_source_object(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::MailMergeDataSourceObject<'a>> {
        if depth > crate::MAX_MAIL_MERGE_NESTING_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge nesting depth exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(&Token::OpenBrace)?;
        self.expect_token(&Token::Control(ControlWord::IgnorableDestination))?;
        self.expect_token(&Token::Control(ControlWord::MailMergeDataSourceObject))?;
        let mut object = crate::MailMergeDataSourceObject::default();
        let mut saw_dynamic_address = false;
        let mut saw_first_row_header = false;

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    object.validate()?;
                    return Ok(object);
                },
                Some(Token::Control(control)) => {
                    match control {
                        ControlWord::MailMergeActiveRecord(value) => {
                            if object.active_record.is_some() {
                                return Err(duplicate_mail_merge("active record"));
                            }
                            object.active_record =
                                Some(nonnegative_mail_merge(*value, "active record")?);
                        },
                        ControlWord::MailMergeColumnDelimiter(value) => {
                            if object.column_delimiter.replace(*value).is_some() {
                                return Err(duplicate_mail_merge("column delimiter"));
                            }
                        },
                        ControlWord::MailMergeColumnCount(value) => {
                            if object.column_count.is_some() {
                                return Err(duplicate_mail_merge("column count"));
                            }
                            object.column_count =
                                Some(nonnegative_mail_merge(*value, "column count")?);
                        },
                        ControlWord::MailMergeDynamicAddress(value) => {
                            if saw_dynamic_address {
                                return Err(duplicate_mail_merge("dynamic-address flag"));
                            }
                            saw_dynamic_address = true;
                            object.dynamic_address = Some(*value);
                        },
                        ControlWord::MailMergeFirstRowHeader(value) => {
                            if saw_first_row_header {
                                return Err(duplicate_mail_merge("first-row-header flag"));
                            }
                            saw_first_row_header = true;
                            object.first_row_header = Some(*value);
                        },
                        ControlWord::MailMergeHash(value) => {
                            if object.hash.replace(*value).is_some() {
                                return Err(duplicate_mail_merge("data-source hash"));
                            }
                        },
                        ControlWord::MailMergeId(value) => {
                            if object.id.replace(*value).is_some() {
                                return Err(duplicate_mail_merge("data-source ID"));
                            }
                        },
                        ControlWord::MailMergeSourceType(value) => {
                            if object
                                .source_type
                                .replace(crate::MailMergeDataSourceType::from_rtf(*value))
                                .is_some()
                            {
                                return Err(duplicate_mail_merge("data-source type"));
                            }
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "misplaced control in RTF mmodso destination".to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    let control = self.mail_merge_child_control()?;
                    match control {
                        ControlWord::MailMergeFilter => set_mail_merge_text(
                            &mut object.filter,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeFilter,
                                depth + 1,
                            )?,
                            "filter",
                        )?,
                        ControlWord::MailMergeName => set_mail_merge_text(
                            &mut object.name,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeName,
                                depth + 1,
                            )?,
                            "data-source name",
                        )?,
                        ControlWord::MailMergeSort => set_mail_merge_text(
                            &mut object.sort,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeSort,
                                depth + 1,
                            )?,
                            "sort",
                        )?,
                        ControlWord::MailMergeTable => set_mail_merge_text(
                            &mut object.table,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeTable,
                                depth + 1,
                            )?,
                            "table",
                        )?,
                        ControlWord::MailMergeUdl => set_mail_merge_text(
                            &mut object.udl,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeUdl,
                                depth + 1,
                            )?,
                            "UDL",
                        )?,
                        ControlWord::MailMergeUdlData => set_mail_merge_text(
                            &mut object.udl_data,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeUdlData,
                                depth + 1,
                            )?,
                            "UDL data",
                        )?,
                        ControlWord::MailMergeUniqueTag => set_mail_merge_text(
                            &mut object.unique_tag,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeUniqueTag,
                                depth + 1,
                            )?,
                            "unique tag",
                        )?,
                        ControlWord::MailMergeRecipientData => {
                            if object.recipient_data.len() >= crate::MAX_MAIL_MERGE_RECIPIENT_DATA {
                                return Err(RtfError::MalformedDocument(
                                    "RTF mail-merge recipient-data count exceeds the safety limit"
                                        .to_string(),
                                ));
                            }
                            let value = self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeRecipientData,
                                depth + 1,
                            )?;
                            object.recipient_data.push(value);
                        },
                        ControlWord::MailMergeFieldMapData => {
                            if object.field_mappings.len() >= crate::MAX_MAIL_MERGE_FIELD_MAPPINGS {
                                return Err(RtfError::MalformedDocument(
                                    "RTF mail-merge field-mapping count exceeds the safety limit"
                                        .to_string(),
                                ));
                            }
                            object
                                .field_mappings
                                .push(self.parse_mail_merge_field_mapping(depth + 1)?);
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "unsupported nested RTF mmodso destination".to_string(),
                            ));
                        },
                    }
                },
                Some(Token::Text(text)) if text.trim().is_empty() => self.pos += 1,
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF mmodso destination contains active or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_mail_merge_field_mapping(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::MailMergeFieldMapping<'a>> {
        if depth > crate::MAX_MAIL_MERGE_NESTING_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge nesting depth exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(&Token::OpenBrace)?;
        self.expect_token(&Token::Control(ControlWord::IgnorableDestination))?;
        self.expect_token(&Token::Control(ControlWord::MailMergeFieldMapData))?;
        let mut column = None;
        let mut name = None;
        let mut mapped_name = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let column_value = column.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF field mapping is missing mmodsofmcolumn".to_string(),
                        )
                    })?;
                    let name_value = name.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF field mapping is missing mmodsoname".to_string(),
                        )
                    })?;
                    let mapping = crate::MailMergeFieldMapping {
                        column: column_value,
                        name: name_value,
                        mapped_name,
                    };
                    mapping.validate()?;
                    return Ok(mapping);
                },
                Some(Token::Control(ControlWord::MailMergeFieldMapColumn(value))) => {
                    if column.is_some() {
                        return Err(duplicate_mail_merge("field-map column"));
                    }
                    column = Some(crate::MailMergeColumnIndex::from_rtf(*value)?);
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    let control = self.mail_merge_child_control()?;
                    match control {
                        ControlWord::MailMergeName => set_mail_merge_text(
                            &mut name,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeName,
                                depth + 1,
                            )?,
                            "field name",
                        )?,
                        ControlWord::MailMergeMappedName => set_mail_merge_text(
                            &mut mapped_name,
                            self.parse_mail_merge_text_destination(
                                &ControlWord::MailMergeMappedName,
                                depth + 1,
                            )?,
                            "mapped field name",
                        )?,
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "unsupported nested RTF field-map destination".to_string(),
                            ));
                        },
                    }
                },
                Some(Token::Text(text)) if text.trim().is_empty() => self.pos += 1,
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF field-map destination contains active or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_data_store_destination(&mut self) -> RtfResult<Vec<u8>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::DataStore))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF datastore destination".to_string(),
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
                            "RTF data-store payload has an odd hexadecimal digit count".to_string(),
                        ));
                    }
                    if data.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF data-store payload cannot be empty".to_string(),
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
                                    "RTF data-store payload contains a non-hexadecimal character"
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
                Some(Token::OpenBrace | Token::Binary(_) | Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF datastore cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > crate::data_store::MAX_DATA_STORE_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF data-store payload exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_math_properties_destination(
        &mut self,
    ) -> RtfResult<crate::DocumentMathProperties> {
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::MathProperties))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF math-properties destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut properties = crate::DocumentMathProperties::default();

        macro_rules! set_once {
            ($field:ident, $value:expr, $name:literal) => {{
                if properties.$field.is_some() {
                    return Err(RtfError::MalformedDocument(
                        concat!("duplicate RTF math property ", $name).to_string(),
                    ));
                }
                properties.$field = Some($value);
            }};
        }

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    properties.validate()?;
                    return Ok(properties);
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Control(control)) => {
                    match *control {
                        ControlWord::MathBreakBinary(value) => set_once!(
                            binary_operator_break,
                            crate::MathBinaryOperatorBreak::from_rtf(value),
                            "mbrkBin"
                        ),
                        ControlWord::MathBreakBinarySubtraction(value) => set_once!(
                            binary_subtraction_break,
                            crate::MathBinarySubtractionBreak::from_rtf(value),
                            "mbrkBinSub"
                        ),
                        ControlWord::MathDefaultJustification(value) => set_once!(
                            default_justification,
                            crate::MathJustification::from_rtf(value),
                            "mdefJc"
                        ),
                        ControlWord::MathDisplayDefaults(value) => set_once!(
                            display_defaults,
                            crate::MathFlag::from_rtf(value),
                            "mdispDef"
                        ),
                        ControlWord::MathInterEquationSpacing(value) => {
                            set_once!(inter_equation_spacing, value, "minterSp");
                        },
                        ControlWord::MathIntegralLimitPlacement(value) => set_once!(
                            integral_limit_placement,
                            crate::MathLimitPlacement::from_rtf(value),
                            "mintLim"
                        ),
                        ControlWord::MathIntraEquationSpacing(value) => {
                            set_once!(intra_equation_spacing, value, "mintraSp");
                        },
                        ControlWord::MathLeftMargin(value) => {
                            set_once!(left_margin, value, "mlMargin");
                        },
                        ControlWord::MathFont(value) => {
                            let font_index = u32::try_from(value).map_err(|_err| {
                                RtfError::MalformedDocument(
                                    "RTF math font index cannot be negative".to_string(),
                                )
                            })?;
                            set_once!(math_font, font_index, "mmathFont");
                        },
                        ControlWord::MathNaryLimitPlacement(value) => set_once!(
                            nary_limit_placement,
                            crate::MathLimitPlacement::from_rtf(value),
                            "mnaryLim"
                        ),
                        ControlWord::MathPostSpacing(value) => {
                            set_once!(post_spacing, value, "mpostSp");
                        },
                        ControlWord::MathPreSpacing(value) => {
                            set_once!(pre_spacing, value, "mpreSp");
                        },
                        ControlWord::MathRightMargin(value) => {
                            set_once!(right_margin, value, "mrMargin");
                        },
                        ControlWord::MathSmallFractions(value) => set_once!(
                            small_fractions,
                            crate::MathFlag::from_rtf(value),
                            "msmallFrac"
                        ),
                        ControlWord::MathWrapIndent(value) => {
                            set_once!(wrap_indent, value, "mwrapIndent");
                        },
                        ControlWord::MathWrapRight(value) => {
                            set_once!(wrap_right, crate::MathFlag::from_rtf(value), "mwrapRight");
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF math-properties destination contains an unsupported control"
                                    .to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_) | Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math-properties destination contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }
}
