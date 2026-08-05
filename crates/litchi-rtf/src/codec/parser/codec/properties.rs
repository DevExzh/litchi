use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_document_variable_destination(&mut self) -> RtfResult<()> {
        // The destination group is one level below the RTF root. Body text is
        // flushed before nested groups, so a nonzero body length also rejects
        // document variables that appear after body content has begun.
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF docvar destination must occur in the root document header".to_string(),
            ));
        }
        if self.document_variables.len() >= MAX_DOCUMENT_VARIABLES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document exceeds {MAX_DOCUMENT_VARIABLES} document variables"
            )));
        }
        self.pos += 2; // \* and \docvar
        let name = self.parse_document_variable_text_group(MAX_DOCUMENT_VARIABLE_NAME_BYTES)?;
        let value = self.parse_document_variable_text_group(MAX_DOCUMENT_VARIABLE_VALUE_BYTES)?;
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF document variable must contain exactly two immediate text groups".to_string(),
            ));
        }
        self.pos += 1;
        let variable = DocumentVariable::new(Cow::Owned(name), Cow::Owned(value))?;
        let added = variable
            .name
            .len()
            .checked_add(variable.value.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument("document-variable size overflow".to_string())
            })?;
        self.document_variable_text_bytes = self
            .document_variable_text_bytes
            .checked_add(added)
            .ok_or_else(|| {
                RtfError::MalformedDocument("document-variable size overflow".to_string())
            })?;
        if self.document_variable_text_bytes > MAX_DOCUMENT_VARIABLE_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document-variable text exceeds {MAX_DOCUMENT_VARIABLE_TEXT_BYTES} bytes"
            )));
        }
        self.document_variables.push(variable);
        Ok(())
    }

    pub(super) fn parse_user_properties_destination(&mut self) -> RtfResult<()> {
        if self.saw_user_properties {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple userprops destinations".to_string(),
            ));
        }
        self.saw_user_properties = true;
        self.pos += 2; // \* and \userprops
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_user_property()?,
                Some(Token::Text(text)) if text.as_bytes().iter().all(u8::is_ascii_whitespace) => {
                    self.pos += 1;
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF userprops may contain only immediate propinfo groups".to_string(),
                    ));
                },
                None => {
                    return Err(RtfError::MalformedDocument(
                        "unterminated RTF userprops destination".to_string(),
                    ));
                },
            }
        }
    }

    pub(super) fn parse_user_property(&mut self) -> RtfResult<()> {
        if self.user_properties.len() >= MAX_USER_PROPERTIES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document exceeds {MAX_USER_PROPERTIES} user properties"
            )));
        }
        let name = self.parse_user_property_text_group(
            ControlWord::PropertyName,
            MAX_USER_PROPERTY_NAME_BYTES,
        )?;
        self.skip_user_property_whitespace();
        let type_code = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::PropertyType(Some(type_code)))) => *type_code,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF propinfo requires an immediate proptype parameter".to_string(),
                ));
            },
        };
        self.pos += 1;
        self.skip_user_property_whitespace();
        let lexical = self.parse_user_property_text_group(
            ControlWord::StaticValue,
            MAX_USER_PROPERTY_VALUE_BYTES,
        )?;
        self.skip_user_property_whitespace();
        let link_value = if matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([Token::OpenBrace, Token::Control(ControlWord::LinkValue)])
        ) {
            Some(self.parse_user_property_text_group(
                ControlWord::LinkValue,
                MAX_USER_PROPERTY_VALUE_BYTES,
            )?)
        } else {
            None
        };
        if self
            .user_properties
            .iter()
            .any(|property| property.name == name)
        {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF user-property name: {name}"
            )));
        }
        let property = UserProperty::new(
            Cow::Owned(name),
            UserPropertyValue::from_lexical(type_code, Cow::Owned(lexical))?,
            link_value.map(Cow::Owned),
        )?;
        self.user_property_text_bytes = self
            .user_property_text_bytes
            .checked_add(property.text_bytes().ok_or_else(|| {
                RtfError::MalformedDocument("user-property size overflow".to_string())
            })?)
            .ok_or_else(|| {
                RtfError::MalformedDocument("user-property size overflow".to_string())
            })?;
        if self.user_property_text_bytes > MAX_USER_PROPERTY_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF user-property text exceeds {MAX_USER_PROPERTY_TEXT_BYTES} bytes"
            )));
        }
        self.user_properties.push(property);
        Ok(())
    }

    pub(super) fn skip_user_property_whitespace(&mut self) {
        while matches!(
            self.tokens.get(self.pos),
            Some(Token::Text(text)) if text.as_bytes().iter().all(u8::is_ascii_whitespace)
        ) {
            self.pos += 1;
        }
    }

    pub(super) fn parse_user_property_text_group(
        &mut self,
        destination: ControlWord<'a>,
        limit: usize,
    ) -> RtfResult<String> {
        self.expect_token(Token::OpenBrace)?;
        self.expect_token(Token::Control(destination))?;
        self.parse_inert_text_group_contents(limit, "user-property")
    }

    pub(super) fn parse_document_variable_text_group(&mut self, limit: usize) -> RtfResult<String> {
        self.expect_token(Token::OpenBrace)?;
        self.parse_inert_text_group_contents(limit, "document-variable")
    }

    pub(super) fn parse_inert_text_group_contents(
        &mut self,
        limit: usize,
        kind: &str,
    ) -> RtfResult<String> {
        let mut bytes = SmallVec::<[u8; 128]>::new();
        let mut output = String::new();
        let mut unicode_skip = self.states.last().map_or(1, |state| state.unicode_skip);
        let mut skip_fallback = 0i32;
        let mut pending_high_surrogate = None;
        loop {
            let token = self.tokens.get(self.pos).ok_or_else(|| {
                RtfError::MalformedDocument(format!("unterminated RTF {kind} text group"))
            })?;
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    return Err(RtfError::MalformedDocument(format!(
                        "nested groups are not allowed in RTF {kind} text"
                    )));
                },
                Token::Binary(_) => {
                    return Err(RtfError::MalformedDocument(format!(
                        "binary data is not allowed in RTF {kind} text"
                    )));
                },
                Token::Text(text) => {
                    let mut transport = SmallVec::<[u8; 128]>::new();
                    append_transport_bytes(&mut transport, text)?;
                    let skip = usize::try_from(skip_fallback.max(0)).unwrap_or(usize::MAX);
                    let skipped = skip.min(transport.len());
                    skip_fallback -= i32::try_from(skipped).unwrap_or(i32::MAX);
                    bytes.extend(transport.into_iter().skip(skipped));
                    self.pos += 1;
                },
                Token::Control(ControlWord::UnicodeSkip(value)) => {
                    unicode_skip = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(value)) => {
                    if !bytes.is_empty() {
                        let decoded = self
                            .states
                            .last()
                            .map_or(RtfEncoding::Standard(Mbcs::WINDOWS_1252), |state| {
                                state.encoding
                            })
                            .decode(&bytes);
                        output.push_str(&decoded);
                        bytes.clear();
                    }
                    let unit = *value as i16 as u16;
                    if (0xD800..=0xDBFF).contains(&unit) {
                        if pending_high_surrogate.replace(unit).is_some() {
                            output.push('\u{FFFD}');
                        }
                    } else if let Some(high) = pending_high_surrogate.take() {
                        output.push(
                            char::decode_utf16([high, unit])
                                .next()
                                .and_then(Result::ok)
                                .unwrap_or('\u{FFFD}'),
                        );
                    } else {
                        output.push(
                            char::decode_utf16([unit])
                                .next()
                                .and_then(Result::ok)
                                .unwrap_or('\u{FFFD}'),
                        );
                    }
                    skip_fallback = unicode_skip.max(0);
                    self.pos += 1;
                },
                Token::Control(control) => {
                    if let Some(text) = control_symbol_text(control) {
                        if !bytes.is_empty() {
                            let decoded = self
                                .states
                                .last()
                                .map_or(RtfEncoding::Standard(Mbcs::WINDOWS_1252), |state| {
                                    state.encoding
                                })
                                .decode(&bytes);
                            output.push_str(&decoded);
                            bytes.clear();
                        }
                        output.push_str(text);
                        self.pos += 1;
                    } else {
                        return Err(RtfError::MalformedDocument(format!(
                            "active controls are not allowed in RTF {kind} text"
                        )));
                    }
                },
            }
            if output.len().saturating_add(bytes.len()) > limit {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {kind} text group exceeds {limit} bytes"
                )));
            }
        }
        if !bytes.is_empty() {
            let decoded = self
                .states
                .last()
                .map_or(RtfEncoding::Standard(Mbcs::WINDOWS_1252), |state| {
                    state.encoding
                })
                .decode(&bytes);
            output.push_str(&decoded);
        }
        if pending_high_surrogate.is_some() {
            output.push('\u{FFFD}');
        }
        if output.len() > limit {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {kind} text group exceeds {limit} bytes"
            )));
        }
        Ok(output)
    }
}
