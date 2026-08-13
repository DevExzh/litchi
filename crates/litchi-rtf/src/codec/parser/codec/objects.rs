use super::{
    ControlWord, Cow, Destination, MAX_GROUP_NESTING_DEPTH, MAX_OBJECT_DATA_BYTES,
    MAX_OBJECT_TEXT_BYTES, Parser, RtfError, RtfResult, Token, control_symbol_text,
};

impl<'a> Parser<'a> {
    /// Parse an `object` destination without activating or updating its content.
    pub(super) fn parse_object_destination(
        &mut self,
    ) -> RtfResult<super::super::super::object::EmbeddedObject<'a>> {
        use super::super::super::object::{ObjectKind, ObjectResultKind};

        let state = self.current_state()?;
        if state.destination != Destination::DocumentBody || state.in_table {
            return Err(RtfError::MalformedDocument(
                "RTF object destination may occur only in the non-table document body".to_string(),
            ));
        }
        let mut object = super::super::super::object::EmbeddedObject::new();
        object.position = self.body_text_len;
        let mut depth = 0usize;
        let mut saw_class = false;
        let mut saw_name = false;
        let mut saw_class_id = false;
        let mut saw_data = false;
        let mut saw_result = false;
        if matches!(
            self.pos
                .checked_sub(1)
                .and_then(|index| self.tokens.get(index)),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF object destination must not be starred".to_string(),
            ));
        }
        self.pos += 1; // consume \object

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ObjectClass) =>
                {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || saw_class
                        || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object class destination placement".to_string(),
                        ));
                    }
                    object.class_name = Cow::Owned(self.parse_object_text_destination()?);
                    saw_class = true;
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::ObjectName) => {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || saw_name
                        || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object name destination placement".to_string(),
                        ));
                    }
                    object.name = Cow::Owned(self.parse_object_text_destination()?);
                    saw_name = true;
                },
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ObjectAlias) =>
                {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || object.alias.is_some()
                        || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object alias destination placement".to_string(),
                        ));
                    }
                    object.alias = Some(Cow::Owned(self.parse_object_text_destination()?));
                },
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ObjectSection) =>
                {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || object.section.is_some()
                        || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object section destination placement".to_string(),
                        ));
                    }
                    object.section = Some(Cow::Owned(self.parse_object_text_destination()?));
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::ObjectTime) => {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || object.time.is_some()
                        || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object time destination placement".to_string(),
                        ));
                    }
                    object.time = Some(self.parse_object_time_destination()?);
                },
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::OleClassId(None)) =>
                {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || saw_class_id
                        || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object CLSID destination placement".to_string(),
                        ));
                    }
                    object.class_id = Cow::Owned(self.parse_object_text_destination()?);
                    saw_class_id = true;
                },
                Token::OpenBrace
                    if matches!(
                        self.nested_control_word(),
                        Some(ControlWord::OleClassId(Some(_)))
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF object CLSID destination must not have a parameter".to_string(),
                    ));
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::ObjectData) => {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || saw_data
                        || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object data destination placement".to_string(),
                        ));
                    }
                    object.data = self.parse_object_hex_destination()?;
                    saw_data = true;
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::Result) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid RTF object result destination placement".to_string(),
                        ));
                    }
                    let (text, pictures) = self.parse_object_result_destination()?;
                    object.result_text = Cow::Owned(text);
                    object.result_picture_indices = pictures;
                    saw_result = true;
                },
                Token::OpenBrace
                    if matches!(
                        self.nested_control_word(),
                        Some(
                            ControlWord::InvalidObjectDestinationParameter
                                | ControlWord::InvalidObjectResultDestinationParameter
                        )
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "parameterized RTF object subdestination is invalid".to_string(),
                    ));
                },
                Token::OpenBrace => {
                    if depth >= MAX_GROUP_NESTING_DEPTH {
                        return Err(RtfError::MalformedDocument(
                            "RTF object nesting depth exceeds the safety limit".to_string(),
                        ));
                    }
                    self.mark_unknown_syntax()?;
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    if saw_data && object.data.is_empty() {
                        self.mark_unknown_syntax()?;
                    }
                    return Ok(object);
                },
                Token::CloseBrace => {
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectEmbedded) => {
                    if object.kind != ObjectKind::Unknown || saw_class_id || saw_data || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object kind".to_string(),
                        ));
                    }
                    object.kind = ObjectKind::Embedded;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectLink) => {
                    if object.kind != ObjectKind::Unknown || saw_class_id || saw_data || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object kind".to_string(),
                        ));
                    }
                    object.kind = ObjectKind::Link;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectAutoLink) => {
                    if object.kind != ObjectKind::Unknown || saw_class_id || saw_data || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object kind".to_string(),
                        ));
                    }
                    object.kind = ObjectKind::AutoLink;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectHtml) => {
                    if object.kind != ObjectKind::Unknown || saw_class_id || saw_data || saw_result
                    {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object kind".to_string(),
                        ));
                    }
                    object.kind = ObjectKind::Html;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectSubscriber(None)) => {
                    object.kind = Self::set_object_kind(
                        object.kind,
                        ObjectKind::Subscriber,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectPublisher(None)) => {
                    object.kind = Self::set_object_kind(
                        object.kind,
                        ObjectKind::Publisher,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectInstallableCommand(None)) => {
                    object.kind = Self::set_object_kind(
                        object.kind,
                        ObjectKind::InstallableCommand,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectOleControl(None)) => {
                    object.kind = Self::set_object_kind(
                        object.kind,
                        ObjectKind::OleControl,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(
                    ControlWord::ObjectSubscriber(Some(_))
                    | ControlWord::ObjectPublisher(Some(_))
                    | ControlWord::ObjectInstallableCommand(Some(_))
                    | ControlWord::ObjectOleControl(Some(_)),
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object kind control must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::ObjectLinkSelf(None)) => {
                    if object.link_self || saw_class_id || saw_data || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object linkself modifier".to_string(),
                        ));
                    }
                    object.link_self = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectLinkSelf(Some(_))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object linkself modifier must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::ObjectWidth(value)) => {
                    object.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectHeight(value)) => {
                    object.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectAlignment(Some(value))) => {
                    Self::set_object_value(
                        &mut object.alignment,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectTranslationY(Some(value))) => {
                    Self::set_object_value(
                        &mut object.translation_y,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectCropTop(Some(value))) => {
                    Self::set_object_value(
                        &mut object.crop_top,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectCropBottom(Some(value))) => {
                    Self::set_object_value(
                        &mut object.crop_bottom,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectCropLeft(Some(value))) => {
                    Self::set_object_value(
                        &mut object.crop_left,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectCropRight(Some(value))) => {
                    Self::set_object_value(
                        &mut object.crop_right,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectScaleX(Some(value))) => {
                    Self::set_object_value(
                        &mut object.scale_x,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectScaleY(Some(value))) => {
                    Self::set_object_value(
                        &mut object.scale_y,
                        *value,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(
                    ControlWord::ObjectAlignment(None)
                    | ControlWord::ObjectTranslationY(None)
                    | ControlWord::ObjectCropTop(None)
                    | ControlWord::ObjectCropBottom(None)
                    | ControlWord::ObjectCropLeft(None)
                    | ControlWord::ObjectCropRight(None)
                    | ControlWord::ObjectScaleX(None)
                    | ControlWord::ObjectScaleY(None),
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object numeric modifier requires a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::ObjectLocked(value)) => {
                    object.locked = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectUpdate(value)) => {
                    object.update_requested = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectSetSize(value)) => {
                    object.set_size = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultMerge(None)) => {
                    if object.merge_result || saw_class_id || saw_data || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF object result merge modifier".to_string(),
                        ));
                    }
                    object.merge_result = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultRtf(None)) => {
                    Self::set_object_result_kind(
                        &mut object,
                        ObjectResultKind::Rtf,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultText(None)) => {
                    Self::set_object_result_kind(
                        &mut object,
                        ObjectResultKind::Text,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultPicture(None)) => {
                    Self::set_object_result_kind(
                        &mut object,
                        ObjectResultKind::Picture,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultBitmap(None)) => {
                    Self::set_object_result_kind(
                        &mut object,
                        ObjectResultKind::Bitmap,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectResultHtml(None)) => {
                    Self::set_object_result_kind(
                        &mut object,
                        ObjectResultKind::Html,
                        saw_class_id || saw_data || saw_result,
                    )?;
                    self.pos += 1;
                },
                Token::Control(
                    ControlWord::ObjectResultMerge(Some(_))
                    | ControlWord::ObjectResultRtf(Some(_))
                    | ControlWord::ObjectResultText(Some(_))
                    | ControlWord::ObjectResultPicture(Some(_))
                    | ControlWord::ObjectResultBitmap(Some(_))
                    | ControlWord::ObjectResultHtml(Some(_)),
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object result modifier must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::InvalidObjectModifierParameter) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object modifier has an invalid parameter".to_string(),
                    ));
                },
                Token::Control(_) | Token::Text(_) | Token::Binary(_) => {
                    self.mark_unknown_syntax()?;
                    self.pos += 1;
                },
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn set_object_kind(
        current: super::super::super::object::ObjectKind,
        value: super::super::super::object::ObjectKind,
        too_late: bool,
    ) -> RtfResult<super::super::super::object::ObjectKind> {
        if current != super::super::super::object::ObjectKind::Unknown || too_late {
            return Err(RtfError::MalformedDocument(
                "invalid or duplicate RTF object kind".to_string(),
            ));
        }
        Ok(value)
    }

    pub(super) fn set_object_value(
        target: &mut Option<i32>,
        value: i32,
        too_late: bool,
    ) -> RtfResult<()> {
        if target.is_some() || too_late {
            return Err(RtfError::MalformedDocument(
                "invalid or duplicate RTF object numeric modifier".to_string(),
            ));
        }
        *target = Some(value);
        Ok(())
    }

    pub(super) fn set_object_result_kind(
        object: &mut super::super::super::object::EmbeddedObject<'_>,
        kind: super::super::super::object::ObjectResultKind,
        too_late: bool,
    ) -> RtfResult<()> {
        if object.result_kind.is_some() || too_late {
            return Err(RtfError::MalformedDocument(
                "invalid or duplicate RTF object result kind".to_string(),
            ));
        }
        object.result_kind = Some(kind);
        Ok(())
    }

    pub(super) fn parse_object_result_destination(&mut self) -> RtfResult<(String, Vec<usize>)> {
        let mut text = String::new();
        let mut picture_indices = Vec::new();
        self.pos += 1; // opening brace
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::Result))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF object result destination".to_string(),
            ));
        }
        self.pos += 1;

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    return Ok((text.trim().to_string(), picture_indices));
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::Picture) => {
                    let first_picture = self.pictures.len();
                    self.parse_group()?;
                    picture_indices.extend(first_picture..self.pictures.len());
                },
                Token::OpenBrace => {
                    self.mark_unknown_syntax()?;
                    self.skip_group()?
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Token::Control(ControlWord::Par | ControlWord::Line) => {
                    text.push('\n');
                    self.pos += 1;
                },
                Token::Control(ControlWord::Tab) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Text(value) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Token::Control(_) | Token::Binary(_) => {
                    self.mark_unknown_syntax()?;
                    self.pos += 1;
                },
            }
            if text.len() > MAX_OBJECT_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF object result text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_object_text_destination(&mut self) -> RtfResult<String> {
        let mut text = String::new();
        self.pos += 1; // opening brace
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace => {
                    return Err(RtfError::MalformedDocument(
                        "nested groups are not allowed in RTF object text metadata".to_string(),
                    ));
                },
                Token::CloseBrace => {
                    self.pos += 1;
                    return Ok(text.trim().to_string());
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Token::Control(
                    ControlWord::IgnorableDestination
                    | ControlWord::ObjectClass
                    | ControlWord::ObjectName
                    | ControlWord::ObjectAlias
                    | ControlWord::ObjectSection
                    | ControlWord::ObjectTime
                    | ControlWord::OleClassId(None),
                ) => {
                    self.pos += 1;
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Text(value) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Token::Control(_) | Token::Binary(_) => {
                    self.mark_unknown_syntax()?;
                    self.pos += 1;
                },
            }
            if text.len() > MAX_OBJECT_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF embedded object metadata exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_object_time_destination(&mut self) -> RtfResult<crate::RtfTimestamp> {
        self.pos += 2; // opening brace and ignorable marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::ObjectTime))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF object time destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut timestamp = crate::RtfTimestamp::default();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(timestamp);
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::Year(value) => timestamp.year = Some(*value),
                    ControlWord::Month(value) => timestamp.month = Some(*value),
                    ControlWord::Day(value) => timestamp.day = Some(*value),
                    ControlWord::Hour(value) => timestamp.hour = Some(*value),
                    ControlWord::Minute(value) => timestamp.minute = Some(*value),
                    ControlWord::Second(value) => timestamp.second = Some(*value),
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF object time destination contains an active control".to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object time destination contains grouped, binary, or text data"
                            .to_string(),
                    ));
                },
            }
            self.pos += 1;
        }
    }

    pub(super) fn parse_object_hex_destination(&mut self) -> RtfResult<Vec<u8>> {
        let mut data = Vec::new();
        let mut high_nibble = None;
        self.pos += 1; // opening brace
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace => {
                    self.mark_unknown_syntax()?;
                    self.skip_group()?;
                },
                Token::CloseBrace => {
                    self.pos += 1;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF objdata contains an odd number of hexadecimal digits".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Token::Text(text) => {
                    let hex_digits = text
                        .bytes()
                        .filter(|byte| !byte.is_ascii_whitespace())
                        .count();
                    let decoded_bytes = (usize::from(high_nibble.is_some()) + hex_digits) / 2;
                    Self::reserve_object_payload(&mut data, decoded_bytes)?;
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF objdata contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            data.push((high << 4) | nibble);
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Token::Binary(bytes) => {
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF objdata binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    Self::reserve_object_payload(&mut data, bytes.len())?;
                    data.extend_from_slice(bytes);
                    self.pos += 1;
                },
                Token::Control(ControlWord::IgnorableDestination | ControlWord::ObjectData) => {
                    self.pos += 1;
                },
                Token::Control(_) => {
                    self.mark_unknown_syntax()?;
                    self.pos += 1;
                },
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn reserve_object_payload(data: &mut Vec<u8>, additional: usize) -> RtfResult<()> {
        Self::reserve_object_payload_with_limit(
            data,
            additional,
            MAX_OBJECT_DATA_BYTES,
            "RTF embedded object data",
            "RTF embedded object data exceeds the safety limit",
        )
    }

    fn reserve_object_payload_with_limit(
        data: &mut Vec<u8>,
        additional: usize,
        limit: usize,
        resource: &'static str,
        message: &'static str,
    ) -> RtfResult<()> {
        let remaining = limit.saturating_sub(data.len());
        if additional > remaining {
            return Err(RtfError::MalformedDocument(message.to_string()));
        }
        data.try_reserve_exact(additional)
            .map_err(|_err| RtfError::AllocationFailed {
                resource,
                requested: data.len().saturating_add(additional),
            })?;
        Ok(())
    }

    pub(super) fn hex_nibble(byte: u8) -> Option<u8> {
        match byte {
            b'0'..=b'9' => Some(byte - b'0'),
            b'a'..=b'f' => Some(byte - b'a' + 10),
            b'A'..=b'F' => Some(byte - b'A' + 10),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Parser;

    #[test]
    fn object_payload_capacity_rejects_one_over_before_reserving() {
        let mut data = vec![0_u8, 1_u8];
        assert!(
            Parser::reserve_object_payload_with_limit(&mut data, 0, 2, "object", "object",).is_ok()
        );
        assert!(
            Parser::reserve_object_payload_with_limit(&mut data, 1, 2, "object", "object",)
                .is_err()
        );
    }
}
