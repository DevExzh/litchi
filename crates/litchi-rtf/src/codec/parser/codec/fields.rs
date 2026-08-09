#![allow(
    clippy::shadow_unrelated,
    reason = "decoding steps deliberately rebind a working value as it is refined through the parse pipeline"
)]
use super::{
    ControlWord, Cow, Destination, DrawingStoryCapture, FormFieldBuilder, ParsedBodyStoryEvent,
    Parser, RtfError, RtfResult, SmallVec, Token, control_symbol_text, parser_classification_error,
    require_parameterless,
};

impl Parser<'_> {
    /// Parse field content.
    ///
    /// Fields in RTF have the format:
    /// {\field{\*\fldinst INSTRUCTION}{\fldrslt RESULT}}
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_field(&mut self) -> RtfResult<()> {
        let state = self.current_state()?;
        let enclosing_destination = state.destination;
        let field_in_table = state.in_table;
        let table_level = state.table_nesting_level.max(1);
        let mut field_owner = match enclosing_destination {
            Destination::FieldResult => crate::FieldOwner::FieldResult,
            Destination::Header => crate::FieldOwner::Header,
            Destination::Footer => crate::FieldOwner::Footer,
            Destination::Footnote => crate::FieldOwner::Footnote,
            Destination::Endnote => crate::FieldOwner::Endnote,
            Destination::DocumentBody if field_in_table => {
                crate::FieldOwner::TableCell(table_level)
            },
            Destination::DocumentBody => crate::FieldOwner::Body,
            Destination::FontTable
            | Destination::ColorTable
            | Destination::StyleSheet
            | Destination::Info
            | Destination::Picture
            | Destination::Result
            | Destination::FieldInstruction
            | Destination::NestedTableProperties
            | Destination::Revision
            | Destination::Other => crate::FieldOwner::Other,
        };
        let field_position = match field_owner {
            crate::FieldOwner::FieldResult => self
                .field_drawing_captures
                .last()
                .map_or(0, |capture| capture.story_offset),
            crate::FieldOwner::Header | crate::FieldOwner::Footer => self.current_hf_story_offset,
            crate::FieldOwner::Footnote | crate::FieldOwner::Endnote => {
                self.current_note_buffer.len()
            },
            crate::FieldOwner::TableCell(1) => self.current_cell_text.len(),
            crate::FieldOwner::TableCell(level) => self
                .nested_table_builders
                .get(usize::from(level - 2))
                .map_or(0, |builder| builder.cell_text.len()),
            crate::FieldOwner::Body => self.body_text_len,
            crate::FieldOwner::Detached
            | crate::FieldOwner::FormField
            | crate::FieldOwner::Other => 0,
        };
        self.pos += 1; // Skip \field
        let mut field_status = crate::FieldStatus::default();
        let mut field_status_seen = 0_u8;
        let mut saw_field_destination = false;

        let mut instruction = SmallVec::<[u8; 128]>::new();
        let mut result = SmallVec::<[u8; 128]>::new();
        self.field_drawing_captures
            .push(DrawingStoryCapture::default());
        let mut form_field = None;
        let mut data_field = None;
        let mut in_instruction;
        let mut in_result;

        // Parse field groups
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    // End of outer field group
                    break;
                },
                Token::Control(ControlWord::Unknown(_, _)) => {
                    let token = self.pos;
                    self.pos += 1;
                    self.preserve_unknown_control_in(token, crate::opaque::Context::Field)?;
                },
                Token::Control(control) => {
                    let status_control = match control {
                        ControlWord::FieldDirty(parameter) => Some(("flddirty", parameter, 1_u8)),
                        ControlWord::FieldEdit(parameter) => Some(("fldedit", parameter, 2_u8)),
                        ControlWord::FieldLock(parameter) => Some(("fldlock", parameter, 4_u8)),
                        ControlWord::FieldPrivate(parameter) => Some(("fldpriv", parameter, 8_u8)),
                        _ => None,
                    };

                    if let Some((name, parameter, bit)) = status_control {
                        if saw_field_destination {
                            return Err(RtfError::MalformedDocument(format!(
                                "RTF field state control \\{name} must precede field destinations"
                            )));
                        }
                        if parameter.is_some() {
                            return Err(RtfError::MalformedDocument(format!(
                                "RTF field state control \\{name} does not accept a parameter"
                            )));
                        }
                        if field_status_seen & bit != 0 {
                            return Err(RtfError::MalformedDocument(format!(
                                "duplicate RTF field state control \\{name}"
                            )));
                        }
                        field_status_seen |= bit;
                        match control {
                            ControlWord::FieldDirty(_) => field_status.dirty = true,
                            ControlWord::FieldEdit(_) => field_status.edited = true,
                            ControlWord::FieldLock(_) => field_status.locked = true,
                            ControlWord::FieldPrivate(_) => field_status.private = true,
                            _ => return Err(parser_classification_error()),
                        }
                    }
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    self.reject_non_body_custom_xml_markup_group()?;
                    let is_status_control = |token: Option<&Token<'_>>| {
                        matches!(
                            token,
                            Some(Token::Control(
                                ControlWord::FieldDirty(_)
                                    | ControlWord::FieldEdit(_)
                                    | ControlWord::FieldLock(_)
                                    | ControlWord::FieldPrivate(_)
                            ))
                        )
                    };
                    let grouped_status = is_status_control(self.tokens.get(self.pos + 1))
                        || (matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::IgnorableDestination))
                        ) && is_status_control(self.tokens.get(self.pos + 2)));
                    if grouped_status {
                        return Err(RtfError::MalformedDocument(
                            "RTF field state controls must occur directly in the field group"
                                .to_string(),
                        ));
                    }
                    saw_field_destination = true;
                    self.pos += 1;
                    // Check for fldinst or fldrslt
                    if self.pos < self.tokens.len() {
                        // Look for \*\fldinst or \fldrslt
                        let is_ignorable = matches!(
                            self.tokens.get(self.pos),
                            Some(Token::Control(ControlWord::IgnorableDestination))
                        );
                        if is_ignorable {
                            self.pos += 1;
                        }

                        if let Some(Token::Control(ControlWord::FieldInstruction)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = true;
                            in_result = false;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldInstruction;
                            }
                        } else if let Some(Token::Control(ControlWord::FieldResult)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = false;
                            in_result = true;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldResult;
                            }
                        } else {
                            // Retain unknown field-owned groups without interpreting them.
                            self.preserve_unknown_destination_in(crate::opaque::Context::Field)?;
                            continue;
                        }

                        // Collect text until the destination's closing brace. Producers often
                        // wrap or split field instructions in formatting groups; those groups
                        // do not change the field-code text and must not discard it.
                        let mut nested_depth = 0usize;
                        while let Some(token) = self.tokens.get(self.pos) {
                            match token {
                                Token::CloseBrace if nested_depth == 0 => {
                                    self.pos += 1;
                                    break;
                                },
                                Token::CloseBrace => {
                                    nested_depth -= 1;
                                    self.pos += 1;
                                },
                                Token::Text(text) => {
                                    let decoded = self.decode_transport_text(text)?;
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Unicode(first)) => {
                                    let decoded = self.parse_style_unicode(
                                        *first,
                                        self.current_state()?.unicode_skip.max(0),
                                    )?;
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                },
                                Token::Control(ControlWord::UnicodeSkip(value)) => {
                                    self.current_state_mut()?.unicode_skip = (*value).max(0);
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Par | ControlWord::Line)
                                    if in_result =>
                                {
                                    result.push(b'\n');
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Page(param)) if in_result => {
                                    require_parameterless(*param, "page")?;
                                    self.current_field_drawing_capture_mut()?.story_events.push(
                                        crate::StoryEvent::PageBreak(crate::PageBreak::new(
                                            result.len(),
                                        )),
                                    );
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Tab) if in_result => {
                                    result.push(b'\t');
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Unknown(_, _)) => {
                                    let token = self.pos;
                                    self.pos += 1;
                                    self.preserve_unknown_control_in(
                                        token,
                                        crate::opaque::Context::Field,
                                    )?;
                                },
                                Token::Control(control)
                                    if control_symbol_text(control).is_some() =>
                                {
                                    let decoded = control_symbol_text(control).unwrap_or_default();
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                    self.pos += 1;
                                },
                                Token::OpenBrace if self.is_custom_xml_markup_group() => {
                                    return Err(RtfError::MalformedDocument(
                                        "RTF custom XML markup destinations are supported only in the main body story"
                                            .to_string(),
                                    ));
                                },
                                Token::OpenBrace if in_result && self.is_root_drawing_group() => {
                                    self.current_field_drawing_capture_mut()?.story_offset =
                                        result.len();
                                    self.parse_group()?;
                                },
                                Token::OpenBrace if in_instruction => {
                                    let destination = match (
                                        self.tokens.get(self.pos + 1),
                                        self.tokens.get(self.pos + 2),
                                    ) {
                                        (
                                            Some(Token::Control(ControlWord::IgnorableDestination)),
                                            Some(Token::Control(control)),
                                        )
                                        | (Some(Token::Control(control)), _) => Some(control),
                                        _ => None,
                                    };
                                    match destination {
                                        Some(ControlWord::FormField) => {
                                            if form_field.is_some() {
                                                return Err(RtfError::MalformedDocument(
                                                    "RTF field contains multiple formfield destinations"
                                                        .to_string(),
                                                ));
                                            }
                                            form_field = Some(self.parse_form_field_destination()?);
                                        },
                                        Some(ControlWord::DataField) => {
                                            if data_field.is_some() {
                                                return Err(RtfError::MalformedDocument(
                                                    "RTF field contains multiple datafield destinations"
                                                        .to_string(),
                                                ));
                                            }
                                            data_field = Some(self.parse_data_field_destination()?);
                                        },
                                        Some(ControlWord::Unknown(_, _))
                                            if matches!(
                                                self.tokens.get(self.pos + 1),
                                                Some(Token::Control(
                                                    ControlWord::IgnorableDestination
                                                ))
                                            ) =>
                                        {
                                            self.pos += 1;
                                            self.preserve_unknown_destination_in(
                                                crate::opaque::Context::Field,
                                            )?;
                                        },
                                        _ => {
                                            nested_depth =
                                                nested_depth.checked_add(1).ok_or_else(|| {
                                                    RtfError::MalformedDocument(
                                                        "field instruction nesting depth overflow"
                                                            .to_string(),
                                                    )
                                                })?;
                                            self.pos += 1;
                                        },
                                    }
                                },
                                Token::OpenBrace
                                    if matches!(
                                        self.tokens.get(self.pos + 1),
                                        Some(Token::Control(ControlWord::Field))
                                    ) =>
                                {
                                    self.current_field_drawing_capture_mut()?.story_offset =
                                        result.len();
                                    self.pos += 1;
                                    self.parse_field()?;
                                    self.skip_until_close_brace()?;
                                },
                                Token::OpenBrace => {
                                    nested_depth =
                                        nested_depth.checked_add(1).ok_or_else(|| {
                                            RtfError::MalformedDocument(
                                                "field instruction nesting depth overflow"
                                                    .to_string(),
                                            )
                                        })?;
                                    self.pos += 1;
                                },
                                Token::Control(_) | Token::Binary(_) => {
                                    self.pos += 1;
                                },
                            }
                            if instruction.len()
                                > super::super::super::form_field::MAX_FORM_FIELD_STRING_BYTES
                                || result.len()
                                    > super::super::super::form_field::MAX_FORM_FIELD_STRING_BYTES
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF field instruction or result exceeds the safety limit"
                                        .to_string(),
                                ));
                            }
                        }
                    }
                },
                Token::Text(_) | Token::Binary(_) => {
                    self.pos += 1;
                },
            }
        }

        let result_text = std::str::from_utf8(&result).map_err(|_err| {
            RtfError::MalformedDocument("RTF field result is not valid UTF-8".to_string())
        })?;
        let field_drawings = self.field_drawing_captures.pop().ok_or_else(|| {
            RtfError::MalformedDocument("RTF field drawing capture stack underflow".to_string())
        })?;

        // Create the generic field record if we have an instruction.
        if !instruction.is_empty()
            && let Ok(inst_str) = std::str::from_utf8(&instruction)
        {
            // Allocate instruction in arena first
            let inst_alloc = self.arena.alloc_str(inst_str);

            // Parse field type from allocated instruction
            let mut field = super::super::super::field::Field::parse_instruction(inst_alloc);
            field.instruction = Cow::Borrowed(inst_alloc);
            field.status = field_status;

            // Add result if available
            if !result.is_empty()
                && let Ok(res_str) = std::str::from_utf8(&result)
            {
                let res_alloc = self.arena.alloc_str(res_str);
                field.result = Cow::Borrowed(res_alloc);
            }

            field.shapes = field_drawings.shapes;
            field.shape_groups = field_drawings.shape_groups;
            field.drawing_order = field_drawings.drawing_order;
            field.result_events = field_drawings.story_events;
            if form_field.is_some() {
                field_owner = crate::FieldOwner::FormField;
            }
            field.owner = field_owner;
            field.position = field_position;
            field.range_end = field_position;
            field.validate()?;

            if self.fields.len() >= crate::field::MAX_GENERIC_FIELDS {
                return Err(RtfError::MalformedDocument(
                    "RTF generic field count exceeds the safety limit".to_string(),
                ));
            }
            let field_index = self.fields.len();
            self.fields.push(field);
            let story_field = crate::StoryField {
                field_index,
                position: field_position,
            };
            match field_owner {
                crate::FieldOwner::Body => self.body_story_events.push(
                    ParsedBodyStoryEvent::Resolved(crate::BodyStoryEvent::Field(field_index)),
                ),
                crate::FieldOwner::Header | crate::FieldOwner::Footer => self
                    .current_hf_story_events
                    .push(crate::StoryEvent::Field(story_field)),
                crate::FieldOwner::Footnote | crate::FieldOwner::Endnote => self
                    .current_note_story_events
                    .push(crate::StoryEvent::Field(story_field)),
                crate::FieldOwner::TableCell(1) => self
                    .current_cell_story_events
                    .push(crate::CellStoryEvent::Field(story_field)),
                crate::FieldOwner::TableCell(level) => self
                    .ensure_nested_builder(level)?
                    .cell_story_events
                    .push(crate::CellStoryEvent::Field(story_field)),
                crate::FieldOwner::FieldResult => self
                    .field_drawing_captures
                    .last_mut()
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF nested field lacks a parent result story".to_string(),
                        )
                    })?
                    .story_events
                    .push(crate::StoryEvent::Field(story_field)),
                crate::FieldOwner::Detached
                | crate::FieldOwner::FormField
                | crate::FieldOwner::Other => {},
            }
        }

        self.current_state_mut()?.destination = enclosing_destination;
        if form_field.is_some() && !result_text.is_empty() {
            self.append_semantic_text(result_text)?;
        }

        if let Some(builder) = form_field {
            if self.form_fields.len() >= super::super::super::form_field::MAX_FORM_FIELDS {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field count exceeds the safety limit".to_string(),
                ));
            }
            let field_type = builder.field_type.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF formfield destination is missing fftype".to_string(),
                )
            })?;
            let to_cow = |value: Option<String>| {
                value.map(|text| Cow::Borrowed(self.arena.alloc_str(&text)))
            };
            let form_field = super::super::super::form_field::FormField {
                field_type,
                text_type: builder.text_type,
                name: to_cow(builder.name),
                max_length: builder.max_length,
                format: to_cow(builder.format),
                default_text: to_cow(builder.default_text),
                default_result: builder.default_result,
                result: builder.result,
                half_point_size: builder.half_point_size,
                protected: builder.protected.unwrap_or(false),
                calculate_on_exit: builder.calculate_on_exit.unwrap_or(false),
                size_automatically: builder.size_automatically.unwrap_or(false),
                own_help: builder.own_help.unwrap_or(false),
                own_status: builder.own_status.unwrap_or(false),
                help_text: to_cow(builder.help_text),
                status_text: to_cow(builder.status_text),
                entry_macro: to_cow(builder.entry_macro),
                exit_macro: to_cow(builder.exit_macro),
                list_entries: builder
                    .list_entries
                    .into_iter()
                    .map(|text| Cow::Borrowed(self.arena.alloc_str(&text)))
                    .collect(),
                has_list_box: builder.has_list_box.unwrap_or(false),
                data: Cow::Borrowed(
                    self.arena
                        .alloc_slice_copy(data_field.as_deref().unwrap_or_default()),
                ),
                result_text: Cow::Borrowed(self.arena.alloc_str(if field_in_table {
                    ""
                } else {
                    result_text
                })),
                position: field_position,
                range_end: if field_in_table {
                    field_position
                } else {
                    self.body_text_len
                },
            };
            form_field.validate()?;
            let added = form_field.text_bytes().ok_or_else(|| {
                RtfError::MalformedDocument("RTF form-field aggregate size overflow".to_string())
            })?;
            self.form_field_text_bytes =
                self.form_field_text_bytes
                    .checked_add(added)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF form-field aggregate size overflow".to_string(),
                        )
                    })?;
            if self.form_field_text_bytes
                > super::super::super::form_field::MAX_FORM_FIELD_TOTAL_BYTES
            {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field aggregate text exceeds the safety limit".to_string(),
                ));
            }
            let index = self.form_fields.len();
            self.form_fields.push(form_field);
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::FormFieldStart(index),
            ));
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::FormFieldEnd(index),
            ));
        } else if data_field.is_some() {
            // Data fields attached to non-form fields are inert legacy payloads and
            // are intentionally not exposed as executable/external content.
        }

        Ok(())
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_form_field_destination(&mut self) -> RtfResult<FormFieldBuilder> {
        self.expect_token(&Token::OpenBrace)?;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF formfield destination must be starred".to_string(),
            ));
        }
        self.pos += 1;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::FormField))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF formfield destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut builder = FormFieldBuilder::default();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    depth -= 1;
                    if depth == 0 {
                        return Ok(builder);
                    }
                },
                Some(Token::OpenBrace) => {
                    let starred = matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    );
                    let control =
                        match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                            (
                                Some(Token::Control(ControlWord::IgnorableDestination)),
                                Some(Token::Control(control)),
                            )
                            | (Some(Token::Control(control)), _) => Some(control),
                            _ => None,
                        };
                    if !starred
                        && matches!(
                            control,
                            Some(
                                ControlWord::FormFieldName
                                    | ControlWord::FormFieldFormat
                                    | ControlWord::FormFieldDefaultText
                                    | ControlWord::FormFieldHelpText
                                    | ControlWord::FormFieldStatusText
                                    | ControlWord::FormFieldEntryMacro
                                    | ControlWord::FormFieldExitMacro
                                    | ControlWord::FormFieldListEntry
                            )
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield text destinations must be starred".to_string(),
                        ));
                    }
                    let target = match control {
                        Some(ControlWord::FormFieldName) => &mut builder.name,
                        Some(ControlWord::FormFieldFormat) => &mut builder.format,
                        Some(ControlWord::FormFieldDefaultText) => &mut builder.default_text,
                        Some(ControlWord::FormFieldHelpText) => &mut builder.help_text,
                        Some(ControlWord::FormFieldStatusText) => &mut builder.status_text,
                        Some(ControlWord::FormFieldEntryMacro) => &mut builder.entry_macro,
                        Some(ControlWord::FormFieldExitMacro) => &mut builder.exit_macro,
                        Some(ControlWord::FormFieldListEntry) => {
                            if builder.list_entries.len()
                                >= super::super::super::form_field::MAX_FORM_FIELD_LIST_ENTRIES
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF form-field list exceeds 25 entries".to_string(),
                                ));
                            }
                            builder
                                .list_entries
                                .push(self.parse_form_field_text_destination()?);
                            continue;
                        },
                        Some(
                            ControlWord::FormFieldType(_)
                            | ControlWord::FormFieldTextType(_)
                            | ControlWord::FormFieldMaxLength(_)
                            | ControlWord::FormFieldProtected(_)
                            | ControlWord::FormFieldRecalculate(_)
                            | ControlWord::FormFieldAutomaticSize(_)
                            | ControlWord::FormFieldDefaultResult(_)
                            | ControlWord::FormFieldResult(_)
                            | ControlWord::FormFieldHalfPointSize(_)
                            | ControlWord::FormFieldOwnHelp(_)
                            | ControlWord::FormFieldOwnStatus(_)
                            | ControlWord::FormFieldHasListBox(_),
                        ) => {
                            self.pos += 1;
                            depth = depth.checked_add(1).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF formfield nesting depth overflow".to_string(),
                                )
                            })?;
                            continue;
                        },
                        Some(_) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF formfield contains an active or unknown nested destination"
                                    .to_string(),
                            ));
                        },
                        None => {
                            self.pos += 1;
                            depth = depth.checked_add(1).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF formfield nesting depth overflow".to_string(),
                                )
                            })?;
                            continue;
                        },
                    };
                    if target.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield contains a duplicate text destination".to_string(),
                        ));
                    }
                    *target = Some(self.parse_form_field_text_destination()?);
                },
                Some(Token::Control(control)) => {
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    concat!("duplicate RTF formfield ", $name).to_string(),
                                ));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::FormFieldType(value) => set_once!(
                            builder.field_type,
                            super::super::super::form_field::FormFieldType::from_rtf(*value)?,
                            "fftype"
                        ),
                        ControlWord::FormFieldTextType(value) => set_once!(
                            builder.text_type,
                            super::super::super::form_field::FormTextType::from_rtf(*value)?,
                            "fftypetxt"
                        ),
                        ControlWord::FormFieldMaxLength(value) => set_once!(
                            builder.max_length,
                            u16::try_from(*value).map_err(|_err| RtfError::MalformedDocument(
                                "RTF ffmaxlen is outside 0..=65535".to_string()
                            ))?,
                            "ffmaxlen"
                        ),
                        ControlWord::FormFieldProtected(value) => {
                            set_once!(builder.protected, *value, "ffprot");
                        },
                        ControlWord::FormFieldRecalculate(value) => {
                            set_once!(builder.calculate_on_exit, *value, "ffrecalc");
                        },
                        ControlWord::FormFieldAutomaticSize(value) => {
                            set_once!(builder.size_automatically, *value, "ffsize");
                        },
                        ControlWord::FormFieldDefaultResult(value) => {
                            set_once!(builder.default_result, *value, "ffdefres");
                        },
                        ControlWord::FormFieldResult(value) => {
                            set_once!(builder.result, *value, "ffres");
                        },
                        ControlWord::FormFieldHalfPointSize(value) => {
                            set_once!(builder.half_point_size, *value, "ffhps");
                        },
                        ControlWord::FormFieldOwnHelp(value) => {
                            set_once!(builder.own_help, *value, "ffownhelp");
                        },
                        ControlWord::FormFieldOwnStatus(value) => {
                            set_once!(builder.own_status, *value, "ffownstat");
                        },
                        ControlWord::FormFieldHasListBox(value) => {
                            set_once!(builder.has_list_box, *value, "ffhaslistbox");
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF formfield contains an unsupported control".to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    if !self.decode_transport_text(text)?.trim().is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield contains orphan text".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF formfield cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_form_field_text_destination(&mut self) -> RtfResult<String> {
        self.expect_token(&Token::OpenBrace)?;
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        self.pos += 1; // destination control, classified by caller
        let mut text = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(text.trim_end_matches(['\r', '\n']).to_string());
                },
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    text.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_) | Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF form-field text contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > super::super::super::form_field::MAX_FORM_FIELD_STRING_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field string exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn parse_data_field_destination(&mut self) -> RtfResult<Vec<u8>> {
        self.expect_token(&Token::OpenBrace)?;
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::DataField))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF datafield destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut high = None;
        let mut data = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF datafield has an odd hexadecimal digit count".to_string(),
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
                                    "RTF datafield contains a non-hexadecimal character"
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
                        "RTF datafield cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > super::super::super::form_field::MAX_FORM_FIELD_DATA_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF datafield exceeds the safety limit".to_string(),
                ));
            }
        }
    }
}
