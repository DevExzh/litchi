use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_annotation_range_marker(&mut self, is_start: bool) -> RtfResult<()> {
        let value = self.parse_ignorable_text_destination()?;
        let reference = value.trim().parse::<i32>().map_err(|_| {
            RtfError::MalformedDocument(
                "RTF annotation range reference must be a signed integer".to_string(),
            )
        })?;
        if !self.annotation_ranges.contains_key(&reference)
            && self.annotation_ranges.len() >= MAX_ANNOTATIONS
        {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range count exceeds the safety limit".to_string(),
            ));
        }
        if is_start {
            if self.annotation_ranges.contains_key(&reference) {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF annotation range start".to_string(),
                ));
            }
            self.annotation_ranges
                .insert(reference, (self.body_text_len, None));
            self.body_story_events
                .push(ParsedBodyStoryEvent::AnnotationStart(reference));
        } else {
            let range = self.annotation_ranges.get_mut(&reference).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF annotation range end has no matching start".to_string(),
                )
            })?;
            if range.1.is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF annotation range end".to_string(),
                ));
            }
            range.1 = Some(self.body_text_len);
            self.body_story_events
                .push(ParsedBodyStoryEvent::AnnotationEnd(reference));
        }
        Ok(())
    }

    pub(super) fn parse_annotation_destination(&mut self) -> RtfResult<()> {
        if self.annotations.len() >= MAX_ANNOTATIONS {
            return Err(RtfError::MalformedDocument(
                "RTF annotation count exceeds the safety limit".to_string(),
            ));
        }
        if !self.pending_annotation_mark {
            return Err(RtfError::MalformedDocument(
                "RTF annotation destination requires a preceding chatn marker".to_string(),
            ));
        }
        self.pending_annotation_mark = false;
        self.pos += 2; // ignorable marker and annotation destination
        let mut reference = None;
        let mut date = None;
        let mut parent_id = None;
        let mut icon = None;
        let mut time = None;
        let mut text = String::new();
        let mut shapes = Vec::new();
        let mut shape_groups = Vec::new();
        let mut drawing_order = Vec::new();
        let mut story_events = Vec::new();
        let mut depth = 1usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    if self.is_root_drawing_group() {
                        let order_start = drawing_order.len();
                        self.parse_story_drawing_group(
                            text.len(),
                            &mut shapes,
                            &mut shape_groups,
                            &mut drawing_order,
                        )?;
                        let added = drawing_order.get(order_start..).ok_or_else(|| {
                            RtfError::ParserError(
                                "RTF annotation drawing order shrank during parsing".to_string(),
                            )
                        })?;
                        story_events.extend(added.iter().copied().map(crate::StoryEvent::Drawing));
                        continue;
                    }
                    let nested =
                        match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                            (
                                Some(Token::Control(ControlWord::IgnorableDestination)),
                                Some(Token::Control(control)),
                            ) => Some(*control),
                            _ => None,
                        };
                    match nested {
                        Some(ControlWord::AnnotationReference) => {
                            if reference.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation reference".to_string(),
                                ));
                            }
                            let value = self.parse_nested_annotation_value()?;
                            reference = Some(value.trim().parse::<i32>().map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF annotation reference must be a signed integer".to_string(),
                                )
                            })?);
                        },
                        Some(ControlWord::AnnotationDate) => {
                            if date.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation date".to_string(),
                                ));
                            }
                            date = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationParent) => {
                            if parent_id.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation parent".to_string(),
                                ));
                            }
                            parent_id = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationIcon) => {
                            if icon.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation icon".to_string(),
                                ));
                            }
                            icon = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationTime) => {
                            if time.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation time".to_string(),
                                ));
                            }
                            time = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(control) if Self::forbidden_annotation_control(&control) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF annotation body cannot contain active data".to_string(),
                            ));
                        },
                        Some(_) => {
                            self.skip_group()?;
                        },
                        _ => {
                            depth += 1;
                            self.pos += 1;
                        },
                    }
                    continue;
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(value)) => {
                    let skipped = fallback_skip.min(value.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = value.chars().skip(skipped).collect();
                    text.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let code = match self.tokens.get(self.pos) {
                        Some(Token::Control(ControlWord::Unicode(code))) => *code,
                        _ => return Err(parser_classification_error()),
                    };
                    text.push_str(&self.parse_navigation_unicode_sequence(code)?);
                    fallback_skip = 0;
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = (*count).max(0);
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => text.push('\n'),
                Some(Token::Control(ControlWord::Page(param))) => {
                    require_parameterless(*param, "page")?;
                    story_events.push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                        text.len(),
                    )));
                },
                Some(Token::Control(ControlWord::Tab)) => text.push('\t'),
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(control)) if Self::forbidden_annotation_control(control) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation body cannot contain active data".to_string(),
                    ));
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation body cannot contain binary data".to_string(),
                    ));
                },
                _ => {},
            }
            self.pos += 1;
            if text.len() > MAX_ANNOTATION_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation text exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }

        let has_reference = reference.is_some();
        let id = reference.unwrap_or(0);
        if has_reference
            && self
                .annotations
                .iter()
                .any(|annotation| annotation.has_reference && annotation.id == id)
        {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF annotation reference".to_string(),
            ));
        }
        let (position, range_end) = match self.annotation_ranges.remove(&id) {
            Some((start, Some(end))) if start <= end => (start, end),
            Some((_start, None)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation range has no matching end".to_string(),
                ));
            },
            Some(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation range end precedes its start".to_string(),
                ));
            },
            None => (self.body_text_len, self.body_text_len),
        };
        let annotation = super::super::super::annotation::Annotation {
            annotation_type: super::super::super::annotation::AnnotationType::Comment,
            id,
            has_reference,
            author: Cow::Owned(std::mem::take(&mut self.pending_annotation_author)),
            initials: Cow::Owned(std::mem::take(&mut self.pending_annotation_initials)),
            date: date.map(Cow::Owned),
            text: Cow::Owned(text.trim_end_matches(['\r', '\n']).to_string()),
            shapes,
            shape_groups,
            drawing_order,
            story_events,
            position,
            range_end,
            parent_id: parent_id.map(Cow::Owned),
            icon: icon.map(Cow::Owned),
            time: time.map(Cow::Owned),
        };
        self.pending_annotation_author_seen = false;
        self.pending_annotation_initials_seen = false;
        annotation.validate()?;
        let index = self.annotations.len();
        self.annotations.push(annotation);
        if !has_reference {
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::AnnotationStart(index),
            ));
            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                crate::BodyStoryEvent::AnnotationEnd(index),
            ));
        }
        Ok(())
    }

    pub(super) fn parse_nested_annotation_value(&mut self) -> RtfResult<String> {
        self.pos += 3; // opening brace, ignorable marker, destination
        let mut value = String::new();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation metadata cannot contain nested groups".to_string(),
                    ));
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    value.push_str(&self.parse_navigation_unicode_sequence(*code)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = (*count).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation metadata contains active or invalid controls".to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation metadata exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }

    pub(super) fn forbidden_annotation_control(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::Field
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
        )
    }

    pub(super) fn finalize_annotations(&self) -> RtfResult<()> {
        if !self.annotation_ranges.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF document contains an orphan annotation range".to_string(),
            ));
        }
        if self.pending_annotation_author_seen
            || self.pending_annotation_initials_seen
            || self.pending_annotation_mark
        {
            return Err(RtfError::MalformedDocument(
                "RTF document contains orphan annotation metadata".to_string(),
            ));
        }
        Ok(())
    }

    pub(super) fn finalize_body_story_events(&mut self) -> RtfResult<Vec<crate::BodyStoryEvent>> {
        let annotation_indices: HashMap<i32, usize> = self
            .annotations
            .iter()
            .enumerate()
            .filter(|(_, annotation)| annotation.has_reference)
            .map(|(index, annotation)| (annotation.id, index))
            .collect();
        let mut output = Vec::with_capacity(self.body_story_events.len());
        for event in self.body_story_events.drain(..) {
            let resolved = match event {
                ParsedBodyStoryEvent::Resolved(event) => event,
                ParsedBodyStoryEvent::AnnotationStart(reference) => {
                    crate::BodyStoryEvent::AnnotationStart(
                        *annotation_indices.get(&reference).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF annotation start has no resolved comment".to_string(),
                            )
                        })?,
                    )
                },
                ParsedBodyStoryEvent::AnnotationEnd(reference) => {
                    crate::BodyStoryEvent::AnnotationEnd(
                        *annotation_indices.get(&reference).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF annotation end has no resolved comment".to_string(),
                            )
                        })?,
                    )
                },
                ParsedBodyStoryEvent::RevisionStart(id) => crate::BodyStoryEvent::RevisionStart(
                    self.revision_event_indices
                        .get(id)
                        .copied()
                        .flatten()
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF revision event has no tracked text".to_string(),
                            )
                        })?,
                ),
                ParsedBodyStoryEvent::RevisionEnd(id) => crate::BodyStoryEvent::RevisionEnd(
                    self.revision_event_indices
                        .get(id)
                        .copied()
                        .flatten()
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF revision event has no tracked text".to_string(),
                            )
                        })?,
                ),
                ParsedBodyStoryEvent::RevisionDeletion(id) => {
                    crate::BodyStoryEvent::RevisionDeletion(
                        self.revision_event_indices
                            .get(id)
                            .copied()
                            .flatten()
                            .ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF deletion event has no tracked text".to_string(),
                                )
                            })?,
                    )
                },
            };
            output.push(resolved);
        }
        Ok(output)
    }
}
