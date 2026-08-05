use super::*;

impl<'a> Parser<'a> {
    /// Parse the unique starred root document-background destination.
    pub(super) fn parse_background_destination(
        &mut self,
    ) -> RtfResult<super::super::super::shape::Shape<'a>> {
        self.pos += 2; // consume \* and \background
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF background destination contains text outside its shape".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF background destination must contain exactly one shape group".to_string(),
            ));
        }
        self.pos += 1;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::Shape(None)))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF background destination must contain an unstarred shp destination".to_string(),
            ));
        }
        let mut shape = self.parse_shape_destination(true)?;
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF background destination contains trailing text".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF background destination must contain exactly one shape group".to_string(),
            ));
        }
        self.pos += 1;
        if shape
            .properties
            .iter()
            .any(|property| property.name.is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF background shape property names cannot be empty".to_string(),
            ));
        }
        shape.is_background = true;
        Ok(shape)
    }

    /// Parse a `shp` destination and its nested shape-property groups.
    pub(super) fn parse_shape_destination(
        &mut self,
        allow_shape_result: bool,
    ) -> RtfResult<super::super::super::shape::Shape<'a>> {
        use super::super::super::shape::{Shape, ShapeType};

        let mut shape = Shape::new(ShapeType::Unknown);
        shape.instruction_present = false;
        let mut depth = 0usize;
        let mut shape_instance_depth = None;
        let mut saw_shape_instance = false;
        let mut saw_property = false;
        let mut saw_shape_info = false;
        let mut right = None;
        let mut bottom = None;
        let mut closed = false;
        let mut saw_shape_result = false;
        let mut saw_shape_text = false;
        let parameter = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::Shape(parameter))) => *parameter,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF shape parser is not at shp".to_string(),
                ));
            },
        };
        require_parameterless(parameter, "shp")?;
        self.pos += 1;

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace
                    if self
                        .nested_control_word()
                        .is_some_and(|control| matches!(control, ControlWord::ShapeText(_))) =>
                {
                    let valid_parent = shape_instance_depth == Some(depth);
                    if !valid_parent || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shptxt must be one final destination after shape properties"
                                .to_string(),
                        ));
                    }
                    let (
                        text,
                        text_shapes,
                        text_shape_groups,
                        text_drawing_order,
                        text_story_events,
                        text_background_color,
                    ) = self.parse_shape_text_destination()?;
                    shape.text = Cow::Owned(text);
                    shape.text_shapes = text_shapes;
                    shape.text_shape_groups = text_shape_groups;
                    shape.text_drawing_order = text_drawing_order;
                    shape.text_story_events = text_story_events;
                    shape.text_destination_present = true;
                    shape.text_formatting = self.current_state().ok().map(|state| state.formatting);
                    if let (Some(formatting), Some(background_color)) =
                        (&mut shape.text_formatting, text_background_color)
                    {
                        formatting.background_color = Some(background_color);
                    }
                    saw_shape_text = true;
                },
                Token::OpenBrace
                    if self
                        .nested_control_word()
                        .is_some_and(|control| matches!(control, ControlWord::ShapeResult(_))) =>
                {
                    if depth != 0 || !allow_shape_result || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shprslt must be a single direct child of a root shape".to_string(),
                        ));
                    }
                    saw_shape_result = true;
                    shape.result = self.parse_shape_result()?;
                },
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ShapeProperty) =>
                {
                    if shape_instance_depth != Some(depth) || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape properties must be direct shpinst children before shptxt"
                                .to_string(),
                        ));
                    }
                    saw_property = true;
                    if shape.properties.len() >= MAX_SHAPE_PROPERTIES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape property count exceeds the safety limit".to_string(),
                        ));
                    }
                    let (name, value, binary_value, theme_value, hyperlink) =
                        self.parse_shape_property_group()?;
                    shape
                        .properties
                        .push(super::super::super::shape::ShapeProperty {
                            name: Cow::Owned(name),
                            value: Cow::Owned(value),
                            binary_value: binary_value.map(Cow::Owned),
                            theme_value,
                            hyperlink,
                        });
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapeInstance),
                        ])
                    ) =>
                {
                    if depth != 0
                        || shape_instance_depth.is_some()
                        || saw_shape_instance
                        || saw_shape_text
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF shpinst must be a single direct shape child before shptxt"
                                .to_string(),
                        ));
                    }
                    shape_instance_depth = Some(depth + 1);
                    saw_shape_instance = true;
                    shape.instruction_present = true;
                    depth += 1;
                    self.pos += 3;
                },
                Token::OpenBrace => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape contains a misplaced or unknown group".to_string(),
                    ));
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Token::CloseBrace => {
                    if shape_instance_depth == Some(depth) {
                        shape_instance_depth = None;
                    }
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeType(value)) => {
                    let _ = value;
                    return Err(RtfError::MalformedDocument(
                        "RTF shpinst destination must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeLeft(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.x = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeTop(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.y = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRight(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    right = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBottom(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    bottom = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWidth(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeHeight(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRotation(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.rotation = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeZOrder(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.geometry.z_order = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWrap(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.info.push(crate::ShapeGroupInfo::Wrap(*value));
                    shape.wrap_mode = match value {
                        1 => super::super::super::shape::WrapMode::None,
                        2 => super::super::super::shape::WrapMode::Square,
                        4 => super::super::super::shape::WrapMode::Tight,
                        3 | 5 => super::super::super::shape::WrapMode::Through,
                        _ => shape.wrap_mode,
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBelowText(value)) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.info.push(crate::ShapeGroupInfo::BelowText(*value));
                    shape.behind_doc = *value;
                    if *value {
                        shape.wrap_mode = super::super::super::shape::WrapMode::Behind;
                    }
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeLockAnchor) => {
                    saw_shape_info = true;
                    if saw_property || saw_shape_text || saw_shape_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape info must precede properties".to_string(),
                        ));
                    }
                    shape.info.push(crate::ShapeGroupInfo::LockAnchor);
                    shape.locked = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::PictureShapeId(Some(value)))
                    if !saw_property && !saw_shape_text && !saw_shape_result =>
                {
                    saw_shape_info = true;
                    shape.info.push(crate::ShapeGroupInfo::ShapeId(*value));
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unknown(name, parameter))
                    if !saw_property && !saw_shape_text && !saw_shape_result =>
                {
                    let info = match (*name, *parameter) {
                        ("shpfhdr", Some(value)) => crate::ShapeGroupInfo::InHeader(value != 0),
                        ("shpbxpage", None) => crate::ShapeGroupInfo::HorizontalPage,
                        ("shpbxmargin", None) => crate::ShapeGroupInfo::HorizontalMargin,
                        ("shpbxcolumn", None) => crate::ShapeGroupInfo::HorizontalColumn,
                        ("shpbxignore", None) => crate::ShapeGroupInfo::IgnoreHorizontal,
                        ("shpbypage", None) => crate::ShapeGroupInfo::VerticalPage,
                        ("shpbymargin", None) => crate::ShapeGroupInfo::VerticalMargin,
                        ("shpbypara", None) => crate::ShapeGroupInfo::VerticalParagraph,
                        ("shpbyignore", None) => crate::ShapeGroupInfo::IgnoreVertical,
                        ("shpwrk", Some(value)) => crate::ShapeGroupInfo::WrapSide(value),
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF shape contains an unknown or malformed shape-info control"
                                    .to_string(),
                            ));
                        },
                    };
                    if shape.info.len() >= 32 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-info count exceeds the safety limit".to_string(),
                        ));
                    }
                    shape.info.push(info);
                    saw_shape_info = true;
                    self.pos += 1;
                },
                Token::Text(value) if !value.bytes().all(|byte| byte.is_ascii_whitespace()) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape destination contains direct text".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeResult(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shprslt destination must be grouped".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeText(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shptxt destination must be grouped and unstarred".to_string(),
                    ));
                },
                Token::Control(_) if saw_shape_text => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape controls must precede shptxt".to_string(),
                    ));
                },
                Token::Text(_) => self.pos += 1,
                Token::Binary(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape cannot contain direct binary data".to_string(),
                    ));
                },
                Token::Control(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape contains a misplaced control".to_string(),
                    ));
                },
            }
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        let fallback_only = allow_shape_result
            && shape.result.is_some()
            && !saw_shape_info
            && !saw_property
            && !saw_shape_text;
        if shape_instance_depth.is_some() || (!saw_shape_instance && !fallback_only) {
            return Err(RtfError::MalformedDocument(
                "RTF shape must contain exactly one starred shpinst".to_string(),
            ));
        }
        shape.validate()?;
        Self::apply_shape_properties(&mut shape);
        if let Some(right) = right {
            shape.geometry.width = right.saturating_sub(shape.geometry.x);
        }
        if let Some(bottom) = bottom {
            shape.geometry.height = bottom.saturating_sub(shape.geometry.y);
        }
        Ok(shape)
    }

    #[allow(clippy::type_complexity)]
    pub(super) fn parse_shape_text_destination(
        &mut self,
    ) -> RtfResult<(
        String,
        Vec<super::super::super::shape::Shape<'a>>,
        Vec<super::super::super::shape::ShapeGroup<'a>>,
        Vec<crate::StoryDrawing>,
        Vec<crate::StoryEvent>,
        Option<ColorRef>,
    )> {
        if let Some(
            [
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::ShapeText(_)),
            ],
        ) = self.tokens.get(self.pos..self.pos + 3)
        {
            return Err(RtfError::MalformedDocument(
                "RTF shptxt destination must not be starred".to_string(),
            ));
        }
        match self.tokens.get(self.pos..self.pos + 2) {
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::ShapeText(None)),
                ],
            ) => {
                self.pos += 2;
            },
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::ShapeText(Some(_))),
                ],
            ) => {
                return Err(RtfError::MalformedDocument(
                    "RTF shptxt destination must not have a parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF shptxt destination".to_string(),
                ));
            },
        }

        let mut text = String::new();
        let mut shapes = Vec::new();
        let mut shape_groups = Vec::new();
        let mut drawing_order = Vec::new();
        let mut story_events = Vec::new();
        let mut depth = 0usize;
        let mut background_stack = vec![self.current_state()?.formatting.background_color];
        let mut observed_background: Option<Option<ColorRef>> = None;
        macro_rules! observe_background {
            ($visible:expr) => {
                if $visible {
                    let current = *background_stack.last().ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF shape-text background state is missing".to_string(),
                        )
                    })?;
                    if let Some(observed) = observed_background {
                        if observed != current {
                            return Err(RtfError::MalformedDocument(
                                "RTF shape text has multiple visible background colors".to_string(),
                            ));
                        }
                    } else {
                        observed_background = Some(current);
                    }
                }
            };
        }
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    return Ok((
                        text,
                        shapes,
                        shape_groups,
                        drawing_order,
                        story_events,
                        observed_background.flatten(),
                    ));
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    background_stack.pop();
                    self.pos += 1;
                },
                Some(Token::OpenBrace) if self.is_root_drawing_group() => {
                    let order_start = drawing_order.len();
                    self.parse_story_drawing_group(
                        text.len(),
                        &mut shapes,
                        &mut shape_groups,
                        &mut drawing_order,
                    )?;
                    let added = drawing_order.get(order_start..).ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF shape-text drawing order shrank during parsing".to_string(),
                        )
                    })?;
                    story_events.extend(added.iter().copied().map(crate::StoryEvent::Drawing));
                },
                Some(Token::OpenBrace) if self.is_custom_xml_markup_group() => {
                    return Err(RtfError::MalformedDocument(
                        "RTF custom XML markup destinations are supported only in the main body story"
                            .to_string(),
                    ));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || self.nested_control_word().is_some_and(|control| {
                        matches!(
                            control,
                            ControlWord::Picture
                                | ControlWord::Object
                                | ControlWord::InvalidObjectDestinationParameter
                                | ControlWord::Shape(_)
                                | ControlWord::ShapeGroup(_)
                                | ControlWord::ShapeResult(_)
                                | ControlWord::LegacyDrawingObject
                        )
                    }) =>
                {
                    self.skip_group()?;
                },
                Some(Token::OpenBrace) => {
                    depth += 1;
                    if depth > 64 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shptxt nesting exceeds the safety limit".to_string(),
                        ));
                    }
                    let inherited = *background_stack.last().ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF shape-text background state is missing".to_string(),
                        )
                    })?;
                    background_stack.push(inherited);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::ShapeText(_))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shptxt cannot contain another shptxt destination".to_string(),
                    ));
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    let decoded = self.parse_destination_unicode_sequence(*code)?;
                    observe_background!(
                        decoded.chars().any(|character| !character.is_whitespace())
                    );
                    text.push_str(&decoded);
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                    text.push('\n');
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Page(param))) => {
                    require_parameterless(*param, "page")?;
                    story_events.push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                        text.len(),
                    )));
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    let decoded = control_symbol_text(control).unwrap_or_default();
                    observe_background!(
                        decoded.chars().any(|character| !character.is_whitespace())
                    );
                    text.push_str(decoded);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::ColorBackground(value))) => {
                    *background_stack.last_mut().ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF shape-text background state is missing".to_string(),
                        )
                    })? = Some(Self::required_character_value(*value, "cb", u16::MAX)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Plain)) => {
                    *background_stack.last_mut().ok_or_else(|| {
                        RtfError::ParserError(
                            "RTF shape-text background state is missing".to_string(),
                        )
                    })? = None;
                    self.pos += 1;
                },
                Some(Token::Text(value)) => {
                    let decoded = self.decode_transport_text(value)?;
                    observe_background!(
                        decoded.chars().any(|character| !character.is_whitespace())
                    );
                    text.push_str(&decoded);
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shptxt contains direct binary data".to_string(),
                    ));
                },
                Some(Token::Control(_)) => self.pos += 1,
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > MAX_SHAPE_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF shape text exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn is_root_drawing_group(&self) -> bool {
        matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::Shape(_) | ControlWord::ShapeGroup(_)),
            ])
        ) || matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::Shape(_) | ControlWord::ShapeGroup(_)),
            ])
        )
    }

    pub(super) fn parse_story_drawing_group(
        &mut self,
        position: usize,
        shapes: &mut Vec<super::super::super::shape::Shape<'a>>,
        shape_groups: &mut Vec<super::super::super::shape::ShapeGroup<'a>>,
        drawing_order: &mut Vec<crate::StoryDrawing>,
    ) -> RtfResult<()> {
        let shape_start = self.shapes.len();
        let group_start = self.shape_groups.len();
        let order_start = self.drawing_order.len();
        let body_event_start = self.body_story_events.len();
        self.parse_group()?;
        self.body_story_events.truncate(body_event_start);
        let added = self.shapes.len().saturating_sub(shape_start)
            + self.shape_groups.len().saturating_sub(group_start);
        if added != 1 {
            return Err(RtfError::MalformedDocument(
                "RTF story drawing group must contain exactly one root shp or shpgrp destination"
                    .to_string(),
            ));
        }
        for mut shape in self.shapes.drain(shape_start..) {
            shape.position = position;
            shapes.push(shape);
        }
        for mut group in self.shape_groups.drain(group_start..) {
            group.position = position;
            shape_groups.push(group);
        }
        for drawing in self.drawing_order.drain(order_start..) {
            match drawing {
                crate::StoryDrawing::Shape(index) => {
                    drawing_order.push(crate::StoryDrawing::Shape(index - shape_start));
                },
                crate::StoryDrawing::ShapeGroup(index) => {
                    drawing_order.push(crate::StoryDrawing::ShapeGroup(index - group_start));
                },
            }
        }
        if shapes.len() > MAX_SHAPES || shape_groups.len() > MAX_SHAPE_GROUPS {
            return Err(RtfError::MalformedDocument(
                "RTF story drawing count exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(super) fn parse_shape_result(&mut self) -> RtfResult<Option<crate::ShapeResult<'a>>> {
        let control_index = match self.tokens.get(self.pos..self.pos + 3) {
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeResult(None)),
                ],
            ) => self.pos + 2,
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeResult(Some(_))),
                ],
            ) => {
                return Err(RtfError::MalformedDocument(
                    "RTF shprslt destination must not have a parameter".to_string(),
                ));
            },
            _ => match self.tokens.get(self.pos..self.pos + 2) {
                Some(
                    [
                        Token::OpenBrace,
                        Token::Control(ControlWord::ShapeResult(None)),
                    ],
                ) => self.pos + 1,
                Some(
                    [
                        Token::OpenBrace,
                        Token::Control(ControlWord::ShapeResult(Some(_))),
                    ],
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shprslt destination must not have a parameter".to_string(),
                    ));
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "invalid RTF shprslt destination".to_string(),
                    ));
                },
            },
        };
        self.pos = control_index + 1;
        self.skip_legacy_whitespace();
        if matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::LegacyDrawingObject),
            ])
        ) {
            self.pos += 1;
            let mut drawing = self.parse_legacy_drawing_at(0)?.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF shprslt legacy drawing must contain one primitive".to_string(),
                )
            })?;
            drawing.position = 0;
            self.skip_legacy_whitespace();
            if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
                return Err(RtfError::MalformedDocument(
                    "RTF shprslt must contain exactly one legacy drawing".to_string(),
                ));
            }
            self.pos += 1;
            let result = crate::ShapeResult { drawing };
            result.validate()?;
            return Ok(Some(result));
        }
        if matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::LegacyDrawingObject),
            ])
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF shprslt legacy drawing destination must be starred".to_string(),
            ));
        }

        let mut depth = 0usize;
        let mut retained_bytes = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF shprslt nesting depth overflow".to_string(),
                        )
                    })?;
                    if depth > 64 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shprslt nesting exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    return Ok(None);
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::Text(value)) => {
                    retained_bytes = retained_bytes.saturating_add(value.len());
                    if retained_bytes > 16 * 1_048_576 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shprslt content exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::Binary(value)) => {
                    retained_bytes = retained_bytes.saturating_add(value.len());
                    if retained_bytes > 16 * 1_048_576 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shprslt content exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::Control(_)) => self.pos += 1,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    #[allow(clippy::type_complexity)]
    pub(super) fn parse_shape_property_group(
        &mut self,
    ) -> RtfResult<(
        String,
        String,
        Option<Vec<u8>>,
        Option<crate::ShapeThemeValue>,
        Option<crate::ShapeHyperlink<'a>>,
    )> {
        #[derive(Clone, Copy, PartialEq, Eq)]
        enum PropertyPart {
            Name,
            Value,
        }

        let mut name = String::new();
        let mut value = String::new();
        let mut binary_value = None;
        let mut theme_value = None;
        let mut hyperlink = None;
        let mut seen_name = false;
        let mut seen_value = false;
        let mut part = None;
        let mut part_depth = None;
        let mut depth = 0usize;
        self.pos += 1; // consume the opening brace
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ShapeHyperlink))
                    ) =>
                {
                    if depth != 0 || part.is_some() || !seen_name || hyperlink.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF hl must be the single direct hyperlink group after sn in sp"
                                .to_string(),
                        ));
                    }
                    hyperlink = Some(self.parse_shape_hyperlink_destination()?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapeHyperlink),
                        ])
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF hl destination must not be starred".to_string(),
                    ));
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::ShapeHyperlinkLocation
                                | ControlWord::ShapeHyperlinkSource
                                | ControlWord::ShapeHyperlinkFriendlyName
                        ))
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape-hyperlink string destinations may occur only inside hl"
                            .to_string(),
                    ));
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) && matches!(
                        self.tokens.get(self.pos + 2),
                        Some(Token::Control(ControlWord::ShapeBinaryValue(_)))
                    ) =>
                {
                    if part != Some(PropertyPart::Value)
                        || binary_value.is_some()
                        || !value.trim().is_empty()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF svb must be the only value inside one sv destination".to_string(),
                        ));
                    }
                    binary_value = Some(self.parse_shape_binary_value()?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) && matches!(
                        self.tokens.get(self.pos + 2),
                        Some(Token::Control(ControlWord::ShapeThemeValue(_)))
                    ) =>
                {
                    if depth != 0
                        || part.is_some()
                        || !seen_value
                        || theme_value.is_some()
                        || binary_value.is_some()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF hsv must be the single direct child after scalar sv in sp"
                                .to_string(),
                        ));
                    }
                    theme_value = Some(self.parse_shape_theme_value()?);
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ShapeBinaryValue(_)))
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF svb destination must be starred".to_string(),
                    ));
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ShapeThemeValue(_)))
                    ) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF hsv destination must be starred".to_string(),
                    ));
                },
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    if !seen_name || !seen_value {
                        return Err(RtfError::MalformedDocument(
                            "RTF sp requires exactly one sn followed by one sv".to_string(),
                        ));
                    }
                    if binary_value.is_some() && !value.trim().is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape property cannot mix scalar and binary values".to_string(),
                        ));
                    }
                    return Ok((
                        name.trim().to_string(),
                        value.trim().to_string(),
                        binary_value,
                        theme_value,
                        hyperlink,
                    ));
                },
                Token::CloseBrace => {
                    if part_depth == Some(depth) {
                        part = None;
                        part_depth = None;
                    }
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapePropertyName) => {
                    if depth != 1 || seen_name || seen_value || theme_value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF sp requires exactly one direct sn before sv".to_string(),
                        ));
                    }
                    seen_name = true;
                    part = Some(PropertyPart::Name);
                    part_depth = Some(depth);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapePropertyValue) => {
                    if depth != 1 || !seen_name || seen_value || theme_value.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF sp requires exactly one direct sv after sn".to_string(),
                        ));
                    }
                    seen_value = true;
                    part = Some(PropertyPart::Value);
                    part_depth = Some(depth);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBinaryValue(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF svb destination must be grouped and starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::ShapeHyperlink
                    | ControlWord::ShapeHyperlinkLocation
                    | ControlWord::ShapeHyperlinkSource
                    | ControlWord::ShapeHyperlinkFriendlyName,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape-hyperlink destinations must be grouped".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeThemeValue(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF hsv destination must be a direct starred group after sv".to_string(),
                    ));
                },
                Token::Control(ControlWord::Unicode(code)) if part.is_some() => {
                    let decoded = self.parse_destination_unicode_sequence(*code)?;
                    match part {
                        Some(PropertyPart::Name) => name.push_str(&decoded),
                        Some(PropertyPart::Value) => value.push_str(&decoded),
                        None => {},
                    }
                },
                Token::Control(control)
                    if part.is_some() && control_symbol_text(control).is_some() =>
                {
                    let decoded = control_symbol_text(control).unwrap_or_default();
                    match part {
                        Some(PropertyPart::Name) => name.push_str(decoded),
                        Some(PropertyPart::Value) => value.push_str(decoded),
                        None => {},
                    }
                    self.pos += 1;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text(text)?;
                    match part {
                        Some(PropertyPart::Name) => name.push_str(&decoded),
                        Some(PropertyPart::Value) => value.push_str(&decoded),
                        None => {},
                    }
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
            if name.len().saturating_add(value.len()) > MAX_SHAPE_PROPERTY_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF shape property exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_shape_theme_value(&mut self) -> RtfResult<crate::ShapeThemeValue> {
        match self.tokens.get(self.pos..self.pos + 3) {
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeThemeValue(None)),
                ],
            ) => self.pos += 3,
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeThemeValue(Some(_))),
                ],
            ) => {
                return Err(RtfError::MalformedDocument(
                    "RTF hsv destination must not have a parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF hsv destination".to_string(),
                ));
            },
        }
        let mut color = None;
        let mut tint = None;
        let mut shade = None;
        let mut whitespace_bytes = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let theme = crate::ShapeThemeValue {
                        color: color.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF hsv requires exactly one accent selector".to_string(),
                            )
                        })?,
                        tint: tint.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF hsv requires exactly one ctint value".to_string(),
                            )
                        })?,
                        shade: shade.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF hsv requires exactly one cshade value".to_string(),
                            )
                        })?,
                    };
                    theme.validate()?;
                    return Ok(theme);
                },
                Some(Token::Control(ControlWord::ShapeThemeColor(value, None)))
                    if color.is_none() =>
                {
                    color = Some(*value);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::ShapeThemeTint(Some(value))))
                    if tint.is_none() =>
                {
                    tint = Some(u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF hsv ctint must be between 0 and 255".to_string(),
                        )
                    })?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::ShapeThemeShade(Some(value))))
                    if shade.is_none() =>
                {
                    shade = Some(u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF hsv cshade must be between 0 and 255".to_string(),
                        )
                    })?);
                    self.pos += 1;
                },
                Some(Token::Text(value))
                    if value.bytes().all(|byte| byte.is_ascii_whitespace()) =>
                {
                    whitespace_bytes = whitespace_bytes.saturating_add(value.len());
                    if whitespace_bytes > 4_096 {
                        return Err(RtfError::MalformedDocument(
                            "RTF hsv whitespace exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF hsv contains duplicate, parameterless, nested, or active content"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_shape_binary_value(&mut self) -> RtfResult<Vec<u8>> {
        match self.tokens.get(self.pos..self.pos + 3) {
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeBinaryValue(None)),
                ],
            ) => self.pos += 3,
            Some(
                [
                    Token::OpenBrace,
                    Token::Control(ControlWord::IgnorableDestination),
                    Token::Control(ControlWord::ShapeBinaryValue(Some(_))),
                ],
            ) => {
                return Err(RtfError::MalformedDocument(
                    "RTF svb destination must not have a parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF svb destination".to_string(),
                ));
            },
        }
        let mut bytes = Vec::new();
        let mut high_nibble = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high_nibble.is_some() || bytes.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF svb payload is empty or has an odd hexadecimal length".to_string(),
                        ));
                    }
                    return Ok(bytes);
                },
                Some(Token::Text(text)) => {
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF svb contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            bytes.push((high << 4) | nibble);
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::Binary(value)) => {
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF svb binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    bytes.extend_from_slice(value);
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF svb contains nested or active content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if bytes.len() > crate::MAX_SHAPE_PROPERTY_BINARY_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF svb payload exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn apply_shape_property(
        shape: &mut super::super::super::shape::Shape<'a>,
        name: &str,
        value: &str,
    ) {
        match name {
            "shapeType" => {
                if let Ok(value) = value.parse() {
                    shape.shape_type = Self::shape_type_from_rtf(value);
                }
            },
            "wzName" => shape.name = Cow::Owned(value.to_string()),
            "fBehindDocument" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.behind_doc = value;
                }
            },
            "fBackground" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.is_background = value;
                }
            },
            "fLockPosition" | "fLockAgainstGrouping" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.locked |= value;
                }
            },
            "fillType" => {
                if let Ok(value) = value.parse::<i32>() {
                    shape.fill.fill_type = match value {
                        0 => super::super::super::shape::FillType::Solid,
                        1 => super::super::super::shape::FillType::Pattern,
                        2 => super::super::super::shape::FillType::Texture,
                        3 => super::super::super::shape::FillType::Picture,
                        4..=8 => super::super::super::shape::FillType::Gradient,
                        9 => super::super::super::shape::FillType::Background,
                        _ => shape.fill.fill_type,
                    };
                }
            },
            "fillColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.color = super::super::super::shape::OfficeArtColor(value);
                }
            },
            "fillBackColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.color2 = Some(super::super::super::shape::OfficeArtColor(value));
                }
            },
            "fillOpacity" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.opacity = super::super::super::shape::OfficeArtOpacity(value);
                }
            },
            "lineColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.line.color = super::super::super::shape::OfficeArtColor(value);
                }
            },
            "lineWidth" => {
                if let Ok(value) = value.parse() {
                    shape.line.width_emu = value;
                }
            },
            "rotation" => {
                if let Ok(value) = value.parse::<i32>() {
                    shape.geometry.rotation = value / 65_536;
                }
            },
            _ => {},
        }
    }

    pub(super) fn apply_shape_properties(shape: &mut super::super::super::shape::Shape<'a>) {
        for index in 0..shape.properties.len() {
            let Some(property) = shape.properties.get(index) else {
                break;
            };
            let name = property.name.to_string();
            let value = property.value.to_string();
            Self::apply_shape_property(shape, &name, &value);
        }

        if let Some(value) = shape
            .properties
            .iter()
            .rev()
            .find(|property| property.name == "fFilled")
            .and_then(|property| Self::parse_shape_bool(&property.value))
        {
            if value {
                if shape.fill.fill_type == super::super::super::shape::FillType::None {
                    shape.fill.fill_type = super::super::super::shape::FillType::Solid;
                }
            } else {
                shape.fill.fill_type = super::super::super::shape::FillType::None;
            }
        }

        if let Some(value) = shape
            .properties
            .iter()
            .rev()
            .find(|property| property.name == "fLine")
            .and_then(|property| Self::parse_shape_bool(&property.value))
        {
            shape.line.visible = value;
        }
    }

    pub(super) fn parse_shape_bool(value: &str) -> Option<bool> {
        value.trim().parse::<i32>().ok().map(|value| value != 0)
    }

    pub(super) fn parse_office_art_u32(value: &str) -> Option<u32> {
        let value = value.trim();
        value
            .parse::<u32>()
            .ok()
            .or_else(|| value.parse::<i32>().ok().map(|value| value as u32))
    }

    pub(super) fn parse_shape_group_destination(
        &mut self,
    ) -> RtfResult<super::super::super::shape::ShapeGroup<'a>> {
        self.parse_shape_group_destination_at_depth(0, true)
    }

    pub(super) fn parse_shape_group_destination_at_depth(
        &mut self,
        nesting_depth: usize,
        root: bool,
    ) -> RtfResult<super::super::super::shape::ShapeGroup<'a>> {
        if nesting_depth >= MAX_SHAPE_GROUP_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF shape group nesting exceeds the safety limit".to_string(),
            ));
        }
        let mut group = super::super::super::shape::ShapeGroup::new();
        let mut right = None;
        let mut bottom = None;
        let mut in_instance = false;
        let mut saw_instance = false;
        let mut saw_child = false;
        let mut saw_result = false;
        let mut closed = false;
        let parameter = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::ShapeGroup(parameter))) => *parameter,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF shape group parser is not at shpgrp".to_string(),
                ));
            },
        };
        require_parameterless(parameter, "shpgrp")?;
        self.pos += 1;

        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ShapeInstance),
                        ])
                    ) =>
                {
                    if in_instance || saw_instance || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group must contain exactly one shpinst".to_string(),
                        ));
                    }
                    saw_instance = true;
                    in_instance = true;
                    self.pos += 3;
                },
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ShapeProperty) =>
                {
                    if !in_instance || saw_child {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group properties must precede direct children inside shpinst"
                                .to_string(),
                        ));
                    }
                    if group.properties.len() >= MAX_SHAPE_PROPERTIES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group property count exceeds the safety limit".to_string(),
                        ));
                    }
                    let (name, value, binary_value, theme_value, hyperlink) =
                        self.parse_shape_property_group()?;
                    if name == "wzName" {
                        group.name = Cow::Owned(value.clone());
                    }
                    group
                        .properties
                        .push(super::super::super::shape::ShapeProperty {
                            name: Cow::Owned(name),
                            value: Cow::Owned(value),
                            binary_value: binary_value.map(Cow::Owned),
                            theme_value,
                            hyperlink,
                        });
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::Shape(_)))
                    ) =>
                {
                    if !in_instance || group.child_order.len() >= MAX_SHAPES_PER_GROUP {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group child is misplaced or exceeds the safety limit"
                                .to_string(),
                        ));
                    }
                    saw_child = true;
                    self.pos += 1;
                    group.add_shape(self.parse_shape_destination(false)?)?;
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ShapeGroup(_)))
                    ) =>
                {
                    if !in_instance
                        || group.groups.len() >= MAX_GROUPS_PER_GROUP
                        || group.child_order.len() >= MAX_SHAPES_PER_GROUP
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF nested shape group is misplaced or exceeds the safety limit"
                                .to_string(),
                        ));
                    }
                    saw_child = true;
                    self.pos += 1;
                    let nested = self.parse_shape_group_destination_at_depth(
                        nesting_depth.saturating_add(1),
                        false,
                    )?;
                    group.add_group(nested)?;
                },
                Token::OpenBrace
                    if self
                        .nested_control_word()
                        .is_some_and(|control| matches!(control, ControlWord::ShapeResult(_))) =>
                {
                    if in_instance || !root || !saw_instance || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shprslt may occur once after root shape-group shpinst".to_string(),
                        ));
                    }
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group shprslt destination must be starred".to_string(),
                        ));
                    }
                    group.result = self.parse_shape_result()?;
                    saw_result = true;
                },
                Token::OpenBrace => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape group contains a misplaced or unknown group".to_string(),
                    ));
                },
                Token::CloseBrace if in_instance => {
                    in_instance = false;
                    self.pos += 1;
                },
                Token::CloseBrace => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Token::Control(ControlWord::ShapeLeft(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.x = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeTop(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.y = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRight(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    right = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBottom(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    bottom = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWidth(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeHeight(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRotation(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.rotation = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeZOrder(value)) => {
                    if saw_child || saw_result {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group geometry must precede children".to_string(),
                        ));
                    }
                    group.geometry.z_order = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::PictureShapeId(Some(value)))
                    if !saw_child && !saw_result =>
                {
                    group.info.push(crate::ShapeGroupInfo::ShapeId(*value));
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWrap(value)) if !saw_child && !saw_result => {
                    group.info.push(crate::ShapeGroupInfo::Wrap(*value));
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBelowText(value)) if !saw_child && !saw_result => {
                    group.info.push(crate::ShapeGroupInfo::BelowText(*value));
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeLockAnchor) if !saw_child && !saw_result => {
                    group.info.push(crate::ShapeGroupInfo::LockAnchor);
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unknown(name, parameter))
                    if !saw_child && !saw_result =>
                {
                    let info = match (*name, *parameter) {
                        ("shpfhdr", Some(value)) => crate::ShapeGroupInfo::InHeader(value != 0),
                        ("shpbxpage", None) => crate::ShapeGroupInfo::HorizontalPage,
                        ("shpbxmargin", None) => crate::ShapeGroupInfo::HorizontalMargin,
                        ("shpbxcolumn", None) => crate::ShapeGroupInfo::HorizontalColumn,
                        ("shpbxignore", None) => crate::ShapeGroupInfo::IgnoreHorizontal,
                        ("shpbypage", None) => crate::ShapeGroupInfo::VerticalPage,
                        ("shpbymargin", None) => crate::ShapeGroupInfo::VerticalMargin,
                        ("shpbypara", None) => crate::ShapeGroupInfo::VerticalParagraph,
                        ("shpbyignore", None) => crate::ShapeGroupInfo::IgnoreVertical,
                        ("shpwrk", Some(value)) => crate::ShapeGroupInfo::WrapSide(value),
                        _ => return Err(RtfError::MalformedDocument(
                            "RTF shape group contains an unknown or malformed shape-info control"
                                .to_string(),
                        )),
                    };
                    if group.info.len() >= 32 {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-group info count exceeds the safety limit".to_string(),
                        ));
                    }
                    group.info.push(info);
                    self.pos += 1;
                },
                Token::Text(text) if text.chars().all(char::is_whitespace) => self.pos += 1,
                Token::Binary(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape group cannot contain direct binary data".to_string(),
                    ));
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shape group contains misplaced content".to_string(),
                    ));
                },
            }
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if in_instance || !saw_instance {
            return Err(RtfError::MalformedDocument(
                "RTF shape group must contain exactly one starred shpinst".to_string(),
            ));
        }
        if let Some(right) = right {
            group.geometry.width = right.saturating_sub(group.geometry.x);
        }
        if let Some(bottom) = bottom {
            group.geometry.height = bottom.saturating_sub(group.geometry.y);
        }
        if root {
            group.validate()?;
        }
        Ok(group)
    }

    pub(super) fn nested_control_word(&self) -> Option<ControlWord<'a>> {
        let mut index = self.pos.checked_add(1)?;
        if matches!(
            self.tokens.get(index),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            index += 1;
        }
        match self.tokens.get(index) {
            Some(Token::Control(control)) => Some(*control),
            _ => None,
        }
    }

    pub(super) fn shape_type_from_rtf(value: i32) -> super::super::super::shape::ShapeType {
        use super::super::super::shape::ShapeType;
        match value {
            1 => ShapeType::Rectangle,
            2 => ShapeType::RoundRectangle,
            3 => ShapeType::Ellipse,
            19 => ShapeType::Arc,
            20 => ShapeType::Line,
            75 => ShapeType::PictureFrame,
            202 => ShapeType::TextBox,
            0 => ShapeType::Group,
            value => ShapeType::Custom(value),
        }
    }
}
