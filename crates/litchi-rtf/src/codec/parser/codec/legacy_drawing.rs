use super::*;

impl<'a> Parser<'a> {
    pub(super) fn apply_legacy_text_box_control(
        builder: &mut LegacyTextBoxBuilder,
        control: &ControlWord,
    ) -> RtfResult<bool> {
        macro_rules! set_once {
            ($slot:expr, $value:expr, $name:literal) => {{
                if $slot.is_some() {
                    return Err(RtfError::MalformedDocument(
                        concat!("duplicate RTF legacy text-box ", $name).to_string(),
                    ));
                }
                $slot = Some($value);
                true
            }};
        }
        Ok(match control {
            ControlWord::LegacyAnchorXPage => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Page,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorXMargin => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Margin,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorXColumn => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Column,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorYPage => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Page,
                "vertical anchor"
            ),
            ControlWord::LegacyAnchorYMargin => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Margin,
                "vertical anchor"
            ),
            ControlWord::LegacyAnchorYParagraph => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Paragraph,
                "vertical anchor"
            ),
            ControlWord::LegacyDrawingX(value) => set_once!(builder.x, *value, "x"),
            ControlWord::LegacyDrawingY(value) => set_once!(builder.y, *value, "y"),
            ControlWord::LegacyDrawingWidth(value) => {
                set_once!(builder.width, *value, "width")
            },
            ControlWord::LegacyDrawingHeightSize(value) => {
                set_once!(builder.height, *value, "height")
            },
            ControlWord::LegacyTextBoxMargin(value) => {
                set_once!(builder.margin, *value, "margin")
            },
            ControlWord::LegacyDrawingHeight(value) => {
                set_once!(builder.z_order, *value, "z-order")
            },
            ControlWord::LegacyTextLeftRightTopBottom => set_once!(
                builder.direction,
                crate::LegacyTextDirection::LeftToRightTopToBottom,
                "direction"
            ),
            ControlWord::LegacyTextLeftRightTopBottomVertical => set_once!(
                builder.direction,
                crate::LegacyTextDirection::LeftToRightTopToBottomVertical,
                "direction"
            ),
            ControlWord::LegacyTextTopBottomRightLeft => set_once!(
                builder.direction,
                crate::LegacyTextDirection::TopToBottomRightToLeft,
                "direction"
            ),
            ControlWord::LegacyTextTopBottomRightLeftVertical => set_once!(
                builder.direction,
                crate::LegacyTextDirection::TopToBottomRightToLeftVertical,
                "direction"
            ),
            ControlWord::LegacyTextBottomTopLeftRight => set_once!(
                builder.direction,
                crate::LegacyTextDirection::BottomToTopLeftToRight,
                "direction"
            ),
            _ => false,
        })
    }

    pub(super) fn parse_legacy_text_box(&mut self) -> RtfResult<Option<crate::LegacyTextBox<'a>>> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing text box may occur only in the document body".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and do
        let mut builder = LegacyTextBoxBuilder::default();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if !builder.saw_text_box {
                        return Ok(None);
                    }
                    let text = builder.text.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF legacy text box lacks dptxbxtext".to_string(),
                        )
                    })?;
                    let text_box = crate::LegacyTextBox {
                        text: Cow::Borrowed(self.arena.alloc_str(&text) as &str),
                        shapes: std::mem::take(&mut builder.shapes).into_iter().collect(),
                        shape_groups: std::mem::take(&mut builder.shape_groups)
                            .into_iter()
                            .collect(),
                        drawing_order: std::mem::take(&mut builder.drawing_order),
                        story_events: std::mem::take(&mut builder.story_events),
                        position: self.body_text_len,
                        horizontal_anchor: builder.horizontal_anchor,
                        vertical_anchor: builder.vertical_anchor,
                        x: builder.x,
                        y: builder.y,
                        width: builder.width,
                        height: builder.height,
                        margin: builder.margin,
                        z_order: builder.z_order,
                        direction: builder.direction.unwrap_or_default(),
                    };
                    text_box.validate()?;
                    if self.legacy_text_boxes.len() >= crate::legacy_text_box::MAX_LEGACY_TEXT_BOXES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text-box count exceeds the safety limit".to_string(),
                        ));
                    }
                    self.legacy_text_box_text_bytes = self
                        .legacy_text_box_text_bytes
                        .checked_add(text_box.text.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF legacy text-box text size overflow".to_string(),
                            )
                        })?;
                    if self.legacy_text_box_text_bytes
                        > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text-box text exceeds the aggregate safety limit"
                                .to_string(),
                        ));
                    }
                    return Ok(Some(text_box));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyTextBoxText))
                    ) =>
                {
                    if !builder.saw_text_box {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy drawing dptxbxtext must follow dptxbx".to_string(),
                        ));
                    }
                    if builder.text.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text box contains duplicate dptxbxtext".to_string(),
                        ));
                    }
                    builder.text = Some(self.parse_legacy_text_box_text(&mut builder)?);
                },
                Some(Token::OpenBrace) => self.skip_group()?,
                Some(Token::Control(ControlWord::LegacyTextBox)) => {
                    if builder.saw_text_box {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy drawing contains duplicate dptxbx".to_string(),
                        ));
                    }
                    builder.saw_text_box = true;
                    self.pos += 1;
                },
                Some(Token::Control(control)) => {
                    Self::apply_legacy_text_box_control(&mut builder, control)?;
                    self.pos += 1;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => self.pos += 1,
                Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing contains orphan text".to_string(),
                    ));
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_legacy_text_box_text(
        &mut self,
        builder: &mut LegacyTextBoxBuilder,
    ) -> RtfResult<String> {
        self.pos += 2; // opening brace and dptxbxtext
        let mut depth = 0usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut text = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    return Ok(text);
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    if self.is_root_drawing_group() {
                        let mut shapes = Vec::new();
                        let mut groups = Vec::new();
                        let mut order = Vec::new();
                        self.parse_story_drawing_group(
                            text.len(),
                            &mut shapes,
                            &mut groups,
                            &mut order,
                        )?;
                        let shape_base = builder.shapes.len();
                        let group_base = builder.shape_groups.len();
                        builder.shapes.extend(
                            shapes
                                .into_iter()
                                .map(super::super::super::shape::Shape::into_owned),
                        );
                        builder.shape_groups.extend(
                            groups
                                .into_iter()
                                .map(super::super::super::shape::ShapeGroup::into_owned),
                        );
                        let order_start = builder.drawing_order.len();
                        builder.drawing_order.extend(order.into_iter().map(
                            |drawing| match drawing {
                                crate::StoryDrawing::Shape(index) => {
                                    crate::StoryDrawing::Shape(shape_base + index)
                                },
                                crate::StoryDrawing::ShapeGroup(index) => {
                                    crate::StoryDrawing::ShapeGroup(group_base + index)
                                },
                            },
                        ));
                        builder.story_events.extend(
                            builder
                                .drawing_order
                                .iter()
                                .skip(order_start)
                                .copied()
                                .map(crate::StoryEvent::Drawing),
                        );
                        continue;
                    }
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::FormField
                        ))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text box contains an active nested destination".to_string(),
                        ));
                    }
                    depth += 1;
                    self.pos += 1;
                },
                Some(Token::Control(control))
                    if Self::apply_legacy_text_box_control(builder, control)? =>
                {
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
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                    text.push('\n');
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Page(param))) => {
                    require_parameterless(*param, "page")?;
                    builder
                        .story_events
                        .push(crate::StoryEvent::PageBreak(crate::PageBreak::new(
                            text.len(),
                        )));
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
                    | ControlWord::FormField,
                )) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy text box contains active content".to_string(),
                    ));
                },
                Some(Token::Control(_)) => self.pos += 1,
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy text box cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF legacy text-box text exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    pub(super) fn is_legacy_drawing_control(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::LegacyTextBox
                | ControlWord::LegacyTextBoxText
                | ControlWord::LegacyDrawingLock
                | ControlWord::LegacyDrawingGroup
                | ControlWord::LegacyDrawingCount(_)
                | ControlWord::LegacyDrawingEndGroup
                | ControlWord::LegacyDrawingArc
                | ControlWord::LegacyDrawingCallout
                | ControlWord::LegacyDrawingEllipse
                | ControlWord::LegacyDrawingLine
                | ControlWord::LegacyDrawingPolygon
                | ControlWord::LegacyDrawingPolyline
                | ControlWord::LegacyDrawingRectangle
                | ControlWord::LegacyDrawingRoundRectangle
                | ControlWord::LegacyDrawingPointX(_)
                | ControlWord::LegacyDrawingPointY(_)
                | ControlWord::LegacyDrawingArcFlipX
                | ControlWord::LegacyDrawingArcFlipY
                | ControlWord::LegacyCalloutType(_)
                | ControlWord::LegacyCalloutAngle(_)
                | ControlWord::LegacyCalloutAccent
                | ControlWord::LegacyCalloutSmartAttach
                | ControlWord::LegacyCalloutBestFit
                | ControlWord::LegacyCalloutMinusX
                | ControlWord::LegacyCalloutMinusY
                | ControlWord::LegacyCalloutBorder
                | ControlWord::LegacyCalloutAttachment(_)
                | ControlWord::LegacyCalloutDescent(_)
                | ControlWord::LegacyCalloutOffset(_)
                | ControlWord::LegacyCalloutLength(_)
                | ControlWord::LegacyDrawingLineStyle(_)
                | ControlWord::LegacyDrawingLineGray(_)
                | ControlWord::LegacyDrawingLineRed(_)
                | ControlWord::LegacyDrawingLineGreen(_)
                | ControlWord::LegacyDrawingLineBlue(_)
                | ControlWord::LegacyDrawingLinePalette
                | ControlWord::LegacyDrawingLineWidth(_)
                | ControlWord::LegacyDrawingFillForegroundGray(_)
                | ControlWord::LegacyDrawingFillForegroundRed(_)
                | ControlWord::LegacyDrawingFillForegroundGreen(_)
                | ControlWord::LegacyDrawingFillForegroundBlue(_)
                | ControlWord::LegacyDrawingFillForegroundPalette
                | ControlWord::LegacyDrawingFillBackgroundGray(_)
                | ControlWord::LegacyDrawingFillBackgroundRed(_)
                | ControlWord::LegacyDrawingFillBackgroundGreen(_)
                | ControlWord::LegacyDrawingFillBackgroundBlue(_)
                | ControlWord::LegacyDrawingFillBackgroundPalette
                | ControlWord::LegacyDrawingFillPattern(_)
                | ControlWord::LegacyDrawingStartArrowFill(_)
                | ControlWord::LegacyDrawingStartArrowLength(_)
                | ControlWord::LegacyDrawingStartArrowWidth(_)
                | ControlWord::LegacyDrawingEndArrowFill(_)
                | ControlWord::LegacyDrawingEndArrowLength(_)
                | ControlWord::LegacyDrawingEndArrowWidth(_)
                | ControlWord::LegacyDrawingShadow
                | ControlWord::LegacyDrawingShadowX(_)
                | ControlWord::LegacyDrawingShadowY(_)
        )
    }

    pub(super) fn legacy_primitive_start(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::LegacyDrawingGroup
                | ControlWord::LegacyDrawingCallout
                | ControlWord::LegacyDrawingLine
                | ControlWord::LegacyDrawingRectangle
                | ControlWord::LegacyTextBox
                | ControlWord::LegacyDrawingEllipse
                | ControlWord::LegacyDrawingPolyline
                | ControlWord::LegacyDrawingArc
        )
    }

    pub(super) fn legacy_do_starts_with_text_box(&self) -> bool {
        self.tokens
            .get(self.pos.saturating_add(2)..)
            .unwrap_or_default()
            .iter()
            .find_map(|token| match token {
                Token::Control(control) if Self::legacy_primitive_start(control) => {
                    Some(matches!(control, ControlWord::LegacyTextBox))
                },
                Token::CloseBrace => Some(false),
                _ => None,
            })
            .unwrap_or(false)
    }

    pub(super) fn skip_legacy_whitespace(&mut self) {
        while matches!(self.tokens.get(self.pos), Some(Token::Text(text)) if text.trim().is_empty())
        {
            self.pos += 1;
        }
    }

    pub(super) fn parse_legacy_drawing(&mut self) -> RtfResult<Option<crate::LegacyDrawing<'a>>> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing may occur only in the document body".to_string(),
            ));
        }
        self.parse_legacy_drawing_at(self.body_text_len)
    }

    pub(super) fn parse_legacy_drawing_at(
        &mut self,
        position: usize,
    ) -> RtfResult<Option<crate::LegacyDrawing<'a>>> {
        self.pos += 2;
        let mut horizontal_anchor = None;
        let mut vertical_anchor = None;
        let mut z_order = None;
        let mut locked = false;
        loop {
            self.skip_legacy_whitespace();
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::LegacyAnchorXPage)) => Self::set_legacy_once(
                    &mut horizontal_anchor,
                    crate::LegacyHorizontalAnchor::Page,
                    "horizontal anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyAnchorXMargin)) => Self::set_legacy_once(
                    &mut horizontal_anchor,
                    crate::LegacyHorizontalAnchor::Margin,
                    "horizontal anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyAnchorXColumn)) => Self::set_legacy_once(
                    &mut horizontal_anchor,
                    crate::LegacyHorizontalAnchor::Column,
                    "horizontal anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyAnchorYPage)) => Self::set_legacy_once(
                    &mut vertical_anchor,
                    crate::LegacyVerticalAnchor::Page,
                    "vertical anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyAnchorYMargin)) => Self::set_legacy_once(
                    &mut vertical_anchor,
                    crate::LegacyVerticalAnchor::Margin,
                    "vertical anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyAnchorYParagraph)) => Self::set_legacy_once(
                    &mut vertical_anchor,
                    crate::LegacyVerticalAnchor::Paragraph,
                    "vertical anchor",
                )?,
                Some(Token::Control(ControlWord::LegacyDrawingHeight(value))) => {
                    Self::set_legacy_once(&mut z_order, *value, "z-order")?
                },
                Some(Token::Control(ControlWord::LegacyDrawingLock)) => {
                    if locked {
                        return Err(Self::legacy_error("duplicate dolock"));
                    }
                    locked = true;
                },
                Some(Token::Control(control)) if Self::legacy_primitive_start(control) => break,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(None);
                },
                Some(_) => {
                    return Err(Self::legacy_error("invalid or out-of-order dohead control"));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        let primitive = self.parse_legacy_primitive(1)?;
        self.skip_legacy_whitespace();
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(Self::legacy_error("trailing content after primitive"));
        }
        self.pos += 1;
        let drawing = crate::LegacyDrawing {
            position,
            horizontal_anchor: horizontal_anchor
                .ok_or_else(|| Self::legacy_error("missing horizontal anchor"))?,
            vertical_anchor: vertical_anchor
                .ok_or_else(|| Self::legacy_error("missing vertical anchor"))?,
            z_order: z_order.ok_or_else(|| Self::legacy_error("missing dodhgt"))?,
            locked,
            primitive,
        };
        drawing.validate()?;
        if self.legacy_drawings.len() >= crate::MAX_LEGACY_DRAWINGS {
            return Err(Self::legacy_error(
                "top-level count exceeds the safety limit",
            ));
        }
        Ok(Some(drawing))
    }

    pub(super) fn legacy_error(message: &str) -> RtfError {
        RtfError::MalformedDocument(format!("RTF legacy drawing {message}"))
    }

    pub(super) fn set_legacy_once<T>(slot: &mut Option<T>, value: T, name: &str) -> RtfResult<()> {
        if slot.replace(value).is_some() {
            return Err(Self::legacy_error(&format!("contains duplicate {name}")));
        }
        Ok(())
    }

    pub(super) fn parse_legacy_geometry(&mut self) -> RtfResult<crate::LegacyDrawingGeometry> {
        self.skip_legacy_whitespace();
        let x = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingX(value))) => *value,
            _ => return Err(Self::legacy_error("geometry must begin with dpx")),
        };
        self.pos += 1;
        self.skip_legacy_whitespace();
        let y = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingY(value))) => *value,
            _ => return Err(Self::legacy_error("dpx must be followed by dpy")),
        };
        self.pos += 1;
        self.skip_legacy_whitespace();
        let width = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingWidth(value))) => *value,
            _ => return Err(Self::legacy_error("dpy must be followed by dpxsize")),
        };
        self.pos += 1;
        self.skip_legacy_whitespace();
        let height = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingHeightSize(value))) => *value,
            _ => return Err(Self::legacy_error("dpxsize must be followed by dpysize")),
        };
        self.pos += 1;
        let geometry = crate::LegacyDrawingGeometry {
            x,
            y,
            width,
            height,
        };
        geometry.validate()?;
        Ok(geometry)
    }

    pub(super) fn parse_legacy_point(&mut self) -> RtfResult<crate::LegacyDrawingPoint> {
        self.skip_legacy_whitespace();
        let x = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingPointX(value))) => *value,
            _ => return Err(Self::legacy_error("point must begin with dpptx")),
        };
        self.pos += 1;
        self.skip_legacy_whitespace();
        let y = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingPointY(value))) => *value,
            _ => return Err(Self::legacy_error("dpptx must be followed by dppty")),
        };
        self.pos += 1;
        Ok(crate::LegacyDrawingPoint { x, y })
    }

    pub(super) fn parse_legacy_primitive(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::LegacyDrawingPrimitive<'a>> {
        if depth > crate::MAX_LEGACY_DRAWING_DEPTH {
            return Err(Self::legacy_error("nesting exceeds the safety limit"));
        }
        self.legacy_drawing_primitives = self
            .legacy_drawing_primitives
            .checked_add(1)
            .ok_or_else(|| Self::legacy_error("primitive count overflow"))?;
        if self.legacy_drawing_primitives > crate::MAX_LEGACY_DRAWING_PRIMITIVES {
            return Err(Self::legacy_error(
                "primitive count exceeds the safety limit",
            ));
        }
        self.skip_legacy_whitespace();
        let control = self
            .tokens
            .get(self.pos)
            .cloned()
            .ok_or(RtfError::UnexpectedEof)?;
        self.pos += 1;
        match control {
            Token::Control(ControlWord::LegacyDrawingGroup) => self.parse_legacy_group(depth),
            Token::Control(ControlWord::LegacyDrawingCallout) => self.parse_legacy_callout(depth),
            Token::Control(ControlWord::LegacyDrawingLine) => {
                let start = self.parse_legacy_point()?;
                let end = self.parse_legacy_point()?;
                let (geometry, properties, _) =
                    self.parse_legacy_simple_tail(LegacySimpleKind::Line)?;
                Ok(crate::LegacyDrawingPrimitive::Line {
                    start,
                    end,
                    geometry,
                    properties,
                })
            },
            Token::Control(ControlWord::LegacyDrawingRectangle) => {
                let (geometry, properties, rounded) =
                    self.parse_legacy_simple_tail(LegacySimpleKind::Rectangle)?;
                Ok(crate::LegacyDrawingPrimitive::Rectangle {
                    rounded: rounded != 0,
                    geometry,
                    properties,
                })
            },
            Token::Control(ControlWord::LegacyDrawingEllipse) => {
                let (geometry, properties, _) =
                    self.parse_legacy_simple_tail(LegacySimpleKind::Ellipse)?;
                Ok(crate::LegacyDrawingPrimitive::Ellipse {
                    geometry,
                    properties,
                })
            },
            Token::Control(ControlWord::LegacyDrawingPolyline) => self.parse_legacy_polyline(),
            Token::Control(ControlWord::LegacyDrawingArc) => {
                let (geometry, properties, flags) =
                    self.parse_legacy_simple_tail(LegacySimpleKind::Arc)?;
                Ok(crate::LegacyDrawingPrimitive::Arc {
                    flip_x: flags & 1 != 0,
                    flip_y: flags & 2 != 0,
                    geometry,
                    properties,
                })
            },
            Token::Control(ControlWord::LegacyTextBox) => self.parse_nested_legacy_text_box(),
            _ => Err(Self::legacy_error("expected a drawing primitive")),
        }
    }

    pub(super) fn parse_legacy_group(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::LegacyDrawingPrimitive<'a>> {
        self.skip_legacy_whitespace();
        let declared = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingCount(value))) if *value > 0 => {
                usize::try_from(*value).map_err(|_| Self::legacy_error("invalid dpcount"))?
            },
            _ => return Err(Self::legacy_error("dpgroup lacks positive dpcount")),
        };
        self.pos += 1;
        let geometry = self.parse_legacy_geometry()?;
        let mut children = Vec::new();
        loop {
            self.skip_legacy_whitespace();
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::LegacyDrawingEndGroup)) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Control(control)) if Self::legacy_primitive_start(control) => {
                    children.push(self.parse_legacy_primitive(depth + 1)?)
                },
                Some(_) => return Err(Self::legacy_error("invalid content in dpgroup")),
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        if declared != children.len() && declared != children.len().saturating_add(1) {
            return Err(Self::legacy_error("dpcount does not match group children"));
        }
        let end_geometry = self.parse_legacy_geometry()?;
        Ok(crate::LegacyDrawingPrimitive::Group {
            geometry,
            children,
            end_geometry,
        })
    }

    pub(super) fn parse_legacy_polyline(&mut self) -> RtfResult<crate::LegacyDrawingPrimitive<'a>> {
        self.skip_legacy_whitespace();
        let closed = if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LegacyDrawingPolygon))
        ) {
            self.pos += 1;
            true
        } else {
            false
        };
        self.skip_legacy_whitespace();
        let count = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyDrawingCount(value))) => usize::try_from(*value)
                .map_err(|_| Self::legacy_error("invalid polyline point count"))?,
            _ => return Err(Self::legacy_error("dppolyline lacks dppolycount")),
        };
        if count == 0 || count > crate::MAX_LEGACY_DRAWING_POINTS {
            return Err(Self::legacy_error(
                "polyline point count exceeds the safety limit",
            ));
        }
        self.pos += 1;
        let mut points = Vec::with_capacity(count);
        for _ in 0..count {
            points.push(self.parse_legacy_point()?);
        }
        self.legacy_drawing_points = self
            .legacy_drawing_points
            .checked_add(count)
            .ok_or_else(|| Self::legacy_error("point count overflow"))?;
        if self.legacy_drawing_points > crate::MAX_LEGACY_DRAWING_TOTAL_POINTS {
            return Err(Self::legacy_error(
                "aggregate point count exceeds the safety limit",
            ));
        }
        let (geometry, properties, _) =
            self.parse_legacy_simple_tail(LegacySimpleKind::Polyline)?;
        Ok(crate::LegacyDrawingPrimitive::Polyline {
            closed,
            points,
            geometry,
            properties,
        })
    }

    pub(super) fn parse_nested_legacy_text_box(
        &mut self,
    ) -> RtfResult<crate::LegacyDrawingPrimitive<'a>> {
        let mut margin = None;
        let mut direction = None;
        let mut text = None;
        let mut shapes = Vec::new();
        let mut shape_groups = Vec::new();
        let mut drawing_order = Vec::new();
        let mut story_events = Vec::new();
        loop {
            self.skip_legacy_whitespace();
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::LegacyTextBoxMargin(value))) => {
                    Self::set_legacy_once(&mut margin, *value, "text-box margin")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyTextLeftRightTopBottom)) => {
                    Self::set_legacy_once(
                        &mut direction,
                        crate::LegacyTextDirection::LeftToRightTopToBottom,
                        "text direction",
                    )?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyTextLeftRightTopBottomVertical)) => {
                    Self::set_legacy_once(
                        &mut direction,
                        crate::LegacyTextDirection::LeftToRightTopToBottomVertical,
                        "text direction",
                    )?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyTextTopBottomRightLeft)) => {
                    Self::set_legacy_once(
                        &mut direction,
                        crate::LegacyTextDirection::TopToBottomRightToLeft,
                        "text direction",
                    )?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyTextTopBottomRightLeftVertical)) => {
                    Self::set_legacy_once(
                        &mut direction,
                        crate::LegacyTextDirection::TopToBottomRightToLeftVertical,
                        "text direction",
                    )?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyTextBottomTopLeftRight)) => {
                    Self::set_legacy_once(
                        &mut direction,
                        crate::LegacyTextDirection::BottomToTopLeftToRight,
                        "text direction",
                    )?;
                    self.pos += 1;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyTextBoxText))
                    ) =>
                {
                    if text.is_some() {
                        return Err(Self::legacy_error("duplicate dptxbxtext"));
                    }
                    let mut dummy = LegacyTextBoxBuilder::default();
                    text = Some(self.parse_legacy_text_box_text(&mut dummy)?);
                    shapes = dummy.shapes;
                    shape_groups = dummy.shape_groups;
                    drawing_order = dummy.drawing_order;
                    story_events = dummy.story_events;
                },
                _ => break,
            }
        }
        let geometry = self.parse_legacy_geometry()?;
        let properties = self.parse_legacy_properties_until_boundary()?;
        let text = text.ok_or_else(|| Self::legacy_error("text box lacks dptxbxtext"))?;
        self.legacy_text_box_text_bytes =
            self.legacy_text_box_text_bytes
                .checked_add(text.len())
                .ok_or_else(|| Self::legacy_error("text-box text size overflow"))?;
        if self.legacy_text_box_text_bytes > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES
        {
            return Err(Self::legacy_error(
                "text-box text exceeds aggregate safety limit",
            ));
        }
        let text_box = crate::LegacyTextBox {
            text: Cow::Borrowed(self.arena.alloc_str(&text)),
            shapes,
            shape_groups,
            drawing_order,
            story_events,
            position: self.body_text_len,
            horizontal_anchor: None,
            vertical_anchor: None,
            x: Some(geometry.x),
            y: Some(geometry.y),
            width: Some(geometry.width),
            height: Some(geometry.height),
            margin,
            z_order: None,
            direction: direction.unwrap_or_default(),
        };
        text_box.validate()?;
        Ok(crate::LegacyDrawingPrimitive::TextBox {
            text_box,
            properties,
        })
    }

    pub(super) fn parse_legacy_callout(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::LegacyDrawingPrimitive<'a>> {
        let mut callout_type = None;
        let mut angle = None;
        let mut attachment = None;
        let mut descent = None;
        let mut accent = false;
        let mut smart_attach = false;
        let mut best_fit = false;
        let mut minus_x = false;
        let mut minus_y = false;
        let mut border = false;
        let mut offset = None;
        let mut length = None;
        loop {
            self.skip_legacy_whitespace();
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::LegacyCalloutType(value))) => {
                    Self::set_legacy_once(&mut callout_type, *value, "callout type")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutAngle(value))) => {
                    let value = u8::try_from(*value)
                        .map_err(|_| Self::legacy_error("invalid callout angle"))?;
                    if !matches!(value, 0 | 30 | 45 | 60 | 90) {
                        return Err(Self::legacy_error("invalid callout angle"));
                    }
                    Self::set_legacy_once(&mut angle, value, "callout angle")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutAttachment(value))) => {
                    Self::set_legacy_once(&mut attachment, *value, "callout attachment")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutDescent(value))) => {
                    Self::set_legacy_once(&mut descent, *value, "callout descent")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutOffset(value))) => {
                    Self::set_legacy_once(&mut offset, *value, "callout offset")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutLength(value))) => {
                    Self::set_legacy_once(&mut length, *value, "callout length")?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutAccent)) if !accent => {
                    accent = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutSmartAttach)) if !smart_attach => {
                    smart_attach = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutBestFit)) if !best_fit => {
                    best_fit = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutMinusX)) if !minus_x => {
                    minus_x = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutMinusY)) if !minus_y => {
                    minus_y = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::LegacyCalloutBorder)) if !border => {
                    border = true;
                    self.pos += 1;
                },
                _ => break,
            }
        }
        let geometry = self.parse_legacy_geometry()?;
        self.skip_legacy_whitespace();
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LegacyDrawingPolyline))
        ) {
            return Err(Self::legacy_error("callout lacks polyline"));
        }
        let polyline = Box::new(self.parse_legacy_primitive(depth + 1)?);
        self.skip_legacy_whitespace();
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LegacyTextBox))
        ) {
            return Err(Self::legacy_error("callout lacks trailing text box"));
        }
        let text_box = Box::new(self.parse_legacy_primitive(depth + 1)?);
        let properties = self.parse_legacy_properties_until_boundary()?;
        Ok(crate::LegacyDrawingPrimitive::Callout(
            crate::LegacyCallout {
                callout_type: callout_type
                    .ok_or_else(|| Self::legacy_error("callout lacks type"))?,
                angle,
                accent,
                smart_attach,
                best_fit,
                minus_x,
                minus_y,
                border,
                attachment,
                descent,
                offset: offset.ok_or_else(|| Self::legacy_error("callout lacks offset"))?,
                length: length.ok_or_else(|| Self::legacy_error("callout lacks length"))?,
                polyline,
                text_box,
                geometry,
                properties,
            },
        ))
    }

    pub(super) fn parse_legacy_simple_tail(
        &mut self,
        kind: LegacySimpleKind,
    ) -> RtfResult<(
        crate::LegacyDrawingGeometry,
        crate::LegacyDrawingProperties,
        u8,
    )> {
        let mut flags = 0u8;
        loop {
            self.skip_legacy_whitespace();
            match (kind, self.tokens.get(self.pos)) {
                (
                    LegacySimpleKind::Rectangle,
                    Some(Token::Control(ControlWord::LegacyDrawingRoundRectangle)),
                ) if flags == 0 => {
                    flags = 1;
                    self.pos += 1;
                },
                (
                    LegacySimpleKind::Arc,
                    Some(Token::Control(ControlWord::LegacyDrawingArcFlipX)),
                ) if flags & 1 == 0 => {
                    flags |= 1;
                    self.pos += 1;
                },
                (
                    LegacySimpleKind::Arc,
                    Some(Token::Control(ControlWord::LegacyDrawingArcFlipY)),
                ) if flags & 2 == 0 => {
                    flags |= 2;
                    self.pos += 1;
                },
                _ => break,
            }
        }
        let geometry = self.parse_legacy_geometry()?;
        let properties = self.parse_legacy_properties_until_boundary()?;
        Ok((geometry, properties, flags))
    }

    pub(super) fn parse_legacy_properties_until_boundary(
        &mut self,
    ) -> RtfResult<crate::LegacyDrawingProperties> {
        let mut builder = LegacyPropertiesBuilder::default();
        loop {
            self.skip_legacy_whitespace();
            let Some(Token::Control(control)) = self.tokens.get(self.pos) else {
                break;
            };
            if Self::legacy_primitive_start(control)
                || matches!(control, ControlWord::LegacyDrawingEndGroup)
            {
                break;
            }
            if !builder.apply(control)? {
                break;
            }
            self.pos += 1;
        }
        builder.finish()
    }
}
