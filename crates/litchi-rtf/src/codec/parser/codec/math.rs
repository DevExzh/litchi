use super::*;

impl<'a> Parser<'a> {
    pub(super) fn is_math_scoped_control(control: &ControlWord<'_>) -> bool {
        use crate::lexer::ControlWord as C;
        matches!(
            control,
            C::MathZoneParagraphProperties
                | C::MathZonePictureFallback
                | C::MathAccent
                | C::MathBar
                | C::MathBorderBox
                | C::MathBox
                | C::MathDelimiter
                | C::MathEquationArray
                | C::MathFraction
                | C::MathFunction
                | C::MathGroupChar
                | C::MathLimitLower
                | C::MathLimitUpper
                | C::MathMatrix
                | C::MathNary
                | C::MathPhantom
                | C::MathRadical
                | C::MathScriptPre
                | C::MathScriptSub
                | C::MathScriptSubSup
                | C::MathScriptSup
                | C::MathRun
                | C::MathElement
                | C::MathNumerator
                | C::MathDenominator
                | C::MathDegree
                | C::MathSubscript
                | C::MathSuperscript
                | C::MathLimit
                | C::MathFunctionName
                | C::MathMatrixRow
                | C::MathAccentProperties
                | C::MathBarProperties
                | C::MathBorderBoxProperties
                | C::MathBoxProperties
                | C::MathDelimiterProperties
                | C::MathEquationArrayProperties
                | C::MathFractionProperties
                | C::MathFunctionProperties
                | C::MathGroupCharProperties
                | C::MathLimitLowerProperties
                | C::MathLimitUpperProperties
                | C::MathMatrixProperties
                | C::MathNaryProperties
                | C::MathPhantomProperties
                | C::MathRadicalProperties
                | C::MathScriptPreProperties
                | C::MathScriptSubProperties
                | C::MathScriptSubSupProperties
                | C::MathScriptSupProperties
                | C::MathRunProperties
                | C::MathControlProperties
                | C::MathMatrixColumns
                | C::MathMatrixColumn
                | C::MathMatrixColumnProperties
                | C::MathArgumentProperties
                | C::MathPropertyArgumentSize(_)
                | C::MathPropertyType(_)
                | C::MathPropertyGrow(_)
                | C::MathPropertyChar(_)
                | C::MathPropertyBeginChar(_)
                | C::MathPropertyEndChar(_)
                | C::MathPropertySeparatorChar(_)
                | C::MathPropertyPosition(_)
                | C::MathPropertyVerticalJustify(_)
                | C::MathPropertyBaseJustify(_)
                | C::MathPropertyJustify(_)
                | C::MathPropertyAlign(_)
                | C::MathPropertyAlignScript(_)
                | C::MathPropertyDegreeHide(_)
                | C::MathPropertyDifferential(_)
                | C::MathPropertyDifferentialStyle(_)
                | C::MathPropertyHideBottom(_)
                | C::MathPropertyHideLeft(_)
                | C::MathPropertyHideRight(_)
                | C::MathPropertyHideTop(_)
                | C::MathPropertyLimitLocation(_)
                | C::MathPropertyPlaceholderHide(_)
                | C::MathPropertySubscriptHide(_)
                | C::MathPropertySuperscriptHide(_)
                | C::MathPropertyStrikeBltr(_)
                | C::MathPropertyStrikeHorizontal(_)
                | C::MathPropertyStrikeTlbr(_)
                | C::MathPropertyStrikeVertical(_)
                | C::MathPropertyStyle(_)
                | C::MathPropertyScript(_)
                | C::MathPropertyTransparent(_)
                | C::MathPropertyShow(_)
                | C::MathPropertyShape(_)
                | C::MathPropertyZeroAscent(_)
                | C::MathPropertyZeroDescent(_)
                | C::MathPropertyZeroWidth(_)
                | C::MathPropertyOperatorEmulator(_)
                | C::MathPropertyNoBreak(_)
                | C::MathPropertyNormalText(_)
                | C::MathPropertyLiteral(_)
                | C::MathPropertyMatrixColumnGap(_)
                | C::MathPropertyMatrixColumnGapRule(_)
                | C::MathPropertyMatrixColumnSpacing(_)
                | C::MathPropertyMatrixCellCount(_)
                | C::MathPropertyMatrixCellJustify(_)
                | C::MathPropertyRowSpacing(_)
                | C::MathPropertyRowSpacingRule(_)
                | C::MathPropertyBreak(_)
        )
    }

    /// Map a structure control word to its structure kind.
    pub(super) fn math_structure_kind(
        control: &ControlWord<'_>,
    ) -> Option<crate::MathStructureKind> {
        use crate::MathStructureKind as K;
        use crate::lexer::ControlWord as C;
        Some(match control {
            C::MathAccent => K::Accent,
            C::MathBar => K::Bar,
            C::MathBorderBox => K::BorderBox,
            C::MathBox => K::Box,
            C::MathDelimiter => K::Delimiter,
            C::MathEquationArray => K::EquationArray,
            C::MathFraction => K::Fraction,
            C::MathFunction => K::Function,
            C::MathGroupChar => K::GroupChar,
            C::MathLimitLower => K::LimitLower,
            C::MathLimitUpper => K::LimitUpper,
            C::MathMatrix => K::Matrix,
            C::MathNary => K::Nary,
            C::MathPhantom => K::Phantom,
            C::MathRadical => K::Radical,
            C::MathScriptPre => K::ScriptPre,
            C::MathScriptSub => K::ScriptSub,
            C::MathScriptSubSup => K::ScriptSubSup,
            C::MathScriptSup => K::ScriptSup,
            _ => return None,
        })
    }

    /// Map a structure property-destination control to its kind.
    pub(super) fn math_structure_properties_kind(
        control: &ControlWord<'_>,
    ) -> Option<crate::MathStructureKind> {
        use crate::MathStructureKind as K;
        use crate::lexer::ControlWord as C;
        Some(match control {
            C::MathAccentProperties => K::Accent,
            C::MathBarProperties => K::Bar,
            C::MathBorderBoxProperties => K::BorderBox,
            C::MathBoxProperties => K::Box,
            C::MathDelimiterProperties => K::Delimiter,
            C::MathEquationArrayProperties => K::EquationArray,
            C::MathFractionProperties => K::Fraction,
            C::MathFunctionProperties => K::Function,
            C::MathGroupCharProperties => K::GroupChar,
            C::MathLimitLowerProperties => K::LimitLower,
            C::MathLimitUpperProperties => K::LimitUpper,
            C::MathMatrixProperties => K::Matrix,
            C::MathNaryProperties => K::Nary,
            C::MathPhantomProperties => K::Phantom,
            C::MathRadicalProperties => K::Radical,
            C::MathScriptPreProperties => K::ScriptPre,
            C::MathScriptSubProperties => K::ScriptSub,
            C::MathScriptSubSupProperties => K::ScriptSubSup,
            C::MathScriptSupProperties => K::ScriptSup,
            _ => return None,
        })
    }

    /// Map an argument control word to its element role.
    pub(super) fn math_element_role(control: &ControlWord<'_>) -> Option<crate::MathElementRole> {
        use crate::MathElementRole as R;
        use crate::lexer::ControlWord as C;
        Some(match control {
            C::MathElement => R::Element,
            C::MathNumerator => R::Numerator,
            C::MathDenominator => R::Denominator,
            C::MathDegree => R::Degree,
            C::MathSubscript => R::Subscript,
            C::MathSuperscript => R::Superscript,
            C::MathLimit => R::Limit,
            C::MathFunctionName => R::FunctionName,
            _ => return None,
        })
    }

    /// Map a property control word to its name and numeric parameter.
    pub(super) fn math_property_name(
        control: &ControlWord<'_>,
    ) -> Option<(crate::MathPropertyName, Option<i32>)> {
        use crate::MathPropertyName as N;
        use crate::lexer::ControlWord as C;
        let (name, param) = match control {
            C::MathPropertyType(param) => (N::Type, param),
            C::MathPropertyGrow(param) => (N::Grow, param),
            C::MathPropertyChar(param) => (N::Char, param),
            C::MathPropertyBeginChar(param) => (N::BeginChar, param),
            C::MathPropertyEndChar(param) => (N::EndChar, param),
            C::MathPropertySeparatorChar(param) => (N::SeparatorChar, param),
            C::MathPropertyPosition(param) => (N::Position, param),
            C::MathPropertyVerticalJustify(param) => (N::VerticalJustify, param),
            C::MathPropertyBaseJustify(param) => (N::BaseJustify, param),
            C::MathPropertyJustify(param) => (N::Justify, param),
            C::MathPropertyAlign(param) => (N::Align, param),
            C::MathPropertyAlignScript(param) => (N::AlignScript, param),
            C::MathPropertyDegreeHide(param) => (N::DegreeHide, param),
            C::MathPropertyDifferential(param) => (N::Differential, param),
            C::MathPropertyDifferentialStyle(param) => (N::DifferentialStyle, param),
            C::MathPropertyHideBottom(param) => (N::HideBottom, param),
            C::MathPropertyHideLeft(param) => (N::HideLeft, param),
            C::MathPropertyHideRight(param) => (N::HideRight, param),
            C::MathPropertyHideTop(param) => (N::HideTop, param),
            C::MathPropertyLimitLocation(param) => (N::LimitLocation, param),
            C::MathPropertyPlaceholderHide(param) => (N::PlaceholderHide, param),
            C::MathPropertySubscriptHide(param) => (N::SubscriptHide, param),
            C::MathPropertySuperscriptHide(param) => (N::SuperscriptHide, param),
            C::MathPropertyStrikeBltr(param) => (N::StrikeBottomLeftToTopRight, param),
            C::MathPropertyStrikeHorizontal(param) => (N::StrikeHorizontal, param),
            C::MathPropertyStrikeTlbr(param) => (N::StrikeTopLeftToBottomRight, param),
            C::MathPropertyStrikeVertical(param) => (N::StrikeVertical, param),
            C::MathPropertyStyle(param) => (N::Style, param),
            C::MathPropertyScript(param) => (N::Script, param),
            C::MathPropertyTransparent(param) => (N::Transparent, param),
            C::MathPropertyShow(param) => (N::Show, param),
            C::MathPropertyShape(param) => (N::Shape, param),
            C::MathPropertyZeroAscent(param) => (N::ZeroAscent, param),
            C::MathPropertyZeroDescent(param) => (N::ZeroDescent, param),
            C::MathPropertyZeroWidth(param) => (N::ZeroWidth, param),
            C::MathPropertyOperatorEmulator(param) => (N::OperatorEmulator, param),
            C::MathPropertyNoBreak(param) => (N::NoBreak, param),
            C::MathPropertyNormalText(param) => (N::NormalText, param),
            C::MathPropertyLiteral(param) => (N::Literal, param),
            C::MathPropertyMatrixColumnGap(param) => (N::MatrixColumnGap, param),
            C::MathPropertyMatrixColumnGapRule(param) => (N::MatrixColumnGapRule, param),
            C::MathPropertyMatrixColumnSpacing(param) => (N::MatrixColumnSpacing, param),
            C::MathPropertyMatrixCellCount(param) => (N::MatrixCellCount, param),
            C::MathPropertyMatrixCellJustify(param) => (N::MatrixCellJustify, param),
            C::MathPropertyRowSpacing(param) => (N::RowSpacing, param),
            C::MathPropertyRowSpacingRule(param) => (N::RowSpacingRule, param),
            C::MathPropertyBreak(param) => (N::Break, param),
            C::MathPropertyArgumentSize(param) => (N::ArgumentSize, param),
            _ => return None,
        };
        Some((name, *param))
    }

    /// Parse an `\mmath` or `\mmathPara` zone destination group.
    ///
    /// Expects `self.pos` at the zone control word and consumes tokens through
    /// the group's closing brace. Math run text is stored in the typed zone
    /// tree only and never enters the body story (like field results).
    pub(super) fn parse_math_zone_destination(&mut self, display: bool) -> RtfResult<()> {
        if self.math_zones.len() >= crate::math::MAX_MATH_ZONES {
            return Err(RtfError::MalformedDocument(
                "RTF math zone count exceeds the safety limit".to_string(),
            ));
        }
        self.pos += 1; // zone control
        let mut paragraph_properties = None;
        let mut content = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => match self.tokens.get(self.pos + 1) {
                    Some(Token::Control(ControlWord::MathZoneParagraphProperties)) => {
                        if !display || paragraph_properties.is_some() || !content.is_empty() {
                            return Err(RtfError::MalformedDocument(
                                "RTF math paragraph properties are misplaced".to_string(),
                            ));
                        }
                        paragraph_properties = Some(self.parse_math_properties_group(
                            crate::MathPropertiesKind::Paragraph,
                            1,
                        )?);
                    },
                    Some(Token::Control(ControlWord::IgnorableDestination))
                        if matches!(
                            self.tokens.get(self.pos + 2),
                            Some(Token::Control(ControlWord::MathZonePictureFallback))
                        ) =>
                    {
                        // Fallback renderings are never modeled; skip them.
                        self.skip_group()?;
                    },
                    Some(Token::Control(_)) => {
                        content.push(self.parse_math_object(1)?);
                        if content.len() > crate::math::MAX_MATH_OBJECTS_PER_CONTAINER {
                            return Err(RtfError::MalformedDocument(
                                "RTF math object count exceeds the safety limit".to_string(),
                            ));
                        }
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF math zone contains unsupported grouped content".to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math zone contains ungrouped, binary, or active data".to_string(),
                    ));
                },
            }
        }
        let kind = if display {
            crate::MathZoneKind::Display
        } else {
            crate::MathZoneKind::Inline
        };
        let zone = crate::MathZone::new(kind, paragraph_properties, content, self.body_text_len)?;
        let index = self.math_zones.len();
        self.math_zones.push(zone);
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::MathZone(index),
        ));
        Ok(())
    }

    /// Parse one math object group; `self.pos` is at its opening brace.
    pub(super) fn parse_math_object(&mut self, depth: usize) -> RtfResult<crate::MathObject<'a>> {
        if depth > crate::math::MAX_MATH_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF math nesting depth exceeds the safety limit".to_string(),
            ));
        }
        let control = match self.tokens.get(self.pos + 1) {
            Some(Token::Control(control)) => *control,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF math object group must start with a control word".to_string(),
                ));
            },
        };
        if control == ControlWord::MathRun {
            return self.parse_math_run(depth);
        }
        let Some(kind) = Self::math_structure_kind(&control) else {
            return Err(RtfError::MalformedDocument(
                "RTF math zone contains an unsupported object destination".to_string(),
            ));
        };
        self.parse_math_structure(kind, depth)
    }

    /// Parse a math structure group; `self.pos` is at its opening brace.
    pub(super) fn parse_math_structure(
        &mut self,
        kind: crate::MathStructureKind,
        depth: usize,
    ) -> RtfResult<crate::MathObject<'a>> {
        self.pos += 2; // opening brace and structure control
        let mut properties = None;
        let mut children = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => match self.tokens.get(self.pos + 1) {
                    Some(Token::Control(control))
                        if Self::math_structure_properties_kind(control) == Some(kind) =>
                    {
                        if properties.is_some() || !children.is_empty() {
                            return Err(RtfError::MalformedDocument(
                                "RTF math structure properties are misplaced".to_string(),
                            ));
                        }
                        properties = Some(self.parse_math_properties_group(
                            crate::MathPropertiesKind::Structure(kind),
                            depth,
                        )?);
                    },
                    Some(Token::Control(ControlWord::MathMatrixRow))
                        if kind == crate::MathStructureKind::Matrix =>
                    {
                        children.push(crate::MathStructureChild::MatrixRow(
                            self.parse_math_matrix_row(depth)?,
                        ));
                    },
                    Some(Token::Control(control)) => {
                        let Some(role) = Self::math_element_role(control) else {
                            return Err(RtfError::MalformedDocument(
                                "RTF math structure contains an unsupported child destination"
                                    .to_string(),
                            ));
                        };
                        children.push(crate::MathStructureChild::Element(
                            self.parse_math_element(role, depth)?,
                        ));
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF math structure contains unsupported grouped content".to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math structure contains ungrouped, binary, or active data".to_string(),
                    ));
                },
            }
            if children.len() > crate::math::MAX_MATH_OBJECTS_PER_CONTAINER {
                return Err(RtfError::MalformedDocument(
                    "RTF math object count exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(crate::MathObject::Structure(crate::MathStructure::new(
            kind, properties, children,
        )?))
    }

    /// Parse a math argument group (`\me`, `\mnum`, ...); `self.pos` is at its
    /// opening brace.
    pub(super) fn parse_math_element(
        &mut self,
        role: crate::MathElementRole,
        depth: usize,
    ) -> RtfResult<crate::MathElement<'a>> {
        self.pos += 2; // opening brace and element control
        let mut argument_properties = None;
        let mut content = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => match self.tokens.get(self.pos + 1) {
                    Some(Token::Control(ControlWord::MathArgumentProperties)) => {
                        if argument_properties.is_some() || !content.is_empty() {
                            return Err(RtfError::MalformedDocument(
                                "RTF math argument properties are misplaced".to_string(),
                            ));
                        }
                        argument_properties = Some(self.parse_math_properties_group(
                            crate::MathPropertiesKind::Argument,
                            depth,
                        )?);
                    },
                    Some(Token::Control(_)) => {
                        content.push(self.parse_math_object(depth + 1)?);
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF math argument contains unsupported grouped content".to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math argument contains ungrouped, binary, or active data".to_string(),
                    ));
                },
            }
            if content.len() > crate::math::MAX_MATH_OBJECTS_PER_CONTAINER {
                return Err(RtfError::MalformedDocument(
                    "RTF math object count exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(crate::MathElement {
            role,
            argument_properties,
            content,
        })
    }

    /// Parse a matrix row group (`\mmr`); `self.pos` is at its opening brace.
    pub(super) fn parse_math_matrix_row(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::MathMatrixRow<'a>> {
        self.pos += 2; // opening brace and row control
        let mut cells = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::MathElement))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF math matrix rows may contain only me cell destinations"
                                .to_string(),
                        ));
                    }
                    cells.push(self.parse_math_element(crate::MathElementRole::Element, depth)?);
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math matrix row contains ungrouped, binary, or active data"
                            .to_string(),
                    ));
                },
            }
            if cells.len() > crate::math::MAX_MATH_OBJECTS_PER_CONTAINER {
                return Err(RtfError::MalformedDocument(
                    "RTF math object count exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(crate::MathMatrixRow { cells })
    }

    /// Parse a math run group (`\mr`); `self.pos` is at its opening brace.
    ///
    /// Character formatting controls inside the run are passive and skipped;
    /// only the text, the `\mnor` flag, and `\mrPr` properties are retained.
    pub(super) fn parse_math_run(&mut self, depth: usize) -> RtfResult<crate::MathObject<'a>> {
        self.pos += 2; // opening brace and run control
        let mut properties = None;
        let mut normal_text = false;
        let mut text = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    if properties.is_some()
                        || !matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::MathRunProperties))
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF math run contains an unsupported group".to_string(),
                        ));
                    }
                    properties = Some(
                        self.parse_math_properties_group(crate::MathPropertiesKind::Run, depth)?,
                    );
                },
                Some(Token::Control(ControlWord::MathPropertyNormalText(_))) => {
                    normal_text = true;
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math run contains binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut text,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        "run",
                    )? {
                        // Passive character formatting is not retained.
                        self.pos += 1;
                    }
                    if text.len() > crate::math::MAX_MATH_RUN_TEXT_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF math run text exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        self.math_text_bytes = self.math_text_bytes.saturating_add(text.len());
        if self.math_text_bytes > crate::math::MAX_MATH_TOTAL_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF math aggregate text exceeds the safety limit".to_string(),
            ));
        }
        Ok(crate::MathObject::Run(crate::MathRun {
            properties,
            normal_text,
            text: Cow::Owned(text),
        }))
    }

    /// Parse a math property destination (`\mfPr`, `\mctrlPr`, ...);
    /// `self.pos` is at its opening brace.
    pub(super) fn parse_math_properties_group(
        &mut self,
        kind: crate::MathPropertiesKind,
        depth: usize,
    ) -> RtfResult<crate::MathProperties<'a>> {
        if depth > crate::math::MAX_MATH_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF math nesting depth exceeds the safety limit".to_string(),
            ));
        }
        self.pos += 2; // opening brace and property-destination control
        let mut properties = Vec::new();
        let mut matrix_columns = Vec::new();
        let mut saw_matrix_columns = false;
        let mut control = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => match self.tokens.get(self.pos + 1) {
                    Some(Token::Control(ControlWord::MathControlProperties)) => {
                        if control.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF math destination contains multiple control properties"
                                    .to_string(),
                            ));
                        }
                        control = Some(Box::new(self.parse_math_properties_group(
                            crate::MathPropertiesKind::Control,
                            depth + 1,
                        )?));
                    },
                    Some(Token::Control(ControlWord::MathMatrixColumns)) => {
                        if kind
                            != crate::MathPropertiesKind::Structure(
                                crate::MathStructureKind::Matrix,
                            )
                            || saw_matrix_columns
                        {
                            return Err(RtfError::MalformedDocument(
                                "RTF math matrix columns are misplaced".to_string(),
                            ));
                        }
                        saw_matrix_columns = true;
                        matrix_columns = self.parse_math_matrix_columns(depth)?;
                    },
                    Some(Token::Control(property_control)) => {
                        let Some((name, param)) = Self::math_property_name(property_control) else {
                            return Err(RtfError::MalformedDocument(
                                "RTF math destination contains an unsupported property".to_string(),
                            ));
                        };
                        properties.push(self.parse_math_property(name, param)?);
                        if properties.len() > crate::math::MAX_MATH_PROPERTIES_PER_DESTINATION {
                            return Err(RtfError::MalformedDocument(
                                "RTF math property count exceeds the safety limit".to_string(),
                            ));
                        }
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF math property destination contains unsupported grouped content"
                                .to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math property destination contains ungrouped, binary, or active data"
                            .to_string(),
                    ));
                },
            }
        }
        let destination = crate::MathProperties {
            kind,
            properties,
            matrix_columns,
            control,
        };
        destination.validate()?;
        Ok(destination)
    }

    /// Parse the matrix columns destination (`\mmcs`) of a matrix property
    /// destination; `self.pos` is at its opening brace.
    pub(super) fn parse_math_matrix_columns(
        &mut self,
        depth: usize,
    ) -> RtfResult<Vec<crate::MathMatrixColumn<'a>>> {
        self.pos += 2; // opening brace and mmcs control
        let mut columns = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    if !matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::MathMatrixColumn))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF math matrix columns may contain only mmc destinations".to_string(),
                        ));
                    }
                    columns.push(self.parse_math_matrix_column(depth)?);
                    if columns.len() > crate::math::MAX_MATH_OBJECTS_PER_CONTAINER {
                        return Err(RtfError::MalformedDocument(
                            "RTF math matrix column count exceeds the safety limit".to_string(),
                        ));
                    }
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math matrix columns contain ungrouped, binary, or active data"
                            .to_string(),
                    ));
                },
            }
        }
        if columns.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF math matrix columns must contain at least one mmc destination".to_string(),
            ));
        }
        Ok(columns)
    }

    /// Parse one matrix column destination (`\mmc`); `self.pos` is at its
    /// opening brace.
    pub(super) fn parse_math_matrix_column(
        &mut self,
        depth: usize,
    ) -> RtfResult<crate::MathMatrixColumn<'a>> {
        self.pos += 2; // opening brace and mmc control
        let mut properties = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    if properties.is_some()
                        || !matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::MathMatrixColumnProperties))
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF math matrix column contains an unsupported group".to_string(),
                        ));
                    }
                    properties = Some(self.parse_math_properties_group(
                        crate::MathPropertiesKind::MatrixColumn,
                        depth,
                    )?);
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math matrix column contains ungrouped, binary, or active data"
                            .to_string(),
                    ));
                },
            }
        }
        let column = crate::MathMatrixColumn { properties };
        column.validate()?;
        Ok(column)
    }

    /// Parse one math property group; `self.pos` is at its opening brace and
    /// the property control word follows it.
    pub(super) fn parse_math_property(
        &mut self,
        name: crate::MathPropertyName,
        param: Option<i32>,
    ) -> RtfResult<crate::MathProperty<'a>> {
        self.pos += 2; // opening brace and property control
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        "property value",
                    )? {
                        return Err(RtfError::MalformedDocument(
                            "RTF math property contains grouped, binary, or active data"
                                .to_string(),
                        ));
                    }
                    if value.len() > crate::math::MAX_MATH_PROPERTY_VALUE_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF math property value exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        let mut value = value.trim().to_string();
        if let Some(param) = param {
            if !value.is_empty() {
                return Err(RtfError::MalformedDocument(
                    "RTF math property has conflicting parameter and text values".to_string(),
                ));
            }
            value = param.to_string();
        }
        self.math_text_bytes = self.math_text_bytes.saturating_add(value.len());
        if self.math_text_bytes > crate::math::MAX_MATH_TOTAL_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF math aggregate text exceeds the safety limit".to_string(),
            ));
        }
        crate::MathProperty::new(name, Cow::Owned(value))
    }
}
