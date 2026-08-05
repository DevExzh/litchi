use super::*;

impl<'a> Parser<'a> {
    pub(super) fn parse_paragraph_group_table(
        &mut self,
    ) -> RtfResult<crate::ParagraphGroupPropertyTable> {
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF pgptbl must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::ParagraphGroupTable))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF pgptbl destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut table = crate::ParagraphGroupPropertyTable::new();
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
                        Some(Token::Control(ControlWord::ParagraphGroup))
                    ) =>
                {
                    let id = u32::try_from(table.entries().len() + 1).map_err(|_| {
                        RtfError::MalformedDocument("RTF paragraph-group ID overflow".to_string())
                    })?;
                    let entry = self.parse_paragraph_group_property(id)?;
                    table.push(entry)?;
                    continue;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pgptbl cannot contain fields, objects, or unknown destinations"
                            .to_string(),
                    ));
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pgptbl destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_paragraph_group_property(
        &mut self,
        id: u32,
    ) -> RtfResult<crate::ParagraphGroupProperty> {
        self.pos += 2; // opening brace and pgp
        let mut parent_id = None;
        let mut nesting = None;
        let mut left = None;
        let mut right = None;
        let mut before = None;
        let mut after = None;
        let mut borders = crate::Borders::new();
        let mut current_border = None;
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let entry = crate::ParagraphGroupProperty {
                        id,
                        parent_id: u32::try_from(parent_id.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks ipgp".to_string())
                        })?)
                        .map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF ipgp reference".to_string())
                        })?,
                        table_nesting_level: u8::try_from(nesting.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks itap".to_string())
                        })?)
                        .map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF pgp itap value".to_string())
                        })?,
                        left_indent: left.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks li".to_string())
                        })?,
                        right_indent: right.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks ri".to_string())
                        })?,
                        space_before: before.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks sb".to_string())
                        })?,
                        space_after: after.ok_or_else(|| {
                            RtfError::MalformedDocument("RTF pgp entry lacks sa".to_string())
                        })?,
                        borders,
                    };
                    entry.validate()?;
                    return Ok(entry);
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ParagraphGroupParent(value) => {
                        if !seen.insert("ipgp") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp ipgp".to_string(),
                            ));
                        }
                        parent_id = Some(*value);
                    },
                    ControlWord::TableNestingLevel(value) => {
                        if !seen.insert("itap") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp itap".to_string(),
                            ));
                        }
                        nesting = Some(value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF pgp itap requires a numeric parameter".to_string(),
                            )
                        })?);
                    },
                    ControlWord::LeftIndent(value) => {
                        if !seen.insert("li") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp li".to_string(),
                            ));
                        }
                        left = Some(*value);
                    },
                    ControlWord::RightIndent(value) => {
                        if !seen.insert("ri") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp ri".to_string(),
                            ));
                        }
                        right = Some(*value);
                    },
                    ControlWord::SpaceBefore(value) => {
                        if !seen.insert("sb") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp sb".to_string(),
                            ));
                        }
                        before = Some(*value);
                    },
                    ControlWord::SpaceAfter(value) => {
                        if !seen.insert("sa") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp sa".to_string(),
                            ));
                        }
                        after = Some(*value);
                    },
                    ControlWord::BorderTop => {
                        if !seen.insert("brdrt") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp top border".to_string(),
                            ));
                        }
                        current_border = Some(0u8);
                    },
                    ControlWord::BorderBottom => {
                        if !seen.insert("brdrb") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp bottom border".to_string(),
                            ));
                        }
                        current_border = Some(1u8);
                    },
                    ControlWord::BorderLeft => {
                        if !seen.insert("brdrl") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp left border".to_string(),
                            ));
                        }
                        current_border = Some(2u8);
                    },
                    ControlWord::BorderRight => {
                        if !seen.insert("brdrr") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp right border".to_string(),
                            ));
                        }
                        current_border = Some(3u8);
                    },
                    ControlWord::BorderNone
                    | ControlWord::BorderSingle
                    | ControlWord::BorderDotted
                    | ControlWord::BorderDashed
                    | ControlWord::BorderDouble
                    | ControlWord::BorderTriple
                    | ControlWord::BorderWave => {
                        let border = match current_border {
                            Some(0) => &mut borders.top,
                            Some(1) => &mut borders.bottom,
                            Some(2) => &mut borders.left,
                            Some(3) => &mut borders.right,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF pgp border style has no side".to_string(),
                                ));
                            },
                        };
                        if border.style != crate::BorderStyle::None
                            && !matches!(control, ControlWord::BorderNone)
                        {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pgp border style".to_string(),
                            ));
                        }
                        border.style = match control {
                            ControlWord::BorderNone => crate::BorderStyle::None,
                            ControlWord::BorderSingle => crate::BorderStyle::Single,
                            ControlWord::BorderDotted => crate::BorderStyle::Dotted,
                            ControlWord::BorderDashed => crate::BorderStyle::Dashed,
                            ControlWord::BorderDouble => crate::BorderStyle::Double,
                            ControlWord::BorderTriple => crate::BorderStyle::Triple,
                            _ => crate::BorderStyle::Wavy,
                        };
                    },
                    ControlWord::BorderWidth(value) => {
                        let border = match current_border {
                            Some(0) => &mut borders.top,
                            Some(1) => &mut borders.bottom,
                            Some(2) => &mut borders.left,
                            Some(3) => &mut borders.right,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF pgp border width has no side".to_string(),
                                ));
                            },
                        };
                        border.width = value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF pgp brdrw requires a numeric parameter".to_string(),
                            )
                        })?;
                    },
                    ControlWord::BorderColor(value) => {
                        let border = match current_border {
                            Some(0) => &mut borders.top,
                            Some(1) => &mut borders.bottom,
                            Some(2) => &mut borders.left,
                            Some(3) => &mut borders.right,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF pgp border color has no side".to_string(),
                                ));
                            },
                        };
                        border.color_ref = u16::try_from(value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF pgp brdrcf requires a numeric parameter".to_string(),
                            )
                        })?)
                        .map_err(|_| {
                            RtfError::MalformedDocument("invalid RTF pgp border color".to_string())
                        })?;
                    },
                    ControlWord::BorderSpace(value) => {
                        let border = match current_border {
                            Some(0) => &mut borders.top,
                            Some(1) => &mut borders.bottom,
                            Some(2) => &mut borders.left,
                            Some(3) => &mut borders.right,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF pgp border space has no side".to_string(),
                                ));
                            },
                        };
                        border.space = value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF pgp brsp requires a numeric parameter".to_string(),
                            )
                        })?;
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "unsupported control in RTF pgp entry".to_string(),
                        ));
                    },
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pgp entry cannot contain nested destinations".to_string(),
                    ));
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pgp entry".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_legacy_section_numbering_level(&mut self) -> RtfResult<()> {
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl destinations must occur at document scope before body text"
                    .to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        let level_index = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacySectionNumberingLevel(value))) => {
                u8::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF pnseclvl index must be between 1 and 9".to_string(),
                    )
                })?
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF pnseclvl destination".to_string(),
                ));
            },
        };
        self.pos += 1;
        let mut format = None;
        let mut start_at = None;
        let mut indent = None;
        let mut space = None;
        let mut hanging = false;
        let mut previous = false;
        let mut alignment = None;
        let mut font_ref = None;
        let mut text_before = String::new();
        let mut text_after = String::new();
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let format = format.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF pnseclvl destination has no numbering format".to_string(),
                        )
                    })?;
                    let mut level = crate::LegacySectionNumberingLevel::new(level_index, format);
                    level.start_at = start_at;
                    level.indent = indent;
                    level.space = space;
                    level.hanging = hanging;
                    level.previous = previous;
                    level.alignment = alignment;
                    level.font_ref = font_ref;
                    level.text_before = Cow::Owned(text_before);
                    level.text_after = Cow::Owned(text_after);
                    return self.legacy_section_numbering.add(level);
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextBefore))
                    ) =>
                {
                    if !seen.insert("text-before") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF pntxtb destination".to_string(),
                        ));
                    }
                    text_before =
                        self.parse_legacy_numbering_text(ControlWord::LegacyNumberingTextBefore)?;
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextAfter))
                    ) =>
                {
                    if !seen.insert("text-after") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF pntxta destination".to_string(),
                        ));
                    }
                    text_after =
                        self.parse_legacy_numbering_text(ControlWord::LegacyNumberingTextAfter)?;
                    continue;
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pnseclvl cannot contain nested fields, objects, or destinations"
                            .to_string(),
                    ));
                },
                Some(Token::Control(control)) => {
                    let (key, new_format) = match control {
                        ControlWord::LegacyNumberingDecimal(param) => {
                            require_parameterless(*param, "pndec")?;
                            ("format", Some(crate::LegacyNumberingFormat::Decimal))
                        },
                        ControlWord::LegacyNumberingUpperRoman(param) => {
                            require_parameterless(*param, "pnucrm")?;
                            ("format", Some(crate::LegacyNumberingFormat::UpperRoman))
                        },
                        ControlWord::LegacyNumberingLowerRoman(param) => {
                            require_parameterless(*param, "pnlcrm")?;
                            ("format", Some(crate::LegacyNumberingFormat::LowerRoman))
                        },
                        ControlWord::LegacyNumberingUpperLetter(param) => {
                            require_parameterless(*param, "pnucltr")?;
                            ("format", Some(crate::LegacyNumberingFormat::UpperLetter))
                        },
                        ControlWord::LegacyNumberingLowerLetter(param) => {
                            require_parameterless(*param, "pnlcltr")?;
                            ("format", Some(crate::LegacyNumberingFormat::LowerLetter))
                        },
                        ControlWord::LegacyNumberingStart(value) => {
                            if !seen.insert("start") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnstart".to_string(),
                                ));
                            }
                            start_at = Some(value.ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF pnstart requires a numeric parameter".to_string(),
                                )
                            })?);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingIndent(value) => {
                            if !seen.insert("indent") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnindent".to_string(),
                                ));
                            }
                            indent = Some(value.ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF pnindent requires a numeric parameter".to_string(),
                                )
                            })?);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingSpace(value) => {
                            if !seen.insert("space") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnsp".to_string(),
                                ));
                            }
                            space = Some(value.ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF pnsp requires a numeric parameter".to_string(),
                                )
                            })?);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingHanging(param) => {
                            require_parameterless(*param, "pnhang")?;
                            if !seen.insert("hanging") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnhang".to_string(),
                                ));
                            }
                            hanging = true;
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingPrevious(param) => {
                            require_parameterless(*param, "pnprev")?;
                            if !seen.insert("previous") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnprev".to_string(),
                                ));
                            }
                            previous = true;
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingAlignLeft(param) => {
                            require_parameterless(*param, "pnql")?;
                            ("alignment", None)
                        },
                        ControlWord::LegacyNumberingAlignCenter(param) => {
                            require_parameterless(*param, "pnqc")?;
                            ("alignment-center", None)
                        },
                        ControlWord::LegacyNumberingAlignRight(param) => {
                            require_parameterless(*param, "pnqr")?;
                            ("alignment-right", None)
                        },
                        ControlWord::LegacyNumberingFont(value) => {
                            if !seen.insert("font") {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF pnf".to_string(),
                                ));
                            }
                            font_ref = Some(
                                u16::try_from(value.ok_or_else(|| {
                                    RtfError::MalformedDocument(
                                        "RTF pnf requires a numeric parameter".to_string(),
                                    )
                                })?)
                                .map_err(|_| {
                                    RtfError::MalformedDocument(
                                        "invalid RTF pnf reference".to_string(),
                                    )
                                })?,
                            );
                            self.pos += 1;
                            continue;
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "unsupported control in RTF pnseclvl destination".to_string(),
                            ));
                        },
                    };
                    if key == "format" {
                        if !seen.insert(key) {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pnseclvl numbering format".to_string(),
                            ));
                        }
                        format = new_format;
                    } else {
                        if !seen.insert("alignment") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pnseclvl alignment".to_string(),
                            ));
                        }
                        alignment = Some(match control {
                            ControlWord::LegacyNumberingAlignCenter(_) => {
                                crate::LegacyNumberingAlignment::Center
                            },
                            ControlWord::LegacyNumberingAlignRight(_) => {
                                crate::LegacyNumberingAlignment::Right
                            },
                            _ => crate::LegacyNumberingAlignment::Left,
                        });
                    }
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pnseclvl destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_legacy_paragraph_numbering(&mut self) -> RtfResult<()> {
        let parameter = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacyParagraphNumbering(parameter))) => *parameter,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF pn destination".to_string(),
                ));
            },
        };
        require_parameterless(parameter, "pn")?;
        let parent = self.states.len().checked_sub(2).ok_or_else(|| {
            RtfError::MalformedDocument("RTF pn destination has no paragraph owner".to_string())
        })?;
        let owner = self.states.get(parent).ok_or_else(|| {
            RtfError::ParserError("RTF pn paragraph owner state is missing".to_string())
        })?;
        if owner.destination != Destination::DocumentBody
            || owner.in_table
            || owner.table_nesting_level >= 2
        {
            return Err(RtfError::MalformedDocument(
                "RTF pn destination must belong to a non-table paragraph".to_string(),
            ));
        }
        if owner.paragraph_content_started {
            return Err(RtfError::MalformedDocument(
                "RTF pn destination must precede paragraph content".to_string(),
            ));
        }
        if owner.paragraph_numbering_declared {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF pn destination in one paragraph".to_string(),
            ));
        }
        if self.legacy_paragraph_numbering.len() >= crate::MAX_LEGACY_PARAGRAPH_NUMBERING_RECORDS {
            return Err(RtfError::MalformedDocument(
                "RTF contains too many pn destinations".to_string(),
            ));
        }
        self.pos += 1;
        let mut level = None;
        let mut format = None;
        let mut alignment = None;
        let mut start_at = None;
        let mut indent = None;
        let mut space = None;
        let mut across = false;
        let mut number_once = false;
        let mut previous = false;
        let mut restart = false;
        let mut hanging = false;
        let mut bidi = None;
        let mut font_ref = None;
        let mut font_size = None;
        let mut color_ref = None;
        let mut bold = None;
        let mut italic = None;
        let mut caps = None;
        let mut small_caps = None;
        let mut strike = None;
        let mut underline = None;
        let mut text_before = None;
        let mut text_after = None;
        let mut revision = crate::LegacyParagraphNumberingRevision::default();
        let mut seen = std::collections::HashSet::new();
        macro_rules! once {
            ($key:expr, $name:expr) => {
                if !seen.insert($key) {
                    return Err(RtfError::MalformedDocument(format!(
                        "duplicate RTF {} in pn destination",
                        $name
                    )));
                }
            };
        }
        fn value(parameter: Option<i32>, name: &str) -> RtfResult<i32> {
            parameter.ok_or_else(|| {
                RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
            })
        }
        fn toggle(parameter: Option<i32>, name: &str) -> RtfResult<bool> {
            match parameter {
                None | Some(1) => Ok(true),
                Some(0) => Ok(false),
                Some(_) => Err(RtfError::MalformedDocument(format!(
                    "RTF {name} accepts only 0 or 1"
                ))),
            }
        }
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let mut record =
                        crate::LegacyParagraphNumbering::new(level.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF pn destination is missing its pnlvl selector".to_string(),
                            )
                        })?);
                    record.format = format;
                    record.alignment = alignment;
                    record.start_at = start_at;
                    record.indent = indent;
                    record.space = space;
                    record.across = across;
                    record.number_once = number_once;
                    record.previous = previous;
                    record.restart = restart;
                    record.hanging = hanging;
                    record.bidi = bidi;
                    record.font_ref = font_ref;
                    record.font_size = font_size;
                    record.color_ref = color_ref;
                    record.bold = bold;
                    record.italic = italic;
                    record.caps = caps;
                    record.small_caps = small_caps;
                    record.strike = strike;
                    record.underline = underline;
                    record.text_before = text_before.map(Cow::Owned);
                    record.text_after = text_after.map(Cow::Owned);
                    record.revision = revision;
                    record.validate()?;
                    let index =
                        u32::try_from(self.legacy_paragraph_numbering.len()).map_err(|_| {
                            RtfError::MalformedDocument("RTF pn record index overflow".to_string())
                        })?;
                    self.legacy_paragraph_numbering.push(record);
                    let owner = self.states.get_mut(parent).ok_or_else(|| {
                        RtfError::ParserError("RTF pn paragraph owner state is missing".to_string())
                    })?;
                    owner.paragraph.legacy_numbering = Some(index);
                    owner.paragraph_numbering_declared = true;
                    if let Some(state) = self.states.last_mut() {
                        state.paragraph.legacy_numbering = Some(index);
                        state.paragraph_numbering_declared = true;
                    }
                    return Ok(());
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextBefore))
                    ) =>
                {
                    once!("text-before", "pntxtb");
                    text_before = Some(
                        self.parse_legacy_numbering_text(ControlWord::LegacyNumberingTextBefore)?,
                    );
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextAfter))
                    ) =>
                {
                    once!("text-after", "pntxta");
                    text_after = Some(
                        self.parse_legacy_numbering_text(ControlWord::LegacyNumberingTextAfter)?,
                    );
                    continue;
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pn destination cannot contain nested active content".to_string(),
                    ));
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::Control(control)) => match control {
                    ControlWord::LegacyNumberingLevel(parameter) => {
                        once!("level", "pnlvl");
                        let v = value(*parameter, "pnlvl")?;
                        level = Some(crate::LegacyParagraphNumberingLevel::Explicit(
                            u8::try_from(v).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pnlvl value must be in 1..=9".to_string(),
                                )
                            })?,
                        ));
                    },
                    ControlWord::LegacyNumberingLevelBullet(parameter) => {
                        require_parameterless(*parameter, "pnlvlblt")?;
                        once!("level", "pnlvlblt");
                        level = Some(crate::LegacyParagraphNumberingLevel::Bullet);
                    },
                    ControlWord::LegacyNumberingLevelBody(parameter) => {
                        require_parameterless(*parameter, "pnlvlbody")?;
                        once!("level", "pnlvlbody");
                        level = Some(crate::LegacyParagraphNumberingLevel::Body);
                    },
                    ControlWord::LegacyNumberingLevelContinue(parameter) => {
                        require_parameterless(*parameter, "pnlvlcont")?;
                        once!("level", "pnlvlcont");
                        level = Some(crate::LegacyParagraphNumberingLevel::Continue);
                    },
                    ControlWord::LegacyNumberingDecimal(parameter) => {
                        require_parameterless(*parameter, "pndec")?;
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::Decimal);
                    },
                    ControlWord::LegacyNumberingUpperRoman(parameter) => {
                        require_parameterless(*parameter, "pnucrm")?;
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::UpperRoman);
                    },
                    ControlWord::LegacyNumberingLowerRoman(parameter) => {
                        require_parameterless(*parameter, "pnlcrm")?;
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::LowerRoman);
                    },
                    ControlWord::LegacyNumberingUpperLetter(parameter) => {
                        require_parameterless(*parameter, "pnucltr")?;
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::UpperLetter);
                    },
                    ControlWord::LegacyNumberingLowerLetter(parameter) => {
                        require_parameterless(*parameter, "pnlcltr")?;
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::LowerLetter);
                    },
                    ControlWord::LegacyNumberingFormat(kind, parameter) => {
                        require_parameterless(*parameter, "pn numbering-format selector")?;
                        once!("format", "numbering format");
                        format = Some(*kind);
                    },
                    ControlWord::Pngblip => {
                        once!("format", "numbering format");
                        format = Some(crate::LegacyParagraphNumberingFormat::GbLip);
                    },
                    ControlWord::LegacyNumberingAlignLeft(parameter)
                    | ControlWord::LegacyNumberingAlignCenter(parameter)
                    | ControlWord::LegacyNumberingAlignRight(parameter) => {
                        require_parameterless(*parameter, "pn alignment")?;
                        once!("alignment", "alignment");
                        alignment = Some(match control {
                            ControlWord::LegacyNumberingAlignCenter(_) => {
                                crate::LegacyParagraphNumberingAlignment::Center
                            },
                            ControlWord::LegacyNumberingAlignRight(_) => {
                                crate::LegacyParagraphNumberingAlignment::Right
                            },
                            _ => crate::LegacyParagraphNumberingAlignment::Left,
                        });
                    },
                    ControlWord::LegacyNumberingStart(parameter) => {
                        once!("start", "pnstart");
                        start_at = Some(value(*parameter, "pnstart")?);
                    },
                    ControlWord::LegacyNumberingIndent(parameter) => {
                        once!("indent", "pnindent");
                        indent = Some(value(*parameter, "pnindent")?);
                    },
                    ControlWord::LegacyNumberingSpace(parameter) => {
                        once!("space", "pnsp");
                        space = Some(value(*parameter, "pnsp")?);
                    },
                    ControlWord::LegacyNumberingFont(parameter) => {
                        once!("font", "pnf");
                        font_ref =
                            Some(u16::try_from(value(*parameter, "pnf")?).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pnf value must be in 0..=65535".to_string(),
                                )
                            })?);
                    },
                    ControlWord::LegacyNumberingFontSize(parameter) => {
                        once!("font-size", "pnfs");
                        font_size =
                            Some(u16::try_from(value(*parameter, "pnfs")?).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pnfs value must be in 1..=65535".to_string(),
                                )
                            })?);
                    },
                    ControlWord::LegacyNumberingColor(parameter) => {
                        once!("color", "pncf");
                        color_ref =
                            Some(u16::try_from(value(*parameter, "pncf")?).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pncf value must be in 0..=65535".to_string(),
                                )
                            })?);
                    },
                    ControlWord::LegacyNumberingAcross(parameter)
                    | ControlWord::LegacyNumberingOnce(parameter)
                    | ControlWord::LegacyNumberingPrevious(parameter)
                    | ControlWord::LegacyNumberingRestart(parameter)
                    | ControlWord::LegacyNumberingHanging(parameter) => {
                        require_parameterless(*parameter, "pn flag")?;
                        let (key, target) = match control {
                            ControlWord::LegacyNumberingAcross(_) => ("across", &mut across),
                            ControlWord::LegacyNumberingOnce(_) => ("once", &mut number_once),
                            ControlWord::LegacyNumberingPrevious(_) => ("previous", &mut previous),
                            ControlWord::LegacyNumberingRestart(_) => ("restart", &mut restart),
                            _ => ("hanging", &mut hanging),
                        };
                        once!(key, key);
                        *target = true;
                    },
                    ControlWord::LegacyNumberingBidiA(parameter)
                    | ControlWord::LegacyNumberingBidiB(parameter) => {
                        require_parameterless(*parameter, "pn bidi selector")?;
                        once!("bidi", "bidi selector");
                        bidi = Some(if matches!(control, ControlWord::LegacyNumberingBidiA(_)) {
                            crate::LegacyParagraphNumberingBidi::A
                        } else {
                            crate::LegacyParagraphNumberingBidi::B
                        });
                    },
                    ControlWord::LegacyNumberingBold(parameter)
                    | ControlWord::LegacyNumberingItalic(parameter)
                    | ControlWord::LegacyNumberingCaps(parameter)
                    | ControlWord::LegacyNumberingSmallCaps(parameter)
                    | ControlWord::LegacyNumberingStrike(parameter) => {
                        let (key, target) = match control {
                            ControlWord::LegacyNumberingBold(_) => ("bold", &mut bold),
                            ControlWord::LegacyNumberingItalic(_) => ("italic", &mut italic),
                            ControlWord::LegacyNumberingCaps(_) => ("caps", &mut caps),
                            ControlWord::LegacyNumberingSmallCaps(_) => {
                                ("small-caps", &mut small_caps)
                            },
                            _ => ("strike", &mut strike),
                        };
                        once!(key, key);
                        *target = Some(toggle(*parameter, key)?);
                    },
                    ControlWord::LegacyNumberingUnderlineToggle(parameter) => {
                        once!("underline", "underline");
                        underline = Some(if toggle(*parameter, "pnul")? {
                            crate::LegacyParagraphNumberingUnderline::Single
                        } else {
                            crate::LegacyParagraphNumberingUnderline::None
                        });
                    },
                    ControlWord::LegacyNumberingUnderline(kind, parameter) => {
                        require_parameterless(*parameter, "pn underline selector")?;
                        once!("underline", "underline");
                        underline = Some(*kind);
                    },
                    ControlWord::LegacyNumberingRevisionAuthor(parameter) => {
                        once!("revision-author", "pnrauth");
                        revision.author =
                            Some(u16::try_from(value(*parameter, "pnrauth")?).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pnrauth value must be in 0..=65535".to_string(),
                                )
                            })?);
                    },
                    ControlWord::LegacyNumberingRevisionDate(parameter) => {
                        once!("revision-date", "pnrdate");
                        revision.date = Some(value(*parameter, "pnrdate")?);
                    },
                    ControlWord::LegacyNumberingRevisionFormat(parameter) => {
                        once!("revision-format", "pnrnfc");
                        revision.number_format = Some(value(*parameter, "pnrnfc")?);
                    },
                    ControlWord::LegacyNumberingRevisionNoTrack(parameter) => {
                        require_parameterless(*parameter, "pnrnot")?;
                        once!("revision-no-track", "pnrnot");
                        revision.no_tracking = true;
                    },
                    ControlWord::LegacyNumberingRevisionParagraph(parameter) => {
                        once!("revision-paragraph", "pnrpnbr");
                        revision.paragraph_number = Some(value(*parameter, "pnrpnbr")?);
                    },
                    ControlWord::LegacyNumberingRevisionRgb(parameter) => {
                        once!("revision-rgb", "pnrrgb");
                        revision.rgb =
                            Some(u32::try_from(value(*parameter, "pnrrgb")?).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF pnrrgb must be non-negative".to_string(),
                                )
                            })?);
                    },
                    ControlWord::LegacyNumberingRevisionStart(parameter) => {
                        once!("revision-start", "pnrstart");
                        revision.start = Some(value(*parameter, "pnrstart")?);
                    },
                    ControlWord::LegacyNumberingRevisionStop(parameter) => {
                        once!("revision-stop", "pnrstop");
                        revision.stop = Some(value(*parameter, "pnrstop")?);
                    },
                    ControlWord::LegacyNumberingRevisionTextStart(parameter) => {
                        once!("revision-text-start", "pnrxst");
                        revision.text_start = Some(value(*parameter, "pnrxst")?);
                    },
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "unsupported control in RTF pn destination".to_string(),
                        ));
                    },
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pn destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
    }

    pub(super) fn parse_legacy_numbering_text(
        &mut self,
        expected: ControlWord<'a>,
    ) -> RtfResult<String> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || self.tokens.get(self.pos + 1) != Some(&Token::Control(expected))
        {
            return Err(RtfError::MalformedDocument(
                "invalid RTF legacy-numbering text destination".to_string(),
            ));
        }
        self.pos += 2;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(value);
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy-numbering text cannot contain nested destinations".to_string(),
                    ));
                },
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                // Character-type selectors affect only which run font decodes
                // the following text. `dbch` is emitted by Word in pnseclvl
                // punctuation destinations and carries no textual payload.
                Some(Token::Control(
                    ControlWord::Unknown("dbch", None) | ControlWord::DoubleByteCharacter(None),
                )) => {},
                Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy-numbering text contains a non-text control".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if value.len() > crate::legacy_numbering::MAX_LEGACY_NUMBERING_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF pnseclvl text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }
}
