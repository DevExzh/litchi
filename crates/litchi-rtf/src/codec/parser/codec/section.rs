use super::*;

impl<'a> Parser<'a> {
    pub(super) fn apply_review_display_control(
        &mut self,
        control: &ControlWord<'_>,
    ) -> RtfResult<()> {
        let (bit, name, parameter) = match control {
            ControlWord::HideReviewMarkup(parameter) => (1, "donotshowmarkup", parameter),
            ControlWord::HideReviewComments(parameter) => (2, "donotshowcomments", parameter),
            ControlWord::HideReviewInsertionsAndDeletions(parameter) => {
                (4, "donotshowinsdel", parameter)
            },
            _ => return Ok(()),
        };
        if parameter.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {name} must not have a numeric parameter"
            )));
        }
        if self.review_display_seen & bit != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF {name} document property"
            )));
        }
        self.review_display_seen |= bit;
        match control {
            ControlWord::HideReviewMarkup(_) => self.review_display.hide_markup = true,
            ControlWord::HideReviewComments(_) => self.review_display.hide_comments = true,
            ControlWord::HideReviewInsertionsAndDeletions(_) => {
                self.review_display.hide_insertions_and_deletions = true;
            },
            _ => return Err(parser_classification_error()),
        }
        Ok(())
    }

    pub(super) fn apply_document_view_control(
        &mut self,
        control: &ControlWord<'_>,
    ) -> RtfResult<()> {
        let (bit, name) = match control {
            ControlWord::DocumentViewKind(_) => (1, "viewkind"),
            ControlWord::DocumentViewScale(_) => (2, "viewscale"),
            ControlWord::DocumentZoomKind(_) => (4, "viewzk"),
            ControlWord::DocumentViewBackgroundShapes(_) => (8, "viewbksp"),
            ControlWord::DocumentViewNoPageBoundaries(_) => (16, "viewnobound"),
            _ => return Ok(()),
        };
        if self.document_view_seen & bit != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF {name} document property"
            )));
        }
        self.document_view_seen |= bit;
        let value = match control {
            ControlWord::DocumentViewKind(value)
            | ControlWord::DocumentViewScale(value)
            | ControlWord::DocumentZoomKind(value)
            | ControlWord::DocumentViewBackgroundShapes(value) => value.ok_or_else(|| {
                RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
            })?,
            ControlWord::DocumentViewNoPageBoundaries(value) => {
                if value.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF viewnobound must not have a numeric parameter".to_string(),
                    ));
                }
                0
            },
            _ => return Err(parser_classification_error()),
        };
        match control {
            ControlWord::DocumentViewKind(_) => {
                self.document_view.kind = Some(crate::DocumentViewKind::from_rtf(value)?);
            },
            ControlWord::DocumentViewScale(_) => {
                let value = u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument(format!(
                        "RTF viewscale must be in 1..={}",
                        crate::MAX_DOCUMENT_VIEW_SCALE_PERCENT
                    ))
                })?;
                if value == 0 || value > crate::MAX_DOCUMENT_VIEW_SCALE_PERCENT {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF viewscale must be in 1..={}",
                        crate::MAX_DOCUMENT_VIEW_SCALE_PERCENT
                    )));
                }
                self.document_view.scale_percent = Some(value);
            },
            ControlWord::DocumentZoomKind(_) => {
                self.document_view.zoom_kind = Some(crate::DocumentZoomKind::from_rtf(value)?);
            },
            ControlWord::DocumentViewBackgroundShapes(_) => {
                self.document_view.background_shapes = Some(match value {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF viewbksp accepts only 0 or 1".to_string(),
                        ));
                    },
                });
            },
            ControlWord::DocumentViewNoPageBoundaries(_) => {
                self.document_view.hide_page_boundaries = true;
            },
            _ => return Err(parser_classification_error()),
        }
        Ok(())
    }

    pub(super) fn apply_document_hyphenation_control(
        &mut self,
        control: &ControlWord<'_>,
    ) -> RtfResult<()> {
        let (bit, name) = match control {
            ControlWord::HyphenateAutomatically(_) => (1, "hyphauto"),
            ControlWord::HyphenateCapitalizedWords(_) => (2, "hyphcaps"),
            ControlWord::HyphenationConsecutiveLines(_) => (4, "hyphconsec"),
            ControlWord::HyphenationHotZone(_) => (8, "hyphhotz"),
            _ => return Ok(()),
        };
        if self.hyphenation_seen & bit != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF {name} document property"
            )));
        }
        self.hyphenation_seen |= bit;
        let strict_toggle = |value: Option<i32>| match value {
            None | Some(1) => Ok(true),
            Some(0) => Ok(false),
            Some(_) => Err(RtfError::MalformedDocument(format!(
                "RTF {name} accepts only 0 or 1"
            ))),
        };
        match control {
            ControlWord::HyphenateAutomatically(value) => {
                self.hyphenation.automatic = Some(strict_toggle(*value)?);
            },
            ControlWord::HyphenateCapitalizedWords(value) => {
                self.hyphenation.capitalized_words = Some(strict_toggle(*value)?);
            },
            ControlWord::HyphenationConsecutiveLines(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF hyphconsec requires a numeric parameter".to_string(),
                    )
                })?;
                self.hyphenation.consecutive_line_limit =
                    Some(u32::try_from(value).map_err(|_| {
                        RtfError::MalformedDocument(format!(
                            "RTF hyphconsec must be in 0..={}",
                            crate::MAX_HYPHENATION_CONSECUTIVE_LINES
                        ))
                    })?);
            },
            ControlWord::HyphenationHotZone(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF hyphhotz requires a numeric parameter".to_string(),
                    )
                })?;
                let value = u32::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument("RTF hyphhotz cannot be negative".to_string())
                })?;
                if value > crate::MAX_HYPHENATION_HOT_ZONE_TWIPS {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF hyphenation hot zone exceeds {} twips",
                        crate::MAX_HYPHENATION_HOT_ZONE_TWIPS
                    )));
                }
                self.hyphenation.hot_zone_twips = Some(value);
            },
            _ => {},
        }
        Ok(())
    }

    pub(super) fn bind_pending_section_break(&mut self, section_index: usize) {
        for event in self.body_story_events.iter_mut().rev() {
            if let ParsedBodyStoryEvent::Resolved(crate::BodyStoryEvent::SectionBreak(boundary)) =
                event
                && boundary.next_section.is_none()
            {
                boundary.next_section = Some(section_index);
                return;
            }
        }
    }

    pub(super) fn begin_section(&mut self) -> RtfResult<()> {
        if self.sections.len() >= MAX_SECTIONS {
            return Err(RtfError::MalformedDocument(
                "RTF section count exceeds the safety limit".to_string(),
            ));
        }
        let inherited = self
            .sections
            .last()
            .map(|section| section.properties.clone())
            .unwrap_or_default();
        let inherited_gutter_override = self
            .section_gutter_overrides
            .last()
            .copied()
            .unwrap_or(false);
        let mut section = super::super::super::section::Section::new();
        section.properties = inherited;
        if self.sections.is_empty() {
            section.properties.margin_gutter = self
                .print_layout_settings
                .document_gutter_twips
                .unwrap_or_default() as i32;
        }
        let section_index = self.sections.len();
        self.sections.push(section);
        self.section_gutter_overrides
            .push(inherited_gutter_override);
        self.bind_pending_section_break(section_index);
        self.section_properties_active = true;
        Ok(())
    }

    pub(super) fn apply_section_control(&mut self, control: &ControlWord<'_>) -> RtfResult<bool> {
        use super::super::super::section::{PageOrientation, SectionBreakType, VerticalAlignment};

        let is_body_scoped_section_control = matches!(
            control,
            ControlWord::LineNumbering(_)
                | ControlWord::LineNumberDistance(_)
                | ControlWord::LineNumberStart(_)
                | ControlWord::LineNumberRestartSection
                | ControlWord::LineNumberRestartPage
                | ControlWord::LineNumberContinuous
                | ControlWord::PageNumberHeadingLevel(_)
                | ControlWord::PageNumberHeadingSeparator(_)
                | ControlWord::SectionLineGrid(_)
                | ControlWord::SectionDocumentGrid(_)
        );
        if is_body_scoped_section_control
            && self.current_state()?.destination != Destination::DocumentBody
        {
            return Ok(true);
        }
        let in_root_document_body = self.states.len() == 2
            && self
                .states
                .last()
                .is_some_and(|state| state.destination == Destination::DocumentBody);

        if let Some(side) = match control {
            ControlWord::PageBorderTop => Some(crate::PageBorderSide::Top),
            ControlWord::PageBorderLeft => Some(crate::PageBorderSide::Left),
            ControlWord::PageBorderBottom => Some(crate::PageBorderSide::Bottom),
            ControlWord::PageBorderRight => Some(crate::PageBorderSide::Right),
            _ => None,
        } {
            let border = self.parse_page_border_run()?;
            if !self.section_properties_active {
                self.begin_section()?;
            }
            let section = self
                .sections
                .last_mut()
                .ok_or_else(|| RtfError::MalformedDocument("no active RTF section".to_string()))?;
            if section.properties.page_borders.get(side).is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF page-border edge".to_string(),
                ));
            }
            section.properties.page_borders.set(side, border);
            self.section_properties_active = true;
            return Ok(true);
        }

        if matches!(control, ControlWord::Section) {
            if in_root_document_body {
                self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                    crate::BodyStoryEvent::SectionBreak(crate::SectionBreak::new(
                        self.body_text_len,
                        None,
                    )),
                ));
            }
            self.current_state_mut()?.section_column_number = None;
            self.section_properties_active = false;
            self.section_note_options_closed = false;
            self.root_section_format_run = false;
            return Ok(true);
        }
        let is_section_note_control = matches!(
            control,
            ControlWord::SectionFootnotePlacement(_)
                | ControlWord::SectionEndnoteHere
                | ControlWord::SectionFootnoteStart(_)
                | ControlWord::SectionEndnoteStart(_)
                | ControlWord::SectionFootnoteRestart(_)
                | ControlWord::SectionEndnoteRestart(_)
                | ControlWord::SectionFootnoteNumbering(_)
                | ControlWord::SectionEndnoteNumbering(_)
        );
        if !is_section_control(control) {
            return Ok(false);
        }
        if matches!(control, ControlWord::SectionDefault) && in_root_document_body {
            self.root_section_format_run = true;
        } else if matches!(control, ControlWord::SectionBreak) {
            self.root_section_format_run = false;
        }
        let in_visible_section_format = self.states.last().is_some_and(|state| {
            state.destination == Destination::DocumentBody && state.visible_section_format
        });
        let in_root_section_prefix = self.states.len() == 2
            && !self.section_note_options_closed
            && self
                .sections
                .last()
                .is_none_or(|section| section.headers_footers.is_empty());
        let in_root_section_format_run = in_root_document_body && self.root_section_format_run;
        if is_section_note_control
            && !in_root_section_prefix
            && !in_visible_section_format
            && !in_root_section_format_run
        {
            return Err(RtfError::MalformedDocument(
                "RTF section note options must precede section content at document root"
                    .to_string(),
            ));
        }

        if !self.section_properties_active {
            self.begin_section()?;
        }
        if matches!(control, ControlWord::SectionDefault) {
            if let Some(overridden) = self.section_gutter_overrides.last_mut() {
                *overridden = false;
            }
        } else if matches!(control, ControlWord::MarginGutter(_))
            && let Some(overridden) = self.section_gutter_overrides.last_mut()
        {
            *overridden = true;
        }
        let section = self.sections.last_mut().ok_or_else(|| {
            RtfError::ParserError("failed to create RTF section state".to_string())
        })?;
        let properties = &mut section.properties;
        match control {
            ControlWord::SectionStyle(value) => {
                properties.section_style = Some(section_style_reference(*value)?);
            },
            ControlWord::SectionRsid(value) => {
                properties.section_rsid = Some(*value as u32);
            },
            ControlWord::TitlePage => properties.title_page = true,
            ControlWord::SectionEndnoteHere => properties.note_options.endnote_here = true,
            ControlWord::SectionDefault => {
                *properties = super::super::super::section::SectionProperties::default();
                properties.margin_gutter = self
                    .print_layout_settings
                    .document_gutter_twips
                    .unwrap_or_default() as i32;
                self.states
                    .last_mut()
                    .ok_or_else(|| RtfError::ParserError("missing RTF parser state".to_string()))?
                    .section_column_number = None;
            },
            ControlWord::PageBorderOptions(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF pgbrdropt requires a numeric parameter".to_string(),
                    )
                })?;
                properties.page_borders.set_option_value(value)?;
            },
            ControlWord::PageBorderSurroundHeader => properties.page_borders.surround_header = true,
            ControlWord::PageBorderSurroundFooter => properties.page_borders.surround_footer = true,
            ControlWord::PageBorderSnap => properties.page_borders.snap_to_text_borders = true,
            ControlWord::SectionFootnotePlacement(value) => {
                properties.note_options.footnote_placement = Some(*value);
            },
            ControlWord::SectionFootnoteStart(value) => {
                if *value <= 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF section footnote starting number must be positive".to_string(),
                    ));
                }
                properties.note_options.footnote_start = Some(*value);
            },
            ControlWord::SectionEndnoteStart(value) => {
                if *value <= 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF section endnote starting number must be positive".to_string(),
                    ));
                }
                properties.note_options.endnote_start = Some(*value);
            },
            ControlWord::SectionFootnoteRestart(value) => {
                properties.note_options.footnote_restart = Some(*value);
            },
            ControlWord::SectionEndnoteRestart(value) => {
                properties.note_options.endnote_restart = Some(*value);
            },
            ControlWord::SectionFootnoteNumbering(value) => {
                properties.note_options.footnote_numbering = Some(*value);
            },
            ControlWord::SectionEndnoteNumbering(value) => {
                properties.note_options.endnote_numbering = Some(*value);
            },
            ControlWord::LeftToRightSection => {
                properties.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftSection => {
                properties.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::SectionBreak | ControlWord::SectionPage => {
                properties.break_type = SectionBreakType::Page;
            },
            ControlWord::SectionContinuous => {
                properties.break_type = SectionBreakType::Continuous;
            },
            ControlWord::SectionColumn => properties.break_type = SectionBreakType::Column,
            ControlWord::SectionEvenPage => properties.break_type = SectionBreakType::EvenPage,
            ControlWord::SectionOddPage => properties.break_type = SectionBreakType::OddPage,
            ControlWord::PageWidth(value) => properties.page_width = *value,
            ControlWord::PageHeight(value) => properties.page_height = *value,
            ControlWord::MarginLeft(value) => properties.margin_left = *value,
            ControlWord::MarginRight(value) => properties.margin_right = *value,
            ControlWord::MarginTop(value) => properties.margin_top = *value,
            ControlWord::MarginBottom(value) => properties.margin_bottom = *value,
            ControlWord::MarginGutter(value) => properties.margin_gutter = *value,
            ControlWord::PaperSourceFirst(value) => {
                properties.paper_source.first = Some(paper_source_bin(*value, "binfsxn")?);
            },
            ControlWord::PaperSourceOther(value) => {
                properties.paper_source.other = Some(paper_source_bin(*value, "binsxn")?);
            },
            ControlWord::HeaderDistance(value) => properties.header_distance = *value,
            ControlWord::FooterDistance(value) => properties.footer_distance = *value,
            ControlWord::Landscape => properties.orientation = PageOrientation::Landscape,
            ControlWord::Columns(value) => {
                let value = value.unwrap_or(1);
                let count = u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument(format!(
                        "RTF section-column count must be in 1..={}",
                        super::super::super::section::MAX_SECTION_COLUMNS
                    ))
                })?;
                if !(1..=super::super::super::section::MAX_SECTION_COLUMNS).contains(&count) {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF section-column count must be in 1..={}",
                        super::super::super::section::MAX_SECTION_COLUMNS
                    )));
                }
                properties.columns.count = count;
                properties.columns.explicit.clear();
                self.states
                    .last_mut()
                    .ok_or_else(|| RtfError::ParserError("missing RTF parser state".to_string()))?
                    .section_column_number = None;
            },
            ControlWord::ColumnSpace(value) => {
                let value = value.unwrap_or(720);
                if !(0..=super::super::super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column default spacing must be in 0..=31680 twips".to_string(),
                    ));
                }
                properties.columns.default_spacing = value;
            },
            ControlWord::ColumnSeparator(value) => properties.columns.separator = *value,
            ControlWord::ColumnNumber(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF colno requires a numeric parameter".to_string(),
                    )
                })?;
                let number = u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF colno must select an existing one-based section column".to_string(),
                    )
                })?;
                let expected =
                    u16::try_from(properties.columns.explicit.len() + 1).unwrap_or(u16::MAX);
                if number != expected || number > properties.columns.count {
                    return Err(RtfError::MalformedDocument(
                        "RTF explicit section columns must use sequential one-based colno values"
                            .to_string(),
                    ));
                }
                properties
                    .columns
                    .explicit
                    .push(super::super::super::section::SectionColumn {
                        width: 0,
                        space_after: None,
                    });
                self.states
                    .last_mut()
                    .ok_or_else(|| RtfError::ParserError("missing RTF parser state".to_string()))?
                    .section_column_number = Some(number);
            },
            ControlWord::ColumnWidth(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument("RTF colw requires a numeric parameter".to_string())
                })?;
                if !(1..=super::super::super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column width must be in 1..=31680 twips".to_string(),
                    ));
                }
                let number = self
                    .states
                    .last()
                    .and_then(|state| state.section_column_number)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF colw requires a preceding colno in the active group".to_string(),
                        )
                    })?;
                let column = properties
                    .columns
                    .explicit
                    .get_mut(usize::from(number - 1))
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF colw refers to an undefined section column".to_string(),
                        )
                    })?;
                if column.width != 0 {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF colw for one section column".to_string(),
                    ));
                }
                column.width = value;
            },
            ControlWord::ColumnSpaceRight(value) => {
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF colsr requires a numeric parameter".to_string(),
                    )
                })?;
                if !(0..=super::super::super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column spacing must be in 0..=31680 twips".to_string(),
                    ));
                }
                let number = self
                    .states
                    .last()
                    .and_then(|state| state.section_column_number)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF colsr requires a preceding colno in the active group".to_string(),
                        )
                    })?;
                let column = properties
                    .columns
                    .explicit
                    .get_mut(usize::from(number - 1))
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF colsr refers to an undefined section column".to_string(),
                        )
                    })?;
                if column.width == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF colsr must follow colw for its section column".to_string(),
                    ));
                }
                if column.space_after.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF colsr for one section column".to_string(),
                    ));
                }
            },
            ControlWord::PageNumberStart(value) => properties.page_number_start = *value,
            ControlWord::SectionVerticalRendering(param) => {
                require_parameterless(*param, "vertsect")?;
                properties.rendering = Some(crate::SectionRendering::Vertical);
            },
            ControlWord::SectionHorizontalRendering(param) => {
                require_parameterless(*param, "horzsect")?;
                properties.rendering = Some(crate::SectionRendering::Horizontal);
            },
            ControlWord::SectionNoColumnBalance(param) => {
                require_parameterless(*param, "nocolbal")?;
                properties.balance_columns = false;
            },
            ControlWord::SectionDefaultColumns(param) => {
                require_parameterless(*param, "sectdefaultcl")?;
                properties.columns = Default::default();
            },
            ControlWord::PageNumberFormat(format) => {
                properties.page_number_format = *format;
            },
            ControlWord::PageNumberRestart(restart) => {
                properties.page_number_restart = Some(*restart);
            },
            ControlWord::PageNumberOffsetX(value) => {
                properties.page_number_offset_x = Some(*value);
            },
            ControlWord::PageNumberOffsetY(value) => {
                properties.page_number_offset_y = Some(*value);
            },
            ControlWord::PageNumberHeadingLevel(value) => {
                let value = value.unwrap_or(0);
                if !(0..=super::super::super::section::MAX_PAGE_NUMBER_HEADING_LEVEL)
                    .contains(&value)
                {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF page-number heading level must be in 0..={}",
                        super::super::super::section::MAX_PAGE_NUMBER_HEADING_LEVEL
                    )));
                }
                properties.page_number_heading.level = Some(value as u8);
            },
            ControlWord::PageNumberHeadingSeparator(separator) => {
                properties.page_number_heading.separator = Some(*separator);
            },
            ControlWord::SectionLineGrid(value) => {
                let value = value.unwrap_or(360);
                if !(0..=super::super::super::section::MAX_SECTION_LINE_GRID_TWIPS).contains(&value)
                {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF section line-grid pitch must be in 0..={} twips",
                        super::super::super::section::MAX_SECTION_LINE_GRID_TWIPS
                    )));
                }
                properties.document_grid.line_grid = Some(value);
            },
            ControlWord::SectionDocumentGrid(grid_type) => {
                properties.document_grid.grid_type = Some(*grid_type);
            },
            ControlWord::SectionRevisionAuthor(value) => {
                properties.revision.author = Some(nonnegative_author_index(*value, "srauth")?);
            },
            ControlWord::SectionRevisionDate(value) => {
                properties.revision.date = Some(*value);
            },
            ControlWord::VerticalAlignTop => {
                properties.vertical_alignment = VerticalAlignment::Top;
            },
            ControlWord::VerticalAlignCenter => {
                properties.vertical_alignment = VerticalAlignment::Center;
            },
            ControlWord::VerticalAlignJustify => {
                properties.vertical_alignment = VerticalAlignment::Justify;
            },
            ControlWord::VerticalAlignBottom => {
                properties.vertical_alignment = VerticalAlignment::Bottom;
            },
            ControlWord::LineNumbering(value) => {
                let value = value.unwrap_or(1);
                if value < 0
                    || value > i32::from(super::super::super::section::MAX_SECTION_LINE_INCREMENT)
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF line-number increment must be in 0..=65535".to_string(),
                    ));
                }
                properties.line_numbering.increment =
                    if value == 0 { None } else { Some(value as u16) };
            },
            ControlWord::LineNumberDistance(value) => {
                let value = value.unwrap_or(360);
                if !(0..=super::super::super::section::MAX_SECTION_LINE_DISTANCE).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF line-number distance must be in 0..=31680 twips".to_string(),
                    ));
                }
                properties.line_numbering.distance = Some(value);
            },
            ControlWord::LineNumberStart(value) => {
                let value = value.unwrap_or(1);
                if value <= 0 || value as u32 > super::super::super::section::MAX_SECTION_LINE_START
                {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF starting line number must be in 1..={}",
                        super::super::super::section::MAX_SECTION_LINE_START
                    )));
                }
                properties.line_numbering.start = Some(value as u32);
            },
            ControlWord::LineNumberRestartSection => {
                properties.line_numbering.restart =
                    Some(super::super::super::section::SectionLineNumberRestart::Section);
            },
            ControlWord::LineNumberRestartPage => {
                properties.line_numbering.restart =
                    Some(super::super::super::section::SectionLineNumberRestart::Page);
            },
            ControlWord::LineNumberContinuous => {
                properties.line_numbering.restart =
                    Some(super::super::super::section::SectionLineNumberRestart::Continuous);
            },
            _ => {},
        }
        Ok(true)
    }

    pub(super) fn parse_page_border_run(&mut self) -> RtfResult<crate::PageBorder> {
        let mut border = crate::PageBorder::default();
        let mut saw_style = false;
        let mut seen = 0u8;
        while let Some(Token::Control(control)) = self.tokens.get(self.pos) {
            let style = match control {
                ControlWord::BorderNone => Some(crate::PageBorderStyle::None),
                ControlWord::BorderSingle => Some(crate::PageBorderStyle::Single),
                ControlWord::BorderThick => Some(crate::PageBorderStyle::Thick),
                ControlWord::BorderDotted => Some(crate::PageBorderStyle::Dotted),
                ControlWord::BorderDashed => Some(crate::PageBorderStyle::Dashed),
                ControlWord::BorderDashSmall => Some(crate::PageBorderStyle::DashSmallGap),
                ControlWord::BorderDotDash => Some(crate::PageBorderStyle::DotDash),
                ControlWord::BorderDotDotDash => Some(crate::PageBorderStyle::DotDotDash),
                ControlWord::BorderDouble => Some(crate::PageBorderStyle::Double),
                ControlWord::BorderTriple => Some(crate::PageBorderStyle::Triple),
                ControlWord::BorderThinThickSmall => {
                    Some(crate::PageBorderStyle::ThinThickSmallGap)
                },
                ControlWord::BorderThickThinSmall => {
                    Some(crate::PageBorderStyle::ThickThinSmallGap)
                },
                ControlWord::BorderThinThickThinSmall => {
                    Some(crate::PageBorderStyle::ThinThickThinSmallGap)
                },
                ControlWord::BorderThinThickMedium => {
                    Some(crate::PageBorderStyle::ThinThickMediumGap)
                },
                ControlWord::BorderThickThinMedium => {
                    Some(crate::PageBorderStyle::ThickThinMediumGap)
                },
                ControlWord::BorderThinThickThinMedium => {
                    Some(crate::PageBorderStyle::ThinThickThinMediumGap)
                },
                ControlWord::BorderThinThickLarge => {
                    Some(crate::PageBorderStyle::ThinThickLargeGap)
                },
                ControlWord::BorderThickThinLarge => {
                    Some(crate::PageBorderStyle::ThickThinLargeGap)
                },
                ControlWord::BorderThinThickThinLarge => {
                    Some(crate::PageBorderStyle::ThinThickThinLargeGap)
                },
                ControlWord::BorderWave => Some(crate::PageBorderStyle::Wavy),
                ControlWord::BorderWavyDouble => Some(crate::PageBorderStyle::DoubleWavy),
                ControlWord::BorderStriped => Some(crate::PageBorderStyle::Striped),
                ControlWord::BorderEmbossed => Some(crate::PageBorderStyle::Embossed),
                ControlWord::BorderEngraved => Some(crate::PageBorderStyle::Engraved),
                ControlWord::BorderOutset => Some(crate::PageBorderStyle::Outset),
                ControlWord::BorderInset => Some(crate::PageBorderStyle::Inset),
                _ => None,
            };
            if let Some(style) = style {
                if saw_style {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF page-border style".to_string(),
                    ));
                }
                saw_style = true;
                border.style = style;
                self.pos += 1;
                continue;
            }
            match control {
                ControlWord::PageBorderArt(value) => {
                    if saw_style {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF page-border style/art".to_string(),
                        ));
                    }
                    let value = value.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF brdrart requires a numeric parameter".to_string(),
                        )
                    })?;
                    border.art = Some(u8::try_from(value).map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF page-border art".to_string())
                    })?);
                    saw_style = true;
                },
                ControlWord::BorderWidth(value) => {
                    if !saw_style || seen & 1 != 0 {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF page-border width".to_string(),
                        ));
                    }
                    border.width = u8::try_from(value.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF page brdrw requires a numeric parameter".to_string(),
                        )
                    })?)
                    .map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF page-border width".to_string())
                    })?;
                    seen |= 1;
                },
                ControlWord::BorderColor(value) => {
                    if !saw_style || seen & 2 != 0 {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF page-border color".to_string(),
                        ));
                    }
                    border.color_ref = u16::try_from(value.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF page brdrcf requires a numeric parameter".to_string(),
                        )
                    })?)
                    .map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF page-border color".to_string())
                    })?;
                    seen |= 2;
                },
                ControlWord::BorderSpace(value) => {
                    if !saw_style || seen & 4 != 0 {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF page-border spacing".to_string(),
                        ));
                    }
                    border.space = u16::try_from(value.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF page brsp requires a numeric parameter".to_string(),
                        )
                    })?)
                    .map_err(|_| {
                        RtfError::MalformedDocument("invalid RTF page-border spacing".to_string())
                    })?;
                    seen |= 4;
                },
                ControlWord::BorderShadow => {
                    if !saw_style || seen & 8 != 0 {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF page-border shadow".to_string(),
                        ));
                    }
                    border.shadow = true;
                    seen |= 8;
                },
                ControlWord::BorderFrame => {
                    if !saw_style || seen & 16 != 0 {
                        return Err(RtfError::MalformedDocument(
                            "invalid or duplicate RTF page-border frame".to_string(),
                        ));
                    }
                    border.frame = true;
                    seen |= 16;
                },
                _ => break,
            }
            self.pos += 1;
        }
        if !saw_style {
            return Err(RtfError::MalformedDocument(
                "RTF page-border edge requires a style or art control".to_string(),
            ));
        }
        border.validate()?;
        Ok(border)
    }
}
