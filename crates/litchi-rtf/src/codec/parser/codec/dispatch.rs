use super::{
    ControlWord, Destination, MAX_OBJECTS, MAX_SHAPE_GROUPS, MAX_SHAPES, ParsedBodyStoryEvent,
    Parser, RootDrawingOwner, RtfError, RtfResult, Token, parser_classification_error,
    require_parameterless,
};

impl Parser<'_> {
    /// Dispatch the destination-specific handling that a group may open with.
    ///
    /// Returns `true` when the group was fully consumed by a specialised
    /// destination parser and the caller must not fall through to generic
    /// content parsing.
    ///
    /// This is deliberately kept out of [`Parser::parse_group`] and marked
    /// `#[inline(never)]`: the dispatch table is very large, so leaving it inline
    /// would put its correspondingly large stack frame on the recursive
    /// group-nesting path and blow the stack after only a handful of levels.
    // The complexity is concentrated here intentionally so every destination
    // transition remains visible in one auditable protocol dispatch table.
    #[allow(
        clippy::cognitive_complexity,
        reason = "destination dispatch keeps every state transition visible in one table"
    )]
    #[inline(never)]
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn dispatch_group_destination(&mut self) -> RtfResult<bool> {
        if let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::Control(
                    ControlWord::FileTable | ControlWord::FileEntry | ControlWord::BlipUid,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF file-table destinations are misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::GeneratedListText) => {
                    self.parse_generated_list_marker(crate::GeneratedListMarkerKind::Modern)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::LegacyGeneratedListText) => {
                    self.parse_generated_list_marker(crate::GeneratedListMarkerKind::Legacy)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::LegacyParagraphNumbering(_)) => {
                    self.parse_legacy_paragraph_numbering()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::LegacyDrawingObject) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing-object destination must be starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::DefaultCharacterProperties(_)
                    | ControlWord::DefaultParagraphProperties(_),
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF defchp and defpap destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::UnicodeAlternate) => {
                    self.parse_unicode_alternate_group()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::FontTable) => {
                    let valid_scope = self.states.len() == 3
                        || (self.unicode_alternate_depth == 1 && self.states.len() == 4);
                    if self.saw_font_table
                        || !valid_scope
                        || self
                            .blocks
                            .iter()
                            .any(|block| !block.text.trim().is_empty())
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF font table must occur exactly once at document scope before body text".to_string(),
                        ));
                    }
                    self.saw_font_table = true;
                    // Mark this as font table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::FontTable;
                    }
                    self.parse_font_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::ColorTable) => {
                    // Mark this as color table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::ColorTable;
                    }
                    self.parse_color_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::UserProperties) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF userprops destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::DocumentVariable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF docvar destination must be starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::IndexEntry
                    | ControlWord::TableOfContentsEntry
                    | ControlWord::TableOfContentsEntryNoPage,
                ) => {
                    self.parse_navigation_entry_destination()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::XmlOpen | ControlWord::XmlClose) => {
                    self.parse_custom_xml_tag_destination()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::XmlAttributeName | ControlWord::XmlAttributeValue) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF custom XML attribute destinations must be starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::ProtectionRangeStart | ControlWord::ProtectionRangeEnd,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF protection-range destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::KinsokuFollowing | ControlWord::KinsokuLeading) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF kinsoku destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::MathZoneInline) => {
                    self.parse_math_zone_destination(false)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::MathZoneDisplay) => {
                    self.parse_math_zone_destination(true)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(control) if Self::is_math_scoped_control(control) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math destinations and properties may occur only inside a math zone"
                            .to_string(),
                    ));
                },
                Token::Control(ControlWord::IgnorableDestination) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::BookmarkStart | ControlWord::BookmarkEnd
                        ))
                    ) {
                        self.parse_bookmark_destination()?;
                        self.states.pop();
                        return Ok(true);
                    }
                    match self.tokens.get(self.pos + 1) {
                        Some(Token::Control(
                            ControlWord::XmlAttributeName | ControlWord::XmlAttributeValue,
                        )) => {
                            self.parse_custom_xml_attribute_destination()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            ControlWord::ProtectionRangeStart | ControlWord::ProtectionRangeEnd,
                        )) => {
                            self.parse_protection_range_destination()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::XmlOpen | ControlWord::XmlClose)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF custom XML tag destinations must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::FileTable)) => {
                            if self.file_table.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple filetbl destinations".to_string(),
                                ));
                            }
                            self.file_table = Some(self.parse_file_table()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::DefaultCharacterProperties(_))) => {
                            self.parse_default_formatting_destination(
                                crate::DefaultFormattingDestination::Character,
                            )?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::DefaultParagraphProperties(_))) => {
                            self.parse_default_formatting_destination(
                                crate::DefaultFormattingDestination::Paragraph,
                            )?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::BlipUid)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid destination may occur only inside pict".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::PictureProperties(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF picprop destination may occur only inside pict".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapeBinaryValue(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF svb destination may occur only inside sv".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapeThemeValue(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF hsv destination may occur only after sv inside sp".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapeResult(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF shprslt destination may occur only inside a root shape"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapeText(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF shptxt destination may occur only inside a shape".to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::Object | ControlWord::InvalidObjectDestinationParameter,
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF object destination must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::GeneratedListText | ControlWord::LegacyGeneratedListText,
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF generated list-marker destinations must not be starred"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::LegacyParagraphNumbering(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF pn destination must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::StyleSortMethod(_)
                            | ControlWord::FromHtml(_)
                            | ControlWord::DocumentType(_)
                            | ControlWord::DefaultTabWidth(_),
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF numeric document property must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(control))
                            if Self::is_legacy_drawing_control(control) =>
                        {
                            return Err(RtfError::MalformedDocument(
                                "RTF legacy drawing controls must occur inside a starred do destination"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::NoteKinds(_)
                            | ControlWord::FootnotePlacement(_)
                            | ControlWord::EndnotePlacement(_)
                            | ControlWord::FootnoteStart(_)
                            | ControlWord::EndnoteStart(_)
                            | ControlWord::FootnoteRestart(_)
                            | ControlWord::EndnoteRestart(_)
                            | ControlWord::FootnoteNumbering(_)
                            | ControlWord::EndnoteNumbering(_),
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF note options must be unstarred root document-format controls"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::LegacyDrawingObject)) => {
                            if self.legacy_do_starts_with_text_box() {
                                if let Some(text_box) = self.parse_legacy_text_box()? {
                                    let index = self.legacy_text_boxes.len();
                                    self.legacy_text_boxes.push(text_box);
                                    self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                                        crate::BodyStoryEvent::LegacyTextBox(index),
                                    ));
                                }
                            } else if let Some(drawing) = self.parse_legacy_drawing()? {
                                let index = self.legacy_drawings.len();
                                self.legacy_drawings.push(drawing);
                                self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                                    crate::BodyStoryEvent::LegacyDrawing(index),
                                ));
                            }
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            control @ (ControlWord::FootnoteSeparator
                            | ControlWord::FootnoteContinuationSeparator
                            | ControlWord::FootnoteContinuationNotice
                            | ControlWord::EndnoteSeparator
                            | ControlWord::EndnoteContinuationSeparator
                            | ControlWord::EndnoteContinuationNotice),
                        )) => {
                            let kind = match control {
                                ControlWord::FootnoteSeparator => {
                                    crate::NoteSeparatorKind::FootnoteSeparator
                                },
                                ControlWord::FootnoteContinuationSeparator => {
                                    crate::NoteSeparatorKind::FootnoteContinuationSeparator
                                },
                                ControlWord::FootnoteContinuationNotice => {
                                    crate::NoteSeparatorKind::FootnoteContinuationNotice
                                },
                                ControlWord::EndnoteSeparator => {
                                    crate::NoteSeparatorKind::EndnoteSeparator
                                },
                                ControlWord::EndnoteContinuationSeparator => {
                                    crate::NoteSeparatorKind::EndnoteContinuationSeparator
                                },
                                _ => crate::NoteSeparatorKind::EndnoteContinuationNotice,
                            };
                            let separator = self.parse_note_separator_destination(kind)?;
                            self.note_separators.add(separator)?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            ControlWord::HideReviewMarkup(_)
                            | ControlWord::HideReviewComments(_)
                            | ControlWord::HideReviewInsertionsAndDeletions(_)
                            | ControlWord::UseXslTransform(_)
                            | ControlWord::ReadOnlyRecommended(_)
                            | ControlWord::SavePreviousPicture(_)
                            | ControlWord::FromText(_)
                            | ControlWord::MakeBackup(_)
                            | ControlWord::DefaultSaveFormat(_)
                            | ControlWord::BoilerplateDocument(_)
                            | ControlWord::Word97CompatibilityMode(_)
                            | ControlWord::PostScriptOverText(_)
                            | ControlWord::HorizontalDocument(_)
                            | ControlWord::VerticalDocument(_)
                            | ControlWord::CompressJustification(_)
                            | ControlWord::ExpandJustification(_)
                            | ControlWord::LineBasedOnGrid(_)
                            | ControlWord::FractionalCharacterWidths(_)
                            | ControlWord::AbstractNumberingCleanup(_)
                            | ControlWord::DocumentEventMask(_)
                            | ControlWord::DrawingGridFollowsMargins(_)
                            | ControlWord::SnapToDrawingGrid(_)
                            | ControlWord::DrawingGridHorizontalSpacing(_)
                            | ControlWord::DrawingGridVerticalSpacing(_)
                            | ControlWord::DrawingGridHorizontalOrigin(_)
                            | ControlWord::DrawingGridVerticalOrigin(_)
                            | ControlWord::DrawingGridHorizontalShow(_)
                            | ControlWord::DrawingGridVerticalShow(_)
                            | ControlWord::ParallelGutter(_)
                            | ControlWord::PrintTwoOnOne(_)
                            | ControlWord::ThemeLanguage(_)
                            | ControlWord::ThemeLanguageEastAsian(_)
                            | ControlWord::ThemeLanguageComplexScript(_)
                            | ControlWord::RelyOnVml(_)
                            | ControlWord::ValidateXml(_)
                            | ControlWord::ShowPlaceholderText(_)
                            | ControlWord::IgnoreMixedContent(_)
                            | ControlWord::SaveInvalidXml(_)
                            | ControlWord::ShowXmlErrors(_)
                            | ControlWord::DoNotEmbedSystemFonts(_)
                            | ControlWord::DoNotEmbedLinguisticData(_)
                            | ControlWord::TrackMoves(_)
                            | ControlWord::TrackFormatting(_)
                            | ControlWord::LockDocumentTheme(_)
                            | ControlWord::LockQuickFormatSet(_)
                            | ControlWord::UseNormalStyleForLists(_)
                            | ControlWord::UpdateStylesFromTemplate(_)
                            | ControlWord::DeclareStyleRestrictions(_)
                            | ControlWord::EnforceStyleRestrictions(_)
                            | ControlWord::StyleRestrictionsBackwardCompatibility(_)
                            | ControlWord::AllowAutoFormatOverride(_)
                            | ControlWord::BookFold(_)
                            | ControlWord::ReverseBookFold(_)
                            | ControlWord::BookFoldSheets(_)
                            | ControlWord::RemovePersonalInformation(_)
                            | ControlWord::RemoveDateTimeInformation(_)
                            | ControlWord::HyphenateAutomatically(_)
                            | ControlWord::HyphenateCapitalizedWords(_)
                            | ControlWord::HyphenationConsecutiveLines(_)
                            | ControlWord::HyphenationHotZone(_)
                            | ControlWord::SuppressRaisedLoweredExtraSpacing(_)
                            | ControlWord::SuppressTopPageExtraSpacing(_)
                            | ControlWord::SuppressSpaceBeforeAfterHardBreak(_)
                            | ControlWord::SuppressWordPerfectExtraLineSpacing(_)
                            | ControlWord::SuppressBottomPageExtraSpacing(_)
                            | ControlWord::DoNotBalanceSbcsDbcs(_)
                            | ControlWord::ExpandSpacingAtShiftReturn(_)
                            | ControlWord::DoNotAddSpaceForUnderline(_)
                            | ControlWord::DoNotUnderlineTrailingSpaces(_)
                            | ControlWord::DoNotTranslateBackslashToYen(_)
                            | ControlWord::LegacyAsianLineBreakingRules(_)
                            | ControlWord::CombineLegacyTableBorders(_)
                            | ControlWord::DoNotAlignTableRowsIndependently(_)
                            | ControlWord::DoNotUseRawTableWidth(_)
                            | ControlWord::KeepTableRowsTogether(_)
                            | ControlWord::DoNotAdjustTableLineHeight(_)
                            | ControlWord::DoNotBreakWrappedTablesAcrossPages(_)
                            | ControlWord::PreventAutofitGrowthIntoMargins(_)
                            | ControlWord::UseWord2003TableStyleRules(_)
                            | ControlWord::DoNotUseWord97ShapeLayout(_)
                            | ControlWord::UseLegacyFootnoteLayout(_)
                            | ControlWord::UseHtmlParagraphAutoSpacing(_)
                            | ControlWord::PreserveLastTabAlignment(_)
                            | ControlWord::UseWord95AutoSpacing(_)
                            | ControlWord::ApplyThaiLineBreakingRules(_)
                            | ControlWord::SnapTextToGridInsideTable(_)
                            | ControlWord::AllowHangingPunctuation(_)
                            | ControlWord::UseAsianLineBreakingRules(_)
                            | ControlWord::CompressPunctuationAtLineStart(_)
                            | ControlWord::NoCompatibilityOptions(_)
                            | ControlWord::NoUiCompatibility(_)
                            | ControlWord::NoFeatureThrottle(_)
                            | ControlWord::ForceCompatibilityUpgrade(_)
                            | ControlWord::PreserveAutofitTableWidthAroundShapes(_)
                            | ControlWord::UseHangingIndentAsNumberingTab(_)
                            | ControlWord::UseLegacyKinsokuCharacters(_)
                            | ControlWord::UseLegacyFloatingObjectIndentation(_)
                            | ControlWord::AllowContextualSpacingInTables(_)
                            | ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(_)
                            | ControlWord::IgnoreTextBoxVerticalAlignment(_)
                            | ControlWord::SplitPageBreakParagraph(_)
                            | ControlWord::UseFixedWidthHangul(_)
                            | ControlWord::UseLegacyAutofitWidthExpansion(_)
                            | ControlWord::UseCachedColumnBalancing(_)
                            | ControlWord::UnderlineNumberingSuffix(_)
                            | ControlWord::DoNotSplitRowsAroundFloatingTables(_)
                            | ControlWord::UseAnsiKerningPairs(_),
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF document-property flag must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(
                            control @ (ControlWord::KinsokuFollowing | ControlWord::KinsokuLeading),
                        )) => {
                            // Word wraps the header-level kinsoku destinations
                            // of codepage documents in \upr/\ud pairs; the ANSI
                            // branch is skipped, so the Unicode branch is the
                            // single parsed representation.
                            let in_unicode_alternate = self.unicode_alternate_depth > 0;
                            if !in_unicode_alternate
                                && (self.states.len() != 3 || self.section_note_options_closed)
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF kinsoku destinations must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            let following = matches!(control, ControlWord::KinsokuFollowing);
                            let duplicate = if following {
                                self.kinsoku.following.is_some()
                            } else {
                                self.kinsoku.leading.is_some()
                            };
                            if duplicate {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF kinsoku destination".to_string(),
                                ));
                            }
                            let value = self.parse_kinsoku_destination(following)?;
                            if following {
                                self.kinsoku.following = Some(value);
                            } else {
                                self.kinsoku.leading = Some(value);
                            }
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::WindowCaption)) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF window caption must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            if self.window_caption.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF window caption destination".to_string(),
                                ));
                            }
                            self.window_caption =
                                Some(self.parse_window_caption_destination(true)?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::XslTransform)) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF XSL transform must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            if self.xsl_transform.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF XSL transform destination".to_string(),
                                ));
                            }
                            self.xsl_transform = Some(self.parse_xsl_transform_destination()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::StyleListFilter(parameter))) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF style-list filter must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            if self.style_list_filter.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF style-list filter destination".to_string(),
                                ));
                            }
                            self.style_list_filter =
                                Some(self.parse_style_list_filter_destination(*parameter)?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            control @ (ControlWord::WriteReservation(_)
                            | ControlWord::WriteReservationHash(_)),
                        )) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF write reservation must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            match control {
                                ControlWord::WriteReservation(_) => {
                                    if self.write_reservations.legacy.is_some() {
                                        return Err(RtfError::MalformedDocument(
                                            "duplicate RTF legacy write-reservation destination"
                                                .to_string(),
                                        ));
                                    }
                                    self.write_reservations.legacy =
                                        Some(self.parse_legacy_write_reservation_destination()?);
                                },
                                ControlWord::WriteReservationHash(_) => {
                                    if self.write_reservations.hash.is_some() {
                                        return Err(RtfError::MalformedDocument(
                                            "duplicate RTF write-reservation hash destination"
                                                .to_string(),
                                        ));
                                    }
                                    self.write_reservations.hash =
                                        Some(self.parse_write_reservation_hash_destination()?);
                                },
                                _ => return Err(parser_classification_error()),
                            }
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::UnicodeAlternateDestination)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF ud destination must be the Unicode branch of upr".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ListTable)) => {
                            self.pos += 1;
                            self.parse_list_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::ListOverrideTable)) => {
                            self.pos += 1;
                            self.parse_list_override_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::RevisionTable)) => {
                            self.pos += 1;
                            self.parse_revision_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::FormField | ControlWord::DataField)) => {
                            return Err(RtfError::MalformedDocument(
                                "orphan RTF formfield/datafield destination".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::Generator)) => {
                            if self.generator.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple generator destinations".to_string(),
                                ));
                            }
                            self.generator = Some(self.parse_generator_destination()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::RevisionSaveTable)) => {
                            self.pos += 1;
                            self.parse_revision_save_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::XmlNamespaceTable)) => {
                            self.pos += 1;
                            self.parse_xml_namespace_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::ProtectionUserTable)) => {
                            self.pos += 1;
                            self.parse_protection_user_table()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            control @ (ControlWord::NextFile | ControlWord::DocumentTemplate),
                        )) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF external document references must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            let duplicate = match control {
                                ControlWord::NextFile => {
                                    self.external_references.next_file.is_some()
                                },
                                ControlWord::DocumentTemplate => {
                                    self.external_references.template.is_some()
                                },
                                _ => return Err(parser_classification_error()),
                            };
                            if duplicate {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF external document reference destination"
                                        .to_string(),
                                ));
                            }
                            let value = self.parse_external_reference_destination(*control)?;
                            match control {
                                ControlWord::NextFile => {
                                    self.external_references.next_file = Some(value);
                                },
                                ControlWord::DocumentTemplate => {
                                    self.external_references.template = Some(value);
                                },
                                _ => return Err(parser_classification_error()),
                            }
                            self.external_references.validate()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(
                            control @ (ControlWord::DocumentViewKind(_)
                            | ControlWord::DocumentViewScale(_)
                            | ControlWord::DocumentZoomKind(_)
                            | ControlWord::DocumentViewBackgroundShapes(_)
                            | ControlWord::DocumentViewNoPageBoundaries(_)),
                        )) => {
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF document-view controls must precede visible text at document root"
                                        .to_string(),
                                ));
                            }
                            let view_control = *control;
                            self.pos += 2;
                            if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
                                return Err(RtfError::MalformedDocument(
                                    "starred RTF document-view group must contain exactly one control"
                                        .to_string(),
                                ));
                            }
                            self.pos += 1;
                            self.apply_document_view_control(&view_control)?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::ThemeData)) => {
                            if self.saw_theme_data {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple theme-data destinations".to_string(),
                                ));
                            }
                            self.saw_theme_data = true;
                            self.theme_data = Some(self.parse_theme_hex_destination(
                                ControlWord::ThemeData,
                                crate::theme::MAX_THEME_DATA_BYTES,
                            )?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::ColorSchemeMapping)) => {
                            if self.saw_color_scheme_mapping {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple color-scheme mappings".to_string(),
                                ));
                            }
                            self.saw_color_scheme_mapping = true;
                            self.color_scheme_mapping = Some(self.parse_theme_hex_destination(
                                ControlWord::ColorSchemeMapping,
                                crate::theme::MAX_COLOR_SCHEME_MAPPING_BYTES,
                            )?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::LatentStyles)) => {
                            if self.latent_styles.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple latentstyles destinations".to_string(),
                                ));
                            }
                            self.latent_styles = Some(self.parse_latent_styles()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::LegacySectionNumberingLevel(_))) => {
                            self.parse_legacy_section_numbering_level()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::ParagraphGroupTable)) => {
                            if self.paragraph_group_table.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple pgptbl destinations".to_string(),
                                ));
                            }
                            self.paragraph_group_table = Some(self.parse_paragraph_group_table()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::DataStore)) => {
                            if self.saw_data_store {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple datastore destinations".to_string(),
                                ));
                            }
                            self.saw_data_store = true;
                            self.data_store = Some(self.parse_data_store_destination()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::MailMerge)) => {
                            if self.mail_merge.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple mailmerge destinations".to_string(),
                                ));
                            }
                            self.mail_merge = Some(self.parse_mail_merge_destination()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::MathProperties)) => {
                            if self.math_properties.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple math-properties destinations"
                                        .to_string(),
                                ));
                            }
                            self.math_properties = Some(self.parse_math_properties_destination()?);
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::DocumentVariable)) => {
                            self.parse_document_variable_destination()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::UserProperties)) => {
                            self.parse_user_properties_destination()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::AnnotationAuthor)) => {
                            if self.pending_annotation_author_seen {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate pending RTF annotation author".to_string(),
                                ));
                            }
                            self.pending_annotation_author =
                                self.parse_ignorable_text_destination()?;
                            self.pending_annotation_author_seen = true;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::AnnotationInitials)) => {
                            if self.pending_annotation_initials_seen {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate pending RTF annotation initials".to_string(),
                                ));
                            }
                            self.pending_annotation_initials =
                                self.parse_ignorable_text_destination()?;
                            self.pending_annotation_initials_seen = true;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeStart)) => {
                            self.parse_annotation_range_marker(true)?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeEnd)) => {
                            self.parse_annotation_range_marker(false)?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::Annotation)) => {
                            self.parse_annotation_destination()?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::Shape(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF shp destination must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapeGroup(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF shpgrp destination must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ListPicture(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF listpicture is misplaced".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ShapePicture(parameter))) => {
                            require_parameterless(*parameter, "shppict")?;
                            self.parse_body_picture_compatibility(
                                crate::PictureCompatibilityKind::ShapePicture,
                                true,
                            )?;
                            self.states.pop();
                            return Ok(true);
                        },
                        Some(Token::Control(ControlWord::NonShapePicture(_))) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF nonshppict destination must not be starred".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::BackgroundDestination(parameter))) => {
                            require_parameterless(*parameter, "background")?;
                            if self.states.len() != 3 || self.section_note_options_closed {
                                return Err(RtfError::MalformedDocument(
                                    "RTF background destination must precede visible text at document root".to_string(),
                                ));
                            }
                            if self.background_shape_index.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple background destinations".to_string(),
                                ));
                            }
                            if self.shapes.len() >= MAX_SHAPES {
                                return Err(RtfError::MalformedDocument(
                                    "RTF shape count exceeds the safety limit".to_string(),
                                ));
                            }
                            let shape = self.parse_background_destination()?;
                            self.background_shape_index = Some(self.shapes.len());
                            self.shapes.push(shape);
                            self.states.pop();
                            return Ok(true);
                        },
                        _ => {},
                    }
                    // Retain an unsupported destination as bounded inert syntax.
                    self.preserve_unknown_destination()?;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Other;
                    }
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::StyleSheet) => {
                    // Parse style definitions without adding their names to body text.
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::StyleSheet;
                    }
                    self.parse_stylesheet()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::ListTable) => {
                    self.parse_list_table()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::ListOverrideTable) => {
                    self.parse_list_override_table()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::RevisionTable) => {
                    self.parse_revision_table()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::FormField | ControlWord::DataField) => {
                    return Err(RtfError::MalformedDocument(
                        "orphan RTF formfield/datafield destination".to_string(),
                    ));
                },
                Token::Control(ControlWord::Generator) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generator destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::WindowCaption) => {
                    if self.states.len() != 3 || self.section_note_options_closed {
                        return Err(RtfError::MalformedDocument(
                            "RTF window caption must precede visible text at document root"
                                .to_string(),
                        ));
                    }
                    if self.window_caption.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF window caption destination".to_string(),
                        ));
                    }
                    self.window_caption = Some(self.parse_window_caption_destination(false)?);
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::XslTransform) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XSL transform destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::StyleListFilter(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF style-list filter destination must be starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::WriteReservation(_) | ControlWord::WriteReservationHash(_),
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF write-reservation destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::RevisionSaveTable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-save table must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::XmlNamespaceTable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace table must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ProtectionUserTable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF protection-user table must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::NextFile | ControlWord::DocumentTemplate) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF external document reference destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ThemeData | ControlWord::ColorSchemeMapping) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF theme destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::LatentStyles | ControlWord::LatentStyleExceptions) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latent-style destinations are misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::LegacySectionNumberingLevel(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pnseclvl destination is misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ParagraphGroupTable | ControlWord::ParagraphGroup) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF paragraph-group destination is misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::DataStore) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF datastore destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::MailMerge) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF mailmerge destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::MathProperties) => {
                    if self.math_properties.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF contains multiple math-properties destinations".to_string(),
                        ));
                    }
                    self.math_properties = Some(self.parse_math_properties_destination()?);
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Info) => {
                    // Parse document metadata without adding it to body text.
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Info;
                    }
                    self.parse_info()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Shape(parameter)) => {
                    require_parameterless(*parameter, "shp")?;
                    let state = self.current_state()?;
                    let owner = if matches!(state.destination, Destination::FieldResult)
                        && !self.field_drawing_captures.is_empty()
                    {
                        RootDrawingOwner::FieldResult
                    } else if self.current_note_separator_active {
                        RootDrawingOwner::NoteSeparator
                    } else if matches!(
                        state.destination,
                        Destination::Footnote | Destination::Endnote
                    ) {
                        RootDrawingOwner::Note
                    } else if self.current_hf_type.is_some() {
                        RootDrawingOwner::HeaderFooter
                    } else if state.in_table
                        && matches!(state.destination, Destination::DocumentBody)
                    {
                        RootDrawingOwner::Cell(state.table_nesting_level.max(1))
                    } else {
                        RootDrawingOwner::Body
                    };
                    if let RootDrawingOwner::Cell(level) = owner {
                        self.drain_nested_to(level)?;
                        if level >= 2 {
                            self.ensure_nested_builder(level)?;
                        }
                    }
                    let count = match owner {
                        RootDrawingOwner::FieldResult => {
                            self.current_field_drawing_capture()?.shapes.len()
                        },
                        RootDrawingOwner::NoteSeparator => {
                            self.current_note_separator_drawings.shapes.len()
                        },
                        RootDrawingOwner::Note => self.current_note_shapes.len(),
                        RootDrawingOwner::HeaderFooter => self.current_hf_shapes.len(),
                        RootDrawingOwner::Cell(1) => self.current_cell_drawings.shapes.len(),
                        RootDrawingOwner::Cell(level) => self
                            .ensure_nested_builder(level)?
                            .cell_drawings
                            .shapes
                            .len(),
                        RootDrawingOwner::Body => self.shapes.len(),
                    };
                    if count >= MAX_SHAPES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape count exceeds the safety limit".to_string(),
                        ));
                    }
                    let mut shape = self.parse_shape_destination(true)?;
                    match owner {
                        RootDrawingOwner::FieldResult => {
                            let capture = self.current_field_drawing_capture_mut()?;
                            shape.position = capture.story_offset;
                            let drawing = crate::StoryDrawing::Shape(capture.shapes.len());
                            capture.drawing_order.push(drawing);
                            capture
                                .story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            capture.shapes.push(shape);
                        },
                        RootDrawingOwner::NoteSeparator => {
                            shape.position = self.current_note_separator_drawings.story_offset;
                            self.current_note_separator_elements.push(
                                crate::NoteSeparatorElement::Drawing(crate::StoryDrawing::Shape(
                                    self.current_note_separator_drawings.shapes.len(),
                                )),
                            );
                            self.current_note_separator_drawings.shapes.push(shape);
                        },
                        RootDrawingOwner::Note => {
                            shape.position = self.current_note_buffer.len();
                            let drawing =
                                crate::StoryDrawing::Shape(self.current_note_shapes.len());
                            self.current_note_drawing_order.push(drawing);
                            self.current_note_story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            self.current_note_shapes.push(shape);
                        },
                        RootDrawingOwner::HeaderFooter => {
                            shape.position = self.current_hf_story_offset;
                            let drawing = crate::StoryDrawing::Shape(self.current_hf_shapes.len());
                            self.current_hf_drawing_order.push(drawing);
                            self.current_hf_story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            self.current_hf_shapes.push(shape);
                        },
                        RootDrawingOwner::Cell(1) => {
                            shape.position = self.current_cell_text.len();
                            let drawing =
                                crate::StoryDrawing::Shape(self.current_cell_drawings.shapes.len());
                            self.current_cell_drawings.drawing_order.push(drawing);
                            self.current_cell_story_events
                                .push(crate::CellStoryEvent::Drawing(drawing));
                            self.current_cell_drawings.shapes.push(shape);
                        },
                        RootDrawingOwner::Cell(level) => {
                            let builder = self.ensure_nested_builder(level)?;
                            shape.position = builder.cell_text.len();
                            let drawing =
                                crate::StoryDrawing::Shape(builder.cell_drawings.shapes.len());
                            builder.cell_drawings.drawing_order.push(drawing);
                            builder
                                .cell_story_events
                                .push(crate::CellStoryEvent::Drawing(drawing));
                            builder.cell_drawings.shapes.push(shape);
                        },
                        RootDrawingOwner::Body => {
                            shape.position = self.body_text_len;
                            let drawing = crate::StoryDrawing::Shape(self.shapes.len());
                            self.drawing_order.push(drawing);
                            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                                crate::BodyStoryEvent::Drawing(drawing),
                            ));
                            self.shapes.push(shape);
                        },
                    }
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::NonShapePicture(parameter)) => {
                    require_parameterless(*parameter, "nonshppict")?;
                    self.parse_body_picture_compatibility(
                        crate::PictureCompatibilityKind::NonShapePicture,
                        false,
                    )?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::ShapePicture(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shppict destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeGroup(parameter)) => {
                    require_parameterless(*parameter, "shpgrp")?;
                    let state = self.current_state()?;
                    let owner = if matches!(state.destination, Destination::FieldResult)
                        && !self.field_drawing_captures.is_empty()
                    {
                        RootDrawingOwner::FieldResult
                    } else if self.current_note_separator_active {
                        RootDrawingOwner::NoteSeparator
                    } else if matches!(
                        state.destination,
                        Destination::Footnote | Destination::Endnote
                    ) {
                        RootDrawingOwner::Note
                    } else if self.current_hf_type.is_some() {
                        RootDrawingOwner::HeaderFooter
                    } else if state.in_table
                        && matches!(state.destination, Destination::DocumentBody)
                    {
                        RootDrawingOwner::Cell(state.table_nesting_level.max(1))
                    } else {
                        RootDrawingOwner::Body
                    };
                    if let RootDrawingOwner::Cell(level) = owner {
                        self.drain_nested_to(level)?;
                        if level >= 2 {
                            self.ensure_nested_builder(level)?;
                        }
                    }
                    let count = match owner {
                        RootDrawingOwner::FieldResult => {
                            self.current_field_drawing_capture()?.shape_groups.len()
                        },
                        RootDrawingOwner::NoteSeparator => {
                            self.current_note_separator_drawings.shape_groups.len()
                        },
                        RootDrawingOwner::Note => self.current_note_shape_groups.len(),
                        RootDrawingOwner::HeaderFooter => self.current_hf_shape_groups.len(),
                        RootDrawingOwner::Cell(1) => self.current_cell_drawings.shape_groups.len(),
                        RootDrawingOwner::Cell(level) => self
                            .ensure_nested_builder(level)?
                            .cell_drawings
                            .shape_groups
                            .len(),
                        RootDrawingOwner::Body => self.shape_groups.len(),
                    };
                    if count >= MAX_SHAPE_GROUPS {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group count exceeds the safety limit".to_string(),
                        ));
                    }
                    let mut group = self.parse_shape_group_destination()?;
                    match owner {
                        RootDrawingOwner::FieldResult => {
                            let capture = self.current_field_drawing_capture_mut()?;
                            group.position = capture.story_offset;
                            let drawing =
                                crate::StoryDrawing::ShapeGroup(capture.shape_groups.len());
                            capture.drawing_order.push(drawing);
                            capture
                                .story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            capture.shape_groups.push(group);
                        },
                        RootDrawingOwner::NoteSeparator => {
                            group.position = self.current_note_separator_drawings.story_offset;
                            self.current_note_separator_elements.push(
                                crate::NoteSeparatorElement::Drawing(
                                    crate::StoryDrawing::ShapeGroup(
                                        self.current_note_separator_drawings.shape_groups.len(),
                                    ),
                                ),
                            );
                            self.current_note_separator_drawings
                                .shape_groups
                                .push(group);
                        },
                        RootDrawingOwner::Note => {
                            group.position = self.current_note_buffer.len();
                            let drawing = crate::StoryDrawing::ShapeGroup(
                                self.current_note_shape_groups.len(),
                            );
                            self.current_note_drawing_order.push(drawing);
                            self.current_note_story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            self.current_note_shape_groups.push(group);
                        },
                        RootDrawingOwner::HeaderFooter => {
                            group.position = self.current_hf_story_offset;
                            let drawing =
                                crate::StoryDrawing::ShapeGroup(self.current_hf_shape_groups.len());
                            self.current_hf_drawing_order.push(drawing);
                            self.current_hf_story_events
                                .push(crate::StoryEvent::Drawing(drawing));
                            self.current_hf_shape_groups.push(group);
                        },
                        RootDrawingOwner::Cell(1) => {
                            group.position = self.current_cell_text.len();
                            let drawing = crate::StoryDrawing::ShapeGroup(
                                self.current_cell_drawings.shape_groups.len(),
                            );
                            self.current_cell_drawings.drawing_order.push(drawing);
                            self.current_cell_story_events
                                .push(crate::CellStoryEvent::Drawing(drawing));
                            self.current_cell_drawings.shape_groups.push(group);
                        },
                        RootDrawingOwner::Cell(level) => {
                            let builder = self.ensure_nested_builder(level)?;
                            group.position = builder.cell_text.len();
                            let drawing = crate::StoryDrawing::ShapeGroup(
                                builder.cell_drawings.shape_groups.len(),
                            );
                            builder.cell_drawings.drawing_order.push(drawing);
                            builder
                                .cell_story_events
                                .push(crate::CellStoryEvent::Drawing(drawing));
                            builder.cell_drawings.shape_groups.push(group);
                        },
                        RootDrawingOwner::Body => {
                            group.position = self.body_text_len;
                            let drawing = crate::StoryDrawing::ShapeGroup(self.shape_groups.len());
                            self.drawing_order.push(drawing);
                            self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                                crate::BodyStoryEvent::Drawing(drawing),
                            ));
                            self.shape_groups.push(group);
                        },
                    }
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Picture) => {
                    // Mark as picture destination and extract
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Picture;
                    }
                    self.parse_picture()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::PictureProperties(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF picprop destination may occur only inside pict".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeBinaryValue(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF svb destination may occur only inside sv".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeThemeValue(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF hsv destination may occur only after sv inside sp".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeResult(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shprslt destination may occur only inside a root shape".to_string(),
                    ));
                },
                Token::Control(ControlWord::ShapeText(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shptxt destination may occur only inside a shape".to_string(),
                    ));
                },
                Token::Control(ControlWord::Object) => {
                    if self.objects.len() >= MAX_OBJECTS {
                        return Err(RtfError::MalformedDocument(
                            "RTF embedded object count exceeds the safety limit".to_string(),
                        ));
                    }
                    let object = self.parse_object_destination()?;
                    let index = self.objects.len();
                    self.objects.push(object);
                    self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
                        crate::BodyStoryEvent::Object(index),
                    ));
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::InvalidObjectDestinationParameter) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object destination must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::Result) => {
                    // Mark as result destination and skip
                    // This contains the rendered result of an embedded object
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Result;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::InvalidObjectResultDestinationParameter) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF object result destination must not have a parameter".to_string(),
                    ));
                },
                Token::Control(ControlWord::Field) => {
                    // Parse field group
                    self.parse_field()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Header) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::Header);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::HeaderFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::HeaderFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::HeaderLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::HeaderLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::HeaderRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::HeaderRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Footer) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::Footer);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::FooterFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::FooterFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::FooterLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::FooterLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::FooterRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type =
                        Some(super::super::super::section::HeaderFooterType::FooterRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(ControlWord::Footnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footnote;
                    }
                    self.parse_note(true)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::Control(
                    ControlWord::FootnoteSeparator
                    | ControlWord::FootnoteContinuationSeparator
                    | ControlWord::FootnoteContinuationNotice
                    | ControlWord::EndnoteSeparator
                    | ControlWord::EndnoteContinuationSeparator
                    | ControlWord::EndnoteContinuationNotice,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note-separator destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::Endnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Endnote;
                    }
                    self.parse_note(false)?;
                    self.states.pop();
                    return Ok(true);
                },
                Token::OpenBrace
                | Token::CloseBrace
                | Token::Control(_)
                | Token::Text(_)
                | Token::Binary(_) => {},
            }
        }

        Ok(false)
    }
}
