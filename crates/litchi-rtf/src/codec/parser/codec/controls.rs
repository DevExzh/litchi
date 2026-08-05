use super::*;

impl<'a> Parser<'a> {
    /// Apply a control word to the current state.
    // This exhaustive state-transition table mirrors the RTF control-word
    // specification; splitting it would obscure coverage and precedence.
    #[allow(clippy::too_many_lines)]
    pub(super) fn apply_control_word(&mut self, control: &ControlWord) -> RtfResult<()> {
        if let ControlWord::Page(parameter) = control {
            require_parameterless(*parameter, "page")?;
            return Err(RtfError::MalformedDocument(
                "RTF page is not permitted in this destination".to_string(),
            ));
        }
        if let ControlWord::EditableRegionStart(parameter) = control {
            require_parameterless(*parameter, "ebcstart")?;
            return Err(RtfError::MalformedDocument(
                "RTF editable-region marks are supported only in the main body story".to_string(),
            ));
        }
        if let ControlWord::EditableRegionEnd(parameter) = control {
            require_parameterless(*parameter, "ebcend")?;
            return Err(RtfError::MalformedDocument(
                "RTF editable-region marks are supported only in the main body story".to_string(),
            ));
        }
        if matches!(
            control,
            ControlWord::SoftPageBreak(_)
                | ControlWord::SoftColumnBreak(_)
                | ControlWord::SoftLineBreak(_)
                | ControlWord::SoftLineHeight(_)
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF soft-break controls are supported only in the main body story".to_string(),
            ));
        }
        if matches!(
            control,
            ControlWord::DefaultFont(_)
                | ControlWord::AssociatedDefaultFont(_)
                | ControlWord::StylesheetDefaultBidiFont(_)
                | ControlWord::StylesheetDefaultDoubleByteFont(_)
                | ControlWord::StylesheetDefaultHighAnsiFont(_)
                | ControlWord::StylesheetDefaultLowAnsiFont(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF default-font selectors must occur in the root document header".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::DefaultFont(v) => (1, "deff", v),
                ControlWord::AssociatedDefaultFont(v) => (2, "adeff", v),
                ControlWord::StylesheetDefaultBidiFont(v) => (4, "stshfbi", v),
                ControlWord::StylesheetDefaultDoubleByteFont(v) => (8, "stshfdbch", v),
                ControlWord::StylesheetDefaultHighAnsiFont(v) => (16, "stshfhich", v),
                ControlWord::StylesheetDefaultLowAnsiFont(v) => (32, "stshfloch", v),
                _ => return Err(parser_classification_error()),
            };
            if self.default_font_selectors_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} default-font selector"
                )));
            }
            let value = u16::try_from(parameter.ok_or_else(|| {
                RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
            })?)
            .map_err(|_| {
                RtfError::MalformedDocument(format!("RTF {name} value must be in 0..=65535"))
            })?;
            self.default_font_selectors_seen |= bit;
            let fonts = &mut self.default_formatting.fonts;
            match control {
                ControlWord::DefaultFont(_) => fonts.primary = Some(value),
                ControlWord::AssociatedDefaultFont(_) => fonts.associated = Some(value),
                ControlWord::StylesheetDefaultBidiFont(_) => fonts.stylesheet_bidi = Some(value),
                ControlWord::StylesheetDefaultDoubleByteFont(_) => {
                    fonts.stylesheet_double_byte = Some(value)
                },
                ControlWord::StylesheetDefaultHighAnsiFont(_) => {
                    fonts.stylesheet_high_ansi = Some(value)
                },
                ControlWord::StylesheetDefaultLowAnsiFont(_) => {
                    fonts.stylesheet_low_ansi = Some(value)
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if let ControlWord::DefaultTabWidth(parameter) = control {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF deftab must precede visible text at document root".to_string(),
                ));
            }
            let value = parameter.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF deftab requires a nonnegative numeric parameter".to_string(),
                )
            })?;
            let value = u32::try_from(value).map_err(|_| {
                RtfError::MalformedDocument(
                    "RTF deftab requires a nonnegative numeric parameter".to_string(),
                )
            })?;
            // Real producers (LibreOffice in particular) restate `\deftab` once
            // per paragraph-properties reset, so an identical redeclaration is
            // idempotent rather than malformed. Only a genuinely conflicting
            // value is ambiguous, and that is still rejected.
            if let Some(existing) = self.default_tab_width_twips
                && existing != value
            {
                return Err(RtfError::MalformedDocument(
                    "conflicting RTF deftab document property".to_string(),
                ));
            }
            self.default_tab_width_twips = Some(value);
            return Ok(());
        }
        if let ControlWord::KinsokuLanguage(parameter) = control {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF ksulang must precede visible text at document root".to_string(),
                ));
            }
            if self.kinsoku.language.is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF ksulang document property".to_string(),
                ));
            }
            let value = parameter.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF ksulang requires a nonnegative numeric parameter".to_string(),
                )
            })?;
            self.kinsoku.language = Some(u32::try_from(value).map_err(|_| {
                RtfError::MalformedDocument(
                    "RTF ksulang requires a nonnegative numeric parameter".to_string(),
                )
            })?);
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::MakeBackup(_)
                | ControlWord::DefaultSaveFormat(_)
                | ControlWord::BoilerplateDocument(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF file-setting flag must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::MakeBackup(parameter) => (1, "makebackup", parameter),
                ControlWord::DefaultSaveFormat(parameter) => (2, "defformat", parameter),
                ControlWord::BoilerplateDocument(parameter) => (4, "doctemp", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.file_settings_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.file_settings_seen |= bit;
            match control {
                ControlWord::MakeBackup(_) => self.file_settings.automatic_backup = true,
                ControlWord::DefaultSaveFormat(_) => {
                    self.file_settings.default_save_format_rtf = true;
                },
                ControlWord::BoilerplateDocument(_) => {
                    self.file_settings.template_or_stationery = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::Word97CompatibilityMode(_) | ControlWord::PostScriptOverText(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF output-setting flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::Word97CompatibilityMode(parameter) => (1, "muser", parameter),
                ControlWord::PostScriptOverText(parameter) => (2, "psover", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.output_settings_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.output_settings_seen |= bit;
            match control {
                ControlWord::Word97CompatibilityMode(_) => {
                    self.output_settings.word97_compatibility_marker = true;
                },
                ControlWord::PostScriptOverText(_) => {
                    self.output_settings.postscript_over_text = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::HorizontalDocument(_)
                | ControlWord::VerticalDocument(_)
                | ControlWord::CompressJustification(_)
                | ControlWord::ExpandJustification(_)
                | ControlWord::LineBasedOnGrid(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF rendering flag must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::HorizontalDocument(parameter) => (1, "horzdoc", parameter),
                ControlWord::VerticalDocument(parameter) => (1, "vertdoc", parameter),
                ControlWord::CompressJustification(parameter) => (2, "jcompress", parameter),
                ControlWord::ExpandJustification(parameter) => (2, "jexpand", parameter),
                ControlWord::LineBasedOnGrid(parameter) => (4, "lnongrid", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.rendering_settings_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate or conflicting RTF {name} rendering property"
                )));
            }
            self.rendering_settings_seen |= bit;
            match control {
                ControlWord::HorizontalDocument(_) => {
                    self.rendering_settings.orientation =
                        Some(crate::DocumentRenderingOrientation::Horizontal);
                },
                ControlWord::VerticalDocument(_) => {
                    self.rendering_settings.orientation =
                        Some(crate::DocumentRenderingOrientation::Vertical);
                },
                ControlWord::CompressJustification(_) => {
                    self.rendering_settings.justification_mode =
                        Some(crate::DocumentJustificationMode::Compress);
                },
                ControlWord::ExpandJustification(_) => {
                    self.rendering_settings.justification_mode =
                        Some(crate::DocumentJustificationMode::Expand);
                },
                ControlWord::LineBasedOnGrid(_) => {
                    self.rendering_settings.line_based_on_grid = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::FractionalCharacterWidths(_)
                | ControlWord::AbstractNumberingCleanup(_)
                | ControlWord::DocumentEventMask(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF processing property must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, _parameter) = match control {
                ControlWord::FractionalCharacterWidths(parameter) => (1, "fracwidth", parameter),
                ControlWord::AbstractNumberingCleanup(parameter) => {
                    (2, "ilfomacatclnup", parameter)
                },
                ControlWord::DocumentEventMask(parameter) => (4, "grfdocevents", parameter),
                _ => return Err(parser_classification_error()),
            };
            if self.processing_settings_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.processing_settings_seen |= bit;
            match control {
                ControlWord::FractionalCharacterWidths(None) => {
                    self.processing_settings
                        .fractional_character_widths_for_printing = true;
                },
                ControlWord::FractionalCharacterWidths(Some(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF fracwidth must not have a numeric parameter".to_string(),
                    ));
                },
                ControlWord::AbstractNumberingCleanup(Some(0)) => {
                    self.processing_settings.abstract_numbering_cleanup =
                        Some(crate::AbstractNumberingCleanupStatus::Reviewed);
                },
                ControlWord::AbstractNumberingCleanup(Some(1)) => {
                    self.processing_settings.abstract_numbering_cleanup =
                        Some(crate::AbstractNumberingCleanupStatus::Incomplete);
                },
                ControlWord::AbstractNumberingCleanup(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF ilfomacatclnup must have value 0 or 1".to_string(),
                    ));
                },
                ControlWord::DocumentEventMask(Some(value @ 0..=0x7fff)) => {
                    self.processing_settings.event_mask = Some(
                        crate::DocumentEventMask::from_bits(*value as u16).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF grfdocevents contains unsupported event bits".to_string(),
                            )
                        })?,
                    );
                },
                ControlWord::DocumentEventMask(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF grfdocevents must have a value from 0 through 32767".to_string(),
                    ));
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DrawingGridFollowsMargins(_)
                | ControlWord::SnapToDrawingGrid(_)
                | ControlWord::DrawingGridHorizontalSpacing(_)
                | ControlWord::DrawingGridVerticalSpacing(_)
                | ControlWord::DrawingGridHorizontalOrigin(_)
                | ControlWord::DrawingGridVerticalOrigin(_)
                | ControlWord::DrawingGridHorizontalShow(_)
                | ControlWord::DrawingGridVerticalShow(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF drawing-grid property must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name) = match control {
                ControlWord::DrawingGridFollowsMargins(_) => (1, "dgmargin"),
                ControlWord::SnapToDrawingGrid(_) => (2, "dgsnap"),
                ControlWord::DrawingGridHorizontalSpacing(_) => (4, "dghspace"),
                ControlWord::DrawingGridVerticalSpacing(_) => (8, "dgvspace"),
                ControlWord::DrawingGridHorizontalOrigin(_) => (16, "dghorigin"),
                ControlWord::DrawingGridVerticalOrigin(_) => (32, "dgvorigin"),
                ControlWord::DrawingGridHorizontalShow(_) => (64, "dghshow"),
                ControlWord::DrawingGridVerticalShow(_) => (128, "dgvshow"),
                _ => return Err(parser_classification_error()),
            };
            if self.drawing_grid_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} drawing-grid property"
                )));
            }
            self.drawing_grid_seen |= bit;
            match control {
                ControlWord::DrawingGridFollowsMargins(None) => {
                    self.drawing_grid.follows_margins = true;
                },
                ControlWord::SnapToDrawingGrid(None) => {
                    self.drawing_grid.snap_to_grid = true;
                },
                ControlWord::DrawingGridFollowsMargins(Some(_))
                | ControlWord::SnapToDrawingGrid(Some(_)) => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} must not have a numeric parameter"
                    )));
                },
                ControlWord::DrawingGridHorizontalSpacing(Some(value @ 0..=32767)) => {
                    self.drawing_grid.horizontal_spacing = Some(
                        crate::DrawingGridSpacing::new(*value as u16).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF dghspace is outside the supported range".to_string(),
                            )
                        })?,
                    );
                },
                ControlWord::DrawingGridVerticalSpacing(Some(value @ 0..=32767)) => {
                    self.drawing_grid.vertical_spacing = Some(
                        crate::DrawingGridSpacing::new(*value as u16).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF dgvspace is outside the supported range".to_string(),
                            )
                        })?,
                    );
                },
                ControlWord::DrawingGridHorizontalOrigin(Some(value @ -32768..=32767)) => {
                    self.drawing_grid.horizontal_origin_twips = Some(*value as i16);
                },
                ControlWord::DrawingGridVerticalOrigin(Some(value @ -32768..=32767)) => {
                    self.drawing_grid.vertical_origin_twips = Some(*value as i16);
                },
                ControlWord::DrawingGridHorizontalShow(Some(value @ 0..=32767)) => {
                    self.drawing_grid.horizontal_line_interval = Some(
                        crate::DrawingGridLineInterval::new(*value as u16).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF dghshow is outside the supported range".to_string(),
                            )
                        })?,
                    );
                },
                ControlWord::DrawingGridVerticalShow(Some(value @ 0..=32767)) => {
                    self.drawing_grid.vertical_line_interval = Some(
                        crate::DrawingGridLineInterval::new(*value as u16).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF dgvshow is outside the supported range".to_string(),
                            )
                        })?,
                    );
                },
                _ => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} has a missing or out-of-range numeric parameter"
                    )));
                },
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::FacingPages(_)
                | ControlWord::MirrorMargins(_)
                | ControlWord::DocumentGutter(_)
                | ControlWord::ParallelGutter(_)
                | ControlWord::PrintTwoOnOne(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF print-layout setting must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name) = match control {
                ControlWord::FacingPages(_) => (1, "facingp"),
                ControlWord::MirrorMargins(_) => (2, "margmirror"),
                ControlWord::DocumentGutter(_) => (4, "gutter"),
                ControlWord::ParallelGutter(_) => (8, "gutterprl"),
                ControlWord::PrintTwoOnOne(_) => (16, "twoonone"),
                _ => return Err(parser_classification_error()),
            };
            if self.print_layout_settings_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.print_layout_settings_seen |= bit;
            match control {
                ControlWord::FacingPages(enabled) => {
                    self.print_layout_settings.facing_pages = *enabled;
                },
                ControlWord::MirrorMargins(parameter) => {
                    if parameter.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF margmirror must not have a numeric parameter".to_string(),
                        ));
                    }
                    self.print_layout_settings.mirror_margins = true;
                },
                ControlWord::DocumentGutter(Some(value @ 0..=31_680)) => {
                    let value = *value as u32;
                    self.print_layout_settings.document_gutter_twips = Some(value);
                    for (index, section) in self.sections.iter_mut().enumerate() {
                        if !self
                            .section_gutter_overrides
                            .get(index)
                            .copied()
                            .unwrap_or(false)
                        {
                            section.properties.margin_gutter = value as i32;
                        }
                    }
                },
                ControlWord::DocumentGutter(_) => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF gutter must have a numeric parameter in 0..={} twips",
                        crate::MAX_DOCUMENT_GUTTER_TWIPS
                    )));
                },
                ControlWord::ParallelGutter(_) => {
                    if let ControlWord::ParallelGutter(Some(_)) = control {
                        return Err(RtfError::MalformedDocument(
                            "RTF gutterprl must not have a numeric parameter".to_string(),
                        ));
                    }
                    self.print_layout_settings.parallel_gutter = true;
                },
                ControlWord::PrintTwoOnOne(_) => {
                    if let ControlWord::PrintTwoOnOne(Some(_)) = control {
                        return Err(RtfError::MalformedDocument(
                            "RTF twoonone must not have a numeric parameter".to_string(),
                        ));
                    }
                    self.print_layout_settings
                        .two_logical_pages_per_physical_page = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::ThemeLanguage(_)
                | ControlWord::ThemeLanguageEastAsian(_)
                | ControlWord::ThemeLanguageComplexScript(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF theme language must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::ThemeLanguage(parameter) => (1, "themelang", parameter),
                ControlWord::ThemeLanguageEastAsian(parameter) => (2, "themelangfe", parameter),
                ControlWord::ThemeLanguageComplexScript(parameter) => (4, "themelangcs", parameter),
                _ => return Err(parser_classification_error()),
            };
            if self.theme_languages_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            let value = parameter.ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "RTF {name} control requires a numeric language ID"
                ))
            })?;
            let language = crate::LanguageId::from_rtf(value)?;
            self.theme_languages_seen |= bit;
            match control {
                ControlWord::ThemeLanguage(_) => self.theme_languages.primary = Some(language),
                ControlWord::ThemeLanguageEastAsian(_) => {
                    self.theme_languages.east_asian = Some(language);
                },
                ControlWord::ThemeLanguageComplexScript(_) => {
                    self.theme_languages.complex_script = Some(language);
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::RelyOnVml(_)
                | ControlWord::ValidateXml(_)
                | ControlWord::ShowPlaceholderText(_)
                | ControlWord::IgnoreMixedContent(_)
                | ControlWord::SaveInvalidXml(_)
                | ControlWord::ShowXmlErrors(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF XML policy must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::RelyOnVml(parameter) => (1, "relyonvml", parameter),
                ControlWord::ValidateXml(parameter) => (2, "validatexml", parameter),
                ControlWord::ShowPlaceholderText(parameter) => (4, "showplaceholdtext", parameter),
                ControlWord::IgnoreMixedContent(parameter) => (8, "ignoremixedcontent", parameter),
                ControlWord::SaveInvalidXml(parameter) => (16, "saveinvalidxml", parameter),
                ControlWord::ShowXmlErrors(parameter) => (32, "showxmlerrors", parameter),
                _ => return Err(parser_classification_error()),
            };
            if self.xml_policies_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            let enabled = match parameter {
                Some(0) => false,
                Some(1) => true,
                _ => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} must have numeric value 0 or 1"
                    )));
                },
            };
            self.xml_policies_seen |= bit;
            match control {
                ControlWord::RelyOnVml(_) => self.xml_policies.rely_on_vml = Some(enabled),
                ControlWord::ValidateXml(_) => {
                    self.xml_policies.validate_custom_xml = Some(enabled);
                },
                ControlWord::ShowPlaceholderText(_) => {
                    self.xml_policies.show_placeholder_text = Some(enabled);
                },
                ControlWord::IgnoreMixedContent(_) => {
                    self.xml_policies.ignore_mixed_content = Some(enabled);
                },
                ControlWord::SaveInvalidXml(_) => {
                    self.xml_policies.save_invalid_xml = Some(enabled);
                },
                ControlWord::ShowXmlErrors(_) => {
                    self.xml_policies.show_xml_errors = Some(enabled);
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DoNotEmbedSystemFonts(_) | ControlWord::DoNotEmbedLinguisticData(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF embedding policy must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::DoNotEmbedSystemFonts(parameter) => {
                    (1, "donotembedsysfont", parameter)
                },
                ControlWord::DoNotEmbedLinguisticData(parameter) => {
                    (2, "donotembedlingdata", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if self.embedding_policies_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            let enabled = match parameter {
                Some(0) => false,
                Some(1) => true,
                _ => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} must have numeric value 0 or 1"
                    )));
                },
            };
            self.embedding_policies_seen |= bit;
            match control {
                ControlWord::DoNotEmbedSystemFonts(_) => {
                    self.embedding_policies.do_not_embed_system_fonts = Some(enabled);
                },
                ControlWord::DoNotEmbedLinguisticData(_) => {
                    self.embedding_policies.do_not_embed_linguistic_data = Some(enabled);
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::TrackMoves(_) | ControlWord::TrackFormatting(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF revision policy must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::TrackMoves(parameter) => (1, "trackmoves", parameter),
                ControlWord::TrackFormatting(parameter) => (2, "trackformatting", parameter),
                _ => return Err(parser_classification_error()),
            };
            if self.revision_policies_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            let enabled = match parameter {
                Some(0) => false,
                Some(1) => true,
                _ => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} must have numeric value 0 or 1"
                    )));
                },
            };
            self.revision_policies_seen |= bit;
            match control {
                ControlWord::TrackMoves(_) => self.revision_policies.track_moves = Some(enabled),
                ControlWord::TrackFormatting(_) => {
                    self.revision_policies.track_formatting = Some(enabled);
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::LockDocumentTheme(_)
                | ControlWord::LockQuickFormatSet(_)
                | ControlWord::UseNormalStyleForLists(_)
                | ControlWord::UpdateStylesFromTemplate(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF style policy must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::LockDocumentTheme(parameter) => (1, "stylelocktheme", parameter),
                ControlWord::LockQuickFormatSet(parameter) => (2, "stylelockqfset", parameter),
                ControlWord::UseNormalStyleForLists(parameter) => {
                    (4, "usenormstyforlist", parameter)
                },
                ControlWord::UpdateStylesFromTemplate(parameter) => (8, "linkstyles", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.style_policies_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.style_policies_seen |= bit;
            match control {
                ControlWord::LockDocumentTheme(_) => self.style_policies.lock_theme = true,
                ControlWord::LockQuickFormatSet(_) => {
                    self.style_policies.lock_quick_format_set = true;
                },
                ControlWord::UseNormalStyleForLists(_) => {
                    self.style_policies.use_normal_style_for_lists = true;
                },
                ControlWord::UpdateStylesFromTemplate(_) => {
                    self.style_policies.update_styles_from_template = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DeclareStyleRestrictions(_)
                | ControlWord::EnforceStyleRestrictions(_)
                | ControlWord::StyleRestrictionsBackwardCompatibility(_)
                | ControlWord::AllowAutoFormatOverride(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF style restriction must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::DeclareStyleRestrictions(parameter) => (1, "stylelock", parameter),
                ControlWord::EnforceStyleRestrictions(parameter) => {
                    (2, "stylelockenforced", parameter)
                },
                ControlWord::StyleRestrictionsBackwardCompatibility(parameter) => {
                    (4, "stylelockbackcomp", parameter)
                },
                ControlWord::AllowAutoFormatOverride(parameter) => {
                    (8, "autofmtoverride", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.style_restrictions_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.style_restrictions_seen |= bit;
            match control {
                ControlWord::DeclareStyleRestrictions(_) => {
                    self.style_restrictions.restrictions_present = true;
                },
                ControlWord::EnforceStyleRestrictions(_) => {
                    self.style_restrictions.enforced = true;
                },
                ControlWord::StyleRestrictionsBackwardCompatibility(_) => {
                    self.style_restrictions.backward_compatibility = true;
                },
                ControlWord::AllowAutoFormatOverride(_) => {
                    self.style_restrictions.allow_auto_format_override = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::BookFold(_)
                | ControlWord::ReverseBookFold(_)
                | ControlWord::BookFoldSheets(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF booklet-printing property must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter, requires_parameter) = match control {
                ControlWord::BookFold(parameter) => (1, "bookfold", parameter, false),
                ControlWord::ReverseBookFold(parameter) => (2, "bookfoldrev", parameter, false),
                ControlWord::BookFoldSheets(parameter) => (4, "bookfoldsheets", parameter, true),
                _ => return Err(parser_classification_error()),
            };
            if requires_parameter != parameter.is_some() {
                let requirement = if requires_parameter {
                    "requires a numeric parameter"
                } else {
                    "must not have a numeric parameter"
                };
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} {requirement}"
                )));
            }
            if self.booklet_printing_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            if let Some(value) = parameter {
                if *value < 0 || *value % 4 != 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF bookfoldsheets must be a nonnegative multiple of four".to_string(),
                    ));
                }
                self.booklet_printing.sheets_per_booklet = Some(*value as u32);
            } else {
                match control {
                    ControlWord::BookFold(_) => self.booklet_printing.book_fold = true,
                    ControlWord::ReverseBookFold(_) => {
                        self.booklet_printing.reverse_book_fold = true;
                    },
                    _ => return Err(parser_classification_error()),
                }
            }
            self.booklet_printing_seen |= bit;
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::RemovePersonalInformation(_) | ControlWord::RemoveDateTimeInformation(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF privacy policy must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::RemovePersonalInformation(parameter) => {
                    (1, "rempersonalinfo", parameter)
                },
                ControlWord::RemoveDateTimeInformation(parameter) => (2, "remdttm", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.privacy_policies_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.privacy_policies_seen |= bit;
            match control {
                ControlWord::RemovePersonalInformation(_) => {
                    self.privacy_policies.remove_personal_information = true;
                },
                ControlWord::RemoveDateTimeInformation(_) => {
                    self.privacy_policies.remove_date_time_information = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::PreserveAutofitTableWidthAroundShapes(_)
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
                | ControlWord::UseAnsiKerningPairs(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF Word 2003 compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::PreserveAutofitTableWidthAroundShapes(parameter) => {
                    (1, "noafcnsttbl", parameter)
                },
                ControlWord::UseHangingIndentAsNumberingTab(parameter) => {
                    (2, "noindnmbrts", parameter)
                },
                ControlWord::UseLegacyKinsokuCharacters(parameter) => (4, "felnbrelev", parameter),
                ControlWord::UseLegacyFloatingObjectIndentation(parameter) => {
                    (8, "indrlsweleven", parameter)
                },
                ControlWord::AllowContextualSpacingInTables(parameter) => {
                    (16, "nocxsptable", parameter)
                },
                ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(parameter) => {
                    (32, "notcvasp", parameter)
                },
                ControlWord::IgnoreTextBoxVerticalAlignment(parameter) => {
                    (64, "notvatxbx", parameter)
                },
                ControlWord::SplitPageBreakParagraph(parameter) => (128, "spltpgpar", parameter),
                ControlWord::UseFixedWidthHangul(parameter) => (256, "hwelev", parameter),
                ControlWord::UseLegacyAutofitWidthExpansion(parameter) => {
                    (512, "afelev", parameter)
                },
                ControlWord::UseCachedColumnBalancing(parameter) => {
                    (1024, "cachedcolbal", parameter)
                },
                ControlWord::UnderlineNumberingSuffix(parameter) => (2048, "utinl", parameter),
                ControlWord::DoNotSplitRowsAroundFloatingTables(parameter) => {
                    (4096, "notbrkcnstfrctbl", parameter)
                },
                ControlWord::UseAnsiKerningPairs(parameter) => (8192, "krnprsnet", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.word_2003_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.word_2003_compatibility_seen |= bit;
            match control {
                ControlWord::PreserveAutofitTableWidthAroundShapes(_) => {
                    self.word_2003_compatibility
                        .preserve_autofit_table_width_around_shapes = true
                },
                ControlWord::UseHangingIndentAsNumberingTab(_) => {
                    self.word_2003_compatibility
                        .use_hanging_indent_as_numbering_tab = true
                },
                ControlWord::UseLegacyKinsokuCharacters(_) => {
                    self.word_2003_compatibility.use_legacy_kinsoku_characters = true
                },
                ControlWord::UseLegacyFloatingObjectIndentation(_) => {
                    self.word_2003_compatibility
                        .use_legacy_floating_object_indentation = true
                },
                ControlWord::AllowContextualSpacingInTables(_) => {
                    self.word_2003_compatibility
                        .allow_contextual_spacing_in_tables = true
                },
                ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(_) => {
                    self.word_2003_compatibility
                        .ignore_cell_vertical_alignment_with_floating_objects = true
                },
                ControlWord::IgnoreTextBoxVerticalAlignment(_) => {
                    self.word_2003_compatibility
                        .ignore_text_box_vertical_alignment = true
                },
                ControlWord::SplitPageBreakParagraph(_) => {
                    self.word_2003_compatibility.split_page_break_paragraph = true
                },
                ControlWord::UseFixedWidthHangul(_) => {
                    self.word_2003_compatibility.use_fixed_width_hangul = true
                },
                ControlWord::UseLegacyAutofitWidthExpansion(_) => {
                    self.word_2003_compatibility
                        .use_legacy_autofit_width_expansion = true
                },
                ControlWord::UseCachedColumnBalancing(_) => {
                    self.word_2003_compatibility.use_cached_column_balancing = true
                },
                ControlWord::UnderlineNumberingSuffix(_) => {
                    self.word_2003_compatibility.underline_numbering_suffix = true
                },
                ControlWord::DoNotSplitRowsAroundFloatingTables(_) => {
                    self.word_2003_compatibility
                        .do_not_split_rows_around_floating_tables = true
                },
                ControlWord::UseAnsiKerningPairs(_) => {
                    self.word_2003_compatibility.use_ansi_kerning_pairs = true
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::NoCompatibilityOptions(_)
                | ControlWord::NoUiCompatibility(_)
                | ControlWord::NoFeatureThrottle(_)
                | ControlWord::ForceCompatibilityUpgrade(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF compatibility policy must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::NoCompatibilityOptions(parameter) => (1, "nocompatoptions", parameter),
                ControlWord::NoUiCompatibility(parameter) => (2, "nouicompat", parameter),
                ControlWord::NoFeatureThrottle(parameter) => (4, "nofeaturethrottle", parameter),
                ControlWord::ForceCompatibilityUpgrade(parameter) => (8, "forceupgrade", parameter),
                _ => return Err(parser_classification_error()),
            };
            if self.compatibility_policy_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            match control {
                ControlWord::NoFeatureThrottle(_) => {
                    let value = parameter.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF nofeaturethrottle requires parameter 0 or 1".to_string(),
                        )
                    })?;
                    self.compatibility_policy.feature_throttle = Some(
                        crate::DocumentFeatureThrottle::from_rtf(value).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF nofeaturethrottle parameter must be 0 or 1".to_string(),
                            )
                        })?,
                    );
                },
                _ if parameter.is_some() => {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF {name} must not have a numeric parameter"
                    )));
                },
                ControlWord::NoCompatibilityOptions(_) => {
                    self.compatibility_policy.reset_options_to_defaults = true;
                },
                ControlWord::NoUiCompatibility(_) => {
                    self.compatibility_policy.feature_throttle =
                        Some(crate::DocumentFeatureThrottle::Unrestricted);
                },
                ControlWord::ForceCompatibilityUpgrade(_) => {
                    self.compatibility_policy.force_upgrade = true;
                },
                _ => return Err(parser_classification_error()),
            }
            self.compatibility_policy_seen |= bit;
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::ApplyThaiLineBreakingRules(_)
                | ControlWord::SnapTextToGridInsideTable(_)
                | ControlWord::AllowHangingPunctuation(_)
                | ControlWord::UseAsianLineBreakingRules(_)
                | ControlWord::CompressPunctuationAtLineStart(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF Asian grid compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::ApplyThaiLineBreakingRules(parameter) => {
                    (1, "ApplyBrkRules", parameter)
                },
                ControlWord::SnapTextToGridInsideTable(parameter) => {
                    (2, "snaptogridincell", parameter)
                },
                ControlWord::AllowHangingPunctuation(parameter) => (4, "wrppunct", parameter),
                ControlWord::UseAsianLineBreakingRules(parameter) => (8, "asianbrkrule", parameter),
                ControlWord::CompressPunctuationAtLineStart(parameter) => {
                    (16, "toplinepunct", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.asian_grid_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.asian_grid_compatibility_seen |= bit;
            match control {
                ControlWord::ApplyThaiLineBreakingRules(_) => {
                    self.asian_grid_compatibility.apply_thai_line_breaking_rules = true;
                },
                ControlWord::SnapTextToGridInsideTable(_) => {
                    self.asian_grid_compatibility.snap_text_to_grid_inside_table = true;
                },
                ControlWord::AllowHangingPunctuation(_) => {
                    self.asian_grid_compatibility.allow_hanging_punctuation = true;
                },
                ControlWord::UseAsianLineBreakingRules(_) => {
                    self.asian_grid_compatibility.use_asian_line_breaking_rules = true;
                },
                ControlWord::CompressPunctuationAtLineStart(_) => {
                    self.asian_grid_compatibility
                        .compress_punctuation_at_line_start = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DoNotUseWord97ShapeLayout(_)
                | ControlWord::UseLegacyFootnoteLayout(_)
                | ControlWord::UseHtmlParagraphAutoSpacing(_)
                | ControlWord::PreserveLastTabAlignment(_)
                | ControlWord::UseWord95AutoSpacing(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF legacy layout compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::DoNotUseWord97ShapeLayout(parameter) => (1, "splytwnine", parameter),
                ControlWord::UseLegacyFootnoteLayout(parameter) => (2, "ftnlytwnine", parameter),
                ControlWord::UseHtmlParagraphAutoSpacing(parameter) => (4, "htmautsp", parameter),
                ControlWord::PreserveLastTabAlignment(parameter) => (8, "useltbaln", parameter),
                ControlWord::UseWord95AutoSpacing(parameter) => (16, "oldas", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.legacy_layout_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.legacy_layout_compatibility_seen |= bit;
            match control {
                ControlWord::DoNotUseWord97ShapeLayout(_) => {
                    self.legacy_layout_compatibility
                        .do_not_use_word_97_shape_layout = true;
                },
                ControlWord::UseLegacyFootnoteLayout(_) => {
                    self.legacy_layout_compatibility.use_legacy_footnote_layout = true;
                },
                ControlWord::UseHtmlParagraphAutoSpacing(_) => {
                    self.legacy_layout_compatibility
                        .use_html_paragraph_auto_spacing = true;
                },
                ControlWord::PreserveLastTabAlignment(_) => {
                    self.legacy_layout_compatibility.preserve_last_tab_alignment = true;
                },
                ControlWord::UseWord95AutoSpacing(_) => {
                    self.legacy_layout_compatibility.use_word_95_auto_spacing = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::CombineLegacyTableBorders(_)
                | ControlWord::DoNotAlignTableRowsIndependently(_)
                | ControlWord::DoNotUseRawTableWidth(_)
                | ControlWord::KeepTableRowsTogether(_)
                | ControlWord::DoNotAdjustTableLineHeight(_)
                | ControlWord::DoNotBreakWrappedTablesAcrossPages(_)
                | ControlWord::PreventAutofitGrowthIntoMargins(_)
                | ControlWord::UseWord2003TableStyleRules(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF table-layout compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::CombineLegacyTableBorders(parameter) => (1, "otblrul", parameter),
                ControlWord::DoNotAlignTableRowsIndependently(parameter) => {
                    (2, "alntblind", parameter)
                },
                ControlWord::DoNotUseRawTableWidth(parameter) => (4, "lytcalctblwd", parameter),
                ControlWord::KeepTableRowsTogether(parameter) => (8, "lyttblrtgr", parameter),
                ControlWord::DoNotAdjustTableLineHeight(parameter) => {
                    (16, "nolnhtadjtbl", parameter)
                },
                ControlWord::DoNotBreakWrappedTablesAcrossPages(parameter) => {
                    (32, "nobrkwrptbl", parameter)
                },
                ControlWord::PreventAutofitGrowthIntoMargins(parameter) => {
                    (64, "nogrowautofit", parameter)
                },
                ControlWord::UseWord2003TableStyleRules(parameter) => {
                    (128, "newtblstyruls", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.table_layout_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.table_layout_compatibility_seen |= bit;
            match control {
                ControlWord::CombineLegacyTableBorders(_) => {
                    self.table_layout_compatibility.combine_borders_like_word_5 = true;
                },
                ControlWord::DoNotAlignTableRowsIndependently(_) => {
                    self.table_layout_compatibility
                        .do_not_align_rows_independently = true;
                },
                ControlWord::DoNotUseRawTableWidth(_) => {
                    self.table_layout_compatibility.do_not_use_raw_table_width = true;
                },
                ControlWord::KeepTableRowsTogether(_) => {
                    self.table_layout_compatibility.keep_rows_together = true;
                },
                ControlWord::DoNotAdjustTableLineHeight(_) => {
                    self.table_layout_compatibility.do_not_adjust_line_height = true;
                },
                ControlWord::DoNotBreakWrappedTablesAcrossPages(_) => {
                    self.table_layout_compatibility
                        .do_not_break_wrapped_tables_across_pages = true;
                },
                ControlWord::PreventAutofitGrowthIntoMargins(_) => {
                    self.table_layout_compatibility
                        .prevent_autofit_growth_into_margins = true;
                },
                ControlWord::UseWord2003TableStyleRules(_) => {
                    self.table_layout_compatibility
                        .use_word_2003_table_style_rules = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DoNotBalanceSbcsDbcs(_)
                | ControlWord::ExpandSpacingAtShiftReturn(_)
                | ControlWord::DoNotAddSpaceForUnderline(_)
                | ControlWord::DoNotUnderlineTrailingSpaces(_)
                | ControlWord::DoNotTranslateBackslashToYen(_)
                | ControlWord::LegacyAsianLineBreakingRules(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF East Asian compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::DoNotBalanceSbcsDbcs(parameter) => (1, "dntblnsbdb", parameter),
                ControlWord::ExpandSpacingAtShiftReturn(parameter) => (2, "expshrtn", parameter),
                ControlWord::DoNotAddSpaceForUnderline(parameter) => (4, "nospaceforul", parameter),
                ControlWord::DoNotUnderlineTrailingSpaces(parameter) => {
                    (8, "noultrlspc", parameter)
                },
                ControlWord::DoNotTranslateBackslashToYen(parameter) => {
                    (16, "noxlattoyen", parameter)
                },
                ControlWord::LegacyAsianLineBreakingRules(parameter) => {
                    (32, "lnbrkrule", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.east_asian_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.east_asian_compatibility_seen |= bit;
            match control {
                ControlWord::DoNotBalanceSbcsDbcs(_) => {
                    self.east_asian_compatibility.do_not_balance_sbcs_dbcs = true;
                },
                ControlWord::ExpandSpacingAtShiftReturn(_) => {
                    self.east_asian_compatibility.expand_spacing_at_shift_return = true;
                },
                ControlWord::DoNotAddSpaceForUnderline(_) => {
                    self.east_asian_compatibility.do_not_add_space_for_underline = true;
                },
                ControlWord::DoNotUnderlineTrailingSpaces(_) => {
                    self.east_asian_compatibility
                        .do_not_underline_trailing_spaces = true;
                },
                ControlWord::DoNotTranslateBackslashToYen(_) => {
                    self.east_asian_compatibility
                        .do_not_translate_backslash_to_yen = true;
                },
                ControlWord::LegacyAsianLineBreakingRules(_) => {
                    self.east_asian_compatibility.use_legacy_line_breaking_rules = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::SuppressRaisedLoweredExtraSpacing(_)
                | ControlWord::SuppressTopPageExtraSpacing(_)
                | ControlWord::SuppressSpaceBeforeAfterHardBreak(_)
                | ControlWord::SuppressWordPerfectExtraLineSpacing(_)
                | ControlWord::SuppressBottomPageExtraSpacing(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF line-spacing compatibility flag must precede visible text at document root"
                        .to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::SuppressRaisedLoweredExtraSpacing(parameter) => {
                    (1, "noextrasprl", parameter)
                },
                ControlWord::SuppressTopPageExtraSpacing(parameter) => (2, "sprstsp", parameter),
                ControlWord::SuppressSpaceBeforeAfterHardBreak(parameter) => {
                    (4, "sprsspbf", parameter)
                },
                ControlWord::SuppressWordPerfectExtraLineSpacing(parameter) => {
                    (8, "sprslnsp", parameter)
                },
                ControlWord::SuppressBottomPageExtraSpacing(parameter) => {
                    (16, "sprsbsp", parameter)
                },
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.line_spacing_compatibility_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.line_spacing_compatibility_seen |= bit;
            match control {
                ControlWord::SuppressRaisedLoweredExtraSpacing(_) => {
                    self.line_spacing_compatibility
                        .suppress_extra_spacing_for_raised_lowered_text = true;
                },
                ControlWord::SuppressTopPageExtraSpacing(_) => {
                    self.line_spacing_compatibility
                        .suppress_extra_spacing_at_top_of_page = true;
                },
                ControlWord::SuppressSpaceBeforeAfterHardBreak(_) => {
                    self.line_spacing_compatibility
                        .suppress_space_before_after_hard_break = true;
                },
                ControlWord::SuppressWordPerfectExtraLineSpacing(_) => {
                    self.line_spacing_compatibility
                        .suppress_wordperfect_extra_line_spacing = true;
                },
                ControlWord::SuppressBottomPageExtraSpacing(_) => {
                    self.line_spacing_compatibility
                        .suppress_extra_spacing_at_bottom_of_page = true;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if matches!(control, ControlWord::FromText(_) | ControlWord::FromHtml(_)) {
            if self.states.len() != 2 || self.section_note_options_closed || self.saw_font_table {
                return Err(RtfError::MalformedDocument(
                    "RTF document origin must occur in the header before font tables and visible text"
                        .to_string(),
                ));
            }
            if self.origin_metadata.origin.is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate or conflicting RTF document-origin property".to_string(),
                ));
            }
            self.origin_metadata.origin = Some(match control {
                ControlWord::FromText(None) => crate::DocumentOrigin::PlainTextEmail,
                ControlWord::FromText(Some(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF fromtext must not have a numeric parameter".to_string(),
                    ));
                },
                ControlWord::FromHtml(version) => crate::DocumentOrigin::HtmlEmail {
                    version: version
                        .map(crate::HtmlEmailVersion::from_rtf_value)
                        .transpose()?,
                },
                _ => return Err(parser_classification_error()),
            });
            return Ok(());
        }
        if let ControlWord::DocumentType(parameter) = control {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF document type must precede visible text at document root".to_string(),
                ));
            }
            if self.origin_metadata.auto_format_type.is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF document type property".to_string(),
                ));
            }
            self.origin_metadata.auto_format_type = Some(
                crate::DocumentAutoFormatType::from_rtf_value(parameter.unwrap_or(0))?,
            );
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::ReadOnlyRecommended(_) | ControlWord::SavePreviousPicture(_)
        ) {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF save preference must precede visible text at document root".to_string(),
                ));
            }
            let (bit, name, parameter) = match control {
                ControlWord::ReadOnlyRecommended(parameter) => {
                    (1, "readonlyrecommended", parameter)
                },
                ControlWord::SavePreviousPicture(parameter) => (2, "saveprevpict", parameter),
                _ => return Err(parser_classification_error()),
            };
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must not have a numeric parameter"
                )));
            }
            if self.save_preferences_seen & bit != 0 {
                return Err(RtfError::MalformedDocument(format!(
                    "duplicate RTF {name} document property"
                )));
            }
            self.save_preferences_seen |= bit;
            match control {
                ControlWord::ReadOnlyRecommended(_) => {
                    self.save_preferences.read_only =
                        crate::DocumentReadOnlyRecommendation::Recommended;
                },
                ControlWord::SavePreviousPicture(_) => {
                    self.save_preferences.thumbnail =
                        crate::DocumentThumbnailPreference::RequiredIfSupported;
                },
                _ => return Err(parser_classification_error()),
            }
            return Ok(());
        }
        if let ControlWord::StyleSortMethod(parameter) = control {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF style-sort method must precede visible text at document root".to_string(),
                ));
            }
            if self.style_sort_method_seen {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF style-sort method document property".to_string(),
                ));
            }
            self.style_sort_method = Some(crate::DocumentStyleSortMethod::from_rtf_value(
                parameter.unwrap_or(1),
            )?);
            self.style_sort_method_seen = true;
            return Ok(());
        }
        if let ControlWord::UseXslTransform(parameter) = control {
            if self.states.len() != 2 || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF usexform must precede visible text at document root".to_string(),
                ));
            }
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF usexform must not have a numeric parameter".to_string(),
                ));
            }
            if self.use_xsl_transform_seen {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF usexform document property".to_string(),
                ));
            }
            self.use_xsl_transform_seen = true;
            self.xsl_transform_usage = crate::DocumentXslTransformUsage::Requested;
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::HideReviewMarkup(_)
                | ControlWord::HideReviewComments(_)
                | ControlWord::HideReviewInsertionsAndDeletions(_)
        ) {
            let at_root = self.states.len() == 2
                && self
                    .states
                    .last()
                    .is_some_and(|state| state.destination == Destination::DocumentBody);
            if !at_root || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF review-display flags must precede visible text at document root"
                        .to_string(),
                ));
            }
            self.apply_review_display_control(control)?;
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::DocumentViewKind(_)
                | ControlWord::DocumentViewScale(_)
                | ControlWord::DocumentZoomKind(_)
                | ControlWord::DocumentViewBackgroundShapes(_)
                | ControlWord::DocumentViewNoPageBoundaries(_)
        ) {
            let at_root = self.states.len() == 2
                && self
                    .states
                    .last()
                    .is_some_and(|state| state.destination == Destination::DocumentBody);
            if !at_root || self.section_note_options_closed {
                return Err(RtfError::MalformedDocument(
                    "RTF document-view controls must precede visible text at document root"
                        .to_string(),
                ));
            }
            self.apply_document_view_control(control)?;
            return Ok(());
        }
        if matches!(
            control,
            ControlWord::HyphenateAutomatically(_)
                | ControlWord::HyphenateCapitalizedWords(_)
                | ControlWord::HyphenationConsecutiveLines(_)
                | ControlWord::HyphenationHotZone(_)
        ) {
            let at_root = self.states.len() == 2
                && self
                    .states
                    .last()
                    .is_some_and(|state| state.destination == Destination::DocumentBody);
            // This flag closes as soon as visible body text (including pending text that has
            // not yet been flushed into a StyleBlock) enters the document stream.
            let body_started = self.section_note_options_closed;
            if !at_root || body_started {
                return Err(RtfError::MalformedDocument(
                    "RTF document hyphenation controls must precede visible text at document root"
                        .to_string(),
                ));
            }
            self.apply_document_hyphenation_control(control)?;
            return Ok(());
        }
        if let ControlWord::TableNestingLevel(parameter) = control {
            let value = parameter.ok_or_else(|| {
                RtfError::MalformedDocument("RTF itap requires a numeric parameter".to_string())
            })?;
            let level = u8::try_from(value).map_err(|_| {
                RtfError::MalformedDocument("RTF itap is outside 0..=32".to_string())
            })?;
            if usize::from(level) > crate::MAX_TABLE_NESTING_DEPTH {
                return Err(RtfError::MalformedDocument(
                    "RTF itap is outside 0..=32".to_string(),
                ));
            }
            let previous = self.current_state()?.table_nesting_level;
            self.current_state_mut()?.table_nesting_level = level;
            let previous = if previous >= 2 { previous } else { 1 };
            let effective = if level >= 2 { level } else { 1 };
            if effective < previous {
                self.drain_nested_to(effective)?;
            }
            return Ok(());
        }
        match control {
            ControlWord::RevisionSaveRoot(value) => {
                if self.states.len() != 2 || self.saw_revision_save_root {
                    return Err(RtfError::MalformedDocument(
                        "RTF rsidroot must occur exactly once at document scope".to_string(),
                    ));
                }
                let value = u32::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF revision root must be a positive signed integer".to_string(),
                    )
                })?;
                if value == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision root must be a positive signed integer".to_string(),
                    ));
                }
                self.saw_revision_save_root = true;
                self.revision_save_root = Some(value);
                return Ok(());
            },
            ControlWord::RevisionSaveId(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF rsid control outside rsidtbl".to_string(),
                ));
            },
            ControlWord::BlipTag(_)
            | ControlWord::BlipUnitsPerInch(_)
            | ControlWord::PictureScaled(_)
            | ControlWord::PictureBitmap(_)
            | ControlWord::PictureBitsPerPixel(_)
            | ControlWord::PictureCropLeft(_)
            | ControlWord::PictureCropRight(_)
            | ControlWord::PictureCropTop(_)
            | ControlWord::PictureCropBottom(_)
            | ControlWord::WindowsBitmapBitsPerPixel(_)
            | ControlWord::WindowsBitmapPlanes(_)
            | ControlWord::WindowsBitmapWidthBytes(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF picture metadata control outside pict".to_string(),
                ));
            },
            control @ (ControlWord::NoteKinds(_)
            | ControlWord::FootnotePlacement(_)
            | ControlWord::EndnotePlacement(_)
            | ControlWord::FootnoteStart(_)
            | ControlWord::EndnoteStart(_)
            | ControlWord::FootnoteRestart(_)
            | ControlWord::EndnoteRestart(_)
            | ControlWord::FootnoteNumbering(_)
            | ControlWord::EndnoteNumbering(_)) => {
                if self.states.len() != 2
                    || self.note_options_closed
                    || self.blocks.iter().any(|block| !block.text.is_empty())
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF note options must precede body text at document root".to_string(),
                    ));
                }
                match control {
                    ControlWord::NoteKinds(value) => {
                        self.note_options.present_kinds = Some(match value {
                            0 => crate::PresentNoteKinds::FootnotesOnly,
                            1 => crate::PresentNoteKinds::EndnotesOnly,
                            2 => crate::PresentNoteKinds::FootnotesAndEndnotes,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF fet value must be between 0 and 2".to_string(),
                                ));
                            },
                        });
                    },
                    ControlWord::FootnotePlacement(value) => {
                        self.note_options.footnote_placement = Some(*value);
                    },
                    ControlWord::EndnotePlacement(value) => {
                        self.note_options.endnote_placement = Some(*value);
                    },
                    ControlWord::FootnoteStart(value) => {
                        if *value <= 0 {
                            return Err(RtfError::MalformedDocument(
                                "RTF footnote starting number must be positive".to_string(),
                            ));
                        }
                        self.note_options.footnote_start = Some(*value);
                    },
                    ControlWord::EndnoteStart(value) => {
                        if *value <= 0 {
                            return Err(RtfError::MalformedDocument(
                                "RTF endnote starting number must be positive".to_string(),
                            ));
                        }
                        self.note_options.endnote_start = Some(*value);
                    },
                    ControlWord::FootnoteRestart(value) => {
                        self.note_options.footnote_restart = Some(*value);
                    },
                    ControlWord::EndnoteRestart(value) => {
                        self.note_options.endnote_restart = Some(*value);
                    },
                    ControlWord::FootnoteNumbering(value) => {
                        self.note_options.footnote_numbering = Some(*value);
                    },
                    ControlWord::EndnoteNumbering(value) => {
                        self.note_options.endnote_numbering = Some(*value);
                    },
                    _ => return Err(parser_classification_error()),
                }
                return Ok(());
            },
            ControlWord::LegacyDrawingObject
            | ControlWord::LegacyTextBox
            | ControlWord::LegacyTextBoxText
            | ControlWord::LegacyAnchorXPage
            | ControlWord::LegacyAnchorXMargin
            | ControlWord::LegacyAnchorXColumn
            | ControlWord::LegacyAnchorYPage
            | ControlWord::LegacyAnchorYMargin
            | ControlWord::LegacyAnchorYParagraph
            | ControlWord::LegacyDrawingHeight(_)
            | ControlWord::LegacyTextBoxMargin(_)
            | ControlWord::LegacyDrawingX(_)
            | ControlWord::LegacyDrawingY(_)
            | ControlWord::LegacyDrawingWidth(_)
            | ControlWord::LegacyDrawingHeightSize(_)
            | ControlWord::LegacyTextLeftRightTopBottom
            | ControlWord::LegacyTextLeftRightTopBottomVertical
            | ControlWord::LegacyTextTopBottomRightLeft
            | ControlWord::LegacyTextTopBottomRightLeftVertical
            | ControlWord::LegacyTextBottomTopLeftRight => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF legacy drawing control outside do".to_string(),
                ));
            },
            ControlWord::GeneratedListText | ControlWord::LegacyGeneratedListText => {
                return Err(RtfError::MalformedDocument(
                    "RTF generated list marker must be a grouped body destination".to_string(),
                ));
            },
            ControlWord::XmlNamespace(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF xmlns control outside xmlnstbl".to_string(),
                ));
            },
            ControlWord::ListPicture(_)
            | ControlWord::ShapePicture(_)
            | ControlWord::NonShapePicture(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF listpicture and shppict are misplaced".to_string(),
                ));
            },
            ControlWord::BackgroundDestination(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF background must be a starred root destination".to_string(),
                ));
            },
            ControlWord::ThemeData | ControlWord::ColorSchemeMapping => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF theme destination control".to_string(),
                ));
            },
            ControlWord::LatentStyles
            | ControlWord::LatentStyleMax(_)
            | ControlWord::LatentStyleLockedDefault(_)
            | ControlWord::LatentStyleSemiHiddenDefault(_)
            | ControlWord::LatentStyleUnhideUsedDefault(_)
            | ControlWord::LatentStyleQuickFormatDefault(_)
            | ControlWord::LatentStylePriorityDefault(_)
            | ControlWord::LatentStyleExceptions
            | ControlWord::LatentStyleLocked(_)
            | ControlWord::LatentStyleSemiHidden(_)
            | ControlWord::LatentStyleUnhideUsed(_)
            | ControlWord::LatentStyleQuickFormat(_)
            | ControlWord::LatentStylePriority(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF latent-style control".to_string(),
                ));
            },
            ControlWord::DataStore => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF datastore destination control".to_string(),
                ));
            },
            ControlWord::MailMerge
            | ControlWord::MailMergeConnectString
            | ControlWord::MailMergeConnectStringData
            | ControlWord::MailMergeDataSource
            | ControlWord::MailMergeHeaderSource
            | ControlWord::MailMergeLinkToQuery(_)
            | ControlWord::MailMergeQuery
            | ControlWord::MailMergeDataSourceObject
            | ControlWord::MailMergeActiveRecord(_)
            | ControlWord::MailMergeColumnDelimiter(_)
            | ControlWord::MailMergeColumnCount(_)
            | ControlWord::MailMergeDynamicAddress(_)
            | ControlWord::MailMergeFirstRowHeader(_)
            | ControlWord::MailMergeFilter
            | ControlWord::MailMergeFieldMapData
            | ControlWord::MailMergeFieldMapColumn(_)
            | ControlWord::MailMergeHash(_)
            | ControlWord::MailMergeId(_)
            | ControlWord::MailMergeMappedName
            | ControlWord::MailMergeName
            | ControlWord::MailMergeRecipientData
            | ControlWord::MailMergeSort
            | ControlWord::MailMergeSourceType(_)
            | ControlWord::MailMergeTable
            | ControlWord::MailMergeUdl
            | ControlWord::MailMergeUdlData
            | ControlWord::MailMergeUniqueTag => {
                return Err(RtfError::MalformedDocument(
                    "orphan or misplaced RTF mail-merge control".to_string(),
                ));
            },
            ControlWord::MathProperties
            | ControlWord::MathBreakBinary(_)
            | ControlWord::MathBreakBinarySubtraction(_)
            | ControlWord::MathDefaultJustification(_)
            | ControlWord::MathDisplayDefaults(_)
            | ControlWord::MathInterEquationSpacing(_)
            | ControlWord::MathIntegralLimitPlacement(_)
            | ControlWord::MathIntraEquationSpacing(_)
            | ControlWord::MathLeftMargin(_)
            | ControlWord::MathFont(_)
            | ControlWord::MathNaryLimitPlacement(_)
            | ControlWord::MathPostSpacing(_)
            | ControlWord::MathPreSpacing(_)
            | ControlWord::MathRightMargin(_)
            | ControlWord::MathSmallFractions(_)
            | ControlWord::MathWrapIndent(_)
            | ControlWord::MathWrapRight(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF document math-properties control".to_string(),
                ));
            },
            ControlWord::DefaultLanguage(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.primary = Some(language);
                let state = self.current_state_mut()?;
                state.formatting.language = Some(language);
                state.formatting.language_no_proof = Some(language);
                return Ok(());
            },
            ControlWord::DefaultLanguageEastAsian(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.east_asian = Some(language);
                let state = self.current_state_mut()?;
                state.formatting.east_asian_language = Some(language);
                state.formatting.east_asian_language_no_proof = Some(language);
                return Ok(());
            },
            ControlWord::DefaultLanguageComplexScript(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.complex_script = Some(language);
                self.current_state_mut()?.formatting.associated.language = Some(language);
                return Ok(());
            },
            ControlWord::LeftToRightDocument => {
                self.document_direction = Some(TextDirection::LeftToRight);
                return Ok(());
            },
            ControlWord::RightToLeftDocument => {
                self.document_direction = Some(TextDirection::RightToLeft);
                return Ok(());
            },
            ControlWord::RightGutter(value) => {
                self.gutter_on_right = *value;
                return Ok(());
            },
            ControlWord::FormProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.forms, *value, "formprot")?;
                return Ok(());
            },
            ControlWord::AnnotationProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(
                    &mut self.info.protection.annotations,
                    *value,
                    "annotprot",
                )?;
                return Ok(());
            },
            ControlWord::RevisionProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.revisions, *value, "revprot")?;
                return Ok(());
            },
            ControlWord::ReadOnlyProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.read_only, *value, "readprot")?;
                return Ok(());
            },
            ControlWord::AllProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.all, *value, "allprot")?;
                return Ok(());
            },
            ControlWord::EnforceProtection(Some(value)) => {
                self.ensure_protection_scope()?;
                Self::set_required_protection_flag(
                    &mut self.info.protection.enforced,
                    *value,
                    "enforceprot",
                )?;
                return Ok(());
            },
            ControlWord::EnforceProtection(None) => {
                return Err(RtfError::MalformedDocument(
                    "RTF enforceprot requires a numeric parameter".to_string(),
                ));
            },
            ControlWord::ProtectionLevel(Some(value)) => {
                self.ensure_protection_scope()?;
                if self.info.protection.level.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF protlevel control".to_string(),
                    ));
                }
                self.info.protection.level = Some(crate::ProtectionLevel::from_rtf(*value)?);
                return Ok(());
            },
            ControlWord::ProtectionLevel(None) => {
                return Err(RtfError::MalformedDocument(
                    "RTF protlevel requires a numeric parameter".to_string(),
                ));
            },
            ControlWord::Password => {
                return Err(RtfError::MalformedDocument(
                    "RTF password hash is misplaced or not a starred info destination".to_string(),
                ));
            },
            control if Self::is_math_scoped_control(control) => {
                return Err(RtfError::MalformedDocument(
                    "RTF math controls may occur only inside a math zone destination".to_string(),
                ));
            },
            _ => {},
        }
        if self.apply_section_control(control)? {
            return Ok(());
        }
        let language_defaults = self.language_defaults;
        let state = self.current_state_mut()?;

        if Self::apply_paragraph_tab_control(state, control)? {
            return Ok(());
        }

        if Self::apply_table_decoration_control(state, control)? {
            return Ok(());
        }

        if Self::apply_character_decoration_control(state, control)? {
            return Ok(());
        }

        if Self::apply_paragraph_border_control(state, control)? {
            return Ok(());
        }

        if Self::apply_paragraph_shading_control(state, control)? {
            return Ok(());
        }

        if apply_associated_character_control(&mut state.formatting.associated, control)? {
            return Ok(());
        }

        match control {
            // Font formatting
            ControlWord::CharacterStyle(value) => {
                state.formatting.character_style = Some(character_style_reference(*value)?);
            },
            ControlWord::FontNumber(n) => {
                state.formatting.font_ref = FontRef::try_from(*n).map_err(|_| {
                    RtfError::MalformedDocument("invalid RTF body font reference".to_string())
                })?;
            },
            ControlWord::Language(value) => {
                state.formatting.language = Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageEastAsian(value) => {
                state.formatting.east_asian_language = Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageNoProof(value) => {
                state.formatting.language_no_proof = Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageEastAsianNoProof(value) => {
                state.formatting.east_asian_language_no_proof =
                    Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::NoProof(value) => state.formatting.no_proof = *value,
            ControlWord::LeftToRightCharacter => {
                state.formatting.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftCharacter => {
                state.formatting.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::LowAnsiCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "loch",
                    crate::CharacterType::LowAnsi,
                )?);
            },
            ControlWord::HighAnsiCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "hich",
                    crate::CharacterType::HighAnsi,
                )?);
            },
            ControlWord::DoubleByteCharacter(parameter) => {
                state.formatting.character_type = Some(character_type_selector(
                    *parameter,
                    "dbch",
                    crate::CharacterType::DoubleByte,
                )?);
            },
            ControlWord::FontComplexScript(value) => {
                state.formatting.complex_script = Some(complex_script_selector(*value)?);
            },
            ControlWord::CharacterGrid(value) => {
                state.formatting.character_grid = Some(character_grid(*value)?);
            },
            ControlWord::AnimatedText(value) => {
                state.formatting.animated_text = animated_text(*value)?;
            },
            ControlWord::FitText(value) => {
                state.formatting.fit_text = fit_text(*value)?;
            },
            ControlWord::EmphasisMark(mark, value) => {
                state.formatting.emphasis_mark = emphasis_mark(*mark, *value)?;
            },
            ControlWord::FontSize(size) => {
                if let Some(nz) = NonZeroU16::new((*size).max(0) as u16) {
                    state.formatting.font_size = nz;
                }
            },
            ControlWord::ColorForeground(c) => {
                state.formatting.color_ref = *c as ColorRef;
            },
            ControlWord::ColorBackground(value) => {
                state.formatting.background_color =
                    Some(Self::required_character_value(*value, "cb", u16::MAX)?);
            },

            // Character formatting
            ControlWord::InsertRsid(value) => {
                state.formatting.insert_rsid = Some(*value as u32);
            },
            ControlWord::DeleteRsid(value) => {
                state.formatting.delete_rsid = Some(*value as u32);
            },
            ControlWord::CharStyleRsid(value) => {
                state.formatting.char_style_rsid = Some(*value as u32);
            },
            ControlWord::Bold(b) => state.formatting.bold = *b,
            ControlWord::Italic(b) => state.formatting.italic = *b,
            ControlWord::Underline(b) => {
                state.formatting.underline = if *b {
                    super::super::super::types::UnderlineStyle::Single
                } else {
                    super::super::super::types::UnderlineStyle::None
                }
            },
            ControlWord::UnderlineNone => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::None
            },
            ControlWord::UnderlineDouble => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Double
            },
            ControlWord::UnderlineDotted => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Dotted
            },
            ControlWord::UnderlineDashed => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Dashed
            },
            ControlWord::UnderlineDashDot => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::DashDot
            },
            ControlWord::UnderlineDashDotDot => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::DashDotDot
            },
            ControlWord::UnderlineWords => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Words
            },
            ControlWord::UnderlineThick => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Thick
            },
            ControlWord::UnderlineWave => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Wave
            },
            ControlWord::UnderlineHairline => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::Hairline
            },
            ControlWord::UnderlineThickDotted => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::ThickDotted
            },
            ControlWord::UnderlineThickDashed => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::ThickDashed
            },
            ControlWord::UnderlineThickDashDot => {
                state.formatting.underline =
                    super::super::super::types::UnderlineStyle::ThickDashDot
            },
            ControlWord::UnderlineThickDashDotDot => {
                state.formatting.underline =
                    super::super::super::types::UnderlineStyle::ThickDashDotDot
            },
            ControlWord::UnderlineThickLongDash => {
                state.formatting.underline =
                    super::super::super::types::UnderlineStyle::ThickLongDash
            },
            ControlWord::UnderlineLongDash => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::LongDash
            },
            ControlWord::UnderlineHeavyWave => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::HeavyWave
            },
            ControlWord::UnderlineDoubleWave => {
                state.formatting.underline = super::super::super::types::UnderlineStyle::DoubleWave
            },
            ControlWord::UnderlineColor(value) => {
                state.formatting.underline_color = Some(Self::required_character_value(
                    Some(*value),
                    "ulc",
                    u16::MAX,
                )?);
            },
            ControlWord::Strike(b) => state.formatting.strike = *b,
            ControlWord::DoubleStrike(b) => state.formatting.double_strike = *b,
            ControlWord::Superscript(b) => {
                state.formatting.superscript = *b;
                if *b {
                    state.formatting.subscript = false;
                }
                state.formatting.character_positioning.set_superscript(*b);
            },
            ControlWord::Subscript(b) => {
                state.formatting.subscript = *b;
                if *b {
                    state.formatting.superscript = false;
                }
                state.formatting.character_positioning.set_subscript(*b);
            },
            ControlWord::NoSuperSub => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.clear_baseline();
            },
            ControlWord::BaselineUp(value) => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.set_raised(*value)?;
            },
            ControlWord::BaselineDown(value) => {
                state.formatting.superscript = false;
                state.formatting.subscript = false;
                state.formatting.character_positioning.set_lowered(*value)?;
            },
            ControlWord::SmallCaps(b) => state.formatting.smallcaps = *b,
            ControlWord::AllCaps(b) => state.formatting.all_caps = *b,
            ControlWord::Hidden(b) => state.formatting.hidden = *b,
            ControlWord::Outline(b) => state.formatting.outline = *b,
            ControlWord::Shadow(b) => state.formatting.shadow = *b,
            ControlWord::Emboss(b) => state.formatting.emboss = *b,
            ControlWord::Imprint(b) => state.formatting.imprint = *b,
            ControlWord::CharSpacing(n) => {
                state
                    .formatting
                    .character_positioning
                    .set_quarter_point_expansion(*n)?;
                state.formatting.char_spacing = *n;
            },
            ControlWord::CharSpacingTwips(n) => {
                state
                    .formatting
                    .character_positioning
                    .set_twip_expansion(*n)?;
                state.formatting.char_spacing = *n;
            },
            ControlWord::CharScale(n) => {
                state.formatting.character_positioning.set_scale(*n)?;
                state.formatting.char_scale = *n;
            },
            ControlWord::Kerning(n) => {
                state.formatting.character_positioning.set_kerning(*n)?;
                state.formatting.kerning = *n;
            },
            ControlWord::Highlight(c) => state.formatting.highlight_color = Some(*c as ColorRef),
            ControlWord::Plain => {
                // Reset to default formatting
                state.formatting = Formatting::default();
                state.formatting.language = language_defaults.primary;
                state.formatting.east_asian_language = language_defaults.east_asian;
                state.formatting.language_no_proof = language_defaults.primary;
                state.formatting.east_asian_language_no_proof = language_defaults.east_asian;
                state.formatting.associated.language = language_defaults.complex_script;
                state.character_border_active = false;
                state.character_border_seen = 0;
            },

            // Paragraph alignment
            ControlWord::ParagraphStyle(value) => {
                state.paragraph.paragraph_style = Some(paragraph_style_reference(*value)?);
            },
            ControlWord::ParagraphRsid(value) => {
                state.paragraph.paragraph_rsid = Some(*value as u32);
            },
            ControlWord::ParagraphRevisionAuthor(value) => {
                state.paragraph.revision.author = Some(nonnegative_author_index(*value, "prauth")?);
            },
            ControlWord::ParagraphRevisionDate(value) => {
                state.paragraph.revision.date = Some(*value);
            },
            ControlWord::OutlineLevel(value) => {
                let level = u8::try_from(*value)
                    .ok()
                    .filter(|level| *level <= 9)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF outline level must be between 0 and 9".to_string(),
                        )
                    })?;
                state.paragraph.outline_level = Some(level);
            },
            ControlWord::LeftAlign => state.paragraph.alignment = Alignment::Left,
            ControlWord::RightAlign => state.paragraph.alignment = Alignment::Right,
            ControlWord::Center => state.paragraph.alignment = Alignment::Center,
            ControlWord::Justify => state.paragraph.alignment = Alignment::Justify,
            ControlWord::LeftToRightParagraph => {
                state.paragraph.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftParagraph => {
                state.paragraph.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::Pard => {
                // Reset to default paragraph properties
                state.paragraph = Paragraph::default();
                state.paragraph_border_side = None;
                state.paragraph_border_seen = 0;
                state.drop_cap_kind = None;
                state.drop_cap_lines = None;
                state.pending_tab_alignment = None;
                state.pending_tab_leader = None;
                state.in_table = false;
            },

            // Paragraph spacing
            ControlWord::SpaceBefore(n) => state.paragraph.spacing.before = *n,
            ControlWord::SpaceAfter(n) => state.paragraph.spacing.after = *n,
            ControlWord::SpaceBetween(n) => state.paragraph.spacing.line = *n,
            ControlWord::LineMultiple(b) => state.paragraph.spacing.line_multiple = *b,
            ControlWord::SpaceBeforeAuto(value) => {
                state.paragraph.spacing_policy.automatic_before =
                    required_paragraph_bool(*value, "sbauto")?
            },
            ControlWord::SpaceAfterAuto(value) => {
                state.paragraph.spacing_policy.automatic_after =
                    required_paragraph_bool(*value, "saauto")?
            },
            ControlWord::ListSpaceBefore(value) => {
                state.paragraph.spacing_policy.list_before =
                    Some(required_list_spacing(*value, "lisb")?)
            },
            ControlWord::ListSpaceAfter(value) => {
                state.paragraph.spacing_policy.list_after =
                    Some(required_list_spacing(*value, "lisa")?)
            },
            ControlWord::NoSnapLineGrid(value) => {
                strict_paragraph_selector(*value, "nosnaplinegrid")?;
                state.paragraph.spacing_policy.snap_to_line_grid = false;
            },
            ControlWord::ContextualSpacing(value) => {
                strict_paragraph_selector(*value, "contextualspace")?;
                state.paragraph.spacing_policy.contextual_spacing = true;
            },

            // Paragraph indentation
            ControlWord::LeftIndent(n) => state.paragraph.indentation.left = *n,
            ControlWord::RightIndent(n) => state.paragraph.indentation.right = *n,
            ControlWord::FirstLineIndent(n) => state.paragraph.indentation.first_line = *n,
            ControlWord::LogicalLeftIndent(v) => {
                state.paragraph.logical_indentation.start =
                    Some(required_paragraph_indent(*v, "lin")?)
            },
            ControlWord::LogicalRightIndent(v) => {
                state.paragraph.logical_indentation.end =
                    Some(required_paragraph_indent(*v, "rin")?)
            },
            ControlWord::CharacterFirstLineIndent(v) => {
                state
                    .paragraph
                    .logical_indentation
                    .first_line_character_units = Some(required_paragraph_indent(*v, "cufi")?)
            },
            ControlWord::CharacterLeftIndent(v) => {
                state.paragraph.logical_indentation.left_character_units =
                    Some(required_paragraph_indent(*v, "culi")?)
            },
            ControlWord::CharacterRightIndent(v) => {
                state.paragraph.logical_indentation.right_character_units =
                    Some(required_paragraph_indent(*v, "curi")?)
            },
            ControlWord::MirrorIndents(v) => {
                strict_paragraph_selector(*v, "indmirror")?;
                state.paragraph.logical_indentation.mirrored = true;
            },

            // Paragraph additional properties
            ControlWord::KeepTogether => state.paragraph.keep_together = true,
            ControlWord::KeepNext => state.paragraph.keep_next = true,
            ControlWord::SideBySide(value) => state.paragraph.side_by_side = *value,
            ControlWord::PageBreakBefore => state.paragraph.page_break_before = true,
            ControlWord::WidowControl => state.paragraph.widow_control = true,
            ControlWord::ParagraphNoLineNumbering(param) => {
                require_parameterless(*param, "noline")?;
                state.paragraph.no_line_numbering = true;
            },
            ControlWord::ParagraphNoAutoTabIndent(param) => {
                require_parameterless(*param, "notabind")?;
                state.paragraph.no_auto_tab_indent = true;
            },
            ControlWord::DropCapLines(_) | ControlWord::DropCapType(_) => {
                Self::apply_drop_cap_control(state, control)?;
            },
            ControlWord::ParagraphHyphenation(value) => {
                state.paragraph.line_breaking.automatic_hyphenation =
                    strict_paragraph_toggle(*value, "hyphpar")?
            },
            ControlWord::AutoSpaceAlphabetic(value) => {
                state.paragraph.line_breaking.auto_space_alphabetic =
                    strict_paragraph_toggle(*value, "aspalpha")?
            },
            ControlWord::AutoSpaceNumbers(value) => {
                state.paragraph.line_breaking.auto_space_numbers =
                    strict_paragraph_toggle(*value, "aspnum")?
            },
            ControlWord::AdjustRightIndent(value) => {
                state.paragraph.line_breaking.adjust_right_indent =
                    strict_paragraph_toggle(*value, "adjustright")?
            },
            ControlWord::WrapDefault(value) => {
                strict_paragraph_selector(*value, "wrapdefault")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::Default;
            },
            ControlWord::NoCharacterWrap(value) => {
                strict_paragraph_selector(*value, "nocwrap")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoCharacterWrap;
            },
            ControlWord::NoWordWrap(value) => {
                strict_paragraph_selector(*value, "nowwrap")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoWordWrap;
            },
            ControlWord::NoOverflow(value) => {
                strict_paragraph_selector(*value, "nooverflow")?;
                state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoOverflow;
            },
            ControlWord::FontAlignAuto(value) => {
                strict_paragraph_selector(*value, "faauto")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Auto;
            },
            ControlWord::FontAlignHanging(value) => {
                strict_paragraph_selector(*value, "fahang")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Hanging;
            },
            ControlWord::FontAlignCenter(value) => {
                strict_paragraph_selector(*value, "facenter")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Center;
            },
            ControlWord::FontAlignRoman(value) => {
                strict_paragraph_selector(*value, "faroman")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Roman;
            },
            ControlWord::FontAlignVariable(value) => {
                strict_paragraph_selector(*value, "favar")?;
                state.paragraph.line_breaking.font_alignment =
                    crate::ParagraphFontAlignment::Variable;
            },
            ControlWord::FontAlignFixed(value) => {
                strict_paragraph_selector(*value, "fafixed")?;
                state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Fixed;
            },
            ControlWord::ListOverrideIndex(value) => {
                state.paragraph.list_override = Some(*value);
            },
            ControlWord::ListLevelIndex(value) => {
                let level = u8::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF paragraph list level is outside the supported range".to_string(),
                    )
                })?;
                if level > 8 {
                    return Err(RtfError::MalformedDocument(
                        "RTF paragraph list level exceeds the nine-level specification limit"
                            .to_string(),
                    ));
                }
                state.paragraph.list_level = Some(level);
            },

            // Tracked revisions
            ControlWord::Revised(value) => {
                if *value {
                    if state.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "conflicting or duplicate RTF revision marker".to_string(),
                        ));
                    }
                    state.revision_type =
                        Some(super::super::super::annotation::RevisionType::Insertion);
                } else if state.revision_type
                    == Some(super::super::super::annotation::RevisionType::Insertion)
                {
                    state.revision_type = None;
                    state.revision_author_id = None;
                    state.revision_date = None;
                    state.revision_event_id = None;
                }
            },
            ControlWord::Deleted(value) => {
                if *value {
                    if state.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "conflicting or duplicate RTF revision marker".to_string(),
                        ));
                    }
                    state.revision_type =
                        Some(super::super::super::annotation::RevisionType::Deletion);
                } else if state.revision_type
                    == Some(super::super::super::annotation::RevisionType::Deletion)
                {
                    state.revision_type = None;
                    state.revision_author_id = None;
                    state.revision_date = None;
                    state.revision_event_id = None;
                }
            },
            ControlWord::RevisionAuthor(value) => {
                if state.revision_type.is_none() || state.revision_author_id.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF revauth requires one active revision marker".to_string(),
                    ));
                }
                state.revision_author_id = Some(*value);
            },
            ControlWord::DeletedRevisionAuthor(value) => {
                if state.revision_type
                    != Some(super::super::super::annotation::RevisionType::Deletion)
                    || state.revision_author_id.is_some()
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF revauthdel requires one active deletion marker".to_string(),
                    ));
                }
                state.revision_author_id = Some(*value);
            },
            ControlWord::RevisionDate(value) => {
                if state.revision_type.is_none() || state.revision_date.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF revdttm requires one active revision marker".to_string(),
                    ));
                }
                state.revision_date = Some(*value);
            },
            ControlWord::DeletedRevisionDate(value) => {
                if state.revision_type
                    != Some(super::super::super::annotation::RevisionType::Deletion)
                    || state.revision_date.is_some()
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF revdttmdel requires one active deletion marker".to_string(),
                    ));
                }
                state.revision_date = Some(*value);
            },

            // Unicode
            ControlWord::UnicodeSkip(n) => state.unicode_skip = *n,
            ControlWord::Unicode(code) => {
                // Unicode characters are handled separately during text parsing
                // since they may span multiple tokens with fallback characters
                // The control word itself doesn't add text here
                let _ = code; // Suppress unused warning
            },

            // Character encoding
            ControlWord::Ansi => {
                state.encoding = RtfEncoding::Standard(Mbcs::WINDOWS_1252);
            },
            ControlWord::AnsiCodePage(cp) => {
                state.encoding = match *cp {
                    437 => RtfEncoding::Cp437,
                    850 => RtfEncoding::Cp850,
                    _ => {
                        let page =
                            u32::try_from(*cp).ok().and_then(Mbcs::new).ok_or_else(|| {
                                RtfError::MalformedDocument(format!(
                                    "unsupported RTF ANSI code page {cp}"
                                ))
                            })?;
                        RtfEncoding::Standard(page)
                    },
                }
            },
            ControlWord::Mac => {
                state.encoding = RtfEncoding::Standard(Mbcs::MACINTOSH);
            },
            ControlWord::Pc => state.encoding = RtfEncoding::Cp437,
            ControlWord::Pca => state.encoding = RtfEncoding::Cp850,

            // Table control words
            ControlWord::InTable => {
                state.in_table = true;
            },
            ControlWord::TableStyle(value) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    state.table_style = Some(table_style_reference(*value)?);
                }
            },
            ControlWord::TableRsid(value) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    state.table_rsid = Some(*value as u32);
                }
            },
            ControlWord::TableRowRevisionAuthor(value) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    state.table_row_revision.author =
                        Some(nonnegative_author_index(*value, "trauth")?);
                }
            },
            ControlWord::TableRowRevisionDate(value) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    state.table_row_revision.date = Some(*value);
                }
            },
            ControlWord::TableRowDefaults => {
                // Start a new row definition
                state.cell_boundaries.clear();
                state.table_style = None;
                state.table_rsid = None;
                state.table_row_padding = Default::default();
                state.table_row_spacing = Default::default();
                state.table_row_positioning = Default::default();
                state.table_row_direction = None;
                state.table_row_layout = Default::default();
                state.table_row_borders = Default::default();
                state.table_row_shading = Default::default();
                state.table_row_geometry = Default::default();
                state.table_default_borders = Default::default();
                state.table_default_padding = Default::default();
                state.table_default_spacing = Default::default();
                state.table_default_width_unit = None;
                state.table_default_width_value = None;
                state.table_autoformat_flags = Default::default();
                state.table_row_banding = Default::default();
                state.table_row_revision = Default::default();
                state.table_row_index_seen = false;
                state.table_row_band_index_seen = false;
                state.table_last_row_seen = false;
                state.table_width_unit = None;
                state.table_width_value = None;
                state.table_leading_width_unit = None;
                state.table_leading_width_value = None;
                state.table_trailing_width_unit = None;
                state.table_trailing_width_value = None;
                state.table_indent_value = None;
                state.table_indent_unit = None;
                state.pending_cell_padding = Default::default();
                state.pending_cell_spacing = Default::default();
                state.pending_cell_layout = Default::default();
                state.pending_cell_merge = Default::default();
                state.pending_cell_revision = None;
                state.pending_cell_borders = Default::default();
                state.pending_cell_shading = Default::default();
                state.pending_cell_width_unit = None;
                state.pending_cell_width_value = None;
                state.table_row_shading_seen = 0;
                state.pending_cell_shading_seen = 0;
                state.active_table_border = None;
                state.active_table_border_seen = 0;
                state.cell_distances.clear();
                state.cell_layouts.clear();
                state.cell_merges.clear();
                state.cell_revisions.clear();
                state.cell_decorations.clear();
                state.cell_widths.clear();
                let destination = state.destination;
                let level = state.table_nesting_level;
                let _ = state;
                if destination == Destination::NestedTableProperties {
                    if level < 2 {
                        return Err(RtfError::MalformedDocument(
                            "RTF nesttableprops lacks itap level 2 or greater".to_string(),
                        ));
                    }
                    let row = &mut self.ensure_nested_builder(level)?.row;
                    row.set_table_style(None);
                    row.set_direction(None);
                    row.set_layout(Default::default());
                } else {
                    self.drain_nested_to(1)?;
                    self.start_table_if_needed();
                    if let Some(row) = &mut self.current_row {
                        row.set_table_style(None);
                        row.set_direction(None);
                        row.set_layout(Default::default());
                    }
                }
            },
            ControlWord::TableRowIndex(value) => {
                if state.table_row_index_seen {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF irow control".to_string(),
                    ));
                }
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument("RTF irow requires a numeric parameter".to_string())
                })?;
                state.table_row_banding.row_index = Some(u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument("RTF irow must be in 0..=65535".to_string())
                })?);
                state.table_row_index_seen = true
            },
            ControlWord::TableRowBandIndex(value) => {
                if state.table_row_band_index_seen {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF irowband control".to_string(),
                    ));
                }
                let value = value.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF irowband requires a numeric parameter".to_string(),
                    )
                })?;
                state.table_row_banding.band_index = Some(if value == -1 {
                    crate::TableRowBandIndex::Header
                } else {
                    crate::TableRowBandIndex::Row(u16::try_from(value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF irowband must be -1 or in 0..=65535".to_string(),
                        )
                    })?)
                });
                state.table_row_band_index_seen = true
            },
            ControlWord::TableLastRow(param) => {
                require_parameterless(*param, "lastrow")?;
                if state.table_last_row_seen {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF lastrow control".to_string(),
                    ));
                }
                state.table_row_banding.last_row = true;
                state.table_last_row_seen = true
            },
            ControlWord::TableAutoformatFlag(flag, param) => {
                require_parameterless(*param, "table autoformat flag")?;
                if !state.table_autoformat_flags.insert(*flag) {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF table autoformat flag".to_string(),
                    ));
                }
            },
            ControlWord::LeftToRightRow(param) => {
                require_parameterless(*param, "ltrrow")?;
                state.table_row_direction = Some(TextDirection::LeftToRight)
            },
            ControlWord::RightToLeftRow(param) => {
                require_parameterless(*param, "rtlrow")?;
                state.table_row_direction = Some(TextDirection::RightToLeft)
            },
            ControlWord::TableRowGap(param) => {
                let value = table_geometry_twips(*param, "trgaph", false)?;
                state
                    .table_row_geometry
                    .set_half_gap_twips(Some(value as u16));
            },
            ControlWord::TableRowLeft(param) => {
                let value = table_geometry_twips(*param, "trleft", true)?;
                state.table_row_geometry.set_left_edge_twips(Some(value));
            },
            ControlWord::TableRowHeight(param) => state
                .table_row_geometry
                .set_height(table_row_height(*param)?),
            ControlWord::TablePreferredWidthUnit(scope, param) => {
                let unit = table_width_unit(*param)?;
                let target = match scope {
                    crate::TableDistanceScope::Row => &mut state.table_width_unit,
                    crate::TableDistanceScope::Cell => &mut state.pending_cell_width_unit,
                };
                if target.replace(unit).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF preferred-width unit is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TablePreferredWidthValue(scope, param) => {
                let value = param.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF preferred-width value requires a parameter".to_string(),
                    )
                })?;
                let target = match scope {
                    crate::TableDistanceScope::Row => &mut state.table_width_value,
                    crate::TableDistanceScope::Cell => &mut state.pending_cell_width_value,
                };
                if target.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF preferred-width value is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableInvisibleWidthUnit(trailing, param) => {
                let unit = table_width_unit(*param)?;
                let target = if *trailing {
                    &mut state.table_trailing_width_unit
                } else {
                    &mut state.table_leading_width_unit
                };
                if target.replace(unit).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF invisible-width unit is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableInvisibleWidthValue(trailing, param) => {
                let value = param.ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF invisible-width value requires a parameter".to_string(),
                    )
                })?;
                let target = if *trailing {
                    &mut state.table_trailing_width_value
                } else {
                    &mut state.table_leading_width_value
                };
                if target.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF invisible-width value is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableAutoFit(param) => {
                state.table_row_geometry.set_auto_fit(match param {
                    Some(0) => false,
                    Some(1) => true,
                    None => {
                        return Err(RtfError::MalformedDocument(
                            "RTF trautofit requires 0 or 1".to_string(),
                        ));
                    },
                    Some(_) => {
                        return Err(RtfError::MalformedDocument(
                            "RTF trautofit accepts only 0 or 1".to_string(),
                        ));
                    },
                })
            },
            ControlWord::TableIndentValue(param) => {
                let value = match param {
                    None => 0,
                    Some(_) => table_geometry_twips(*param, "tblind", true)?,
                };
                if state.table_indent_value.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF tblind is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableIndentUnit(param) => {
                let unit = table_indent_unit(*param)?;
                if state.table_indent_unit.replace(unit).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF tblindtype is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableRowHeader(param) => {
                require_parameterless(*param, "trhdr")?;
                state.table_row_layout.header = true
            },
            ControlWord::TableRowKeep(param) => {
                require_parameterless(*param, "trkeep")?;
                state.table_row_layout.keep_together = true
            },
            ControlWord::TableRowKeepFollow(param) => {
                require_parameterless(*param, "trkeepfollow")?;
                state.table_row_layout.keep_with_following = true
            },
            ControlWord::TableRowAlignment(value, param) => {
                require_parameterless(*param, "table row alignment")?;
                state.table_row_layout.alignment = Some(*value)
            },
            ControlWord::TableCellVerticalAlignment(value, param) => {
                require_parameterless(*param, "cell vertical alignment")?;
                state.pending_cell_layout.vertical_alignment = Some(*value)
            },
            ControlWord::TableCellTextFlow(value, param) => {
                require_parameterless(*param, "cell text flow")?;
                state.pending_cell_layout.text_flow = Some(*value)
            },
            ControlWord::TableCellFitText(param) => {
                require_parameterless(*param, "clFitText")?;
                state.pending_cell_layout.fit_text = true
            },
            ControlWord::TableCellNoWrap(param) => {
                require_parameterless(*param, "clNoWrap")?;
                state.pending_cell_layout.no_wrap = true
            },
            ControlWord::TableCellHideMark(param) => {
                require_parameterless(*param, "clhidemark")?;
                state.pending_cell_layout.hide_mark = true
            },
            ControlWord::TableCellMerge(axis, role, param) => {
                require_parameterless(*param, "table cell merge")?;
                let pending = match axis {
                    crate::TableCellMergeAxis::Horizontal => {
                        &mut state.pending_cell_merge.horizontal
                    },
                    crate::TableCellMergeAxis::Vertical => &mut state.pending_cell_merge.vertical,
                };
                if pending.replace(*role).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF cell definition has duplicate or conflicting merge roles on one axis"
                            .to_string(),
                    ));
                }
            },
            ControlWord::CellRevisionMark(kind) => match &mut state.pending_cell_revision {
                Some(revision) if revision.kind != *kind => {
                    return Err(RtfError::MalformedDocument(
                        "RTF cell definition has conflicting revision markers".to_string(),
                    ));
                },
                Some(_) => {},
                slot @ None => {
                    *slot = Some(crate::CellRevision {
                        kind: *kind,
                        metadata: crate::RevisionMetadata::default(),
                    });
                },
            },
            ControlWord::CellRevisionAuthor(kind, value) => {
                let revision = pending_cell_revision(state, *kind, kind.author_control_word())?;
                revision.metadata.author = Some(nonnegative_author_index(
                    *value,
                    kind.author_control_word(),
                )?);
            },
            ControlWord::CellRevisionDate(kind, value) => {
                let revision = pending_cell_revision(state, *kind, kind.date_control_word())?;
                revision.metadata.date = Some(*value);
            },
            ControlWord::TableRightToLeft(value) => {
                let direction = Some(if *value {
                    TextDirection::RightToLeft
                } else {
                    TextDirection::LeftToRight
                });
                let destination = state.destination;
                let level = state.table_nesting_level;
                let _ = state;
                if destination == Destination::NestedTableProperties {
                    self.ensure_nested_builder(level)?
                        .table
                        .set_direction(direction);
                } else {
                    self.start_table_if_needed();
                    if let Some(table) = &mut self.current_table {
                        table.set_direction(direction);
                    }
                }
            },
            ControlWord::CellX(boundary) => {
                // Cell boundary definition
                state.cell_boundaries.push(*boundary);
                if state.cell_distances.len() >= crate::MAX_TABLE_CELLS_PER_ROW {
                    return Err(RtfError::MalformedDocument(
                        "RTF row exceeds 4096 cell definitions".to_string(),
                    ));
                }
                let width = resolve_preferred_width(
                    state.pending_cell_width_unit.take(),
                    state.pending_cell_width_value.take(),
                )?;
                state.cell_widths.push(width);
                state.cell_distances.push((
                    std::mem::take(&mut state.pending_cell_padding),
                    std::mem::take(&mut state.pending_cell_spacing),
                ));
                state
                    .cell_layouts
                    .push(std::mem::take(&mut state.pending_cell_layout));
                state
                    .cell_merges
                    .push(std::mem::take(&mut state.pending_cell_merge));
                state
                    .cell_revisions
                    .push(state.pending_cell_revision.take());
                state.cell_decorations.push((
                    std::mem::take(&mut state.pending_cell_borders),
                    std::mem::take(&mut state.pending_cell_shading),
                ));
                state.pending_cell_shading_seen = 0;
                state.active_table_border = None;
                state.active_table_border_seen = 0;
            },
            ControlWord::TableDistanceValue(target, value) => {
                apply_table_distance(state, *target, *value, false)?
            },
            ControlWord::TableDistanceUnit(target, value) => {
                apply_table_distance(state, *target, *value, true)?
            },
            ControlWord::TableDefaultDistanceValue(kind, edge, value) => {
                let distances = match kind {
                    crate::TableDistanceKind::Padding => &mut state.table_default_padding,
                    crate::TableDistanceKind::Spacing => &mut state.table_default_spacing,
                };
                apply_table_distance_side(distances, *edge, *value, false)?
            },
            ControlWord::TableDefaultDistanceUnit(kind, edge, value) => {
                let distances = match kind {
                    crate::TableDistanceKind::Padding => &mut state.table_default_padding,
                    crate::TableDistanceKind::Spacing => &mut state.table_default_spacing,
                };
                apply_table_distance_side(distances, *edge, *value, true)?
            },
            ControlWord::TableDefaultCellWidthUnit(param) => {
                let unit = table_width_unit(*param)?;
                if state.table_default_width_unit.replace(unit).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF tscellwidthfts is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableDefaultCellWidthValue(param) => {
                let value = param.ok_or_else(|| {
                    RtfError::MalformedDocument("RTF tscellwidth requires a parameter".to_string())
                })?;
                if state.table_default_width_value.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF tscellwidth is duplicated".to_string(),
                    ));
                }
            },
            ControlWord::TableHorizontalReference(value, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    require_parameterless(*param, "floating-table horizontal reference")?;
                    state.table_row_positioning.horizontal_reference = Some(*value)
                }
            },
            ControlWord::TableVerticalReference(value, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    require_parameterless(*param, "floating-table vertical reference")?;
                    state.table_row_positioning.vertical_reference = Some(*value)
                }
            },
            ControlWord::TableHorizontalPosition(value, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    require_parameterless(*param, "floating-table horizontal position")?;
                    state.table_row_positioning.horizontal_position = Some(*value)
                }
            },
            ControlWord::TableVerticalPosition(value, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    require_parameterless(*param, "floating-table vertical position")?;
                    state.table_row_positioning.vertical_position = Some(*value)
                }
            },
            ControlWord::TableHorizontalOffset(negative, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    let value = floating_table_offset(*param, *negative, "horizontal")?;
                    state.table_row_positioning.horizontal_position = Some(if *negative {
                        crate::TableHorizontalPosition::NegativeOffset(value)
                    } else {
                        crate::TableHorizontalPosition::Offset(value)
                    })
                }
            },
            ControlWord::TableVerticalOffset(negative, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    let value = floating_table_offset(*param, *negative, "vertical")?;
                    state.table_row_positioning.vertical_position = Some(if *negative {
                        crate::TableVerticalPosition::NegativeOffset(value)
                    } else {
                        crate::TableVerticalPosition::Offset(value)
                    })
                }
            },
            ControlWord::TableWrapDistance(edge, param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    *state.table_row_positioning.wrap_distances.side_mut(*edge) =
                        Some(floating_table_wrap_distance(*param)?)
                }
            },
            ControlWord::TableNoOverlap(param) => {
                if matches!(
                    state.destination,
                    Destination::DocumentBody | Destination::NestedTableProperties
                ) {
                    state.table_row_positioning.no_overlap = match *param {
                        None | Some(1) => true,
                        Some(0) => false,
                        Some(_) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF tabsnoovrlp accepts only 0 or 1".to_string(),
                            ));
                        },
                    }
                }
            },
            ControlWord::NestedTableCell(param) => {
                require_parameterless(*param, "nestcell")?;
                let destination = state.destination;
                let level = state.table_nesting_level;
                let _ = state;
                if destination != Destination::DocumentBody || level < 2 {
                    return Err(RtfError::MalformedDocument(
                        "RTF nestcell requires visible nested-table text and itap 2 or greater"
                            .to_string(),
                    ));
                }
                self.finalize_nested_cell(level)?;
            },
            ControlWord::NestedTableRow(param) => {
                require_parameterless(*param, "nestrow")?;
                let destination = state.destination;
                let level = state.table_nesting_level;
                let _ = state;
                if destination != Destination::NestedTableProperties || level < 2 {
                    return Err(RtfError::MalformedDocument(
                        "RTF nestrow requires a nesttableprops destination and itap 2 or greater"
                            .to_string(),
                    ));
                }
                self.finalize_nested_row(level)?;
            },
            ControlWord::NestedTableProperties(_) | ControlWord::NoNestedTables(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF nested-table destination control is misplaced".to_string(),
                ));
            },
            ControlWord::TableCell => {
                // Cell break - finalize current cell
                self.start_table_if_needed();
                self.finalize_cell(true)?;
            },
            ControlWord::TableRow => {
                // Row break - finalize current row
                let row_geometry = resolve_row_geometry(state)?;
                let row_cell_defaults = crate::TableRowCellDefaults {
                    borders: state.table_default_borders.clone(),
                    padding: state.table_default_padding.clone(),
                    spacing: state.table_default_spacing.clone(),
                    preferred_cell_width: resolve_preferred_width(
                        state.table_default_width_unit,
                        state.table_default_width_value,
                    )?,
                };
                let table_style = state.table_style;
                let table_rsid = state.table_rsid;
                let row_padding = state.table_row_padding.clone();
                let row_spacing = state.table_row_spacing.clone();
                let row_positioning = state.table_row_positioning.clone();
                let row_direction = state.table_row_direction;
                let row_layout = state.table_row_layout;
                let row_borders = state.table_row_borders.clone();
                let row_shading = state.table_row_shading;
                let autoformat_flags = state.table_autoformat_flags;
                let banding = state.table_row_banding;
                let row_revision = state.table_row_revision;
                let boundaries = state.cell_boundaries.clone();
                let cell_distances = state.cell_distances.clone();
                let cell_layouts = state.cell_layouts.clone();
                let cell_merges = state.cell_merges.clone();
                let cell_revisions = state.cell_revisions.clone();
                let cell_decorations = state.cell_decorations.clone();
                let cell_widths = state.cell_widths.clone();
                let _ = state;
                self.drain_nested_to(1)?;
                self.finalize_cell(false)?;
                if let Some(row) = &mut self.current_row {
                    if !boundaries.is_empty() && boundaries.len() != row.cell_count() {
                        return Err(RtfError::MalformedDocument(
                            "RTF row cell boundaries do not match cell count".to_string(),
                        ));
                    }
                    for (index, cell) in row.cells_mut().iter_mut().enumerate() {
                        if let Some((padding, spacing)) = cell_distances.get(index) {
                            cell.set_padding(padding.clone());
                            cell.set_spacing(spacing.clone());
                        }
                        if let Some(layout) = cell_layouts.get(index) {
                            cell.set_layout(*layout);
                        }
                        if let Some(merge) = cell_merges.get(index) {
                            cell.set_merge(*merge);
                        }
                        if let Some(revision) = cell_revisions.get(index) {
                            cell.set_revision(*revision);
                        }
                        cell.set_right_boundary(boundaries.get(index).copied());
                        cell.set_preferred_width(cell_widths.get(index).copied().flatten());
                        if let Some((borders, shading)) = cell_decorations.get(index) {
                            cell.set_borders(borders.clone());
                            cell.set_shading(*shading);
                        }
                    }
                    row.set_table_style(table_style);
                    row.set_table_rsid(table_rsid);
                    row.set_direction(row_direction);
                    row.set_layout(row_layout);
                    row.set_padding(row_padding);
                    row.set_spacing(row_spacing);
                    row.set_cell_defaults(row_cell_defaults);
                    row.set_positioning(row_positioning);
                    row.set_borders(row_borders);
                    row.set_shading(row_shading);
                    row.set_geometry(row_geometry);
                    row.set_autoformat_flags(autoformat_flags);
                    row.set_banding(banding);
                    row.set_revision(row_revision);
                }
                self.finalize_row()?;
            },

            ControlWord::ProtectionUserTable
            | ControlWord::NextFile
            | ControlWord::DocumentTemplate
            | ControlWord::WindowCaption
            | ControlWord::XslTransform
            | ControlWord::StyleListFilter(_)
            | ControlWord::WriteReservation(_)
            | ControlWord::WriteReservationHash(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF destination control is misplaced".to_string(),
                ));
            },
            _ => {
                // Ignore unknown or unhandled control words
            },
        }

        Ok(())
    }
}
