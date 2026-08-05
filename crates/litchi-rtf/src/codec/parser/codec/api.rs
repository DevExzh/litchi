use super::*;

impl<'a> Parser<'a> {
    /// Create a new parser.
    pub fn new(tokens: &'a [Token<'a>], arena: &'a Bump) -> Self {
        Self {
            tokens,
            pos: 0,
            states: vec![State::default()],
            font_table: RefCell::new(FontTable::new()),
            saw_font_table: false,
            file_table: None,
            unicode_alternate_depth: 0,
            color_table: RefCell::new(ColorTable::new()),
            blocks: Vec::new(),
            arena,
            tables: Vec::new(),
            current_table: None,
            current_row: None,
            current_cell_text: SmallVec::new(),
            current_cell_nested: Vec::new(),
            current_cell_drawings: DrawingStoryCapture::default(),
            current_cell_story_events: Vec::new(),
            nested_table_builders: Vec::new(),
            logical_table_count: 0,
            pictures: Vec::new(),
            picture_compatibility_records: Vec::new(),
            fields: Vec::new(),
            field_drawing_captures: Vec::new(),
            form_fields: Vec::new(),
            form_field_text_bytes: 0,
            generator: None,
            revision_save_ids: Vec::new(),
            saw_revision_save_table: false,
            revision_save_root: None,
            saw_revision_save_root: false,
            xml_namespaces: Vec::new(),
            saw_xml_namespace_table: false,
            xml_namespace_text_bytes: 0,
            custom_xml_tags: Vec::new(),
            open_custom_xml_tags: Vec::new(),
            custom_xml_spans: Vec::new(),
            pending_custom_xml_attribute: None,
            next_custom_xml_order: 0,
            custom_xml_text_bytes: 0,
            math_zones: Vec::new(),
            math_text_bytes: 0,
            protection_ranges: Vec::new(),
            open_protection_ranges: HashMap::new(),
            protection_range_spans: Vec::new(),
            next_protection_range_order: 0,
            editable_regions: Vec::new(),
            open_editable_regions: Vec::new(),
            editable_region_spans: Vec::new(),
            next_editable_region_order: 0,
            protection_users: Vec::new(),
            saw_protection_user_table: false,
            protection_user_text_bytes: 0,
            hyphenation: crate::DocumentHyphenation::default(),
            hyphenation_seen: 0,
            external_references: crate::DocumentExternalReferences::default(),
            document_view: crate::DocumentView::default(),
            document_view_seen: 0,
            review_display: crate::DocumentReviewDisplay::default(),
            review_display_seen: 0,
            window_caption: None,
            kinsoku: crate::DocumentKinsoku::default(),
            xsl_transform: None,
            xsl_transform_usage: crate::DocumentXslTransformUsage::default(),
            use_xsl_transform_seen: false,
            style_list_filter: None,
            style_sort_method: None,
            style_sort_method_seen: false,
            save_preferences: crate::DocumentSavePreferences::default(),
            save_preferences_seen: 0,
            write_reservations: crate::DocumentWriteReservations::default(),
            origin_metadata: crate::DocumentOriginMetadata::default(),
            file_settings: crate::DocumentFileSettings::default(),
            file_settings_seen: 0,
            output_settings: crate::DocumentOutputSettings::default(),
            output_settings_seen: 0,
            rendering_settings: crate::DocumentRenderingSettings::default(),
            rendering_settings_seen: 0,
            processing_settings: crate::DocumentProcessingSettings::default(),
            processing_settings_seen: 0,
            drawing_grid: crate::DocumentDrawingGrid::default(),
            drawing_grid_seen: 0,
            print_layout_settings: crate::DocumentPrintLayoutSettings::default(),
            print_layout_settings_seen: 0,
            section_gutter_overrides: Vec::new(),
            theme_languages: crate::DocumentThemeLanguages::default(),
            theme_languages_seen: 0,
            xml_policies: crate::DocumentXmlPolicies::default(),
            xml_policies_seen: 0,
            embedding_policies: crate::DocumentEmbeddingPolicies::default(),
            embedding_policies_seen: 0,
            revision_policies: crate::DocumentRevisionPolicies::default(),
            revision_policies_seen: 0,
            style_policies: crate::DocumentStylePolicies::default(),
            style_policies_seen: 0,
            style_restrictions: crate::DocumentStyleRestrictions::default(),
            style_restrictions_seen: 0,
            booklet_printing: crate::DocumentBookletPrinting::default(),
            booklet_printing_seen: 0,
            privacy_policies: crate::DocumentPrivacyPolicies::default(),
            privacy_policies_seen: 0,
            line_spacing_compatibility: crate::DocumentLineSpacingCompatibility::default(),
            line_spacing_compatibility_seen: 0,
            east_asian_compatibility: crate::DocumentEastAsianCompatibility::default(),
            east_asian_compatibility_seen: 0,
            table_layout_compatibility: crate::DocumentTableLayoutCompatibility::default(),
            table_layout_compatibility_seen: 0,
            legacy_layout_compatibility: crate::DocumentLegacyLayoutCompatibility::default(),
            legacy_layout_compatibility_seen: 0,
            asian_grid_compatibility: crate::DocumentAsianGridCompatibility::default(),
            asian_grid_compatibility_seen: 0,
            compatibility_policy: crate::DocumentCompatibilityPolicy::default(),
            compatibility_policy_seen: 0,
            word_2003_compatibility: crate::DocumentWord2003Compatibility::default(),
            word_2003_compatibility_seen: 0,
            theme_data: None,
            saw_theme_data: false,
            color_scheme_mapping: None,
            saw_color_scheme_mapping: false,
            latent_styles: None,
            data_store: None,
            saw_data_store: false,
            mail_merge: None,
            math_properties: None,
            default_tab_width_twips: None,
            language_defaults: crate::DocumentLanguageDefaults::default(),
            default_formatting: crate::DocumentDefaultFormatting::default(),
            default_font_selectors_seen: 0,
            saw_info_group: false,
            document_direction: None,
            gutter_on_right: false,
            objects: Vec::new(),
            document_variables: Vec::new(),
            document_variable_text_bytes: 0,
            user_properties: Vec::new(),
            user_property_text_bytes: 0,
            navigation_entries: Vec::new(),
            navigation_entry_text_bytes: 0,
            generated_list_markers: Vec::new(),
            generated_list_marker_text_bytes: 0,
            saw_user_properties: false,
            list_table: super::super::super::list::ListTable::new(),
            saw_list_table: false,
            list_override_table: super::super::super::list::ListOverrideTable::new(),
            saw_list_override_table: false,
            legacy_section_numbering: crate::LegacySectionNumbering::new(),
            legacy_paragraph_numbering: Vec::new(),
            paragraph_group_table: None,
            sections: Vec::new(),
            section_properties_active: false,
            section_note_options_closed: false,
            root_section_format_run: false,
            bookmarks: super::super::super::bookmark::BookmarkTable::new(),
            open_bookmarks: HashMap::new(),
            bookmark_spans: Vec::new(),
            body_text_len: 0,
            body_boundaries: Vec::new(),
            next_bookmark_order: 0,
            shapes: Vec::new(),
            drawing_order: Vec::new(),
            body_story_events: Vec::new(),
            revision_event_indices: Vec::new(),
            background_shape_index: None,
            legacy_text_boxes: Vec::new(),
            legacy_drawings: Vec::new(),
            legacy_text_box_text_bytes: 0,
            legacy_drawing_primitives: 0,
            legacy_drawing_points: 0,
            shape_groups: Vec::new(),
            stylesheet: super::super::super::stylesheet::StyleSheet::new(),
            saw_stylesheet: false,
            info: super::super::super::info::DocumentInfo::new(),
            annotations: Vec::new(),
            annotation_ranges: HashMap::new(),
            pending_annotation_author: String::new(),
            pending_annotation_author_seen: false,
            pending_annotation_initials: String::new(),
            pending_annotation_initials_seen: false,
            pending_annotation_mark: false,
            notes: Vec::new(),
            note_options: crate::NoteOptions::default(),
            note_options_closed: false,
            note_separators: crate::NoteSeparatorTable::new(),
            current_note_separator_active: false,
            current_note_separator_elements: Vec::new(),
            current_note_separator_drawings: DrawingStoryCapture::default(),
            revisions: Vec::new(),
            revision_authors: Vec::new(),
            saw_revision_table: false,
            revision_author_text_bytes: 0,
            revision_text_bytes: 0,
            current_header_footer: None,
            current_note_buffer: SmallVec::new(),
            current_note_shapes: Vec::new(),
            current_note_shape_groups: Vec::new(),
            current_note_drawing_order: Vec::new(),
            current_note_story_events: Vec::new(),
            current_hf_shapes: Vec::new(),
            current_hf_shape_groups: Vec::new(),
            current_hf_drawing_order: Vec::new(),
            current_hf_story_events: Vec::new(),
            current_hf_story_offset: 0,
            current_hf_type: None,
        }
    }

    /// Parse the token stream into a document.
    pub fn parse(mut self) -> RtfResult<ParsedDocument<'a>> {
        // Validate document structure
        if self.tokens.is_empty() {
            return Err(RtfError::MalformedDocument(
                "Empty token stream".to_string(),
            ));
        }
        #[derive(Clone, Copy)]
        struct NoteGuardContext {
            body_flow: bool,
            visible_field_result: bool,
            visible_section_format: bool,
            direct_header_footer: bool,
            direct_field_instruction: bool,
            inert_section_format: bool,
        }

        let mut contexts: Vec<NoteGuardContext> = Vec::new();
        for (index, token) in self.tokens.iter().enumerate() {
            match token {
                Token::OpenBrace => {
                    if let Some(parent) = contexts.last_mut() {
                        parent.inert_section_format = false;
                    }
                    let parent = contexts.last().copied().unwrap_or(NoteGuardContext {
                        body_flow: true,
                        visible_field_result: false,
                        visible_section_format: false,
                        direct_header_footer: false,
                        direct_field_instruction: false,
                        inert_section_format: false,
                    });
                    let mut destination_index = index + 1;
                    let starred = matches!(
                        self.tokens.get(destination_index),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    );
                    if starred {
                        destination_index += 1;
                    }
                    let destination = match self.tokens.get(destination_index) {
                        Some(Token::Control(control)) => Some(control),
                        _ => None,
                    };
                    let context = match destination {
                        Some(ControlWord::FieldResult) if parent.body_flow => NoteGuardContext {
                            body_flow: true,
                            visible_field_result: true,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::Field) => NoteGuardContext {
                            body_flow: parent.body_flow,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::SectionDefault) if parent.body_flow && !starred => {
                            NoteGuardContext {
                                body_flow: true,
                                visible_field_result: parent.visible_field_result,
                                visible_section_format: true,
                                direct_header_footer: false,
                                direct_field_instruction: false,
                                inert_section_format: false,
                            }
                        },
                        Some(
                            ControlWord::Header
                            | ControlWord::HeaderFirst
                            | ControlWord::HeaderLeft
                            | ControlWord::HeaderRight
                            | ControlWord::Footer
                            | ControlWord::FooterFirst
                            | ControlWord::FooterLeft
                            | ControlWord::FooterRight,
                        ) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: true,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::FieldInstruction) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: true,
                            inert_section_format: false,
                        },
                        Some(
                            ControlWord::Annotation
                            | ControlWord::Footnote
                            | ControlWord::Endnote
                            | ControlWord::Object
                            | ControlWord::Result
                            | ControlWord::Picture
                            | ControlWord::Shape(_)
                            | ControlWord::ShapeGroup(_),
                        ) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        _ if starred => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        _ => NoteGuardContext {
                            body_flow: parent.body_flow,
                            visible_field_result: parent.visible_field_result,
                            visible_section_format: parent.visible_section_format,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                    };
                    contexts.push(context);
                },
                Token::CloseBrace => {
                    contexts.pop();
                    if contexts.is_empty() {
                        break;
                    }
                },
                Token::Control(ControlWord::SectionDefault) => {
                    if let Some(context) = contexts.last_mut()
                        && (context.direct_header_footer || context.direct_field_instruction)
                    {
                        context.inert_section_format = true;
                    }
                },
                Token::Control(ControlWord::SectionBreak) => {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Text(text) if !text.is_empty() => {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Control(control)
                    if matches!(
                        control,
                        ControlWord::Par
                            | ControlWord::Line
                            | ControlWord::Tab
                            | ControlWord::Unicode(_)
                    ) || control_symbol_text(control).is_some() =>
                {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Control(
                    ControlWord::NoteKinds(_)
                    | ControlWord::FootnotePlacement(_)
                    | ControlWord::EndnotePlacement(_)
                    | ControlWord::FootnoteStart(_)
                    | ControlWord::EndnoteStart(_)
                    | ControlWord::FootnoteRestart(_)
                    | ControlWord::EndnoteRestart(_)
                    | ControlWord::FootnoteNumbering(_)
                    | ControlWord::EndnoteNumbering(_),
                ) if contexts.len() != 1 => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note options must be root document-format controls".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::SectionFootnotePlacement(_)
                    | ControlWord::SectionFootnoteStart(_)
                    | ControlWord::SectionEndnoteStart(_)
                    | ControlWord::SectionFootnoteRestart(_)
                    | ControlWord::SectionEndnoteRestart(_)
                    | ControlWord::SectionFootnoteNumbering(_)
                    | ControlWord::SectionEndnoteNumbering(_),
                ) if contexts.len() != 1
                    && !contexts.last().is_some_and(|context| {
                        context.visible_field_result
                            || context.visible_section_format
                            || context.inert_section_format
                    }) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF section note options must be root controls or visible field-result formatting"
                            .to_string(),
                    ));
                },
                _ => {},
            }
        }

        // Expect opening brace
        if !matches!(self.tokens.first(), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "Document must start with {".to_string(),
            ));
        }

        // Parse document content
        self.parse_group()?;

        // Finalize any remaining table
        self.finalize_table()?;
        self.finalize_bookmarks()?;
        self.finalize_custom_xml_tags()?;
        self.finalize_protection_ranges()?;
        self.finalize_editable_regions()?;
        self.finalize_annotations()?;
        crate::story::validate_boundaries(&self.blocks, &self.body_boundaries)?;
        let body_story_events = self.finalize_body_story_events()?;

        let revision_save = if self.saw_revision_save_table || self.saw_revision_save_root {
            Some(crate::RevisionSaveMetadata::new(
                self.revision_save_ids,
                self.revision_save_root,
            )?)
        } else {
            None
        };
        let theme = match (self.theme_data, self.color_scheme_mapping) {
            (Some(data), mapping) => Some(crate::DocumentTheme::new(
                Cow::Owned(data),
                mapping.map(Cow::Owned),
            )?),
            (None, Some(_)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF color-scheme mapping is orphaned without theme data".to_string(),
                ));
            },
            (None, None) => None,
        };
        let data_store = self
            .data_store
            .map(|data| crate::DocumentDataStore::new(Cow::Owned(data)))
            .transpose()?;
        let protection_user_table = self
            .saw_protection_user_table
            .then(|| crate::ProtectionUserTable::new(self.protection_users))
            .transpose()?;

        for section in &self.sections {
            section.properties.columns.validate()?;
        }

        Ok(ParsedDocument {
            font_table: self.font_table.into_inner(),
            file_table: self.file_table,
            color_table: self.color_table.into_inner(),
            blocks: self.blocks,
            tables: self.tables,
            pictures: self.pictures,
            picture_compatibility_records: self.picture_compatibility_records,
            fields: self.fields,
            form_fields: self.form_fields,
            generator: self.generator,
            revision_save,
            xml_namespaces: self.xml_namespaces,
            saw_xml_namespace_table: self.saw_xml_namespace_table,
            custom_xml_tags: self.custom_xml_tags,
            math_zones: self.math_zones,
            protection_ranges: self.protection_ranges,
            editable_regions: self.editable_regions,
            protection_user_table,
            hyphenation: self.hyphenation,
            external_references: self.external_references,
            document_view: self.document_view,
            review_display: self.review_display,
            window_caption: self.window_caption,
            kinsoku: self.kinsoku,
            xsl_transform: self.xsl_transform,
            xsl_transform_usage: self.xsl_transform_usage,
            style_list_filter: self.style_list_filter,
            style_sort_method: self.style_sort_method,
            save_preferences: self.save_preferences,
            write_reservations: self.write_reservations,
            origin_metadata: self.origin_metadata,
            file_settings: self.file_settings,
            output_settings: self.output_settings,
            rendering_settings: self.rendering_settings,
            processing_settings: self.processing_settings,
            drawing_grid: self.drawing_grid,
            print_layout_settings: self.print_layout_settings,
            theme_languages: self.theme_languages,
            xml_policies: self.xml_policies,
            embedding_policies: self.embedding_policies,
            revision_policies: self.revision_policies,
            style_policies: self.style_policies,
            style_restrictions: self.style_restrictions,
            booklet_printing: self.booklet_printing,
            privacy_policies: self.privacy_policies,
            line_spacing_compatibility: self.line_spacing_compatibility,
            east_asian_compatibility: self.east_asian_compatibility,
            table_layout_compatibility: self.table_layout_compatibility,
            legacy_layout_compatibility: self.legacy_layout_compatibility,
            asian_grid_compatibility: self.asian_grid_compatibility,
            compatibility_policy: self.compatibility_policy,
            word_2003_compatibility: self.word_2003_compatibility,
            theme,
            latent_styles: self.latent_styles,
            data_store,
            mail_merge: self.mail_merge,
            math_properties: self.math_properties,
            default_tab_width_twips: self.default_tab_width_twips,
            language_defaults: self.language_defaults,
            default_formatting: self.default_formatting,
            document_direction: self.document_direction,
            gutter_on_right: self.gutter_on_right,
            objects: self.objects,
            document_variables: self.document_variables,
            user_properties: self.user_properties,
            navigation_entries: self.navigation_entries,
            generated_list_markers: self.generated_list_markers,
            list_table: self.list_table,
            list_override_table: self.list_override_table,
            legacy_section_numbering: self.legacy_section_numbering,
            legacy_paragraph_numbering: self.legacy_paragraph_numbering,
            paragraph_group_table: self.paragraph_group_table,
            sections: self.sections,
            bookmarks: self.bookmarks,
            shapes: self.shapes,
            drawing_order: self.drawing_order,
            body_boundaries: self.body_boundaries,
            body_story_events,
            background_shape_index: self.background_shape_index,
            legacy_text_boxes: self.legacy_text_boxes,
            legacy_drawings: self.legacy_drawings,
            shape_groups: self.shape_groups,
            stylesheet: self.stylesheet,
            info: self.info,
            annotations: self.annotations,
            notes: self.notes,
            note_options: self.note_options,
            note_separators: self.note_separators,
            revisions: self.revisions,
            revision_authors: self.revision_authors,
        })
    }
}
