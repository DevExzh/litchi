//! RTF document-header output.

#![allow(
    clippy::shadow_reuse,
    reason = "serialization helpers deliberately rebind a working value as the output is assembled"
)]
use super::super::{
    AbstractNumberingCleanupStatus, Charset, DefaultTabWidthPolicy, DocumentAsianGridCompatibility,
    DocumentAutoFormatType, DocumentBookletPrinting, DocumentCompatibilityPolicy,
    DocumentDefaultFonts, DocumentDrawingGrid, DocumentEastAsianCompatibility,
    DocumentEmbeddingPolicies, DocumentExternalReferences, DocumentFileSettings,
    DocumentHyphenation, DocumentJustificationMode, DocumentKinsoku, DocumentLanguageDefaults,
    DocumentLegacyLayoutCompatibility, DocumentLineSpacingCompatibility, DocumentOrigin,
    DocumentOutputSettings, DocumentPrintLayoutSettings, DocumentPrivacyPolicies,
    DocumentProcessingSettings, DocumentReadOnlyRecommendation, DocumentRenderingOrientation,
    DocumentRenderingSettings, DocumentReviewDisplay, DocumentRevisionPolicies,
    DocumentSavePreferences, DocumentStyleListFilter, DocumentStylePolicies,
    DocumentStyleRestrictions, DocumentStyleSortMethod, DocumentTableLayoutCompatibility,
    DocumentThemeLanguages, DocumentThumbnailPreference, DocumentView, DocumentWindowCaption,
    DocumentWord2003Compatibility, DocumentWriteReservations, DocumentXmlPolicies,
    DocumentXslTransform, DocumentXslTransformUsage, HtmlEmailVersion, MAX_DEFAULT_TAB_WIDTH_TWIPS,
    RtfDocument, RtfWriter, TextDirection, Write, io,
};

impl<W: Write> RtfWriter<W> {
    /// Write document header
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_header(&mut self) -> io::Result<()> {
        let default_tab_width = match self.options.default_tab_width {
            DefaultTabWidthPolicy::PreserveDocument => None,
            DefaultTabWidthPolicy::Override(width) => Some(width),
        };
        self.write_document_header_with_origin(None, default_tab_width, None)
    }

    pub(in super::super) fn write_document_header_with_origin(
        &mut self,
        origin: Option<DocumentOrigin>,
        default_tab_width: Option<u32>,
        default_fonts: Option<&DocumentDefaultFonts>,
    ) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("rtf", Some(1))?;

        match self.options.charset {
            Charset::Ansi(page) => {
                let parameter = i32::try_from(page.id()).map_err(|_err| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF code page exceeds i32")
                })?;
                self.write_control_word("ansi", None)?;
                self.write_control_word("ansicpg", Some(parameter))?;
            },
            Charset::Mac => self.write_control_word("mac", None)?,
            Charset::Pc => self.write_control_word("pc", None)?,
            Charset::Pca => self.write_control_word("pca", None)?,
        }

        match origin {
            Some(DocumentOrigin::PlainTextEmail) => {
                self.write_control_word("fromtext", None)?;
            },
            Some(DocumentOrigin::HtmlEmail { version }) => {
                self.write_control_word("fromhtml", version.map(HtmlEmailVersion::rtf_value))?;
            },
            None => {},
        }

        self.write_control_word(
            "deff",
            Some(i32::from(
                default_fonts
                    .and_then(|fonts| fonts.primary)
                    .unwrap_or(self.options.default_font),
            )),
        )?;
        if let Some(fonts) = default_fonts {
            if let Some(value) = fonts.associated {
                self.write_control_word("adeff", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_double_byte {
                self.write_control_word("stshfdbch", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_low_ansi {
                self.write_control_word("stshfloch", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_high_ansi {
                self.write_control_word("stshfhich", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_bidi {
                self.write_control_word("stshfbi", Some(i32::from(value)))?;
            }
        }
        if let Some(width) = default_tab_width {
            let width = i32::try_from(width).map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!("RTF deftab width {width} exceeds {MAX_DEFAULT_TAB_WIDTH_TWIPS}"),
                )
            })?;
            self.write_control_word("deftab", Some(width))?;
        }

        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_auto_format_type(
        &mut self,
        document_type: Option<DocumentAutoFormatType>,
    ) -> io::Result<()> {
        if let Some(document_type) = document_type {
            self.write_control_word("doctype", Some(document_type.rtf_value()))?;
        }
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_language_defaults(
        &mut self,
        defaults: &DocumentLanguageDefaults,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(language) = defaults.primary {
            self.write_control_word("deflang", Some(language.rtf_value()))?;
        }
        if let Some(language) = defaults.east_asian {
            self.write_control_word("deflangfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = defaults.complex_script {
            self.write_control_word("adeflang", Some(language.rtf_value()))?;
        }
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_direction(&mut self, doc: &RtfDocument<'_>) -> io::Result<()> {
        if let Some(direction) = doc.document_direction() {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrdoc",
                    TextDirection::RightToLeft => "rtldoc",
                },
                None,
            )?;
        }
        if doc.gutter_on_right() {
            self.write_control_word("rtlgutter", None)?;
        }
        Ok(())
    }

    /// Write explicit passive document-level hyphenation settings.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_hyphenation(
        &mut self,
        hyphenation: &DocumentHyphenation,
    ) -> io::Result<()> {
        hyphenation
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(value) = hyphenation.hot_zone_twips {
            self.write_control_word("hyphhotz", Some(value.cast_signed()))?;
        }
        if let Some(value) = hyphenation.consecutive_line_limit {
            self.write_control_word("hyphconsec", Some(value.cast_signed()))?;
        }
        if let Some(value) = hyphenation.capitalized_words {
            self.write_control_word("hyphcaps", Some(i32::from(value)))?;
        }
        if let Some(value) = hyphenation.automatic {
            self.write_control_word("hyphauto", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write inert names without opening, resolving, or invoking referenced files.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_external_references(
        &mut self,
        references: &DocumentExternalReferences<'_>,
    ) -> io::Result<()> {
        references
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (control, value) in [
            ("nextfile", references.next_file.as_deref()),
            ("template", references.template.as_deref()),
        ] {
            let Some(value) = value else { continue };
            self.write_str("{\\*")?;
            self.write_control_word(control, None)?;
            self.write_str(" ")?;
            self.write_destination_text(value)?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write passive compatibility and output flags in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_output_settings(
        &mut self,
        settings: &DocumentOutputSettings,
    ) -> io::Result<()> {
        if settings.word97_compatibility_marker {
            self.write_control_word("muser", None)?;
        }
        if settings.postscript_over_text {
            self.write_control_word("psover", None)?;
        }
        Ok(())
    }

    /// Write passive rendering flags in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_rendering_settings(
        &mut self,
        settings: &DocumentRenderingSettings,
    ) -> io::Result<()> {
        if let Some(orientation) = settings.orientation {
            self.write_control_word(
                match orientation {
                    DocumentRenderingOrientation::Horizontal => "horzdoc",
                    DocumentRenderingOrientation::Vertical => "vertdoc",
                },
                None,
            )?;
        }
        if let Some(justification_mode) = settings.justification_mode {
            self.write_control_word(
                match justification_mode {
                    DocumentJustificationMode::Compress => "jcompress",
                    DocumentJustificationMode::Expand => "jexpand",
                },
                None,
            )?;
        }
        if settings.line_based_on_grid {
            self.write_control_word("lnongrid", None)?;
        }
        Ok(())
    }

    /// Write passive printing, cleanup, and event properties in stable order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_processing_settings(
        &mut self,
        settings: &DocumentProcessingSettings,
    ) -> io::Result<()> {
        if settings.fractional_character_widths_for_printing {
            self.write_control_word("fracwidth", None)?;
        }
        if let Some(cleanup) = settings.abstract_numbering_cleanup {
            self.write_control_word(
                "ilfomacatclnup",
                Some(match cleanup {
                    AbstractNumberingCleanupStatus::Reviewed => 0,
                    AbstractNumberingCleanupStatus::Incomplete => 1,
                }),
            )?;
        }
        if let Some(event_mask) = settings.event_mask {
            self.write_control_word("grfdocevents", Some(i32::from(event_mask.bits())))?;
        }
        Ok(())
    }

    /// Write passive document-level drawing-grid properties in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_drawing_grid(&mut self, grid: &DocumentDrawingGrid) -> io::Result<()> {
        if let Some(value) = grid.horizontal_spacing {
            self.write_control_word("dghspace", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.vertical_spacing {
            self.write_control_word("dgvspace", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.horizontal_origin_twips {
            self.write_control_word("dghorigin", Some(i32::from(value)))?;
        }
        if let Some(value) = grid.vertical_origin_twips {
            self.write_control_word("dgvorigin", Some(i32::from(value)))?;
        }
        if let Some(value) = grid.horizontal_line_interval {
            self.write_control_word("dghshow", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.vertical_line_interval {
            self.write_control_word("dgvshow", Some(i32::from(value.get())))?;
        }
        if grid.snap_to_grid {
            self.write_control_word("dgsnap", None)?;
        }
        if grid.follows_margins {
            self.write_control_word("dgmargin", None)?;
        }
        Ok(())
    }

    /// Write passive print-layout settings in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_print_layout_settings(
        &mut self,
        settings: &DocumentPrintLayoutSettings,
    ) -> io::Result<()> {
        settings
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if settings.facing_pages {
            self.write_control_word("facingp", None)?;
        }
        if settings.mirror_margins {
            self.write_control_word("margmirror", None)?;
        }
        if let Some(value) = settings.document_gutter_twips {
            self.write_control_word("gutter", Some(value.cast_signed()))?;
        }
        if settings.parallel_gutter {
            self.write_control_word("gutterprl", None)?;
        }
        if settings.two_logical_pages_per_physical_page {
            self.write_control_word("twoonone", None)?;
        }
        Ok(())
    }

    /// Write passive theme languages in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_theme_languages(
        &mut self,
        languages: &DocumentThemeLanguages,
    ) -> io::Result<()> {
        if let Some(language) = languages.primary {
            self.write_control_word("themelang", Some(language.rtf_value()))?;
        }
        if let Some(language) = languages.east_asian {
            self.write_control_word("themelangfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = languages.complex_script {
            self.write_control_word("themelangcs", Some(language.rtf_value()))?;
        }
        Ok(())
    }

    /// Write passive web-save and custom-XML policies in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_xml_policies(
        &mut self,
        policies: &DocumentXmlPolicies,
    ) -> io::Result<()> {
        for (name, value) in [
            ("relyonvml", policies.rely_on_vml),
            ("validatexml", policies.validate_custom_xml),
            ("showplaceholdtext", policies.show_placeholder_text),
            ("ignoremixedcontent", policies.ignore_mixed_content),
            ("saveinvalidxml", policies.save_invalid_xml),
            ("showxmlerrors", policies.show_xml_errors),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, Some(i32::from(value)))?;
            }
        }
        Ok(())
    }

    /// Write passive embedding policies in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_embedding_policies(
        &mut self,
        policies: &DocumentEmbeddingPolicies,
    ) -> io::Result<()> {
        if let Some(value) = policies.do_not_embed_system_fonts {
            self.write_control_word("donotembedsysfont", Some(i32::from(value)))?;
        }
        if let Some(value) = policies.do_not_embed_linguistic_data {
            self.write_control_word("donotembedlingdata", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write passive revision policies in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_revision_policies(
        &mut self,
        policies: &DocumentRevisionPolicies,
    ) -> io::Result<()> {
        if let Some(value) = policies.track_moves {
            self.write_control_word("trackmoves", Some(i32::from(value)))?;
        }
        if let Some(value) = policies.track_formatting {
            self.write_control_word("trackformatting", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write passive style policies in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_style_policies(
        &mut self,
        policies: &DocumentStylePolicies,
    ) -> io::Result<()> {
        if policies.update_styles_from_template {
            self.write_control_word("linkstyles", None)?;
        }
        if policies.lock_theme {
            self.write_control_word("stylelocktheme", None)?;
        }
        if policies.lock_quick_format_set {
            self.write_control_word("stylelockqfset", None)?;
        }
        if policies.use_normal_style_for_lists {
            self.write_control_word("usenormstyforlist", None)?;
        }
        Ok(())
    }

    /// Write passive legacy style restrictions in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_style_restrictions(
        &mut self,
        restrictions: &DocumentStyleRestrictions,
    ) -> io::Result<()> {
        if restrictions.restrictions_present {
            self.write_control_word("stylelock", None)?;
        }
        if restrictions.enforced {
            self.write_control_word("stylelockenforced", None)?;
        }
        if restrictions.backward_compatibility {
            self.write_control_word("stylelockbackcomp", None)?;
        }
        if restrictions.allow_auto_format_override {
            self.write_control_word("autofmtoverride", None)?;
        }
        Ok(())
    }

    /// Write passive booklet-printing metadata in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_booklet_printing(
        &mut self,
        printing: &DocumentBookletPrinting,
    ) -> io::Result<()> {
        if printing.book_fold {
            self.write_control_word("bookfold", None)?;
        }
        if printing.reverse_book_fold {
            self.write_control_word("bookfoldrev", None)?;
        }
        if let Some(value) = printing.sheets_per_booklet {
            if value > i32::MAX as u32 || value % 4 != 0 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "booklet sheets must be an RTF signed nonnegative multiple of four",
                ));
            }
            self.write_control_word("bookfoldsheets", Some(value.cast_signed()))?;
        }
        Ok(())
    }

    /// Write passive privacy-removal requests in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_privacy_policies(
        &mut self,
        policies: &DocumentPrivacyPolicies,
    ) -> io::Result<()> {
        if policies.remove_personal_information {
            self.write_control_word("rempersonalinfo", None)?;
        }
        if policies.remove_date_time_information {
            self.write_control_word("remdttm", None)?;
        }
        Ok(())
    }

    /// Write passive Word 2003 compatibility flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_word_2003_compatibility(
        &mut self,
        compatibility: &DocumentWord2003Compatibility,
    ) -> io::Result<()> {
        if compatibility.preserve_autofit_table_width_around_shapes {
            self.write_control_word("noafcnsttbl", None)?;
        }
        if compatibility.use_hanging_indent_as_numbering_tab {
            self.write_control_word("noindnmbrts", None)?;
        }
        if compatibility.use_legacy_kinsoku_characters {
            self.write_control_word("felnbrelev", None)?;
        }
        if compatibility.use_legacy_floating_object_indentation {
            self.write_control_word("indrlsweleven", None)?;
        }
        if compatibility.allow_contextual_spacing_in_tables {
            self.write_control_word("nocxsptable", None)?;
        }
        if compatibility.ignore_cell_vertical_alignment_with_floating_objects {
            self.write_control_word("notcvasp", None)?;
        }
        if compatibility.ignore_text_box_vertical_alignment {
            self.write_control_word("notvatxbx", None)?;
        }
        if compatibility.split_page_break_paragraph {
            self.write_control_word("spltpgpar", None)?;
        }
        if compatibility.use_fixed_width_hangul {
            self.write_control_word("hwelev", None)?;
        }
        if compatibility.use_legacy_autofit_width_expansion {
            self.write_control_word("afelev", None)?;
        }
        if compatibility.use_cached_column_balancing {
            self.write_control_word("cachedcolbal", None)?;
        }
        if compatibility.underline_numbering_suffix {
            self.write_control_word("utinl", None)?;
        }
        if compatibility.do_not_split_rows_around_floating_tables {
            self.write_control_word("notbrkcnstfrctbl", None)?;
        }
        if compatibility.use_ansi_kerning_pairs {
            self.write_control_word("krnprsnet", None)?;
        }
        Ok(())
    }

    /// Write passive document compatibility policy in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_compatibility_policy(
        &mut self,
        policy: &DocumentCompatibilityPolicy,
    ) -> io::Result<()> {
        if policy.reset_options_to_defaults {
            self.write_control_word("nocompatoptions", None)?;
        }
        if let Some(feature_throttle) = policy.feature_throttle {
            self.write_control_word("nofeaturethrottle", Some(feature_throttle.rtf_value()))?;
        }
        if policy.force_upgrade {
            self.write_control_word("forceupgrade", None)?;
        }
        Ok(())
    }

    /// Write passive Asian grid compatibility flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_asian_grid_compatibility(
        &mut self,
        compatibility: &DocumentAsianGridCompatibility,
    ) -> io::Result<()> {
        if compatibility.apply_thai_line_breaking_rules {
            self.write_control_word("ApplyBrkRules", None)?;
        }
        if compatibility.snap_text_to_grid_inside_table {
            self.write_control_word("snaptogridincell", None)?;
        }
        if compatibility.allow_hanging_punctuation {
            self.write_control_word("wrppunct", None)?;
        }
        if compatibility.use_asian_line_breaking_rules {
            self.write_control_word("asianbrkrule", None)?;
        }
        if compatibility.compress_punctuation_at_line_start {
            self.write_control_word("toplinepunct", None)?;
        }
        Ok(())
    }

    /// Write passive legacy automatic-layout flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_legacy_layout_compatibility(
        &mut self,
        compatibility: &DocumentLegacyLayoutCompatibility,
    ) -> io::Result<()> {
        if compatibility.do_not_use_word_97_shape_layout {
            self.write_control_word("splytwnine", None)?;
        }
        if compatibility.use_legacy_footnote_layout {
            self.write_control_word("ftnlytwnine", None)?;
        }
        if compatibility.use_html_paragraph_auto_spacing {
            self.write_control_word("htmautsp", None)?;
        }
        if compatibility.preserve_last_tab_alignment {
            self.write_control_word("useltbaln", None)?;
        }
        if compatibility.use_word_95_auto_spacing {
            self.write_control_word("oldas", None)?;
        }
        Ok(())
    }

    /// Write passive table-layout compatibility flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_table_layout_compatibility(
        &mut self,
        compatibility: &DocumentTableLayoutCompatibility,
    ) -> io::Result<()> {
        if compatibility.combine_borders_like_word_5 {
            self.write_control_word("otblrul", None)?;
        }
        if compatibility.do_not_align_rows_independently {
            self.write_control_word("alntblind", None)?;
        }
        if compatibility.do_not_use_raw_table_width {
            self.write_control_word("lytcalctblwd", None)?;
        }
        if compatibility.keep_rows_together {
            self.write_control_word("lyttblrtgr", None)?;
        }
        if compatibility.do_not_adjust_line_height {
            self.write_control_word("nolnhtadjtbl", None)?;
        }
        if compatibility.do_not_break_wrapped_tables_across_pages {
            self.write_control_word("nobrkwrptbl", None)?;
        }
        if compatibility.prevent_autofit_growth_into_margins {
            self.write_control_word("nogrowautofit", None)?;
        }
        if compatibility.use_word_2003_table_style_rules {
            self.write_control_word("newtblstyruls", None)?;
        }
        Ok(())
    }

    /// Write passive East Asian compatibility flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_east_asian_compatibility(
        &mut self,
        compatibility: &DocumentEastAsianCompatibility,
    ) -> io::Result<()> {
        if compatibility.do_not_balance_sbcs_dbcs {
            self.write_control_word("dntblnsbdb", None)?;
        }
        if compatibility.expand_spacing_at_shift_return {
            self.write_control_word("expshrtn", None)?;
        }
        if compatibility.do_not_add_space_for_underline {
            self.write_control_word("nospaceforul", None)?;
        }
        if compatibility.do_not_underline_trailing_spaces {
            self.write_control_word("noultrlspc", None)?;
        }
        if compatibility.do_not_translate_backslash_to_yen {
            self.write_control_word("noxlattoyen", None)?;
        }
        if compatibility.use_legacy_line_breaking_rules {
            self.write_control_word("lnbrkrule", None)?;
        }
        Ok(())
    }

    /// Write passive line-spacing compatibility flags in specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_line_spacing_compatibility(
        &mut self,
        compatibility: &DocumentLineSpacingCompatibility,
    ) -> io::Result<()> {
        if compatibility.suppress_extra_spacing_for_raised_lowered_text {
            self.write_control_word("noextrasprl", None)?;
        }
        if compatibility.suppress_extra_spacing_at_top_of_page {
            self.write_control_word("sprstsp", None)?;
        }
        if compatibility.suppress_space_before_after_hard_break {
            self.write_control_word("sprsspbf", None)?;
        }
        if compatibility.suppress_wordperfect_extra_line_spacing {
            self.write_control_word("sprslnsp", None)?;
        }
        if compatibility.suppress_extra_spacing_at_bottom_of_page {
            self.write_control_word("sprsbsp", None)?;
        }
        Ok(())
    }

    /// Write passive file and template flags in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_file_settings(
        &mut self,
        settings: &DocumentFileSettings,
    ) -> io::Result<()> {
        if settings.automatic_backup {
            self.write_control_word("makebackup", None)?;
        }
        if settings.default_save_format_rtf {
            self.write_control_word("defformat", None)?;
        }
        if settings.template_or_stationery {
            self.write_control_word("doctemp", None)?;
        }
        Ok(())
    }

    /// Write passive view metadata in stable `viewkind`, `viewscale`, `viewzk` order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_view(&mut self, view: &DocumentView) -> io::Result<()> {
        view.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(kind) = view.kind {
            self.write_control_word("viewkind", Some(kind.rtf_value()))?;
        }
        if let Some(scale) = view.scale_percent {
            self.write_control_word("viewscale", Some(i32::from(scale)))?;
        }
        if let Some(kind) = view.zoom_kind {
            self.write_control_word("viewzk", Some(kind.rtf_value()))?;
        }
        if let Some(value) = view.background_shapes {
            self.write_control_word("viewbksp", Some(i32::from(value)))?;
        }
        if view.hide_page_boundaries {
            self.write_control_word("viewnobound", None)?;
        }
        Ok(())
    }

    /// Write passive review-display flags in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_review_display(&mut self, display: &DocumentReviewDisplay) -> io::Result<()> {
        if display.hide_markup {
            self.write_control_word("donotshowmarkup", None)?;
        }
        if display.hide_comments {
            self.write_control_word("donotshowcomments", None)?;
        }
        if display.hide_insertions_and_deletions {
            self.write_control_word("donotshowinsdel", None)?;
        }
        Ok(())
    }

    /// Write an inert document-window caption as the canonical starred destination.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_window_caption(
        &mut self,
        caption: Option<&DocumentWindowCaption<'_>>,
    ) -> io::Result<()> {
        let Some(caption) = caption else {
            return Ok(());
        };
        caption
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\windowcaption ")?;
        self.write_destination_text(caption.text.as_ref())?;
        self.write_str("}")
    }

    /// Write the inert custom kinsoku character sets and their language.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_kinsoku(&mut self, kinsoku: &DocumentKinsoku<'_>) -> io::Result<()> {
        kinsoku
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(following) = &kinsoku.following {
            self.write_str("{\\*\\fchars ")?;
            self.write_destination_text(following.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(leading) = &kinsoku.leading {
            self.write_str("{\\*\\lchars ")?;
            self.write_destination_text(leading.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(language) = kinsoku.language {
            self.write_control_word("ksulang", Some(language.cast_signed()))?;
        }
        Ok(())
    }

    /// Write an inert custom XSL transform location as its required starred destination.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_xsl_transform(
        &mut self,
        transform: Option<&DocumentXslTransform<'_>>,
    ) -> io::Result<()> {
        let Some(transform) = transform else {
            return Ok(());
        };
        transform
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\xform ")?;
        self.write_destination_text(transform.location.as_ref())?;
        self.write_str("}")
    }

    /// Write passive requested transform usage without applying the transform.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_xsl_transform_usage(
        &mut self,
        usage: DocumentXslTransformUsage,
    ) -> io::Result<()> {
        if usage.is_requested() {
            self.write_control_word("usexform", None)?;
        }
        Ok(())
    }

    /// Write passive style-list filter suggestions as exactly four hexadecimal digits.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_style_list_filter(
        &mut self,
        filter: Option<DocumentStyleListFilter>,
    ) -> io::Result<()> {
        let Some(filter) = filter else {
            return Ok(());
        };
        filter
            .validate_for_write()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\wgrffmtfilter ")?;
        self.write_str(&format!("{:04X}", filter.bits()))?;
        self.write_str("}")
    }

    /// Write a passive style-list sorting suggestion with an explicit value.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_style_sort_method(
        &mut self,
        method: Option<DocumentStyleSortMethod>,
    ) -> io::Result<()> {
        if let Some(method) = method {
            self.write_control_word("stylesortmethod", Some(method.rtf_value()))?;
        }
        Ok(())
    }

    /// Write opaque reservation metadata without authenticating or decrypting it.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_write_reservations(
        &mut self,
        reservations: &DocumentWriteReservations<'_>,
    ) -> io::Result<()> {
        reservations
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(hash) = &reservations.hash {
            self.write_str("{\\*\\writereservhash ")?;
            for byte in hash.data.iter() {
                self.write_str(&format!("{byte:02X}"))?;
            }
            self.write_str("}")?;
        }
        if let Some(legacy) = &reservations.legacy {
            self.write_str("{\\*\\writereservation ")?;
            self.write_destination_text(legacy.data.as_ref())?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write passive save preferences in stable specification order.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_document_save_preferences(
        &mut self,
        preferences: &DocumentSavePreferences,
    ) -> io::Result<()> {
        if preferences.read_only == DocumentReadOnlyRecommendation::Recommended {
            self.write_control_word("readonlyrecommended", None)?;
        }
        if preferences.thumbnail == DocumentThumbnailPreference::RequiredIfSupported {
            self.write_control_word("saveprevpict", None)?;
        }
        Ok(())
    }
}
