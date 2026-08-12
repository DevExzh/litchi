//! RTF body-content output.

#![allow(
    clippy::shadow_reuse,
    reason = "serialization helpers deliberately rebind a working value as the output is assembled"
)]
use super::super::{
    Alignment, AnimatedTextEffect, AssociatedCharacterBaseline, AssociatedUnderlineStyle, Border,
    BorderStyle, Borders, CharacterBaseline, CharacterExpansion, CharacterGrid, CharacterType,
    EmphasisMark, FormField, FormFieldType, Formatting, GeneratedListMarker,
    GeneratedListMarkerKind, IndexPageReference, LegacyParagraphNumberingAlignment,
    LegacyParagraphNumberingBidi, LegacyParagraphNumberingFormat, LegacyParagraphNumberingLevel,
    LegacyParagraphNumberingUnderline, NavigationEntry, Paragraph, ParagraphFontAlignment,
    ParagraphWrapping, RevisionMetadata, RtfWriter, Shading, ShadingPattern, StyleBlock,
    TabAlignment, TabLeader, TabStop, TextDirection, UnderlineStyle, Write, io,
};

fn explicit_shading_pattern_word(
    pattern: ShadingPattern,
    character: bool,
) -> io::Result<&'static str> {
    let word = match (character, pattern) {
        (false, ShadingPattern::Horizontal) => "bghoriz",
        (false, ShadingPattern::Vertical) => "bgvert",
        (false, ShadingPattern::ForwardDiagonal) => "bgfdiag",
        (false, ShadingPattern::BackwardDiagonal) => "bgbdiag",
        (false, ShadingPattern::Cross) => "bgcross",
        (false, ShadingPattern::DiagonalCross) => "bgdcross",
        (false, ShadingPattern::DarkHorizontal) => "bgdkhoriz",
        (false, ShadingPattern::DarkVertical) => "bgdkvert",
        (false, ShadingPattern::DarkForwardDiagonal) => "bgdkfdiag",
        (false, ShadingPattern::DarkBackwardDiagonal) => "bgdkbdiag",
        (false, ShadingPattern::DarkCross) => "bgdkcross",
        (false, ShadingPattern::DarkDiagonalCross) => "bgdkdcross",
        (true, ShadingPattern::Horizontal) => "chbghoriz",
        (true, ShadingPattern::Vertical) => "chbgvert",
        (true, ShadingPattern::ForwardDiagonal) => "chbgfdiag",
        (true, ShadingPattern::BackwardDiagonal) => "chbgbdiag",
        (true, ShadingPattern::Cross) => "chbgcross",
        (true, ShadingPattern::DiagonalCross) => "chbgdcross",
        (true, ShadingPattern::DarkHorizontal) => "chbgdkhoriz",
        (true, ShadingPattern::DarkVertical) => "chbgdkvert",
        (true, ShadingPattern::DarkForwardDiagonal) => "chbgdkfdiag",
        (true, ShadingPattern::DarkBackwardDiagonal) => "chbgdkbdiag",
        (true, ShadingPattern::DarkCross) => "chbgdkcross",
        (true, ShadingPattern::DarkDiagonalCross) => "chbgdkdcross",
        _ => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "numeric RTF shading pattern cannot use an explicit pattern control",
            ));
        },
    };
    Ok(word)
}

impl<W: Write> RtfWriter<W> {
    /// Write one inert generated list-marker destination.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_generated_list_marker(
        &mut self,
        marker: &GeneratedListMarker<'_>,
    ) -> io::Result<()> {
        marker
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(
            match marker.kind {
                GeneratedListMarkerKind::Modern => "listtext",
                GeneratedListMarkerKind::Legacy => "pntext",
            },
            None,
        )?;
        self.write_str(" ")?;
        let mut segments = marker.text.split('\t').peekable();
        while let Some(segment) = segments.next() {
            self.write_destination_text(segment)?;
            if segments.peek().is_some() {
                self.write_control_word("tab", None)?;
            }
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_form_field_start(
        &mut self,
        field: &FormField<'_>,
    ) -> io::Result<()> {
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\field{\\*\\fldinst ")?;
        self.write_str(match field.field_type {
            FormFieldType::Text => "FORMTEXT",
            FormFieldType::CheckBox => "FORMCHECKBOX",
            FormFieldType::DropDown => "FORMDROPDOWN",
        })?;
        if !field.data.is_empty() {
            self.write_str("{\\*\\datafield ")?;
            for byte in field.data.iter() {
                write!(self.writer, "{byte:02x}")?;
            }
            self.write_str("}")?;
        }
        self.write_str("{\\*\\formfield{")?;
        self.write_control_word("fftype", Some(field.field_type.to_rtf()))?;
        if let Some(value) = field.text_type {
            self.write_control_word("fftypetxt", Some(value.to_rtf()))?;
        }
        if let Some(value) = field.max_length {
            self.write_control_word("ffmaxlen", Some(i32::from(value)))?;
        }
        if let Some(value) = field.half_point_size {
            self.write_control_word("ffhps", Some(value))?;
        }
        if field.protected {
            self.write_control_word("ffprot", None)?;
        }
        if field.calculate_on_exit {
            self.write_control_word("ffrecalc", None)?;
        }
        if field.size_automatically {
            self.write_control_word("ffsize", None)?;
        }
        if field.own_help {
            self.write_control_word("ffownhelp", None)?;
        }
        if field.own_status {
            self.write_control_word("ffownstat", None)?;
        }
        if field.has_list_box {
            self.write_control_word("ffhaslistbox", None)?;
        }
        if let Some(value) = field.default_result {
            self.write_control_word("ffdefres", Some(value))?;
        }
        if let Some(value) = field.result {
            self.write_control_word("ffres", Some(value))?;
        }
        self.write_form_field_value("ffname", field.name.as_deref())?;
        self.write_form_field_value("ffformat", field.format.as_deref())?;
        self.write_form_field_value("ffdeftext", field.default_text.as_deref())?;
        self.write_form_field_value("ffhelptext", field.help_text.as_deref())?;
        self.write_form_field_value("ffstattext", field.status_text.as_deref())?;
        self.write_form_field_value("ffentrymcr", field.entry_macro.as_deref())?;
        self.write_form_field_value("ffexitmcr", field.exit_macro.as_deref())?;
        for entry in &field.list_entries {
            self.write_form_field_value("ffl", Some(entry.as_ref()))?;
        }
        self.write_str("}}}")?; // formfield and fldinst
        self.write_str("{\\fldrslt ")
    }

    pub(in super::super) fn write_form_field_value(
        &mut self,
        control: &str,
        value: Option<&str>,
    ) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    /// Write an inert source mark. Marks are canonicalized as hidden; any
    /// originally visible entry text remains in the ordinary body stream.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_navigation_entry(&mut self, entry: &NavigationEntry<'_>) -> io::Result<()> {
        entry
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        self.write_str("{")?;
        match entry {
            NavigationEntry::Index(entry) => {
                self.write_control_word("xe", None)?;
                self.write_control_word("v", None)?;
                if let Some(index_id) = entry.index_id {
                    self.write_control_word("xef", Some(i32::from(index_id)))?;
                }
                if entry.bold_page_number {
                    self.write_control_word("bxe", None)?;
                }
                if entry.italic_page_number {
                    self.write_control_word("ixe", None)?;
                }
                self.write_str(" ")?;
                self.write_destination_text(entry.text.as_ref())?;
                match &entry.page_reference {
                    IndexPageReference::CurrentPage => {},
                    IndexPageReference::ReplacementText(value) => {
                        self.write_str("{")?;
                        self.write_control_word("txe", None)?;
                        self.write_str(" ")?;
                        self.write_destination_text(value.as_ref())?;
                        self.write_str("}")?;
                    },
                    IndexPageReference::BookmarkRange(value) => {
                        self.write_str("{")?;
                        self.write_control_word("rxe", None)?;
                        self.write_str(" ")?;
                        self.write_destination_text(value.as_ref())?;
                        self.write_str("}")?;
                    },
                }
                if let Some(yomi) = &entry.yomi {
                    self.write_str("{")?;
                    self.write_control_word("yxe", None)?;
                    self.write_str("{\\*")?;
                    self.write_control_word("pxe", None)?;
                    self.write_str(" ")?;
                    self.write_destination_text(yomi.as_ref())?;
                    self.write_str("}}")?;
                }
            },
            NavigationEntry::TableOfContents(entry) => {
                self.write_control_word(
                    if entry.suppress_page_number {
                        "tcn"
                    } else {
                        "tc"
                    },
                    None,
                )?;
                self.write_control_word("v", None)?;
                self.write_control_word("tcf", Some(i32::from(entry.table_id)))?;
                self.write_control_word("tcl", Some(i32::from(entry.level)))?;
                self.write_str(" ")?;
                self.write_destination_text(entry.text.as_ref())?;
            },
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_style_block_fragment(
        &mut self,
        block: &StyleBlock<'_>,
        start: usize,
        end: usize,
        body_position: usize,
        boundaries: &[crate::story::Boundary],
        boundary: &mut usize,
    ) -> io::Result<()> {
        let text = block.text.get(start..end).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark boundary splits a UTF-8 character",
            )
        })?;
        let fragment = StyleBlock::new(
            std::borrow::Cow::Borrowed(text),
            block.formatting,
            block.paragraph,
        );
        let fragment_position = body_position.checked_add(start).ok_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidInput, "RTF body text size overflow")
        })?;
        self.write_style_block_with_boundaries(&fragment, fragment_position, boundaries, boundary)
    }

    pub(in super::super) fn write_style_block_with_boundaries(
        &mut self,
        block: &StyleBlock<'_>,
        body_position: usize,
        boundaries: &[crate::story::Boundary],
        boundary: &mut usize,
    ) -> io::Result<()> {
        self.write_str("{")?;

        // Write character formatting
        self.write_formatting(&block.formatting)?;

        // Write paragraph properties
        self.write_paragraph_properties(&block.paragraph)?;

        // Delimit the final control word from body text that starts with a letter.
        self.write_str(" ")?;

        self.write_body_text(block.text.as_ref(), body_position, boundaries, boundary)?;

        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn write_body_text(
        &mut self,
        text: &str,
        body_position: usize,
        boundaries: &[crate::story::Boundary],
        boundary: &mut usize,
    ) -> io::Result<()> {
        let mut fragment_start = 0usize;
        for (offset, character) in text.char_indices() {
            if character != '\n' {
                continue;
            }
            let fragment = text.get(fragment_start..offset).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF body text boundary splits a UTF-8 character",
                )
            })?;
            self.write_text(fragment)?;

            let position = body_position.checked_add(offset).ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "RTF body text size overflow")
            })?;
            match boundaries.get(*boundary).copied() {
                Some(value) if value.position == position => {
                    self.write_str(match value.kind {
                        crate::text::Break::Paragraph => "\\par ",
                        crate::text::Break::Line => "\\line ",
                    })?;
                    *boundary = (*boundary).checked_add(1).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body text boundary index overflow",
                        )
                    })?;
                },
                Some(value) if value.position < position => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF body text boundary precedes its line-feed byte",
                    ));
                },
                _ => self.write_str("\\'0a")?,
            }
            fragment_start = offset.saturating_add(character.len_utf8());
        }
        let remainder = text.get(fragment_start..).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body text boundary splits a UTF-8 character",
            )
        })?;
        self.write_text(remainder)
    }

    /// Write character formatting
    pub(crate) fn write_formatting(&mut self, fmt: &Formatting) -> io::Result<()> {
        if let Some(character_style) = fmt.character_style {
            self.write_control_word("cs", Some(i32::from(character_style)))?;
        }
        if let Some(insert_rsid) = fmt.insert_rsid {
            self.write_control_word("insrsid", Some(insert_rsid.cast_signed()))?;
        }
        if let Some(delete_rsid) = fmt.delete_rsid {
            self.write_control_word("delrsid", Some(delete_rsid.cast_signed()))?;
        }
        if let Some(char_style_rsid) = fmt.char_style_rsid {
            self.write_control_word("charrsid", Some(char_style_rsid.cast_signed()))?;
        }
        if let Some(direction) = fmt.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrch",
                    TextDirection::RightToLeft => "rtlch",
                },
                None,
            )?;
        }

        if let Some(complex_script) = fmt.complex_script {
            self.write_control_word("fcs", Some(i32::from(complex_script)))?;
        }
        if let Some(character_type) = fmt.character_type {
            self.write_control_word(
                match character_type {
                    CharacterType::LowAnsi => "loch",
                    CharacterType::HighAnsi => "hich",
                    CharacterType::DoubleByte => "dbch",
                },
                None,
            )?;
        }
        if let Some(character_grid) = fmt.character_grid {
            self.write_control_word(
                "cgrid",
                match character_grid {
                    CharacterGrid::Parameterless => None,
                    CharacterGrid::Value(value) => Some(i32::from(value)),
                },
            )?;
        }
        if fmt.animated_text != AnimatedTextEffect::None {
            self.write_control_word("animtext", Some(fmt.animated_text.rtf_value()))?;
        }
        if let Some(value) = fmt.fit_text.rtf_value() {
            self.write_control_word("fittext", Some(value))?;
        }
        if fmt.emphasis_mark != EmphasisMark::None {
            self.write_control_word(fmt.emphasis_mark.control_word(), None)?;
        }

        if let Some(language) = fmt.language {
            self.write_control_word("lang", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.east_asian_language {
            self.write_control_word("langfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.language_no_proof {
            self.write_control_word("langnp", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.east_asian_language_no_proof {
            self.write_control_word("langfenp", Some(language.rtf_value()))?;
        }
        if fmt.no_proof {
            self.write_control_word("noproof", None)?;
        }

        fmt.associated
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(bold) = fmt.associated.bold {
            self.write_control_word("ab", Some(i32::from(bold)))?;
        }
        if let Some(all_caps) = fmt.associated.all_caps {
            self.write_control_word("acaps", Some(i32::from(all_caps)))?;
        }
        if let Some(color_ref) = fmt.associated.color_ref {
            self.write_control_word("acf", Some(i32::from(color_ref)))?;
        }
        if let Some(AssociatedCharacterBaseline::LoweredHalfPoints(value)) = fmt.associated.baseline
        {
            self.write_control_word("adn", Some(i32::from(value)))?;
        }
        if let Some(expansion) = fmt.associated.expansion_quarter_points {
            self.write_control_word("aexpnd", Some(i32::from(expansion)))?;
        }
        if let Some(font_ref) = fmt.associated.font_ref {
            self.write_control_word("af", Some(i32::from(font_ref)))?;
        }
        if let Some(font_size) = fmt.associated.font_size {
            self.write_control_word("afs", Some(i32::from(font_size.get())))?;
        }
        if let Some(italic) = fmt.associated.italic {
            self.write_control_word("ai", Some(i32::from(italic)))?;
        }
        if let Some(language) = fmt.associated.language {
            self.write_control_word("alang", Some(language.rtf_value()))?;
        }
        if let Some(outline) = fmt.associated.outline {
            self.write_control_word("aoutl", Some(i32::from(outline)))?;
        }
        if let Some(small_caps) = fmt.associated.small_caps {
            self.write_control_word("ascaps", Some(i32::from(small_caps)))?;
        }
        if let Some(shadow) = fmt.associated.shadow {
            self.write_control_word("ashad", Some(i32::from(shadow)))?;
        }
        if let Some(strike) = fmt.associated.strike {
            self.write_control_word("astrike", Some(i32::from(strike)))?;
        }
        if let Some(underline) = fmt.associated.underline {
            self.write_control_word(
                match underline {
                    AssociatedUnderlineStyle::None => "aulnone",
                    AssociatedUnderlineStyle::Single => "aul",
                    AssociatedUnderlineStyle::Dotted => "auld",
                    AssociatedUnderlineStyle::Double => "auldb",
                    AssociatedUnderlineStyle::Words => "aulw",
                },
                None,
            )?;
        }
        if let Some(AssociatedCharacterBaseline::RaisedHalfPoints(value)) = fmt.associated.baseline
        {
            self.write_control_word("aup", Some(i32::from(value)))?;
        }

        // Font
        if fmt.font_ref != 0 {
            self.write_control_word("f", Some(i32::from(fmt.font_ref)))?;
        }

        // Font size
        self.write_control_word("fs", Some(i32::from(fmt.font_size.get())))?;

        // Color
        if fmt.color_ref != 0 {
            self.write_control_word("cf", Some(i32::from(fmt.color_ref)))?;
        }

        // Exact character background color, independent of highlighting.
        if let Some(background_color) = fmt.background_color {
            self.write_control_word("cb", Some(i32::from(background_color)))?;
        }

        // Highlight
        if let Some(highlight) = fmt.highlight_color {
            self.write_control_word("highlight", Some(i32::from(highlight)))?;
        }

        if let Some(border) = fmt.character_border {
            border
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_control_word("chbrdr", None)?;
            self.write_control_word(border.style.control_word(), None)?;
            self.write_control_word("brdrw", Some(i32::from(border.width)))?;
            self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
            self.write_control_word("brsp", Some(i32::from(border.space)))?;
            if border.shadow {
                self.write_control_word("brdrsh", None)?;
            }
            if border.frame {
                self.write_control_word("brdrframe", None)?;
            }
        }

        if let Some(shading) = fmt.character_shading {
            shading
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if let Some(amount) = shading.amount {
                self.write_control_word("chshdng", Some(i32::from(amount)))?;
            } else if let Some(pattern) = shading.pattern {
                self.write_control_word(explicit_shading_pattern_word(pattern, true)?, None)?;
            }
            if let Some(color) = shading.foreground_color {
                self.write_control_word("chcfpat", Some(i32::from(color)))?;
            }
            if let Some(color) = shading.background_color {
                self.write_control_word("chcbpat", Some(i32::from(color)))?;
            }
        }

        // Bold
        if fmt.bold {
            self.write_control_word("b", None)?;
        }

        // Italic
        if fmt.italic {
            self.write_control_word("i", None)?;
        }

        // Underline
        match fmt.underline {
            UnderlineStyle::None => {},
            UnderlineStyle::Single => self.write_control_word("ul", None)?,
            UnderlineStyle::Double => self.write_control_word("uldb", None)?,
            UnderlineStyle::Dotted => self.write_control_word("uld", None)?,
            UnderlineStyle::Dashed => self.write_control_word("uldash", None)?,
            UnderlineStyle::DashDot => self.write_control_word("uldashd", None)?,
            UnderlineStyle::DashDotDot => self.write_control_word("uldashdd", None)?,
            UnderlineStyle::Words => self.write_control_word("ulw", None)?,
            UnderlineStyle::Thick => self.write_control_word("ulth", None)?,
            UnderlineStyle::Wave => self.write_control_word("ulwave", None)?,
            UnderlineStyle::Hairline => self.write_control_word("ulhair", None)?,
            UnderlineStyle::ThickDotted => self.write_control_word("ulthd", None)?,
            UnderlineStyle::ThickDashed => self.write_control_word("ulthdash", None)?,
            UnderlineStyle::ThickDashDot => self.write_control_word("ulthdashd", None)?,
            UnderlineStyle::ThickDashDotDot => self.write_control_word("ulthdashdd", None)?,
            UnderlineStyle::ThickLongDash => self.write_control_word("ulthldash", None)?,
            UnderlineStyle::LongDash => self.write_control_word("ulldash", None)?,
            UnderlineStyle::HeavyWave => self.write_control_word("ulhwave", None)?,
            UnderlineStyle::DoubleWave => self.write_control_word("ululdbwave", None)?,
        }
        if let Some(underline_color) = fmt.underline_color {
            self.write_control_word("ulc", Some(i32::from(underline_color)))?;
        }

        // Strike
        if fmt.strike {
            self.write_control_word("strike", None)?;
        }

        // Double strike
        if fmt.double_strike {
            self.write_control_word("striked", None)?;
        }

        fmt.character_positioning
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        match fmt.character_positioning.baseline {
            CharacterBaseline::Normal if fmt.superscript => {
                self.write_control_word("super", None)?;
            },
            CharacterBaseline::Normal if fmt.subscript => self.write_control_word("sub", None)?,
            CharacterBaseline::Normal => {},
            CharacterBaseline::Superscript => self.write_control_word("super", None)?,
            CharacterBaseline::Subscript => self.write_control_word("sub", None)?,
            CharacterBaseline::RaisedHalfPoints(value) => {
                self.write_control_word("up", Some(i32::from(value)))?;
            },
            CharacterBaseline::LoweredHalfPoints(value) => {
                self.write_control_word("dn", Some(i32::from(value)))?;
            },
        }

        // Small caps
        if fmt.smallcaps {
            self.write_control_word("scaps", None)?;
        }

        // All caps
        if fmt.all_caps {
            self.write_control_word("caps", None)?;
        }

        // Hidden
        if fmt.hidden {
            self.write_control_word("v", None)?;
        }

        // Outline
        if fmt.outline {
            self.write_control_word("outl", None)?;
        }

        // Shadow
        if fmt.shadow {
            self.write_control_word("shad", None)?;
        }

        // Emboss
        if fmt.emboss {
            self.write_control_word("embo", None)?;
        }

        // Imprint
        if fmt.imprint {
            self.write_control_word("impr", None)?;
        }

        match fmt.character_positioning.expansion {
            CharacterExpansion::None if fmt.char_spacing != 0 => {
                self.write_control_word("expnd", Some(fmt.char_spacing))?;
            },
            CharacterExpansion::None => {},
            CharacterExpansion::QuarterPoints(value) => {
                self.write_control_word("expnd", Some(i32::from(value)))?;
            },
            CharacterExpansion::Twips(value) => {
                self.write_control_word("expndtw", Some(i32::from(value)))?;
            },
        }
        let scale = if fmt.character_positioning.horizontal_scale_percent == 100 {
            fmt.char_scale
        } else {
            i32::from(fmt.character_positioning.horizontal_scale_percent)
        };
        if scale != 100 {
            self.write_control_word("charscalex", Some(scale))?;
        }
        let kerning = if fmt.character_positioning.kerning_half_points != 0 {
            i32::from(fmt.character_positioning.kerning_half_points)
        } else {
            fmt.kerning
        };
        if kerning != 0 {
            self.write_control_word("kerning", Some(kerning))?;
        }

        Ok(())
    }

    pub(in super::super) fn write_legacy_paragraph_numbering(
        &mut self,
        index: Option<u32>,
    ) -> io::Result<()> {
        let Some(index) = index else {
            return Ok(());
        };
        let record = self
            .legacy_paragraph_numbering
            .get(index as usize)
            .cloned()
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF paragraph references a missing legacy pn record",
                )
            })?;
        record
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word("pn", None)?;
        match record.level {
            LegacyParagraphNumberingLevel::Explicit(value) => {
                self.write_control_word("pnlvl", Some(i32::from(value)))?;
            },
            LegacyParagraphNumberingLevel::Bullet => self.write_control_word("pnlvlblt", None)?,
            LegacyParagraphNumberingLevel::Body => self.write_control_word("pnlvlbody", None)?,
            LegacyParagraphNumberingLevel::Continue => {
                self.write_control_word("pnlvlcont", None)?;
            },
        }
        if let Some(format) = record.format {
            self.write_control_word(
                match format {
                    LegacyParagraphNumberingFormat::Aiueo => "pnaiu",
                    LegacyParagraphNumberingFormat::AiueoDbChar => "pnaiud",
                    LegacyParagraphNumberingFormat::AiueoExtended => "pnaiueo",
                    LegacyParagraphNumberingFormat::AiueoExtendedDbChar => "pnaiueod",
                    LegacyParagraphNumberingFormat::Chosung => "pnchosung",
                    LegacyParagraphNumberingFormat::CardinalText => "pncard",
                    LegacyParagraphNumberingFormat::Decimal => "pndec",
                    LegacyParagraphNumberingFormat::DecimalWithPeriod => "pndecd",
                    LegacyParagraphNumberingFormat::UpperRoman => "pnucrm",
                    LegacyParagraphNumberingFormat::LowerRoman => "pnlcrm",
                    LegacyParagraphNumberingFormat::UpperLetter => "pnucltr",
                    LegacyParagraphNumberingFormat::LowerLetter => "pnlcltr",
                    LegacyParagraphNumberingFormat::Ordinal => "pnord",
                    LegacyParagraphNumberingFormat::OrdinalText => "pnordt",
                    LegacyParagraphNumberingFormat::ChineseCounting => "pncnum",
                    LegacyParagraphNumberingFormat::ChineseCountingDbChar => "pndbnum",
                    LegacyParagraphNumberingFormat::ChineseCountingKorean => "pndbnumd",
                    LegacyParagraphNumberingFormat::ChineseCountingLegal => "pndbnumk",
                    LegacyParagraphNumberingFormat::ChineseCountingThousand => "pndbnuml",
                    LegacyParagraphNumberingFormat::ChineseCountingTraditional => "pndbnumt",
                    LegacyParagraphNumberingFormat::Ganada => "pnganada",
                    LegacyParagraphNumberingFormat::GbCounting => "pngbnum",
                    LegacyParagraphNumberingFormat::GbCountingDbChar => "pngbnumd",
                    LegacyParagraphNumberingFormat::GbCountingKorean => "pngbnumk",
                    LegacyParagraphNumberingFormat::GbCountingLegal => "pngbnuml",
                    LegacyParagraphNumberingFormat::GbLip => "pngblip",
                    LegacyParagraphNumberingFormat::Iroha => "pniroha",
                    LegacyParagraphNumberingFormat::IrohaDbChar => "pnirohad",
                    LegacyParagraphNumberingFormat::Zodiac => "pnzodiac",
                    LegacyParagraphNumberingFormat::ZodiacDbChar => "pnzodiacd",
                    LegacyParagraphNumberingFormat::ZodiacLegal => "pnzodiacl",
                },
                None,
            )?;
        }
        if let Some(value) = record.alignment {
            self.write_control_word(
                match value {
                    LegacyParagraphNumberingAlignment::Left => "pnql",
                    LegacyParagraphNumberingAlignment::Center => "pnqc",
                    LegacyParagraphNumberingAlignment::Right => "pnqr",
                },
                None,
            )?;
        }
        for (enabled, name) in [
            (record.across, "pnacross"),
            (record.number_once, "pnnumonce"),
            (record.previous, "pnprev"),
            (record.restart, "pnrestart"),
            (record.hanging, "pnhang"),
        ] {
            if enabled {
                self.write_control_word(name, None)?;
            }
        }
        if let Some(value) = record.bidi {
            self.write_control_word(
                match value {
                    LegacyParagraphNumberingBidi::A => "pnbidia",
                    LegacyParagraphNumberingBidi::B => "pnbidib",
                },
                None,
            )?;
        }
        if let Some(value) = record.start_at {
            self.write_control_word("pnstart", Some(value))?;
        }
        if let Some(value) = record.indent {
            self.write_control_word("pnindent", Some(value))?;
        }
        if let Some(value) = record.space {
            self.write_control_word("pnsp", Some(value))?;
        }
        if let Some(value) = record.font_ref {
            self.write_control_word("pnf", Some(i32::from(value)))?;
        }
        if let Some(value) = record.font_size {
            self.write_control_word("pnfs", Some(i32::from(value)))?;
        }
        if let Some(value) = record.color_ref {
            self.write_control_word("pncf", Some(i32::from(value)))?;
        }
        for (value, name) in [
            (record.bold, "pnb"),
            (record.italic, "pni"),
            (record.caps, "pncaps"),
            (record.small_caps, "pnscaps"),
            (record.strike, "pnstrike"),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, (!value).then_some(0))?;
            }
        }
        if let Some(value) = record.underline {
            self.write_control_word(
                match value {
                    LegacyParagraphNumberingUnderline::None => "pnulnone",
                    LegacyParagraphNumberingUnderline::Single => "pnul",
                    LegacyParagraphNumberingUnderline::Dotted => "pnuld",
                    LegacyParagraphNumberingUnderline::Dashed => "pnuldash",
                    LegacyParagraphNumberingUnderline::DashDot => "pnuldashd",
                    LegacyParagraphNumberingUnderline::DashDotDot => "pnuldashdd",
                    LegacyParagraphNumberingUnderline::Double => "pnuldb",
                    LegacyParagraphNumberingUnderline::Hairline => "pnulhair",
                    LegacyParagraphNumberingUnderline::Thick => "pnulth",
                    LegacyParagraphNumberingUnderline::Words => "pnulw",
                    LegacyParagraphNumberingUnderline::Wave => "pnulwave",
                },
                None,
            )?;
        }
        let revision = &record.revision;
        if let Some(value) = revision.author {
            self.write_control_word("pnrauth", Some(i32::from(value)))?;
        }
        if let Some(value) = revision.date {
            self.write_control_word("pnrdate", Some(value))?;
        }
        if let Some(value) = revision.number_format {
            self.write_control_word("pnrnfc", Some(value))?;
        }
        if revision.no_tracking {
            self.write_control_word("pnrnot", None)?;
        }
        if let Some(value) = revision.paragraph_number {
            self.write_control_word("pnrpnbr", Some(value))?;
        }
        if let Some(value) = revision.rgb {
            self.write_control_word("pnrrgb", Some(value.cast_signed()))?;
        }
        if let Some(value) = revision.start {
            self.write_control_word("pnrstart", Some(value))?;
        }
        if let Some(value) = revision.stop {
            self.write_control_word("pnrstop", Some(value))?;
        }
        if let Some(value) = revision.text_start {
            self.write_control_word("pnrxst", Some(value))?;
        }
        if let Some(value) = record.text_before {
            self.write_str("{")?;
            self.write_control_word("pntxtb", None)?;
            self.write_str(" ")?;
            self.write_text(value.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(value) = record.text_after {
            self.write_str("{")?;
            self.write_control_word("pntxta", None)?;
            self.write_str(" ")?;
            self.write_text(value.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    /// Write author/date metadata for a structural revision marker, after
    /// validating the author index.
    pub(in super::super) fn write_revision_metadata(
        &mut self,
        author_control: &'static str,
        date_control: &'static str,
        metadata: RevisionMetadata,
    ) -> io::Result<()> {
        metadata
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(author) = metadata.author {
            self.write_control_word(author_control, Some(author))?;
        }
        if let Some(date) = metadata.date {
            self.write_control_word(date_control, Some(date))?;
        }
        Ok(())
    }

    /// Write paragraph properties
    pub(crate) fn write_paragraph_properties(&mut self, para: &Paragraph) -> io::Result<()> {
        if let Some(paragraph_style) = para.paragraph_style {
            self.write_control_word("s", Some(i32::from(paragraph_style)))?;
        }
        if let Some(paragraph_rsid) = para.paragraph_rsid {
            self.write_control_word("pararsid", Some(paragraph_rsid.cast_signed()))?;
        }
        if let Some(outline_level) = para.outline_level {
            self.write_control_word("outlinelevel", Some(i32::from(outline_level)))?;
        }
        self.write_revision_metadata("prauth", "prdate", para.revision)?;
        self.write_legacy_paragraph_numbering(para.legacy_numbering)?;
        if let Some(direction) = para.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrpar",
                    TextDirection::RightToLeft => "rtlpar",
                },
                None,
            )?;
        }

        // Alignment
        match para.alignment {
            Alignment::Left => self.write_control_word("ql", None)?,
            Alignment::Right => self.write_control_word("qr", None)?,
            Alignment::Center => self.write_control_word("qc", None)?,
            Alignment::Justify => self.write_control_word("qj", None)?,
        }

        // Spacing
        if para.spacing.before != 0 {
            self.write_control_word("sb", Some(para.spacing.before))?;
        }
        if para.spacing.after != 0 {
            self.write_control_word("sa", Some(para.spacing.after))?;
        }
        if let Some(value) = para.spacing_policy.list_before {
            self.write_control_word(
                "lisb",
                Some(i32::try_from(value).map_err(|_err| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF lisb exceeds i32")
                })?),
            )?;
        }
        if let Some(value) = para.spacing_policy.list_after {
            self.write_control_word(
                "lisa",
                Some(i32::try_from(value).map_err(|_err| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF lisa exceeds i32")
                })?),
            )?;
        }
        if para.spacing_policy.automatic_before {
            self.write_control_word("sbauto", Some(1))?;
        }
        if para.spacing_policy.automatic_after {
            self.write_control_word("saauto", Some(1))?;
        }
        if !para.spacing_policy.snap_to_line_grid {
            self.write_control_word("nosnaplinegrid", None)?;
        }
        if para.spacing_policy.contextual_spacing {
            self.write_control_word("contextualspace", None)?;
        }
        if para.spacing.line != 0 {
            self.write_control_word("sl", Some(para.spacing.line))?;
            if para.spacing.line_multiple {
                self.write_control_word("slmult", Some(1))?;
            }
        }

        // Indentation
        if para.indentation.left != 0 {
            self.write_control_word("li", Some(para.indentation.left))?;
        }
        if para.indentation.right != 0 {
            self.write_control_word("ri", Some(para.indentation.right))?;
        }
        if para.indentation.first_line != 0 {
            self.write_control_word("fi", Some(para.indentation.first_line))?;
        }
        let logical = para.logical_indentation;
        if let Some(v) = logical.start {
            self.write_control_word("lin", Some(v))?;
        }
        if let Some(v) = logical.end {
            self.write_control_word("rin", Some(v))?;
        }
        if let Some(v) = logical.first_line_character_units {
            self.write_control_word("cufi", Some(v))?;
        }
        if let Some(v) = logical.left_character_units {
            self.write_control_word("culi", Some(v))?;
        }
        if let Some(v) = logical.right_character_units {
            self.write_control_word("curi", Some(v))?;
        }
        if logical.mirrored {
            self.write_control_word("indmirror", None)?;
        }

        // Borders (if any)
        self.write_borders(&para.borders)?;

        // Shading (if any)
        self.write_shading(&para.shading)?;

        // Custom tab stops, retained in declaration order.
        for tab in &para.tab_stops {
            self.write_tab_stop(*tab)?;
        }

        if let Some(drop_cap) = para.drop_cap {
            drop_cap.validate().map_err(|error| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!("invalid RTF paragraph drop cap: {error}"),
                )
            })?;
            self.write_control_word("dropcapli", Some(i32::from(drop_cap.line_count())))?;
            self.write_control_word("dropcapt", Some(drop_cap.kind().as_rtf_value()))?;
        }

        // Keep together
        if para.keep_together {
            self.write_control_word("keep", None)?;
        }

        // Keep with next
        if para.keep_next {
            self.write_control_word("keepn", None)?;
        }

        // Side-by-side
        if para.side_by_side {
            self.write_control_word("sbys", None)?;
        }

        // Page break before
        if para.page_break_before {
            self.write_control_word("pagebb", None)?;
        }

        // Widow control
        if para.widow_control {
            self.write_control_word("widctlpar", None)?;
        }
        if para.no_line_numbering {
            self.write_control_word("noline", None)?;
        }
        if para.no_auto_tab_indent {
            self.write_control_word("notabind", None)?;
        }

        let breaking = para.line_breaking;
        if breaking.automatic_hyphenation {
            self.write_control_word("hyphpar", None)?;
        }
        match breaking.wrapping {
            ParagraphWrapping::Default => {},
            ParagraphWrapping::NoCharacterWrap => self.write_control_word("nocwrap", None)?,
            ParagraphWrapping::NoWordWrap => self.write_control_word("nowwrap", None)?,
            ParagraphWrapping::NoOverflow => self.write_control_word("nooverflow", None)?,
        }
        if breaking.auto_space_alphabetic {
            self.write_control_word("aspalpha", None)?;
        }
        if breaking.auto_space_numbers {
            self.write_control_word("aspnum", None)?;
        }
        match breaking.font_alignment {
            ParagraphFontAlignment::Auto => {},
            ParagraphFontAlignment::Hanging => self.write_control_word("fahang", None)?,
            ParagraphFontAlignment::Center => self.write_control_word("facenter", None)?,
            ParagraphFontAlignment::Roman => self.write_control_word("faroman", None)?,
            ParagraphFontAlignment::Variable => self.write_control_word("favar", None)?,
            ParagraphFontAlignment::Fixed => self.write_control_word("fafixed", None)?,
        }
        if breaking.adjust_right_indent {
            self.write_control_word("adjustright", None)?;
        }

        if let Some(list_override) = para.list_override {
            self.write_control_word("ls", Some(list_override))?;
        }
        if let Some(list_level) = para.list_level {
            if list_level > 8 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF paragraph list levels must be between zero and eight",
                ));
            }
            self.write_control_word("ilvl", Some(i32::from(list_level)))?;
        }

        Ok(())
    }

    /// Write borders
    pub(in super::super) fn write_borders(&mut self, borders: &Borders) -> io::Result<()> {
        if !borders.has_any_border() {
            return Ok(());
        }

        // Top border
        if borders.top.is_visible() {
            self.write_border("brdrt", &borders.top)?;
        }

        // Bottom border
        if borders.bottom.is_visible() {
            self.write_border("brdrb", &borders.bottom)?;
        }

        // Left border
        if borders.left.is_visible() {
            self.write_border("brdrl", &borders.left)?;
        }

        // Right border
        if borders.right.is_visible() {
            self.write_border("brdrr", &borders.right)?;
        }

        // Bar border
        if borders.bar.is_visible() {
            self.write_border("brdrbar", &borders.bar)?;
        }

        // Between border
        if borders.between.is_visible() {
            self.write_border("brdrbtw", &borders.between)?;
        }

        Ok(())
    }

    /// Write a single border
    pub(in super::super) fn write_border(
        &mut self,
        control: &str,
        border: &Border,
    ) -> io::Result<()> {
        self.write_control_word(control, None)?;

        // Border style
        let style_word = match border.style {
            BorderStyle::None => return Ok(()),
            BorderStyle::Single => "brdrs",
            BorderStyle::Thick => "brdrth",
            BorderStyle::Dotted => "brdrdot",
            BorderStyle::Dashed => "brdrdash",
            BorderStyle::DashSmallGap => "brdrdashsm",
            BorderStyle::DotDash => "brdrdashd",
            BorderStyle::DotDotDash => "brdrdashdd",
            BorderStyle::Double => "brdrdb",
            BorderStyle::Triple => "brdrtriple",
            BorderStyle::ThickThinSmall => "brdrtnthsg",
            BorderStyle::ThinThickSmall | BorderStyle::ThickThinMedium => "brdrtnthmg",
            BorderStyle::ThinThickThinSmall => "brdrtnthtnsg",
            BorderStyle::ThinThickMedium => "brdrthtnmg",
            BorderStyle::ThinThickThinMedium => "brdrtnthtnmg",
            BorderStyle::ThickThinLarge => "brdrtnthlg",
            BorderStyle::ThinThickLarge => "brdrththlg",
            BorderStyle::ThinThickThinLarge => "brdrtnthtnlg",
            BorderStyle::Wavy => "brdrwavy",
            BorderStyle::WavyDouble => "brdrwavydb",
            BorderStyle::Striped => "brdrdashdotstr",
            BorderStyle::Embossed => "brdremboss",
            BorderStyle::Engraved => "brdrengrave",
            BorderStyle::Outset => "brdroutset",
            BorderStyle::Inset => "brdrinset",
            BorderStyle::Hairline => "brdrhair",
        };
        self.write_control_word(style_word, None)?;

        // Border width
        self.write_control_word("brdrw", Some(border.width))?;

        // Border color
        if border.color_ref != 0 {
            self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
        }

        // Border space
        if border.space != 0 {
            self.write_control_word("brsp", Some(border.space))?;
        }

        // Border shadow
        if border.shadow {
            self.write_control_word("brdrsh", None)?;
        }

        // Border frame
        if border.frame {
            self.write_control_word("brdrframe", None)?;
        }

        Ok(())
    }

    /// Write shading
    pub(in super::super) fn write_shading(&mut self, shading: &Shading) -> io::Result<()> {
        shading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if !shading.is_present() {
            return Ok(());
        }

        let pattern_value = match (shading.amount, shading.pattern) {
            (Some(amount), _) => Some(i32::from(amount)),
            (None, Some(ShadingPattern::Clear)) => Some(0),
            (None, Some(ShadingPattern::Solid)) => Some(10_000),
            (None, Some(ShadingPattern::Percent5)) => Some(500),
            (None, Some(ShadingPattern::Percent10)) => Some(1000),
            (None, Some(ShadingPattern::Percent12)) => Some(1250),
            (None, Some(ShadingPattern::Percent15)) => Some(1500),
            (None, Some(ShadingPattern::Percent20)) => Some(2000),
            (None, Some(ShadingPattern::Percent25)) => Some(2500),
            (None, Some(ShadingPattern::Percent30)) => Some(3000),
            (None, Some(ShadingPattern::Percent35)) => Some(3500),
            (None, Some(ShadingPattern::Percent40)) => Some(4000),
            (None, Some(ShadingPattern::Percent45)) => Some(4500),
            (None, Some(ShadingPattern::Percent50)) => Some(5000),
            (None, Some(ShadingPattern::Percent55)) => Some(5500),
            (None, Some(ShadingPattern::Percent60)) => Some(6000),
            (None, Some(ShadingPattern::Percent62)) => Some(6250),
            (None, Some(ShadingPattern::Percent65)) => Some(6500),
            (None, Some(ShadingPattern::Percent70)) => Some(7000),
            (None, Some(ShadingPattern::Percent75)) => Some(7500),
            (None, Some(ShadingPattern::Percent80)) => Some(8000),
            (None, Some(ShadingPattern::Percent85)) => Some(8500),
            (None, Some(ShadingPattern::Percent87)) => Some(8750),
            (None, Some(ShadingPattern::Percent90)) => Some(9000),
            (None, Some(ShadingPattern::Percent95)) => Some(9500),
            (None, None | Some(_)) => None,
        };

        if let Some(pattern_value) = pattern_value {
            self.write_control_word("shading", Some(pattern_value))?;
        } else if let Some(pattern) = shading.pattern {
            self.write_control_word(explicit_shading_pattern_word(pattern, false)?, None)?;
        }

        // Foreground color
        if let Some(color) = shading.foreground_color {
            self.write_control_word("cfpat", Some(i32::from(color)))?;
        }

        // Background color
        if let Some(color) = shading.background_color {
            self.write_control_word("cbpat", Some(i32::from(color)))?;
        }

        Ok(())
    }

    /// Write tab stop
    ///
    pub(in super::super) fn write_tab_stop(&mut self, tab: TabStop) -> io::Result<()> {
        // The left kind is implicit. A bar tab uses `tbN` as its terminator.
        match tab.alignment {
            TabAlignment::Left | TabAlignment::Bar => {},
            TabAlignment::Right => self.write_control_word("tqr", None)?,
            TabAlignment::Center => self.write_control_word("tqc", None)?,
            TabAlignment::Decimal => self.write_control_word("tqdec", None)?,
        }

        // Tab leader
        match tab.leader {
            TabLeader::None => {},
            TabLeader::Dot => self.write_control_word("tldot", None)?,
            TabLeader::MiddleDot => self.write_control_word("tlmdot", None)?,
            TabLeader::Hyphen => self.write_control_word("tlhyph", None)?,
            TabLeader::Underscore => self.write_control_word("tlul", None)?,
            TabLeader::ThickLine => self.write_control_word("tlth", None)?,
            TabLeader::Equal => self.write_control_word("tleq", None)?,
        }

        self.write_control_word(
            if tab.alignment == TabAlignment::Bar {
                "tb"
            } else {
                "tx"
            },
            Some(tab.position),
        )?;

        Ok(())
    }
}
