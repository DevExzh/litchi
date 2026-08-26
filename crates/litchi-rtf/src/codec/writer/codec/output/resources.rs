//! RTF resource-table and style output.

#![allow(
    clippy::shadow_reuse,
    reason = "serialization helpers deliberately rebind a working value as the output is assembled"
)]
use super::super::{
    DefaultFormattingDestination, DocumentDataStore, DocumentDefaultFormatting, DocumentGenerator,
    DocumentMathProperties, DocumentTheme, EmbeddedFontFormat, EndnoteRestart, FileLocation,
    FileTable, FontFamily, FontPitch, FootnoteRestart, HashSet, ImageType, LatentStyles,
    LegacyNumberingAlignment, LegacyNumberingFormat, LegacySectionNumbering, List, ListFollow,
    ListJustification, ListLevel, ListLevelType, ListOverrideTable, ListTable, MailMerge,
    NoteNumberingStyle, NoteOptions, NotePlacement, NoteSeparatorElement, NoteSeparatorKind,
    NoteSeparatorTable, ParagraphGroupPropertyTable, Picture, PresentNoteKinds,
    ProtectionUserTable, Revision, RevisionAuthor, RevisionSaveMetadata, RtfWriter, StoryDrawing,
    Style, StyleSheet, StyleType, Write, XmlNamespace, annotation, io,
};

impl<W: Write> RtfWriter<W> {
    /// Write font table
    pub(in super::super) fn write_font_table(&mut self) -> io::Result<()> {
        self.font_table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if self.font_table.fonts().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("fonttbl", None)?;

        // Clone fonts to avoid borrowing issues
        let fonts: Vec<_> = self.font_table.fonts().to_vec();
        for (idx, font) in fonts.iter().enumerate() {
            let font_ref = u16::try_from(idx).map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF font-table index exceeds the u16 range",
                )
            })?;
            if !self.font_table.is_defined(font_ref) {
                continue;
            }
            font.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_str("{")?;
            self.write_control_word("f", Some(i32::from(font_ref)))?;

            // Write font family
            match font.family {
                FontFamily::Roman => self.write_control_word("froman", None)?,
                FontFamily::Swiss => self.write_control_word("fswiss", None)?,
                FontFamily::Modern => self.write_control_word("fmodern", None)?,
                FontFamily::Script => self.write_control_word("fscript", None)?,
                FontFamily::Decor => self.write_control_word("fdecor", None)?,
                FontFamily::Tech => self.write_control_word("ftech", None)?,
                FontFamily::Nil => self.write_control_word("fnil", None)?,
            }
            if font.bidi {
                self.write_control_word("fbidi", None)?;
            }

            // Write charset
            if let Some(charset) = font.charset {
                self.write_control_word("fcharset", Some(i32::from(charset.id())))?;
            }
            if let Some(theme) = font.theme {
                self.write_control_word(theme.control_word(), None)?;
            }
            self.write_control_word(
                "fprq",
                Some(match font.pitch {
                    FontPitch::Default => 0,
                    FontPitch::Fixed => 1,
                    FontPitch::Variable => 2,
                }),
            )?;
            if let Some(code_page) = font.code_page {
                self.write_control_word("cpg", Some(i32::from(code_page.id())))?;
            }
            if let Some(panose) = font.panose {
                self.write_str("{\\*")?;
                self.write_control_word("panose", None)?;
                self.write_str(" ")?;
                for byte in panose {
                    write!(self.writer, "{byte:02x}")?;
                }
                self.write_str("}")?;
            }
            if let Some(name) = font.non_tagged_name.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("fname", None)?;
                self.write_str(" ")?;
                self.write_text(name)?;
                self.write_str("}")?;
            }
            if let Some(embedded) = &font.embedded {
                self.write_str("{\\*")?;
                self.write_control_word("fontemb", None)?;
                self.write_control_word(
                    match embedded.format {
                        EmbeddedFontFormat::Nil => "ftnil",
                        EmbeddedFontFormat::TrueType => "fttruetype",
                    },
                    None,
                )?;
                if let Some(name) = embedded.file_name.as_deref() {
                    self.write_str("{\\*")?;
                    self.write_control_word("fontfile", None)?;
                    if let Some(code_page) = embedded.file_code_page {
                        self.write_control_word("cpg", Some(i32::from(code_page.id())))?;
                    }
                    self.write_str(" ")?;
                    self.write_text(name)?;
                    self.write_str("}")?;
                }
                if let Some(data) = &embedded.data {
                    self.write_str(" ")?;
                    for byte in data {
                        write!(self.writer, "{byte:02x}")?;
                    }
                }
                self.write_str("}")?;
            }

            // Write font name
            self.write_str(" ")?;
            self.write_text(font.name.as_ref())?;
            if let Some(name) = font.alternate_name.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("falt", None)?;
                self.write_str(" ")?;
                self.write_text(name)?;
                self.write_str("}")?;
            }
            self.write_str(";")?;
            self.write_str("}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write the optional external-file metadata table.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_file_table(&mut self, table: Option<&FileTable<'_>>) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{\\*")?;
        self.write_control_word("filetbl", None)?;
        for entry in table.entries() {
            self.write_str("{")?;
            self.write_control_word("file", None)?;
            self.write_control_word("fid", Some(entry.id.cast_signed()))?;
            if let Some(level) = entry.relative_path_level {
                self.write_control_word("frelative", Some(i32::from(level)))?;
            }
            if let Some(os) = entry.operating_system {
                self.write_control_word("fosnum", Some(i32::from(os)))?;
            }
            if entry.valid_on.mac {
                self.write_control_word("fvalidmac", None)?;
            }
            if entry.valid_on.dos {
                self.write_control_word("fvaliddos", None)?;
            }
            if entry.valid_on.ntfs {
                self.write_control_word("fvalidntfs", None)?;
            }
            if entry.valid_on.hpfs {
                self.write_control_word("fvalidhpfs", None)?;
            }
            match entry.location {
                FileLocation::Local => {},
                FileLocation::Network => self.write_control_word("fnetwork", None)?,
                FileLocation::NonFileSystem => {
                    self.write_control_word("fnonfilesys", None)?;
                },
            }
            self.write_str(" ")?;
            self.write_text(entry.name.as_ref())?;
            self.write_str(";}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write color table
    pub(in super::super) fn write_color_table(&mut self) -> io::Result<()> {
        if self.color_table.colors().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("colortbl", None)?;

        // Clone colors to avoid borrowing issues
        let colors: Vec<_> = self.color_table.colors().to_vec();
        for (index, color) in colors.iter().enumerate() {
            let automatic = u16::try_from(index)
                .ok()
                .is_some_and(|reference| self.color_table.is_automatic(reference));
            if !automatic {
                self.write_control_word("red", Some(i32::from(color.red)))?;
                self.write_control_word("green", Some(i32::from(color.green)))?;
                self.write_control_word("blue", Some(i32::from(color.blue)))?;
            }
            self.write_str(";")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write the list-definition table.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_picture(&mut self, picture: &Picture<'_>) -> io::Result<()> {
        picture
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\pict")?;
        match picture.image_type {
            ImageType::Emf => self.write_control_word("emfblip", None)?,
            ImageType::Wmf => self.write_control_word("wmetafile", Some(8))?,
            ImageType::Png => self.write_control_word("pngblip", None)?,
            ImageType::Jpeg => self.write_control_word("jpegblip", None)?,
            ImageType::Dib if picture.bitmap.windows_bitmap => {
                self.write_control_word("wbitmap", Some(0))?;
            },
            ImageType::Dib => self.write_control_word("dibitmap", Some(0))?,
            ImageType::Pict => self.write_control_word("macpict", None)?,
            ImageType::Unknown => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "cannot write a picture with unknown image type",
                ));
            },
        }
        for (control, value) in [
            ("picw", picture.width),
            ("pich", picture.height),
            ("picwgoal", picture.goal_width),
            ("pichgoal", picture.goal_height),
            ("picscalex", picture.scale_x),
            ("picscaley", picture.scale_y),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if picture.scaled {
            self.write_control_word("picscaled", None)?;
        }
        for (control, value) in [
            ("piccropl", picture.crop.left),
            ("piccropr", picture.crop.right),
            ("piccropt", picture.crop.top),
            ("piccropb", picture.crop.bottom),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if picture.bitmap.bitmap_source {
            self.write_control_word("picbmp", None)?;
        }
        for (control, value) in [
            ("picbpp", picture.bitmap.bits_per_pixel.map(i32::from)),
            (
                "wbmbitspixel",
                picture.bitmap.windows_bits_per_pixel.map(i32::from),
            ),
            ("wbmplanes", picture.bitmap.planes.map(i32::from)),
            (
                "wbmwidthbytes",
                picture
                    .bitmap
                    .width_bytes
                    .and_then(|value| i32::try_from(value).ok()),
            ),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if let Some(identity) = &picture.identity {
            if let Some(tag) = identity.tag {
                self.write_control_word("bliptag", Some(tag))?;
            }
            if let Some(upi) = identity.units_per_inch {
                self.write_control_word("blipupi", Some(i32::from(upi)))?;
            }
            if let Some(uid) = &identity.uid {
                self.write_str("{\\*\\blipuid ")?;
                for byte in uid.iter() {
                    write!(self.writer, "{byte:02x}")?;
                }
                self.write_str("}")?;
            }
        }
        if let Some(properties) = &picture.shape_properties {
            self.write_picture_shape_properties(properties)?;
        }
        self.write_str(" ")?;
        for byte in picture.data.iter() {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")
    }

    /// Write the list-definition table.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_list_table(&mut self, table: &ListTable<'_>) -> io::Result<()> {
        self.write_list_table_with_pictures(table, &[])
    }

    pub(in super::super) fn write_list_table_with_pictures(
        &mut self,
        table: &ListTable<'_>,
        pictures: &[Picture<'_>],
    ) -> io::Result<()> {
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if table.lists().is_empty() && table.picture_bullet_count == 0 {
            return Ok(());
        }
        if table.lists().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list table exceeds the supported list count",
            ));
        }
        self.write_str("{\\*\\listtable")?;
        if table.picture_bullet_count != 0 {
            self.write_str("{\\*\\listpicture")?;
            for slot in 0..table.picture_bullet_count as usize {
                let Some(index) = table
                    .picture_bullet_picture_indices()
                    .get(slot)
                    .copied()
                    .flatten()
                else {
                    continue;
                };
                let picture = pictures.get(index).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF list-picture index is outside the document picture store",
                    )
                })?;
                self.write_str("{\\*\\shppict")?;
                self.write_picture(picture)?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        for list in table.lists() {
            self.write_list_definition(list)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn write_list_definition(&mut self, list: &List<'_>) -> io::Result<()> {
        if list.levels.len() > 9 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF lists cannot contain more than nine levels",
            ));
        }
        if list.simple && list.levels.len() > 1 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "a simple RTF list cannot contain more than one level",
            ));
        }
        if list.simple && list.hybrid {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "an RTF list cannot be both simple and hybrid",
            ));
        }
        self.write_str("{")?;
        self.write_control_word("list", None)?;
        self.write_control_word("listtemplateid", Some(list.template_id))?;
        if list.simple {
            self.write_control_word("listsimple", None)?;
        }
        if list.hybrid {
            self.write_control_word("listhybrid", None)?;
        }
        for level in &list.levels {
            self.write_list_level(level)?;
        }
        if !list.name.is_empty() {
            self.write_str("{")?;
            self.write_control_word("listname", None)?;
            self.write_str(" ")?;
            self.write_text(list.name.as_ref())?;
            self.write_str(";}")?;
        }
        if !list.style_name.is_empty() {
            self.write_str("{\\*")?;
            self.write_control_word("liststylename", None)?;
            self.write_str(" ")?;
            self.write_text(list.style_name.as_ref())?;
            self.write_str(";}")?;
        }
        if let Some(priority) = list.style_priority {
            self.write_control_word("spriority", Some(priority))?;
        }
        self.write_control_word("listid", Some(list.id))?;
        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn write_list_level(&mut self, level: &ListLevel<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("listlevel", None)?;
        self.write_control_word(
            "levelnfc",
            Some(Self::list_level_type_value(level.level_type)),
        )?;
        self.write_control_word(
            "leveljc",
            Some(match level.justification {
                ListJustification::Left => 0,
                ListJustification::Center => 1,
                ListJustification::Right => 2,
            }),
        )?;
        self.write_control_word(
            "levelfollow",
            Some(match level.follow {
                ListFollow::Tab => 0,
                ListFollow::Space => 1,
                ListFollow::Nothing => 2,
            }),
        )?;
        self.write_control_word("levelstartat", Some(level.start_at))?;
        self.write_control_word("levelspace", Some(level.space))?;
        self.write_control_word("levelindent", Some(level.indent))?;
        if let Some(left_indent) = level.left_indent {
            self.write_control_word("li", Some(left_indent))?;
        }
        if let Some(first_line_indent) = level.first_line_indent {
            self.write_control_word("fi", Some(first_line_indent))?;
        }
        for tab in &level.tabs {
            self.write_control_word("tx", Some(*tab))?;
        }
        if level.tentative {
            self.write_control_word("lvltentative", None)?;
        }
        if level.legal_format {
            self.write_control_word("levellegal", None)?;
        }
        if level.no_restart {
            self.write_control_word("levelnorestart", None)?;
        }
        if level.legacy {
            self.write_control_word("levelold", None)?;
        }
        if level.include_previous {
            self.write_control_word("levelprev", None)?;
        }
        if level.include_previous_space {
            self.write_control_word("levelprevspace", None)?;
        }
        if let Some(template_id) = level.template_id {
            self.write_control_word("leveltemplateid", Some(template_id))?;
        }
        if let Some(picture_index) = level.picture_index {
            self.write_control_word(
                "levelpicture",
                Some(i32::try_from(picture_index).map_err(|_err| {
                    io::Error::new(io::ErrorKind::InvalidInput, "invalid list picture index")
                })?),
            )?;
        }
        self.write_list_level_text(level.number_text.as_ref(), level.number_positions.as_ref())?;
        if level.font_ref != 0 {
            self.write_control_word("f", Some(i32::from(level.font_ref)))?;
        }
        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn list_level_type_value(level_type: ListLevelType) -> i32 {
        match level_type {
            ListLevelType::Decimal => 0,
            ListLevelType::UpperRoman => 1,
            ListLevelType::LowerRoman => 2,
            ListLevelType::UpperLetter => 3,
            ListLevelType::LowerLetter => 4,
            ListLevelType::Ordinal => 5,
            ListLevelType::CardinalText => 6,
            ListLevelType::OrdinalText => 7,
            ListLevelType::Bullet => 23,
            ListLevelType::None => 255,
            ListLevelType::Other(value) => value,
        }
    }

    pub(in super::super) fn write_list_level_text(
        &mut self,
        text: &str,
        positions: &str,
    ) -> io::Result<()> {
        let count = u8::try_from(text.chars().count()).map_err(|_err| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list level text cannot exceed 255 characters",
            )
        })?;
        self.write_str("{")?;
        self.write_control_word("leveltext", None)?;
        self.write_hex_byte(count)?;
        for ch in text.chars() {
            if u32::from(ch) <= u8::MAX.into() && (ch.is_control() || !ch.is_ascii()) {
                self.write_hex_byte(ch as u8)?;
            } else {
                let mut buffer = [0; 4];
                self.write_text(ch.encode_utf8(&mut buffer))?;
            }
        }
        self.write_str(";}")?;

        self.write_str("{")?;
        self.write_control_word("levelnumbers", None)?;
        if positions.is_empty() {
            for (index, ch) in text.chars().enumerate() {
                if u32::from(ch) <= 8 {
                    let position = u8::try_from(index + 1).map_err(|_err| {
                        io::Error::new(io::ErrorKind::InvalidInput, "invalid RTF list placeholder")
                    })?;
                    self.write_hex_byte(position)?;
                }
            }
        } else {
            for byte in positions.bytes() {
                self.write_hex_byte(byte)?;
            }
        }
        self.write_str(";}")?;
        Ok(())
    }

    pub(in super::super) fn write_hex_byte(&mut self, value: u8) -> io::Result<()> {
        write!(self.writer, "\\'{value:02x}")
    }

    /// Write the list-override table.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_list_override_table(&mut self, table: &ListOverrideTable) -> io::Result<()> {
        if table.overrides().is_empty() {
            return Ok(());
        }
        if table.overrides().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list override table exceeds the supported entry count",
            ));
        }
        self.write_str("{\\*\\listoverridetable")?;
        for entry in table.overrides() {
            self.write_str("{")?;
            self.write_control_word("listoverride", None)?;
            self.write_control_word("listid", Some(entry.list_id))?;
            self.write_control_word(
                "listoverridecount",
                Some(i32::from(entry.level_count_override.unwrap_or_else(|| {
                    u8::try_from(entry.levels.len())
                        .unwrap_or_else(|_| u8::from(entry.start_at_override.is_some()))
                }))),
            )?;
            if entry.levels.is_empty()
                && let Some(start_at) = entry.start_at_override
            {
                self.write_str("{")?;
                self.write_control_word("lfolevel", None)?;
                self.write_control_word("listoverridestartat", None)?;
                self.write_control_word("levelstartat", Some(start_at))?;
                self.write_str("}")?;
            }
            for level in &entry.levels {
                self.write_str("{")?;
                self.write_control_word("lfolevel", None)?;
                if level.format_override {
                    self.write_control_word("listoverrideformat", None)?;
                }
                if let Some(start_at) = level.start_at {
                    self.write_control_word("listoverridestartat", None)?;
                    self.write_control_word("levelstartat", Some(start_at))?;
                }
                self.write_str("}")?;
            }
            self.write_control_word("ls", Some(entry.index))?;
            self.write_str("}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write ordered inert legacy section-numbering defaults.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_legacy_section_numbering(
        &mut self,
        numbering: &LegacySectionNumbering<'_>,
    ) -> io::Result<()> {
        numbering
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for level in numbering.levels() {
            self.write_str("{\\*")?;
            self.write_control_word("pnseclvl", Some(i32::from(level.level)))?;
            self.write_control_word(
                match level.format {
                    LegacyNumberingFormat::Decimal => "pndec",
                    LegacyNumberingFormat::UpperRoman => "pnucrm",
                    LegacyNumberingFormat::LowerRoman => "pnlcrm",
                    LegacyNumberingFormat::UpperLetter => "pnucltr",
                    LegacyNumberingFormat::LowerLetter => "pnlcltr",
                },
                None,
            )?;
            if let Some(alignment) = level.alignment {
                self.write_control_word(
                    match alignment {
                        LegacyNumberingAlignment::Left => "pnql",
                        LegacyNumberingAlignment::Center => "pnqc",
                        LegacyNumberingAlignment::Right => "pnqr",
                    },
                    None,
                )?;
            }
            if let Some(start_at) = level.start_at {
                self.write_control_word("pnstart", Some(start_at))?;
            }
            if let Some(indent) = level.indent {
                self.write_control_word("pnindent", Some(indent))?;
            }
            if let Some(space) = level.space {
                self.write_control_word("pnsp", Some(space))?;
            }
            if level.hanging {
                self.write_control_word("pnhang", None)?;
            }
            if level.previous {
                self.write_control_word("pnprev", None)?;
            }
            if let Some(font_ref) = level.font_ref {
                self.write_control_word("pnf", Some(i32::from(font_ref)))?;
            }
            if !level.text_before.is_empty() {
                self.write_str("{")?;
                self.write_control_word("pntxtb", None)?;
                self.write_str(" ")?;
                self.write_text(level.text_before.as_ref())?;
                self.write_str("}")?;
            }
            if !level.text_after.is_empty() {
                self.write_str("{")?;
                self.write_control_word("pntxta", None)?;
                self.write_str(" ")?;
                self.write_text(level.text_after.as_ref())?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write the inert paragraph-group property table.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_paragraph_group_table(
        &mut self,
        table: Option<&ParagraphGroupPropertyTable>,
    ) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\pgptbl")?;
        for entry in table.entries() {
            self.write_str("{")?;
            self.write_control_word("pgp", None)?;
            self.write_control_word(
                "ipgp",
                Some(i32::try_from(entry.parent_id).map_err(|_err| {
                    io::Error::new(io::ErrorKind::InvalidInput, "invalid pgp parent ID")
                })?),
            )?;
            self.write_control_word("itap", Some(i32::from(entry.table_nesting_level)))?;
            self.write_control_word("li", Some(entry.left_indent))?;
            self.write_control_word("ri", Some(entry.right_indent))?;
            self.write_control_word("sb", Some(entry.space_before))?;
            self.write_control_word("sa", Some(entry.space_after))?;
            self.write_borders(&entry.borders)?;
            self.write_str("}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write explicit document-level footnote and endnote configuration.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_note_options(&mut self, options: &NoteOptions) -> io::Result<()> {
        options
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        if let Some(value) = options.present_kinds {
            self.write_control_word(
                "fet",
                Some(match value {
                    PresentNoteKinds::FootnotesOnly => 0,
                    PresentNoteKinds::EndnotesOnly => 1,
                    PresentNoteKinds::FootnotesAndEndnotes => 2,
                }),
            )?;
        }
        if let Some(value) = options.footnote_placement {
            self.write_control_word(
                match value {
                    NotePlacement::EndOfSection => "endnotes",
                    NotePlacement::EndOfDocument => "enddoc",
                    NotePlacement::BeneathText => "ftntj",
                    NotePlacement::BottomOfPage => "ftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_start {
            self.write_control_word("ftnstart", Some(value))?;
        }
        if let Some(value) = options.footnote_restart {
            self.write_control_word(
                match value {
                    FootnoteRestart::Continuous => "ftnrstcont",
                    FootnoteRestart::EachSection => "ftnrestart",
                    FootnoteRestart::EachPage => "ftnrstpg",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_numbering {
            self.write_control_word(Self::note_numbering_control(value, false), None)?;
        }
        if let Some(value) = options.endnote_placement {
            self.write_control_word(
                match value {
                    NotePlacement::EndOfSection => "aendnotes",
                    NotePlacement::EndOfDocument => "aenddoc",
                    NotePlacement::BeneathText => "aftntj",
                    NotePlacement::BottomOfPage => "aftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_start {
            self.write_control_word("aftnstart", Some(value))?;
        }
        if let Some(value) = options.endnote_restart {
            self.write_control_word(
                match value {
                    EndnoteRestart::Continuous => "aftnrstcont",
                    EndnoteRestart::EachSection => "aftnrestart",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_numbering {
            self.write_control_word(Self::note_numbering_control(value, true), None)?;
        }
        Ok(())
    }

    pub(in super::super) fn note_numbering_control(
        style: NoteNumberingStyle,
        endnote: bool,
    ) -> &'static str {
        match (endnote, style) {
            (false, NoteNumberingStyle::Arabic) => "ftnnar",
            (false, NoteNumberingStyle::LowercaseLetter) => "ftnnalc",
            (false, NoteNumberingStyle::UppercaseLetter) => "ftnnauc",
            (false, NoteNumberingStyle::LowercaseRoman) => "ftnnrlc",
            (false, NoteNumberingStyle::UppercaseRoman) => "ftnnruc",
            (false, NoteNumberingStyle::Chicago) => "ftnnchi",
            (false, NoteNumberingStyle::KoreanChosung) => "ftnnchosung",
            (false, NoteNumberingStyle::Circle) => "ftnncnum",
            (false, NoteNumberingStyle::KanjiDigitless) => "ftnndbnum",
            (false, NoteNumberingStyle::KanjiWithDigit) => "ftnndbnumd",
            (false, NoteNumberingStyle::KanjiThree) => "ftnndbnumt",
            (false, NoteNumberingStyle::KanjiFour) => "ftnndbnumk",
            (false, NoteNumberingStyle::DoubleByte) => "ftnndbar",
            (false, NoteNumberingStyle::KoreanGanada) => "ftnnganada",
            (false, NoteNumberingStyle::ChineseOne) => "ftnngbnum",
            (false, NoteNumberingStyle::ChineseTwo) => "ftnngbnumd",
            (false, NoteNumberingStyle::ChineseThree) => "ftnngbnuml",
            (false, NoteNumberingStyle::ChineseFour) => "ftnngbnumk",
            (false, NoteNumberingStyle::ZodiacOne) => "ftnnzodiac",
            (false, NoteNumberingStyle::ZodiacTwo) => "ftnnzodiacd",
            (false, NoteNumberingStyle::ZodiacThree) => "ftnnzodiacl",
            (true, NoteNumberingStyle::Arabic) => "aftnnar",
            (true, NoteNumberingStyle::LowercaseLetter) => "aftnnalc",
            (true, NoteNumberingStyle::UppercaseLetter) => "aftnnauc",
            (true, NoteNumberingStyle::LowercaseRoman) => "aftnnrlc",
            (true, NoteNumberingStyle::UppercaseRoman) => "aftnnruc",
            (true, NoteNumberingStyle::Chicago) => "aftnnchi",
            (true, NoteNumberingStyle::KoreanChosung) => "aftnnchosung",
            (true, NoteNumberingStyle::Circle) => "aftnncnum",
            (true, NoteNumberingStyle::KanjiDigitless) => "aftnndbnum",
            (true, NoteNumberingStyle::KanjiWithDigit) => "aftnndbnumd",
            (true, NoteNumberingStyle::KanjiThree) => "aftnndbnumt",
            (true, NoteNumberingStyle::KanjiFour) => "aftnndbnumk",
            (true, NoteNumberingStyle::DoubleByte) => "aftnndbar",
            (true, NoteNumberingStyle::KoreanGanada) => "aftnnganada",
            (true, NoteNumberingStyle::ChineseOne) => "aftnngbnum",
            (true, NoteNumberingStyle::ChineseTwo) => "aftnngbnumd",
            (true, NoteNumberingStyle::ChineseThree) => "aftnngbnuml",
            (true, NoteNumberingStyle::ChineseFour) => "aftnngbnumk",
            (true, NoteNumberingStyle::ZodiacOne) => "aftnnzodiac",
            (true, NoteNumberingStyle::ZodiacTwo) => "aftnnzodiacd",
            (true, NoteNumberingStyle::ZodiacThree) => "aftnnzodiacl",
        }
    }

    /// Write ordered semantic note-separator destinations.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_note_separators(&mut self, table: &NoteSeparatorTable<'_>) -> io::Result<()> {
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for separator in table.entries() {
            self.write_str("{\\*")?;
            self.write_control_word(
                match separator.kind {
                    NoteSeparatorKind::FootnoteSeparator => "ftnsep",
                    NoteSeparatorKind::FootnoteContinuationSeparator => "ftnsepc",
                    NoteSeparatorKind::FootnoteContinuationNotice => "ftncn",
                    NoteSeparatorKind::EndnoteSeparator => "aftnsep",
                    NoteSeparatorKind::EndnoteContinuationSeparator => "aftnsepc",
                    NoteSeparatorKind::EndnoteContinuationNotice => "aftncn",
                },
                None,
            )?;
            self.write_str(" ")?;
            for element in &separator.elements {
                match element {
                    NoteSeparatorElement::Text(text) => self.write_text(text.as_ref())?,
                    NoteSeparatorElement::SeparatorMark => {
                        self.write_control_word("chftnsep", None)?;
                        self.write_str(" ")?;
                    },
                    NoteSeparatorElement::ContinuationSeparatorMark => {
                        self.write_control_word("chftnsepc", None)?;
                        self.write_str(" ")?;
                    },
                    NoteSeparatorElement::ParagraphBreak => {
                        self.write_control_word("par", None)?;
                        self.write_str(" ")?;
                    },
                    NoteSeparatorElement::LineBreak => {
                        self.write_control_word("line", None)?;
                        self.write_str(" ")?;
                    },
                    NoteSeparatorElement::Drawing(StoryDrawing::Shape(index)) => {
                        let shape = separator.shapes.get(*index).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF note separator references a missing shape",
                            )
                        })?;
                        self.write_root_shape(shape)?;
                    },
                    NoteSeparatorElement::Drawing(StoryDrawing::ShapeGroup(index)) => {
                        let group = separator.shape_groups.get(*index).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF note separator references a missing shape group",
                            )
                        })?;
                        self.write_shape_group(group, true)?;
                    },
                }
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write the revision-author table referenced by tracked-change runs.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_revision_table(
        &mut self,
        authors: &[RevisionAuthor<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        if authors.is_empty() && revisions.is_empty() {
            return Ok(());
        }
        if authors.len() > annotation::MAX_REVISION_AUTHORS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF revision-author table exceeds the safety limit",
            ));
        }
        let author_bytes = authors.iter().try_fold(0usize, |total, author| {
            author
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            total.checked_add(author.name.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF aggregate revision-author size overflow",
                )
            })
        })?;
        if author_bytes > annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF aggregate revision-author text exceeds the safety limit",
            ));
        }
        for revision in revisions {
            revision
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let index = usize::try_from(revision.id).map_err(|_err| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author indices cannot be negative",
                )
            })?;
            let author = authors.get(index).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author index is outside revtbl",
                )
            })?;
            if author.name != revision.author {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author does not match its revtbl entry",
                ));
            }
        }

        self.write_str("{\\*\\revtbl")?;
        for author in authors {
            self.write_str("{")?;
            self.write_text(author.name.as_ref())?;
            self.write_str(";}")?;
        }
        self.write_str("}")?;
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_generator(&mut self, generator: Option<&DocumentGenerator<'_>>) -> io::Result<()> {
        let Some(generator) = generator else {
            return Ok(());
        };
        generator
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\generator ")?;
        self.write_destination_text(generator.value.as_ref())?;
        self.write_str(";}")
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_revision_save_metadata(
        &mut self,
        metadata: Option<&RevisionSaveMetadata>,
    ) -> io::Result<()> {
        let Some(metadata) = metadata else {
            return Ok(());
        };
        metadata
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\rsidtbl ")?;
        for id in metadata.ids() {
            self.write_control_word("rsid", Some((*id).cast_signed()))?;
        }
        self.write_str("}")?;
        if let Some(root) = metadata.root() {
            self.write_control_word("rsidroot", Some(root.cast_signed()))?;
        }
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_xml_namespace_table(
        &mut self,
        namespaces: Option<&[XmlNamespace<'_>]>,
    ) -> io::Result<()> {
        let Some(namespaces) = namespaces else {
            return Ok(());
        };
        if namespaces.len() > crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF XML namespace count exceeds the safety limit",
            ));
        }
        let mut total = 0usize;
        let mut ids = HashSet::with_capacity(namespaces.len());
        self.write_str("{\\*\\xmlnstbl ")?;
        for namespace in namespaces {
            namespace
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if !ids.insert(namespace.id) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF XML namespace IDs must be unique",
                ));
            }
            total = total
                .checked_add(namespace.namespace.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF XML namespace aggregate size overflow",
                    )
                })?;
            if total > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF XML namespace aggregate text exceeds the safety limit",
                ));
            }
            self.write_str("{")?;
            self.write_control_word("xmlns", Some(namespace.id.cast_signed()))?;
            self.write_str(" ")?;
            self.write_destination_text(namespace.namespace.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    /// Write inert range-protection usernames without resolving any identity.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_protection_user_table(
        &mut self,
        table: Option<&ProtectionUserTable<'_>>,
    ) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\protusertbl")?;
        for user in table.users() {
            self.write_str("{")?;
            self.write_destination_text(user.name.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_theme(&mut self, theme: Option<&DocumentTheme<'_>>) -> io::Result<()> {
        let Some(theme) = theme else {
            return Ok(());
        };
        theme
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_hex_destination("themedata", theme.data.as_ref())?;
        if let Some(mapping) = theme.color_scheme_mapping.as_deref() {
            self.write_hex_destination("colorschememapping", mapping)?;
        }
        Ok(())
    }

    pub(in super::super) fn write_hex_destination(
        &mut self,
        control: &str,
        data: &[u8],
    ) -> io::Result<()> {
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        for byte in data {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_latent_styles(&mut self, styles: Option<&LatentStyles<'_>>) -> io::Result<()> {
        let Some(styles) = styles else {
            return Ok(());
        };
        styles
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\latentstyles")?;
        self.write_control_word("lsdstimax", Some(styles.max_style_index.cast_signed()))?;
        self.write_optional_bool("lsdlockeddef", styles.locked_default)?;
        self.write_optional_bool("lsdsemihiddendef", styles.semi_hidden_default)?;
        self.write_optional_bool("lsdunhideuseddef", styles.unhide_when_used_default)?;
        self.write_optional_bool("lsdqformatdef", styles.quick_format_default)?;
        if let Some(priority) = styles.priority_default {
            self.write_control_word("lsdprioritydef", Some(i32::from(priority)))?;
        }
        if !styles.exceptions.is_empty() {
            self.write_str("{\\lsdlockedexcept ")?;
            for exception in &styles.exceptions {
                self.write_optional_bool("lsdlocked", exception.locked)?;
                self.write_optional_bool("lsdsemihidden", exception.semi_hidden)?;
                self.write_optional_bool("lsdunhideused", exception.unhide_when_used)?;
                self.write_optional_bool("lsdqformat", exception.quick_format)?;
                if let Some(priority) = exception.priority {
                    self.write_control_word("lsdpriority", Some(i32::from(priority)))?;
                }
                self.write_destination_text(exception.name.as_ref())?;
                self.write_str(";")?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_data_store(
        &mut self,
        data_store: Option<&DocumentDataStore<'_>>,
    ) -> io::Result<()> {
        let Some(data_store) = data_store else {
            return Ok(());
        };
        data_store
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_hex_destination("datastore", data_store.data.as_ref())
    }

    /// Write inert RTF 1.9.1 mail-merge metadata without evaluating it.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_mail_merge(&mut self, merge: Option<&MailMerge<'_>>) -> io::Result<()> {
        let Some(merge) = merge else {
            return Ok(());
        };
        merge
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\mailmerge")?;
        self.write_mail_merge_text("mmconnectstr", merge.connect_string.as_deref())?;
        self.write_mail_merge_text("mmconnectstrdata", merge.connect_string_data.as_deref())?;
        self.write_mail_merge_text("mmquery", merge.query.as_deref())?;
        self.write_mail_merge_text("mmdatasource", merge.data_source.as_deref())?;
        self.write_mail_merge_text("mmheadersource", merge.header_source.as_deref())?;
        if merge.link_to_query {
            self.write_control_word("mmlinktoquery", None)?;
        }
        if let Some(object) = &merge.data_source_object {
            self.write_str("{\\*\\mmodso")?;
            if let Some(value) = object.active_record {
                self.write_control_word("mmodsoactive", Some(value.cast_signed()))?;
            }
            if let Some(value) = object.column_delimiter {
                self.write_control_word("mmodsocoldelim", Some(value))?;
            }
            if let Some(value) = object.column_count {
                self.write_control_word("mmodsocolumn", Some(value.cast_signed()))?;
            }
            self.write_optional_bool("mmodsodynaddr", object.dynamic_address)?;
            self.write_optional_bool("mmodsofhdr", object.first_row_header)?;
            if let Some(value) = object.hash {
                self.write_control_word("mmodsohash", Some(value))?;
            }
            if let Some(value) = object.id {
                self.write_control_word("mmodsolid", Some(value))?;
            }
            if let Some(value) = object.source_type {
                self.write_control_word("mmodsosrc", Some(value.rtf_value()))?;
            }
            self.write_mail_merge_text("mmodsofilter", object.filter.as_deref())?;
            self.write_mail_merge_text("mmodsoname", object.name.as_deref())?;
            self.write_mail_merge_text("mmodsosort", object.sort.as_deref())?;
            self.write_mail_merge_text("mmodsotable", object.table.as_deref())?;
            self.write_mail_merge_text("mmodsoudl", object.udl.as_deref())?;
            self.write_mail_merge_text("mmodsoudldata", object.udl_data.as_deref())?;
            self.write_mail_merge_text("mmodsouniquetag", object.unique_tag.as_deref())?;
            for mapping in &object.field_mappings {
                self.write_str("{\\*\\mmodsofldmpdata")?;
                self.write_control_word(
                    "mmodsofmcolumn",
                    Some(mapping.column.rtf_value().map_err(|error| {
                        io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                    })?),
                )?;
                self.write_mail_merge_text("mmodsoname", Some(mapping.name.as_ref()))?;
                self.write_mail_merge_text("mmodsomappedname", mapping.mapped_name.as_deref())?;
                self.write_str("}")?;
            }
            for value in &object.recipient_data {
                self.write_mail_merge_text("mmodsorecipdata", Some(value.as_ref()))?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_mail_merge_text(
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
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_math_properties(
        &mut self,
        properties: Option<&DocumentMathProperties>,
    ) -> io::Result<()> {
        let Some(properties) = properties else {
            return Ok(());
        };
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\mmathPr")?;
        if let Some(value) = properties.binary_operator_break {
            self.write_control_word("mbrkBin", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.binary_subtraction_break {
            self.write_control_word("mbrkBinSub", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.default_justification {
            self.write_control_word("mdefJc", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.display_defaults {
            self.write_control_word("mdispDef", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.inter_equation_spacing {
            self.write_control_word("minterSp", Some(value))?;
        }
        if let Some(value) = properties.integral_limit_placement {
            self.write_control_word("mintLim", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.intra_equation_spacing {
            self.write_control_word("mintraSp", Some(value))?;
        }
        if let Some(value) = properties.left_margin {
            self.write_control_word("mlMargin", Some(value))?;
        }
        if let Some(value) = properties.math_font {
            self.write_control_word("mmathFont", Some(value.cast_signed()))?;
        }
        if let Some(value) = properties.nary_limit_placement {
            self.write_control_word("mnaryLim", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.post_spacing {
            self.write_control_word("mpostSp", Some(value))?;
        }
        if let Some(value) = properties.pre_spacing {
            self.write_control_word("mpreSp", Some(value))?;
        }
        if let Some(value) = properties.right_margin {
            self.write_control_word("mrMargin", Some(value))?;
        }
        if let Some(value) = properties.small_fractions {
            self.write_control_word("msmallFrac", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.wrap_indent {
            self.write_control_word("mwrapIndent", Some(value))?;
        }
        if let Some(value) = properties.wrap_right {
            self.write_control_word("mwrapRight", Some(value.rtf_value()))?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_optional_bool(
        &mut self,
        control: &str,
        value: Option<bool>,
    ) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write an RTF stylesheet destination.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_stylesheet(&mut self, stylesheet: &StyleSheet<'_>) -> io::Result<()> {
        if stylesheet.styles().is_empty() {
            return Ok(());
        }
        stylesheet
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{")?;
        self.write_control_word("stylesheet", None)?;
        for style in stylesheet.styles() {
            self.write_style_definition(style)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    pub(in super::super) fn write_default_formatting_destinations(
        &mut self,
        defaults: &DocumentDefaultFormatting,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for kind in defaults.destination_order() {
            self.write_str("{\\*")?;
            match kind {
                DefaultFormattingDestination::Character => {
                    let value = defaults.character().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "missing defchp value")
                    })?;
                    self.write_control_word("defchp", None)?;
                    self.write_formatting(&value.formatting)?;
                    for (control, font) in [
                        ("loch", value.low_ansi_font),
                        ("hich", value.high_ansi_font),
                        ("dbch", value.double_byte_font),
                    ] {
                        if let Some(font) = font {
                            self.write_control_word(control, None)?;
                            self.write_control_word("af", Some(i32::from(font)))?;
                        }
                    }
                },
                DefaultFormattingDestination::Paragraph => {
                    let value = defaults.paragraph().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "missing defpap value")
                    })?;
                    self.write_control_word("defpap", None)?;
                    self.write_paragraph_properties(&value.paragraph)?;
                    if let Some(level) = value.table_nesting_level {
                        self.write_control_word("itap", Some(i32::from(level)))?;
                    }
                },
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    pub(in super::super) fn write_style_definition(&mut self, style: &Style<'_>) -> io::Result<()> {
        if style.name.contains(';') {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF style names cannot contain a semicolon",
            ));
        }
        self.write_str("{")?;
        let control = match style.style_type {
            StyleType::Paragraph => "s",
            StyleType::Character => "cs",
            StyleType::Section => "ds",
            StyleType::Table => "ts",
        };
        if style.style_type != StyleType::Paragraph {
            self.write_str("\\*")?;
        }
        self.write_control_word(control, Some(i32::from(style.id)))?;
        if style.style_type == StyleType::Table {
            if style.table_conditional.row_defaults_marker {
                self.write_control_word("tsrowd", None)?;
            }
        } else if !style.table_conditional.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table-style conditional metadata requires a table style",
            ));
        }
        self.write_formatting(&style.formatting)?;
        if let Some(paragraph) = &style.paragraph {
            self.write_paragraph_properties(paragraph)?;
        }
        if let Some(value) = style.based_on {
            self.write_control_word("sbasedon", Some(i32::from(value)))?;
        }
        if let Some(value) = style.next_style {
            self.write_control_word("snext", Some(i32::from(value)))?;
        }
        if let Some(value) = style.linked_style {
            self.write_control_word("slink", Some(i32::from(value)))?;
        }
        if style.additive {
            self.write_control_word("additive", None)?;
        }
        if style.auto_update {
            self.write_control_word("sautoupd", None)?;
        }
        if style.hidden {
            self.write_control_word("shidden", None)?;
        }
        if style.locked {
            self.write_control_word("slocked", None)?;
        }
        if style.semi_hidden {
            self.write_control_word("ssemihidden", None)?;
        }
        if style.unhide_when_used {
            self.write_control_word("sunhideused", None)?;
        }
        if style.quick_format {
            self.write_control_word("sqformat", None)?;
        }
        if let Some(value) = style.priority {
            self.write_control_word("spriority", Some(value))?;
        }
        if let Some(value) = style.revision_id {
            self.write_control_word("styrsid", Some(value))?;
        }
        if style.personal {
            self.write_control_word("spersonal", None)?;
        }
        if style.compose {
            self.write_control_word("scompose", None)?;
        }
        if style.reply {
            self.write_control_word("sreply", None)?;
        }
        if style.style_type == StyleType::Table {
            let conditional = &style.table_conditional;
            for (flag, word) in [
                (conditional.first_row, "tscfirstrow"),
                (conditional.last_row, "tsclastrow"),
                (conditional.first_column, "tscfirstcol"),
                (conditional.last_column, "tsclastcol"),
                (conditional.band_horizontal_odd, "tscbandhorzodd"),
                (conditional.band_horizontal_even, "tscbandhorzeven"),
                (conditional.band_vertical_odd, "tscbandvertodd"),
                (conditional.band_vertical_even, "tscbandverteven"),
            ] {
                if flag {
                    self.write_control_word(word, None)?;
                }
            }
            if let Some(size) = conditional.horizontal_band_size {
                self.write_control_word("tscbandsh", Some(i32::from(size)))?;
            }
            if let Some(size) = conditional.vertical_band_size {
                self.write_control_word("tscbandsv", Some(i32::from(size)))?;
            }
        }
        self.write_str(" ")?;
        self.write_text(style.name.as_ref())?;
        self.write_str(";}")?;
        Ok(())
    }
}
