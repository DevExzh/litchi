//! RTF metadata and destination output.

use super::super::*;

impl<W: Write> RtfWriter<W> {
    /// Write the standard RTF document-information destination.
    pub fn write_document_info(&mut self, info: &DocumentInfo<'_>) -> io::Result<()> {
        let has_info = info.title.is_some()
            || info.subject.is_some()
            || info.author.is_some()
            || info.manager.is_some()
            || info.company.is_some()
            || info.operator.is_some()
            || info.category.is_some()
            || info.keywords.is_some()
            || info.comment.is_some()
            || info.document_comment.is_some()
            || info.hyperlink_base.is_some()
            || info.version.is_some()
            || info.revision.is_some()
            || info.creation_time.is_some()
            || info.revision_time.is_some()
            || info.print_time.is_some()
            || info.backup_time.is_some()
            || info.creation_timestamp.is_some()
            || info.revision_timestamp.is_some()
            || info.print_timestamp.is_some()
            || info.backup_timestamp.is_some()
            || info.editing_time.is_some()
            || info.pages.is_some()
            || info.words.is_some()
            || info.characters.is_some()
            || info.characters_with_spaces.is_some()
            || info.id.is_some()
            || info.protection.password_hash.is_some();
        info.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        if has_info {
            self.write_str("{")?;
            self.write_control_word("info", None)?;
            self.write_info_text("title", info.title.as_deref())?;
            self.write_info_text("subject", info.subject.as_deref())?;
            self.write_info_text("author", info.author.as_deref())?;
            self.write_info_text("manager", info.manager.as_deref())?;
            self.write_info_text("company", info.company.as_deref())?;
            self.write_info_text("operator", info.operator.as_deref())?;
            self.write_info_text("category", info.category.as_deref())?;
            self.write_info_text("keywords", info.keywords.as_deref())?;
            self.write_info_text("comment", info.comment.as_deref())?;
            self.write_info_text("doccomm", info.document_comment.as_deref())?;
            self.write_info_text("hlinkbase", info.hyperlink_base.as_deref())?;
            self.write_info_time(
                "creatim",
                info.creation_timestamp,
                info.creation_time.as_deref(),
            )?;
            self.write_info_time(
                "revtim",
                info.revision_timestamp,
                info.revision_time.as_deref(),
            )?;
            self.write_info_time("printim", info.print_timestamp, info.print_time.as_deref())?;
            self.write_info_time("buptim", info.backup_timestamp, info.backup_time.as_deref())?;
            self.write_optional_u32("version", info.version)?;
            self.write_optional_u32("vern", info.revision)?;
            self.write_optional_u32("edmins", info.editing_time)?;
            self.write_optional_u32("nofpages", info.pages)?;
            self.write_optional_u32("nofwords", info.words)?;
            self.write_optional_u32("nofchars", info.characters)?;
            self.write_optional_u32("nofcharsws", info.characters_with_spaces)?;
            self.write_optional_u32("id", info.id)?;
            if let Some(hash) = info.protection.password_hash.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("password", None)?;
                self.write_str(" ")?;
                self.write_str(hash)?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        self.write_protection_controls(&info.protection)
    }

    pub(in super::super) fn write_protection_controls(
        &mut self,
        protection: &crate::DocumentProtection<'_>,
    ) -> io::Result<()> {
        for (control, value) in [
            ("formprot", protection.forms),
            ("annotprot", protection.annotations),
            ("revprot", protection.revisions),
            ("readprot", protection.read_only),
            ("allprot", protection.all),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, (!value).then_some(0))?;
            }
        }
        if let Some(value) = protection.enforced {
            self.write_control_word("enforceprot", Some(i32::from(value)))?;
        }
        if let Some(level) = protection.level {
            self.write_control_word("protlevel", Some(level.rtf_value()))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_info_text(
        &mut self,
        control: &str,
        value: Option<&str>,
    ) -> io::Result<()> {
        let Some(value) = value else { return Ok(()) };
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    pub(in super::super) fn write_info_time(
        &mut self,
        control: &str,
        typed: Option<RtfTimestamp>,
        legacy: Option<&str>,
    ) -> io::Result<()> {
        let timestamp = match (typed, legacy) {
            (Some(value), _) => value,
            (None, Some(value)) => RtfTimestamp::from_legacy(value)
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?,
            (None, None) => return Ok(()),
        };
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        for (name, value) in [
            ("yr", timestamp.year),
            ("mo", timestamp.month),
            ("dy", timestamp.day),
            ("hr", timestamp.hour),
            ("min", timestamp.minute),
            ("sec", timestamp.second),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, Some(value))?;
            }
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_optional_u32(
        &mut self,
        control: &str,
        value: Option<u32>,
    ) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value as i32))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_optional_i32(
        &mut self,
        control: &str,
        value: Option<i32>,
    ) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value))?;
        }
        Ok(())
    }

    /// Write the canonical starred RTF user-properties destination.
    pub fn write_user_properties(
        &mut self,
        properties: &[crate::UserProperty<'_>],
    ) -> io::Result<()> {
        if properties.is_empty() {
            return Ok(());
        }
        if properties.len() > crate::user_property::MAX_USER_PROPERTIES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF user-property count limit exceeded",
            ));
        }
        let mut names = std::collections::HashSet::with_capacity(properties.len());
        let mut aggregate = 0usize;
        for property in properties {
            property
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if !names.insert(property.name.as_ref()) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "duplicate RTF user-property name",
                ));
            }
            aggregate = aggregate
                .checked_add(property.text_bytes().ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "user-property size overflow")
                })?)
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "user-property size overflow")
                })?;
            if aggregate > crate::user_property::MAX_USER_PROPERTY_TEXT_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF user-property aggregate text limit exceeded",
                ));
            }
        }

        self.write_str("{\\*")?;
        self.write_control_word("userprops", None)?;
        for property in properties {
            self.write_str("{")?;
            self.write_control_word("propname", None)?;
            self.write_str(" ")?;
            self.write_destination_text(property.name.as_ref())?;
            self.write_str("}")?;
            self.write_control_word("proptype", Some(property.value.type_code()))?;
            self.write_str("{")?;
            self.write_control_word("staticval", None)?;
            self.write_str(" ")?;
            self.write_destination_text(property.value.lexical())?;
            self.write_str("}")?;
            if let Some(link) = &property.link_value {
                self.write_str("{")?;
                self.write_control_word("linkval", None)?;
                self.write_str(" ")?;
                self.write_destination_text(link.as_ref())?;
                self.write_str("}")?;
            }
        }
        self.write_str("}")
    }

    /// Write ordered standard RTF document-variable destinations.
    pub fn write_document_variables(
        &mut self,
        variables: &[crate::DocumentVariable<'_>],
    ) -> io::Result<()> {
        if variables.len() > crate::document_variable::MAX_DOCUMENT_VARIABLES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF document-variable count limit exceeded",
            ));
        }
        let mut aggregate = 0usize;
        for variable in variables {
            variable
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            aggregate = aggregate
                .checked_add(variable.name.len())
                .and_then(|size| size.checked_add(variable.value.len()))
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "document-variable size overflow",
                    )
                })?;
            if aggregate > crate::document_variable::MAX_DOCUMENT_VARIABLE_TEXT_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF document-variable aggregate text limit exceeded",
                ));
            }
            self.write_str("{\\*")?;
            self.write_control_word("docvar", None)?;
            self.write_str(" {")?;
            self.write_destination_text(variable.name.as_ref())?;
            self.write_str("}{")?;
            self.write_destination_text(variable.value.as_ref())?;
            self.write_str("}}")?;
        }
        Ok(())
    }

    pub(in super::super) fn write_destination_text(&mut self, text: &str) -> io::Result<()> {
        for character in text.chars() {
            match character {
                '\\' => self.write_str("\\\\")?,
                '{' => self.write_str("\\{")?,
                '}' => self.write_str("\\}")?,
                character if character.is_ascii_control() => {
                    write!(self.writer, "\\'{:02x}", character as u8)?;
                },
                character if character.is_ascii() => write!(self.writer, "{character}")?,
                character => {
                    for unit in character.encode_utf16(&mut [0; 2]).iter().copied() {
                        self.write_control_word("u", Some(i32::from(unit as i16)))?;
                        self.write_str("?")?;
                    }
                },
            }
        }
        Ok(())
    }

    /// Write a bookmark start destination.
    pub fn write_bookmark_start(&mut self, bookmark: &Bookmark<'_>) -> io::Result<()> {
        if bookmark.name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark name cannot be empty",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("bkmkstart", None)?;
        self.write_optional_i32("bkmkcolf", bookmark.first_column)?;
        self.write_optional_i32("bkmkcoll", bookmark.last_column)?;
        if bookmark.is_public {
            self.write_control_word("bkmkpub", None)?;
        }
        self.write_str(" ")?;
        self.write_text(bookmark.name.as_ref())?;
        self.write_str("}")
    }

    /// Write a bookmark end destination.
    pub fn write_bookmark_end(&mut self, name: &str) -> io::Result<()> {
        if name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark name cannot be empty",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("bkmkend", None)?;
        self.write_str(" ")?;
        self.write_text(name)?;
        self.write_str("}")
    }

    /// Write a custom XML tag open destination and its inert attributes.
    pub fn write_custom_xml_open(&mut self, tag: &crate::CustomXmlTag<'_>) -> io::Result<()> {
        tag.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word("xmlopen", None)?;
        if let Some(namespace) = tag.namespace {
            self.write_control_word("xmlns", Some(namespace as i32))?;
        }
        self.write_str(" ")?;
        self.write_destination_text(tag.name.as_ref())?;
        self.write_str("}")?;
        for attribute in &tag.attributes {
            self.write_str("{\\*")?;
            self.write_control_word("xmlattrname", None)?;
            self.write_str(" ")?;
            self.write_destination_text(attribute.name.as_ref())?;
            self.write_str("}{\\*")?;
            self.write_control_word("xmlattrvalue", None)?;
            self.write_str(" ")?;
            self.write_destination_text(attribute.value.as_ref())?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write a custom XML tag close destination.
    pub fn write_custom_xml_close(&mut self, tag: &crate::CustomXmlTag<'_>) -> io::Result<()> {
        if tag.name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF custom XML tag name cannot be empty",
            ));
        }
        self.write_str("{")?;
        self.write_control_word("xmlclose", None)?;
        self.write_str(" ")?;
        self.write_destination_text(tag.name.as_ref())?;
        self.write_str("}")
    }

    /// Write a protection-exception range marker destination.
    pub fn write_protection_range_marker(
        &mut self,
        control: &str,
        range: &crate::ProtectionRange<'_>,
    ) -> io::Result<()> {
        range
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(range.id.as_ref())?;
        self.write_str("}")
    }

    pub(in super::super) fn math_structure_control(kind: crate::MathStructureKind) -> &'static str {
        use crate::MathStructureKind as K;
        match kind {
            K::Accent => "macc",
            K::Bar => "mbar",
            K::BorderBox => "mborderBox",
            K::Box => "mbox",
            K::Delimiter => "md",
            K::EquationArray => "meqArr",
            K::Fraction => "mf",
            K::Function => "mfunc",
            K::GroupChar => "mgroupChr",
            K::LimitLower => "mlimlow",
            K::LimitUpper => "mlimupp",
            K::Matrix => "mm",
            K::Nary => "mnary",
            K::Phantom => "mphant",
            K::Radical => "mrad",
            K::ScriptPre => "msPre",
            K::ScriptSub => "msSub",
            K::ScriptSubSup => "msSubSup",
            K::ScriptSup => "msSup",
        }
    }

    pub(in super::super) fn math_structure_properties_control(
        kind: crate::MathStructureKind,
    ) -> &'static str {
        use crate::MathStructureKind as K;
        match kind {
            K::Accent => "maccPr",
            K::Bar => "mbarPr",
            K::BorderBox => "mborderBoxPr",
            K::Box => "mboxPr",
            K::Delimiter => "mdPr",
            K::EquationArray => "meqArrPr",
            K::Fraction => "mfPr",
            K::Function => "mfuncPr",
            K::GroupChar => "mgroupChrPr",
            K::LimitLower => "mlimlowPr",
            K::LimitUpper => "mlimuppPr",
            K::Matrix => "mmPr",
            K::Nary => "mnaryPr",
            K::Phantom => "mphantPr",
            K::Radical => "mradPr",
            K::ScriptPre => "msPrePr",
            K::ScriptSub => "msSubPr",
            K::ScriptSubSup => "msSubSupPr",
            K::ScriptSup => "msSupPr",
        }
    }

    pub(in super::super) fn math_element_control(role: crate::MathElementRole) -> &'static str {
        use crate::MathElementRole as R;
        match role {
            R::Element => "me",
            R::Numerator => "mnum",
            R::Denominator => "mden",
            R::Degree => "mdeg",
            R::Subscript => "msub",
            R::Superscript => "msup",
            R::Limit => "mlim",
            R::FunctionName => "mfName",
        }
    }

    pub(in super::super) fn math_property_control(name: crate::MathPropertyName) -> &'static str {
        use crate::MathPropertyName as N;
        match name {
            N::Type => "mtype",
            N::Grow => "mgrow",
            N::Char => "mchr",
            N::BeginChar => "mbegChr",
            N::EndChar => "mendChr",
            N::SeparatorChar => "msepChr",
            N::Position => "mpos",
            N::VerticalJustify => "mvertJc",
            N::BaseJustify => "mbaseJc",
            N::Justify => "mjc",
            N::Align => "maln",
            N::AlignScript => "malnScr",
            N::DegreeHide => "mdegHide",
            N::Differential => "mdiff",
            N::DifferentialStyle => "mdiffSty",
            N::HideBottom => "mhideBot",
            N::HideLeft => "mhideLeft",
            N::HideRight => "mhideRight",
            N::HideTop => "mhideTop",
            N::LimitLocation => "mlimLoc",
            N::PlaceholderHide => "mplcHide",
            N::SubscriptHide => "msubHide",
            N::SuperscriptHide => "msupHide",
            N::StrikeBottomLeftToTopRight => "mstrikeBLTR",
            N::StrikeHorizontal => "mstrikeH",
            N::StrikeTopLeftToBottomRight => "mstrikeTLBR",
            N::StrikeVertical => "mstrikeV",
            N::Style => "msty",
            N::Script => "mscr",
            N::Transparent => "mtransp",
            N::Show => "mshow",
            N::Shape => "mshp",
            N::ZeroAscent => "mzeroAsc",
            N::ZeroDescent => "mzeroDesc",
            N::ZeroWidth => "mzeroWid",
            N::OperatorEmulator => "mopEmu",
            N::NoBreak => "mnoBreak",
            N::NormalText => "mnor",
            N::Literal => "mlit",
            N::MatrixColumnGap => "mcGp",
            N::MatrixColumnGapRule => "mcGpRule",
            N::MatrixColumnSpacing => "mcSp",
            N::MatrixCellCount => "mcount",
            N::MatrixCellJustify => "mmcJc",
            N::RowSpacing => "mrSp",
            N::RowSpacingRule => "mrSpRule",
            N::Break => "mbrk",
            N::ArgumentSize => "margSz",
        }
    }

    /// Write an inert math zone destination.
    pub fn write_math_zone(&mut self, zone: &crate::MathZone<'_>) -> io::Result<()> {
        zone.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(
            match zone.kind {
                crate::MathZoneKind::Inline => "mmath",
                crate::MathZoneKind::Display => "mmathPara",
            },
            None,
        )?;
        if let Some(properties) = &zone.paragraph_properties {
            self.write_math_properties_group("mmathParaPr", properties)?;
        }
        for object in &zone.content {
            self.write_math_object(object)?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_math_object(
        &mut self,
        object: &crate::MathObject<'_>,
    ) -> io::Result<()> {
        match object {
            crate::MathObject::Structure(structure) => self.write_math_structure(structure),
            crate::MathObject::Run(run) => self.write_math_run(run),
        }
    }

    pub(in super::super) fn write_math_structure(
        &mut self,
        structure: &crate::MathStructure<'_>,
    ) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word(Self::math_structure_control(structure.kind), None)?;
        if let Some(properties) = &structure.properties {
            self.write_math_properties_group(
                Self::math_structure_properties_control(structure.kind),
                properties,
            )?;
        }
        for child in &structure.children {
            match child {
                crate::MathStructureChild::Element(element) => self.write_math_element(element)?,
                crate::MathStructureChild::MatrixRow(row) => {
                    self.write_str("{")?;
                    self.write_control_word("mmr", None)?;
                    for cell in &row.cells {
                        self.write_math_element(cell)?;
                    }
                    self.write_str("}")?;
                },
            }
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_math_element(
        &mut self,
        element: &crate::MathElement<'_>,
    ) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word(Self::math_element_control(element.role), None)?;
        if let Some(properties) = &element.argument_properties {
            self.write_math_properties_group("margPr", properties)?;
        }
        self.write_str(" ")?;
        for object in &element.content {
            self.write_math_object(object)?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_math_run(&mut self, run: &crate::MathRun<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("mr", None)?;
        if let Some(properties) = &run.properties {
            self.write_math_properties_group("mrPr", properties)?;
        }
        if run.normal_text {
            self.write_control_word("mnor", None)?;
        }
        self.write_str(" ")?;
        self.write_destination_text(run.text.as_ref())?;
        self.write_str("}")
    }

    pub(in super::super) fn write_math_properties_group(
        &mut self,
        destination: &str,
        properties: &crate::MathProperties<'_>,
    ) -> io::Result<()> {
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(destination, None)?;
        for property in &properties.properties {
            self.write_str("{")?;
            self.write_control_word(Self::math_property_control(property.name), None)?;
            if !property.value.is_empty() {
                self.write_str(" ")?;
                self.write_destination_text(property.value.as_ref())?;
            }
            self.write_str("}")?;
        }
        if !properties.matrix_columns.is_empty() {
            self.write_str("{")?;
            self.write_control_word("mmcs", None)?;
            for column in &properties.matrix_columns {
                self.write_str("{")?;
                self.write_control_word("mmc", None)?;
                if let Some(column_properties) = &column.properties {
                    self.write_math_properties_group("mmcPr", column_properties)?;
                }
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        if let Some(control) = &properties.control {
            self.write_math_properties_group("mctrlPr", control)?;
        }
        self.write_str("}")
    }

    /// Write an annotation range-start destination.
    pub fn write_annotation_start(&mut self, annotation: &Annotation<'_>) -> io::Result<()> {
        if !annotation.has_reference {
            return Ok(());
        }
        self.write_str("{\\*")?;
        self.write_control_word("atrfstart", None)?;
        self.write_str(" ")?;
        write!(self.writer, "{}", annotation.id)?;
        self.write_str("}")
    }

    /// Write an annotation range end, author metadata, and inert comment body.
    pub fn write_annotation_end(&mut self, annotation: &Annotation<'_>) -> io::Result<()> {
        if annotation.annotation_type != AnnotationType::Comment {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "only comment annotations use the RTF annotation destination",
            ));
        }
        annotation
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        if annotation.has_reference {
            self.write_str("{\\*")?;
            self.write_control_word("atrfend", None)?;
            self.write_str(" ")?;
            write!(self.writer, "{}", annotation.id)?;
            self.write_str("}")?;
        }
        self.write_annotation_value("atnid", Some(annotation.initials.as_ref()))?;
        self.write_annotation_value("atnauthor", Some(annotation.author.as_ref()))?;
        self.write_control_word("chatn", None)?;
        self.write_str("{\\*")?;
        self.write_control_word("annotation", None)?;
        self.write_str(" ")?;
        let reference = annotation.has_reference.then(|| annotation.id.to_string());
        self.write_annotation_value("atnref", reference.as_deref())?;
        self.write_annotation_value("atndate", annotation.date.as_deref())?;
        self.write_annotation_value("atnparent", annotation.parent_id.as_deref())?;
        self.write_annotation_value("atnicn", annotation.icon.as_deref())?;
        self.write_annotation_value("atntime", annotation.time.as_deref())?;
        self.write_field_story(
            annotation.text.as_ref(),
            &annotation.shapes,
            &annotation.shape_groups,
            &annotation.drawing_order,
            &annotation.story_events,
            &[],
            crate::FieldOwner::Other,
            DrawingStoryTextMode::Destination,
            0,
        )?;
        self.write_str("}")
    }
}
