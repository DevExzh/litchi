//! RTF document writer/serializer.
//!
//! This module provides functionality to write RTF documents from structured data.
//! It supports all RTF features including formatting, tables, pictures, fields, lists, and more.

use super::*;
use std::io::{self, Write};

/// RTF writer options
#[derive(Debug, Clone)]
pub struct WriterOptions {
    /// Use ANSI code page
    pub use_ansi: bool,
    /// ANSI code page number (default 1252 for Western European)
    pub code_page: u16,
    /// Indent RTF output for readability
    pub indent: bool,
    /// Default font index
    pub default_font: u16,
    /// Default tab width (in twips)
    pub default_tab_width: i32,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self {
            use_ansi: true,
            code_page: 1252,
            indent: false,
            default_font: 0,
            default_tab_width: 720, // 0.5 inch
        }
    }
}

/// RTF document writer
///
/// Provides functionality to serialize RTF documents to a writer.
/// All fields are used internally during the writing process.
pub struct RtfWriter<W: Write> {
    /// Output writer
    writer: W,
    /// Writer options
    options: WriterOptions,
    /// Current indentation level (reserved for formatted output)
    #[allow(dead_code)]
    indent_level: usize,
    /// Font table
    font_table: FontTable<'static>,
    /// Color table
    color_table: ColorTable,
}

#[derive(Clone, Copy)]
enum BodyEventKind<'b, 'a> {
    NavigationEntry(&'b crate::NavigationEntry<'a>),
    BookmarkStart(&'b Bookmark<'a>),
    BookmarkEnd(&'b Bookmark<'a>),
    AnnotationStart(&'b Annotation<'a>),
    AnnotationEnd(&'b Annotation<'a>),
    RevisionStart(&'b Revision<'a>),
    RevisionEnd,
    RevisionDeletion(&'b Revision<'a>),
    FormFieldStart(&'b crate::FormField<'a>),
    FormFieldEnd,
}

#[derive(Clone, Copy)]
struct BodyEvent<'b, 'a> {
    offset: usize,
    order: u8,
    kind: BodyEventKind<'b, 'a>,
}

impl<W: Write> RtfWriter<W> {
    /// Create a new RTF writer
    pub fn new(writer: W) -> Self {
        Self::with_options(writer, WriterOptions::default())
    }

    /// Create a new RTF writer with options
    pub fn with_options(writer: W, options: WriterOptions) -> Self {
        Self {
            writer,
            options,
            indent_level: 0,
            font_table: FontTable::new(),
            color_table: ColorTable::new(),
        }
    }

    /// Write a complete RTF document
    pub fn write_document<'a>(&mut self, doc: &RtfDocument<'a>) -> io::Result<()> {
        if doc.sections().len() > 1 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF document model does not retain body boundaries for multiple sections",
            ));
        }
        // Collect font and color tables from document by cloning them
        // We need to convert the lifetime to 'static for storage
        let font_table: FontTable<'static> = FontTable {
            fonts: doc
                .font_table()
                .fonts()
                .iter()
                .map(|f| Font {
                    name: std::borrow::Cow::Owned(f.name.to_string()),
                    family: f.family,
                    charset: f.charset,
                })
                .collect(),
        };
        let color_table = doc.color_table().clone();

        self.font_table = font_table;
        self.color_table = color_table;

        // Write document header
        self.write_document_header()?;

        self.write_language_defaults(doc.language_defaults())?;

        // Write font table
        self.write_font_table()?;

        // Write color table
        self.write_color_table()?;

        // Write named paragraph, character, section, and table styles.
        self.write_stylesheet(doc.stylesheet())?;

        // Write list definitions before body paragraphs reference them.
        self.write_list_table(doc.list_table())?;
        self.write_list_override_table(doc.list_override_table())?;

        // Revision controls reference this author table by numeric index.
        self.write_revision_table(doc.revision_authors(), doc.revisions())?;

        self.write_revision_save_metadata(doc.revision_save_metadata())?;

        self.write_xml_namespace_table(doc.xml_namespaces())?;

        self.write_theme(doc.theme())?;

        self.write_latent_styles(doc.latent_styles())?;

        self.write_data_store(doc.data_store())?;

        self.write_math_properties(doc.math_properties())?;

        // Producer provenance is inert header metadata.
        self.write_generator(doc.generator())?;

        // Write document properties before body content.
        self.write_document_info(doc.info())?;

        // User-defined properties are header-level inert metadata.
        self.write_user_properties(doc.user_properties())?;

        // Document variables are header-level inert metadata.
        self.write_document_variables(doc.document_variables())?;

        // Headers and footers belong to the section definition before body text.
        for section in doc.sections() {
            self.write_section(section)?;
        }

        // Write document content and reinsert positional bookmark/comment markers.
        self.write_blocks_with_markup(
            doc.blocks(),
            doc.bookmarks(),
            doc.annotations(),
            doc.revisions(),
            doc.navigation_entries(),
            doc.form_fields(),
        )?;

        // Write tables
        for table in doc.tables() {
            self.write_table(table)?;
        }

        // Close document
        self.write_str("}")?;

        Ok(())
    }

    /// Write document header
    pub fn write_document_header(&mut self) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("rtf", Some(1))?;

        if self.options.use_ansi {
            self.write_control_word("ansi", None)?;
            self.write_control_word("ansicpg", Some(self.options.code_page as i32))?;
        }

        self.write_control_word("deff", Some(self.options.default_font as i32))?;
        self.write_control_word("deftab", Some(self.options.default_tab_width))?;

        Ok(())
    }

    pub fn write_language_defaults(
        &mut self,
        defaults: &crate::DocumentLanguageDefaults,
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

    /// Write font table
    fn write_font_table(&mut self) -> io::Result<()> {
        if self.font_table.fonts().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("fonttbl", None)?;

        // Clone fonts to avoid borrowing issues
        let fonts: Vec<_> = self.font_table.fonts().to_vec();
        for (idx, font) in fonts.iter().enumerate() {
            self.write_str("{")?;
            self.write_control_word("f", Some(idx as i32))?;

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

            // Write charset
            if font.charset != 0 {
                self.write_control_word("fcharset", Some(font.charset as i32))?;
            }

            // Write font name
            self.write_str(" ")?;
            self.write_text(font.name.as_ref())?;
            self.write_str(";")?;
            self.write_str("}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write color table
    fn write_color_table(&mut self) -> io::Result<()> {
        if self.color_table.colors().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("colortbl", None)?;

        // Clone colors to avoid borrowing issues
        let colors: Vec<_> = self.color_table.colors().to_vec();
        for color in &colors {
            self.write_control_word("red", Some(color.red as i32))?;
            self.write_control_word("green", Some(color.green as i32))?;
            self.write_control_word("blue", Some(color.blue as i32))?;
            self.write_str(";")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write the list-definition table.
    pub fn write_list_table(&mut self, table: &ListTable<'_>) -> io::Result<()> {
        if table.lists().is_empty() {
            return Ok(());
        }
        if table.lists().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list table exceeds the supported list count",
            ));
        }
        self.write_str("{\\*\\listtable")?;
        for list in table.lists() {
            self.write_list_definition(list)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn write_list_definition(&mut self, list: &List<'_>) -> io::Result<()> {
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
        self.write_control_word("listid", Some(list.id))?;
        self.write_str("}")?;
        Ok(())
    }

    fn write_list_level(&mut self, level: &ListLevel<'_>) -> io::Result<()> {
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
        self.write_list_level_text(level.number_text.as_ref())?;
        if level.font_ref != 0 {
            self.write_control_word("f", Some(i32::from(level.font_ref)))?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn list_level_type_value(level_type: ListLevelType) -> i32 {
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

    fn write_list_level_text(&mut self, text: &str) -> io::Result<()> {
        let count = u8::try_from(text.chars().count()).map_err(|_| {
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
        for (index, ch) in text.chars().enumerate() {
            if u32::from(ch) <= 8 {
                let position = u8::try_from(index + 1).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "invalid RTF list placeholder")
                })?;
                self.write_hex_byte(position)?;
            }
        }
        self.write_str(";}")?;
        Ok(())
    }

    fn write_hex_byte(&mut self, value: u8) -> io::Result<()> {
        write!(self.writer, "\\'{value:02x}")
    }

    /// Write the list-override table.
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
                    u8::from(entry.start_at_override.is_some())
                }))),
            )?;
            if let Some(start_at) = entry.start_at_override {
                self.write_str("{")?;
                self.write_control_word("lfolevel", None)?;
                self.write_control_word("listoverridestartat", None)?;
                self.write_control_word("levelstartat", Some(start_at))?;
                self.write_str("}")?;
            }
            self.write_control_word("ls", Some(entry.index))?;
            self.write_str("}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write the revision-author table referenced by tracked-change runs.
    pub fn write_revision_table(
        &mut self,
        authors: &[crate::RevisionAuthor<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        if authors.is_empty() && revisions.is_empty() {
            return Ok(());
        }
        if authors.len() > crate::annotation::MAX_REVISION_AUTHORS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF revision-author table exceeds the safety limit",
            ));
        }
        let author_bytes = authors.iter().try_fold(0usize, |total, author| {
            author.validate().map_err(|error| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    error.to_string(),
                )
            })?;
            total.checked_add(author.name.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF aggregate revision-author size overflow",
                )
            })
        })?;
        if author_bytes > crate::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF aggregate revision-author text exceeds the safety limit",
            ));
        }
        for revision in revisions {
            revision.validate().map_err(|error| {
                io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
            })?;
            let index = usize::try_from(revision.id).map_err(|_| {
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

    pub fn write_generator(
        &mut self,
        generator: Option<&crate::DocumentGenerator<'_>>,
    ) -> io::Result<()> {
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

    pub fn write_revision_save_metadata(
        &mut self,
        metadata: Option<&crate::RevisionSaveMetadata>,
    ) -> io::Result<()> {
        let Some(metadata) = metadata else {
            return Ok(());
        };
        metadata
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\rsidtbl ")?;
        for id in metadata.ids() {
            self.write_control_word("rsid", Some(*id as i32))?;
        }
        self.write_str("}")?;
        if let Some(root) = metadata.root() {
            self.write_control_word("rsidroot", Some(root as i32))?;
        }
        Ok(())
    }

    pub fn write_xml_namespace_table(
        &mut self,
        namespaces: Option<&[crate::XmlNamespace<'_>]>,
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
        self.write_str("{\\*\\xmlnstbl ")?;
        for (index, namespace) in namespaces.iter().enumerate() {
            namespace
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if namespaces[..index]
                .iter()
                .any(|existing| existing.id == namespace.id)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF XML namespace IDs must be unique",
                ));
            }
            total = total.checked_add(namespace.namespace.len()).ok_or_else(|| {
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
            self.write_control_word("xmlns", Some(namespace.id as i32))?;
            self.write_str(" ")?;
            self.write_destination_text(namespace.namespace.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    pub fn write_theme(&mut self, theme: Option<&crate::DocumentTheme<'_>>) -> io::Result<()> {
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

    fn write_hex_destination(&mut self, control: &str, data: &[u8]) -> io::Result<()> {
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        for byte in data {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")
    }

    pub fn write_latent_styles(
        &mut self,
        styles: Option<&crate::LatentStyles<'_>>,
    ) -> io::Result<()> {
        let Some(styles) = styles else {
            return Ok(());
        };
        styles
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\latentstyles")?;
        self.write_control_word("lsdstimax", Some(styles.max_style_index as i32))?;
        self.write_optional_bool("lsdlockeddef", styles.locked_default)?;
        self.write_optional_bool("lsdsemihiddendef", styles.semi_hidden_default)?;
        self.write_optional_bool(
            "lsdunhideuseddef",
            styles.unhide_when_used_default,
        )?;
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

    pub fn write_data_store(
        &mut self,
        data_store: Option<&crate::DocumentDataStore<'_>>,
    ) -> io::Result<()> {
        let Some(data_store) = data_store else {
            return Ok(());
        };
        data_store
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_hex_destination("datastore", data_store.data.as_ref())
    }

    pub fn write_math_properties(
        &mut self,
        properties: Option<&crate::DocumentMathProperties>,
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
            self.write_control_word("mmathFont", Some(value as i32))?;
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

    fn write_optional_bool(&mut self, control: &str, value: Option<bool>) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write an RTF stylesheet destination.
    pub fn write_stylesheet(&mut self, stylesheet: &StyleSheet<'_>) -> io::Result<()> {
        if stylesheet.styles().is_empty() {
            return Ok(());
        }
        if stylesheet.styles().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF stylesheet exceeds the supported style count",
            ));
        }

        self.write_str("{")?;
        self.write_control_word("stylesheet", None)?;
        for style in stylesheet.styles() {
            self.write_style_definition(style)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn write_style_definition(&mut self, style: &Style<'_>) -> io::Result<()> {
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
        self.write_str(" ")?;
        self.write_text(style.name.as_ref())?;
        self.write_str(";}")?;
        Ok(())
    }

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
            || info.version.is_some()
            || info.revision.is_some()
            || info.creation_time.is_some()
            || info.revision_time.is_some()
            || info.print_time.is_some()
            || info.backup_time.is_some()
            || info.editing_time.is_some()
            || info.pages.is_some()
            || info.words.is_some()
            || info.characters.is_some()
            || info.characters_with_spaces.is_some()
            || info.id.is_some();
        if !has_info {
            return Ok(());
        }

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
        self.write_info_time("creatim", info.creation_time.as_deref())?;
        self.write_info_time("revtim", info.revision_time.as_deref())?;
        self.write_info_time("printim", info.print_time.as_deref())?;
        self.write_info_time("buptim", info.backup_time.as_deref())?;
        self.write_optional_i32("version", info.version)?;
        self.write_optional_i32("vern", info.revision)?;
        self.write_optional_i32("edmins", info.editing_time)?;
        self.write_optional_i32("nofpages", info.pages)?;
        self.write_optional_i32("nofwords", info.words)?;
        self.write_optional_i32("nofchars", info.characters)?;
        self.write_optional_i32("nofcharsws", info.characters_with_spaces)?;
        self.write_optional_i32("id", info.id)?;
        self.write_str("}")
    }

    fn write_info_text(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else { return Ok(()) };
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_text(value)?;
        self.write_str("}")
    }

    fn write_info_time(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else { return Ok(()) };
        let (date, time) = value.split_once('T').ok_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidInput, "RTF info time must contain T")
        })?;
        let date: Vec<i32> = date
            .split('-')
            .map(|part| part.parse())
            .collect::<Result<_, _>>()
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "invalid RTF info date"))?;
        let time: Vec<i32> = time
            .split(':')
            .map(|part| part.parse())
            .collect::<Result<_, _>>()
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "invalid RTF info time"))?;
        if date.len() != 3 || time.len() != 3 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF info time must use YYYY-MM-DDTHH:MM:SS",
            ));
        }
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        for (name, value) in [
            ("yr", date[0]),
            ("mo", date[1]),
            ("dy", date[2]),
            ("hr", time[0]),
            ("min", time[1]),
            ("sec", time[2]),
        ] {
            self.write_control_word(name, Some(value))?;
        }
        self.write_str("}")
    }

    fn write_optional_i32(&mut self, control: &str, value: Option<i32>) -> io::Result<()> {
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
                .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "document-variable size overflow"))?;
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

    fn write_destination_text(&mut self, text: &str) -> io::Result<()> {
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
        self.write_destination_text(annotation.text.as_ref())?;
        self.write_str("}")
    }

    fn write_annotation_value(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    fn write_blocks_with_markup(
        &mut self,
        blocks: &[StyleBlock<'_>],
        bookmarks: &BookmarkTable<'_>,
        annotations: &[Annotation<'_>],
        revisions: &[Revision<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        form_fields: &[crate::FormField<'_>],
    ) -> io::Result<()> {
        if bookmarks.bookmarks().is_empty()
            && annotations.is_empty()
            && revisions.is_empty()
            && navigation_entries.is_empty()
            && form_fields.is_empty()
        {
            for block in blocks {
                self.write_style_block(block)?;
            }
            return Ok(());
        }

        let body: String = blocks.iter().map(|block| block.text.as_ref()).collect();
        let event_count = bookmarks
            .bookmarks()
            .len()
            .saturating_add(annotations.len())
            .saturating_add(revisions.len())
            .saturating_mul(2);
        let event_count = event_count.saturating_add(navigation_entries.len());
        let event_count = event_count.saturating_add(form_fields.len().saturating_mul(2));
        let mut events = Vec::with_capacity(event_count);
        if form_fields.len() > crate::form_field::MAX_FORM_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF form-field count exceeds the safety limit",
            ));
        }
        let mut form_field_bytes = 0usize;
        let mut form_field_ranges: Vec<&crate::FormField<'_>> = form_fields.iter().collect();
        form_field_ranges.sort_by_key(|field| (field.position, field.range_end));
        let mut previous_form_end = 0usize;
        for field in form_field_ranges {
            field
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let result = body.get(field.position..field.range_end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field range is outside body text or splits a character",
                )
            })?;
            if result != field.result_text {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result does not match its visible body range",
                ));
            }
            if field.position != field.range_end && field.position < previous_form_end {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result ranges cannot overlap",
                ));
            }
            if field.position != field.range_end {
                previous_form_end = field.range_end;
            }
            form_field_bytes = form_field_bytes
                .checked_add(field.text_bytes().ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?;
            if form_field_bytes > crate::form_field::MAX_FORM_FIELD_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field aggregate text exceeds the safety limit",
                ));
            }
            let empty = field.position == field.range_end;
            events.push(BodyEvent {
                offset: field.position,
                order: 1,
                kind: BodyEventKind::FormFieldStart(field),
            });
            events.push(BodyEvent {
                offset: field.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::FormFieldEnd,
            });
        }
        if navigation_entries.len() > crate::navigation_entry::MAX_NAVIGATION_ENTRIES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF navigation-entry count limit exceeded",
            ));
        }
        let mut navigation_text_bytes = 0usize;
        for entry in navigation_entries {
            entry
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if body.get(entry.position()..entry.position()).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry position is outside body text or splits a character",
                ));
            }
            navigation_text_bytes = navigation_text_bytes
                .checked_add(entry.text_bytes().ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "navigation-entry size overflow")
                })?)
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "navigation-entry size overflow")
                })?;
            if navigation_text_bytes
                > crate::navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry aggregate text limit exceeded",
                ));
            }
            events.push(BodyEvent {
                offset: entry.position(),
                order: 1,
                kind: BodyEventKind::NavigationEntry(entry),
            });
        }
        for bookmark in bookmarks.bookmarks() {
            let end = bookmark
                .position
                .checked_add(bookmark.content.len())
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF bookmark range overflow")
                })?;
            let content = body.get(bookmark.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark range is outside body text or splits a character",
                )
            })?;
            if content != bookmark.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark content does not match its body range",
                ));
            }
            let empty = bookmark.content.is_empty();
            events.push(BodyEvent {
                offset: bookmark.position,
                order: 1,
                kind: BodyEventKind::BookmarkStart(bookmark),
            });
            events.push(BodyEvent {
                offset: end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::BookmarkEnd(bookmark),
            });
        }
        for annotation in annotations {
            annotation
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if annotation.range_end < annotation.position
                || body
                    .get(annotation.position..annotation.range_end)
                    .is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF annotation range is outside body text or splits a character",
                ));
            }
            let empty = annotation.position == annotation.range_end;
            events.push(BodyEvent {
                offset: annotation.position,
                order: 1,
                kind: BodyEventKind::AnnotationStart(annotation),
            });
            events.push(BodyEvent {
                offset: annotation.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::AnnotationEnd(annotation),
            });
        }
        let mut revision_ranges: Vec<&Revision<'_>> = revisions
            .iter()
            .filter(|revision| revision.revision_type == RevisionType::Insertion)
            .collect();
        revision_ranges.sort_by_key(|revision| (revision.position, revision.range_end));
        let mut previous_end = 0usize;
        for revision in revision_ranges {
            if revision.range_end <= revision.position
                || revision.position < previous_end
                || body.get(revision.position..revision.range_end).is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision ranges overlap, leave the body, or split a character",
                ));
            }
            let content = &body[revision.position..revision.range_end];
            if content != revision.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision content does not match its body range",
                ));
            }
            previous_end = revision.range_end;
            events.push(BodyEvent {
                offset: revision.position,
                order: 1,
                kind: BodyEventKind::RevisionStart(revision),
            });
            events.push(BodyEvent {
                offset: revision.range_end,
                order: 0,
                kind: BodyEventKind::RevisionEnd,
            });
        }
        for revision in revisions
            .iter()
            .filter(|revision| revision.revision_type == RevisionType::Deletion)
        {
            revision.validate().map_err(|error| {
                io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
            })?;
            if body.get(..revision.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF deletion position is outside body text or splits a character",
                ));
            }
            events.push(BodyEvent {
                offset: revision.position,
                order: 0,
                kind: BodyEventKind::RevisionDeletion(revision),
            });
        }
        events.sort_by_key(|event| (event.offset, event.order));

        let mut event_index = 0usize;
        let mut body_offset = 0usize;
        for block in blocks {
            let block_end = body_offset + block.text.len();
            let mut local_offset = 0usize;
            while event_index < events.len() && events[event_index].offset <= block_end {
                let event_offset = events[event_index].offset;
                if event_offset < body_offset {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF bookmark events are not ordered",
                    ));
                }
                let local_end = event_offset - body_offset;
                if local_end > local_offset {
                    self.write_style_block_fragment(block, local_offset, local_end)?;
                    local_offset = local_end;
                }
                while event_index < events.len() && events[event_index].offset == event_offset {
                    self.write_body_event(events[event_index])?;
                    event_index += 1;
                }
            }
            if local_offset < block.text.len() {
                self.write_style_block_fragment(block, local_offset, block.text.len())?;
            }
            body_offset = block_end;
        }
        while event_index < events.len() && events[event_index].offset == body_offset {
            self.write_body_event(events[event_index])?;
            event_index += 1;
        }
        if event_index != events.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark range extends beyond body text",
            ));
        }
        Ok(())
    }

    fn write_body_event(&mut self, event: BodyEvent<'_, '_>) -> io::Result<()> {
        match event.kind {
            BodyEventKind::NavigationEntry(entry) => self.write_navigation_entry(entry),
            BodyEventKind::BookmarkStart(bookmark) => self.write_bookmark_start(bookmark),
            BodyEventKind::BookmarkEnd(bookmark) => self.write_bookmark_end(bookmark.name.as_ref()),
            BodyEventKind::AnnotationStart(annotation) => self.write_annotation_start(annotation),
            BodyEventKind::AnnotationEnd(annotation) => self.write_annotation_end(annotation),
            BodyEventKind::RevisionStart(revision) => self.write_revision_start(revision),
            BodyEventKind::RevisionEnd => self.write_str("}"),
            BodyEventKind::RevisionDeletion(revision) => self.write_revision(revision),
            BodyEventKind::FormFieldStart(field) => self.write_form_field_start(field),
            BodyEventKind::FormFieldEnd => self.write_str("}}"),
        }
    }

    fn write_form_field_start(&mut self, field: &crate::FormField<'_>) -> io::Result<()> {
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\field{\\*\\fldinst ")?;
        self.write_str(match field.field_type {
            crate::FormFieldType::Text => "FORMTEXT",
            crate::FormFieldType::CheckBox => "FORMCHECKBOX",
            crate::FormFieldType::DropDown => "FORMDROPDOWN",
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
        if let Some(value) = field.half_point_size {
            self.write_control_word("ffhps", Some(value))?;
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

    fn write_form_field_value(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
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
    pub fn write_navigation_entry(
        &mut self,
        entry: &crate::NavigationEntry<'_>,
    ) -> io::Result<()> {
        entry
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        self.write_str("{")?;
        match entry {
            crate::NavigationEntry::Index(entry) => {
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
                    crate::IndexPageReference::CurrentPage => {},
                    crate::IndexPageReference::ReplacementText(value) => {
                        self.write_str("{")?;
                        self.write_control_word("txe", None)?;
                        self.write_str(" ")?;
                        self.write_destination_text(value.as_ref())?;
                        self.write_str("}")?;
                    },
                    crate::IndexPageReference::BookmarkRange(value) => {
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
            crate::NavigationEntry::TableOfContents(entry) => {
                self.write_control_word(
                    if entry.suppress_page_number { "tcn" } else { "tc" },
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

    fn write_style_block_fragment(
        &mut self,
        block: &StyleBlock<'_>,
        start: usize,
        end: usize,
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
        self.write_style_block(&fragment)
    }

    /// Write a style block
    fn write_style_block(&mut self, block: &StyleBlock) -> io::Result<()> {
        self.write_str("{")?;

        // Write character formatting
        self.write_formatting(&block.formatting)?;

        // Write paragraph properties
        self.write_paragraph_properties(&block.paragraph)?;

        // Delimit the final control word from body text that starts with a letter.
        self.write_str(" ")?;

        // Write text content
        self.write_text(block.text.as_ref())?;

        self.write_str("}")?;
        Ok(())
    }

    /// Write character formatting
    fn write_formatting(&mut self, fmt: &Formatting) -> io::Result<()> {
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

        // Font
        if fmt.font_ref != 0 {
            self.write_control_word("f", Some(fmt.font_ref as i32))?;
        }

        // Font size
        self.write_control_word("fs", Some(fmt.font_size.get() as i32))?;

        // Color
        if fmt.color_ref != 0 {
            self.write_control_word("cf", Some(fmt.color_ref as i32))?;
        }

        // Highlight
        if let Some(highlight) = fmt.highlight_color {
            self.write_control_word("highlight", Some(highlight as i32))?;
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
        }

        // Strike
        if fmt.strike {
            self.write_control_word("strike", None)?;
        }

        // Double strike
        if fmt.double_strike {
            self.write_control_word("striked", None)?;
        }

        // Superscript
        if fmt.superscript {
            self.write_control_word("super", None)?;
        }

        // Subscript
        if fmt.subscript {
            self.write_control_word("sub", None)?;
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

        // Character spacing
        if fmt.char_spacing != 0 {
            self.write_control_word("expnd", Some(fmt.char_spacing))?;
        }

        // Character scale
        if fmt.char_scale != 100 {
            self.write_control_word("charscalex", Some(fmt.char_scale))?;
        }

        // Kerning
        if fmt.kerning != 0 {
            self.write_control_word("kerning", Some(fmt.kerning))?;
        }

        Ok(())
    }

    /// Write paragraph properties
    fn write_paragraph_properties(&mut self, para: &Paragraph) -> io::Result<()> {
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

        // Borders (if any)
        self.write_borders(&para.borders)?;

        // Shading (if any)
        self.write_shading(&para.shading)?;

        // Note: Tab stops would be written here if they were part of Paragraph
        // For now, they would need to be passed separately or stored elsewhere

        // Keep together
        if para.keep_together {
            self.write_control_word("keep", None)?;
        }

        // Keep with next
        if para.keep_next {
            self.write_control_word("keepn", None)?;
        }

        // Page break before
        if para.page_break_before {
            self.write_control_word("pagebb", None)?;
        }

        // Widow control
        if para.widow_control {
            self.write_control_word("widctlpar", None)?;
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
    fn write_borders(&mut self, borders: &Borders) -> io::Result<()> {
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

        Ok(())
    }

    /// Write a single border
    fn write_border(&mut self, control: &str, border: &Border) -> io::Result<()> {
        self.write_control_word(control, None)?;

        // Border style
        let style_word = match border.style {
            BorderStyle::None => return Ok(()),
            BorderStyle::Single => "brdrs",
            BorderStyle::Dotted => "brdrdot",
            BorderStyle::Dashed => "brdrdash",
            BorderStyle::Double => "brdrdb",
            BorderStyle::Triple => "brdrtriple",
            BorderStyle::ThickThinSmall => "brdrtnthsg",
            BorderStyle::ThinThickSmall => "brdrtnthmg",
            BorderStyle::ThinThickThinSmall => "brdrtnthtnsg",
            BorderStyle::ThickThinMedium => "brdrtnthmg",
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
        };
        self.write_control_word(style_word, None)?;

        // Border width
        self.write_control_word("brdrw", Some(border.width))?;

        // Border color
        if border.color_ref != 0 {
            self.write_control_word("brdrcf", Some(border.color_ref as i32))?;
        }

        // Border space
        if border.space != 0 {
            self.write_control_word("brsp", Some(border.space))?;
        }

        Ok(())
    }

    /// Write shading
    fn write_shading(&mut self, shading: &Shading) -> io::Result<()> {
        if !shading.is_visible() {
            return Ok(());
        }

        // Shading pattern
        let pattern_value = match shading.pattern {
            ShadingPattern::Clear => return Ok(()),
            ShadingPattern::Solid => 10000,
            ShadingPattern::Percent5 => 500,
            ShadingPattern::Percent10 => 1000,
            ShadingPattern::Percent12 => 1250,
            ShadingPattern::Percent15 => 1500,
            ShadingPattern::Percent20 => 2000,
            ShadingPattern::Percent25 => 2500,
            ShadingPattern::Percent30 => 3000,
            ShadingPattern::Percent35 => 3500,
            ShadingPattern::Percent40 => 4000,
            ShadingPattern::Percent45 => 4500,
            ShadingPattern::Percent50 => 5000,
            ShadingPattern::Percent55 => 5500,
            ShadingPattern::Percent60 => 6000,
            ShadingPattern::Percent62 => 6250,
            ShadingPattern::Percent65 => 6500,
            ShadingPattern::Percent70 => 7000,
            ShadingPattern::Percent75 => 7500,
            ShadingPattern::Percent80 => 8000,
            ShadingPattern::Percent85 => 8500,
            ShadingPattern::Percent87 => 8750,
            ShadingPattern::Percent90 => 9000,
            ShadingPattern::Percent95 => 9500,
            _ => 0, // Other patterns need specific control words
        };

        if pattern_value > 0 {
            self.write_control_word("shading", Some(pattern_value))?;
        }

        // Foreground color
        if shading.foreground_color != 0 {
            self.write_control_word("cfpat", Some(shading.foreground_color as i32))?;
        }

        // Background color
        if shading.background_color != 0 {
            self.write_control_word("cbpat", Some(shading.background_color as i32))?;
        }

        Ok(())
    }

    /// Write tab stop
    ///
    /// # Note
    ///
    /// This method is provided for completeness but is not currently used in document
    /// serialization. It will be integrated once tab stops are fully implemented in
    /// the paragraph properties.
    #[allow(dead_code)]
    fn write_tab_stop(&mut self, tab: &TabStop) -> io::Result<()> {
        // Tab alignment
        match tab.alignment {
            TabAlignment::Left => self.write_control_word("tql", None)?,
            TabAlignment::Right => self.write_control_word("tqr", None)?,
            TabAlignment::Center => self.write_control_word("tqc", None)?,
            TabAlignment::Decimal => self.write_control_word("tqdec", None)?,
            TabAlignment::Bar => self.write_control_word("tb", None)?,
        }

        // Tab leader
        match tab.leader {
            TabLeader::None => {},
            TabLeader::Dot => self.write_control_word("tldot", None)?,
            TabLeader::Hyphen => self.write_control_word("tlhyph", None)?,
            TabLeader::Underscore => self.write_control_word("tlul", None)?,
            TabLeader::ThickLine => self.write_control_word("tlth", None)?,
            TabLeader::Equal => self.write_control_word("tleq", None)?,
        }

        // Tab position
        self.write_control_word("tx", Some(tab.position))?;

        Ok(())
    }

    /// Write a table
    fn write_table(&mut self, table: &Table) -> io::Result<()> {
        for row in table.rows() {
            self.write_table_row(row)?;
        }
        Ok(())
    }

    /// Write a table row
    fn write_table_row(&mut self, row: &Row) -> io::Result<()> {
        // Row defaults
        self.write_control_word("trowd", None)?;

        // Cell boundaries
        let cell_width = 2880; // Default cell width (2 inches)
        for (i, _cell) in row.cells().iter().enumerate() {
            let boundary = cell_width * ((i + 1) as i32);
            self.write_control_word("cellx", Some(boundary))?;
        }

        // Write cells
        for cell in row.cells() {
            self.write_str("{")?;
            self.write_control_word("intbl", None)?;
            self.write_text(cell.text())?;
            self.write_control_word("cell", None)?;
            self.write_str("}")?;
        }

        // Row end
        self.write_control_word("row", None)?;
        self.write_str("\n")?;

        Ok(())
    }

    /// Write a control word
    pub fn write_control_word(&mut self, word: &str, param: Option<i32>) -> io::Result<()> {
        self.write_str("\\")?;
        self.write_str(word)?;
        if let Some(p) = param {
            write!(self.writer, "{}", p)?;
        }
        Ok(())
    }

    /// Write plain text (with proper escaping)
    pub fn write_text(&mut self, text: &str) -> io::Result<()> {
        for ch in text.chars() {
            match ch {
                '\\' => self.write_str("\\\\")?,
                '{' => self.write_str("\\{")?,
                '}' => self.write_str("\\}")?,
                '\n' => self.write_control_word("par", None)?,
                '\t' => self.write_control_word("tab", None)?,
                c if c.is_ascii() => {
                    write!(self.writer, "{}", c)?;
                },
                c => {
                    // Write Unicode character
                    let code = c as i32;
                    self.write_control_word("u", Some(code))?;
                    // Fallback character
                    self.write_str("?")?;
                },
            }
        }
        Ok(())
    }

    /// Write a string
    pub fn write_str(&mut self, s: &str) -> io::Result<()> {
        self.writer.write_all(s.as_bytes())
    }

    /// Flush the writer
    pub fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }

    /// Write a header or footer
    pub fn write_header_footer(&mut self, hf: &HeaderFooter) -> io::Result<()> {
        self.write_str("{")?;

        // Write header/footer type control word
        match hf.header_type {
            HeaderFooterType::Header => self.write_control_word("header", None)?,
            HeaderFooterType::HeaderFirst => self.write_control_word("headerf", None)?,
            HeaderFooterType::HeaderLeft => self.write_control_word("headerl", None)?,
            HeaderFooterType::HeaderRight => self.write_control_word("headerr", None)?,
            HeaderFooterType::Footer => self.write_control_word("footer", None)?,
            HeaderFooterType::FooterFirst => self.write_control_word("footerf", None)?,
            HeaderFooterType::FooterLeft => self.write_control_word("footerl", None)?,
            HeaderFooterType::FooterRight => self.write_control_word("footerr", None)?,
        }

        // Write paragraphs
        for para in &hf.paragraphs {
            self.write_formatting(&para.formatting)?;
            self.write_paragraph_properties(&para.paragraph)?;
            self.write_str(" ")?;
            self.write_text(para.text.as_ref())?;
            self.write_control_word("par", None)?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write a footnote or endnote
    pub fn write_note(&mut self, note: &Note) -> io::Result<()> {
        self.write_str("{")?;

        // Write note type control word
        if note.is_footnote {
            self.write_control_word("footnote", None)?;
        } else {
            self.write_control_word("endnote", None)?;
        }

        // Write reference number/marker
        if !note.reference.is_empty()
            && let Ok(num) = note.reference.parse::<i32>()
        {
            self.write_control_word("chftn", Some(num))?;
        }

        // Write note content
        self.write_str(" {")?;
        self.write_formatting(&note.formatting)?;
        self.write_text(note.content.as_ref())?;
        self.write_str("}")?;

        self.write_str("}")?;
        Ok(())
    }

    /// Write a hyperlink field
    pub fn write_hyperlink(&mut self, url: &str, display_text: &str) -> io::Result<()> {
        let instruction = format!(
            "HYPERLINK {}",
            crate::field::quoted_field_operand(url)
        );
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    /// Write an internal bookmark hyperlink without exposing raw field syntax.
    pub fn write_internal_hyperlink(
        &mut self,
        bookmark: &str,
        display_text: &str,
    ) -> io::Result<()> {
        let instruction = format!(
            "HYPERLINK \\l {}",
            crate::field::quoted_field_operand(bookmark)
        );
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    fn write_hyperlink_instruction(
        &mut self,
        instruction: &str,
        display_text: &str,
    ) -> io::Result<()> {
        self.write_str("{\\field")?;

        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(instruction)?;
        self.write_str("}}")?;

        // Field result (display text)
        self.write_str("{\\fldrslt{")?;
        self.write_control_word("ul", None)?; // Underline hyperlinks by default
        self.write_control_word("cf", Some(1))?; // Blue color for hyperlinks
        self.write_text(display_text)?;
        self.write_str("}}}")?;

        Ok(())
    }

    /// Write a field (generic)
    pub fn write_field(&mut self, field: &Field) -> io::Result<()> {
        self.write_str("{\\field")?;

        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(field.instruction.as_ref())?;
        self.write_str("}}")?;

        // Field result
        if !field.result.is_empty() {
            self.write_str("{\\fldrslt{")?;
            self.write_text(field.result.as_ref())?;
            self.write_str("}}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write a revision mark (track changes)
    pub fn write_revision(&mut self, revision: &Revision) -> io::Result<()> {
        revision
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_revision_start(revision)?;
        self.write_text(revision.content.as_ref())?;
        self.write_str("}")?;
        Ok(())
    }

    fn write_revision_start(&mut self, revision: &Revision<'_>) -> io::Result<()> {
        self.write_str("{")?;
        let (kind, author, date) = match revision.revision_type {
            RevisionType::Insertion => ("revised", "revauth", "revdttm"),
            RevisionType::Deletion => ("deleted", "revauthdel", "revdttmdel"),
            RevisionType::FormatChange | RevisionType::MovedFrom | RevisionType::MovedTo => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "this RTF revision kind has no lossless scoped-run representation",
                ));
            },
        };
        self.write_control_word(kind, None)?;
        self.write_control_word(author, Some(revision.id))?;
        if let Some(date_value) = revision.date.as_deref() {
            let packed = date_value.parse::<i32>().map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision dates must contain the packed signed DTTM value",
                )
            })?;
            self.write_control_word(date, Some(packed))?;
        }
        self.write_str(" ")?;
        Ok(())
    }

    /// Write a section with headers and footers
    pub fn write_section(&mut self, section: &Section) -> io::Result<()> {
        // Write section properties
        self.write_control_word("sectd", None)?;

        match section.properties.break_type {
            SectionBreakType::Continuous => self.write_control_word("sbknone", None)?,
            SectionBreakType::Column => self.write_control_word("sbkcol", None)?,
            SectionBreakType::Page => self.write_control_word("sbkpage", None)?,
            SectionBreakType::EvenPage => self.write_control_word("sbkeven", None)?,
            SectionBreakType::OddPage => self.write_control_word("sbkodd", None)?,
        }

        // Page size
        self.write_control_word("pgwsxn", Some(section.properties.page_width))?;
        self.write_control_word("pghsxn", Some(section.properties.page_height))?;

        // Margins
        self.write_control_word("marglsxn", Some(section.properties.margin_left))?;
        self.write_control_word("margrsxn", Some(section.properties.margin_right))?;
        self.write_control_word("margtsxn", Some(section.properties.margin_top))?;
        self.write_control_word("margbsxn", Some(section.properties.margin_bottom))?;
        self.write_control_word("guttersxn", Some(section.properties.margin_gutter))?;

        // Header/footer distance
        self.write_control_word("headery", Some(section.properties.header_distance))?;
        self.write_control_word("footery", Some(section.properties.footer_distance))?;

        if section.properties.orientation == PageOrientation::Landscape {
            self.write_control_word("lndscpsxn", None)?;
        }
        self.write_control_word("cols", Some(i32::from(section.properties.columns)))?;
        self.write_control_word("colsx", Some(section.properties.column_space))?;
        self.write_control_word("pgnstarts", Some(section.properties.page_number_start))?;
        self.write_control_word(
            match section.properties.page_number_format {
                PageNumberFormat::Decimal => "pgndec",
                PageNumberFormat::UpperRoman => "pgnucrm",
                PageNumberFormat::LowerRoman => "pgnlcrm",
                PageNumberFormat::UpperLetter => "pgnucltr",
                PageNumberFormat::LowerLetter => "pgnlcltr",
            },
            None,
        )?;
        self.write_control_word(
            match section.properties.vertical_alignment {
                VerticalAlignment::Top => "vertalt",
                VerticalAlignment::Center => "vertalc",
                VerticalAlignment::Justify => "vertalj",
                VerticalAlignment::Bottom => "vertalb",
            },
            None,
        )?;
        if section.properties.line_numbering {
            self.write_control_word("linemod", Some(1))?;
            self.write_control_word(
                if section.properties.line_number_restart {
                    "lineppage"
                } else {
                    "linecont"
                },
                None,
            )?;
        }

        // Write all headers and footers for this section
        for hf in &section.headers_footers {
            self.write_header_footer(hf)?;
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_document() {
        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);

        writer.write_document_header().unwrap();
        writer.write_text("Hello World").unwrap();
        writer.write_str("}").unwrap();

        let result = String::from_utf8(output).unwrap();
        assert!(result.contains("rtf1"));
        assert!(result.contains("Hello World"));
    }

    #[test]
    fn test_control_words() {
        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);

        writer.write_control_word("test", Some(42)).unwrap();
        writer.write_control_word("flag", None).unwrap();

        let result = String::from_utf8(output).unwrap();
        assert_eq!(result, "\\test42\\flag");
    }

    #[test]
    fn document_info_writer_round_trips() {
        let mut info = DocumentInfo::new().with_title(std::borrow::Cow::Borrowed("Résumé 你"));
        info.author = Some(std::borrow::Cow::Borrowed("Ada"));
        info.creation_time = Some(std::borrow::Cow::Borrowed("2026-07-15T12:34:56"));
        info.pages = Some(3);
        info.characters_with_spaces = Some(42);

        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);
        writer.write_document_header().unwrap();
        writer.write_document_info(&info).unwrap();
        writer.write_text("Body").unwrap();
        writer.write_str("}").unwrap();

        let rtf = String::from_utf8(output).unwrap();
        let parsed = RtfDocument::parse(&rtf).unwrap();
        assert_eq!(parsed.info().title.as_deref(), Some("Résumé 你"));
        assert_eq!(parsed.info().author.as_deref(), Some("Ada"));
        assert_eq!(
            parsed.info().creation_time.as_deref(),
            Some("2026-07-15T12:34:56")
        );
        assert_eq!(parsed.info().pages, Some(3));
        assert_eq!(parsed.info().characters_with_spaces, Some(42));
        assert_eq!(parsed.text(), "Body");
    }

    #[test]
    fn document_writer_round_trips_bookmark_ranges() {
        let source = r#"{\rtf1\ansi Start {\*\bkmkstart\bkmkcolf2\bkmkcoll4\bkmkpub Link}R\'e9sum\'e9 \u20320?{\*\bkmkend Link} end}"#;
        let document = RtfDocument::parse(source).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        let bookmark = reparsed.bookmarks().get("Link").unwrap();
        assert_eq!(bookmark.content, "Résumé 你");
        assert_eq!(bookmark.first_column, Some(2));
        assert_eq!(bookmark.last_column, Some(4));
        assert!(bookmark.is_public);
    }

    #[test]
    fn document_writer_preserves_bookmark_in_empty_body() {
        let document =
            RtfDocument::parse(r#"{\rtf1{\*\bkmkstart Empty}{\*\bkmkend Empty}}"#).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        let bookmark = reparsed.bookmarks().get("Empty").unwrap();
        assert_eq!(bookmark.position, 0);
        assert!(bookmark.content.is_empty());
        assert!(reparsed.text().is_empty());
    }

    #[test]
    fn document_writer_round_trips_annotations() {
        let source = r#"{\rtf1\ansi Before {\*\atrfstart 12}range{\*\atrfend 12}{\*\atnid AM}{\*\atnauthor Ada M}\chatn{\*\annotation{\*\atnref 12}{\*\atndate 12345}{\*\atnparent 4}{\*\atnicn 3}{\*\atntime 99}Review \u20320? now} after}"#;
        let document = RtfDocument::parse(source).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.annotations().len(), 1);
        let annotation = &reparsed.annotations()[0];
        assert_eq!(annotation.id, 12);
        assert_eq!(annotation.author, "Ada M");
        assert_eq!(annotation.initials, "AM");
        assert_eq!(annotation.date.as_deref(), Some("12345"));
        assert_eq!(annotation.parent_id.as_deref(), Some("4"));
        assert_eq!(annotation.icon.as_deref(), Some("3"));
        assert_eq!(annotation.time.as_deref(), Some("99"));
        assert_eq!(annotation.text, "Review 你 now");
        assert_eq!(annotation.position, "Before ".len());
        assert_eq!(annotation.range_end, "Before range".len());
    }

    #[test]
    fn document_writer_round_trips_headers_and_footers() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi\sectd\sbkodd\pgwsxn11000\pghsxn15000\marglsxn910\margrsxn810\margtsxn710\margbsxn610\guttersxn130\headery310\footery410\lndscpsxn\cols3\colsx370\pgnstarts6\pgnlcltr\vertalb\linemod1\lineppage{\header Header \u20320? one\par Header two}{\footer Footer}Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        assert_eq!(reparsed.sections().len(), 1);
        let section = &reparsed.sections()[0];
        assert_eq!(section.properties, document.sections()[0].properties);
        assert_eq!(
            section.get_header(HeaderFooterType::Header).unwrap().text(),
            "Header 你 one\nHeader two"
        );
        assert_eq!(
            section.get_header(HeaderFooterType::Footer).unwrap().text(),
            "Footer"
        );
    }

    #[test]
    fn document_writer_round_trips_stylesheets() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\stylesheet{\s0\snext0 Normal;}{\s1\b\qc\sbasedon0\snext0\slink2\sautoupd\shidden\slocked\ssemihidden\sunhideused\sqformat\spriority9\styrsid42 Heading \u20320?;}{\*\cs2\i\additive\slink1 Emphasis;}{\*\ds3 Section;}{\*\ts4 Table;}}Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        assert_eq!(reparsed.stylesheet().styles().len(), 5);
        let heading = reparsed
            .stylesheet()
            .get_typed(StyleType::Paragraph, 1)
            .unwrap();
        assert_eq!(heading.name, "Heading 你");
        assert!(heading.formatting.bold);
        assert_eq!(heading.paragraph.unwrap().alignment, Alignment::Center);
        assert_eq!(heading.linked_style, Some(2));
        assert!(heading.auto_update);
        assert!(heading.hidden);
        assert!(heading.locked);
        assert!(heading.semi_hidden);
        assert!(heading.unhide_when_used);
        assert!(heading.quick_format);
        assert_eq!(heading.priority, Some(9));
        assert_eq!(heading.revision_id, Some(42));

        let character = reparsed
            .stylesheet()
            .get_typed(StyleType::Character, 2)
            .unwrap();
        assert!(character.additive);
        assert!(character.formatting.italic);
        assert!(
            reparsed
                .stylesheet()
                .get_typed(StyleType::Section, 3)
                .is_some()
        );
        assert!(
            reparsed
                .stylesheet()
                .get_typed(StyleType::Table, 4)
                .is_some()
        );
    }

    #[test]
    fn document_writer_round_trips_list_tables() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\*\listtable{\list\listtemplateid42\listhybrid{\listlevel\levelnfc0\leveljc2\levelfollow1\levelstartat3\levelspace120\levelindent360{\leveltext\'02\'00.;}{\levelnumbers\'01;}\f2}{\listlevel\levelnfc77\leveljc0\levelfollow2\levelstartat1{\leveltext\'01\u8226?;}{\levelnumbers;}}{\listname Outline;}\listid77}}{\*\listoverridetable{\listoverride\listid77\listoverridecount1{\lfolevel\listoverridestartat\levelstartat9}\ls4}}\pard\ls4\ilvl1 Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        let paragraph = reparsed.blocks().last().unwrap().paragraph;
        assert_eq!(paragraph.list_override, Some(4));
        assert_eq!(paragraph.list_level, Some(1));
        let list = reparsed.list_table().get(77).unwrap();
        assert_eq!(list.template_id, 42);
        assert!(list.hybrid);
        assert_eq!(list.name, "Outline");
        assert_eq!(list.levels.len(), 2);
        assert_eq!(list.levels[0].number_text, "\0.");
        assert_eq!(list.levels[0].follow, ListFollow::Space);
        assert_eq!(list.levels[1].level_type, ListLevelType::Other(77));
        assert_eq!(list.levels[1].number_text, "•");
        assert_eq!(list.levels[1].follow, ListFollow::Nothing);
        let list_override = reparsed.list_override_table().get(4).unwrap();
        assert_eq!(list_override.list_id, 77);
        assert_eq!(list_override.level_count_override, Some(1));
        assert_eq!(list_override.start_at_override, Some(9));
    }

    #[test]
    fn document_writer_round_trips_tracked_revision_ranges() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\*\revtbl{Unknown;}{Ada;}}Before {\deleted\revauthdel1\revdttmdel123 old}{\revised\revauth1\revdttm-456 new \u20320?} after}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap_or_else(|error| {
            panic!(
                "failed to parse revision writer output: {error}\n{}",
                String::from_utf8_lossy(&output)
            )
        });
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.revisions().len(), 2);
        for (actual, expected) in reparsed.revisions().iter().zip(document.revisions()) {
            assert_eq!(actual.revision_type, expected.revision_type);
            assert_eq!(actual.id, expected.id);
            assert_eq!(actual.author, expected.author);
            assert_eq!(actual.date, expected.date);
            assert_eq!(actual.content, expected.content);
            assert_eq!(actual.position, expected.position);
            assert_eq!(actual.range_end, expected.range_end);
        }
    }

    #[test]
    fn document_writer_rejects_ambiguous_multiple_sections() {
        let document =
            RtfDocument::parse(r#"{\rtf1\sectd{\header First}One\sect\sectd{\header Second}Two}"#)
                .unwrap();
        assert_eq!(document.sections().len(), 2);
        let error = RtfWriter::new(Vec::new())
            .write_document(&document)
            .unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
    }
}
