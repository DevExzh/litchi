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
    /// List table (reserved for writing lists)
    #[allow(dead_code)]
    list_table: ListTable<'static>,
    /// List override table (reserved for writing lists)
    #[allow(dead_code)]
    list_override_table: ListOverrideTable,
    /// Stylesheet (reserved for writing styles)
    #[allow(dead_code)]
    stylesheet: StyleSheet<'static>,
}

#[derive(Clone, Copy)]
enum BodyEventKind<'b, 'a> {
    BookmarkStart(&'b Bookmark<'a>),
    BookmarkEnd(&'b Bookmark<'a>),
    AnnotationStart(&'b Annotation<'a>),
    AnnotationEnd(&'b Annotation<'a>),
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
            list_table: ListTable::new(),
            list_override_table: ListOverrideTable::new(),
            stylesheet: StyleSheet::new(),
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

        // Write font table
        self.write_font_table()?;

        // Write color table
        self.write_color_table()?;

        // Write document properties before body content.
        self.write_document_info(doc.info())?;

        // Headers and footers belong to the section definition before body text.
        for section in doc.sections() {
            self.write_section(section)?;
        }

        // Write document content and reinsert positional bookmark/comment markers.
        self.write_blocks_with_markup(doc.blocks(), doc.bookmarks(), doc.annotations())?;

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
        self.write_str("{\\*")?;
        self.write_control_word("atrfend", None)?;
        self.write_str(" ")?;
        write!(self.writer, "{}", annotation.id)?;
        self.write_str("}")?;
        self.write_annotation_value("atnid", Some(annotation.initials.as_ref()))?;
        self.write_annotation_value("atnauthor", Some(annotation.author.as_ref()))?;
        self.write_control_word("chatn", None)?;
        self.write_str("{\\*")?;
        self.write_control_word("annotation", None)?;
        self.write_annotation_value("atnref", Some(&annotation.id.to_string()))?;
        self.write_annotation_value("atndate", annotation.date.as_deref())?;
        self.write_annotation_value("atnparent", annotation.parent_id.as_deref())?;
        self.write_annotation_value("atnicn", annotation.icon.as_deref())?;
        self.write_annotation_value("atntime", annotation.time.as_deref())?;
        self.write_text(annotation.text.as_ref())?;
        self.write_str("}")
    }

    fn write_annotation_value(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value.filter(|value| !value.is_empty()) else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_text(value)?;
        self.write_str("}")
    }

    fn write_blocks_with_markup(
        &mut self,
        blocks: &[StyleBlock<'_>],
        bookmarks: &BookmarkTable<'_>,
        annotations: &[Annotation<'_>],
    ) -> io::Result<()> {
        if bookmarks.bookmarks().is_empty() && annotations.is_empty() {
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
            .saturating_mul(2);
        let mut events = Vec::with_capacity(event_count);
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
            BodyEventKind::BookmarkStart(bookmark) => self.write_bookmark_start(bookmark),
            BodyEventKind::BookmarkEnd(bookmark) => self.write_bookmark_end(bookmark.name.as_ref()),
            BodyEventKind::AnnotationStart(annotation) => self.write_annotation_start(annotation),
            BodyEventKind::AnnotationEnd(annotation) => self.write_annotation_end(annotation),
        }
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
        self.write_str("{\\field")?;

        // Field instruction
        self.write_str("{\\*\\fldinst{HYPERLINK \"")?;
        self.write_text(url)?;
        self.write_str("\"}}")?;

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
        self.write_str("{")?;

        // Write revision type control word
        match revision.revision_type {
            RevisionType::Insertion => {
                self.write_control_word("revised", None)?;
                self.write_control_word("revauth", Some(revision.id))?;
                if !revision.author.is_empty() {
                    // Write author in annotation
                    self.write_str("{\\*\\atnauthor ")?;
                    self.write_text(revision.author.as_ref())?;
                    self.write_str("}")?;
                }
            },
            RevisionType::Deletion => {
                self.write_control_word("deleted", None)?;
                self.write_control_word("revauthdel", Some(revision.id))?;
            },
            RevisionType::FormatChange => {
                self.write_control_word("revprop", None)?;
            },
            RevisionType::MovedFrom => {
                self.write_control_word("movedfrom", None)?;
            },
            RevisionType::MovedTo => {
                self.write_control_word("movedto", None)?;
            },
        }

        // Write content
        self.write_text(revision.content.as_ref())?;

        self.write_str("}")?;
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
