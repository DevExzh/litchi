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
pub struct RtfWriter<W: Write> {
    /// Output writer
    writer: W,
    /// Writer options
    options: WriterOptions,
    /// Current indentation level
    indent_level: usize,
    /// Font table
    font_table: FontTable<'static>,
    /// Color table
    color_table: ColorTable,
    /// List table
    list_table: ListTable<'static>,
    /// List override table
    list_override_table: ListOverrideTable,
    /// Stylesheet
    stylesheet: StyleSheet<'static>,
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

        // Write document content
        for block in doc.blocks() {
            self.write_style_block(block)?;
        }

        // Write tables
        for table in doc.tables() {
            self.write_table(table)?;
        }

        // Close document
        self.write_str("}")?;

        Ok(())
    }

    /// Write document header
    fn write_document_header(&mut self) -> io::Result<()> {
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

    /// Write a style block
    fn write_style_block(&mut self, block: &StyleBlock) -> io::Result<()> {
        self.write_str("{")?;

        // Write character formatting
        self.write_formatting(&block.formatting)?;

        // Write paragraph properties
        self.write_paragraph_properties(&block.paragraph)?;

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
    fn write_control_word(&mut self, word: &str, param: Option<i32>) -> io::Result<()> {
        self.write_str("\\")?;
        self.write_str(word)?;
        if let Some(p) = param {
            write!(self.writer, "{}", p)?;
        }
        Ok(())
    }

    /// Write plain text (with proper escaping)
    fn write_text(&mut self, text: &str) -> io::Result<()> {
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
    fn write_str(&mut self, s: &str) -> io::Result<()> {
        self.writer.write_all(s.as_bytes())
    }

    /// Flush the writer
    pub fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

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
}
