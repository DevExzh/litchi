/// Style writer support for DOCX documents.
///
/// This module provides functionality for creating and writing document styles.
use crate::ooxml::docx::enums::WdStyleType;
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A mutable style definition for writing.
///
/// Styles define reusable formatting that can be applied to paragraphs,
/// characters, tables, and lists. This includes built-in styles (like "Heading 1")
/// and custom user-defined styles.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi::ooxml::docx::writer::MutableStyle;
/// use litchi::ooxml::docx::enums::WdStyleType;
///
/// // Create a custom paragraph style
/// let mut style = MutableStyle::new("MyStyle", "My Custom Style", WdStyleType::Paragraph);
/// style.set_based_on(Some("Normal".to_string()));
/// style.set_font_size(Some(24)); // 12pt (half-points)
/// style.set_bold(true);
/// ```
#[derive(Debug, Clone)]
pub struct MutableStyle {
    /// Style identifier (required, e.g., "Heading1")
    style_id: String,
    /// UI-visible name (e.g., "Heading 1")
    name: String,
    /// Type of style (paragraph, character, table, or list)
    style_type: WdStyleType,
    /// Whether this is the default style for its type
    is_default: bool,
    /// Whether this is a custom (user-defined) style
    is_custom: bool,
    /// ID of the style this is based on
    based_on: Option<String>,
    /// UI priority for display ordering (lower = higher priority)
    priority: Option<i32>,
    /// Whether to show in quick style gallery
    is_quick_style: bool,
    /// Whether hidden from UI
    is_hidden: bool,
    /// Whether locked (formatting protection)
    is_locked: bool,
    /// Font family name (e.g., "Calibri", "Times New Roman")
    font_name: Option<String>,
    /// Font size in half-points (e.g., 24 = 12pt)
    font_size: Option<u32>,
    /// Bold formatting
    bold: bool,
    /// Italic formatting
    italic: bool,
    /// Underline formatting
    underline: bool,
    /// Font color (RGB hex format, e.g., "FF0000" for red)
    color: Option<String>,
    /// Paragraph alignment for paragraph styles
    alignment: Option<String>,
    /// Space before paragraph in twips (1/1440 inch)
    space_before: Option<u32>,
    /// Space after paragraph in twips
    space_after: Option<u32>,
    /// Line spacing (e.g., "240" for single spacing)
    line_spacing: Option<u32>,
    /// Left indent in twips
    indent_left: Option<i32>,
    /// Right indent in twips
    indent_right: Option<i32>,
    /// First line indent in twips (negative for hanging)
    indent_first_line: Option<i32>,
}

impl MutableStyle {
    /// Create a new style with the given ID, name, and type.
    ///
    /// # Arguments
    ///
    /// * `style_id` - Unique identifier for the style (e.g., "MyStyle1")
    /// * `name` - Display name for the style (e.g., "My Custom Style")
    /// * `style_type` - Type of style (paragraph, character, table, or list)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let style = MutableStyle::new(
    ///     "MyHeading",
    ///     "My Custom Heading",
    ///     WdStyleType::Paragraph
    /// );
    /// ```
    pub fn new(
        style_id: impl Into<String>,
        name: impl Into<String>,
        style_type: WdStyleType,
    ) -> Self {
        Self {
            style_id: style_id.into(),
            name: name.into(),
            style_type,
            is_default: false,
            is_custom: true,
            based_on: None,
            priority: None,
            is_quick_style: false,
            is_hidden: false,
            is_locked: false,
            font_name: None,
            font_size: None,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            alignment: None,
            space_before: None,
            space_after: None,
            line_spacing: None,
            indent_left: None,
            indent_right: None,
            indent_first_line: None,
        }
    }

    /// Get the style identifier.
    #[inline]
    pub fn style_id(&self) -> &str {
        &self.style_id
    }

    /// Get the style name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the style name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    /// Get the style type.
    #[inline]
    pub fn style_type(&self) -> WdStyleType {
        self.style_type
    }

    /// Set whether this is the default style for its type.
    pub fn set_default(&mut self, is_default: bool) {
        self.is_default = is_default;
    }

    /// Check if this is the default style.
    #[inline]
    pub fn is_default(&self) -> bool {
        self.is_default
    }

    /// Set whether this is a custom style.
    pub fn set_custom(&mut self, is_custom: bool) {
        self.is_custom = is_custom;
    }

    /// Check if this is a custom style.
    #[inline]
    pub fn is_custom(&self) -> bool {
        self.is_custom
    }

    /// Set the base style ID.
    ///
    /// The style inherits formatting from the base style.
    pub fn set_based_on(&mut self, based_on: Option<String>) {
        self.based_on = based_on;
    }

    /// Get the base style ID.
    #[inline]
    pub fn based_on(&self) -> Option<&str> {
        self.based_on.as_deref()
    }

    /// Set the UI priority (lower values appear first).
    pub fn set_priority(&mut self, priority: Option<i32>) {
        self.priority = priority;
    }

    /// Set whether to show in quick style gallery.
    pub fn set_quick_style(&mut self, is_quick_style: bool) {
        self.is_quick_style = is_quick_style;
    }

    /// Set whether hidden from UI.
    pub fn set_hidden(&mut self, is_hidden: bool) {
        self.is_hidden = is_hidden;
    }

    /// Set whether locked (formatting protection).
    pub fn set_locked(&mut self, is_locked: bool) {
        self.is_locked = is_locked;
    }

    /// Set the font name.
    pub fn set_font_name(&mut self, font_name: Option<String>) {
        self.font_name = font_name;
    }

    /// Set the font size in half-points (e.g., 24 = 12pt).
    pub fn set_font_size(&mut self, font_size: Option<u32>) {
        self.font_size = font_size;
    }

    /// Set bold formatting.
    pub fn set_bold(&mut self, bold: bool) {
        self.bold = bold;
    }

    /// Set italic formatting.
    pub fn set_italic(&mut self, italic: bool) {
        self.italic = italic;
    }

    /// Set underline formatting.
    pub fn set_underline(&mut self, underline: bool) {
        self.underline = underline;
    }

    /// Set the font color (RGB hex format, e.g., "FF0000" for red).
    pub fn set_color(&mut self, color: Option<String>) {
        self.color = color;
    }

    /// Set paragraph alignment ("left", "center", "right", "justify").
    pub fn set_alignment(&mut self, alignment: Option<String>) {
        self.alignment = alignment;
    }

    /// Set space before paragraph in twips (1/1440 inch).
    pub fn set_space_before(&mut self, space_before: Option<u32>) {
        self.space_before = space_before;
    }

    /// Set space after paragraph in twips.
    pub fn set_space_after(&mut self, space_after: Option<u32>) {
        self.space_after = space_after;
    }

    /// Set line spacing (e.g., 240 for single, 480 for double).
    pub fn set_line_spacing(&mut self, line_spacing: Option<u32>) {
        self.line_spacing = line_spacing;
    }

    /// Set left indent in twips.
    pub fn set_indent_left(&mut self, indent_left: Option<i32>) {
        self.indent_left = indent_left;
    }

    /// Set right indent in twips.
    pub fn set_indent_right(&mut self, indent_right: Option<i32>) {
        self.indent_right = indent_right;
    }

    /// Set first line indent in twips (negative for hanging indent).
    pub fn set_indent_first_line(&mut self, indent_first_line: Option<i32>) {
        self.indent_first_line = indent_first_line;
    }

    /// Generate XML for this style.
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(512);

        // Start style element
        write!(
            &mut xml,
            r#"<w:style w:type="{}" w:styleId="{}""#,
            self.style_type.to_xml(),
            escape_xml(&self.style_id)
        )?;

        if self.is_default {
            xml.push_str(r#" w:default="1""#);
        }
        if self.is_custom {
            xml.push_str(r#" w:customStyle="1""#);
        }

        xml.push('>');

        // Name element
        write!(&mut xml, r#"<w:name w:val="{}"/>"#, escape_xml(&self.name))?;

        // Based on
        if let Some(ref based_on) = self.based_on {
            write!(&mut xml, r#"<w:basedOn w:val="{}"/>"#, escape_xml(based_on))?;
        }

        // Priority
        if let Some(priority) = self.priority {
            write!(&mut xml, r#"<w:uiPriority w:val="{}"/>"#, priority)?;
        }

        // Quick style
        if self.is_quick_style {
            xml.push_str("<w:qFormat/>");
        }

        // Hidden
        if self.is_hidden {
            xml.push_str("<w:semiHidden/>");
        }

        // Locked
        if self.is_locked {
            xml.push_str("<w:locked/>");
        }

        // Paragraph properties (for paragraph and table styles)
        if matches!(self.style_type, WdStyleType::Paragraph | WdStyleType::Table) {
            let has_para_props = self.alignment.is_some()
                || self.space_before.is_some()
                || self.space_after.is_some()
                || self.line_spacing.is_some()
                || self.indent_left.is_some()
                || self.indent_right.is_some()
                || self.indent_first_line.is_some();

            if has_para_props {
                xml.push_str("<w:pPr>");

                if let Some(ref alignment) = self.alignment {
                    write!(&mut xml, r#"<w:jc w:val="{}"/>"#, alignment)?;
                }

                if self.space_before.is_some()
                    || self.space_after.is_some()
                    || self.line_spacing.is_some()
                {
                    xml.push_str("<w:spacing");
                    if let Some(before) = self.space_before {
                        write!(&mut xml, r#" w:before="{}""#, before)?;
                    }
                    if let Some(after) = self.space_after {
                        write!(&mut xml, r#" w:after="{}""#, after)?;
                    }
                    if let Some(line) = self.line_spacing {
                        write!(&mut xml, r#" w:line="{}""#, line)?;
                    }
                    xml.push_str("/>");
                }

                if self.indent_left.is_some()
                    || self.indent_right.is_some()
                    || self.indent_first_line.is_some()
                {
                    xml.push_str("<w:ind");
                    if let Some(left) = self.indent_left {
                        write!(&mut xml, r#" w:left="{}""#, left)?;
                    }
                    if let Some(right) = self.indent_right {
                        write!(&mut xml, r#" w:right="{}""#, right)?;
                    }
                    if let Some(first_line) = self.indent_first_line {
                        if first_line >= 0 {
                            write!(&mut xml, r#" w:firstLine="{}""#, first_line)?;
                        } else {
                            write!(&mut xml, r#" w:hanging="{}""#, -first_line)?;
                        }
                    }
                    xml.push_str("/>");
                }

                xml.push_str("</w:pPr>");
            }
        }

        // Run properties (character formatting)
        let has_run_props = self.font_name.is_some()
            || self.font_size.is_some()
            || self.bold
            || self.italic
            || self.underline
            || self.color.is_some();

        if has_run_props {
            xml.push_str("<w:rPr>");

            if let Some(ref font_name) = self.font_name {
                write!(
                    &mut xml,
                    r#"<w:rFonts w:ascii="{}" w:hAnsi="{}" w:cs="{}"/>"#,
                    escape_xml(font_name),
                    escape_xml(font_name),
                    escape_xml(font_name)
                )?;
            }

            if self.bold {
                xml.push_str("<w:b/>");
            }

            if self.italic {
                xml.push_str("<w:i/>");
            }

            if self.underline {
                xml.push_str(r#"<w:u w:val="single"/>"#);
            }

            if let Some(size) = self.font_size {
                write!(&mut xml, r#"<w:sz w:val="{}"/>"#, size)?;
                write!(&mut xml, r#"<w:szCs w:val="{}"/>"#, size)?;
            }

            if let Some(ref color) = self.color {
                write!(&mut xml, r#"<w:color w:val="{}"/>"#, escape_xml(color))?;
            }

            xml.push_str("</w:rPr>");
        }

        xml.push_str("</w:style>");

        Ok(xml)
    }

    /// Factory methods for common built-in styles
    ///
    /// Create a "Normal" paragraph style (base style).
    pub fn normal() -> Self {
        let mut style = Self::new("Normal", "Normal", WdStyleType::Paragraph);
        style.set_default(true);
        style.set_custom(false);
        style.set_font_name(Some("Calibri".to_string()));
        style.set_font_size(Some(22)); // 11pt
        style
    }

    /// Create a "Heading 1" style.
    pub fn heading_1() -> Self {
        let mut style = Self::new("Heading1", "Heading 1", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_font_name(Some("Calibri Light".to_string()));
        style.set_font_size(Some(32)); // 16pt
        style.set_color(Some("2F5496".to_string())); // Blue
        style.set_space_before(Some(240)); // 12pt before
        style.set_space_after(Some(0));
        style.set_priority(Some(9));
        style.set_quick_style(true);
        style
    }

    /// Create a "Heading 2" style.
    pub fn heading_2() -> Self {
        let mut style = Self::new("Heading2", "Heading 2", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_font_name(Some("Calibri Light".to_string()));
        style.set_font_size(Some(26)); // 13pt
        style.set_color(Some("2F5496".to_string()));
        style.set_space_before(Some(40)); // 2pt before
        style.set_space_after(Some(0));
        style.set_priority(Some(9));
        style.set_quick_style(true);
        style
    }

    /// Create a "Heading 3" style.
    pub fn heading_3() -> Self {
        let mut style = Self::new("Heading3", "Heading 3", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_font_name(Some("Calibri Light".to_string()));
        style.set_font_size(Some(24)); // 12pt
        style.set_color(Some("1F3763".to_string()));
        style.set_space_before(Some(40));
        style.set_space_after(Some(0));
        style.set_priority(Some(9));
        style.set_quick_style(true);
        style
    }

    /// Create a "Title" style.
    pub fn title() -> Self {
        let mut style = Self::new("Title", "Title", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_font_name(Some("Calibri Light".to_string()));
        style.set_font_size(Some(56)); // 28pt
        style.set_space_after(Some(0));
        style.set_priority(Some(10));
        style.set_quick_style(true);
        style
    }

    /// Create a default character style.
    pub fn default_paragraph_font() -> Self {
        let mut style = Self::new(
            "DefaultParagraphFont",
            "Default Paragraph Font",
            WdStyleType::Character,
        );
        style.set_default(true);
        style.set_custom(false);
        style.set_priority(Some(1));
        style
    }

    /// Create a TOC heading style (for "Table of Contents" title).
    pub fn toc_heading() -> Self {
        let mut style = Self::new("TOCHeading", "TOC Heading", WdStyleType::Paragraph);
        style.set_based_on(Some("Heading1".to_string()));
        style.set_custom(false);
        style.set_priority(Some(39));
        style
    }

    /// Create a TOC level 1 style.
    pub fn toc1() -> Self {
        let mut style = Self::new("TOC1", "toc 1", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(39));
        style.set_space_after(Some(100)); // 5pt after
        style
    }

    /// Create a TOC level 2 style.
    pub fn toc2() -> Self {
        let mut style = Self::new("TOC2", "toc 2", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(39));
        style.set_indent_left(Some(440)); // 0.31 inch
        style.set_space_after(Some(100));
        style
    }

    /// Create a TOC level 3 style.
    pub fn toc3() -> Self {
        let mut style = Self::new("TOC3", "toc 3", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(39));
        style.set_indent_left(Some(880)); // 0.61 inch
        style.set_space_after(Some(100));
        style
    }

    /// Create a hyperlink character style.
    pub fn hyperlink() -> Self {
        let mut style = Self::new("Hyperlink", "Hyperlink", WdStyleType::Character);
        style.set_based_on(Some("DefaultParagraphFont".to_string()));
        style.set_custom(false);
        style.set_priority(Some(99));
        style.set_color(Some("0563C1".to_string())); // Blue color
        style.set_underline(true);
        style
    }

    /// Create a header paragraph style.
    pub fn header() -> Self {
        let mut style = Self::new("Header", "header", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(99));
        style
    }

    /// Create a footer paragraph style.
    pub fn footer() -> Self {
        let mut style = Self::new("Footer", "footer", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(99));
        style
    }

    /// Create a footnote text paragraph style.
    pub fn footnote_text() -> Self {
        let mut style = Self::new("FootnoteText", "footnote text", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(99));
        style.set_font_size(Some(20)); // 10pt (half-points)
        style
    }

    /// Create an endnote text paragraph style.
    pub fn endnote_text() -> Self {
        let mut style = Self::new("EndnoteText", "endnote text", WdStyleType::Paragraph);
        style.set_based_on(Some("Normal".to_string()));
        style.set_custom(false);
        style.set_priority(Some(99));
        style.set_font_size(Some(20)); // 10pt (half-points)
        style
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Generate a complete styles.xml document from a list of styles.
///
/// # Arguments
///
/// * `styles` - Collection of styles to include in the document
///
/// # Returns
///
/// XML string representing the complete styles.xml content
pub fn generate_styles_xml(styles: &[MutableStyle]) -> Result<String> {
    let mut xml = String::with_capacity(4096);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    );
    xml.push_str(
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );

    // Add default document defaults
    xml.push_str("<w:docDefaults>");
    xml.push_str("<w:rPrDefault><w:rPr>");
    xml.push_str(r#"<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>"#);
    xml.push_str(r#"<w:sz w:val="22"/>"#);
    xml.push_str(r#"<w:szCs w:val="22"/>"#);
    xml.push_str("</w:rPr></w:rPrDefault>");
    xml.push_str("<w:pPrDefault/>");
    xml.push_str("</w:docDefaults>");

    // Add each style
    for style in styles {
        let style_xml = style.to_xml()?;
        xml.push_str(&style_xml);
    }

    xml.push_str("</w:styles>");

    Ok(xml)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_basic_style() {
        let style = MutableStyle::new("MyStyle", "My Custom Style", WdStyleType::Paragraph);
        assert_eq!(style.style_id(), "MyStyle");
        assert_eq!(style.name(), "My Custom Style");
        assert_eq!(style.style_type(), WdStyleType::Paragraph);
        assert!(style.is_custom());
    }

    #[test]
    fn test_style_formatting() {
        let mut style = MutableStyle::new("Formatted", "Formatted Style", WdStyleType::Character);
        style.set_bold(true);
        style.set_italic(true);
        style.set_underline(true);
        style.set_font_size(Some(24));
        style.set_color(Some("FF0000".to_string()));

        let xml = style.to_xml().unwrap();
        assert!(xml.contains("<w:b/>"));
        assert!(xml.contains("<w:i/>"));
        assert!(xml.contains("<w:u"));
        assert!(xml.contains(r#"w:val="24""#));
        assert!(xml.contains(r#"w:val="FF0000""#));
    }

    #[test]
    fn test_paragraph_properties() {
        let mut style = MutableStyle::new("ParaStyle", "Paragraph Style", WdStyleType::Paragraph);
        style.set_alignment(Some("center".to_string()));
        style.set_space_before(Some(240));
        style.set_space_after(Some(120));
        style.set_indent_left(Some(720));

        let xml = style.to_xml().unwrap();
        assert!(xml.contains(r#"<w:jc w:val="center"/>"#));
        assert!(xml.contains(r#"w:before="240""#));
        assert!(xml.contains(r#"w:after="120""#));
        assert!(xml.contains(r#"w:left="720""#));
    }

    #[test]
    fn test_heading_styles() {
        let h1 = MutableStyle::heading_1();
        assert_eq!(h1.style_id(), "Heading1");
        assert_eq!(h1.based_on(), Some("Normal"));

        let h2 = MutableStyle::heading_2();
        assert_eq!(h2.style_id(), "Heading2");

        let h3 = MutableStyle::heading_3();
        assert_eq!(h3.style_id(), "Heading3");
    }

    #[test]
    fn test_normal_style() {
        let normal = MutableStyle::normal();
        assert_eq!(normal.style_id(), "Normal");
        assert!(normal.is_default());
        assert!(!normal.is_custom());
    }

    #[test]
    fn test_generate_styles_xml() {
        let styles = vec![
            MutableStyle::normal(),
            MutableStyle::heading_1(),
            MutableStyle::heading_2(),
        ];

        let xml = generate_styles_xml(&styles).unwrap();
        // Debug output
        eprintln!("Generated styles XML length: {}", xml.len());
        assert!(xml.contains("<?xml version"));
        assert!(xml.contains("<w:styles"));
        assert!(xml.contains("<w:docDefaults>"));
        assert!(xml.contains("Normal"));
        assert!(xml.contains("Heading1"));
        assert!(xml.contains("Heading2"));
        assert!(xml.contains("</w:styles>"));
    }

    #[test]
    fn test_xml_escaping() {
        let mut style = MutableStyle::new("Test<>&\"'", "Name<>&\"'", WdStyleType::Paragraph);
        style.set_based_on(Some("Base<>&\"'".to_string()));

        let xml = style.to_xml().unwrap();
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&quot;"));
        assert!(xml.contains("&apos;"));
    }
}
