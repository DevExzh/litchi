/// Run types and implementation for DOCX documents.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::UnderlineStyle;
// Import section types for PageNumberFormat
use super::section::PageNumberFormat;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Run content type.
#[derive(Debug, Clone)]
pub enum RunContent {
    /// Plain text
    Text(String),
    /// Page number field
    PageNumber(PageNumberFormat),
    /// Page count field (total pages)
    PageCount,
    /// Tab character
    Tab,
    /// Page break
    PageBreak,
    /// Footnote reference
    FootnoteReference(u32),
    /// Endnote reference
    EndnoteReference(u32),
}

/// A mutable run.
///
/// Runs contain text and character formatting.
#[derive(Debug)]
pub struct MutableRun {
    /// Run content
    pub(crate) content: RunContent,
    /// Run properties
    pub(crate) properties: RunProperties,
}

impl MutableRun {
    pub(crate) fn new() -> Self {
        Self {
            content: RunContent::Text(String::new()),
            properties: RunProperties::default(),
        }
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: &str) {
        self.content = RunContent::Text(text.to_string());
    }

    /// Get the text content.
    pub fn get_text(&self) -> String {
        match &self.content {
            RunContent::Text(s) => s.clone(),
            _ => String::new(),
        }
    }

    /// Make the text bold.
    pub fn bold(&mut self, bold: bool) -> &mut Self {
        self.properties.bold = Some(bold);
        self
    }

    /// Make the text italic.
    pub fn italic(&mut self, italic: bool) -> &mut Self {
        self.properties.italic = Some(italic);
        self
    }

    /// Set underline style.
    pub fn underline(&mut self, style: UnderlineStyle) -> &mut Self {
        self.properties.underline = Some(style);
        self
    }

    /// Set font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&mut self, size: u32) -> &mut Self {
        self.properties.font_size = Some(size);
        self
    }

    /// Set font name.
    pub fn font_name(&mut self, name: &str) -> &mut Self {
        self.properties.font_name = Some(name.to_string());
        self
    }

    /// Set text color using hex RGB (e.g., "FF0000" for red).
    pub fn color(&mut self, color: &str) -> &mut Self {
        self.properties.color = Some(color.to_string());
        self
    }

    /// Set text highlight color.
    pub fn highlight(&mut self, color: &str) -> &mut Self {
        self.properties.highlight = Some(color.to_string());
        self
    }

    /// Add a line break.
    pub fn add_break(&mut self) -> &mut Self {
        self.properties.has_break = true;
        self
    }

    /// Add a page break.
    pub fn add_page_break(&mut self) -> &mut Self {
        self.content = RunContent::PageBreak;
        self
    }

    /// Add a page number field.
    pub fn add_page_number(&mut self, format: PageNumberFormat) -> &mut Self {
        self.content = RunContent::PageNumber(format);
        self
    }

    /// Add a page count field (total pages).
    pub fn add_page_count(&mut self) -> &mut Self {
        self.content = RunContent::PageCount;
        self
    }

    /// Add a tab character.
    pub fn add_tab(&mut self) -> &mut Self {
        self.content = RunContent::Tab;
        self
    }

    /// Add a footnote reference.
    pub fn add_footnote_reference(&mut self, id: u32) -> &mut Self {
        self.content = RunContent::FootnoteReference(id);
        self
    }

    /// Add an endnote reference.
    pub fn add_endnote_reference(&mut self, id: u32) -> &mut Self {
        self.content = RunContent::EndnoteReference(id);
        self
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:r>");

        // Write run properties
        if self.properties.has_properties() {
            xml.push_str("<w:rPr>");

            if let Some(bold) = self.properties.bold
                && bold
            {
                xml.push_str("<w:b/>");
            }

            if let Some(italic) = self.properties.italic
                && italic
            {
                xml.push_str("<w:i/>");
            }

            if let Some(underline_style) = self.properties.underline {
                write!(xml, "<w:u w:val=\"{}\"/>", underline_style.as_str())
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(size) = self.properties.font_size {
                write!(xml, "<w:sz w:val=\"{}\"/>", size)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(ref font_name) = self.properties.font_name {
                write!(
                    xml,
                    "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
                    escape_xml(font_name),
                    escape_xml(font_name)
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(ref color) = self.properties.color {
                write!(xml, "<w:color w:val=\"{}\"/>", color)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(ref highlight) = self.properties.highlight {
                write!(xml, "<w:highlight w:val=\"{}\"/>", highlight)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if self.properties.no_proof {
                xml.push_str("<w:noProof/>");
            }

            if self.properties.web_hidden {
                xml.push_str("<w:webHidden/>");
            }

            xml.push_str("</w:rPr>");
        }

        // Write content based on type
        match &self.content {
            RunContent::Text(text) if !text.is_empty() => {
                write!(
                    xml,
                    "<w:t xml:space=\"preserve\">{}</w:t>",
                    escape_xml(text)
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            },
            RunContent::PageNumber(format) => {
                // Field begin
                xml.push_str("<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    xml.push_str("<w:rPr>");
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    xml.push_str("</w:rPr>");
                }
                // Field instruction
                write!(
                    xml,
                    "<w:instrText xml:space=\"preserve\">PAGE \\* {}</w:instrText></w:r><w:r>",
                    format.as_str()
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                // Field separate
                xml.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    xml.push_str("<w:rPr>");
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    xml.push_str("</w:rPr>");
                }
                // Placeholder text
                xml.push_str("<w:t>1</w:t></w:r><w:r>");
                // Field end
                xml.push_str("<w:fldChar w:fldCharType=\"end\"/>");
            },
            RunContent::PageCount => {
                xml.push_str("<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    xml.push_str("<w:rPr>");
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    xml.push_str("</w:rPr>");
                }
                xml.push_str(
                    "<w:instrText xml:space=\"preserve\">NUMPAGES</w:instrText></w:r><w:r>",
                );
                xml.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    xml.push_str("<w:rPr>");
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    xml.push_str("</w:rPr>");
                }
                xml.push_str("<w:t>1</w:t></w:r><w:r>");
                xml.push_str("<w:fldChar w:fldCharType=\"end\"/>");
            },
            RunContent::Tab => {
                xml.push_str("<w:tab/>");
            },
            RunContent::PageBreak => {
                xml.push_str("<w:br w:type=\"page\"/>");
            },
            RunContent::FootnoteReference(id) => {
                write!(xml, "<w:footnoteReference w:id=\"{}\"/>", id)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            },
            RunContent::EndnoteReference(id) => {
                write!(xml, "<w:endnoteReference w:id=\"{}\"/>", id)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            },
            _ => {},
        }

        // Write line break if set
        if self.properties.has_break {
            xml.push_str("<w:br/>");
        }

        xml.push_str("</w:r>");

        Ok(())
    }
}

/// Run properties.
#[derive(Debug, Default)]
pub(crate) struct RunProperties {
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) underline: Option<UnderlineStyle>,
    pub(crate) font_size: Option<u32>,
    pub(crate) font_name: Option<String>,
    pub(crate) color: Option<String>,
    pub(crate) highlight: Option<String>,
    pub(crate) has_break: bool,
    pub(crate) no_proof: bool,
    pub(crate) web_hidden: bool,
}

impl RunProperties {
    pub(crate) fn has_properties(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.font_size.is_some()
            || self.font_name.is_some()
            || self.color.is_some()
            || self.highlight.is_some()
            || self.no_proof
            || self.web_hidden
    }
}
