/// Hyperlink support for DOCX documents.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// A mutable hyperlink in a document.
///
/// Hyperlinks can be added to paragraphs to create clickable links to URLs.
#[derive(Debug)]
pub struct MutableHyperlink {
    /// Hyperlink URL
    pub(crate) url: String,
    /// Display text
    pub(crate) text: String,
    /// Optional tooltip text
    pub(crate) tooltip: Option<String>,
}

impl MutableHyperlink {
    /// Create a new hyperlink.
    pub fn new(url: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            url: url.into(),
            text: text.into(),
            tooltip: None,
        }
    }

    /// Set the tooltip text.
    pub fn set_tooltip(&mut self, tooltip: impl Into<String>) -> &mut Self {
        self.tooltip = Some(tooltip.into());
        self
    }

    /// Serialize the hyperlink to XML.
    pub(crate) fn to_xml(&self, xml: &mut String, r_id: &str) -> Result<()> {
        write!(xml, r#"<w:hyperlink r:id="{}">"#, r_id)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("<w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>");
        write!(xml, "<w:t>{}</w:t>", escape_xml(&self.text))
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</w:r></w:hyperlink>");
        Ok(())
    }
}
