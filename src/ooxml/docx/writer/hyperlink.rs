use super::field::MutableField;
use super::run::MutableRun;
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

/// Elements that can appear in a hyperlink.
#[derive(Debug)]
pub(crate) enum HyperlinkElement {
    Run(MutableRun),
    Field(MutableField),
}

/// A mutable hyperlink in a document.
///
/// Hyperlinks can be added to paragraphs to create clickable links to URLs or document anchors.
#[derive(Debug)]
pub struct MutableHyperlink {
    /// Hyperlink URL (for external links)
    pub(crate) url: Option<String>,
    /// Anchor name (for internal bookmarks)
    pub(crate) anchor: Option<String>,
    /// Display text (simple hyperlink, kept for backward compatibility)
    pub(crate) text: Option<String>,
    /// Elements (runs, fields) in this hyperlink
    pub(crate) elements: Vec<HyperlinkElement>,
    /// Optional tooltip text
    pub(crate) tooltip: Option<String>,
}

impl MutableHyperlink {
    /// Create a new hyperlink to a URL.
    pub fn new(url: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            url: Some(url.into()),
            anchor: None,
            text: Some(text.into()),
            elements: Vec::new(),
            tooltip: None,
        }
    }

    /// Create a new hyperlink to an anchor (bookmark).
    pub fn new_anchor(anchor: impl Into<String>) -> Self {
        Self {
            url: None,
            anchor: Some(anchor.into()),
            text: None,
            elements: Vec::new(),
            tooltip: None,
        }
    }

    /// Add a run to the hyperlink.
    pub fn add_run(&mut self, run: MutableRun) {
        self.elements.push(HyperlinkElement::Run(run));
    }

    /// Set the tooltip text.
    pub fn set_tooltip(&mut self, tooltip: impl Into<String>) -> &mut Self {
        self.tooltip = Some(tooltip.into());
        self
    }

    /// Serialize the hyperlink to XML.
    pub(crate) fn to_xml(&self, xml: &mut String, r_id: Option<&str>) -> Result<()> {
        // Start hyperlink element
        if let Some(anchor) = &self.anchor {
            // Internal anchor hyperlink
            write!(xml, r#"<w:hyperlink w:anchor="{}">"#, escape_xml(anchor))
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        } else if let Some(rid) = r_id {
            // External URL hyperlink
            write!(xml, r#"<w:hyperlink r:id="{}">"#, rid)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        } else {
            return Err(OoxmlError::Xml(
                "Hyperlink must have either anchor or r:id".to_string(),
            ));
        }

        // If we have elements, use them
        if !self.elements.is_empty() {
            for element in &self.elements {
                match element {
                    HyperlinkElement::Run(run) => {
                        run.to_xml(xml)?;
                    },
                    HyperlinkElement::Field(field) => {
                        xml.push_str("<w:r>");
                        let field_xml = field.to_xml()?;
                        xml.push_str(&field_xml);
                        xml.push_str("</w:r>");
                    },
                }
            }
        } else if let Some(text) = &self.text {
            // Fallback to simple text (backward compatibility)
            xml.push_str("<w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>");
            write!(xml, "<w:t>{}</w:t>", escape_xml(text))
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            xml.push_str("</w:r>");
        }

        xml.push_str("</w:hyperlink>");
        Ok(())
    }
}
