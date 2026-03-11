//! Hyperlink support for DOCX documents.

use super::field::MutableField;
use super::run::MutableRun;
use crate::common::xml::escape_xml;
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

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

#[cfg(feature = "fonts")]
use crate::fonts::CollectGlyphs;
#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;
#[cfg(feature = "fonts")]
use std::collections::HashMap;

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableHyperlink {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();

        // Collect from elements (runs)
        for element in &self.elements {
            if let HyperlinkElement::Run(run) = element {
                for (font, bitmap) in run.collect_glyphs() {
                    *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                }
            }
        }

        // Collect from fallback text if no elements
        if self.elements.is_empty()
            && let Some(text) = &self.text
        {
            // Hyperlink style usually defaults to Calibri in Word,
            // but here we just use the default font name for fallback text.
            let font_name = "Calibri".to_string();
            let bitmap = glyphs.entry(font_name).or_insert_with(RoaringBitmap::new);
            for c in text.chars() {
                bitmap.insert(c as u32);
            }
        }

        glyphs
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_new() {
        let link = MutableHyperlink::new("https://example.com", "Click here");
        assert_eq!(link.url, Some("https://example.com".to_string()));
        assert_eq!(link.text, Some("Click here".to_string()));
        assert!(link.anchor.is_none());
        assert!(link.elements.is_empty());
        assert!(link.tooltip.is_none());
    }

    #[test]
    fn test_hyperlink_new_anchor() {
        let link = MutableHyperlink::new_anchor("bookmark1");
        assert!(link.url.is_none());
        assert!(link.text.is_none());
        assert_eq!(link.anchor, Some("bookmark1".to_string()));
        assert!(link.elements.is_empty());
    }

    #[test]
    fn test_hyperlink_set_tooltip() {
        let mut link = MutableHyperlink::new("https://example.com", "Click here");
        link.set_tooltip("Example tooltip");
        assert_eq!(link.tooltip, Some("Example tooltip".to_string()));
    }

    #[test]
    fn test_hyperlink_add_run() {
        let mut link = MutableHyperlink::new_anchor("bookmark1");
        let run = MutableRun::new();
        link.add_run(run);
        assert_eq!(link.elements.len(), 1);
    }

    #[test]
    fn test_hyperlink_to_xml_with_anchor() {
        let link = MutableHyperlink::new_anchor("bookmark1");
        let mut xml = String::new();
        let result = link.to_xml(&mut xml, None);
        assert!(result.is_ok());
        assert!(xml.contains("<w:hyperlink"));
        assert!(xml.contains("w:anchor=\"bookmark1\""));
        assert!(xml.contains("</w:hyperlink>"));
    }

    #[test]
    fn test_hyperlink_to_xml_with_url() {
        let link = MutableHyperlink::new("https://example.com", "Click here");
        let mut xml = String::new();
        let result = link.to_xml(&mut xml, Some("rId1"));
        assert!(result.is_ok());
        assert!(xml.contains("<w:hyperlink"));
        assert!(xml.contains("r:id=\"rId1\""));
        assert!(xml.contains("Click here"));
        assert!(xml.contains("<w:rStyle w:val=\"Hyperlink\"/>"));
    }

    #[test]
    fn test_hyperlink_to_xml_without_r_id_or_anchor() {
        let link = MutableHyperlink::new("https://example.com", "Click here");
        let mut xml = String::new();
        let result = link.to_xml(&mut xml, None);
        assert!(result.is_err());
    }

    #[test]
    fn test_hyperlink_element_debug() {
        let element = HyperlinkElement::Run(MutableRun::new());
        let debug_str = format!("{:?}", element);
        assert!(debug_str.contains("Run"));
    }
}
