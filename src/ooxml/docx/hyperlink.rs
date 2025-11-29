/// Hyperlink support for reading hyperlinks from Word documents.
///
/// This module provides types and methods for accessing hyperlinks in Word documents.
/// Hyperlinks can point to external URLs, email addresses, or internal document locations (bookmarks).
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::rel::Relationships;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A hyperlink in a Word document.
///
/// Represents a `<w:hyperlink>` element. Hyperlinks contain text and a target URL.
/// They can be external (pointing to a web URL or file) or internal (pointing to a bookmark).
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get all hyperlinks from the document
/// let hyperlinks = doc.hyperlinks()?;
/// for link in hyperlinks {
///     println!("Text: {}", link.text());
///     if let Some(url) = link.url() {
///         println!("URL: {}", url);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// The display text of the hyperlink
    text: String,
    /// The target URL (None for internal bookmarks)
    url: Option<String>,
    /// The relationship ID (rId) if external
    r_id: Option<String>,
    /// The bookmark anchor if internal
    anchor: Option<String>,
    /// Tooltip text (optional)
    tooltip: Option<String>,
}

impl Hyperlink {
    /// Create a new Hyperlink.
    ///
    /// # Arguments
    ///
    /// * `text` - The display text
    /// * `url` - The target URL (for external links)
    /// * `r_id` - The relationship ID (for external links)
    /// * `anchor` - The bookmark anchor (for internal links)
    /// * `tooltip` - Optional tooltip text
    pub fn new(
        text: String,
        url: Option<String>,
        r_id: Option<String>,
        anchor: Option<String>,
        tooltip: Option<String>,
    ) -> Self {
        Self {
            text,
            url,
            r_id,
            anchor,
            tooltip,
        }
    }

    /// Get the display text of the hyperlink.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get the target URL of the hyperlink (if external).
    #[inline]
    pub fn url(&self) -> Option<&str> {
        self.url.as_deref()
    }

    /// Get the relationship ID of the hyperlink (if external).
    #[inline]
    pub fn r_id(&self) -> Option<&str> {
        self.r_id.as_deref()
    }

    /// Get the bookmark anchor of the hyperlink (if internal).
    #[inline]
    pub fn anchor(&self) -> Option<&str> {
        self.anchor.as_deref()
    }

    /// Get the tooltip text of the hyperlink.
    #[inline]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }

    /// Check if this is an external hyperlink (has a URL).
    #[inline]
    pub fn is_external(&self) -> bool {
        self.url.is_some()
    }

    /// Check if this is an internal hyperlink (has an anchor).
    #[inline]
    pub fn is_internal(&self) -> bool {
        self.anchor.is_some()
    }

    /// Extract hyperlinks from paragraph XML bytes.
    ///
    /// # Arguments
    ///
    /// * `para_xml` - The paragraph XML bytes
    /// * `rels` - Relationships for resolving rIds to URLs
    ///
    /// # Returns
    ///
    /// A vector of hyperlinks found in the paragraph
    pub(crate) fn extract_from_paragraph(
        para_xml: &[u8],
        rels: &Relationships,
    ) -> Result<Vec<Hyperlink>> {
        let mut reader = Reader::from_reader(para_xml);
        reader.config_mut().trim_text(true);

        let mut hyperlinks = Vec::new();
        let mut in_hyperlink = false;
        let mut current_text = String::new();
        let mut current_r_id: Option<String> = None;
        let mut current_anchor: Option<String> = None;
        let mut current_tooltip: Option<String> = None;
        let mut in_text = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    match e.local_name().as_ref() {
                        b"hyperlink" => {
                            in_hyperlink = true;
                            current_text.clear();
                            current_r_id = None;
                            current_anchor = None;
                            current_tooltip = None;

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"id" => {
                                        // External link - has relationship ID
                                        current_r_id =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    b"anchor" => {
                                        // Internal link - has anchor/bookmark
                                        current_anchor =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    b"tooltip" => {
                                        current_tooltip =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    _ => {},
                                }
                            }
                        },
                        b"t" if in_hyperlink => {
                            in_text = true;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(e)) if in_hyperlink && in_text => {
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    current_text.push_str(text);
                },
                Ok(Event::End(e)) => {
                    match e.local_name().as_ref() {
                        b"hyperlink" => {
                            // End of hyperlink element - create Hyperlink object
                            let url = if let Some(ref rid) = current_r_id {
                                // Look up the URL from relationships
                                rels.get(rid).and_then(|rel| {
                                    if rel.is_external() {
                                        Some(rel.target_ref().to_string())
                                    } else {
                                        None
                                    }
                                })
                            } else {
                                None
                            };

                            hyperlinks.push(Hyperlink::new(
                                current_text.clone(),
                                url,
                                current_r_id.clone(),
                                current_anchor.clone(),
                                current_tooltip.clone(),
                            ));

                            in_hyperlink = false;
                        },
                        b"t" => {
                            in_text = false;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Empty(e)) if in_hyperlink && e.local_name().as_ref() == b"t" => {
                    // Empty text element
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(hyperlinks)
    }

    /// Extract all hyperlinks from document XML bytes.
    ///
    /// # Arguments
    ///
    /// * `doc_xml` - The document XML bytes
    /// * `rels` - Relationships for resolving rIds to URLs
    ///
    /// # Returns
    ///
    /// A vector of all hyperlinks found in the document
    pub(crate) fn extract_from_document(
        doc_xml: &[u8],
        rels: &Relationships,
    ) -> Result<Vec<Hyperlink>> {
        let mut reader = Reader::from_reader(doc_xml);
        reader.config_mut().trim_text(true);

        let mut hyperlinks = Vec::new();
        let mut in_hyperlink = false;
        let mut current_text = String::new();
        let mut current_r_id: Option<String> = None;
        let mut current_anchor: Option<String> = None;
        let mut current_tooltip: Option<String> = None;
        let mut in_text = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    match e.local_name().as_ref() {
                        b"hyperlink" => {
                            in_hyperlink = true;
                            current_text.clear();
                            current_r_id = None;
                            current_anchor = None;
                            current_tooltip = None;

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"id" => {
                                        current_r_id =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    b"anchor" => {
                                        current_anchor =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    b"tooltip" => {
                                        current_tooltip =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    },
                                    _ => {},
                                }
                            }
                        },
                        b"t" if in_hyperlink => {
                            in_text = true;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(e)) if in_hyperlink && in_text => {
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    current_text.push_str(text);
                },
                Ok(Event::End(e)) => {
                    match e.local_name().as_ref() {
                        b"hyperlink" => {
                            // End of hyperlink element
                            let url = if let Some(ref rid) = current_r_id {
                                rels.get(rid).and_then(|rel| {
                                    if rel.is_external() {
                                        Some(rel.target_ref().to_string())
                                    } else {
                                        None
                                    }
                                })
                            } else {
                                None
                            };

                            hyperlinks.push(Hyperlink::new(
                                current_text.clone(),
                                url,
                                current_r_id.clone(),
                                current_anchor.clone(),
                                current_tooltip.clone(),
                            ));

                            in_hyperlink = false;
                        },
                        b"t" => {
                            in_text = false;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(hyperlinks)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_creation() {
        let link = Hyperlink::new(
            "Click here".to_string(),
            Some("https://example.com".to_string()),
            Some("rId1".to_string()),
            None,
            Some("Example website".to_string()),
        );

        assert_eq!(link.text(), "Click here");
        assert_eq!(link.url(), Some("https://example.com"));
        assert_eq!(link.r_id(), Some("rId1"));
        assert_eq!(link.tooltip(), Some("Example website"));
        assert!(link.is_external());
        assert!(!link.is_internal());
    }

    #[test]
    fn test_internal_hyperlink() {
        let link = Hyperlink::new(
            "Go to section".to_string(),
            None,
            None,
            Some("section1".to_string()),
            None,
        );

        assert!(!link.is_external());
        assert!(link.is_internal());
        assert_eq!(link.anchor(), Some("section1"));
    }
}
