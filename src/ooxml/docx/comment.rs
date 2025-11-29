/// Comment support for reading comments from Word documents.
///
/// This module provides types and methods for accessing comments in Word documents.
/// Comments contain author information, text content, and timestamps.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A comment in a Word document.
///
/// Represents a `<w:comment>` element. Comments include author information,
/// text content, and optional date information.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for comment in doc.comments()? {
///     println!("Comment by {}: {}", comment.author(), comment.text()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Comment {
    /// The comment ID
    id: u32,
    /// Author name
    author: String,
    /// Author initials (optional)
    initials: Option<String>,
    /// Date of comment creation (optional)
    date: Option<String>,
    /// The raw XML bytes for this comment
    xml_bytes: Vec<u8>,
}

impl Comment {
    /// Create a new Comment.
    ///
    /// # Arguments
    ///
    /// * `id` - The comment ID
    /// * `author` - Author name
    /// * `initials` - Author initials
    /// * `date` - Date of comment creation
    /// * `xml_bytes` - The XML content of the comment
    pub fn new(
        id: u32,
        author: String,
        initials: Option<String>,
        date: Option<String>,
        xml_bytes: Vec<u8>,
    ) -> Self {
        Self {
            id,
            author,
            initials,
            date,
            xml_bytes,
        }
    }

    /// Get the comment ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the author name.
    #[inline]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Get the author initials.
    #[inline]
    pub fn initials(&self) -> Option<&str> {
        self.initials.as_deref()
    }

    /// Get the comment date.
    #[inline]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Get the XML bytes of this comment.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml_bytes
    }

    /// Extract all text content from this comment.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for comment in doc.comments()? {
    ///     println!("{}: {}", comment.author(), comment.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let estimated_capacity = self.xml_bytes.len() / 8;
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    result.push_str(text);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        result.shrink_to_fit();
        Ok(result)
    }

    /// Extract all comments from a comments.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The comments part
    ///
    /// # Returns
    ///
    /// A vector of comments
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Vec<Comment>> {
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut comments = Vec::new();
        let mut current_comment_xml = Vec::with_capacity(4096);
        let mut in_comment = false;
        let mut depth = 0;
        let mut current_id: Option<u32> = None;
        let mut current_author = String::new();
        let mut current_initials: Option<String> = None;
        let mut current_date: Option<String> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"comment" && !in_comment {
                        in_comment = true;
                        depth = 1;
                        current_comment_xml.clear();
                        current_id = None;
                        current_author.clear();
                        current_initials = None;
                        current_date = None;

                        // Parse attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"id" => {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_id = atoi_simd::parse::<u32>(id_str.as_bytes()).ok();
                                },
                                b"author" => {
                                    current_author =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                },
                                b"initials" => {
                                    current_initials =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                },
                                b"date" => {
                                    current_date =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                },
                                _ => {},
                            }
                        }

                        current_comment_xml.extend_from_slice(b"<w:comment>");
                    } else if in_comment {
                        depth += 1;
                        current_comment_xml.extend_from_slice(b"<");
                        current_comment_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_comment_xml.extend_from_slice(b" ");
                            current_comment_xml.extend_from_slice(attr.key.as_ref());
                            current_comment_xml.extend_from_slice(b"=\"");
                            current_comment_xml.extend_from_slice(&attr.value);
                            current_comment_xml.extend_from_slice(b"\"");
                        }
                        current_comment_xml.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_comment {
                        current_comment_xml.extend_from_slice(b"</");
                        current_comment_xml.extend_from_slice(e.name().as_ref());
                        current_comment_xml.extend_from_slice(b">");

                        if e.local_name().as_ref() == b"comment" && depth == 1 {
                            if let Some(id) = current_id {
                                comments.push(Comment::new(
                                    id,
                                    current_author.clone(),
                                    current_initials.clone(),
                                    current_date.clone(),
                                    current_comment_xml.clone(),
                                ));
                            }
                            in_comment = false;
                        } else {
                            depth -= 1;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    if in_comment {
                        current_comment_xml.extend_from_slice(b"<");
                        current_comment_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_comment_xml.extend_from_slice(b" ");
                            current_comment_xml.extend_from_slice(attr.key.as_ref());
                            current_comment_xml.extend_from_slice(b"=\"");
                            current_comment_xml.extend_from_slice(&attr.value);
                            current_comment_xml.extend_from_slice(b"\"");
                        }
                        current_comment_xml.extend_from_slice(b"/>");
                    }
                },
                Ok(Event::Text(e)) if in_comment => {
                    current_comment_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::CData(e)) if in_comment => {
                    current_comment_xml.extend_from_slice(b"<![CDATA[");
                    current_comment_xml.extend_from_slice(e.as_ref());
                    current_comment_xml.extend_from_slice(b"]]>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(comments)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_comment_creation() {
        let xml = b"<w:comment><w:p><w:r><w:t>Test comment</w:t></w:r></w:p></w:comment>";
        let comment = Comment::new(
            1,
            "John Doe".to_string(),
            Some("JD".to_string()),
            Some("2024-01-01".to_string()),
            xml.to_vec(),
        );

        assert_eq!(comment.id(), 1);
        assert_eq!(comment.author(), "John Doe");
        assert_eq!(comment.initials(), Some("JD"));
        assert_eq!(comment.date(), Some("2024-01-01"));
        assert_eq!(comment.text().unwrap(), "Test comment");
    }
}
