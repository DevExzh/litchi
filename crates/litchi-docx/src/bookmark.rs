/// Bookmark support for reading bookmarks from Word documents.
///
/// This module provides types and methods for accessing bookmarks in Word documents.
/// Bookmarks mark locations or regions in a document.
use crate::error::{Error, Result};
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A bookmark in a Word document.
///
/// Represents a `<w:bookmarkStart>` element. Bookmarks mark specific
/// locations or regions in a document for navigation or reference.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for bookmark in doc.bookmarks()? {
///     println!("Bookmark: {} (ID: {})", bookmark.name(), bookmark.id());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Bookmark {
    /// Bookmark ID
    id: u32,
    /// Bookmark name
    name: String,
}

impl Bookmark {
    /// Create a new Bookmark.
    ///
    /// # Arguments
    ///
    /// * `id` - The bookmark ID
    /// * `name` - The bookmark name
    #[must_use]
    pub fn new(id: u32, name: String) -> Self {
        Self { id, name }
    }

    /// Get the bookmark ID.
    #[inline]
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the bookmark name.
    #[inline]
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Extract all bookmarks from document XML bytes.
    ///
    /// # Arguments
    ///
    /// * `doc_xml` - The document XML bytes
    ///
    /// # Returns
    ///
    /// A vector of bookmarks
    pub(crate) fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<Bookmark>> {
        let mut reader = Reader::from_reader(doc_xml);
        reader.config_mut().trim_text(true);

        let mut bookmarks = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e) | Event::Start(e))
                    if e.local_name().as_ref() == b"bookmarkStart" =>
                {
                    let mut id: Option<u32> = None;
                    let mut name = String::new();

                    // Parse attributes; a repeated name keeps its first value.
                    for attr in first_wins(&e).flatten() {
                        match attr.key.local_name().as_ref() {
                            b"id" => {
                                let id_str = String::from_utf8_lossy(&attr.value);
                                id = atoi_simd::parse::<u32, false, false>(id_str.as_bytes()).ok();
                            },
                            b"name" => {
                                name = String::from_utf8_lossy(&attr.value).into_owned();
                            },
                            _ => {},
                        }
                    }

                    // Skip system bookmarks (starting with _)
                    if let Some(bookmark_id) = id
                        && !name.is_empty()
                        && !name.starts_with('_')
                    {
                        bookmarks.push(Bookmark::new(bookmark_id, name));
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(bookmarks)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_bookmark_creation() {
        let bookmark = Bookmark::new(1, "Section1".to_string());
        assert_eq!(bookmark.id(), 1);
        assert_eq!(bookmark.name(), "Section1");
    }

    /// ` n00000=""` to ` n19999=""`, then 20,000 repeats of the last name: a
    /// tag quick-xml's checked iterator reads in `O(n²)` when every item is read.
    fn repeated_attributes() -> String {
        let mut attributes = String::new();
        for index in 0..20_000 {
            attributes.push_str(&format!(" n{index:05}=\"\""));
        }
        for _ in 0..20_000 {
            attributes.push_str(" n19999=\"\"");
        }
        attributes
    }

    fn bookmarks(start: &str) -> Vec<(u32, String)> {
        let xml = format!(r#"<w:document xmlns:w="urn:w"><w:body>{start}</w:body></w:document>"#);
        Bookmark::extract_from_document(xml.as_bytes())
            .unwrap()
            .into_iter()
            .map(|bookmark| (bookmark.id(), bookmark.name().to_owned()))
            .collect()
    }

    #[test]
    fn repeated_bookmark_attributes_keep_their_first_value() {
        assert_eq!(
            bookmarks(r#"<w:bookmarkStart w:id="3" w:name="single"/>"#),
            [(3, "single".to_owned())]
        );
        assert_eq!(
            bookmarks(r#"<w:bookmarkStart w:id="4" w:name="first" w:id="5" w:name="second"/>"#),
            [(4, "first".to_owned())]
        );
        let repeated = repeated_attributes();
        assert_eq!(
            bookmarks(&format!(
                r#"<w:bookmarkStart w:id="6" w:name="mark"{repeated} w:name="late" w:id="7"/>"#
            )),
            [(6, "mark".to_owned())]
        );
    }
}
