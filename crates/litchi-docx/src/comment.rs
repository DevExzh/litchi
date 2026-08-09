#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::error::{Error, Result};
/// Comment support for reading comments from Word documents.
///
/// This module provides types and methods for accessing comments in Word documents.
/// Comments contain author information, text content, and timestamps.
use crate::namespace::scan_word_element_ranges;
use crate::paragraph::extract_word_text;
use litchi_core::XmlSlice;
use litchi_opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::sync::Arc;

#[derive(Debug, Clone)]
enum CommentXmlData {
    Owned(Box<[u8]>),
    Shared(XmlSlice),
}

impl CommentXmlData {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Shared(slice) => slice.as_bytes(),
        }
    }
}

/// A comment in a Word document.
///
/// Represents a `<w:comment>` element. Comments include author information,
/// text content, and optional date information.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
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
    xml_data: CommentXmlData,
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
    #[must_use]
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
            xml_data: CommentXmlData::Owned(xml_bytes.into_boxed_slice()),
        }
    }

    fn from_arc_range(
        id: u32,
        author: String,
        initials: Option<String>,
        date: Option<String>,
        source: Arc<Vec<u8>>,
        start: u32,
        length: u32,
    ) -> Self {
        Self {
            id,
            author,
            initials,
            date,
            xml_data: CommentXmlData::Shared(XmlSlice::new(source, start, length)),
        }
    }

    /// Get the comment ID.
    #[inline]
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the author name.
    #[inline]
    #[must_use]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Get the author initials.
    #[inline]
    #[must_use]
    pub fn initials(&self) -> Option<&str> {
        self.initials.as_deref()
    }

    /// Get the comment date.
    #[inline]
    #[must_use]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Get the XML bytes of this comment.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Extract all text content from this comment.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for comment in doc.comments()? {
    ///     println!("{}: {}", comment.author(), comment.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
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
        let source = litchi_ooxml_common::mce::process_part_arc(part)?;
        let mut comments = Vec::new();
        scan_word_element_ranges(
            source.as_slice(),
            &[b"comment".as_slice()],
            |_, start, length| {
                let start_index = start as usize;
                let end_index = start_index.checked_add(length as usize).ok_or_else(|| {
                    Error::InvalidFormat("Word comment range overflow".to_string())
                })?;
                let metadata = parse_comment_metadata(&source[start_index..end_index])?;
                if let Some((id, author, initials, date)) = metadata {
                    comments.push(Comment::from_arc_range(
                        id,
                        author,
                        initials,
                        date,
                        Arc::clone(&source),
                        start,
                        length,
                    ));
                }
                Ok(())
            },
        )?;
        Ok(comments)
    }
}

type CommentMetadata = (u32, String, Option<String>, Option<String>);

fn parse_comment_metadata(xml_bytes: &[u8]) -> Result<Option<CommentMetadata>> {
    let mut reader = Reader::from_reader(xml_bytes);
    let element = loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => break element,
            Ok(Event::Eof) => {
                return Err(Error::InvalidFormat(
                    "missing Word comment element".to_string(),
                ));
            },
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    };

    let mut id = None;
    let mut author = String::new();
    let mut initials = None;
    let mut date = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"id" => {
                id = atoi_simd::parse::<u32, false, false>(attribute.value.as_ref()).ok();
            },
            b"author" => author = decode_comment_attribute(attribute.value.as_ref())?,
            b"initials" => {
                initials = Some(decode_comment_attribute(attribute.value.as_ref())?);
            },
            b"date" => date = Some(decode_comment_attribute(attribute.value.as_ref())?),
            _ => {},
        }
    }
    Ok(id.map(|id| (id, author, initials, date)))
}

fn decode_comment_attribute(value: &[u8]) -> Result<String> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    quick_xml::escape::unescape(value)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| Error::Xml(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn comments_part(xml: &[u8]) -> BlobPart {
        BlobPart::new(
            PackURI::new("/word/comments.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
                .to_string(),
            xml.to_vec(),
        )
    }

    #[test]
    fn test_comment_creation() {
        let xml = b"<w:comment><w:p><w:r><w:t xml:space=\"preserve\"> Test &amp; comment </w:t><w:tab/><w:br/></w:r></w:p></w:comment>";
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
        assert_eq!(comment.text().unwrap(), " Test & comment \t\n");
    }

    #[test]
    fn extracts_aliased_comment_slices_and_decodes_metadata() {
        let xml = br#"<c:comments xmlns:c="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
            <false:comment false:id="9"><false:p><false:r><false:t>ignored</false:t></false:r></false:p></false:comment>
            <c:comment c:id="1" c:author="A &amp; B" c:initials="AB" c:date="2026-07-14T00:00:00Z"><c:p><c:r><c:t><![CDATA[x < y]]></c:t></c:r></c:p></c:comment>
            <c:comment c:id="2"/>
        </c:comments>"#;
        let part = comments_part(xml);

        let comments = Comment::extract_from_part(&part).unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].id(), 1);
        assert_eq!(comments[0].author(), "A & B");
        assert_eq!(comments[0].initials(), Some("AB"));
        assert_eq!(comments[0].date(), Some("2026-07-14T00:00:00Z"));
        assert_eq!(comments[0].text().unwrap(), "x < y");
        assert!(
            comments[0]
                .xml_bytes()
                .starts_with(br#"<c:comment c:id="1""#)
        );
        assert_eq!(comments[1].id(), 2);
        assert_eq!(comments[1].author(), "");
        assert_eq!(comments[1].text().unwrap(), "");
    }

    #[test]
    fn rejects_unterminated_comment_slices() {
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"><w:p/>"#;
        let part = comments_part(xml);
        assert!(Comment::extract_from_part(&part).is_err());
    }
}
