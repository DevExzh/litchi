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
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
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

    fn self_contained_xml(&self) -> Result<Vec<u8>> {
        match self {
            Self::Owned(bytes) => Ok(bytes.to_vec()),
            Self::Shared(slice) => crate::namespace::self_contained_element_xml(
                slice.arc().as_slice(),
                slice.start(),
                u32::try_from(slice.len()).map_err(|_source_error| {
                    Error::InvalidFormat("Word comment range exceeds u32".to_string())
                })?,
            ),
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
    ///
    /// The bytes are an XML fragment whose namespace declarations are those in
    /// scope in `word/comments.xml`, which since change 0653 are not repeated
    /// on the fragment's own root element. Use
    /// [`self_contained_xml`](Self::self_contained_xml) to parse it on its own.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// The `w14:paraId` of the comment's last paragraph (the last
    /// `w:p` child of `w:comment`), by which Word's modern-comment parts
    /// (`commentsExtended`, `commentsIds`) key the comment ([MS-DOCX]
    /// `CT_CommentEx`; record 0345). For a one-paragraph comment, the
    /// common case and every comment OfficeCLI writes, it is the first
    /// paragraph's. `None` when that paragraph has none.
    #[must_use]
    pub fn paragraph_id(&self) -> Option<u32> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        let mut depth = 0usize;
        let mut last = None;
        loop {
            let (element, empty) = match reader.read_event() {
                Ok(Event::Start(element)) => (element, false),
                Ok(Event::Empty(element)) => (element, true),
                Ok(Event::End(_)) => {
                    depth = depth.saturating_sub(1);
                    continue;
                },
                Ok(Event::Eof) | Err(_) => return last.flatten(),
                _ => continue,
            };
            if depth == 1 && element.local_name().as_ref() == b"p" {
                last = Some(
                    element
                        .unchecked_attributes()
                        .filter_map(std::result::Result::ok)
                        .find(|attribute| attribute.key.local_name().as_ref() == b"paraId")
                        .and_then(|attribute| {
                            u32::from_str_radix(std::str::from_utf8(&attribute.value).ok()?, 16)
                                .ok()
                        }),
                );
            }
            if !empty {
                depth += 1;
            }
        }
    }

    /// This comment's XML with every namespace declaration it inherits from
    /// `word/comments.xml` re-declared on its own `w:comment`, so that it
    /// parses standalone.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained range or the part around it cannot be
    /// read, or when a bound on the namespace scan is reached.
    pub fn self_contained_xml(&self) -> Result<Vec<u8>> {
        self.xml_data.self_contained_xml()
    }

    /// The comment's direct block children, indexed with the scanner the
    /// document body uses (see [`crate::Blocks`]).
    ///
    /// The comment's XML is first made self-contained, so the blocks parse
    /// without `word/comments.xml`.
    ///
    /// # Errors
    ///
    /// Returns an error when the comment's XML is malformed or a bound on the
    /// index is reached.
    pub fn blocks(&self) -> Result<crate::Blocks> {
        let xml = self.self_contained_xml()?;
        let length = u32::try_from(xml.len()).map_err(|_source_error| {
            Error::InvalidFormat("Word comment exceeds u32".to_string())
        })?;
        crate::Blocks::index(
            Arc::new(xml),
            0,
            length,
            &[],
            crate::document::ContainerRoot::COMMENT,
            crate::parts::memo::RawPositions::Unproven,
        )
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
        Self::extract_from_payload(&part.blob_arc())
    }

    /// Extract the comments of a comments-part payload.
    pub(crate) fn extract_from_payload(payload: &Arc<Vec<u8>>) -> Result<Vec<Comment>> {
        // Preserve modeled Word extensions while selecting active MCE
        // content. Native Word marks w14 ignorable, but its paragraph IDs
        // are the keys linking comments to reply and durable-ID metadata.
        let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
        capabilities.understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
        capabilities.understand_namespace(crate::revision::WORD_2023_DATE_UTC_NAMESPACE);
        let source = match litchi_ooxml_common::mce::process_markup_compatibility(
            payload.as_slice(),
            &capabilities,
            &litchi_ooxml_common::mce::Limits::default(),
        )?
        .xml
        {
            std::borrow::Cow::Borrowed(_) => Arc::clone(payload),
            std::borrow::Cow::Owned(value) => Arc::new(value),
        };
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
    for attribute in element.checked_attributes() {
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
    fn word_2023_revision_utc_is_retained_in_comment_mce_projection() {
        let word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let mce = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        let date_utc = crate::revision::WORD_2023_DATE_UTC_NAMESPACE;
        let xml = format!(
            r#"<w:comments xmlns:w="{word}" xmlns:mc="{mce}" xmlns:du="{date_utc}" mc:Ignorable="du"><w:comment w:id="1" w:author="Alice"><w:p><w:ins w:id="7" w:author="Bob" du:dateUtc="2026-07-17T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:comment></w:comments>"#
        );
        let comments = Comment::extract_from_part(&comments_part(xml.as_bytes())).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text().unwrap(), "added");
        assert!(
            comments[0]
                .xml_bytes()
                .windows(b"du:dateUtc".len())
                .any(|window| window == b"du:dateUtc")
        );
    }

    #[test]
    fn ignorable_word_2010_paragraph_ids_survive_comment_projection() {
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ids="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="ids"><w:comment w:id="1"><w:p ids:paraId="12345678"/><w:p ids:paraId="9ABCDEF0"/></w:comment></w:comments>"#;
        let comments = Comment::extract_from_part(&comments_part(xml)).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].paragraph_id(), Some(0x9ABCDEF0));
    }

    #[test]
    fn rejects_unterminated_comment_slices() {
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"><w:p/>"#;
        let part = comments_part(xml);
        assert!(Comment::extract_from_part(&part).is_err());
    }
}
