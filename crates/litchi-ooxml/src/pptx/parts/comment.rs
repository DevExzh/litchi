//! Comment parts for PowerPoint presentations.
//!
//! This module provides types for working with comments in PPTX files.

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use litchi_core::xml::escape_xml;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

/// A comment author.
///
/// Represents information about a comment author from the comment authors part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentAuthor {
    /// Author ID
    pub id: u32,
    /// Author name
    pub name: String,
    /// Author initials
    pub initials: String,
}

impl CommentAuthor {
    /// Create a new comment author.
    pub fn new(id: u32, name: impl Into<String>, initials: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            initials: initials.into(),
        }
    }

    /// Generate XML element for this author.
    pub fn to_xml(&self) -> String {
        format!(
            r#"<p:cmAuthor id="{}" name="{}" initials="{}" lastIdx="0" clrIdx="{}"/>"#,
            self.id,
            escape_xml(&self.name),
            escape_xml(&self.initials),
            self.id % 6 // Color index cycles through 0-5
        )
    }
}

/// A comment in a presentation.
///
/// Comments are annotations attached to specific positions on slides.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment author ID
    pub author_id: u32,
    /// Comment text
    pub text: String,
    /// X position in EMUs
    pub x: i64,
    /// Y position in EMUs
    pub y: i64,
    /// Comment date/time as string (ISO 8601 format)
    pub datetime: Option<String>,
    /// Comment index
    pub index: Option<u32>,
}

impl Comment {
    /// Create a new comment.
    pub fn new(author_id: u32, text: impl Into<String>, x: i64, y: i64) -> Self {
        Self {
            author_id,
            text: text.into(),
            x,
            y,
            datetime: None,
            index: None,
        }
    }

    /// Create a new comment with datetime.
    pub fn with_datetime(mut self, datetime: impl Into<String>) -> Self {
        self.datetime = Some(datetime.into());
        self
    }

    /// Create a new comment with index.
    pub fn with_index(mut self, index: u32) -> Self {
        self.index = Some(index);
        self
    }

    /// Generate XML element for this comment.
    pub fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(256);

        xml.push_str(r#"<p:cm authorId=""#);
        xml.push_str(&self.author_id.to_string());
        xml.push('"');

        if let Some(ref dt) = self.datetime {
            xml.push_str(r#" dt=""#);
            xml.push_str(&escape_xml(dt));
            xml.push('"');
        }

        if let Some(idx) = self.index {
            xml.push_str(r#" idx=""#);
            xml.push_str(&idx.to_string());
            xml.push('"');
        }

        xml.push('>');

        // Position
        xml.push_str(r#"<p:pos x=""#);
        xml.push_str(&self.x.to_string());
        xml.push_str(r#"" y=""#);
        xml.push_str(&self.y.to_string());
        xml.push_str(r#""/>"#);

        // Text
        xml.push_str("<p:text>");
        xml.push_str(&escape_xml(&self.text));
        xml.push_str("</p:text>");

        xml.push_str("</p:cm>");

        xml
    }
}

/// Generate comments part XML.
///
/// Creates the complete `/ppt/comments/commentN.xml` content.
pub fn generate_comments_xml(comments: &[Comment]) -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    );

    for comment in comments {
        xml.push_str(&comment.to_xml());
    }

    xml.push_str("</p:cmLst>");

    xml
}

/// Generate comment authors part XML.
///
/// Creates the complete `/ppt/commentAuthors.xml` content.
pub fn generate_comment_authors_xml(authors: &[CommentAuthor]) -> String {
    let mut xml = String::with_capacity(512);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    );

    for author in authors {
        xml.push_str(&author.to_xml());
    }

    xml.push_str("</p:cmAuthorLst>");

    xml
}

/// Comments part - contains comments for a slide.
///
/// Corresponds to `/ppt/comments/commentN.xml` in the package.
pub struct CommentsPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> CommentsPart<'a> {
    /// Create a CommentsPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the comments.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return all comments from this part.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let comments_part = CommentsPart::from_part(part)?;
    /// let comments = comments_part.comments()?;
    /// for comment in comments {
    ///     println!("Comment: {}", comment.text);
    /// }
    /// ```
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let xml = litchi_ooxml_common::mce::process_ooxml(self.xml_bytes())?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut comments = Vec::new();
        let mut pending: Option<PendingComment> = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("comment XML nesting is too deep".to_string())
                    })?;
                    if is_presentationml_name(&namespace, element.name(), b"cm") {
                        if pending.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested PowerPoint comments are invalid".to_string(),
                            ));
                        }
                        pending = Some(parse_comment_start(&element, decoder, depth)?);
                    } else if let Some(comment) = pending.as_mut()
                        && depth == comment.depth + 1
                    {
                        if is_presentationml_name(&namespace, element.name(), b"pos") {
                            comment.set_position(&element, decoder)?;
                        } else if is_presentationml_name(&namespace, element.name(), b"text") {
                            comment.start_text(depth)?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("comment XML nesting is too deep".to_string())
                    })?;
                    if is_presentationml_name(&namespace, element.name(), b"cm") {
                        if pending.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested PowerPoint comments are invalid".to_string(),
                            ));
                        }
                        let comment =
                            parse_comment_start(&element, decoder, child_depth)?.finish()?;
                        comments.push(comment);
                    } else if let Some(comment) = pending.as_mut()
                        && child_depth == comment.depth + 1
                    {
                        if is_presentationml_name(&namespace, element.name(), b"pos") {
                            comment.set_position(&element, decoder)?;
                        } else if is_presentationml_name(&namespace, element.name(), b"text") {
                            comment.mark_empty_text()?;
                        }
                    }
                },
                Event::Text(text) if pending.as_ref().is_some_and(PendingComment::in_text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    pending
                        .as_mut()
                        .ok_or_else(|| {
                            OoxmlError::InvalidFormat("missing PowerPoint comment".to_string())
                        })?
                        .text
                        .push_str(
                            &quick_xml::escape::unescape(&decoded)
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                        );
                },
                Event::CData(text) if pending.as_ref().is_some_and(PendingComment::in_text) => {
                    pending
                        .as_mut()
                        .ok_or_else(|| {
                            OoxmlError::InvalidFormat("missing PowerPoint comment".to_string())
                        })?
                        .text
                        .push_str(
                            &text
                                .xml_content(XmlVersion::Explicit1_0)
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                        );
                },
                Event::GeneralRef(reference)
                    if pending.as_ref().is_some_and(PendingComment::in_text) =>
                {
                    pending
                        .as_mut()
                        .ok_or_else(|| {
                            OoxmlError::InvalidFormat("missing PowerPoint comment".to_string())
                        })?
                        .text
                        .push_str(&decode_xml_reference(&reference)?);
                },
                Event::End(element) => {
                    if let Some(comment) = pending.as_mut()
                        && comment.text_depth == Some(depth)
                        && is_presentationml_name(&namespace, element.name(), b"text")
                    {
                        comment.text_depth = None;
                    } else if pending.as_ref().is_some_and(|comment| {
                        comment.depth == depth
                            && is_presentationml_name(&namespace, element.name(), b"cm")
                    }) {
                        let comment = pending
                            .take()
                            .ok_or_else(|| {
                                OoxmlError::InvalidFormat("missing PowerPoint comment".to_string())
                            })?
                            .finish()?;
                        comments.push(comment);
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid comment XML nesting".to_string())
                    })?;
                },
                Event::Eof if pending.is_some() || depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated PowerPoint comments XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(comments)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

struct PendingComment {
    depth: usize,
    author_id: u32,
    text: String,
    position: Option<(i64, i64)>,
    datetime: Option<String>,
    index: Option<u32>,
    text_seen: bool,
    text_depth: Option<usize>,
}

impl PendingComment {
    fn set_position(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.position.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "duplicate PowerPoint comment position".to_string(),
            ));
        }
        self.position = Some((
            required_i64_attribute(element, b"x", decoder, "comment x coordinate")?,
            required_i64_attribute(element, b"y", decoder, "comment y coordinate")?,
        ));
        Ok(())
    }

    fn start_text(&mut self, depth: usize) -> Result<()> {
        if self.text_seen {
            return Err(OoxmlError::InvalidFormat(
                "duplicate PowerPoint comment text".to_string(),
            ));
        }
        self.text_seen = true;
        self.text_depth = Some(depth);
        Ok(())
    }

    fn mark_empty_text(&mut self) -> Result<()> {
        self.start_text(usize::MAX)?;
        self.text_depth = None;
        Ok(())
    }

    fn in_text(&self) -> bool {
        self.text_depth.is_some()
    }

    fn finish(self) -> Result<Comment> {
        let (x, y) = self.position.ok_or_else(|| {
            OoxmlError::InvalidFormat("PowerPoint comment is missing its position".to_string())
        })?;
        if !self.text_seen {
            return Err(OoxmlError::InvalidFormat(
                "PowerPoint comment is missing its text element".to_string(),
            ));
        }
        Ok(Comment {
            author_id: self.author_id,
            text: self.text,
            x,
            y,
            datetime: self.datetime,
            index: self.index,
        })
    }
}

fn parse_comment_start(
    element: &BytesStart<'_>,
    decoder: Decoder,
    depth: usize,
) -> Result<PendingComment> {
    Ok(PendingComment {
        depth,
        author_id: required_u32_attribute(element, b"authorId", decoder, "comment author ID")?,
        text: String::new(),
        position: None,
        datetime: unqualified_attribute_value(element, b"dt", decoder)?,
        index: optional_u32_attribute(element, b"idx", decoder, "comment index")?,
        text_seen: false,
        text_depth: None,
    })
}

fn required_u32_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32_attribute(element, name, decoder, description)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))
}

fn optional_u32_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value.parse::<u32>().map_err(|_| {
                OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'"))
            })
        })
        .transpose()
}

fn required_i64_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))?;
    value
        .parse::<i64>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))
}

/// Comment authors part - contains author information.
///
/// Corresponds to `/ppt/commentAuthors.xml` in the package.
pub struct CommentAuthorsPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> CommentAuthorsPart<'a> {
    /// Create a CommentAuthorsPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the comment authors.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return all comment authors from this part.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let authors_part = CommentAuthorsPart::from_part(part)?;
    /// let authors = authors_part.authors()?;
    /// for author in authors {
    ///     println!("Author: {}", author.name);
    /// }
    /// ```
    pub fn authors(&self) -> Result<Vec<CommentAuthor>> {
        let xml = litchi_ooxml_common::mce::process_ooxml(self.xml_bytes())?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut authors = Vec::new();
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "comment-author XML nesting is too deep".to_string(),
                        )
                    })?;
                    if is_presentationml_name(&namespace, element.name(), b"cmAuthor") {
                        push_author(&mut authors, parse_author(&element, decoder)?)?;
                    }
                },
                Event::Empty(element)
                    if is_presentationml_name(&namespace, element.name(), b"cmAuthor") =>
                {
                    push_author(&mut authors, parse_author(&element, decoder)?)?;
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid comment-author XML nesting".to_string())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated comment-author XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(authors)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

fn parse_author(element: &BytesStart<'_>, decoder: Decoder) -> Result<CommentAuthor> {
    let id = required_u32_attribute(element, b"id", decoder, "comment-author ID")?;
    let name = unqualified_attribute_value(element, b"name", decoder)?.ok_or_else(|| {
        OoxmlError::InvalidFormat("missing comment-author name attribute".to_string())
    })?;
    let initials =
        unqualified_attribute_value(element, b"initials", decoder)?.ok_or_else(|| {
            OoxmlError::InvalidFormat("missing comment-author initials attribute".to_string())
        })?;
    Ok(CommentAuthor { id, name, initials })
}

fn push_author(authors: &mut Vec<CommentAuthor>, author: CommentAuthor) -> Result<()> {
    if authors.iter().any(|existing| existing.id == author.id) {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate PowerPoint comment-author ID {}",
            author.id
        )));
    }
    authors.push(author);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn part(path: &str, xml: impl Into<Vec<u8>>) -> BlobPart {
        BlobPart::new(
            PackURI::new(path).unwrap(),
            "application/xml".to_string(),
            xml.into(),
        )
    }

    #[test]
    fn test_comment_author_new() {
        let author = CommentAuthor::new(1, "John Doe", "JD");
        assert_eq!(author.id, 1);
        assert_eq!(author.name, "John Doe");
        assert_eq!(author.initials, "JD");
    }

    #[test]
    fn test_comment_author_to_xml() {
        let author = CommentAuthor::new(5, "Jane Smith", "JS");
        let xml = author.to_xml();
        assert!(xml.contains("id=\"5\""));
        assert!(xml.contains("name=\"Jane Smith\""));
        assert!(xml.contains("initials=\"JS\""));
        assert!(xml.contains("clrIdx=\"5\"")); // 5 % 6 = 5
    }

    #[test]
    fn test_comment_author_xml_escaping() {
        let author = CommentAuthor::new(1, "John <Doe>", "J&D");
        let xml = author.to_xml();
        assert!(xml.contains("name=\"John &lt;Doe&gt;\""));
        assert!(xml.contains("initials=\"J&amp;D\""));
    }

    #[test]
    fn test_comment_author_clone() {
        let author = CommentAuthor::new(1, "Test", "T");
        let cloned = author.clone();
        assert_eq!(cloned.id, author.id);
        assert_eq!(cloned.name, author.name);
        assert_eq!(cloned.initials, author.initials);
    }

    #[test]
    fn test_comment_author_debug() {
        let author = CommentAuthor::new(1, "Test", "T");
        let debug = format!("{:?}", author);
        assert!(debug.contains("CommentAuthor"));
        assert!(debug.contains("Test"));
    }

    #[test]
    fn test_comment_author_equality() {
        let a1 = CommentAuthor::new(1, "Test", "T");
        let a2 = CommentAuthor::new(1, "Test", "T");
        let a3 = CommentAuthor::new(2, "Other", "O");
        assert_eq!(a1, a2);
        assert_ne!(a1, a3);
    }

    #[test]
    fn test_comment_new() {
        let comment = Comment::new(1, "Hello!", 100, 200);
        assert_eq!(comment.author_id, 1);
        assert_eq!(comment.text, "Hello!");
        assert_eq!(comment.x, 100);
        assert_eq!(comment.y, 200);
        assert_eq!(comment.datetime, None);
        assert_eq!(comment.index, None);
    }

    #[test]
    fn test_comment_with_datetime() {
        let comment = Comment::new(1, "Test", 0, 0).with_datetime("2025-01-15T10:30:00");
        assert_eq!(comment.datetime, Some("2025-01-15T10:30:00".to_string()));
    }

    #[test]
    fn test_comment_with_index() {
        let comment = Comment::new(1, "Test", 0, 0).with_index(5);
        assert_eq!(comment.index, Some(5));
    }

    #[test]
    fn test_comment_to_xml_basic() {
        let comment = Comment::new(1, "Hello!", 100, 200);
        let xml = comment.to_xml();
        assert!(xml.contains("authorId=\"1\""));
        assert!(xml.contains("<p:text>Hello!</p:text>"));
        assert!(xml.contains("x=\"100\""));
        assert!(xml.contains("y=\"200\""));
    }

    #[test]
    fn test_comment_to_xml_with_all_fields() {
        let comment = Comment::new(2, "Test comment", 500, 600)
            .with_datetime("2025-03-13T14:30:00")
            .with_index(3);
        let xml = comment.to_xml();
        assert!(xml.contains("authorId=\"2\""));
        assert!(xml.contains("dt=\"2025-03-13T14:30:00\""));
        assert!(xml.contains("idx=\"3\""));
        assert!(xml.contains("x=\"500\""));
        assert!(xml.contains("y=\"600\""));
        assert!(xml.contains("<p:text>Test comment</p:text>"));
    }

    #[test]
    fn test_comment_to_xml_escaping() {
        let comment = Comment::new(1, "<script>alert('xss')</script>", 0, 0);
        let xml = comment.to_xml();
        assert!(xml.contains("&lt;script&gt;"));
        assert!(!xml.contains("<script>"));
    }

    #[test]
    fn test_comment_clone() {
        let comment = Comment::new(1, "Test", 100, 200)
            .with_datetime("2025-01-01")
            .with_index(1);
        let cloned = comment.clone();
        assert_eq!(cloned.author_id, comment.author_id);
        assert_eq!(cloned.text, comment.text);
        assert_eq!(cloned.x, comment.x);
        assert_eq!(cloned.y, comment.y);
        assert_eq!(cloned.datetime, comment.datetime);
        assert_eq!(cloned.index, comment.index);
    }

    #[test]
    fn test_comment_debug() {
        let comment = Comment::new(1, "Test", 100, 200);
        let debug = format!("{:?}", comment);
        assert!(debug.contains("Comment"));
        assert!(debug.contains("Test"));
    }

    #[test]
    fn test_generate_comments_xml_empty() {
        let xml = generate_comments_xml(&[]);
        assert!(xml.contains("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
        assert!(xml.contains("<p:cmLst"));
        assert!(
            xml.contains("xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"")
        );
        assert!(xml.contains("</p:cmLst>"));
    }

    #[test]
    fn test_generate_comments_xml_with_comments() {
        let comments = vec![
            Comment::new(1, "First comment", 100, 200),
            Comment::new(2, "Second comment", 300, 400),
        ];
        let xml = generate_comments_xml(&comments);
        assert!(xml.contains("<p:cmLst"));
        assert!(xml.contains("authorId=\"1\""));
        assert!(xml.contains("authorId=\"2\""));
        assert!(xml.contains("<p:text>First comment</p:text>"));
        assert!(xml.contains("<p:text>Second comment</p:text>"));
        assert!(xml.contains("</p:cmLst>"));
    }

    #[test]
    fn test_generate_comment_authors_xml_empty() {
        let xml = generate_comment_authors_xml(&[]);
        assert!(xml.contains("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
        assert!(xml.contains("<p:cmAuthorLst"));
        assert!(xml.contains("</p:cmAuthorLst>"));
    }

    #[test]
    fn test_generate_comment_authors_xml_with_authors() {
        let authors = vec![
            CommentAuthor::new(1, "Alice", "A"),
            CommentAuthor::new(2, "Bob", "B"),
        ];
        let xml = generate_comment_authors_xml(&authors);
        assert!(xml.contains("<p:cmAuthorLst"));
        assert!(xml.contains("id=\"1\""));
        assert!(xml.contains("id=\"2\""));
        assert!(xml.contains("name=\"Alice\""));
        assert!(xml.contains("name=\"Bob\""));
        assert!(xml.contains("</p:cmAuthorLst>"));
    }

    #[test]
    fn comments_resolve_namespaces_and_decode_content() {
        let xml = format!(
            r#"<q:cmLst xmlns:q="{P}" xmlns:f="urn:foreign">
                <f:cm authorId="9"><f:pos x="9" y="9"/><f:text>Spoof</f:text></f:cm>
                <q:cm authorId="2" dt="2026-07-14T10:30:00&amp;08:00" idx="3">
                    <q:pos x="-10" y="20"/><q:text>A &amp; <![CDATA[B < C]]></q:text>
                </q:cm><q:cm authorId="4"><q:pos x="0" y="0"/><q:text/></q:cm>
            </q:cmLst>"#
        );
        let blob = part("/ppt/comments/comment1.xml", xml);
        let comments = CommentsPart::from_part(&blob).unwrap().comments().unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].author_id, 2);
        assert_eq!(comments[0].text, "A & B < C");
        assert_eq!((comments[0].x, comments[0].y), (-10, 20));
        assert_eq!(
            comments[0].datetime.as_deref(),
            Some("2026-07-14T10:30:00&08:00")
        );
        assert_eq!(comments[0].index, Some(3));
        assert_eq!(comments[1].text, "");
    }

    #[test]
    fn generated_comments_round_trip() {
        let expected = vec![
            Comment::new(1, "A & B < C", 100, 200)
                .with_datetime("2026-07-14T10:30:00+08:00")
                .with_index(7),
        ];
        let xml = generate_comments_xml(&expected);
        let blob = part("/ppt/comments/comment1.xml", xml);
        let parsed = CommentsPart::from_part(&blob).unwrap().comments().unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].author_id, expected[0].author_id);
        assert_eq!(parsed[0].text, expected[0].text);
        assert_eq!((parsed[0].x, parsed[0].y), (100, 200));
        assert_eq!(parsed[0].datetime, expected[0].datetime);
        assert_eq!(parsed[0].index, expected[0].index);
    }

    #[test]
    fn comment_authors_resolve_strict_aliases_and_decode_attributes() {
        let xml = r#"<x:cmAuthorLst xmlns:x="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:f="urn:foreign"><f:cmAuthor id="0" name="Spoof" initials="S"/>
            <x:cmAuthor id="5" name="A &amp; B" initials="A&lt;B"/></x:cmAuthorLst>"#;
        let blob = part("/ppt/commentAuthors.xml", xml);
        let authors = CommentAuthorsPart::from_part(&blob)
            .unwrap()
            .authors()
            .unwrap();
        assert_eq!(authors, [CommentAuthor::new(5, "A & B", "A<B")]);
    }

    #[test]
    fn malformed_comments_and_authors_are_rejected() {
        let invalid_comments = [
            format!(
                r#"<p:cmLst xmlns:p="{P}"><p:cm><p:pos x="0" y="0"/><p:text/></p:cm></p:cmLst>"#
            ),
            format!(
                r#"<p:cmLst xmlns:p="{P}"><p:cm authorId="x"><p:pos x="0" y="0"/><p:text/></p:cm></p:cmLst>"#
            ),
            format!(
                r#"<p:cmLst xmlns:p="{P}"><p:cm authorId="1"><p:pos x="x" y="0"/><p:text/></p:cm></p:cmLst>"#
            ),
            format!(r#"<p:cmLst xmlns:p="{P}"><p:cm authorId="1"><p:text/></p:cm></p:cmLst>"#),
            format!(r#"<p:cmLst xmlns:p="{P}"><p:cm authorId="1"><p:pos x="0" y="0"/>"#),
        ];
        for xml in invalid_comments {
            let blob = part("/ppt/comments/comment1.xml", xml);
            assert!(CommentsPart::from_part(&blob).unwrap().comments().is_err());
        }

        let invalid_authors = [
            format!(
                r#"<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor name="A" initials="A"/></p:cmAuthorLst>"#
            ),
            format!(
                r#"<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor id="1" name="A" initials="A"/><p:cmAuthor id="1" name="B" initials="B"/></p:cmAuthorLst>"#
            ),
            format!(r#"<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor id="1" name="A" initials="A">"#),
        ];
        for xml in invalid_authors {
            let blob = part("/ppt/commentAuthors.xml", xml);
            assert!(
                CommentAuthorsPart::from_part(&blob)
                    .unwrap()
                    .authors()
                    .is_err()
            );
        }
    }
}
