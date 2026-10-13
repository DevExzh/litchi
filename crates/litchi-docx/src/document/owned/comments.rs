//! Comment authoring and removal staged with a main-document edit (record
//! 0251c).
//!
//! A comment lives in the comments part and is anchored in the main
//! document by a range and a reference run. Adding one stages the part (or
//! its replacement) as a graph change and the anchor as an element
//! replacement, so one commit publishes both. Removal strips the comment's
//! entry, its markers and its reference runs, and the matching entries of
//! the modern-comment parts.

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::Event;

use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use crate::Package;

const COMMENTS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const STRICT_COMMENTS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
/// The modern-comment parts that index comments by their paragraphs.
const COMMENTS_EXTENDED_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_IDS_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const COMMENTS_EXTENSIBLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";

/// A comment to add: its author, optional initials and date, and its text.
///
/// ```
/// use litchi_docx::document::NewComment;
///
/// let comment = NewComment::new("Ada", "Check this figure").with_initials("AL");
/// assert_eq!(comment.author(), "Ada");
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewComment {
    author: String,
    initials: Option<String>,
    date: Option<String>,
    text: String,
}

impl NewComment {
    /// A comment by `author` holding one paragraph of `text`.
    #[must_use]
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            initials: None,
            date: None,
            text: text.into(),
        }
    }

    /// Set the author's initials.
    #[must_use]
    pub fn with_initials(mut self, initials: impl Into<String>) -> Self {
        self.initials = Some(initials.into());
        self
    }

    /// Set the comment's date, an `xsd:dateTime` written as given. No date
    /// is read from a clock.
    #[must_use]
    pub fn with_date(mut self, date: impl Into<String>) -> Self {
        self.date = Some(date.into());
        self
    }

    /// The author.
    #[must_use]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// The text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

impl Edit<Package> {
    /// Add `comment`, anchored to the direct runs `range` (inclusive) of the
    /// paragraph at `paragraph`, or to all its content for `None`, and
    /// return the comment's identifier.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewComment};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Reviewed text");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let id = edit.add_comment_at(
    ///     &ElementPath::new([ElementStep::Paragraph(Position::new(0))]),
    ///     None,
    ///     &NewComment::new("Ada", "Looks good"),
    /// )?;
    /// let next = edit.commit()?.into_package();
    /// let comments = next.document()?.comments()?;
    /// assert_eq!(comments[0].id(), id);
    /// assert_eq!(comments[0].text()?, "Looks good");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing paragraph or run, invalid text,
    /// a malformed comments part, or a resource-limit error.
    pub fn add_comment_at(
        &mut self,
        paragraph: &ElementPath,
        range: Option<(Position, Position)>,
        comment: &NewComment,
    ) -> Result<u32> {
        let strict = self.inner.is_strict();
        let part = self.related_part(comments_relationship(strict))?;
        let current = match &part {
            Some((name, _)) => Some(
                self.part_payload(name)?
                    .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?,
            ),
            None => None,
        };
        let part_ids = match &current {
            Some(xml) => comment_ids(xml)?,
            None => Vec::new(),
        };
        let id = part_ids
            .iter()
            .copied()
            .chain(self.inner.max_comment_marker_id()?)
            .max()
            .map_or(Ok(0), |maximum| {
                maximum.checked_add(1).ok_or(Error::Refused {
                    position: 0,
                    reason: Refusal::InvalidPropertyValue,
                })
            })?;
        self.inner.anchor_comment_at(paragraph, range, id)?;
        match (part, current) {
            (Some((name, _)), Some(xml)) => {
                let updated = append_comment(&xml, id, comment)?;
                self.replace_part(&name, updated)?;
            },
            _ => {
                let namespace = if strict {
                    STRICT_WORD_NAMESPACE
                } else {
                    WORD_NAMESPACE
                };
                let empty = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="{namespace}"></w:comments>"#
                );
                let xml = append_comment(empty.as_bytes(), id, comment)?;
                self.stage_related_part(
                    comments_relationship(strict),
                    ("", "comments", ".xml"),
                    true,
                    COMMENTS_CONTENT_TYPE,
                    xml,
                )?;
            },
        }
        Ok(id)
    }

    /// Remove comment `id`: its entry in the comments part, its range
    /// markers and reference runs in the main document, and its entries in
    /// the modern-comment parts.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown comment, or a
    /// typed error for a malformed part.
    pub fn remove_comment(&mut self, id: u32) -> Result<&mut Self> {
        let strict = self.inner.is_strict();
        let (name, _) =
            self.related_part(comments_relationship(strict))?
                .ok_or(Error::Refused {
                    position: 0,
                    reason: Refusal::BlockNotFound,
                })?;
        let xml = self
            .part_payload(&name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
        let Some((start, end)) = comment_range(&xml, id)? else {
            return Err(Error::Refused {
                position: 0,
                reason: Refusal::BlockNotFound,
            });
        };
        let paragraph_ids = attribute_values(&xml[start..end], b"p", b"paraId")?;
        let mut updated = xml[..start].to_vec();
        updated.extend_from_slice(&xml[end..]);
        self.replace_part(&name, updated)?;
        self.inner.strip_comment_markers(id)?;
        if !paragraph_ids.is_empty() {
            self.strip_modern_entries(&paragraph_ids)?;
        }
        Ok(self)
    }

    /// Remove the `commentsExtended`, `commentsIds` and `commentsExtensible`
    /// entries of the comment paragraphs `paragraph_ids`.
    fn strip_modern_entries(&mut self, paragraph_ids: &[String]) -> Result<()> {
        let mut durable_ids = Vec::new();
        if let Some((name, _)) = self.related_part(COMMENTS_IDS_RELATIONSHIP)?
            && let Some(xml) = self.part_payload(&name)?
        {
            let (updated, removed) =
                without_elements(&xml, b"commentId", b"paraId", paragraph_ids)?;
            for element in removed {
                durable_ids.extend(attribute_values(&element, b"commentId", b"durableId")?);
            }
            self.replace_part(&name, updated)?;
        }
        if let Some((name, _)) = self.related_part(COMMENTS_EXTENDED_RELATIONSHIP)?
            && let Some(xml) = self.part_payload(&name)?
        {
            let (updated, _) = without_elements(&xml, b"commentEx", b"paraId", paragraph_ids)?;
            self.replace_part(&name, updated)?;
        }
        if !durable_ids.is_empty()
            && let Some((name, _)) = self.related_part(COMMENTS_EXTENSIBLE_RELATIONSHIP)?
            && let Some(xml) = self.part_payload(&name)?
        {
            let (updated, _) =
                without_elements(&xml, b"commentExtensible", b"durableId", &durable_ids)?;
            self.replace_part(&name, updated)?;
        }
        Ok(())
    }
}

const fn comments_relationship(strict: bool) -> &'static str {
    if strict {
        STRICT_COMMENTS_RELATIONSHIP
    } else {
        COMMENTS_RELATIONSHIP
    }
}

fn invalid(message: &str) -> Error {
    Error::Document(crate::Error::InvalidFormat(message.to_owned()))
}

/// Walk the direct `comment` children of a comments part: `(range, id)`.
fn comments(xml: &[u8]) -> Result<(Vec<((usize, usize), u32)>, usize, Vec<u8>)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut found = Vec::new();
    let mut open = None::<(usize, Option<u32>)>;
    let mut prefix = Vec::new();
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("comments offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| invalid("invalid comments part XML"))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("comments offset overflow"))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    prefix = element
                        .name()
                        .prefix()
                        .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec());
                } else if depth == 1 && element.local_name().as_ref() == b"comment" {
                    open = Some((start, id_of(&element)));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"comment" {
                    if let Some(id) = id_of(&element) {
                        found.push(((start, end), id));
                    }
                } else if depth == 0 {
                    return Err(invalid("the comments part has an empty root"));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some((comment_start, id)) = open.take()
                    && let Some(id) = id
                {
                    found.push(((comment_start, end), id));
                }
                if depth == 0 {
                    return Ok((found, start, prefix));
                }
            },
            Event::Eof => return Err(invalid("the comments part is unterminated")),
            _ => {},
        }
    }
}

fn id_of(element: &quick_xml::events::BytesStart<'_>) -> Option<u32> {
    element
        .unchecked_attributes()
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.local_name().as_ref() == b"id")
        .and_then(|attribute| std::str::from_utf8(&attribute.value).ok()?.parse().ok())
}

fn comment_ids(xml: &[u8]) -> Result<Vec<u32>> {
    Ok(comments(xml)?.0.into_iter().map(|(_, id)| id).collect())
}

fn comment_range(xml: &[u8], id: u32) -> Result<Option<(usize, usize)>> {
    Ok(comments(xml)?
        .0
        .into_iter()
        .find(|(_, found)| *found == id)
        .map(|(range, _)| range))
}

/// The comments part with `comment` appended as comment `id`.
fn append_comment(xml: &[u8], id: u32, comment: &NewComment) -> Result<Vec<u8>> {
    let (_, close, prefix) = comments(xml)?;
    let prefix = String::from_utf8(prefix).map_err(|_| invalid("invalid comments prefix"))?;
    let q = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    let escape = |text: &str| quick_xml::escape::escape(text).into_owned();
    if comment.author.is_empty() {
        return Err(Error::Refused {
            position: 0,
            reason: Refusal::InvalidPropertyValue,
        });
    }
    let mut entry = format!(
        r#"<{c} {id}="{value}" {author}="{name}""#,
        c = q("comment"),
        id = q("id"),
        value = id,
        author = q("author"),
        name = escape(&comment.author),
    );
    if let Some(date) = &comment.date {
        entry.push_str(&format!(r#" {}="{}""#, q("date"), escape(date)));
    }
    if let Some(initials) = &comment.initials {
        entry.push_str(&format!(r#" {}="{}""#, q("initials"), escape(initials)));
    }
    entry.push_str(&format!(
        r#"><{p}><{ppr}><{style} {val}="CommentText"/></{ppr}><{r}><{rpr}><{rstyle} {val}="CommentReference"/></{rpr}><{annotation}/></{r}><{r}><{t} xml:space="preserve">{text}</{t}></{r}></{p}></{c}>"#,
        p = q("p"),
        ppr = q("pPr"),
        style = q("pStyle"),
        val = q("val"),
        r = q("r"),
        rpr = q("rPr"),
        rstyle = q("rStyle"),
        annotation = q("annotationRef"),
        t = q("t"),
        text = escape(&comment.text),
        c = q("comment"),
    ));
    let mut output = Vec::with_capacity(xml.len() + entry.len());
    output.extend_from_slice(&xml[..close]);
    output.extend_from_slice(entry.as_bytes());
    output.extend_from_slice(&xml[close..]);
    Ok(output)
}

/// The values of attribute `attribute` (by local name) on every element
/// named `local` in a fragment.
fn attribute_values(xml: &[u8], local: &[u8], attribute: &[u8]) -> Result<Vec<String>> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut values = Vec::new();
    loop {
        match reader
            .read_event()
            .map_err(|_| invalid("invalid comment XML"))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == local =>
            {
                values.extend(
                    element
                        .unchecked_attributes()
                        .filter_map(std::result::Result::ok)
                        .filter(|found| found.key.local_name().as_ref() == attribute)
                        .filter_map(|found| String::from_utf8(found.value.into_owned()).ok()),
                );
            },
            Event::Eof => return Ok(values),
            _ => {},
        }
    }
}

/// The part without every element named `local` whose attribute
/// `attribute` holds one of `values`, and the removed elements.
fn without_elements(
    xml: &[u8],
    local: &[u8],
    attribute: &[u8],
    values: &[String],
) -> Result<(Vec<u8>, Vec<Vec<u8>>)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut ranges = Vec::new();
    let mut open = None::<(usize, usize)>;
    let mut depth = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("comment part offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| invalid("invalid comment part XML"))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("comment part offset overflow"))?;
        let matches = |element: &quick_xml::events::BytesStart<'_>| {
            element.local_name().as_ref() == local
                && element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .any(|found| {
                        found.key.local_name().as_ref() == attribute
                            && values
                                .iter()
                                .any(|value| value.as_bytes() == found.value.as_ref())
                    })
        };
        match event {
            Event::Empty(element) if open.is_none() && matches(&element) => {
                ranges.push((start, end));
            },
            Event::Start(element) => {
                depth += 1;
                if open.is_none() && matches(&element) {
                    open = Some((start, depth));
                }
            },
            Event::End(_) => {
                if let Some((open_start, open_depth)) = open
                    && open_depth == depth
                {
                    ranges.push((open_start, end));
                    open = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let mut output = Vec::with_capacity(xml.len());
    let mut removed = Vec::new();
    let mut cursor = 0usize;
    for (start, end) in ranges {
        output.extend_from_slice(&xml[cursor..start]);
        removed.push(xml[start..end].to_vec());
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok((output, removed))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn comments_parts_are_appended_and_entries_removed() {
        let xml = br#"<w:comments xmlns:w="urn:w"><w:comment w:id="3" w:author="A"><w:p w14:paraId="1A"/></w:comment></w:comments>"#;
        assert_eq!(comment_ids(xml).unwrap(), [3]);
        let appended = append_comment(xml, 4, &NewComment::new("B & C", "x<y")).unwrap();
        let text = String::from_utf8(appended.clone()).unwrap();
        assert!(text.contains(r#"<w:comment w:id="4" w:author="B &amp; C">"#));
        assert!(text.contains("x&lt;y"));
        assert_eq!(comment_ids(&appended).unwrap(), [3, 4]);
        assert!(comment_range(&appended, 4).unwrap().is_some());
        let ids = br#"<w16cid:commentsIds xmlns:w16cid="urn:c"><w16cid:commentId w16cid:paraId="1A" w16cid:durableId="9"/><w16cid:commentId w16cid:paraId="2B" w16cid:durableId="8"/></w16cid:commentsIds>"#;
        let (updated, removed) =
            without_elements(ids, b"commentId", b"paraId", &["1A".to_owned()]).unwrap();
        assert_eq!(removed.len(), 1);
        assert!(!String::from_utf8(updated).unwrap().contains("1A"));
    }
}
