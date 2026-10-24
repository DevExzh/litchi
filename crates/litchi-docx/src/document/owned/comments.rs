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
use super::stories::text_run;
use super::threads::keyed_paragraph_id;
use crate::Package;
use crate::part_scan::{self, qualified};

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
    /// The comment this one replies to (record 0345).
    parent: Option<u32>,
    /// The resolved state (record 0345).
    done: Option<bool>,
    /// The direct run a whole-paragraph range starts after (record 0345).
    start_after: Option<Position>,
    /// Whether the comment has range markers (record 0345).
    range: bool,
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
            parent: None,
            done: None,
            start_after: None,
            range: true,
        }
    }

    /// Make the comment a reply to comment `parent` (its native
    /// identifier), threaded through `commentsExtended`
    /// (`w15:paraIdParent`) as Word threads replies (record 0345).
    #[must_use]
    pub const fn with_parent(mut self, parent: u32) -> Self {
        self.parent = Some(parent);
        self
    }

    /// Set the comment's resolved state (`w15:done` in `commentsExtended`,
    /// record 0345).
    #[must_use]
    pub const fn with_done(mut self, done: bool) -> Self {
        self.done = Some(done);
        self
    }

    /// For a whole-paragraph anchor, start the comment's range after the
    /// paragraph's direct run `run` instead of at the paragraph start
    /// (OfficeCLI's `runStart`, record 0345). A run the paragraph lacks
    /// keeps the paragraph start.
    #[must_use]
    pub const fn with_start_after(mut self, run: Position) -> Self {
        self.start_after = Some(run);
        self
    }

    /// Whether the comment spans a range (`true`, the default: range
    /// markers and a reference run) or is a point comment carrying only
    /// its reference run (`false`; OfficeCLI's `range=false`, record 0345).
    #[must_use]
    pub const fn with_range(mut self, range: bool) -> Self {
        self.range = range;
        self
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

/// A change of an existing comment (record 0335): its author, initials,
/// date and text. Each part left unset keeps the comment's current value.
///
/// ```
/// use litchi_docx::document::CommentChange;
///
/// let change = CommentChange::new().with_author("Grace").with_text("Fixed");
/// assert!(!change.is_empty());
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CommentChange {
    author: Option<String>,
    initials: Option<String>,
    date: Option<String>,
    text: Option<String>,
    done: Option<bool>,
}

impl CommentChange {
    /// A change that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The new author (`w:author`, not empty).
    #[must_use]
    pub fn with_author(mut self, author: impl Into<String>) -> Self {
        self.author = Some(author.into());
        self
    }

    /// The new initials (`w:initials`).
    #[must_use]
    pub fn with_initials(mut self, initials: impl Into<String>) -> Self {
        self.initials = Some(initials.into());
        self
    }

    /// The new date (`w:date`), an `xsd:dateTime` written as given.
    #[must_use]
    pub fn with_date(mut self, date: impl Into<String>) -> Self {
        self.date = Some(date.into());
        self
    }

    /// The new text, replacing the comment's text as a footnote's text is
    /// replaced ([`Edit::set_note_text`]).
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// The new resolved state (`w15:done` in `commentsExtended`, record
    /// 0345).
    #[must_use]
    pub const fn with_done(mut self, done: bool) -> Self {
        self.done = Some(done);
        self
    }

    /// Whether the change sets nothing.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.author.is_none()
            && self.initials.is_none()
            && self.date.is_none()
            && self.text.is_none()
            && self.done.is_none()
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
        let strict = self.is_strict()?;
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
            .chain(self.inner_mut()?.max_comment_marker_id()?)
            .max()
            .map_or(Ok(0), |maximum| {
                maximum.checked_add(1).ok_or(Error::Refused {
                    position: 0,
                    reason: Refusal::InvalidPropertyValue,
                })
            })?;
        self.inner_mut()?.anchor_comment_at(
            paragraph,
            range,
            id,
            comment.start_after,
            !comment.range,
        )?;
        // Replies and resolution (record 0345): the comment's paragraph
        // takes a fresh identifier, which `commentsExtended` keys on.
        let threaded = comment.parent.is_some() || comment.done.is_some();
        let mut current = current.map(|xml| xml.as_ref().clone());
        let mut fresh = self.fresh_paragraph_ids(current.as_deref())?;
        let parent = match comment.parent {
            Some(parent) => {
                let xml = current.as_mut().ok_or(Error::Refused {
                    position: 0,
                    reason: Refusal::BlockNotFound,
                })?;
                Some(
                    keyed_paragraph_id(xml, parent, &mut fresh)?.ok_or(Error::Refused {
                        position: 0,
                        reason: Refusal::BlockNotFound,
                    })?,
                )
            },
            None => None,
        };
        let paragraph_id = threaded.then(|| fresh.next()).transpose()?;
        match (part, current) {
            (Some((name, _)), Some(xml)) => {
                let updated = append_comment(&xml, id, comment, paragraph_id)?;
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
                let xml = append_comment(empty.as_bytes(), id, comment, paragraph_id)?;
                self.stage_related_part(
                    comments_relationship(strict),
                    ("", "comments", ".xml"),
                    true,
                    COMMENTS_CONTENT_TYPE,
                    xml,
                )?;
            },
        }
        if let Some(paragraph_id) = paragraph_id {
            let mut entries = Vec::with_capacity(2);
            if let Some(parent) = parent {
                // Word writes an entry for every comment: the thread's
                // root gets one when it has none.
                entries.push(super::threads::Entry {
                    paragraph: parent,
                    parent: None,
                    done: None,
                });
            }
            entries.push(super::threads::Entry {
                paragraph: paragraph_id,
                parent,
                done: Some(comment.done.unwrap_or(false)),
            });
            self.upsert_comment_entries(&entries, &fresh.assigned)?;
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
        let strict = self.is_strict()?;
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
        self.inner_mut()?.strip_comment_markers(id)?;
        if !paragraph_ids.is_empty() {
            self.strip_modern_entries(&paragraph_ids)?;
        }
        Ok(self)
    }

    /// Change comment `id`: its author, initials and date attributes, and
    /// its text. The text is replaced as a note's is: the comment's first
    /// run that is not its `w:annotationRef` mark holds the text (keeping
    /// its properties) and its other such runs are removed; a comment
    /// holding a field is refused rather than flattened. Its paragraphs,
    /// their identifiers (which the modern-comment parts key on) and the
    /// mark stay. Returns whether anything changed.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{CommentChange, ElementPath, ElementStep, NewComment};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Reviewed text");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let id = edit.add_comment_at(&paragraph, None, &NewComment::new("Ada", "Draft"))?;
    /// let package = edit.commit()?.into_package();
    /// let mut edit = package.edit()?;
    /// let change = CommentChange::new().with_author("Grace").with_text("Final");
    /// assert!(edit.set_comment(id, &change)?);
    /// let comments = edit.commit()?.into_package().document()?.comments()?;
    /// assert_eq!((comments[0].author(), comments[0].text()?.as_str()), ("Grace", "Final"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown comment,
    /// [`Refusal::ComplexContent`] for a comment holding a field,
    /// [`Refusal::InvalidPropertyValue`] for an empty author or text XML
    /// cannot carry, or a typed error for a malformed part.
    pub fn set_comment(&mut self, id: u32, change: &CommentChange) -> Result<bool> {
        let invalid_value = || Error::Refused {
            position: 0,
            reason: Refusal::InvalidPropertyValue,
        };
        if change.author.as_deref().is_some_and(str::is_empty) {
            return Err(invalid_value());
        }
        let controls = |text: &str| text.chars().any(char::is_control);
        if [&change.author, &change.initials, &change.date]
            .into_iter()
            .flatten()
            .any(|value| controls(value))
        {
            return Err(invalid_value());
        }
        let strict = self.is_strict()?;
        let not_found = || Error::Refused {
            position: 0,
            reason: Refusal::BlockNotFound,
        };
        let (name, _) = self
            .related_part(comments_relationship(strict))?
            .ok_or_else(not_found)?;
        let xml = self
            .part_payload(&name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
        let tree = part_scan::scan(&xml, 4).map_err(Error::Document)?;
        let comment = tree
            .children(0)
            .find(|child| {
                let element = &tree.elements[*child];
                element.is(b"comment")
                    && element
                        .attribute(b"id")
                        .and_then(|value| value.trim().parse::<u32>().ok())
                        == Some(id)
            })
            .ok_or_else(not_found)?;
        let element = &tree.elements[comment];
        let mut edits: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        let mut tag = xml[element.start..element.tag_end].to_vec();
        for (local, value) in [
            ("author", &change.author),
            ("initials", &change.initials),
            ("date", &change.date),
        ] {
            let Some(value) = value else { continue };
            let key = crate::picture_spans::attribute_named(&tag, local)
                .unwrap_or_else(|| qualified(&element.prefix, local));
            tag = crate::picture_spans::with_attribute(&tag, &key, value)
                .ok_or_else(|| invalid("invalid comment XML"))?;
        }
        if tag.as_slice() != &xml[element.start..element.tag_end] {
            edits.push((element.start, element.tag_end, tag));
        }
        if let Some(text) = &change.text {
            if part_scan::holds_any(&xml[element.start..element.end], part_scan::FIELD_MARKUP)
                .map_err(Error::Document)?
            {
                return Err(Error::Refused {
                    position: 0,
                    reason: Refusal::ComplexContent,
                });
            }
            let mut first_content = true;
            let paragraphs: Vec<usize> = tree
                .children(comment)
                .filter(|child| tree.elements[*child].is(b"p"))
                .collect();
            for paragraph in &paragraphs {
                for run in tree
                    .children(*paragraph)
                    .filter(|child| tree.elements[*child].is(b"r"))
                    .collect::<Vec<_>>()
                {
                    if tree.child(run, b"annotationRef").is_some() {
                        continue;
                    }
                    let run_element = &tree.elements[run];
                    if first_content {
                        first_content = false;
                        let properties = tree.child(run, b"rPr").map(|child| {
                            let rpr = &tree.elements[child];
                            xml[rpr.start..rpr.end].to_vec()
                        });
                        edits.push((
                            run_element.start,
                            run_element.end,
                            text_run(&run_element.prefix, properties.as_deref(), text)?,
                        ));
                    } else {
                        edits.push((run_element.start, run_element.end, Vec::new()));
                    }
                }
            }
            if first_content {
                let markup = text_run(&element.prefix, None, text)?;
                match paragraphs.first() {
                    Some(&paragraph) => {
                        let paragraph = &tree.elements[paragraph];
                        if paragraph.empty {
                            let mut replacement =
                                xml[paragraph.start..paragraph.tag_end - 2].to_vec();
                            replacement.push(b'>');
                            replacement.extend_from_slice(&markup);
                            replacement.extend_from_slice(
                                format!("</{}>", qualified(&paragraph.prefix, "p")).as_bytes(),
                            );
                            edits.push((paragraph.start, paragraph.end, replacement));
                        } else {
                            edits.push((paragraph.close, paragraph.close, markup));
                        }
                    },
                    None if element.empty => return Err(invalid("invalid comment XML")),
                    None => {
                        let p = qualified(&element.prefix, "p");
                        let mut paragraph = format!("<{p}>").into_bytes();
                        paragraph.extend_from_slice(&markup);
                        paragraph.extend_from_slice(format!("</{p}>").as_bytes());
                        edits.push((element.close, element.close, paragraph));
                    },
                }
            }
        }
        edits.sort_by_key(|(start, end, _)| (*start, *end != *start));
        let mut output = Vec::with_capacity(xml.len());
        let mut cursor = 0usize;
        for (start, end, bytes) in edits {
            output.extend_from_slice(&xml[cursor..start]);
            output.extend_from_slice(&bytes);
            cursor = end;
        }
        output.extend_from_slice(&xml[cursor..]);
        // The resolved state (record 0345): the comment's last paragraph
        // keys its `commentsExtended` entry, given an identifier if it has
        // none.
        let mut threaded = false;
        if let Some(done) = change.done {
            let mut fresh = self.fresh_paragraph_ids(Some(&output))?;
            let paragraph =
                keyed_paragraph_id(&mut output, id, &mut fresh)?.ok_or(Error::Refused {
                    position: 0,
                    reason: Refusal::ComplexContent,
                })?;
            threaded = self.upsert_comment_entries(
                &[super::threads::Entry {
                    paragraph,
                    parent: None,
                    done: Some(done),
                }],
                &fresh.assigned,
            )?;
        }
        if output.as_slice() == xml.as_slice() {
            return Ok(threaded);
        }
        self.replace_part(&name, output)?;
        Ok(true)
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

/// The comments part with `comment` appended as comment `id`; its
/// paragraph carries `paragraph_id` (`w14:paraId`) when given (record
/// 0345).
fn append_comment(
    xml: &[u8],
    id: u32,
    comment: &NewComment,
    paragraph_id: Option<u32>,
) -> Result<Vec<u8>> {
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
    let paragraph_attributes = match paragraph_id {
        Some(paragraph_id) => super::threads::paragraph_id_attributes(xml, paragraph_id),
        None => String::new(),
    };
    entry.push_str(&format!(
        r#"><{p}{paragraph_attributes}><{ppr}><{style} {val}="CommentText"/></{ppr}><{r}><{rpr}><{rstyle} {val}="CommentReference"/></{rpr}><{annotation}/></{r}><{r}><{t} xml:space="preserve">{text}</{t}></{r}></{p}></{c}>"#,
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
        let appended = append_comment(xml, 4, &NewComment::new("B & C", "x<y"), None).unwrap();
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
