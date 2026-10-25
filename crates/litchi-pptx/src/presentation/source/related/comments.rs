//! Detached legacy comment snapshots and the selected-slide eager reader.
//!
//! The wire codecs in [`crate::comments`] remain the single parser for legacy
//! `PresentationML` comments.  This module only projects their values into the
//! related-part API.  Both source-backed and eager reads use the same
//! projection, so native comment identifiers are retained in one place and
//! author names never require callers to parse XML themselves.

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};

use crate::comments::{self, Author as ParsedAuthor, Comment as ParsedComment};
use crate::error::{Error, Result};
use crate::parts::{PresentationPart, SlidePart, related_by_type};

use super::RelatedPart;

/// One semantic comment entry from a legacy comments part.
#[derive(Clone)]
pub struct Comment {
    pub(super) text: String,
    pub(super) date_time: Option<String>,
    pub(super) x: i64,
    pub(super) y: i64,
    pub(super) author_id: u32,
    pub(super) index: u32,
    pub(super) author_name: Option<Arc<str>>,
    pub(super) initials: Option<Arc<str>>,
}

impl fmt::Debug for Comment {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Comment")
            .field("has_author", &self.author_name.is_some())
            .finish_non_exhaustive()
    }
}

impl Comment {
    /// The authored comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The authored timestamp, if present.
    #[must_use]
    pub fn date_time(&self) -> Option<&str> {
        self.date_time.as_deref()
    }

    /// The authored comment location in slide coordinates.
    #[must_use]
    pub fn location(&self) -> (i64, i64) {
        (self.x, self.y)
    }

    /// The resolved legacy comment-author display name, when an author part
    /// was selected and contained this comment's native author ID.
    #[must_use]
    pub fn author_name(&self) -> Option<&str> {
        self.author_name.as_deref()
    }

    /// The resolved legacy comment-author initials, when an author part was
    /// selected and contained this comment's native author ID.
    #[must_use]
    pub fn initials(&self) -> Option<&str> {
        self.initials.as_deref()
    }
}

/// A detached comments snapshot selected from one slide.
#[derive(Clone)]
pub struct Comments {
    pub(super) data: Arc<CommentsData>,
}

pub(super) struct CommentsData {
    pub(super) related: RelatedPart,
    pub(super) author_related: Option<RelatedPart>,
    pub(super) entries: Box<[Comment]>,
}

impl fmt::Debug for Comments {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Comments")
            .field("comment_count", &self.data.entries.len())
            .field("has_authors", &self.data.author_related.is_some())
            .finish_non_exhaustive()
    }
}

impl Comments {
    /// The comments in authored document order.
    #[must_use]
    pub fn entries(&self) -> &[Comment] {
        &self.data.entries
    }

    /// Number of comments in this snapshot.
    #[must_use]
    pub fn len(&self) -> usize {
        self.data.entries.len()
    }

    /// Whether this snapshot contains no comments.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.entries.is_empty()
    }

    /// Project parsed comments and, when selected, their parsed author list.
    ///
    /// A non-empty comments part must carry a selected presentation-level
    /// author relationship. Every native author ID must resolve before the
    /// snapshot is published. Empty comments parts remain valid without an
    /// author part.
    pub(super) fn from_parts(
        related: RelatedPart,
        author_related: Option<RelatedPart>,
        parsed: Vec<ParsedComment>,
        authors: Option<&[ParsedAuthor]>,
    ) -> Result<Self> {
        if !parsed.is_empty() && authors.is_none() {
            return Err(Error::Relationship(
                "presentation comments require a comment-author part".into(),
            ));
        }
        let author_index = author_index(authors)?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(parsed.len())
            .map_err(|source| Error::Allocation {
                resource: "presentation comment projections",
                source,
            })?;

        for comment in parsed {
            let (author_name, initials) = match author_index.as_ref() {
                None => (None, None),
                Some(index) => {
                    let author = index.get(&comment.author_id).ok_or_else(|| {
                        Error::Relationship(
                            "presentation comment author reference is unresolved".into(),
                        )
                    })?;
                    (Some(Arc::clone(&author.0)), Some(Arc::clone(&author.1)))
                },
            };
            entries.push(Comment {
                text: comment.text,
                date_time: comment.date_time,
                x: comment.x,
                y: comment.y,
                author_id: comment.author_id,
                index: comment.index,
                author_name,
                initials,
            });
        }

        Ok(Self {
            data: Arc::new(CommentsData {
                related,
                author_related,
                entries: entries.into_boxed_slice(),
            }),
        })
    }
}

type AuthorIndex = HashMap<u32, (Arc<str>, Arc<str>)>;

fn author_index(authors: Option<&[ParsedAuthor]>) -> Result<Option<AuthorIndex>> {
    let Some(authors) = authors else {
        return Ok(None);
    };
    let mut index = HashMap::new();
    index
        .try_reserve(authors.len())
        .map_err(|source| Error::Allocation {
            resource: "presentation comment author index",
            source,
        })?;
    for author in authors {
        if index
            .insert(
                author.id,
                (
                    Arc::<str>::from(author.name.as_str()),
                    Arc::<str>::from(author.initials.as_str()),
                ),
            )
            .is_some()
        {
            return Err(Error::Relationship(
                "presentation comment authors contain duplicate IDs".into(),
            ));
        }
    }
    Ok(Some(index))
}

/// Read the selected eager slide's optional legacy comments part.
///
/// The walk follows only the slide comments relationship and the one
/// presentation-level author relationship. It does not call
/// the package-wide comments graph or serialize and reopen the package.
pub(crate) fn load_eager<'a>(
    package: &'a OpcPackage,
    slide: &SlidePart<'a>,
    limits: comments::Limits,
) -> Result<Option<Comments>> {
    let Some((relationship, part)) = related_by_type(
        package,
        slide.part(),
        comments::COMMENTS_REL,
        "comments",
        ct::PML_COMMENTS,
    )?
    else {
        return Ok(None);
    };
    let processed = super::process_eager_part(
        part,
        comments::MAX_PART_BYTES.min(limits.max_input_bytes),
        package.read_limits(),
    )?;
    if !part.rels().is_empty() {
        return Err(Error::Relationship(
            "comments part must not have relationships".into(),
        ));
    }
    let parsed = comments::parse_slide_comments_strict(processed.blob(), limits)?;
    let authors = selected_eager_authors(package, &parsed, limits)?;
    let (author_related, author_values) = match authors {
        Some((part, authors)) => (Some(part), Some(authors)),
        None => (None, None),
    };
    Comments::from_parts(
        RelatedPart::from_eager(part, relationship.r_id()),
        author_related,
        parsed,
        author_values.as_deref(),
    )
    .map(Some)
}

fn selected_eager_authors(
    package: &OpcPackage,
    parsed_comments: &[ParsedComment],
    limits: comments::Limits,
) -> Result<Option<(RelatedPart, Vec<ParsedAuthor>)>> {
    // A comments part with no entries does not need the presentation author
    // graph. A non-empty part requires the selected author closure.
    if parsed_comments.is_empty() {
        return Ok(None);
    }

    let presentation = PresentationPart::from_package(package)?;
    let mut selected = None;
    for relationship in presentation.part().rels().iter() {
        if !crate::parts::is_relationship_type(
            relationship.reltype(),
            comments::AUTHORS_REL,
            "commentAuthors",
        ) {
            continue;
        }
        if selected.is_some() {
            return Err(Error::Relationship(
                "presentation has multiple comment-author relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::Relationship(
                "comment-author relationship must be internal".into(),
            ));
        }
        selected = Some(relationship);
    }

    let relationship = selected.ok_or_else(|| {
        Error::Relationship("presentation comments require a comment-author part".into())
    })?;
    let target = relationship
        .target_partname()
        .map_err(|_| Error::Relationship("comment-author relationship target is invalid".into()))?;
    let part = package
        .get_part(&target)
        .map_err(|_| Error::Relationship("comment-author target part is missing".into()))?;
    crate::parts::validate_content_type(part, comments::AUTHORS_CONTENT_TYPE)?;
    let processed = super::process_eager_part(
        part,
        comments::MAX_PART_BYTES.min(limits.max_input_bytes),
        package.read_limits(),
    )?;
    if !part.rels().is_empty() {
        return Err(Error::Relationship(
            "comment-author part must not have relationships".into(),
        ));
    }
    let authors = comments::parse_comment_authors(processed.blob(), limits)?;
    Ok(Some((
        RelatedPart::from_eager(part, relationship.r_id()),
        authors,
    )))
}
