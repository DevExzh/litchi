//! Source-backed per-slide modern-comment acquisition.
//!
//! This module owns only the source graph walk.  The public [`List`] value is
//! projected by `projection`; the source owner is never retained by that
//! value.  A selected slide contributes one comments part and, when the part
//! contains comments, the one presentation-level author part they reference.
//!
//! ```
//! use std::io::Cursor;
//! use litchi_pptx::SourceBackedPresentation;
//!
//! let mut builder = litchi_pptx::builder::Builder::new()?;
//! builder.presentation()?.add_slide()?;
//! let source = SourceBackedPresentation::from_reader(Cursor::new(builder.build()?.to_bytes()?))?;
//! assert!(source.slide(0).unwrap().modern_comments()?.is_none());
//! # Ok::<(), litchi_pptx::Error>(())
//! ```

use std::collections::HashSet;

use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use crate::error::{Error, Result};
use crate::modern_comments::{self as codec, Authors, List as ParsedList};

use super::related::RelatedPart;
use super::{SourceInner, SourcePart, SourceSlide};

mod projection;

pub use projection::{Comment, List, Reply, raw};

const COMMENTS_RELATIONSHIP: &str = codec::MODERN_COMMENT_RELATIONSHIP_TYPE;
const COMMENTS_CONTENT_TYPE: &str = codec::MODERN_COMMENT_CONTENT_TYPE;
const AUTHORS_RELATIONSHIP: &str = codec::MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE;
const AUTHORS_CONTENT_TYPE: &str = codec::MODERN_COMMENT_AUTHOR_CONTENT_TYPE;

fn relationship_error(message: &'static str) -> Error {
    Error::Relationship(message.into())
}

fn invalid_error(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn redact_codec_error(error: Error, resource: &'static str) -> Error {
    match error {
        Error::Invalid(_) => Error::Invalid(format!("invalid {resource}")),
        Error::Xml(_) => Error::Xml(format!("invalid {resource} XML")),
        other => other,
    }
}

fn target_uri(owner: &PackURI, target: &str) -> Result<PackURI> {
    PackURI::from_rel_ref(owner.base_uri(), target)
        .map_err(|_| relationship_error("source modern-comment relationship target is invalid"))
}

fn related_part(part: &SourcePart, relationship_id: &str) -> RelatedPart {
    RelatedPart::new(part.clone(), relationship_id)
}

fn selected_comments_relationship(
    owner: &SourceInner,
    slide: &SourceSlide,
) -> Result<Option<(String, PackURI)>> {
    let mut selected = None;
    for relationship in &slide.data.binding.slide.relationships {
        owner.package.check_execution()?;
        if relationship.kind != COMMENTS_RELATIONSHIP {
            continue;
        }
        if selected.is_some() {
            return Err(relationship_error(
                "source slide has multiple modern-comment relationships",
            ));
        }
        if relationship.mode != TargetMode::Internal {
            return Err(relationship_error(
                "source modern-comment relationship must be internal",
            ));
        }
        selected = Some((relationship.id.clone(), relationship.target.clone()));
    }
    selected
        .map(|(relationship_id, target)| {
            target_uri(&slide.data.part_uri, &target).map(|target| (relationship_id, target))
        })
        .transpose()
}

fn selected_authors_part(
    owner: &SourceInner,
    comments: &ParsedList,
) -> Result<Option<(String, SourcePart, Authors)>> {
    if comments.comments.is_empty() {
        return Ok(None);
    }

    let mut selected = None;
    for relationship in owner._presentation.source_relationships().iter() {
        owner.package.check_execution()?;
        if relationship.reltype() != AUTHORS_RELATIONSHIP {
            continue;
        }
        if selected.is_some() {
            return Err(relationship_error(
                "presentation has multiple modern-comment author relationships",
            ));
        }
        if relationship.target_mode() == TargetMode::External {
            return Err(relationship_error(
                "source modern-comment author relationship must be internal",
            ));
        }
        selected = Some(relationship);
    }

    let relationship = selected.ok_or_else(|| {
        relationship_error("source modern comments require a modern-comment author part")
    })?;
    let target = relationship.target_partname().map_err(|_| {
        relationship_error("source modern-comment author relationship target is invalid")
    })?;
    let part = super::related::load_part(owner, &target, AUTHORS_CONTENT_TYPE, codec::MAX_BYTES)?;
    if !part.source_relationships().is_empty() {
        return Err(relationship_error(
            "source modern-comment author part must not have relationships",
        ));
    }
    let authors = Authors::from_processed(part.blob())
        .map_err(|error| redact_codec_error(error, "modern-comment authors"))?;
    Ok(Some((relationship.r_id().to_owned(), part, authors)))
}

fn validate_unique_ids(
    comments: &ParsedList,
    mut check: impl FnMut() -> Result<()>,
    resource: &'static str,
) -> Result<()> {
    let mut count = comments.comments.len();
    for comment in &comments.comments {
        check()?;
        count = count
            .checked_add(comment.replies.len())
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::CommentEntries,
                ),
                observed: usize::MAX,
                maximum: codec::MAX_COMMENTS.saturating_add(codec::MAX_REPLIES),
            })?;
    }

    let mut ids = HashSet::new();
    ids.try_reserve(count)
        .map_err(|source| Error::Allocation { resource, source })?;
    for comment in &comments.comments {
        check()?;
        if !ids.insert(comment.id.as_str()) {
            return Err(invalid_error(
                "source modern comments contain duplicate comment or reply IDs",
            ));
        }
        for reply in &comment.replies {
            check()?;
            if !ids.insert(reply.id.as_str()) {
                return Err(invalid_error(
                    "source modern comments contain duplicate comment or reply IDs",
                ));
            }
        }
    }
    Ok(())
}

fn validate_author_references(
    comments: &ParsedList,
    authors: &Authors,
    mut check: impl FnMut() -> Result<()>,
    resource: &'static str,
) -> Result<()> {
    let mut ids = HashSet::new();
    ids.try_reserve(authors.authors.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    for author in &authors.authors {
        check()?;
        if !ids.insert(author.id.as_str()) {
            return Err(invalid_error(
                "source modern comment authors contain duplicate IDs",
            ));
        }
    }

    for comment in &comments.comments {
        check()?;
        if !ids.contains(comment.author_id.as_str()) {
            return Err(relationship_error(
                "source modern comment author reference is unresolved",
            ));
        }
        if let Some(assigned) = &comment.assigned_to {
            for author_id in assigned {
                check()?;
                if !ids.contains(author_id.as_str()) {
                    return Err(relationship_error(
                        "source modern comment assignment reference is unresolved",
                    ));
                }
            }
        }
        for reply in &comment.replies {
            check()?;
            if !ids.contains(reply.author_id.as_str()) {
                return Err(relationship_error(
                    "source modern comment reply author reference is unresolved",
                ));
            }
        }
    }
    Ok(())
}

fn acquire(owner: &SourceInner, slide: &SourceSlide) -> Result<Option<List>> {
    let Some((relationship_id, target)) = selected_comments_relationship(owner, slide)? else {
        return Ok(None);
    };

    let comments_part =
        super::related::load_part(owner, &target, COMMENTS_CONTENT_TYPE, codec::MAX_BYTES)?;
    if !comments_part.source_relationships().is_empty() {
        return Err(relationship_error(
            "source modern-comment part must not have relationships",
        ));
    }
    let comments = ParsedList::from_processed(comments_part.blob())
        .map_err(|error| redact_codec_error(error, "modern comments"))?;
    validate_unique_ids(
        &comments,
        || owner.package.check_execution().map_err(Error::from),
        "source modern comment IDs",
    )?;
    let authors = selected_authors_part(owner, &comments)?;
    if let Some((_, _, parsed_authors)) = &authors {
        validate_author_references(
            &comments,
            parsed_authors,
            || owner.package.check_execution().map_err(Error::from),
            "source modern comment author IDs",
        )?;
    }
    let author_part = authors
        .map(|(relationship_id, part, authors)| (related_part(&part, &relationship_id), authors));
    let list = List::from_parts(
        related_part(&comments_part, &relationship_id),
        comments,
        author_part,
    )
    .map_err(|error| redact_codec_error(error, "modern comments"))?;
    Ok(Some(list))
}

fn eager_processed_part(
    part: &dyn Part,
    source_limits: litchi_opc::SourceLimits,
) -> Result<BlobPart> {
    super::related::process_eager_part(part, codec::MAX_BYTES, source_limits)
}

fn eager_target_part<'a>(
    package: &'a OpcPackage,
    target: &PackURI,
    content_type: &str,
) -> Result<&'a dyn Part> {
    let part = package
        .get_part(target)
        .map_err(|_| relationship_error("modern-comment relationship target part is missing"))?;
    if part.content_type() != content_type {
        return Err(Error::ContentType {
            expected: content_type.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    Ok(part)
}

fn selected_eager_comments_relationship(
    slide: &crate::parts::SlidePart<'_>,
) -> Result<Option<(String, PackURI)>> {
    let mut selected = None;
    for relationship in slide.part().rels().iter() {
        if relationship.reltype() != COMMENTS_RELATIONSHIP {
            continue;
        }
        if selected.is_some() {
            return Err(relationship_error(
                "slide has multiple modern-comment relationships",
            ));
        }
        if relationship.target_mode() == TargetMode::External {
            return Err(relationship_error(
                "modern-comment relationship must be internal",
            ));
        }
        let target = relationship
            .target_partname()
            .map_err(|_| relationship_error("modern-comment relationship target is invalid"))?;
        selected = Some((relationship.r_id().to_owned(), target));
    }
    Ok(selected)
}

fn selected_eager_authors_part<'a>(
    package: &'a OpcPackage,
    comments: &ParsedList,
) -> Result<Option<(String, &'a dyn Part, Authors)>> {
    if comments.comments.is_empty() {
        return Ok(None);
    }

    let presentation = crate::parts::PresentationPart::from_package(package)?;
    let mut selected = None;
    for relationship in presentation.part().rels().iter() {
        if relationship.reltype() != AUTHORS_RELATIONSHIP {
            continue;
        }
        if selected.is_some() {
            return Err(relationship_error(
                "presentation has multiple modern-comment author relationships",
            ));
        }
        if relationship.target_mode() == TargetMode::External {
            return Err(relationship_error(
                "modern-comment author relationship must be internal",
            ));
        }
        let target = relationship.target_partname().map_err(|_| {
            relationship_error("modern-comment author relationship target is invalid")
        })?;
        selected = Some((relationship.r_id().to_owned(), target));
    }

    let Some((relationship_id, target)) = selected else {
        return Err(relationship_error(
            "modern comments require a modern-comment author part",
        ));
    };
    let part = eager_target_part(package, &target, AUTHORS_CONTENT_TYPE)?;
    if !part.rels().is_empty() {
        return Err(relationship_error(
            "modern-comment author part must not have relationships",
        ));
    }
    let processed = eager_processed_part(part, package.read_limits())?;
    let authors = Authors::from_processed(processed.blob())
        .map_err(|error| redact_codec_error(error, "modern-comment authors"))?;
    Ok(Some((relationship_id, part, authors)))
}

/// Acquire the selected eager slide's optional modern-comment list.
///
/// This follows only the selected slide's comments relationship and, when
/// needed, the presentation-level author relationship. It intentionally does
/// not build the package-wide modern-comment graph.
pub(crate) fn acquire_eager(
    package: &OpcPackage,
    slide: &crate::parts::SlidePart<'_>,
) -> Result<Option<List>> {
    let Some((relationship_id, target)) = selected_eager_comments_relationship(slide)? else {
        return Ok(None);
    };
    let comments_part = eager_target_part(package, &target, COMMENTS_CONTENT_TYPE)?;
    if !comments_part.rels().is_empty() {
        return Err(relationship_error(
            "modern-comment part must not have relationships",
        ));
    }
    let processed = eager_processed_part(comments_part, package.read_limits())?;
    let comments = ParsedList::from_processed(processed.blob())
        .map_err(|error| redact_codec_error(error, "modern comments"))?;
    validate_unique_ids(&comments, || Ok(()), "modern comment IDs")?;
    let authors = selected_eager_authors_part(package, &comments)?;
    if let Some((_, _, parsed_authors)) = &authors {
        validate_author_references(
            &comments,
            parsed_authors,
            || Ok(()),
            "modern comment author IDs",
        )?;
    }
    let author_part = authors.map(|(relationship_id, part, authors)| {
        (RelatedPart::from_eager(part, &relationship_id), authors)
    });
    let list = List::from_parts(
        RelatedPart::from_eager(comments_part, &relationship_id),
        comments,
        author_part,
    )
    .map_err(|error| redact_codec_error(error, "modern comments"))?;
    Ok(Some(list))
}

impl SourceSlide {
    /// Acquire the selected slide's optional modern-comment list.
    ///
    /// Only this slide's modern-comment part and its presentation-level author
    /// part are loaded.  Missing slide relationships return `None`; malformed
    /// or incomplete selected graphs are refused.
    ///
    /// # Errors
    ///
    /// Returns typed source, execution, limit, relationship, allocation, or
    /// XML errors when the selected graph cannot be acquired safely.
    pub fn modern_comments(&self) -> Result<Option<List>> {
        super::related::with_fences(&self.owner, || acquire(&self.owner, self))
    }
}
