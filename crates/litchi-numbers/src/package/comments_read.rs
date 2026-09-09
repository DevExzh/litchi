//! Bounded semantic readers for Numbers comment metadata.
//!
//! Archive object lookup stays in the parent comment adapter. This module
//! owns only the borrowed Buffa author projection and its fallible copies, so
//! native identifiers never become part of the public value model.

use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::annotation_author_codec;

use super::{CommentAuthor, Error, LimitKind, Path};

const MAX_AUTHOR_NESTING: u32 = 64;

/// Decode one strict annotation-author payload into archive-free display
/// metadata. The returned strings are bounded copies; no reference into the
/// source payload escapes this function.
pub(super) fn decode_author(
    data: &[u8],
    max_text: usize,
    max_references: usize,
    path: Path,
) -> Result<(CommentAuthor, usize), Error> {
    let source_bytes = data.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    let fields = data.len().clamp(1, WireLimits::MAX_FIELDS);
    let work = data
        .len()
        .saturating_mul(64)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let options = annotation_author_codec::DecodeOptions::new(
        source_bytes,
        fields,
        work,
        MAX_AUTHOR_NESTING,
        max_references.max(1),
        // The codec must account for every wire string, including repeated
        // public IDs that are not projected into the returned author. The
        // semantic budget is charged separately while copying the selected
        // name and public ID below.
        source_bytes,
        fields,
    );
    let (snapshot, _report) =
        annotation_author_codec::decode_annotation_author_with_report(data, options)
            .map_err(|error| map_error(error, path))?;

    let mut retained = 0usize;
    let display_name = snapshot
        .name()
        .map(|value| copy_text(value, max_text, &mut retained, path))
        .transpose()?;
    let public_id = snapshot
        .public_id()
        .map(|value| copy_text(value, max_text, &mut retained, path))
        .transpose()?;
    Ok((CommentAuthor::new(display_name, public_id), retained))
}

fn copy_text(
    value: &str,
    maximum: usize,
    retained: &mut usize,
    path: Path,
) -> Result<Box<str>, Error> {
    let next = retained
        .checked_add(value.len())
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::TextBytes,
            observed: usize::MAX,
            maximum,
            path,
        })?;
    if next > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::TextBytes,
            observed: next,
            maximum,
            path,
        });
    }
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|_| Error::Allocation {
            amount: value.len(),
            path,
        })?;
    copied.push_str(value);
    *retained = next;
    Ok(copied.into_boxed_str())
}

fn map_error(error: annotation_author_codec::DecodeError, path: Path) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    match limit {
        annotation_author_codec::DecodeLimit::Bytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed,
            maximum,
            path,
        },
        annotation_author_codec::DecodeLimit::OutputBytes { observed, maximum }
        | annotation_author_codec::DecodeLimit::Retained { observed, maximum } => {
            Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                observed,
                maximum,
                path,
            }
        },
        annotation_author_codec::DecodeLimit::Fields { observed, maximum } => {
            Error::LimitExceeded {
                kind: LimitKind::WireFields,
                observed,
                maximum,
                path,
            }
        },
        annotation_author_codec::DecodeLimit::Work { observed, maximum }
        | annotation_author_codec::DecodeLimit::Scratch { observed, maximum } => {
            Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed,
                maximum,
                path,
            }
        },
        annotation_author_codec::DecodeLimit::Nesting { observed, maximum } => {
            Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
                path,
            }
        },
        annotation_author_codec::DecodeLimit::References { observed, maximum } => {
            Error::LimitExceeded {
                kind: LimitKind::References,
                observed,
                maximum,
                path,
            }
        },
        annotation_author_codec::DecodeLimit::Text { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::TextBytes,
            observed,
            maximum,
            path,
        },
        annotation_author_codec::DecodeLimit::Allocations { observed, .. } => Error::Allocation {
            amount: observed,
            path,
        },
        _ => Error::InvalidSource { path },
    }
}
