//! Atomic staging for chunked same-slide replacement batches.

use crate::{Error, Result};

const MAX_CHUNK_COUNT: usize = crate::shape::Limits::DEFAULT.shapes();
const MAX_SELECTIONS: usize = crate::shape::Limits::DEFAULT.shapes();

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ChunkAdmission {
    selections: usize,
    text_bytes: usize,
}

impl ChunkAdmission {
    #[inline]
    #[must_use]
    pub(crate) const fn selections(self) -> usize {
        self.selections
    }

    #[inline]
    #[must_use]
    pub(crate) const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

#[derive(Clone, Copy)]
struct ChunkPolicy {
    chunk_kind: crate::error::CodecLimitKind,
    plan_resource: &'static str,
    per_chunk_limit: usize,
    selection_kind: crate::error::CodecLimitKind,
    max_text_bytes: usize,
    max_output_bytes: usize,
}

fn text_policy(max_text_bytes: usize, max_output_bytes: usize) -> ChunkPolicy {
    ChunkPolicy {
        chunk_kind: crate::error::CodecLimitKind::TextReplacements,
        plan_resource: "opened-presentation shape text chunk plan",
        per_chunk_limit: super::MAX_SHAPE_TEXT_REPLACEMENTS,
        selection_kind: crate::error::CodecLimitKind::TextReplacements,
        max_text_bytes,
        max_output_bytes,
    }
}

fn property_policy(max_text_bytes: usize, max_output_bytes: usize) -> ChunkPolicy {
    ChunkPolicy {
        chunk_kind: crate::error::CodecLimitKind::ShapeProperties,
        plan_resource: "opened-presentation shape property chunk plan",
        per_chunk_limit: super::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
        selection_kind: crate::error::CodecLimitKind::ShapeProperties,
        max_text_bytes,
        max_output_bytes,
    }
}

/// Stage several text chunks as one same-slide operation.
///
/// Each chunk retains the regular 256-selection and aggregate text limits.
/// The flattened operation is admitted against the slide shape count and the
/// output budget, then resolved and published once so duplicate selectors and
/// overlapping shape spans are checked globally.
#[cfg(test)]
pub(crate) fn stage_shape_text_chunks<'a>(
    xml: &[u8],
    chunks: &[&[super::ShapeTextReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<(Option<Vec<u8>>, usize)> {
    stage_shape_text_chunks_with_limits(
        xml,
        chunks,
        max_text_bytes,
        max_output_bytes,
        crate::shape::Limits::DEFAULT,
    )
}

pub(crate) fn stage_shape_text_chunks_with_limits<'a>(
    xml: &[u8],
    chunks: &[&[super::ShapeTextReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let policy = text_policy(max_text_bytes, max_output_bytes);
    let admission = admit_chunks(chunks, |replacement| replacement.text().len(), policy)?;
    stage_shape_text_chunks_admitted_with_limits(
        xml,
        chunks,
        max_text_bytes,
        max_output_bytes,
        admission,
        shape_limits,
    )
}

pub(crate) fn stage_shape_text_chunks_admitted_with_limits<'a>(
    xml: &[u8],
    chunks: &[&[super::ShapeTextReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
    admission: ChunkAdmission,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let policy = text_policy(max_text_bytes, max_output_bytes);
    let replacements = flatten_chunks(chunks, policy, admission)?;
    super::xml::stage_shape_texts_admitted_with_limits(
        xml,
        &replacements,
        max_text_bytes,
        max_output_bytes,
        MAX_SELECTIONS,
        max_output_bytes,
        shape_limits,
    )
}

pub(crate) fn stage_shape_property_chunks_with_limits<'a>(
    xml: &[u8],
    chunks: &[&[super::ShapePropertiesReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let policy = property_policy(max_text_bytes, max_output_bytes);
    let admission = admit_chunks(
        chunks,
        |replacement| replacement.text().map_or(0, str::len),
        policy,
    )?;
    stage_shape_property_chunks_admitted_with_limits(
        xml,
        chunks,
        max_text_bytes,
        max_output_bytes,
        admission,
        shape_limits,
    )
}

pub(crate) fn stage_shape_property_chunks_admitted_with_limits<'a>(
    xml: &[u8],
    chunks: &[&[super::ShapePropertiesReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
    admission: ChunkAdmission,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let policy = property_policy(max_text_bytes, max_output_bytes);
    let replacements = flatten_chunks(chunks, policy, admission)?;
    super::properties::stage_shape_properties_admitted_with_limits(
        xml,
        &replacements,
        max_text_bytes,
        max_output_bytes,
        MAX_SELECTIONS,
        max_output_bytes,
        shape_limits,
    )
}

/// Admit text chunks without allocating replacement descriptors. The returned
/// scalar admission is suitable for a managed caller's reservation before it
/// asks the stage function to flatten and publish the operation.
pub(crate) fn admit_shape_text_chunks<'a>(
    chunks: &[&[super::ShapeTextReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<ChunkAdmission> {
    admit_chunks(
        chunks,
        |replacement| replacement.text().len(),
        text_policy(max_text_bytes, max_output_bytes),
    )
}

/// Admit property chunks without allocating replacement descriptors. Text is
/// counted only for updates that carry a text value; geometry-only updates do
/// not consume the text budget.
pub(crate) fn admit_shape_property_chunks<'a>(
    chunks: &[&[super::ShapePropertiesReplacement<'a>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<ChunkAdmission> {
    admit_chunks(
        chunks,
        |replacement| replacement.text().map_or(0, str::len),
        property_policy(max_text_bytes, max_output_bytes),
    )
}

/// Validate chunk-local limits and admit the whole operation without
/// allocating replacement descriptors. `item_bytes` returns the replacement
/// text size (zero for a property update without text).
fn admit_chunks<T: Copy>(
    chunks: &[&[T]],
    item_bytes: impl Fn(T) -> usize,
    policy: ChunkPolicy,
) -> Result<ChunkAdmission> {
    if chunks.len() > MAX_CHUNK_COUNT {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(policy.chunk_kind),
            observed: chunks.len(),
            maximum: MAX_CHUNK_COUNT,
        });
    }

    let mut total_selections = 0usize;
    let mut total_text_bytes = 0usize;
    for chunk in chunks {
        if chunk.len() > policy.per_chunk_limit {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(policy.selection_kind),
                observed: chunk.len(),
                maximum: policy.per_chunk_limit,
            });
        }
        total_selections = total_selections
            .checked_add(chunk.len())
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(policy.selection_kind),
                observed: usize::MAX,
                maximum: MAX_SELECTIONS,
            })?;
        if total_selections > MAX_SELECTIONS {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(policy.selection_kind),
                observed: total_selections,
                maximum: MAX_SELECTIONS,
            });
        }

        let chunk_text_bytes = chunk.iter().try_fold(0usize, |total, item| {
            total
                .checked_add(item_bytes(*item))
                .ok_or_else(invalid_aggregate_text_size)
        })?;
        if chunk_text_bytes > policy.max_text_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: chunk_text_bytes,
                maximum: policy.max_text_bytes,
            });
        }
        total_text_bytes = total_text_bytes
            .checked_add(chunk_text_bytes)
            .ok_or_else(invalid_aggregate_text_size)?;
        if total_text_bytes > policy.max_output_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: total_text_bytes,
                maximum: policy.max_output_bytes,
            });
        }
    }

    Ok(ChunkAdmission {
        selections: total_selections,
        text_bytes: total_text_bytes,
    })
}

/// Flatten admitted borrowed descriptors exactly once. The private caller
/// provides the scalar result of `admit_chunks`, so this pass only reserves and
/// copies descriptors and cannot allocate based on unvalidated input.
fn flatten_chunks<T: Copy>(
    chunks: &[&[T]],
    policy: ChunkPolicy,
    admission: ChunkAdmission,
) -> Result<Vec<T>> {
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(admission.selections())
        .map_err(|source| Error::Allocation {
            resource: policy.plan_resource,
            source,
        })?;
    for chunk in chunks {
        replacements.extend_from_slice(chunk);
    }
    Ok(replacements)
}

#[inline]
fn invalid_aggregate_text_size() -> Error {
    Error::Invalid("opened-presentation aggregate shape text size overflow".into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::opened::ShapeTextReplacement;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn slide() -> Vec<u8> {
        format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="First"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>one</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Second"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>two</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes()
    }

    fn is_limit(error: &Error, kind: crate::error::CodecLimitKind, limit: usize) -> bool {
        matches!(
            error,
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(actual_kind),
                maximum: actual_limit,
                ..
            } if *actual_limit == limit && *actual_kind == kind
        )
    }

    #[test]
    fn each_chunk_keeps_the_legacy_selection_and_text_ceiling() {
        let replacements =
            vec![ShapeTextReplacement::at(0, "x"); super::super::MAX_SHAPE_TEXT_REPLACEMENTS + 1];
        let chunks: [&[ShapeTextReplacement<'_>]; 1] = [&replacements];
        let error = admit_chunks(
            &chunks,
            |replacement| replacement.text().len(),
            text_policy(usize::MAX, usize::MAX),
        )
        .expect_err("a chunk over 256 selections must be refused");
        assert!(is_limit(
            &error,
            crate::error::CodecLimitKind::TextReplacements,
            super::super::MAX_SHAPE_TEXT_REPLACEMENTS
        ));

        let replacements = [ShapeTextReplacement::at(0, "12345")];
        let chunks = [&replacements[..]];
        let error = admit_chunks(
            &chunks,
            |replacement| replacement.text().len(),
            text_policy(4, usize::MAX),
        )
        .expect_err("a chunk over the text ceiling must be refused");
        assert!(is_limit(&error, crate::error::CodecLimitKind::TextBytes, 4));
    }

    #[test]
    fn chunk_total_text_budget_is_checked_before_flattening() {
        let first = [ShapeTextReplacement::at(0, "1234")];
        let second = [ShapeTextReplacement::at(1, "5678")];
        let chunks = [&first[..], &second[..]];
        let error = admit_chunks(
            &chunks,
            |replacement| replacement.text().len(),
            text_policy(4, 7),
        )
        .expect_err("the whole operation must honor its output budget");
        assert!(is_limit(&error, crate::error::CodecLimitKind::TextBytes, 7));
    }

    #[test]
    fn chunk_count_ceiling_is_finite() {
        let empty: &[ShapeTextReplacement<'static>] = &[];
        let chunks = vec![empty; MAX_CHUNK_COUNT + 1];
        let error = admit_shape_text_chunks(&chunks, 128, 1 << 20)
            .expect_err("an unbounded chunk list must be refused");
        assert!(is_limit(
            &error,
            crate::error::CodecLimitKind::TextReplacements,
            MAX_CHUNK_COUNT
        ));
    }

    fn selection_chunks(count: usize) -> Vec<Vec<ShapeTextReplacement<'static>>> {
        let mut chunks = Vec::new();
        let mut remaining = count;
        while remaining != 0 {
            let length = remaining.min(super::super::MAX_SHAPE_TEXT_REPLACEMENTS);
            chunks.push(vec![ShapeTextReplacement::at(0, ""); length]);
            remaining -= length;
        }
        chunks
    }

    #[test]
    fn total_shape_selection_ceiling_accepts_boundary_and_rejects_one_more() {
        let owned = selection_chunks(MAX_SELECTIONS);
        let chunks: Vec<&[ShapeTextReplacement<'static>]> =
            owned.iter().map(Vec::as_slice).collect();
        let admission = admit_shape_text_chunks(&chunks, 128, 1 << 20)
            .expect("the exact shape-count boundary must be admitted");
        assert_eq!(admission.selections(), MAX_SELECTIONS);
        assert_eq!(admission.text_bytes(), 0);

        let owned = selection_chunks(MAX_SELECTIONS + 1);
        let chunks: Vec<&[ShapeTextReplacement<'static>]> =
            owned.iter().map(Vec::as_slice).collect();
        let error = admit_shape_text_chunks(&chunks, 128, 1 << 20)
            .expect_err("one selection over the shape-count ceiling must be refused");
        assert!(is_limit(
            &error,
            crate::error::CodecLimitKind::TextReplacements,
            MAX_SELECTIONS
        ));
    }

    #[test]
    fn duplicate_selectors_across_chunks_are_one_atomic_refusal() {
        let first = [ShapeTextReplacement::at(0, "first")];
        let second = [ShapeTextReplacement::at(0, "second")];
        let chunks = [&first[..], &second[..]];
        let error = stage_shape_text_chunks(&slide(), &chunks, 128, 1 << 20)
            .expect_err("duplicate selectors must be rejected globally");
        assert!(matches!(error, Error::Refused(_)));
    }

    #[test]
    fn unchanged_chunked_text_is_a_noop() -> Result<()> {
        let replacement = [ShapeTextReplacement::at(0, "one")];
        let chunks = [&replacement[..]];
        let (candidate, changed) = stage_shape_text_chunks(&slide(), &chunks, 128, 1 << 20)?;
        assert!(candidate.is_none());
        assert_eq!(changed, 0);
        Ok(())
    }

    #[test]
    fn chunked_text_combines_chunks_into_one_candidate() -> Result<()> {
        let first = [ShapeTextReplacement::at(0, "updated first")];
        let second = [ShapeTextReplacement::at(1, "updated second")];
        let chunks = [&first[..], &second[..]];
        let (candidate, changed) = stage_shape_text_chunks(&slide(), &chunks, 128, 1 << 20)?;
        assert_eq!(changed, 2);
        let scene = crate::shape::Scene::read(
            candidate
                .as_deref()
                .ok_or_else(|| Error::Invalid("chunked stage returned no candidate".into()))?,
            crate::shape::Limits::DEFAULT,
        )?;
        assert_eq!(scene.at(0)?.common().text(), Some("updated first"));
        assert_eq!(scene.at(1)?.common().text(), Some("updated second"));
        Ok(())
    }
}
