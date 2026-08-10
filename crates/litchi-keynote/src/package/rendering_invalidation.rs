//! Root rendering-cache policy shared by focused Keynote transactions.
//!
//! Native rendering-change evidence makes the ownership boundary intentionally
//! narrow: the caller may invalidate the three root JPEG previews while
//! preserving every slide component and slide-node playback cache.

use litchi_iwa_archive::package::{Catalog, Entry};
use thiserror::Error;

const ROOT_PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A private, content-redacted rendering-cache planning failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub(super) enum RenderingInvalidationError {
    #[error("the Keynote root rendering-cache topology is ambiguous")]
    InvalidSource,
    #[error("could not allocate {amount} entries for Keynote rendering-cache invalidation")]
    Allocation { amount: usize },
}

/// Exact existing root previews selected for deletion.
///
/// The fixed native member names remain private to the package adapter; no
/// object identifiers or arbitrary caller-supplied package paths cross this
/// boundary.
pub(super) struct RootPreviewPlan {
    names: Box<[&'static str]>,
}

impl std::fmt::Debug for RootPreviewPlan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("RootPreviewPlan")
            .field("count", &self.names.len())
            .finish()
    }
}

impl RootPreviewPlan {
    /// Borrow the exact normalized root-member names authorized for deletion.
    #[must_use]
    pub(super) fn names(&self) -> &[&'static str] {
        &self.names
    }

    /// Return how many of the three optional previews exist in the source.
    #[must_use]
    pub(super) fn len(&self) -> usize {
        self.names.len()
    }
}

/// Select existing root previews and reject duplicate normalized ownership.
pub(super) fn root_preview_deletions(
    catalog: &Catalog,
) -> Result<RootPreviewPlan, RenderingInvalidationError> {
    let names = select_names(catalog.iter().map(Entry::name))?;
    Ok(RootPreviewPlan { names })
}

/// Return whether all root previews are absent, rejecting duplicate topology.
pub(super) fn root_previews_absent(catalog: &Catalog) -> Result<bool, RenderingInvalidationError> {
    Ok(collect_root_previews(catalog)?.iter().all(Option::is_none))
}

/// Verify that a playback-only transaction preserved every root preview.
///
/// Presence, decoded or opaque payload bytes, normalized and raw names,
/// compression/header metadata, and the exact compressed stream must agree.
/// Central-directory offsets are intentionally excluded because an unrelated
/// edited member may move an otherwise byte-identical retained entry.
pub(super) fn root_previews_preserved(
    source: &Catalog,
    candidate: &Catalog,
) -> Result<bool, RenderingInvalidationError> {
    let source_previews = collect_root_previews(source)?;
    let candidate_previews = collect_root_previews(candidate)?;
    for (source_preview, candidate_preview) in source_previews.into_iter().zip(candidate_previews) {
        match (source_preview, candidate_preview) {
            (None, None) => {},
            (Some(source_entry), Some(candidate_entry))
                if preview_entry_preserved(source_entry, candidate_entry) => {},
            _ => return Ok(false),
        }
    }
    Ok(true)
}

fn select_names<'a>(
    names: impl IntoIterator<Item = &'a str>,
) -> Result<Box<[&'static str]>, RenderingInvalidationError> {
    let mut seen = [false; ROOT_PREVIEW_NAMES.len()];
    for name in names {
        let Some(index) = ROOT_PREVIEW_NAMES
            .iter()
            .position(|candidate| *candidate == name)
        else {
            continue;
        };
        if std::mem::replace(&mut seen[index], true) {
            return Err(RenderingInvalidationError::InvalidSource);
        }
    }
    let count = seen.iter().filter(|present| **present).count();
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(count)
        .map_err(|_allocation| RenderingInvalidationError::Allocation { amount: count })?;
    selected.extend(
        ROOT_PREVIEW_NAMES
            .iter()
            .zip(seen)
            .filter_map(|(&name, present)| present.then_some(name)),
    );
    Ok(selected.into_boxed_slice())
}

fn collect_root_previews(
    catalog: &Catalog,
) -> Result<[Option<&Entry>; ROOT_PREVIEW_NAMES.len()], RenderingInvalidationError> {
    let mut previews = [None; ROOT_PREVIEW_NAMES.len()];
    for entry in catalog.iter() {
        let Some(index) = ROOT_PREVIEW_NAMES
            .iter()
            .position(|name| *name == entry.name())
        else {
            continue;
        };
        if previews[index].replace(entry).is_some() {
            return Err(RenderingInvalidationError::InvalidSource);
        }
    }
    Ok(previews)
}

fn preview_entry_preserved(source: &Entry, candidate: &Entry) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.data() == candidate.data()
        && source.metadata() == candidate.metadata()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && source.raw_record().compressed_data() == candidate.raw_record().compressed_data()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

pub(super) fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;

    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..LOCAL_HEADER_OFFSET.start] == candidate[..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::{Limits, package::EntryEdit};

    use super::*;

    #[test]
    fn selects_only_exact_root_previews_in_native_order() -> Result<(), RenderingInvalidationError>
    {
        let selected = select_names([
            "Index/preview.jpg",
            "preview-web.jpg",
            "preview.jpg",
            "Metadata/preview-micro.jpg",
        ])?;
        assert_eq!(selected.as_ref(), ["preview.jpg", "preview-web.jpg"]);
        Ok(())
    }

    #[test]
    fn duplicate_root_preview_is_ambiguous() {
        assert_eq!(
            select_names(["preview.jpg", "preview.jpg"]),
            Err(RenderingInvalidationError::InvalidSource)
        );
    }

    #[test]
    fn deletion_plan_removes_all_existing_root_previews() -> Result<(), Box<dyn std::error::Error>>
    {
        let source_bytes = litchi_iwa_archive::package::to_bytes(
            [
                ("preview.jpg", b"large".as_slice()),
                ("preview-micro.jpg", b"micro".as_slice()),
                ("preview-web.jpg", b"web".as_slice()),
                ("Index/preview.jpg", b"nested".as_slice()),
            ],
            Limits::default(),
        )?;
        let source = Catalog::from_bytes(&source_bytes)?;
        let plan = root_preview_deletions(&source)?;
        assert_eq!(plan.len(), 3);

        let candidate_bytes =
            source.reassemble_with_deletions_to_bytes(&[], plan.names(), Limits::default())?;
        let candidate = Catalog::from_bytes(&candidate_bytes)?;
        assert!(root_previews_absent(&candidate)?);
        assert!(
            candidate
                .iter()
                .any(|entry| entry.name() == "Index/preview.jpg")
        );
        Ok(())
    }

    #[test]
    fn unrelated_edit_preserves_previews_but_preview_edit_does_not()
    -> Result<(), Box<dyn std::error::Error>> {
        let source_bytes = litchi_iwa_archive::package::to_bytes(
            [
                ("Index/Document.iwa", b"before".as_slice()),
                ("preview.jpg", b"large".as_slice()),
                ("preview-micro.jpg", b"micro".as_slice()),
            ],
            Limits::default(),
        )?;
        let source = Catalog::from_bytes(&source_bytes)?;
        let playback_bytes = source.reassemble_to_bytes(
            &[EntryEdit::new(
                "Index/Document.iwa",
                b"a much longer replacement payload",
            )],
            Limits::default(),
        )?;
        let playback = Catalog::from_bytes(&playback_bytes)?;
        assert!(root_previews_preserved(&source, &playback)?);

        let changed_preview_bytes = source.reassemble_to_bytes(
            &[EntryEdit::new("preview.jpg", b"different")],
            Limits::default(),
        )?;
        let changed_preview = Catalog::from_bytes(&changed_preview_bytes)?;
        assert!(!root_previews_preserved(&source, &changed_preview)?);
        Ok(())
    }

    #[test]
    fn central_record_comparison_ignores_only_the_relocated_local_offset() {
        let source = [0_u8; 50];
        let mut relocated = source;
        relocated[42..46].copy_from_slice(&123_u32.to_le_bytes());
        assert!(central_record_preserved(&source, &relocated));

        relocated[38] = 1;
        assert!(!central_record_preserved(&source, &relocated));
    }
}
