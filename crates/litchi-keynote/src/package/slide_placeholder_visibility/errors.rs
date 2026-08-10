//! Content-redacted lower-layer error and source adapters.

use litchi_iwa_archive::SourceCatalog;

use super::{Error, LimitKind, Package, PhysicalSource, ReadError, keynote_placeholder_text_codec};

pub(super) fn placeholder_options(
    package: &Package,
    source: &[u8],
) -> Result<keynote_placeholder_text_codec::DecodeOptions, Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let nesting = u32::try_from(limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    Ok(keynote_placeholder_text_codec::DecodeOptions::new(
        source.len(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        nesting,
    ))
}

pub(super) fn root_preview_count(package: &Package) -> Result<usize, Error> {
    crate::package::rendering_invalidation::root_preview_deletions(
        physical_catalog(package)?.package(),
    )
    .map(|plan| plan.len())
    .map_err(map_rendering_invalidation_error)
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the semantic-only source feature adds a typed unsupported-source branch"
)]
pub(super) fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    #[cfg(feature = "internal-iwork-source")]
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
    #[cfg(not(feature = "internal-iwork-source"))]
    {
        let PhysicalSource::Package(source) = &package.state.source;
        Ok(source)
    }
}

pub(super) fn map_placeholder_error(error: keynote_placeholder_text_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.message_byte_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.recursion_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

pub(super) fn map_rendering_invalidation_error(
    error: crate::package::rendering_invalidation::RenderingInvalidationError,
) -> Error {
    match error {
        crate::package::rendering_invalidation::RenderingInvalidationError::InvalidSource => {
            Error::InvalidSource
        },
        crate::package::rendering_invalidation::RenderingInvalidationError::Allocation {
            amount,
        } => Error::Allocation { amount },
    }
}

pub(super) fn map_slide_preview_error(
    invalidation: crate::package::slide_preview::InvalidationError,
) -> Error {
    match invalidation {
        crate::package::slide_preview::InvalidationError::InvalidSource => Error::InvalidSource,
        crate::package::slide_preview::InvalidationError::Wire(error) => map_wire_error(error),
        crate::package::slide_preview::InvalidationError::Archive(error) => map_core_error(error),
    }
}

pub(super) fn map_read_error(read_error: ReadError) -> Error {
    match read_error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                crate::package::SemanticLimitKind::Objects => LimitKind::Entries,
                crate::package::SemanticLimitKind::Slides => LimitKind::Slides,
                crate::package::SemanticLimitKind::References => LimitKind::References,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                crate::package::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                crate::package::PayloadLimitKind::Fields => LimitKind::WireFields,
                crate::package::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                crate::package::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
    }
}

pub(super) fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> Error {
    match archive_error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

pub(super) fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::EntryBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

pub(super) fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::Entries,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}
