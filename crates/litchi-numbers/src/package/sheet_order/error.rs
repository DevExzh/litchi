use super::super::{Error as ReadError, SemanticLimitKind};
use super::{Error, LimitKind};

#[expect(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
pub(super) fn map_sheet_order_codec_error(
    error: litchi_iwa_protos::numbers_sheet_order_codec::DecodeError,
) -> Error {
    use litchi_iwa_protos::numbers_sheet_order_codec::DecodeLimit;

    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::References { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::PayloadReferences,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(_) | None => Error::InvalidSource,
    }
}

pub(super) fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

pub(super) fn map_candidate_read_error(error: ReadError) -> Error {
    match error {
        ReadError::InputTooLarge { observed, maximum }
        | ReadError::Archive(litchi_iwa_archive::Error::Limit {
            kind: litchi_iwa_archive::LimitKind::InputBytes,
            observed,
            maximum,
        }) => Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed,
            maximum,
        },
        other => map_read_error(other),
    }
}

pub(super) fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Common(common_error) => map_wire_error(common_error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::PayloadObjects,
                SemanticLimitKind::Sheets => LimitKind::Sheets,
                SemanticLimitKind::References => LimitKind::PayloadReferences,
                SemanticLimitKind::FormulaRenderDepth | SemanticLimitKind::FormulaDepth => {
                    LimitKind::WireNesting
                },
                SemanticLimitKind::FormulaRenderWork | SemanticLimitKind::FormulaWork => {
                    LimitKind::WireWork
                },
                SemanticLimitKind::OutputTextBytes
                | SemanticLimitKind::FormulaWireBytes
                | SemanticLimitKind::TextBytes => LimitKind::PayloadBytes,
                _ => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        ReadError::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
        },
        _ => Error::InvalidSource,
    }
}

pub(super) fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        _ => Error::InvalidSource,
    }
}

#[expect(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
pub(super) fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                _ => LimitKind::PayloadBytes,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

#[expect(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
pub(super) fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}
