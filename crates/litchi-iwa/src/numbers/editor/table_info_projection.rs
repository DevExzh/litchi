//! Bounded ownership projection for Numbers `TableInfoArchive` messages.
//!
//! The legacy editor has a few discovery paths that need only the required
//! table-model edge.  Keep those paths away from a complete generated
//! `TableInfoArchive`; mutation and geometry paths still use the generated
//! value when they need its complete state.  The accepted source bytes remain
//! the authority for every later rewrite.

use std::num::NonZeroU64;

use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::table_info_codec;

use crate::{Error, Result};

const TABLE_INFO_PROJECTION_RECURSION_LIMIT: u32 = 64;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const LEGACY_TABLE_INFO_SUPER_PREFIX: [u8; 2] = [0x0a, 0x00];

/// Project only the required native table-model edge from one `TableInfo`.
///
/// The strict raw scan and Buffa lazy view are both bounded by the source
/// length and the common hard ceilings.  Unknown fields are validated and
/// ignored by the private projection; the original message remains available
/// to source-preserving callers.  A non-zero typed reference prevents a
/// malformed zero edge from becoming an ownership candidate.
pub(super) fn model_reference(source: &[u8]) -> Result<NonZeroU64> {
    decode_model_reference(source, decode_options(source), false)
}

/// Project the model edge for one typed Numbers table-info message.
///
/// The historical type-6003 alias may omit the empty `DrawableArchive`
/// envelope.  Canonical type 6000 remains strict: it must carry `super`, even
/// when the remaining fields would otherwise look like a valid table-info
/// payload.  Keeping the message type at this boundary prevents a malformed
/// canonical candidate from receiving the legacy compatibility envelope.
pub(super) fn model_reference_for_type(message_type: u32, source: &[u8]) -> Result<NonZeroU64> {
    if message_type != LEGACY_TABLE_INFO_MESSAGE_TYPE {
        return model_reference(source);
    }
    decode_model_reference(source, decode_options(source), true)
}

fn decode_model_reference(
    source: &[u8],
    options: table_info_codec::DecodeOptions,
    allow_legacy_super_omission: bool,
) -> Result<NonZeroU64> {
    // The public codec currently validates the complete TableInfo envelope,
    // while a historical type-6003 payload may omit its empty Drawable
    // `super` field.  Probe the source first so byte and wire errors retain
    // their original limits, then add only the compatibility envelope when
    // the codec reports that `super` is absent.  The source itself is never
    // changed or retained by this projection.
    match table_info_codec::decode_table_model_reference(source, options) {
        Ok(reference) => Ok(reference.identifier()),
        Err(error)
            if allow_legacy_super_omission
                && error.missing_required_field() == Some("TST.TableInfoArchive.super") =>
        {
            let capacity = source
                .len()
                .checked_add(LEGACY_TABLE_INFO_SUPER_PREFIX.len())
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers legacy table-info projection source size overflowed".to_owned(),
                    )
                })?;
            let mut compatibility = Vec::new();
            compatibility
                .try_reserve_exact(capacity)
                .map_err(|_error| {
                    Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                        resource: "Numbers legacy table-info projection source",
                        amount: capacity,
                    })
                })?;
            compatibility.extend_from_slice(&LEGACY_TABLE_INFO_SUPER_PREFIX);
            compatibility.extend_from_slice(source);
            let projected_options = compatibility_options(options, source.len(), capacity);
            table_info_codec::decode_table_model_reference(
                compatibility.as_slice(),
                projected_options,
            )
            .map(|reference| reference.identifier())
            .map_err(map_decode_error)
        },
        Err(error) => Err(map_decode_error(error)),
    }
}

fn compatibility_options(
    options: table_info_codec::DecodeOptions,
    source_len: usize,
    projected_len: usize,
) -> table_info_codec::DecodeOptions {
    // `decode_options` derives all three ceilings from the source width.  Grow
    // those derived ceilings for the two-byte compatibility envelope, while
    // retaining a caller's intentionally smaller explicit limit.  The hard
    // common ceilings remain in force even for an unusually large source.
    let source_bytes = source_len.clamp(1, WireLimits::MAX_INPUT_BYTES);
    let projected_bytes = projected_len.clamp(1, WireLimits::MAX_INPUT_BYTES);
    let source_fields = source_len.clamp(1, WireLimits::MAX_FIELDS);
    let projected_fields = projected_len.clamp(1, WireLimits::MAX_FIELDS);
    let source_work = source_len
        .saturating_mul(4)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let projected_work = projected_len
        .saturating_mul(4)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);

    options
        .with_max_message_bytes(expand_derived_limit(
            options.max_message_bytes(),
            source_bytes,
            projected_bytes,
        ))
        .with_max_fields(expand_derived_limit(
            options.max_fields(),
            source_fields,
            projected_fields,
        ))
        .with_max_work_bytes(expand_derived_limit(
            options.max_work_bytes(),
            source_work,
            projected_work,
        ))
}

const fn expand_derived_limit(current: usize, source: usize, projected: usize) -> usize {
    if current >= source {
        if projected > current {
            projected
        } else {
            current
        }
    } else {
        current
    }
}

fn decode_options(source: &[u8]) -> table_info_codec::DecodeOptions {
    table_info_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(4)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        TABLE_INFO_PROJECTION_RECURSION_LIMIT,
    )
}

fn map_decode_error(error: table_info_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed,
            limit: maximum,
        });
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed,
            limit: maximum,
        });
    }
    match error.wire_resource_limit() {
        Some(table_info_codec::WireResourceLimit::Bytes { observed, maximum }) => {
            let limit = maximum.unwrap_or(WireLimits::MAX_INPUT_BYTES);
            let observed = observed.unwrap_or_else(|| limit.saturating_add(1));
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit,
            })
        },
        Some(table_info_codec::WireResourceLimit::Nesting { observed, maximum }) => {
            let limit = usize::try_from(maximum.unwrap_or(TABLE_INFO_PROJECTION_RECURSION_LIMIT))
                .unwrap_or(usize::MAX);
            let observed = usize::try_from(
                observed.unwrap_or_else(|| TABLE_INFO_PROJECTION_RECURSION_LIMIT.saturating_add(1)),
            )
            .unwrap_or(usize::MAX);
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed,
                limit,
            })
        },
        Some(_) => Error::InvalidFormat(format!(
            "Numbers TableInfo model-reference projection failed: {error}"
        )),
        None => Error::InvalidFormat(format!(
            "Numbers TableInfo model-reference projection failed: {error}"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source() -> Vec<u8> {
        // Historical type 6003 payloads may carry only the model edge.
        vec![0x12, 0x02, 0x08, 0x2a]
    }

    fn canonical_source() -> Vec<u8> {
        [vec![0x0a, 0x00], source()].concat()
    }

    #[test]
    fn projects_legacy_type_6003_without_a_super_envelope() {
        assert_eq!(
            model_reference_for_type(LEGACY_TABLE_INFO_MESSAGE_TYPE, &source())
                .expect("valid legacy projection")
                .get(),
            42
        );
    }

    #[test]
    fn canonical_type_6000_requires_the_super_envelope() {
        assert!(model_reference_for_type(6_000, &source()).is_err());
        assert_eq!(
            model_reference_for_type(6_000, &canonical_source())
                .expect("valid canonical projection")
                .get(),
            42
        );
    }

    #[test]
    fn canonical_unknown_fields_stay_opaque() {
        let mut payload = canonical_source();
        payload.extend_from_slice(&[0x18, 0x7f]);
        assert_eq!(
            model_reference(&payload)
                .expect("unknown field is opaque")
                .get(),
            42
        );

        let mut noncanonical = canonical_source();
        noncanonical.extend_from_slice(&[0x18, 0x81, 0x00]);
        assert!(model_reference(&noncanonical).is_err());
    }

    #[test]
    fn malformed_required_edge_is_rejected_before_candidate_use() {
        assert!(model_reference(&[0x0a, 0x00]).is_err());
        assert!(model_reference(&[0x12, 0x02, 0x08, 0x00]).is_err());
        // Known but unselected parent metadata still has canonical framing
        // requirements when it is present.
        assert!(model_reference(&[0x08, 0x01, 0x12, 0x02, 0x08, 0x2a]).is_err());
    }

    #[test]
    fn typed_limits_are_preserved() {
        let payload = source();

        let error = decode_model_reference(
            &payload,
            table_info_codec::DecodeOptions::new(payload.len() - 1, 3, 16, 2),
            true,
        )
        .expect_err("byte limit");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                ..
            })
        ));

        let error = decode_model_reference(
            &payload,
            table_info_codec::DecodeOptions::new(payload.len(), 1, 16, 2),
            true,
        )
        .expect_err("field limit");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                ..
            })
        ));

        let error = decode_model_reference(
            &payload,
            table_info_codec::DecodeOptions::new(payload.len(), 3, 16, 0),
            true,
        )
        .expect_err("nesting limit");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                ..
            })
        ));
    }
}
