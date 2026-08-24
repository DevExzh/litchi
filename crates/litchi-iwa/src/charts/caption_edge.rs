//! Bounded, format-neutral chart-caption reference access.
//!
//! The chart archive remains an opaque source-owned payload at this boundary.
//! Only the selected `TSP.Reference.identifier` edge is projected or
//! rewritten by the strict Buffa-backed codec; generated protobuf references
//! do not cross into chart owners.

use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::chart_caption_codec::{
    self as chart_caption_codec, ChartCaptionWrite, DecodeError, DecodeOptions, WireResourceLimit,
};

use crate::package::PackageLimits;
use crate::{Error, Result};

const CHART_CAPTION_RECURSION_LIMIT: u32 = 8;
const CHART_CAPTION_FIELD_MULTIPLIER: usize = 4;
const CHART_CAPTION_WORK_MULTIPLIER: usize = 64;

/// Decode the selected chart-caption edge without exposing generated wire
/// types to host graph owners.
pub(crate) fn chart_caption_identifier(
    limits: PackageLimits,
    source: &[u8],
) -> Result<Option<u64>> {
    let options = chart_caption_decode_options(limits, source, false)?;
    chart_caption_codec::decode_chart_caption_identifier(source, options)
        .map_err(map_chart_caption_error)
}

/// Rewrite the selected chart-caption edge and return the verified candidate
/// payload. The codec report is consumed internally so callers receive only
/// semantic bytes and the package's normal typed error surface.
pub(crate) fn rewrite_chart_caption_identifier(
    limits: PackageLimits,
    source: &[u8],
    identifier: u64,
) -> Result<Vec<u8>> {
    let options = chart_caption_decode_options(limits, source, true)?;
    chart_caption_codec::rewrite_chart_caption_with_report(
        source,
        ChartCaptionWrite::new(identifier),
        options,
    )
    .map(|(candidate, _report)| candidate)
    .map_err(map_chart_caption_error)
}

fn chart_caption_decode_options(
    limits: PackageLimits,
    source: &[u8],
    rewrite: bool,
) -> Result<DecodeOptions> {
    let stream_limit = limits
        .max_iwa_stream_bytes()
        .min(WireLimits::MAX_INPUT_BYTES)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let message_limit = limits
        .archive_limits()
        .max_message_bytes()
        .min(stream_limit)
        .min(WireLimits::MAX_INPUT_BYTES)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let max_message_bytes = source.len().max(1).min(message_limit);
    let max_fields = source
        .len()
        .saturating_mul(CHART_CAPTION_FIELD_MULTIPLIER)
        .clamp(1, WireLimits::MAX_FIELDS);
    let max_work_bytes = source
        .len()
        .saturating_mul(CHART_CAPTION_WORK_MULTIPLIER)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let max_output_bytes = if rewrite {
        message_limit.min(WireLimits::MAX_OUTPUT_BYTES)
    } else {
        source.len().max(1).min(stream_limit)
    };

    Ok(DecodeOptions::new(
        max_message_bytes,
        max_fields,
        max_work_bytes,
        CHART_CAPTION_RECURSION_LIMIT,
    )
    .with_max_output_bytes(max_output_bytes))
}

fn map_chart_caption_error(error: DecodeError) -> Error {
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
    if let Some((observed, maximum)) = error.output_limit_values() {
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed,
            limit: maximum,
        });
    }
    if let Some(amount) = error.allocation_amount() {
        return Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork chart-caption output",
            amount,
        });
    }
    match error.wire_resource_limit() {
        Some(WireResourceLimit::Bytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit: maximum,
            })
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                limit: usize::try_from(maximum).unwrap_or(usize::MAX),
            })
        },
        None => Error::InvalidFormat(format!(
            "iWork chart-caption edge failed strict validation: {error}"
        )),
        Some(_) => Error::InvalidFormat(format!(
            "iWork chart-caption edge failed strict validation: {error}"
        )),
    }
}
