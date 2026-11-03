//! The chart-lifecycle values every format stores on its chart drawable
//! (record 0416): the native chart kind and series direction in the
//! drawable's `TSCH.ChartArchive` extension, read and rewritten in place
//! (the host's `set_*_chart_kind` and `set_*_chart_direction`, which assign
//! `chart_type` and `series_direction`). Every other byte of the payload is
//! kept. [`author`] writes the private graph of a new chart; [`graph`]
//! lists and remaps the private references of an existing one.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::chart::Direction;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::wire::{WireView, patch_varint_field, transform_length_delimited_field};

pub mod author;
pub mod graph;

/// The `TSCH.ChartArchive` extension of `TSCH.ChartDrawableArchive`.
const CHART_UNITY: u32 = 10_000;
const CHART_TYPE: u32 = 1;
const CHART_SERIES_DIRECTION: u32 = 5;

/// Why a chart-lifecycle value could not be read or written.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The chart payload is malformed or over the caller's limits.
    #[error(transparent)]
    Wire(#[from] litchi_iwa_common::Error),
    /// The payload has no chart extension, or a value is repeated or
    /// mistyped.
    #[error("the chart drawable has no well-formed chart archive")]
    Invalid,
    /// The chart data cannot form the requested kind, or its grid cannot be
    /// encoded.
    #[error("the chart data cannot form the chart")]
    InvalidData,
    /// A new chart object exceeds the archive limits.
    #[error("a new chart object exceeds the archive limits")]
    Object,
}

/// The kind and series direction a chart drawable payload stores (`None`
/// when the field is absent).
///
/// # Errors
///
/// [`Error`] for a malformed payload.
pub fn values(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(Option<Kind>, Option<Direction>), Error> {
    let view = WireView::parse(payload, limits)?;
    let mut unity = None;
    for field in view.fields() {
        if field.number() == CHART_UNITY {
            if field.wire_type() != 2 || unity.replace(field.payload()).is_some() {
                return Err(Error::Invalid);
            }
        }
    }
    let unity = WireView::parse(unity.ok_or(Error::Invalid)?, limits)?;
    let (mut kind, mut direction) = (None, None);
    for field in unity.fields() {
        let slot = match field.number() {
            CHART_TYPE => &mut kind,
            CHART_SERIES_DIRECTION => &mut direction,
            _ => continue,
        };
        if field.wire_type() != 0 || slot.is_some() {
            return Err(Error::Invalid);
        }
        let (value, length) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
            .map_err(|_error| Error::Invalid)?;
        if length != field.payload().len() {
            return Err(Error::Invalid);
        }
        *slot = Some(value);
    }
    let native = |value: u64| i32::try_from(value as i64).map_err(|_error| Error::Invalid);
    Ok((
        kind.map(native).transpose()?.map(Kind::from_native),
        direction
            .map(native)
            .transpose()?
            .map(Direction::from_native),
    ))
}

/// Whether `kind` can be assigned to a chart (a supported concrete kind,
/// the host's `require_creatable_kind`).
#[must_use]
pub const fn assignable_kind(kind: Kind) -> bool {
    kind.native_value() != Kind::Undefined.native_value() && !kind.is_unsupported()
}

#[allow(
    clippy::cast_sign_loss,
    reason = "protobuf enum varints are sign-extended"
)]
fn native(value: i32) -> u64 {
    i64::from(value) as u64
}

fn with_value(payload: &[u8], number: u32, value: u64, present: bool) -> Result<Vec<u8>, Error> {
    Ok(transform_length_delimited_field(
        payload,
        CHART_UNITY,
        |unity| patch_varint_field(unity, number, present, Some(value)),
    )?)
}

/// `payload` with the chart kind set to `kind`.
///
/// # Errors
///
/// [`Error`] for a malformed payload.
pub fn with_kind(payload: &[u8], kind: Kind, limits: WireLimits) -> Result<Vec<u8>, Error> {
    let (current, _) = values(payload, limits)?;
    with_value(
        payload,
        CHART_TYPE,
        native(kind.native_value()),
        current.is_some(),
    )
}

/// `payload` with the series direction set to `direction`.
///
/// # Errors
///
/// [`Error`] for a malformed payload.
pub fn with_direction(
    payload: &[u8],
    direction: Direction,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let (_, current) = values(payload, limits)?;
    with_value(
        payload,
        CHART_SERIES_DIRECTION,
        native(direction.native_value()),
        current.is_some(),
    )
}
