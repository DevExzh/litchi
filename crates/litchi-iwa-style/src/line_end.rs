//! `TSD.LineEndArchive` payloads of the standard line endpoints (record
//! 0371): the host's `shapes::line_end::native::endpoint_archive`, field by
//! field in tag order, and the identifier-based read back. Lifted from the
//! Keynote leaf's line-endpoint writer (record 0372) so the Pages and
//! Numbers owners share it.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::line::{Endpoint, Endpoints};
use litchi_iwa_common::varint::encode_varint_into;
use litchi_iwa_common::wire::WireView;
use thiserror::Error as ThisError;

const LINE_END_PATH: u32 = 1;
const LINE_END_JOIN: u32 = 2;
const LINE_END_POINT: u32 = 3;
const LINE_END_FILLED: u32 = 4;
const LINE_END_IDENTIFIER: u32 = 5;
const MITER_JOIN: u64 = 0;
const ROUND_JOIN: u64 = 1;
const CANONICAL_MIN: f32 = 0.0;
const CANONICAL_MID: f32 = 3.0;
const CANONICAL_MAX: f32 = 6.0;
const CIRCLE_INSET: f32 = 0.5;
const CIRCLE_CONTROL_LOW: f32 = 1.619_288_1;
const CIRCLE_CONTROL_HIGH: f32 = 4.380_712;
const OPEN_ARROW_TIP_Y: f32 = 5.0;
const FILLED_ARROW_NOTCH_Y: f32 = 1.5;
const DIAMOND_ATTACHMENT_Y: f32 = 0.331_4;
const FILLED_SQUARE_INSET: f32 = 0.5;
const OPEN_SQUARE_INSET: f32 = 1.0;
const OPEN_SQUARE_ATTACHMENT_Y: f32 = 0.2;
const OPEN_CIRCLE_ATTACHMENT_Y: f32 = -0.3;
const LINE_ATTACHMENT_Y: f32 = -0.8;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const CURVE_TO: u64 = 4;
const CLOSE_SUBPATH: u64 = 5;

/// A line end the reader does not know.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UnknownLineEnd;

impl std::fmt::Display for UnknownLineEnd {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("the line end is not a standard endpoint")
    }
}

impl std::error::Error for UnknownLineEnd {}

/// A bounded line-end decode failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum DecodeError {
    /// The payload is malformed or names an endpoint outside the standard set.
    #[error(transparent)]
    Unknown(#[from] UnknownLineEnd),
    /// The caller's finite wire profile rejected the payload.
    #[error(transparent)]
    Wire(#[from] litchi_iwa_common::Error),
}

#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn key(mut self, number: u32, wire_type: u8) -> Self {
        encode_varint_into(&mut self.0, (u64::from(number) << 3) | u64::from(wire_type));
        self
    }

    fn varint(self, number: u32, value: u64) -> Self {
        let mut writer = self.key(number, 0);
        encode_varint_into(&mut writer.0, value);
        writer
    }

    fn float(self, number: u32, value: f32) -> Self {
        let mut writer = self.key(number, 5);
        writer.0.extend_from_slice(&value.to_le_bytes());
        writer
    }

    fn bytes(self, number: u32, payload: &[u8]) -> Self {
        let mut writer = self.key(number, 2);
        encode_varint_into(&mut writer.0, payload.len() as u64);
        writer.0.extend_from_slice(payload);
        writer
    }

    fn message(self, number: u32, message: &Self) -> Self {
        self.bytes(number, &message.0)
    }
}

fn point(x: f32, y: f32) -> Writer {
    Writer::default().float(1, x).float(2, y)
}

fn element(kind: u64, points: &[(f32, f32)]) -> Writer {
    points
        .iter()
        .fold(Writer::default().varint(1, kind), |writer, &(x, y)| {
            writer.message(2, &point(x, y))
        })
}

fn line_end(
    identifier: &str,
    join: u64,
    elements: &[Writer],
    end_point: Option<(f32, f32)>,
    filled: Option<bool>,
) -> Vec<u8> {
    let path = elements
        .iter()
        .fold(Writer::default(), |path, element| path.message(1, element));
    let mut writer = Writer::default()
        .message(LINE_END_PATH, &path)
        .varint(LINE_END_JOIN, join);
    if let Some((x, y)) = end_point {
        writer = writer.message(LINE_END_POINT, &point(x, y));
    }
    if let Some(filled) = filled {
        writer = writer.varint(LINE_END_FILLED, u64::from(filled));
    }
    writer.bytes(LINE_END_IDENTIFIER, identifier.as_bytes()).0
}

fn circle_end(identifier: &str, endpoint_y: f32, filled: Option<bool>) -> Vec<u8> {
    let low = CIRCLE_INSET;
    let high = CANONICAL_MAX - CIRCLE_INSET;
    line_end(
        identifier,
        MITER_JOIN,
        &[
            element(MOVE_TO, &[(high, CANONICAL_MID)]),
            element(
                CURVE_TO,
                &[
                    (high, CIRCLE_CONTROL_HIGH),
                    (CIRCLE_CONTROL_HIGH, high),
                    (CANONICAL_MID, high),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (CIRCLE_CONTROL_LOW, high),
                    (low, CIRCLE_CONTROL_HIGH),
                    (low, CANONICAL_MID),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (low, CIRCLE_CONTROL_LOW),
                    (CIRCLE_CONTROL_LOW, low),
                    (CANONICAL_MID, low),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (CIRCLE_CONTROL_HIGH, low),
                    (high, CIRCLE_CONTROL_LOW),
                    (high, CANONICAL_MID),
                ],
            ),
            element(CLOSE_SUBPATH, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        filled,
    )
}

fn square_end(identifier: &str, inset: f32, filled: Option<bool>, endpoint_y: f32) -> Vec<u8> {
    let far = CANONICAL_MAX - inset;
    line_end(
        identifier,
        MITER_JOIN,
        &[
            element(MOVE_TO, &[(inset, inset)]),
            element(LINE_TO, &[(far, inset)]),
            element(LINE_TO, &[(far, far)]),
            element(LINE_TO, &[(inset, far)]),
            element(CLOSE_SUBPATH, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        filled,
    )
}

/// The `TSD.LineEndArchive` payload the host writes for `endpoint`.
#[must_use]
pub fn encode(endpoint: Endpoint) -> Vec<u8> {
    match endpoint {
        Endpoint::None => line_end("none", MITER_JOIN, &[], None, None),
        Endpoint::SimpleArrow => line_end(
            "simple arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            Some(true),
        ),
        Endpoint::FilledCircle => circle_end("filled circle", CIRCLE_INSET, Some(true)),
        Endpoint::FilledDiamond => line_end(
            "filled diamond",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::OpenArrow => line_end(
            "open arrow",
            ROUND_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(MOVE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            None,
        ),
        Endpoint::FilledArrow => line_end(
            "filled arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, FILLED_ARROW_NOTCH_Y)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, FILLED_ARROW_NOTCH_Y)),
            Some(true),
        ),
        Endpoint::FilledSquare => square_end(
            "filled square",
            FILLED_SQUARE_INSET,
            Some(true),
            FILLED_SQUARE_INSET,
        ),
        Endpoint::OpenSquare => square_end(
            "open square",
            OPEN_SQUARE_INSET,
            None,
            OPEN_SQUARE_ATTACHMENT_Y,
        ),
        Endpoint::OpenCircle => circle_end("open circle", OPEN_CIRCLE_ATTACHMENT_Y, None),
        Endpoint::InvertedArrow => line_end(
            "inverted arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::Line => line_end(
            "line",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
            ],
            Some((CANONICAL_MID, LINE_ATTACHMENT_Y)),
            None,
        ),
    }
}

/// The endpoint a stored `TSD.LineEndArchive` names (`None` payload: no
/// decoration), by its identifier as the host reads it.
///
/// # Errors
///
/// Decode a line-end payload under the caller's finite wire limits.
///
/// The identifier is the only semantic field needed by the public endpoint
/// value, but the complete payload is still parsed so oversized or malformed
/// nested geometry cannot bypass the owning package's budget.
pub fn decode(payload: Option<&[u8]>, limits: WireLimits) -> Result<Endpoint, DecodeError> {
    let Some(payload) = payload else {
        return Ok(Endpoint::None);
    };
    let view = WireView::parse(payload, limits)?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() != LINE_END_IDENTIFIER {
            continue;
        }
        if field.wire_type() != 2 || identifier.is_some() {
            return Err(UnknownLineEnd.into());
        }
        identifier = Some(std::str::from_utf8(field.payload()).map_err(|_error| UnknownLineEnd)?);
    }
    let identifier = identifier.unwrap_or("none");
    Ok(match identifier {
        "none" => Endpoint::None,
        "simple arrow" => Endpoint::SimpleArrow,
        "filled circle" => Endpoint::FilledCircle,
        "filled diamond" => Endpoint::FilledDiamond,
        "open arrow" => Endpoint::OpenArrow,
        "filled arrow" => Endpoint::FilledArrow,
        "filled square" => Endpoint::FilledSquare,
        "open square" => Endpoint::OpenSquare,
        "open circle" => Endpoint::OpenCircle,
        "inverted arrow" => Endpoint::InvertedArrow,
        "line" => Endpoint::Line,
        _ => return Err(UnknownLineEnd.into()),
    })
}

// ---- The endpoints of a straight-line shape (the host's
// `shapes::line_end` read and `set_shape_line_endpoints`), shared by the
// Pages and Numbers owners over a style workspace ----

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_INFO_SHAPE: u32 = 1;
const SHAPE_PATH_SOURCE: u32 = 3;
const SHAPE_HEAD_LINE_END: u32 = 4;
const SHAPE_TAIL_LINE_END: u32 = 5;
const PATH_SOURCE_HORIZONTAL_FLIP: u32 = 1;

/// Why a line-endpoint read or write failed. No change escapes a failed
/// write's workspace step.
#[derive(Debug)]
#[non_exhaustive]
pub enum EndpointsError {
    /// The shape or style graph failed (the engine's typed failure).
    Graph(crate::graph::Error),
    /// A stored line end is malformed, outside the standard set or over the
    /// caller's wire limits.
    Decode(DecodeError),
    /// The default decorations would clear an exclusive variation that has
    /// no parent style.
    Orphan,
    /// The written decorations did not read back.
    Verification,
}

impl From<crate::graph::Error> for EndpointsError {
    fn from(error: crate::graph::Error) -> Self {
        Self::Graph(error)
    }
}

impl From<DecodeError> for EndpointsError {
    fn from(error: DecodeError) -> Self {
        Self::Decode(error)
    }
}

impl From<litchi_iwa_common::Error> for EndpointsError {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self::Decode(DecodeError::Wire(error))
    }
}

/// The shape's own (legacy per-shape) head and tail decorations and its
/// horizontal flip.
struct Own {
    head: Option<Vec<u8>>,
    tail: Option<Vec<u8>>,
    flipped: bool,
}

fn single<'a>(view: &WireView<'a>, number: u32) -> Result<Option<&'a [u8]>, EndpointsError> {
    let mut found = None;
    for field in view.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 || found.replace(field.payload()).is_some() {
            return Err(
                crate::graph::Error::Invalid("a shape field is repeated or mistyped").into(),
            );
        }
    }
    Ok(found)
}

fn own(
    workspace: &crate::graph::Workspace<'_>,
    drawable: u64,
    limits: WireLimits,
) -> Result<Own, EndpointsError> {
    let (_, payload) = workspace.unique_payload(
        drawable,
        SHAPE_INFO_MESSAGE_TYPE,
        "a shape must have exactly one ShapeInfo payload",
    )?;
    let info = WireView::parse(payload, limits)?;
    let shape = single(&info, SHAPE_INFO_SHAPE)?
        .ok_or(crate::graph::Error::Invalid("a shape info has no shape"))?;
    let shape = WireView::parse(shape, limits)?;
    let mut flipped = false;
    if let Some(source) = single(&shape, SHAPE_PATH_SOURCE)? {
        let source = WireView::parse(source, limits)?;
        for field in source.fields() {
            if field.number() != PATH_SOURCE_HORIZONTAL_FLIP {
                continue;
            }
            if field.wire_type() != 0 {
                return Err(crate::graph::Error::Invalid("a path flip is mistyped").into());
            }
            let (value, length) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                .map_err(|_error| crate::graph::Error::Invalid("a path flip is malformed"))?;
            if length != field.payload().len() {
                return Err(crate::graph::Error::Invalid("a path flip is malformed").into());
            }
            flipped = value != 0;
        }
    }
    let owned = |payload: Option<&[u8]>| -> Result<Option<Vec<u8>>, EndpointsError> {
        let Some(payload) = payload else {
            return Ok(None);
        };
        WireView::parse(payload, limits)?;
        let mut owned = Vec::new();
        owned.try_reserve_exact(payload.len()).map_err(|_error| {
            DecodeError::Wire(litchi_iwa_common::Error::Allocation {
                resource: "line end",
                amount: payload.len(),
            })
        })?;
        owned.extend_from_slice(payload);
        Ok(Some(owned))
    };
    Ok(Own {
        head: owned(single(&shape, SHAPE_HEAD_LINE_END)?)?,
        tail: owned(single(&shape, SHAPE_TAIL_LINE_END)?)?,
        flipped,
    })
}

/// The effective head and tail of a style chain, each inherited on its own.
fn inherited(
    workspace: &crate::graph::Workspace<'_>,
    style: u64,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>), EndpointsError> {
    Ok((
        crate::shape::effective(workspace, style, crate::shape::Slot::HeadLineEnd)?,
        crate::shape::effective(workspace, style, crate::shape::Slot::TailLineEnd)?,
    ))
}

fn resolved(
    workspace: &crate::graph::Workspace<'_>,
    own: &Own,
    drawable: u64,
    limits: WireLimits,
) -> Result<Endpoints, EndpointsError> {
    let (mut head, mut tail) = (own.head.clone(), own.tail.clone());
    if head.is_none() || tail.is_none() {
        let style = crate::shape::target(workspace, drawable)?.style();
        let (inherited_head, inherited_tail) = inherited(workspace, style)?;
        head = head.or(inherited_head);
        tail = tail.or(inherited_tail);
    }
    let mut endpoints = Endpoints {
        start: decode(tail.as_deref(), limits)?,
        end: decode(head.as_deref(), limits)?,
    };
    if own.flipped {
        std::mem::swap(&mut endpoints.start, &mut endpoints.end);
    }
    Ok(endpoints)
}

/// The directed start and end decorations straight-line shape `drawable`
/// shows: its own legacy decorations, else its style chain's, swapped when
/// its path is flipped horizontally (the host's `shape_line_endpoints`).
///
/// # Errors
///
/// [`EndpointsError`] for an invalid shape or style graph or a line end
/// outside the standard set or the limits.
pub fn endpoints(
    workspace: &crate::graph::Workspace<'_>,
    drawable: u64,
    limits: WireLimits,
) -> Result<Endpoints, EndpointsError> {
    resolved(
        workspace,
        &own(workspace, drawable, limits)?,
        drawable,
        limits,
    )
}

/// Store `wanted` as straight-line shape `drawable`'s decorations through
/// the shape-style plan layer (the host's `set_shape_line_endpoints`): an
/// exclusive variation whose parent shows no decorations is cleared for the
/// default, any other write stores both slots. Returns whether anything
/// changed.
///
/// # Errors
///
/// As [`endpoints`]; [`EndpointsError::Orphan`] and
/// [`EndpointsError::Verification`].
pub fn set_endpoints(
    workspace: &mut crate::graph::Workspace<'_>,
    drawable: u64,
    wanted: Endpoints,
    limits: WireLimits,
) -> Result<bool, EndpointsError> {
    use crate::shape::{self as engine, Slot, Styles as _, Value};
    let own = own(workspace, drawable, limits)?;
    if resolved(workspace, &own, drawable, limits)? == wanted {
        return Ok(false);
    }
    let mut stored = wanted;
    if own.flipped {
        std::mem::swap(&mut stored.start, &mut stored.end);
    }
    let target = engine::target(workspace, drawable)?;
    let payload = workspace
        .shape_style(target.style())
        .ok_or(crate::graph::Error::Invalid("a shape style is missing"))?
        .to_vec();
    let disposable = target.exclusive() && engine::direct_slots(&payload)?.is_some();
    let mut clear = false;
    if disposable && wanted == Endpoints::default() {
        let parent = engine::parent_of(&payload)?.ok_or(EndpointsError::Orphan)?;
        let (head, tail) = inherited(workspace, parent)?;
        let inherited = Endpoints {
            start: decode(tail.as_deref(), limits)?,
            end: decode(head.as_deref(), limits)?,
        };
        clear = inherited == Endpoints::default();
    }
    let writes = if clear {
        vec![(Slot::HeadLineEnd, None), (Slot::TailLineEnd, None)]
    } else {
        vec![
            (Slot::HeadLineEnd, Some(Value::message(encode(stored.end)))),
            (
                Slot::TailLineEnd,
                Some(Value::message(encode(stored.start))),
            ),
        ]
    };
    // A shared or non-minimal style takes a derived child (the host's
    // copy-on-write); the plan layer decides exactly as the host does.
    let plan = engine::plan_set(&*workspace, target, &writes)?;
    engine::apply(workspace, drawable, plan)?;
    if resolved(workspace, &own, drawable, limits)? != wanted {
        return Err(EndpointsError::Verification);
    }
    Ok(true)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_common::WireLimits;

    use super::{Endpoint, decode, encode};

    #[test]
    fn every_endpoint_round_trips_through_its_identifier() {
        for endpoint in [
            Endpoint::None,
            Endpoint::SimpleArrow,
            Endpoint::FilledCircle,
            Endpoint::FilledDiamond,
            Endpoint::OpenArrow,
            Endpoint::FilledArrow,
            Endpoint::FilledSquare,
            Endpoint::OpenSquare,
            Endpoint::OpenCircle,
            Endpoint::InvertedArrow,
            Endpoint::Line,
        ] {
            assert_eq!(
                decode(Some(&encode(endpoint)), WireLimits::DEFAULT),
                Ok(endpoint)
            );
        }
        assert_eq!(decode(None, WireLimits::DEFAULT), Ok(Endpoint::None));
    }

    #[test]
    fn bounded_decode_rejects_an_oversized_payload_before_projection() {
        let payload = encode(Endpoint::Line);
        let limits = WireLimits::default()
            .with_input_bytes(payload.len().saturating_sub(1))
            .expect("positive input limit");

        assert!(decode(Some(&payload), limits).is_err());
    }
}
