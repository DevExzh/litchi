//! `TSD.LineEndArchive` payloads of the standard line endpoints (record
//! 0371): the host's `shapes::line_end::native::endpoint_archive`, field by
//! field in tag order, and the identifier-based read back. Lifted from the
//! Keynote leaf's line-endpoint writer (record 0372) so the Pages and
//! Numbers owners share it.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::line::Endpoint;
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
