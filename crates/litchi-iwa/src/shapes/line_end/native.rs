//! Canonical native path data for iWork line endpoint decorations.

use crate::protobuf::{tsd, tsp};

use super::Endpoint;

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

pub(super) fn endpoint_archive(endpoint: Endpoint) -> tsd::LineEndArchive {
    use tsp::path::ElementType;

    let miter = tsd::LineJoin::MiterJoin as i32;
    let round = tsd::LineJoin::RoundJoin as i32;
    match endpoint {
        Endpoint::None => line_end("none", miter, vec![], None, None),
        Endpoint::SimpleArrow => line_end(
            "simple arrow",
            miter,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(ElementType::CloseSubpath, &[]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            Some(true),
        ),
        Endpoint::FilledCircle => circle_end("filled circle", CIRCLE_INSET, Some(true)),
        Endpoint::FilledDiamond => line_end(
            "filled diamond",
            miter,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(ElementType::CloseSubpath, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::OpenArrow => line_end(
            "open arrow",
            round,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(ElementType::MoveTo, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            None,
        ),
        Endpoint::FilledArrow => line_end(
            "filled arrow",
            miter,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(
                    ElementType::LineTo,
                    &[(CANONICAL_MID, FILLED_ARROW_NOTCH_Y)],
                ),
                element(ElementType::CloseSubpath, &[]),
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
            miter,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(ElementType::LineTo, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(ElementType::CloseSubpath, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::Line => line_end(
            "line",
            miter,
            vec![
                element(ElementType::MoveTo, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(ElementType::LineTo, &[(CANONICAL_MAX, CANONICAL_MIN)]),
            ],
            Some((CANONICAL_MID, LINE_ATTACHMENT_Y)),
            None,
        ),
    }
}

fn circle_end(identifier: &str, endpoint_y: f32, is_filled: Option<bool>) -> tsd::LineEndArchive {
    use tsp::path::ElementType;

    let low = CIRCLE_INSET;
    let high = CANONICAL_MAX - CIRCLE_INSET;
    line_end(
        identifier,
        tsd::LineJoin::MiterJoin as i32,
        vec![
            element(ElementType::MoveTo, &[(high, CANONICAL_MID)]),
            element(
                ElementType::CurveTo,
                &[
                    (high, CIRCLE_CONTROL_HIGH),
                    (CIRCLE_CONTROL_HIGH, high),
                    (CANONICAL_MID, high),
                ],
            ),
            element(
                ElementType::CurveTo,
                &[
                    (CIRCLE_CONTROL_LOW, high),
                    (low, CIRCLE_CONTROL_HIGH),
                    (low, CANONICAL_MID),
                ],
            ),
            element(
                ElementType::CurveTo,
                &[
                    (low, CIRCLE_CONTROL_LOW),
                    (CIRCLE_CONTROL_LOW, low),
                    (CANONICAL_MID, low),
                ],
            ),
            element(
                ElementType::CurveTo,
                &[
                    (CIRCLE_CONTROL_HIGH, low),
                    (high, CIRCLE_CONTROL_LOW),
                    (high, CANONICAL_MID),
                ],
            ),
            element(ElementType::CloseSubpath, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        is_filled,
    )
}

fn square_end(
    identifier: &str,
    inset: f32,
    is_filled: Option<bool>,
    endpoint_y: f32,
) -> tsd::LineEndArchive {
    use tsp::path::ElementType;

    let far = CANONICAL_MAX - inset;
    line_end(
        identifier,
        tsd::LineJoin::MiterJoin as i32,
        vec![
            element(ElementType::MoveTo, &[(inset, inset)]),
            element(ElementType::LineTo, &[(far, inset)]),
            element(ElementType::LineTo, &[(far, far)]),
            element(ElementType::LineTo, &[(inset, far)]),
            element(ElementType::CloseSubpath, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        is_filled,
    )
}

fn line_end(
    identifier: &str,
    line_join: i32,
    elements: Vec<tsp::path::Element>,
    end_point: Option<(f32, f32)>,
    is_filled: Option<bool>,
) -> tsd::LineEndArchive {
    tsd::LineEndArchive {
        path: Some(tsp::Path { elements }),
        line_join: Some(line_join),
        end_point: end_point.map(|(x, y)| tsp::Point { x, y }),
        is_filled,
        identifier: Some(identifier.to_owned()),
    }
}

fn element(r#type: tsp::path::ElementType, points: &[(f32, f32)]) -> tsp::path::Element {
    tsp::path::Element {
        r#type: r#type as i32,
        points: points.iter().map(|&(x, y)| tsp::Point { x, y }).collect(),
    }
}
