use super::model::{Instruction, PathInfo, PathKind};
use crate::prop::Array;
use crate::{Error, Result};

pub(crate) fn validate(
    shape_path: PathKind,
    vertices: Option<Array<'_>>,
    segment_info: Option<Array<'_>>,
) -> Result<()> {
    if segment_info.is_some() && vertices.is_none() {
        return Err(Error::MalformedGeometry {
            reason: "pSegmentInfo_complex requires pVertices_complex",
        });
    }
    if matches!(shape_path, PathKind::Complex)
        && segment_info.is_none_or(|array| array.element_count() == 0)
    {
        return Err(Error::MalformedGeometry {
            reason: "complex shapePath requires non-empty pSegmentInfo_complex",
        });
    }

    let (Some(vertex_array), Some(segment_array)) = (vertices, segment_info) else {
        return Ok(());
    };

    let mut point_count = 0usize;
    let mut exact = true;
    for info in segment_array.elements().filter_map(path_info) {
        validate_instruction(info)?;
        let Some(consumed) = info.point_count() else {
            exact = false;
            continue;
        };
        point_count = point_count
            .checked_add(consumed)
            .ok_or(Error::MalformedGeometry {
                reason: "path point count overflows usize",
            })?;
    }
    if exact && point_count != usize::from(vertex_array.element_count()) {
        return Err(Error::MalformedGeometry {
            reason: "pSegmentInfo_complex consumes a different number of vertices",
        });
    }
    Ok(())
}

fn path_info(data: &[u8]) -> Option<PathInfo> {
    let bytes: [u8; 2] = data.try_into().ok()?;
    Some(PathInfo::from_raw(u16::from_le_bytes(bytes)))
}

fn validate_instruction(info: PathInfo) -> Result<()> {
    let invalid = match info.instruction() {
        Instruction::MoveTo | Instruction::End => info.segments() != 0,
        Instruction::Close => info.segments() != 1,
        Instruction::LineTo
        | Instruction::CurveTo
        | Instruction::Escape(_)
        | Instruction::ClientEscape(_)
        | Instruction::Unknown(_) => false,
    };
    if invalid {
        return Err(Error::MalformedGeometry {
            reason: "MSOPATHINFO segment count is invalid for its instruction",
        });
    }

    if let Some(escape) = info.instruction().escape() {
        if matches!(
            escape,
            super::model::EscapeKind::AngleEllipseTo | super::model::EscapeKind::AngleEllipse
        ) && !info.segments().is_multiple_of(3)
        {
            return Err(Error::MalformedGeometry {
                reason: "angle ellipse escape does not contain complete POINT triples",
            });
        }
        if matches!(
            escape,
            super::model::EscapeKind::ArcTo
                | super::model::EscapeKind::Arc
                | super::model::EscapeKind::ClockwiseArcTo
                | super::model::EscapeKind::ClockwiseArc
        ) && !info.segments().is_multiple_of(4)
        {
            return Err(Error::MalformedGeometry {
                reason: "arc escape does not contain complete POINT quadruples",
            });
        }
        if matches!(
            escape,
            super::model::EscapeKind::FillColor | super::model::EscapeKind::LineColor
        ) && info.segments() != 1
        {
            return Err(Error::MalformedGeometry {
                reason: "color escape must contain exactly one POINT",
            });
        }
    }
    Ok(())
}
