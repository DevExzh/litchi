//! Strict structural classification for iWork shape path sources.

use crate::protobuf::{tsd, tsp, tswp};
use crate::{Error, Result};

/// Structural path family used by an ordinary iWork shape.
///
/// Rectangle paths are distinguished from arbitrary Bézier paths so
/// source-built presets can round-trip without relying on display names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ShapePathKind {
    Rectangle,
    BezierPath,
    PointPath,
    ScalarPath,
    Callout,
    ConnectionLine,
    EditableBezierPath,
}

pub(crate) fn shape_path_kind(shape: &tswp::ShapeInfoArchive) -> Result<ShapePathKind> {
    let path = shape.super_.pathsource.as_ref().ok_or_else(|| {
        Error::InvalidFormat("ordinary iWork shape has no path source".to_owned())
    })?;
    let families = [
        path.point_path_source.is_some(),
        path.scalar_path_source.is_some(),
        path.bezier_path_source.is_some(),
        path.callout_path_source.is_some(),
        path.connection_line_path_source.is_some(),
        path.editable_bezier_path_source.is_some(),
    ]
    .into_iter()
    .filter(|present| *present)
    .count();
    if families != 1 {
        return Err(Error::InvalidFormat(format!(
            "ordinary iWork shape must have exactly one path family, found {families}"
        )));
    }
    if let Some(bezier) = &path.bezier_path_source {
        return Ok(if is_rectangle_path(bezier) {
            ShapePathKind::Rectangle
        } else {
            ShapePathKind::BezierPath
        });
    }
    if path.point_path_source.is_some() {
        Ok(ShapePathKind::PointPath)
    } else if path.scalar_path_source.is_some() {
        Ok(ShapePathKind::ScalarPath)
    } else if path.callout_path_source.is_some() {
        Ok(ShapePathKind::Callout)
    } else if path.connection_line_path_source.is_some() {
        Ok(ShapePathKind::ConnectionLine)
    } else {
        Ok(ShapePathKind::EditableBezierPath)
    }
}

fn is_rectangle_path(bezier: &tsd::BezierPathSourceArchive) -> bool {
    use tsp::path::ElementType;

    let Some(size) = bezier.natural_size.as_ref() else {
        return false;
    };
    let Some(path) = bezier.path.as_ref() else {
        return false;
    };
    let expected = [
        (ElementType::MoveTo, Some((0.0, 0.0))),
        (ElementType::LineTo, Some((size.width, 0.0))),
        (ElementType::LineTo, Some((size.width, size.height))),
        (ElementType::LineTo, Some((0.0, size.height))),
        (ElementType::CloseSubpath, None),
        (ElementType::MoveTo, Some((0.0, 0.0))),
    ];
    path.elements.len() == expected.len()
        && path
            .elements
            .iter()
            .zip(expected)
            .all(|(element, expected)| {
                ElementType::try_from(element.r#type).ok() == Some(expected.0)
                    && match expected.1 {
                        Some((x, y)) => {
                            element.points.len() == 1
                                && element.points[0].x == x
                                && element.points[0].y == y
                        },
                        None => element.points.is_empty(),
                    }
            })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_missing_and_ambiguous_path_families() {
        let mut shape = tswp::ShapeInfoArchive::default();
        shape.super_.pathsource = Some(tsd::PathSourceArchive::default());
        assert!(shape_path_kind(&shape).is_err());

        let path = shape.super_.pathsource.as_mut().unwrap();
        path.bezier_path_source = Some(tsd::BezierPathSourceArchive::default());
        path.scalar_path_source = Some(tsd::ScalarPathSourceArchive::default());
        assert!(shape_path_kind(&shape).is_err());
    }
}
