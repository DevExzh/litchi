//! Bulk EMF geometry operations.
//!
//! These routines deliberately use safe scalar code. LLVM auto-vectorizes these
//! tight loops on supported targets, while the implementation remains portable
//! and cannot perform an unaligned or out-of-bounds intrinsic load on hostile
//! metafile input.

use crate::emf::records::types::{PointL, XForm};

/// Transform all points in-place.
#[inline]
pub fn transform_points_simd(points: &mut [PointL], xform: &XForm) {
    for point in points {
        let (x, y) = xform.transform_point(f64::from(point.x), f64::from(point.y));
        point.x = saturating_f64_to_i32(x);
        point.y = saturating_f64_to_i32(y);
    }
}

#[inline]
fn saturating_f64_to_i32(value: f64) -> i32 {
    if value.is_nan() {
        0
    } else if value <= f64::from(i32::MIN) {
        i32::MIN
    } else if value >= f64::from(i32::MAX) {
        i32::MAX
    } else {
        value.round() as i32
    }
}

/// Return `(min_x, min_y, max_x, max_y)` for a non-empty point slice.
#[inline]
pub fn calculate_bounds_simd(points: &[PointL]) -> Option<(i32, i32, i32, i32)> {
    let first = points.first()?;
    let mut bounds = (first.x, first.y, first.x, first.y);
    for point in &points[1..] {
        bounds.0 = bounds.0.min(point.x);
        bounds.1 = bounds.1.min(point.y);
        bounds.2 = bounds.2.max(point.x);
        bounds.3 = bounds.3.max(point.y);
    }
    Some(bounds)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn transforms_and_bounds_points_safely() {
        let mut points = vec![PointL { x: 1, y: 2 }, PointL { x: -3, y: 4 }];
        let xform = XForm {
            m11: 2.0,
            m12: 0.0,
            m21: 0.0,
            m22: -1.0,
            dx: 5.0,
            dy: 1.0,
        };
        transform_points_simd(&mut points, &xform);
        assert_eq!((points[0].x, points[0].y), (7, -1));
        assert_eq!(calculate_bounds_simd(&points), Some((-1, -3, 7, -1)));
    }
}
