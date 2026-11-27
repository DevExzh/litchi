//! Coordinate transformation from WMF logical units to SVG user units.

use super::simd::SimdTransform;
use super::state::{DeviceRect, MappingState};

#[derive(Debug, Clone, Copy)]
pub(super) struct CoordinateTransform {
    bbox_left: f64,
    bbox_top: f64,
    bbox_width: f64,
    bbox_height: f64,
    svg_width: f64,
    svg_height: f64,
}

impl CoordinateTransform {
    pub(super) fn new(bbox: (f64, f64, f64, f64), svg_width: f64, svg_height: f64) -> Self {
        let (left, top, right, bottom) = bbox;
        let left = finite_or(left.min(right), 0.0);
        let top = finite_or(top.min(bottom), 0.0);
        let width = finite_positive((right - left).abs(), 1.0);
        let height = finite_positive((bottom - top).abs(), 1.0);
        Self {
            bbox_left: left,
            bbox_top: top,
            bbox_width: width,
            bbox_height: height,
            svg_width: finite_positive(svg_width, 1.0),
            svg_height: finite_positive(svg_height, 1.0),
        }
    }

    #[inline]
    pub(super) fn device_point(&self, x: f64, y: f64) -> (f64, f64) {
        (
            (finite_or(x, self.bbox_left) - self.bbox_left) * self.svg_width / self.bbox_width,
            (finite_or(y, self.bbox_top) - self.bbox_top) * self.svg_height / self.bbox_height,
        )
    }

    #[inline]
    pub(super) fn point(&self, mapping: &MappingState, x: i16, y: i16) -> (f64, f64) {
        let (x, y) = mapping.point(x, y);
        self.device_point(x, y)
    }

    pub(super) fn rect(&self, rect: DeviceRect) -> DeviceRect {
        let (x1, y1) = self.device_point(rect.left, rect.top);
        let (x2, y2) = self.device_point(rect.right, rect.bottom);
        DeviceRect::new(x1, y1, x2, y2)
    }

    #[inline]
    pub(super) fn device_width(&self, width: f64) -> f64 {
        finite_or(width.abs() * self.svg_width / self.bbox_width, 0.0)
    }

    #[inline]
    pub(super) fn device_height(&self, height: f64) -> f64 {
        finite_or(height.abs() * self.svg_height / self.bbox_height, 0.0)
    }

    pub(super) fn logical_vector(&self, mapping: &MappingState, dx: f64, dy: f64) -> (f64, f64) {
        let (dx, dy) = mapping.vector(dx, dy);
        (self.device_width(dx), self.device_height(dy))
    }

    pub(super) fn determinant_sign(&self, mapping: &MappingState) -> f64 {
        let (sx, sy) = mapping.scale();
        (sx * sy).signum()
    }

    pub(super) fn canvas_rect(&self) -> DeviceRect {
        DeviceRect::new(0.0, 0.0, self.svg_width, self.svg_height)
    }

    #[inline]
    pub(super) fn transform_points_batch(
        &self,
        mapping: &MappingState,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
    ) -> usize {
        let (mapping_x, mapping_y) = mapping.scale();
        let base_x = self.svg_width / self.bbox_width;
        let base_y = self.svg_height / self.bbox_height;
        let transform = SimdTransform::new(
            mapping_x * base_x,
            mapping_y * base_y,
            (mapping.viewport_origin.0 - mapping.window_origin.0 * mapping_x - self.bbox_left)
                * base_x,
            (mapping.viewport_origin.1 - mapping.window_origin.1 * mapping_y - self.bbox_top)
                * base_y,
        );
        transform.transform_batch(xs, ys, out_x, out_y)
    }

    pub(super) fn transform_and_format_points(
        &self,
        mapping: &MappingState,
        xs: &[i16],
        ys: &[i16],
        buffer: &mut String,
        separator: char,
    ) {
        use crate::svg_utils::write_num;

        let len = xs.len().min(ys.len());
        let mut out_x = vec![0.0; len];
        let mut out_y = vec![0.0; len];
        let count = self.transform_points_batch(mapping, xs, ys, &mut out_x, &mut out_y);
        for index in 0..count {
            if index != 0 {
                buffer.push(separator);
            }
            write_num(buffer, out_x[index]);
            buffer.push(',');
            write_num(buffer, out_y[index]);
        }
    }
}

#[inline]
fn finite_or(value: f64, fallback: f64) -> f64 {
    if value.is_finite() { value } else { fallback }
}

#[inline]
fn finite_positive(value: f64, fallback: f64) -> f64 {
    if value.is_finite() && value > 0.0 {
        value
    } else {
        fallback
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reversed_and_degenerate_bounds_stay_finite() {
        for bbox in [(10.0, 20.0, -10.0, -20.0), (4.0, 4.0, 4.0, 4.0)] {
            let transform = CoordinateTransform::new(bbox, 100.0, 100.0);
            let point = transform.device_point(4.0, 4.0);
            assert!(point.0.is_finite() && point.1.is_finite());
        }
    }

    #[test]
    fn logical_mapping_is_composed_with_canvas_mapping() {
        let transform = CoordinateTransform::new((0.0, 0.0, 200.0, 100.0), 400.0, 200.0);
        let mapping = MappingState {
            mode: super::super::super::constants::map_mode::MM_ANISOTROPIC,
            window_extent: (100.0, 100.0),
            viewport_extent: (200.0, 100.0),
            ..MappingState::default()
        };
        assert_eq!(transform.point(&mapping, 50, 50), (200.0, 100.0));
    }
}
