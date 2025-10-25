//! Coordinate transformation from WMF logical units to SVG user units
//!
//! Implements the same transformation as libwmf's `svg_translate`, `svg_width`, `svg_height`:
//! - Maps WMF bbox coordinates to SVG canvas (0, 0) to (width, height)
//! - Preserves aspect ratio
//! - Scales dimensions proportionally

/// Coordinate transformer matching libwmf behavior
#[derive(Debug, Clone, Copy)]
pub struct CoordinateTransform {
    bbox_left: f64,
    bbox_top: f64,
    bbox_width: f64,
    bbox_height: f64,
    svg_width: f64,
    svg_height: f64,
}

impl CoordinateTransform {
    pub fn new(bbox: (i16, i16, i16, i16), svg_width: f64, svg_height: f64) -> Self {
        let (left, top, right, bottom) = bbox;
        Self {
            bbox_left: left as f64,
            bbox_top: top as f64,
            bbox_width: (right - left).abs() as f64,
            bbox_height: (bottom - top).abs() as f64,
            svg_width,
            svg_height,
        }
    }

    /// Transform WMF point to SVG coordinates (matches libwmf svg_translate)
    #[inline]
    pub fn point(&self, x: i16, y: i16) -> (f64, f64) {
        let tx = (x as f64 - self.bbox_left) * self.svg_width / self.bbox_width;
        let ty = (y as f64 - self.bbox_top) * self.svg_height / self.bbox_height;
        (tx, ty)
    }

    /// Scale WMF width to SVG units (matches libwmf svg_width)
    #[inline]
    pub fn width(&self, w: f64) -> f64 {
        w * self.svg_width / self.bbox_width
    }

    /// Scale WMF height to SVG units (matches libwmf svg_height)
    #[inline]
    pub fn height(&self, h: f64) -> f64 {
        h * self.svg_height / self.bbox_height
    }
}
