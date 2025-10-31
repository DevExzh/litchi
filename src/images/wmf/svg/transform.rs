//! Coordinate transformation from WMF logical units to SVG user units
//!
//! Implements the same transformation as libwmf's `svg_translate`, `svg_width`, `svg_height`:
//! - Maps WMF bbox coordinates to SVG canvas (0, 0) to (width, height)
//! - Preserves aspect ratio
//! - Scales dimensions proportionally

use super::simd::SimdTransform;

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

    /// Transform multiple points in batch using SIMD acceleration
    ///
    /// This method processes multiple points efficiently using SIMD instructions
    /// when available. It's significantly faster than calling `point()` in a loop.
    ///
    /// # Arguments
    ///
    /// * `xs` - Input x-coordinates (i16 WMF coordinates)
    /// * `ys` - Input y-coordinates (i16 WMF coordinates)
    /// * `out_x` - Output buffer for transformed x-coordinates
    /// * `out_y` - Output buffer for transformed y-coordinates
    ///
    /// # Example
    ///
    /// ```ignore
    /// let xs = vec![100i16, 200, 300];
    /// let ys = vec![50i16, 100, 150];
    /// let mut out_x = vec![0.0; 3];
    /// let mut out_y = vec![0.0; 3];
    /// transform.transform_points_batch(&xs, &ys, &mut out_x, &mut out_y);
    /// ```
    #[inline]
    pub fn transform_points_batch(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
    ) {
        let simd = SimdTransform::new(
            self.bbox_left,
            self.bbox_top,
            self.svg_width / self.bbox_width,
            self.svg_height / self.bbox_height,
        );
        simd.transform_batch(xs, ys, out_x, out_y);
    }

    /// Transform points in-place and write to string buffer (optimized for SVG generation)
    ///
    /// This is a specialized method that transforms points and immediately formats them
    /// as comma-separated pairs, avoiding intermediate allocations.
    ///
    /// # Arguments
    ///
    /// * `xs` - Input x-coordinates
    /// * `ys` - Input y-coordinates  
    /// * `buffer` - Output string buffer for formatted points
    /// * `separator` - Separator between point pairs (typically ' ' for SVG)
    ///
    /// # Performance
    ///
    /// Uses SIMD for transformation, then formats directly to output string.
    /// Reuses a thread-local buffer to avoid allocations on each call.
    pub fn transform_and_format_points(
        &self,
        xs: &[i16],
        ys: &[i16],
        buffer: &mut String,
        separator: char,
    ) {
        use crate::images::svg_utils::write_num;

        let len = xs.len().min(ys.len());
        if len == 0 {
            return;
        }

        // For small counts, just use scalar to avoid allocation overhead
        if len <= 4 {
            for i in 0..len {
                if i > 0 {
                    buffer.push(separator);
                }
                let (tx, ty) = self.point(xs[i], ys[i]);
                write_num(buffer, tx);
                buffer.push(',');
                write_num(buffer, ty);
            }
            return;
        }

        // Use SIMD for larger batches
        let mut out_x = vec![0.0; len];
        let mut out_y = vec![0.0; len];

        self.transform_points_batch(xs, ys, &mut out_x, &mut out_y);

        for i in 0..len {
            if i > 0 {
                buffer.push(separator);
            }
            write_num(buffer, out_x[i]);
            buffer.push(',');
            write_num(buffer, out_y[i]);
        }
    }
}
