//! Equation pictures: PNG renderings of an equation, for the picture an
//! office file keeps beside an equation for readers that cannot lay it out.

use crate::{Error, Format, Limits, Result};

/// CSS pixels per point (96 pixels per inch, 72 points per inch).
const PIXELS_PER_POINT: f64 = 96.0 / 72.0;

/// How [`latex_to_png`] and [`omml_to_png`] draw an equation.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PngOptions {
    font_size_pt: f64,
    scale: f64,
    max_pixels: u64,
}

impl PngOptions {
    /// The largest accepted font size, in points.
    pub const MAX_FONT_SIZE_PT: f64 = 4000.0;
    /// The largest accepted scale.
    pub const MAX_SCALE: f64 = 16.0;

    /// 12 points at 1× scale (96 pixels per inch), at most 16 777 216 pixels.
    pub const DEFAULT: Self = Self {
        font_size_pt: 12.0,
        scale: 1.0,
        max_pixels: 16_777_216,
    };

    /// Options for a font size in points, at 1× scale (96 pixels per inch).
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for a size that is not finite and
    /// positive or exceeds [`PngOptions::MAX_FONT_SIZE_PT`].
    pub fn new(font_size_pt: f64) -> Result<Self> {
        if !(font_size_pt.is_finite()
            && font_size_pt > 0.0
            && font_size_pt <= Self::MAX_FONT_SIZE_PT)
        {
            return Err(Error::InvalidLimit(
                "the font size must be positive and at most PngOptions::MAX_FONT_SIZE_PT points",
            ));
        }
        Ok(Self {
            font_size_pt,
            ..Self::DEFAULT
        })
    }

    /// Draw at `scale` pixels per CSS pixel: 2.0 is 192 pixels per inch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for a scale that is not finite and
    /// positive or exceeds [`PngOptions::MAX_SCALE`].
    pub fn with_scale(self, scale: f64) -> Result<Self> {
        if !(scale.is_finite() && scale > 0.0 && scale <= Self::MAX_SCALE) {
            return Err(Error::InvalidLimit(
                "the scale must be positive and at most PngOptions::MAX_SCALE",
            ));
        }
        Ok(Self { scale, ..self })
    }

    /// Refuse pictures of more than `max_pixels` pixels (width × height).
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for zero.
    pub const fn with_max_pixels(self, max_pixels: u64) -> Result<Self> {
        if max_pixels == 0 {
            return Err(Error::InvalidLimit("the pixel limit must be positive"));
        }
        Ok(Self { max_pixels, ..self })
    }

    /// The font size, in points.
    #[must_use]
    pub const fn font_size_pt(&self) -> f64 {
        self.font_size_pt
    }

    /// Pixels per CSS pixel.
    #[must_use]
    pub const fn scale(&self) -> f64 {
        self.scale
    }

    /// The most pixels a picture may have.
    #[must_use]
    pub const fn max_pixels(&self) -> u64 {
        self.max_pixels
    }

    fn raster(&self) -> mathema::RasterOptions {
        let mut options = mathema::RasterOptions {
            format: mathema::ImageFormat::Png,
            scale: self.scale,
            background: None,
            ..mathema::RasterOptions::default()
        };
        options.render.font_size_px = self.font_size_pt * PIXELS_PER_POINT;
        options.render.fg_color = "#000000".to_owned();
        options
    }
}

impl Default for PngOptions {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A PNG picture of an equation.
#[derive(Clone, PartialEq)]
pub struct Png {
    bytes: Vec<u8>,
    width: u32,
    height: u32,
    scale: f64,
    max_pixels: u64,
}

impl std::fmt::Debug for Png {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Png")
            .field("bytes", &self.bytes.len())
            .field("width", &self.width)
            .field("height", &self.height)
            .field("scale", &self.scale)
            .finish()
    }
}

impl Png {
    /// The PNG file.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The PNG file, by value.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// The width, in pixels.
    #[must_use]
    pub const fn width(&self) -> u32 {
        self.width
    }

    /// The height, in pixels.
    #[must_use]
    pub const fn height(&self) -> u32 {
        self.height
    }

    /// The picture's size in points (`width` and `height`): its pixel size
    /// at the scale it was drawn at, so it matches the requested font size.
    #[must_use]
    pub fn size_pt(&self) -> (f64, f64) {
        let factor = self.scale * PIXELS_PER_POINT;
        (
            f64::from(self.width) / factor,
            f64::from(self.height) / factor,
        )
    }

    /// This picture on a transparent canvas of `width` × `height` pixels,
    /// its top-left corner at (`x`, `y`) and its size multiplied by `zoom`
    /// (at most 1). The result keeps this picture's scale and pixel limit,
    /// so [`Png::size_pt`] gives the canvas size at that scale. A document
    /// uses it to show an equation inside a larger frame, as an office
    /// application's picture of a text box does.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for a zoom outside `(0, 1]` or a
    /// position that is not finite and non-negative, and [`Error::Render`]
    /// for an empty canvas, a canvas beyond the pixel limit, or a picture
    /// that cannot be decoded or encoded.
    pub fn placed(&self, width: u32, height: u32, x: f64, y: f64, zoom: f64) -> Result<Self> {
        if !(zoom.is_finite() && zoom > 0.0 && zoom <= 1.0) {
            return Err(Error::InvalidLimit("the zoom must be in (0, 1]"));
        }
        if !(x.is_finite() && y.is_finite() && x >= 0.0 && y >= 0.0) {
            return Err(Error::InvalidLimit(
                "the position must be finite and non-negative",
            ));
        }
        if width == 0 || height == 0 || u64::from(width) * u64::from(height) > self.max_pixels {
            return Err(Error::Render(format!(
                "a {width}x{height} canvas is empty or exceeds the {}-pixel limit",
                self.max_pixels
            )));
        }
        let picture = tiny_skia::Pixmap::decode_png(&self.bytes)
            .map_err(|error| Error::Render(format!("the picture does not decode: {error}")))?;
        let mut canvas = tiny_skia::Pixmap::new(width, height)
            .ok_or_else(|| Error::Render(format!("cannot allocate a {width}x{height} canvas")))?;
        #[allow(
            clippy::cast_possible_truncation,
            reason = "a canvas is at most a few thousand pixels in each direction"
        )]
        let transform =
            tiny_skia::Transform::from_row(zoom as f32, 0.0, 0.0, zoom as f32, x as f32, y as f32);
        canvas.draw_pixmap(
            0,
            0,
            picture.as_ref(),
            &tiny_skia::PixmapPaint {
                quality: if zoom < 1.0 {
                    tiny_skia::FilterQuality::Bicubic
                } else {
                    tiny_skia::FilterQuality::Nearest
                },
                ..tiny_skia::PixmapPaint::default()
            },
            transform,
            None,
        );
        let bytes = canvas
            .encode_png()
            .map_err(|error| Error::Render(format!("the canvas does not encode: {error}")))?;
        Ok(Self {
            bytes,
            width,
            height,
            scale: self.scale,
            max_pixels: self.max_pixels,
        })
    }
}

/// Render LaTeX math as a PNG picture with a transparent background.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for LaTeX that does not parse, the limit
/// errors when the input exceeds `limits`, and [`Error::Render`] when the
/// picture would exceed [`PngOptions::max_pixels`] or cannot be drawn.
pub fn latex_to_png(latex: &str, options: &PngOptions, limits: &Limits) -> Result<Png> {
    let parsed = mathema::from_latex_with(latex, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Latex))?;
    render(&parsed, options)
}

/// Render an OMML equation (`m:oMath`, `m:oMathPara` or a fragment of
/// them) as a PNG picture with a transparent background.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for malformed XML or OMML, the limit errors
/// when the input exceeds `limits`, and [`Error::Render`] when the picture
/// would exceed [`PngOptions::max_pixels`] or cannot be drawn.
pub fn omml_to_png(omml: &str, options: &PngOptions, limits: &Limits) -> Result<Png> {
    let parsed = mathema::from_omml_with(omml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    render(&parsed, options)
}

/// The size, in points, of the picture [`omml_to_png`] draws for `omml`
/// at `options`' font size, whatever its scale; measured without drawing,
/// so a caller can choose a scale that fits [`PngOptions::max_pixels`].
///
/// # Errors
///
/// As [`omml_to_png`], without the pixel limit.
pub fn omml_size_pt(omml: &str, options: &PngOptions, limits: &Limits) -> Result<(f64, f64)> {
    let parsed = mathema::from_omml_with(omml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    let (width, height) = layout_size(&parsed, &options.raster())?;
    Ok((
        (width + PADDING) / PIXELS_PER_POINT,
        (height + PADDING) / PIXELS_PER_POINT,
    ))
}

/// The rasterizer pads each side by two pixels at 1× scale.
const PADDING: f64 = 4.0;

/// The layout's size in pixels at 1× scale, read from a text-only SVG
/// before any pixel buffer is allocated: the rasterizer allocates the whole
/// picture at once.
fn layout_size(
    parsed: &mathema::ConvertResult,
    raster: &mathema::RasterOptions,
) -> Result<(f64, f64)> {
    let svg = parsed
        .to_svg_text_with(&raster.render)
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    let (width, height) = svg_size(&svg)
        .ok_or_else(|| Error::Render("the equation layout has no size".to_owned()))?;
    // The SVG adds 0.08 em on every side, which the raster does not.
    let margin = 0.16 * raster.render.font_size_px;
    Ok(((width - margin).max(0.0), (height - margin).max(0.0)))
}

pub(crate) fn render(parsed: &mathema::ConvertResult, options: &PngOptions) -> Result<Png> {
    let raster = options.raster();
    let (width, height) = layout_size(parsed, &raster)?;
    let padded = |extent: f64| ((extent + PADDING) * options.scale).ceil();
    let pixels = padded(width) * padded(height);
    if !pixels.is_finite() || pixels > options.max_pixels as f64 {
        return Err(Error::Render(format!(
            "the picture would exceed the {}-pixel limit",
            options.max_pixels
        )));
    }
    let bytes = parsed
        .to_png_with(&raster)
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    let (width, height) = png_size(&bytes)
        .ok_or_else(|| Error::Render("the renderer returned no PNG header".to_owned()))?;
    if u64::from(width) * u64::from(height) > options.max_pixels {
        return Err(Error::Render(format!(
            "the picture exceeds the {}-pixel limit",
            options.max_pixels
        )));
    }
    Ok(Png {
        bytes,
        width,
        height,
        scale: options.scale,
        max_pixels: options.max_pixels,
    })
}

/// The `width` and `height` attributes of an SVG root, in pixels.
fn svg_size(svg: &str) -> Option<(f64, f64)> {
    let header = &svg[..svg.find('>')?];
    let attribute = |name: &str| -> Option<f64> {
        let start = header.find(name)? + name.len();
        let rest = &header[start..];
        rest[..rest.find('"')?].parse().ok()
    };
    Some((attribute(" width=\"")?, attribute(" height=\"")?))
}

/// The pixel size in a PNG's `IHDR` chunk (ISO/IEC 15948, 11.2.2).
fn png_size(png: &[u8]) -> Option<(u32, u32)> {
    if !png.starts_with(b"\x89PNG\r\n\x1a\n") || png.get(12..16)? != b"IHDR" {
        return None;
    }
    let read = |at: usize| -> Option<u32> {
        Some(u32::from_be_bytes(png.get(at..at + 4)?.try_into().ok()?))
    };
    Some((read(16)?, read(20)?))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_svg_and_png_sizes() {
        assert_eq!(
            svg_size(r#"<svg xmlns="x" width="12.5" height="3" viewBox="0 0 1 1">"#),
            Some((12.5, 3.0))
        );
        assert_eq!(svg_size("<svg>"), None);
        assert_eq!(png_size(b"GIF89a"), None);
    }
}
