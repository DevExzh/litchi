//! Bounded, deterministic SVG rasterization.

use std::io::{Cursor, Write};
use std::sync::Arc;

use image::{ExtendedColorType, ImageEncoder, ImageFormat, ImageReader, Rgba};
use litchi_core::error::{Error, Result};
use resvg::{tiny_skia, usvg};

/// Resource ceilings applied before SVG rasterization allocates its pixel buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "the max prefix makes each independently configurable resource ceiling explicit"
)]
pub(crate) struct RasterLimits {
    /// Maximum target or embedded-raster height.
    pub(crate) max_height: u32,
    /// Maximum UTF-8 SVG input size.
    pub(crate) max_input_bytes: usize,
    /// Maximum encoded PNG or JPEG output size.
    pub(crate) max_output_bytes: usize,
    /// Maximum target or embedded-raster pixel count.
    pub(crate) max_pixels: u64,
    /// Maximum target or embedded-raster width.
    pub(crate) max_width: u32,
}

struct LimitedWriter {
    inner: Vec<u8>,
    maximum: usize,
}

impl Default for RasterLimits {
    fn default() -> Self {
        Self {
            max_height: 8192,
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
            max_pixels: 32 * 1024 * 1024,
            max_width: 8192,
        }
    }
}

impl LimitedWriter {
    fn into_inner(self) -> Vec<u8> {
        self.inner
    }

    const fn new(maximum: usize) -> Self {
        Self {
            inner: Vec::new(),
            maximum,
        }
    }
}

impl Write for LimitedWriter {
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }

    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let end = self
            .inner
            .len()
            .checked_add(buffer.len())
            .ok_or_else(|| std::io::Error::other("encoded output length overflows"))?;
        if end > self.maximum {
            return Err(std::io::Error::other("encoded output limit exceeded"));
        }
        self.inner.try_reserve(buffer.len()).map_err(|error| {
            std::io::Error::other(format!("encoded allocation failed: {error}"))
        })?;
        self.inner.extend_from_slice(buffer);
        Ok(buffer.len())
    }
}

/// Renders an SVG string at an exact target size and encodes it as PNG or JPEG.
///
/// External image paths are deliberately ignored. Embedded raster images are
/// accepted only when their encoded bytes and dimensions satisfy `limits`.
/// JPEG rendering makes `background` opaque before drawing, so transparency is
/// composited instead of being silently discarded by the encoder.
#[allow(
    clippy::cast_precision_loss,
    reason = "tiny-skia transforms use f32 and validated raster dimensions are at most u32"
)]
pub(crate) fn rasterize_svg(
    svg: &str,
    width: u32,
    height: u32,
    background: Rgba<u8>,
    format: ImageFormat,
    limits: &RasterLimits,
) -> Result<Vec<u8>> {
    validate_request(svg, width, height, format, limits)?;

    let mut options = usvg::Options {
        image_href_resolver: bounded_image_resolver(*limits),
        ..usvg::Options::default()
    };
    load_system_fonts(&mut options);
    let tree = usvg::Tree::from_str(svg, &options)
        .map_err(|error| parse(format!("failed to parse SVG: {error}")))?;

    let byte_len = pixel_buffer_len(width, height)?;
    let mut pixels = Vec::new();
    pixels
        .try_reserve_exact(byte_len)
        .map_err(|source| Error::Allocation {
            resource: "SVG raster pixel buffer",
            source,
        })?;
    pixels.resize(byte_len, 0);

    {
        let mut pixmap = tiny_skia::PixmapMut::from_bytes(&mut pixels, width, height)
            .ok_or_else(|| parse("target dimensions cannot be represented by the SVG renderer"))?;
        let fill_alpha = if format == ImageFormat::Jpeg {
            u8::MAX
        } else {
            background.0[3]
        };
        pixmap.fill(tiny_skia::Color::from_rgba8(
            background.0[0],
            background.0[1],
            background.0[2],
            fill_alpha,
        ));

        let source = tree.size();
        let transform = tiny_skia::Transform::from_scale(
            width as f32 / source.width(),
            height as f32 / source.height(),
        );
        resvg::render(&tree, transform, &mut pixmap);
    }

    encode_pixels(pixels, width, height, format, limits.max_output_bytes)
}

fn bounded_image_resolver(limits: RasterLimits) -> usvg::ImageHrefResolver<'static> {
    let default_data = usvg::ImageHrefResolver::default_data_resolver();
    let resolve_data = Box::new(
        move |mime: &str, data: Arc<Vec<u8>>, options: &usvg::Options<'_>| {
            if data.len() > limits.max_input_bytes || !embedded_dimensions_fit(&data, limits) {
                return None;
            }
            let kind = default_data(mime, Arc::clone(&data), options)?;
            if matches!(kind, usvg::ImageKind::SVG(_)) {
                None
            } else {
                Some(kind)
            }
        },
    );
    let resolve_string = Box::new(|_: &str, _: &usvg::Options<'_>| None);
    usvg::ImageHrefResolver {
        resolve_data,
        resolve_string,
    }
}

fn demultiply_rgba(pixels: &mut [u8]) {
    for pixel in pixels.chunks_exact_mut(4) {
        let alpha = pixel[3];
        if alpha == 0 {
            pixel[..3].fill(0);
        } else if alpha != u8::MAX {
            for channel in &mut pixel[..3] {
                let scaled = u32::from(*channel) * 255 + u32::from(alpha) / 2;
                *channel = u8::try_from((scaled / u32::from(alpha)).min(255)).unwrap_or(u8::MAX);
            }
        }
    }
}

fn embedded_dimensions_fit(data: &[u8], limits: RasterLimits) -> bool {
    let Ok(reader) = ImageReader::new(Cursor::new(data)).with_guessed_format() else {
        return false;
    };
    let Ok((width, height)) = reader.into_dimensions() else {
        return false;
    };
    dimensions_fit(width, height, &limits).is_ok()
}

fn encode_pixels(
    mut pixels: Vec<u8>,
    width: u32,
    height: u32,
    format: ImageFormat,
    maximum: usize,
) -> Result<Vec<u8>> {
    let mut output = LimitedWriter::new(maximum);
    let result = match format {
        ImageFormat::Png => {
            demultiply_rgba(&mut pixels);
            image::codecs::png::PngEncoder::new(&mut output).write_image(
                &pixels,
                width,
                height,
                ExtendedColorType::Rgba8,
            )
        },
        ImageFormat::Jpeg => {
            rgba_to_rgb(&mut pixels);
            image::codecs::jpeg::JpegEncoder::new_with_quality(&mut output, 90).write_image(
                &pixels,
                width,
                height,
                ExtendedColorType::Rgb8,
            )
        },
        _ => return Err(parse("SVG raster output must be PNG or JPEG")),
    };
    result.map_err(|error| parse(format!("failed to encode SVG raster: {error}")))?;
    Ok(output.into_inner())
}

fn load_system_fonts(options: &mut usvg::Options<'_>) {
    let font_database = options.fontdb_mut();
    font_database.load_system_fonts();
    let fallback_family = font_database
        .faces()
        .flat_map(|face| face.families.iter().map(|family| &family.0))
        .min()
        .cloned();
    if let Some(family) = fallback_family {
        font_database.set_serif_family(family.clone());
        font_database.set_sans_serif_family(family.clone());
        font_database.set_cursive_family(family.clone());
        font_database.set_fantasy_family(family.clone());
        font_database.set_monospace_family(family);
    }
}

fn dimensions_fit(width: u32, height: u32, limits: &RasterLimits) -> Result<()> {
    if width == 0 || height == 0 {
        return Err(parse("target dimensions must be nonzero"));
    }
    if width > limits.max_width {
        return Err(parse(format!(
            "image width is {width}; limit is {}",
            limits.max_width
        )));
    }
    if height > limits.max_height {
        return Err(parse(format!(
            "image height is {height}; limit is {}",
            limits.max_height
        )));
    }
    let pixels = u64::from(width)
        .checked_mul(u64::from(height))
        .ok_or_else(|| parse("image pixel count overflows"))?;
    if pixels > limits.max_pixels {
        return Err(parse(format!(
            "image has {pixels} pixels; limit is {}",
            limits.max_pixels
        )));
    }
    Ok(())
}

fn parse(error: impl Into<String>) -> Error {
    Error::ParseError(error.into())
}

fn pixel_buffer_len(width: u32, height: u32) -> Result<usize> {
    let pixels = u64::from(width)
        .checked_mul(u64::from(height))
        .ok_or_else(|| parse("pixel buffer dimensions overflow"))?;
    let bytes = pixels
        .checked_mul(4)
        .ok_or_else(|| parse("pixel buffer byte count overflows"))?;
    usize::try_from(bytes).map_err(|_| parse("pixel buffer byte count does not fit usize"))
}

fn rgba_to_rgb(pixels: &mut Vec<u8>) {
    let pixel_count = pixels.len() / 4;
    for index in 0..pixel_count {
        let source = index * 4;
        let target = index * 3;
        pixels[target] = pixels[source];
        pixels[target + 1] = pixels[source + 1];
        pixels[target + 2] = pixels[source + 2];
    }
    pixels.truncate(pixel_count * 3);
}

fn validate_request(
    svg: &str,
    width: u32,
    height: u32,
    format: ImageFormat,
    limits: &RasterLimits,
) -> Result<()> {
    if !matches!(format, ImageFormat::Png | ImageFormat::Jpeg) {
        return Err(parse("SVG raster output must be PNG or JPEG"));
    }
    if limits.max_output_bytes == 0 {
        return Err(parse("encoded output byte limit must be nonzero"));
    }
    if svg.len() > limits.max_input_bytes {
        return Err(parse(format!(
            "SVG input exceeds the {}-byte limit",
            limits.max_input_bytes
        )));
    }
    dimensions_fit(width, height, limits)?;
    let _ = pixel_buffer_len(width, height)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const RECT: &str = concat!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" "#,
        r##"viewBox="0 0 10 10"><rect width="10" height="10" fill="#ff0000"/></svg>"##,
    );

    fn decode(bytes: &[u8], format: ImageFormat) -> image::DynamicImage {
        image::load_from_memory_with_format(bytes, format).unwrap()
    }

    fn render(format: ImageFormat) -> Result<Vec<u8>> {
        rasterize_svg(
            RECT,
            8,
            6,
            Rgba([255, 255, 255, 255]),
            format,
            &RasterLimits::default(),
        )
    }

    #[test]
    fn rejects_invalid_requests_before_parsing() {
        let limits = RasterLimits {
            max_height: 4,
            max_input_bytes: 3,
            max_output_bytes: 0,
            max_pixels: 4,
            max_width: 4,
        };
        assert!(rasterize_svg("not svg", 0, 1, Rgba([0; 4]), ImageFormat::Gif, &limits).is_err());

        let limits = RasterLimits::default();
        assert!(rasterize_svg("not svg", 0, 1, Rgba([0; 4]), ImageFormat::Png, &limits).is_err());
        assert!(rasterize_svg("not svg", 1, 0, Rgba([0; 4]), ImageFormat::Png, &limits).is_err());
    }

    #[test]
    fn rejects_input_dimension_pixel_and_output_limits() {
        let base = RasterLimits::default();
        let input = RasterLimits {
            max_input_bytes: RECT.len() - 1,
            ..base
        };
        assert!(rasterize_svg(RECT, 1, 1, Rgba([0; 4]), ImageFormat::Png, &input).is_err());

        let width = RasterLimits {
            max_width: 7,
            ..base
        };
        assert!(rasterize_svg(RECT, 8, 1, Rgba([0; 4]), ImageFormat::Png, &width).is_err());
        let height = RasterLimits {
            max_height: 7,
            ..base
        };
        assert!(rasterize_svg(RECT, 1, 8, Rgba([0; 4]), ImageFormat::Png, &height).is_err());
        let pixels = RasterLimits {
            max_pixels: 63,
            ..base
        };
        assert!(rasterize_svg(RECT, 8, 8, Rgba([0; 4]), ImageFormat::Png, &pixels).is_err());
        let output = RasterLimits {
            max_output_bytes: 8,
            ..base
        };
        assert!(rasterize_svg(RECT, 8, 8, Rgba([0; 4]), ImageFormat::Png, &output).is_err());
    }

    #[test]
    fn rejects_nested_svg_images() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">"#,
            r#"<image width="1" height="1" href="data:image/svg+xml;base64,"#,
            "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPScxJyBoZWlnaHQ9JzEn",
            "PjxyZWN0IHdpZHRoPScxJyBoZWlnaHQ9JzEnIGZpbGw9J2JsYWNrJy8+PC9zdmc+",
            r#""/></svg>"#,
        );
        let bytes = rasterize_svg(
            svg,
            1,
            1,
            Rgba([255; 4]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap();
        assert_eq!(
            decode(&bytes, ImageFormat::Png)
                .to_rgba8()
                .get_pixel(0, 0)
                .0,
            [255; 4]
        );
    }

    #[test]
    fn renders_png_at_exact_target_size() {
        let bytes = render(ImageFormat::Png).unwrap();
        let image = decode(&bytes, ImageFormat::Png).to_rgba8();
        assert_eq!(image.dimensions(), (8, 6));
        assert_eq!(image.get_pixel(4, 3).0, [255, 0, 0, 255]);
    }

    #[test]
    fn png_preserves_composited_alpha() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">"#,
            r##"<rect width="1" height="1" fill="#ff0000" fill-opacity="0.5"/></svg>"##,
        );
        let bytes = rasterize_svg(
            svg,
            1,
            1,
            Rgba([0, 0, 255, 0]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap();
        let pixel = decode(&bytes, ImageFormat::Png)
            .to_rgba8()
            .get_pixel(0, 0)
            .0;
        assert!(pixel[0] >= 250);
        assert!(pixel[2] <= 5);
        assert!((127..=129).contains(&pixel[3]));
    }

    #[test]
    fn jpeg_flattens_transparency_onto_background() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">"#,
            r##"<rect width="1" height="1" fill="#ff0000" fill-opacity="0.5"/></svg>"##,
        );
        let bytes = rasterize_svg(
            svg,
            16,
            16,
            Rgba([0, 0, 255, 0]),
            ImageFormat::Jpeg,
            &RasterLimits::default(),
        )
        .unwrap();
        let pixel = decode(&bytes, ImageFormat::Jpeg)
            .to_rgb8()
            .get_pixel(8, 8)
            .0;
        assert!((115..=140).contains(&pixel[0]), "red={}", pixel[0]);
        assert!(pixel[1] <= 10, "green={}", pixel[1]);
        assert!((115..=140).contains(&pixel[2]), "blue={}", pixel[2]);
    }

    #[test]
    fn rendering_is_deterministic() {
        assert_eq!(
            render(ImageFormat::Png).unwrap(),
            render(ImageFormat::Png).unwrap()
        );
        assert_eq!(
            render(ImageFormat::Jpeg).unwrap(),
            render(ImageFormat::Jpeg).unwrap()
        );
    }

    #[test]
    fn renders_svg_text() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="64" height="32">"#,
            r#"<text x="2" y="26" font-family="sans-serif" font-size="28" fill="black">I</text>"#,
            "</svg>",
        );
        let bytes = rasterize_svg(
            svg,
            64,
            32,
            Rgba([255; 4]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap();
        let image = decode(&bytes, ImageFormat::Png).to_rgb8();
        assert!(image.pixels().any(|pixel| pixel.0 != [255; 3]));
    }

    #[test]
    fn malformed_svg_is_reported() {
        let error = rasterize_svg(
            "<svg>",
            1,
            1,
            Rgba([0; 4]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap_err();
        assert!(error.to_string().contains("failed to parse SVG"));
    }

    #[test]
    fn ignores_external_image_paths() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">"#,
            r#"<image width="1" height="1" href="/definitely/not/a/portable/resource.png"/>"#,
            "</svg>",
        );
        let bytes = rasterize_svg(
            svg,
            1,
            1,
            Rgba([7, 11, 13, 255]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap();
        assert_eq!(
            decode(&bytes, ImageFormat::Png)
                .to_rgba8()
                .get_pixel(0, 0)
                .0,
            [7, 11, 13, 255]
        );
    }
}
