//! Bounded, deterministic SVG rasterization.

use std::io::{Cursor, Write};
use std::sync::{Arc, OnceLock};

use image::{ExtendedColorType, ImageEncoder, ImageFormat, ImageReader, Rgba};
use litchi_core::error::{Error, Result};
use resvg::{tiny_skia, usvg};

/// Resource ceilings applied before SVG rasterization allocates its pixel buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
pub(crate) fn rasterize_svg(
    svg: &str,
    width: u32,
    height: u32,
    background: Rgba<u8>,
    format: ImageFormat,
    limits: &RasterLimits,
) -> Result<Vec<u8>> {
    validate_request(svg, width, height, format, limits)?;

    let fonts = system_fonts();
    let fallback_faces = Arc::clone(&fonts.fallback);
    let options = usvg::Options {
        image_href_resolver: bounded_image_resolver(*limits),
        fontdb: Arc::clone(&fonts.database),
        font_resolver: usvg::FontResolver {
            select_font: stand_in_font_selector(),
            // Offer each resident face once; shaping skips faces that still
            // lack the character, so the search stays bounded.
            select_fallback: Box::new(move |_, used, _| {
                fallback_faces
                    .iter()
                    .copied()
                    .find(|face| !used.contains(face))
            }),
        },
        ..usvg::Options::default()
    };
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
        let (scale_x, scale_y) = (
            width as f32 / source.width(),
            height as f32 / source.height(),
        );
        if !group_fits(tree.root(), scale_x, scale_y) {
            return Err(Error::Unsupported(
                "picture geometry extends beyond the rasterizer's numeric range".to_string(),
            ));
        }
        let transform = tiny_skia::Transform::from_scale(scale_x, scale_y);
        resvg::render(&tree, transform, &mut pixmap);
    }

    encode_pixels(pixels, width, height, format, limits.max_output_bytes)
}

/// Largest coordinate magnitude, in output pixels, accepted for rendering.
///
/// tiny-skia's edge builder overflows its fixed-point arithmetic, and panics,
/// for geometry about a billion pixels away; real pictures stay far inside
/// this bound, which leaves two orders of magnitude of margin.
const MAX_RASTER_COORDINATE: f32 = 1.0e7;

/// Whether every node, clip path, mask and pattern under `group` stays within
/// [`MAX_RASTER_COORDINATE`] once scaled to output pixels.
fn group_fits(group: &usvg::Group, scale_x: f32, scale_y: f32) -> bool {
    let fits = |rect: usvg::Rect| {
        [
            rect.left() * scale_x,
            rect.right() * scale_x,
            rect.top() * scale_y,
            rect.bottom() * scale_y,
        ]
        .iter()
        .all(|value| value.is_finite() && value.abs() <= MAX_RASTER_COORDINATE)
    };
    let paint_fits = |paint: Option<&usvg::Paint>| match paint {
        Some(usvg::Paint::Pattern(pattern)) => group_fits(pattern.root(), scale_x, scale_y),
        _ => true,
    };
    if group
        .clip_path()
        .is_some_and(|clip| !group_fits(clip.root(), scale_x, scale_y))
        || group
            .mask()
            .is_some_and(|mask| !group_fits(mask.root(), scale_x, scale_y))
    {
        return false;
    }
    group.children().iter().all(|node| {
        fits(node.abs_stroke_bounding_box())
            && match node {
                usvg::Node::Group(child) => group_fits(child, scale_x, scale_y),
                usvg::Node::Path(path) => {
                    paint_fits(path.fill().map(usvg::Fill::paint))
                        && paint_fits(path.stroke().map(usvg::Stroke::paint))
                },
                usvg::Node::Image(_) | usvg::Node::Text(_) => true,
            }
    })
}

/// Installed families that stand in for faces that metafiles name.
///
/// Liberation and the Chrome OS core fonts (Arimo, Tinos, Cousine, Carlito,
/// Caladea) are metric-compatible with the Windows core fonts they replace,
/// so text set with them keeps the widths the picture was laid out with.
const FACE_STAND_INS: &[(&[&str], &[&str])] = &[
    (
        &["Arial", "Helvetica", "Helv", "Arial MT"],
        &["Liberation Sans", "Arimo"],
    ),
    (
        &["Arial Narrow", "Helvetica Narrow"],
        &["Liberation Sans Narrow"],
    ),
    (
        &["Times New Roman", "Times", "Tms Rmn", "Times New Roman PS"],
        &["Liberation Serif", "Tinos"],
    ),
    (&["Courier New", "Courier"], &["Liberation Mono", "Cousine"]),
    (&["Calibri"], &["Carlito"]),
    (&["Cambria"], &["Caladea"]),
];

/// Families that carry the mathematical symbols, arrows, geometric shapes and
/// dingbats that symbol-face text is decoded to; they draw those characters
/// when the selected face lacks them. They do not stand in for the symbol
/// faces themselves: a glyph position without a Unicode equivalent keeps the
/// missing glyph of the selected text face, which the common text faces leave
/// empty, instead of the missing-glyph box of a symbol family.
const SYMBOL_FAMILIES: &[&str] = &[
    "DejaVu Sans",
    "Noto Sans Symbols",
    "Noto Sans Symbols 2",
    "Noto Sans Math",
    "OpenSymbol",
];

/// Installed families to try in place of the face `name`.
fn face_stand_ins(name: &str) -> &'static [&'static str] {
    FACE_STAND_INS
        .iter()
        .find(|(names, _)| names.iter().any(|known| known.eq_ignore_ascii_case(name)))
        .map_or(&[], |(_, stand_ins)| stand_ins)
}

/// Font database query families for an SVG `font-family` list: each family
/// followed by its stand-ins, then the serif default that usvg's own
/// selector ends with.
fn query_families<'a>(families: &'a [usvg::FontFamily]) -> Vec<usvg::fontdb::Family<'a>> {
    use usvg::fontdb::Family;
    let mut query = Vec::with_capacity(families.len() * 3 + 1);
    for family in families {
        match family {
            usvg::FontFamily::Named(name) => {
                query.push(Family::Name(name));
                query.extend(
                    face_stand_ins(name)
                        .iter()
                        .map(|&stand_in| Family::Name(stand_in)),
                );
            },
            usvg::FontFamily::Serif => query.push(Family::Serif),
            usvg::FontFamily::SansSerif => query.push(Family::SansSerif),
            usvg::FontFamily::Cursive => query.push(Family::Cursive),
            usvg::FontFamily::Fantasy => query.push(Family::Fantasy),
            usvg::FontFamily::Monospace => query.push(Family::Monospace),
        }
    }
    query.push(Family::Serif);
    query
}

/// Font selector that replaces a face the system lacks by an installed
/// stand-in before falling back to the generic families; usvg's default
/// selector falls back to the serif default directly, so a picture that
/// names Arial alone would be drawn with a serif face.
fn stand_in_font_selector() -> usvg::FontSelectionFn<'static> {
    Box::new(|font, database| {
        database.query(&usvg::fontdb::Query {
            families: &query_families(font.families()),
            weight: usvg::fontdb::Weight(font.weight()),
            stretch: font.stretch().into(),
            style: font.style().into(),
        })
    })
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
    for pixel in pixels.as_chunks_mut::<4>().0.iter_mut() {
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

/// System fonts shared by every rasterization in the process.
struct SystemFonts {
    database: Arc<usvg::fontdb::Database>,
    /// Memory-resident faces consulted for characters the requested font lacks.
    fallback: Arc<[usvg::fontdb::ID]>,
}

const SANS_SERIF_FAMILIES: &[&str] = &[
    "Arial",
    "Liberation Sans",
    "Arimo",
    "Helvetica",
    "DejaVu Sans",
    "Noto Sans",
    "FreeSans",
];
const SERIF_FAMILIES: &[&str] = &[
    "Times New Roman",
    "Liberation Serif",
    "Tinos",
    "Times",
    "DejaVu Serif",
    "Noto Serif",
    "FreeSerif",
];
const MONOSPACE_FAMILIES: &[&str] = &[
    "Courier New",
    "Liberation Mono",
    "Cousine",
    "Courier",
    "DejaVu Sans Mono",
    "Noto Sans Mono",
    "FreeMono",
];
/// Largest font file kept resident in memory for generic-family fallback.
const MAX_RESIDENT_FONT_BYTES: u64 = 32 * 1024 * 1024;

/// Returns the process-wide system font set.
///
/// Enumerating every installed font costs far more than rendering a typical
/// metafile, so the database is built once. fontdb re-reads a file-backed
/// face on every glyph lookup, so the generic-family faces are loaded into
/// memory, and per-character fallback consults only those faces instead of
/// reading every installed font for each missing glyph.
fn system_fonts() -> &'static SystemFonts {
    static FONTS: OnceLock<SystemFonts> = OnceLock::new();
    FONTS.get_or_init(build_system_fonts)
}

fn build_system_fonts() -> SystemFonts {
    let mut database = usvg::fontdb::Database::new();
    database.load_system_fonts();
    let deterministic = database
        .faces()
        .flat_map(|face| face.families.iter().map(|family| &family.0))
        .min()
        .cloned();
    let sans_serif = preferred_family(&database, SANS_SERIF_FAMILIES).or(deterministic.clone());
    let serif = preferred_family(&database, SERIF_FAMILIES).or(deterministic.clone());
    let monospace = preferred_family(&database, MONOSPACE_FAMILIES).or(deterministic);
    if let Some(family) = &sans_serif {
        database.set_sans_serif_family(family.clone());
        database.set_cursive_family(family.clone());
        database.set_fantasy_family(family.clone());
    }
    if let Some(family) = &serif {
        database.set_serif_family(family.clone());
    }
    if let Some(family) = &monospace {
        database.set_monospace_family(family.clone());
    }
    let families: Vec<String> = [serif, sans_serif, monospace]
        .into_iter()
        .flatten()
        .collect();
    let mut fallback = make_resident(&mut database, &families, false);
    // Symbol-face text decodes to symbols that text faces rarely carry; the
    // regular face of each symbol family is enough to draw them.
    let symbols: Vec<String> = SYMBOL_FAMILIES
        .iter()
        .map(|family| (*family).to_string())
        .collect();
    fallback.extend(make_resident(&mut database, &symbols, true));
    SystemFonts {
        database: Arc::new(database),
        fallback: fallback.into(),
    }
}

fn preferred_family(database: &usvg::fontdb::Database, candidates: &[&str]) -> Option<String> {
    candidates
        .iter()
        .find(|candidate| {
            database.faces().any(|face| {
                face.families
                    .iter()
                    .any(|family| family.0.eq_ignore_ascii_case(candidate))
            })
        })
        .map(|family| (*family).to_string())
}

/// Replaces the file-backed faces of `families` with in-memory copies and
/// returns their identifiers in family order; `regular_only` limits the
/// copies to the upright, normal-weight, normal-width faces.
fn make_resident(
    database: &mut usvg::fontdb::Database,
    families: &[String],
    regular_only: bool,
) -> Vec<usvg::fontdb::ID> {
    let mut paths = Vec::new();
    for family in families {
        for face in database.faces() {
            if !face
                .families
                .iter()
                .any(|name| name.0.eq_ignore_ascii_case(family))
            {
                continue;
            }
            if regular_only
                && (face.weight != usvg::fontdb::Weight::NORMAL
                    || face.style != usvg::fontdb::Style::Normal
                    || face.stretch != usvg::fontdb::Stretch::Normal)
            {
                continue;
            }
            if let usvg::fontdb::Source::File(path) = &face.source
                && !paths.contains(path)
            {
                paths.push(path.clone());
            }
        }
    }
    let mut resident = Vec::new();
    for path in paths {
        let small_enough = std::fs::metadata(&path)
            .map(|metadata| metadata.len() <= MAX_RESIDENT_FONT_BYTES)
            .unwrap_or(false);
        let Some(data) = small_enough.then(|| std::fs::read(&path).ok()).flatten() else {
            continue;
        };
        let stale: Vec<_> = database
            .faces()
            .filter(|face| matches!(&face.source, usvg::fontdb::Source::File(old) if *old == path))
            .map(|face| face.id)
            .collect();
        for id in stale {
            database.remove_face(id);
        }
        resident.extend(database.load_font_source(usvg::fontdb::Source::Binary(Arc::new(data))));
    }
    resident
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

    #[test]
    fn geometry_far_outside_the_canvas_is_rejected_instead_of_panicking() {
        let svg = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 10 10">"#,
            r#"<path d="M0 0L0 -2147483590" stroke="black" stroke-width="846725122"/></svg>"#,
        );
        let result = rasterize_svg(
            svg,
            10,
            10,
            Rgba([255, 255, 255, 255]),
            ImageFormat::Png,
            &RasterLimits::default(),
        );
        assert!(matches!(result, Err(Error::Unsupported(_))), "{result:?}");
    }

    #[test]
    fn system_fonts_are_loaded_once_per_process() {
        assert!(std::ptr::eq(system_fonts(), system_fonts()));
        assert!(Arc::ptr_eq(
            &system_fonts().database,
            &system_fonts().database
        ));
    }

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

    fn render_text(family: &str, text: &str) -> Vec<u8> {
        let svg = format!(
            concat!(
                r#"<svg xmlns="http://www.w3.org/2000/svg" width="96" height="48">"#,
                r#"<text x="4" y="38" font-family="{}" font-size="36" fill="black">{}</text>"#,
                "</svg>",
            ),
            family, text
        );
        rasterize_svg(
            &svg,
            96,
            48,
            Rgba([255; 4]),
            ImageFormat::Png,
            &RasterLimits::default(),
        )
        .unwrap()
    }

    fn installed(family: &str) -> bool {
        preferred_family(&system_fonts().database, &[family]).is_some()
    }

    #[test]
    fn stand_ins_follow_the_face_they_replace() {
        use usvg::fontdb::Family;
        let families = [
            usvg::FontFamily::Named("arial".to_owned()),
            usvg::FontFamily::SansSerif,
            usvg::FontFamily::Named("Unlisted Face".to_owned()),
        ];
        assert_eq!(
            query_families(&families),
            [
                Family::Name("arial"),
                Family::Name("Liberation Sans"),
                Family::Name("Arimo"),
                Family::SansSerif,
                Family::Name("Unlisted Face"),
                Family::Serif,
            ]
        );
        assert_eq!(
            face_stand_ins("Times New Roman"),
            ["Liberation Serif", "Tinos"]
        );
        assert_eq!(
            face_stand_ins("Courier New"),
            ["Liberation Mono", "Cousine"]
        );
        assert!(face_stand_ins("MT Extra").is_empty());
    }

    #[test]
    fn a_missing_windows_face_renders_with_its_metric_compatible_stand_in() {
        if installed("Arial") || !installed("Liberation Sans") {
            return;
        }
        assert_eq!(
            render_text("Arial", "Ag"),
            render_text("Liberation Sans", "Ag")
        );
    }

    #[test]
    fn symbol_characters_draw_from_a_symbol_fallback_face() {
        if !SYMBOL_FAMILIES.iter().any(|family| installed(family)) {
            return;
        }
        // MT Extra's ellipsis and h-bar, which text faces often lack.
        let symbols = render_text("MT Extra, serif", "\u{22ef}\u{210f}");
        let image = decode(&symbols, ImageFormat::Png).to_rgb8();
        assert!(image.pixels().any(|pixel| pixel.0 != [255; 3]));
        let missing = render_text("MT Extra, serif", "\u{10fffd}\u{10fffd}");
        assert_ne!(symbols, missing, "the symbols drew missing-glyph boxes");
    }
}
