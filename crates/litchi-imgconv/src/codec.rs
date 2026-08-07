use std::{
    borrow::Cow,
    io::{Cursor, Read, Seek, SeekFrom, Write},
};

use image::{DynamicImage, ImageFormat, ImageReader};
use litchi_core::error::{Error, Result};
use litchi_odraw::image::{Blip, Compression, Meta};

use crate::raster::{RasterLimits, rasterize_svg};

/// The raw metafile format supplied to [`convert_metafile`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InputFormat {
    /// Enhanced Metafile.
    Emf,
    /// Windows Metafile.
    Wmf,
}

/// A requested representation for a raw metafile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutputFormat {
    /// Preserve vector content as SVG, except for a metafile made exclusively
    /// of bitmap-painting records, which is emitted as PNG.
    Auto,
    /// SVG vector output.
    Svg,
    /// PNG raster output.
    Png,
    /// JPEG raster output.
    Jpeg,
}

/// The representation actually produced by a conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConvertedFormat {
    /// SVG vector output.
    Svg,
    /// PNG raster output.
    Png,
    /// JPEG raster output.
    Jpeg,
}

impl ConvertedFormat {
    /// Returns the conventional filename extension.
    #[must_use]
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Svg => "svg",
            Self::Png => "png",
            Self::Jpeg => "jpg",
        }
    }

    /// Returns the IANA media type.
    #[must_use]
    pub const fn mime_type(self) -> &'static str {
        match self {
            Self::Svg => "image/svg+xml",
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
        }
    }
}

/// A non-fatal fact about a completed conversion.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConversionDiagnostic {
    /// A stable, machine-readable diagnostic code.
    pub code: &'static str,
    /// A concise human-readable explanation.
    pub message: String,
}

/// Metadata describing how an image representation was selected.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConversionReport {
    /// Input metafile family.
    pub input: InputFormat,
    /// Requested output representation.
    pub requested: OutputFormat,
    /// Selected output representation.
    pub selected: ConvertedFormat,
    /// Informational diagnostics generated while selecting the output.
    pub diagnostics: Vec<ConversionDiagnostic>,
}

/// Bytes and metadata produced by [`convert_metafile`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConvertedImage {
    /// Encoded image bytes.
    pub bytes: Vec<u8>,
    /// Actual output representation.
    pub format: ConvertedFormat,
    /// IANA media type for [`Self::bytes`].
    pub mime_type: &'static str,
    /// Conventional output filename extension.
    pub extension: &'static str,
    /// Selection and conversion metadata.
    pub report: ConversionReport,
}

/// Explicit resource ceilings for Office image decoding and rendering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum stored BLIP file-data bytes.
    pub max_encoded_bytes: usize,
    /// Maximum decompressed metafile bytes.
    pub max_uncompressed_bytes: usize,
    /// Maximum decoded or rendered width.
    pub max_width: u32,
    /// Maximum decoded or rendered height.
    pub max_height: u32,
    /// Maximum decoded or rendered pixel count.
    pub max_pixels: u64,
    /// Maximum encoded output bytes.
    pub max_output_bytes: usize,
    /// Maximum metafile records accepted for one conversion.
    pub max_records: usize,
    /// Maximum GDI objects a playback implementation may retain.
    pub max_objects: usize,
    /// Maximum nested saved graphics states a playback implementation may retain.
    pub max_state_depth: usize,
    /// Maximum points accepted in one vector path or polygon.
    pub max_path_points: usize,
    /// Maximum SVG elements produced by metafile playback.
    pub max_svg_elements: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_encoded_bytes: 256 * 1024 * 1024,
            max_uncompressed_bytes: 256 * 1024 * 1024,
            max_width: 8192,
            max_height: 8192,
            max_pixels: 32 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
            max_records: 100_000,
            max_objects: 16_384,
            max_state_depth: 256,
            max_path_points: 1_000_000,
            max_svg_elements: 100_000,
        }
    }
}

/// Output sizing and resource limits for a conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Options {
    /// Optional target width.
    pub width: Option<u32>,
    /// Optional target height.
    pub height: Option<u32>,
    /// Decode and output ceilings.
    pub limits: Limits,
}

impl Options {
    /// Sets the target width while retaining the source aspect ratio when
    /// height is absent.
    pub const fn width(mut self, width: u32) -> Self {
        self.width = Some(width);
        self
    }

    /// Sets the target height while retaining the source aspect ratio when
    /// width is absent.
    pub const fn height(mut self, height: u32) -> Self {
        self.height = Some(height);
        self
    }

    /// Replaces all resource ceilings.
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }
}

/// Converts raw EMF or WMF bytes into a bounded, typed image representation.
///
/// `Auto` keeps vector content as SVG. It chooses PNG only for a metafile
/// whose parsed records are exclusively bitmap-painting operations; this
/// avoids flattening ordinary vector drawings merely because they contain an
/// embedded raster image.
pub fn convert_metafile(
    data: &[u8],
    input: InputFormat,
    output: OutputFormat,
    options: Options,
) -> Result<ConvertedImage> {
    validate_limits(&options.limits)?;
    check_encoded(data.len(), &options.limits)?;
    let (svg, source_width, source_height, bitmap_only, mut diagnostics) =
        raw_metafile_svg(data, input, &options.limits)?;
    let selected = match output {
        OutputFormat::Auto if bitmap_only => ConvertedFormat::Png,
        OutputFormat::Auto | OutputFormat::Svg => ConvertedFormat::Svg,
        OutputFormat::Png => ConvertedFormat::Png,
        OutputFormat::Jpeg => ConvertedFormat::Jpeg,
    };
    let bytes = match selected {
        ConvertedFormat::Svg => {
            if svg.len() > options.limits.max_output_bytes {
                return Err(limit("SVG output", options.limits.max_output_bytes));
            }
            svg.into_bytes()
        },
        ConvertedFormat::Png | ConvertedFormat::Jpeg => {
            let (width, height) = target_dimensions(
                source_width,
                source_height,
                options.width,
                options.height,
                &options.limits,
            )?;
            rasterize_svg(
                &svg,
                width,
                height,
                image::Rgba([255, 255, 255, 255]),
                converted_image_format(selected),
                &raster_limits(&options.limits),
            )?
        },
    };
    if output == OutputFormat::Auto {
        let message = if bitmap_only {
            "auto selected PNG because the metafile contains only bitmap-painting records"
        } else {
            "auto selected SVG to preserve vector content"
        };
        diagnostics.push(ConversionDiagnostic {
            code: if bitmap_only {
                "auto-raster-only-metafile"
            } else {
                "auto-vector-first"
            },
            message: message.to_string(),
        });
    }
    Ok(ConvertedImage {
        bytes,
        format: selected,
        mime_type: selected.mime_type(),
        extension: selected.extension(),
        report: ConversionReport {
            input,
            requested: output,
            selected,
            diagnostics,
        },
    })
}

pub(crate) fn rasterize_raw_metafile(
    data: &[u8],
    input: InputFormat,
    format: ImageFormat,
    options: Options,
) -> Result<Vec<u8>> {
    let output = match format {
        ImageFormat::Png => OutputFormat::Png,
        ImageFormat::Jpeg => OutputFormat::Jpeg,
        ImageFormat::WebP => return rasterize_raw_metafile_webp(data, input, options),
        _ => {
            return Err(Error::Unsupported(format!(
                "unsupported metafile raster output: {format:?}"
            )));
        },
    };
    let converted = convert_metafile(data, input, output, options)?;
    reject_lossy_diagnostics(&converted.report)?;
    Ok(converted.bytes)
}

pub(crate) fn reject_lossy_diagnostics(report: &ConversionReport) -> Result<()> {
    let lossy = report.diagnostics.iter().find(|diagnostic| {
        !matches!(
            diagnostic.code,
            "noncanonical-emf-eof-size-last" | "comment"
        )
    });
    if let Some(diagnostic) = lossy {
        return Err(Error::Unsupported(format!(
            "conversion requires reported approximation ({}): {}",
            diagnostic.code, diagnostic.message
        )));
    }
    Ok(())
}

fn rasterize_raw_metafile_webp(
    data: &[u8],
    input: InputFormat,
    options: Options,
) -> Result<Vec<u8>> {
    let converted = convert_metafile(data, input, OutputFormat::Png, options)?;
    reject_lossy_diagnostics(&converted.report)?;
    let png = converted.bytes;
    let image = image::load_from_memory_with_format(&png, ImageFormat::Png)
        .map_err(|error| parse(format!("failed to decode intermediate PNG: {error}")))?;
    encode(&image, ImageFormat::WebP, options.limits.max_output_bytes)
}

fn raw_metafile_svg(
    data: &[u8],
    input: InputFormat,
    limits: &Limits,
) -> Result<(String, u32, u32, bool, Vec<ConversionDiagnostic>)> {
    match input {
        InputFormat::Emf => {
            preflight_emf_header(data, limits)?;
            let parser = crate::emf::EmfParser::new_compatible(data)?;
            check_record_count(parser.records.len(), limits)?;
            check_emf_playback_limits(&parser, limits)?;
            let width = positive(parser.width());
            let height = positive(parser.height());
            let bitmap_only = parser
                .records
                .iter()
                .any(|record| is_emf_bitmap_paint(record.record_type))
                && parser.records.iter().all(|record| {
                    is_emf_bitmap_paint(record.record_type)
                        || is_emf_non_drawing(record.record_type)
                });
            let (svg, emfplus_diagnostics) =
                crate::emf::EmfSvgConverter::with_limits(&parser, *limits)
                    .convert_with_diagnostics()?;
            let mut diagnostics: Vec<_> = parser
                .warnings()
                .iter()
                .map(|warning| match *warning {
                    crate::emf::parser::EmfParserWarning::EofSizeLastMismatch {
                        expected,
                        found,
                    } => ConversionDiagnostic {
                        code: "noncanonical-emf-eof-size-last",
                        message: format!(
                            "accepted legacy EMF EOF SizeLast {found}; the specification requires {expected}"
                        ),
                    },
                })
                .collect();
            diagnostics.extend(emfplus_diagnostics.into_iter().map(|diagnostic| {
                ConversionDiagnostic {
                    code: diagnostic.code,
                    message: match diagnostic.record_offset {
                        Some(offset) => format!("{} (EMF+ offset {offset})", diagnostic.message),
                        None => diagnostic.message,
                    },
                }
            }));
            Ok((svg, width, height, bitmap_only, diagnostics))
        },
        InputFormat::Wmf => {
            preflight_wmf_records(data, limits)?;
            let parser = crate::wmf::WmfParser::new(data)?;
            check_record_count(parser.records.len(), limits)?;
            check_wmf_playback_limits(&parser, limits)?;
            let width = positive(parser.width());
            let height = positive(parser.height());
            let bitmap_only = parser
                .records
                .iter()
                .any(|record| is_wmf_bitmap_paint(record.function))
                && parser.records.iter().all(|record| {
                    is_wmf_bitmap_paint(record.function) || is_wmf_non_drawing(record.function)
                });
            let (svg, warnings) =
                crate::wmf::WmfSvgConverter::with_limits(parser, *limits).to_svg_reported()?;
            let diagnostics = warnings
                .into_iter()
                .map(|message| ConversionDiagnostic {
                    code: "wmf-playback-warning",
                    message,
                })
                .collect();
            Ok((svg, width, height, bitmap_only, diagnostics))
        },
    }
}

fn validate_limits(limits: &Limits) -> Result<()> {
    if limits.max_encoded_bytes == 0
        || limits.max_uncompressed_bytes == 0
        || limits.max_width == 0
        || limits.max_height == 0
        || limits.max_pixels == 0
        || limits.max_output_bytes == 0
        || limits.max_records == 0
        || limits.max_objects == 0
        || limits.max_state_depth == 0
        || limits.max_path_points == 0
        || limits.max_svg_elements == 0
    {
        return Err(parse("all conversion limits must be greater than zero"));
    }
    Ok(())
}

fn preflight_emf_header(data: &[u8], limits: &Limits) -> Result<()> {
    let records = read_u32_at(data, 52)
        .ok_or_else(|| parse("EMF header is too short for its record count"))?;
    let records = usize::try_from(records).map_err(|_| parse("EMF record count overflow"))?;
    if records > limits.max_records {
        return Err(parse(format!(
            "EMF declares {records} records; limit is {}",
            limits.max_records
        )));
    }
    let handles = data
        .get(56..58)
        .map(|bytes| usize::from(u16::from_le_bytes([bytes[0], bytes[1]])))
        .ok_or_else(|| parse("EMF header is too short for its handle count"))?;
    if handles > limits.max_objects {
        return Err(parse(format!(
            "EMF declares {handles} object handles; limit is {}",
            limits.max_objects
        )));
    }
    Ok(())
}

fn preflight_wmf_records(data: &[u8], limits: &Limits) -> Result<()> {
    const PLACEABLE_KEY: u32 = 0x9ac6_cdd7;
    let header_offset = if read_u32_at(data, 0) == Some(PLACEABLE_KEY) {
        22usize
    } else {
        0usize
    };
    let object_offset = header_offset
        .checked_add(10)
        .ok_or_else(|| parse("WMF header offset overflow"))?;
    let objects = usize::from(
        read_u16_at(data, object_offset)
            .ok_or_else(|| parse("WMF header is too short for its object count"))?,
    );
    if objects > limits.max_objects {
        return Err(parse(format!(
            "WMF declares {objects} object slots; limit is {}",
            limits.max_objects
        )));
    }
    let mut offset = header_offset
        .checked_add(18)
        .ok_or_else(|| parse("WMF record offset overflow"))?;
    let mut records = 0usize;
    loop {
        let words = read_u32_at(data, offset)
            .ok_or_else(|| parse("WMF record header is truncated during limit preflight"))?;
        let function = read_u16_at(
            data,
            offset
                .checked_add(4)
                .ok_or_else(|| parse("WMF function offset overflow"))?,
        )
        .ok_or_else(|| parse("WMF record function is truncated during limit preflight"))?;
        let bytes = usize::try_from(words)
            .map_err(|_| parse("WMF record size overflow"))?
            .checked_mul(2)
            .ok_or_else(|| parse("WMF record byte size overflow"))?;
        if bytes < 6 {
            return Err(parse("WMF record is smaller than its header"));
        }
        offset = offset
            .checked_add(bytes)
            .ok_or_else(|| parse("WMF record range overflow"))?;
        if offset > data.len() {
            return Err(parse(
                "WMF record extends beyond input during limit preflight",
            ));
        }
        records = records
            .checked_add(1)
            .ok_or_else(|| parse("WMF record count overflow"))?;
        if records > limits.max_records {
            return Err(parse(format!(
                "WMF has more than {} records",
                limits.max_records
            )));
        }
        if function == crate::wmf::record::EOF {
            break;
        }
    }
    Ok(())
}

fn check_emf_playback_limits(parser: &crate::emf::EmfParser, limits: &Limits) -> Result<()> {
    if usize::from(parser.header.num_handles) > limits.max_objects {
        return Err(parse(format!(
            "EMF object table exceeds limit {}",
            limits.max_objects
        )));
    }
    let mut saved_depth = 0usize;
    let mut open_path_points = None::<usize>;
    for record in &parser.records {
        match record.record_type {
            33 => {
                saved_depth = saved_depth
                    .checked_add(1)
                    .ok_or_else(|| parse("EMF saved DC depth overflow"))?;
                if saved_depth > limits.max_state_depth {
                    return Err(parse(format!(
                        "EMF saved DC depth exceeds limit {}",
                        limits.max_state_depth
                    )));
                }
            },
            34 => saved_depth = 0,
            59 => open_path_points = Some(0),
            60 | 68 => open_path_points = None,
            _ => {},
        }
        if let Some(points) = emf_record_point_count(record) {
            if points > limits.max_path_points {
                return Err(parse(format!(
                    "EMF record has {points} points; limit is {}",
                    limits.max_path_points
                )));
            }
            if let Some(total) = open_path_points.as_mut() {
                *total = total
                    .checked_add(points)
                    .ok_or_else(|| parse("EMF path point count overflow"))?;
                if *total > limits.max_path_points {
                    return Err(parse(format!(
                        "EMF path has more than {} points",
                        limits.max_path_points
                    )));
                }
            }
        }
    }
    Ok(())
}

fn emf_record_point_count(record: &crate::emf::parser::EmfRecord) -> Option<usize> {
    let offset = match record.record_type {
        2..=6 | 56 | 85..=89 | 92 => 16,
        7 | 8 | 90 | 91 => 20,
        _ => return None,
    };
    read_u32_at(&record.data, offset).and_then(|value| usize::try_from(value).ok())
}

fn check_wmf_playback_limits(parser: &crate::wmf::WmfParser, limits: &Limits) -> Result<()> {
    if usize::from(parser.header.num_objects) > limits.max_objects {
        return Err(parse(format!(
            "WMF object table exceeds limit {}",
            limits.max_objects
        )));
    }
    let mut saved_depth = 0usize;
    for record in &parser.records {
        match crate::wmf::record::canonical(record.function) {
            crate::wmf::record::SAVE_DC => {
                saved_depth = saved_depth
                    .checked_add(1)
                    .ok_or_else(|| parse("WMF saved DC depth overflow"))?;
                if saved_depth > limits.max_state_depth {
                    return Err(parse(format!(
                        "WMF saved DC depth exceeds limit {}",
                        limits.max_state_depth
                    )));
                }
            },
            crate::wmf::record::RESTORE_DC => saved_depth = 0,
            crate::wmf::record::POLYGON | crate::wmf::record::POLYLINE => {
                check_wmf_points(read_u16_at(&record.params, 0), limits)?;
            },
            crate::wmf::record::POLYPOLYGON => {
                let polygons = usize::from(read_u16_at(&record.params, 0).unwrap_or(0));
                let mut total = 0usize;
                for index in 0..polygons {
                    let offset = 2usize
                        .checked_add(
                            index
                                .checked_mul(2)
                                .ok_or_else(|| parse("WMF count overflow"))?,
                        )
                        .ok_or_else(|| parse("WMF count offset overflow"))?;
                    total = total
                        .checked_add(usize::from(
                            read_u16_at(&record.params, offset).unwrap_or(0),
                        ))
                        .ok_or_else(|| parse("WMF point count overflow"))?;
                }
                if total > limits.max_path_points {
                    return Err(parse(format!(
                        "WMF poly-polygon has {total} points; limit is {}",
                        limits.max_path_points
                    )));
                }
            },
            _ => {},
        }
    }
    Ok(())
}

fn check_wmf_points(value: Option<u16>, limits: &Limits) -> Result<()> {
    let points = usize::from(value.unwrap_or(0));
    if points > limits.max_path_points {
        return Err(parse(format!(
            "WMF record has {points} points; limit is {}",
            limits.max_path_points
        )));
    }
    Ok(())
}

fn read_u16_at(data: &[u8], offset: usize) -> Option<u16> {
    let bytes = data.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_at(data: &[u8], offset: usize) -> Option<u32> {
    let bytes = data.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn is_emf_bitmap_paint(record_type: u32) -> bool {
    matches!(record_type, 76..=81 | 114 | 116)
}

fn is_emf_non_drawing(record_type: u32) -> bool {
    matches!(
        record_type,
        1 | 9..=14 | 16..=40 | 48..=52 | 57..=61 | 67..=68 | 82 | 93..=95 | 98..=104
            | 107 | 109..=113 | 115 | 119..=122
    )
}

fn is_wmf_bitmap_paint(function: u16) -> bool {
    matches!(
        crate::wmf::record::canonical(function),
        0x0922 | 0x0940 | 0x0B23 | 0x0B41 | 0x0D33 | 0x0F43
    )
}

fn is_wmf_non_drawing(function: u16) -> bool {
    matches!(
        crate::wmf::record::canonical(function),
        0x0000 | 0x001E | 0x0035 | 0x0037 | 0x00F7 | 0x0102..=0x0108 | 0x0127 | 0x012C
            | 0x012D | 0x012E | 0x0139 | 0x0142 | 0x0149 | 0x01F0 | 0x0201 | 0x0209 | 0x020A
            | 0x020B..=0x020F | 0x0211 | 0x0214 | 0x0220 | 0x0231 | 0x0234 | 0x02FA..=0x02FC
            | 0x0410 | 0x0412 | 0x0415 | 0x0416 | 0x0436
    )
}

fn converted_image_format(format: ConvertedFormat) -> ImageFormat {
    match format {
        ConvertedFormat::Svg => unreachable!("SVG is not rasterized"),
        ConvertedFormat::Png => ImageFormat::Png,
        ConvertedFormat::Jpeg => ImageFormat::Jpeg,
    }
}

fn raster_limits(limits: &Limits) -> RasterLimits {
    RasterLimits {
        max_height: limits.max_height,
        max_input_bytes: limits.max_uncompressed_bytes,
        max_output_bytes: limits.max_output_bytes,
        max_pixels: limits.max_pixels,
        max_width: limits.max_width,
    }
}

/// Returns uncompressed BLIP file data without rendering it.
///
/// Bitmap data remains borrowed. Uncompressed metafiles remain borrowed;
/// RFC1950-compressed metafiles allocate exactly one bounded output buffer.
pub fn decode_data<'data>(blip: &Blip<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    check_encoded(blip.data().len(), limits)?;
    match blip {
        Blip::Emf(meta) | Blip::Wmf(meta) | Blip::Pict(meta) => inflate(meta, limits),
        Blip::Jpeg(bitmap) | Blip::Png(bitmap) | Blip::Dib(bitmap) | Blip::Tiff(bitmap) => {
            Ok(Cow::Borrowed(bitmap.data()))
        },
        Blip::Opaque(_) => Err(parse("cannot decode an unknown OfficeArt BLIP kind")),
    }
}

/// Decodes or renders a BLIP and encodes it in a requested raster format.
pub fn convert(blip: &Blip<'_>, format: ImageFormat, options: Options) -> Result<Vec<u8>> {
    validate_limits(&options.limits)?;
    match blip {
        Blip::Emf(meta) => return render_emf(meta, format, options),
        Blip::Wmf(meta) => return render_wmf(meta, format, options),
        Blip::Pict(meta) => {
            let image = render_pict(meta, options)?;
            return encode(&image, format, options.limits.max_output_bytes);
        },
        Blip::Jpeg(_) | Blip::Png(_) | Blip::Dib(_) | Blip::Tiff(_) => {
            let image = decode_bitmap(blip, options)?;
            return encode(&image, format, options.limits.max_output_bytes);
        },
        Blip::Opaque(_) => return Err(parse("cannot render an unknown OfficeArt BLIP kind")),
    }
}

/// Converts a BLIP to PNG under explicit options.
pub fn to_png(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::Png, options)
}

/// Converts a BLIP to JPEG under explicit options.
pub fn to_jpeg(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::Jpeg, options)
}

/// Converts a BLIP to WebP under explicit options.
pub fn to_webp(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::WebP, options)
}

/// Converts an EMF or WMF BLIP to SVG under explicit input/output limits.
pub fn to_svg(blip: &Blip<'_>, options: Options) -> Result<String> {
    validate_limits(&options.limits)?;
    let data = decode_data(blip, &options.limits)?;
    let svg = match blip {
        Blip::Emf(_) => crate::emf::convert_emf_to_svg_with_options(&data, options)?,
        Blip::Wmf(meta) => {
            let data = wmf_with_header(meta, data, &options.limits)?;
            crate::wmf::convert_wmf_to_svg_with_options(&data, options)?
        },
        Blip::Pict(_) => return Err(parse("PICT to SVG rendering is not implemented")),
        _ => return Err(parse("SVG conversion requires an EMF or WMF BLIP")),
    };
    if svg.len() > options.limits.max_output_bytes {
        return Err(limit("SVG output", options.limits.max_output_bytes));
    }
    Ok(svg)
}

fn check_encoded(actual: usize, limits: &Limits) -> Result<()> {
    if actual > limits.max_encoded_bytes {
        return Err(limit("encoded BLIP", limits.max_encoded_bytes));
    }
    Ok(())
}

fn check_record_count(actual: usize, limits: &Limits) -> Result<()> {
    if actual > limits.max_records {
        return Err(parse(format!(
            "metafile has {actual} records; limit is {}",
            limits.max_records
        )));
    }
    if actual > limits.max_svg_elements {
        return Err(parse(format!(
            "metafile can produce more than {} SVG elements",
            limits.max_svg_elements
        )));
    }
    Ok(())
}

fn inflate<'data>(meta: &Meta<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    let header = meta.header();
    let expected = usize::try_from(header.size)
        .map_err(|_| parse("metafile cbSize does not fit this platform"))?;
    if expected > limits.max_uncompressed_bytes {
        return Err(limit(
            "decompressed metafile",
            limits.max_uncompressed_bytes,
        ));
    }
    if header.compression == Compression::None {
        if meta.data().len() != expected {
            return Err(parse("uncompressed metafile length does not match cbSize"));
        }
        return Ok(Cow::Borrowed(meta.data()));
    }

    validate_zlib_header(meta.data())?;
    let mut decoder = flate2::read::ZlibDecoder::new(meta.data());
    let maximum = limits
        .max_uncompressed_bytes
        .checked_add(1)
        .ok_or_else(|| parse("decompression limit cannot be represented"))?;
    let mut output = Vec::with_capacity(expected.min(limits.max_uncompressed_bytes));
    {
        let maximum =
            u64::try_from(maximum).map_err(|_| parse("decompression limit does not fit u64"))?;
        let mut bounded = (&mut decoder).take(maximum);
        bounded
            .read_to_end(&mut output)
            .map_err(|error| parse(format!("RFC1950 decompression failed: {error}")))?;
    }
    if output.len() > limits.max_uncompressed_bytes {
        return Err(limit(
            "decompressed metafile",
            limits.max_uncompressed_bytes,
        ));
    }
    if output.len() != expected {
        return Err(parse(format!(
            "metafile cbSize is {expected}, but decompression produced {} bytes",
            output.len()
        )));
    }
    let encoded_len = u64::try_from(meta.data().len())
        .map_err(|_| parse("compressed metafile length does not fit u64"))?;
    if decoder.total_in() != encoded_len {
        return Err(parse(
            "compressed metafile has trailing bytes after RFC1950 data",
        ));
    }
    Ok(Cow::Owned(output))
}

fn validate_zlib_header(data: &[u8]) -> Result<()> {
    let header = data
        .get(..2)
        .ok_or_else(|| parse("compressed metafile has no RFC1950 header"))?;
    let cmf = header[0];
    let flags = header[1];
    let check = u16::from_be_bytes([cmf, flags]);
    if cmf & 0x0F != 8 || cmf >> 4 > 7 || check % 31 != 0 {
        return Err(parse("compressed metafile is not RFC1950-wrapped DEFLATE"));
    }
    if flags & 0x20 != 0 {
        return Err(parse("RFC1950 preset dictionaries are not supported"));
    }
    Ok(())
}

fn decode_bitmap(blip: &Blip<'_>, options: Options) -> Result<DynamicImage> {
    let data = decode_data(blip, &options.limits)?;
    let adapted = if matches!(blip, Blip::Dib(_)) {
        dib_to_bmp(&data, &options.limits)?
    } else {
        data
    };
    let probe = ImageReader::new(Cursor::new(adapted.as_ref()))
        .with_guessed_format()
        .map_err(|error| parse(format!("failed to identify bitmap: {error}")))?;
    let (source_width, source_height) = probe
        .into_dimensions()
        .map_err(|error| parse(format!("failed to read bitmap dimensions: {error}")))?;
    check_dimensions(source_width, source_height, &options.limits)?;
    let cursor = Cursor::new(adapted.as_ref());
    let mut reader = ImageReader::new(cursor)
        .with_guessed_format()
        .map_err(|error| parse(format!("failed to identify bitmap: {error}")))?;
    let mut image_limits = image::Limits::default();
    image_limits.max_image_width = Some(options.limits.max_width);
    image_limits.max_image_height = Some(options.limits.max_height);
    image_limits.max_alloc = Some(
        options
            .limits
            .max_pixels
            .checked_mul(8)
            .ok_or_else(|| parse("pixel allocation limit overflows"))?,
    );
    reader.limits(image_limits);
    let image = reader
        .decode()
        .map_err(|error| parse(format!("failed to decode bitmap: {error}")))?;
    let (width, height) = target_dimensions(
        image.width(),
        image.height(),
        options.width,
        options.height,
        &options.limits,
    )?;
    if image.width() == width && image.height() == height {
        Ok(image)
    } else {
        Ok(DynamicImage::ImageRgba8(image::imageops::resize(
            &image,
            width,
            height,
            image::imageops::FilterType::Lanczos3,
        )))
    }
}

fn render_emf(meta: &Meta<'_>, format: ImageFormat, options: Options) -> Result<Vec<u8>> {
    let data = inflate_checked(meta, &options.limits)?;
    rasterize_raw_metafile(&data, InputFormat::Emf, format, options)
}

fn render_wmf(meta: &Meta<'_>, format: ImageFormat, options: Options) -> Result<Vec<u8>> {
    let data = inflate_checked(meta, &options.limits)?;
    let data = wmf_with_header(meta, data, &options.limits)?;
    rasterize_raw_metafile(&data, InputFormat::Wmf, format, options)
}

fn render_pict(meta: &Meta<'_>, options: Options) -> Result<DynamicImage> {
    let data = inflate_checked(meta, &options.limits)?;
    let parser = crate::pict::PictParser::new_bounded(
        &data,
        options.limits.max_records,
        options.limits.max_uncompressed_bytes,
    )?;
    let (width, height) = target_dimensions(
        positive(parser.width()),
        positive(parser.height()),
        options.width,
        options.height,
        &options.limits,
    )?;
    crate::pict::PictConverter::new(
        parser,
        crate::pict::PictToRasterOptions {
            width: Some(width),
            height: Some(height),
            background_color: image::Rgba([255, 255, 255, 255]),
            max_width: options.limits.max_width,
            max_height: options.limits.max_height,
            max_pixels: options.limits.max_pixels,
            max_output_bytes: options.limits.max_output_bytes,
        },
    )
    .convert_to_image()
}

fn inflate_checked<'data>(meta: &Meta<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    check_encoded(meta.data().len(), limits)?;
    inflate(meta, limits)
}

fn positive(value: i32) -> u32 {
    value.unsigned_abs().max(1)
}

fn target_dimensions(
    source_width: u32,
    source_height: u32,
    width: Option<u32>,
    height: Option<u32>,
    limits: &Limits,
) -> Result<(u32, u32)> {
    let source_width = source_width.max(1);
    let source_height = source_height.max(1);
    let (width, height) = match (width, height) {
        (Some(0), _) | (_, Some(0)) => return Err(parse("target dimensions must be nonzero")),
        (Some(width), Some(height)) => (width, height),
        (Some(width), None) => {
            let height = proportional(source_height, width, source_width)?;
            (width, height)
        },
        (None, Some(height)) => {
            let width = proportional(source_width, height, source_height)?;
            (width, height)
        },
        (None, None) => (source_width, source_height),
    };
    check_dimensions(width, height, limits)?;
    Ok((width, height))
}

fn proportional(source_other: u32, target: u32, source_axis: u32) -> Result<u32> {
    let numerator = u64::from(source_other)
        .checked_mul(u64::from(target))
        .ok_or_else(|| parse("aspect-ratio calculation overflows"))?;
    let rounded = numerator
        .checked_add(u64::from(source_axis / 2))
        .ok_or_else(|| parse("aspect-ratio rounding overflows"))?
        / u64::from(source_axis);
    u32::try_from(rounded.max(1)).map_err(|_| parse("target dimension exceeds u32"))
}

fn check_dimensions(width: u32, height: u32, limits: &Limits) -> Result<()> {
    if width > limits.max_width {
        return Err(bound("image width", width, limits.max_width));
    }
    if height > limits.max_height {
        return Err(bound("image height", height, limits.max_height));
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

fn wmf_with_header<'data>(
    meta: &Meta<'data>,
    data: Cow<'data, [u8]>,
    limits: &Limits,
) -> Result<Cow<'data, [u8]>> {
    if data.starts_with(&0x9AC6_CDD7u32.to_le_bytes()) {
        return Ok(data);
    }
    let total = data
        .len()
        .checked_add(22)
        .ok_or_else(|| parse("WMF placeable-header length overflows"))?;
    if total > limits.max_uncompressed_bytes {
        return Err(limit(
            "WMF with placeable header",
            limits.max_uncompressed_bytes,
        ));
    }
    let header = meta.header();
    let bounds = if header.bounds == litchi_odraw::image::Rect::default() {
        let width = emu_to_twips(header.extent.x)?;
        let height = emu_to_twips(header.extent.y)?;
        (0, 0, width, height)
    } else {
        (
            checked_i16(header.bounds.left, "WMF left bound")?,
            checked_i16(header.bounds.top, "WMF top bound")?,
            checked_i16(header.bounds.right, "WMF right bound")?,
            checked_i16(header.bounds.bottom, "WMF bottom bound")?,
        )
    };
    let mut output = Vec::with_capacity(total);
    output.extend_from_slice(&0x9AC6_CDD7u32.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output.extend_from_slice(&bounds.0.to_le_bytes());
    output.extend_from_slice(&bounds.1.to_le_bytes());
    output.extend_from_slice(&bounds.2.to_le_bytes());
    output.extend_from_slice(&bounds.3.to_le_bytes());
    output.extend_from_slice(&1440u16.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    let checksum = output[..20].chunks_exact(2).fold(0u16, |sum, bytes| {
        sum ^ u16::from_le_bytes([bytes[0], bytes[1]])
    });
    output.extend_from_slice(&checksum.to_le_bytes());
    output.extend_from_slice(&data);
    Ok(Cow::Owned(output))
}

fn emu_to_twips(value: i32) -> Result<i16> {
    let scaled = i64::from(value)
        .checked_mul(1440)
        .ok_or_else(|| parse("EMU-to-twips conversion overflows"))?
        / 914_400;
    i16::try_from(scaled).map_err(|_| parse("WMF extent does not fit a placeable header"))
}

fn checked_i16(value: i32, field: &str) -> Result<i16> {
    i16::try_from(value).map_err(|_| parse(format!("{field} does not fit i16")))
}

fn dib_to_bmp<'data>(data: &'data [u8], limits: &Limits) -> Result<Cow<'data, [u8]>> {
    if data.starts_with(b"BM") {
        return Ok(Cow::Borrowed(data));
    }
    let header_size = read_u32(data, 0, "DIB header size")?;
    let header_size_usize =
        usize::try_from(header_size).map_err(|_| parse("DIB header size does not fit usize"))?;
    if header_size_usize < 12 || header_size_usize > data.len() {
        return Err(parse("DIB header size is invalid"));
    }
    let (bit_count, colors_used, palette_entry, masks) = if header_size == 12 {
        (
            u32::from(read_u16(data, 10, "DIB bit count")?),
            0,
            3usize,
            0usize,
        )
    } else if header_size >= 40 {
        let bit_count = u32::from(read_u16(data, 14, "DIB bit count")?);
        let compression = read_u32(data, 16, "DIB compression")?;
        let colors_used = read_u32(data, 32, "DIB palette size")?;
        let masks = if header_size == 40 {
            match compression {
                3 => 12,
                6 => 16,
                _ => 0,
            }
        } else {
            0
        };
        (bit_count, colors_used, 4usize, masks)
    } else {
        return Err(parse("unsupported DIB header layout"));
    };
    let palette_colors = if colors_used != 0 {
        colors_used
    } else if bit_count <= 8 {
        1u32.checked_shl(bit_count)
            .ok_or_else(|| parse("DIB palette size overflows"))?
    } else {
        0
    };
    let palette = usize::try_from(palette_colors)
        .ok()
        .and_then(|count| count.checked_mul(palette_entry))
        .ok_or_else(|| parse("DIB palette byte length overflows"))?;
    let pixel_offset = 14usize
        .checked_add(header_size_usize)
        .and_then(|value| value.checked_add(masks))
        .and_then(|value| value.checked_add(palette))
        .ok_or_else(|| parse("BMP pixel offset overflows"))?;
    let total = data
        .len()
        .checked_add(14)
        .ok_or_else(|| parse("BMP file size overflows"))?;
    if total > limits.max_encoded_bytes.saturating_add(14) {
        return Err(limit("adapted BMP", limits.max_encoded_bytes));
    }
    if pixel_offset > total {
        return Err(parse("DIB palette extends past the image data"));
    }
    let total_u32 = u32::try_from(total).map_err(|_| parse("BMP file size exceeds u32"))?;
    let pixel_u32 =
        u32::try_from(pixel_offset).map_err(|_| parse("BMP pixel offset exceeds u32"))?;
    let mut bmp = Vec::with_capacity(total);
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&total_u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&pixel_u32.to_le_bytes());
    bmp.extend_from_slice(data);
    Ok(Cow::Owned(bmp))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| parse(format!("{field} extent overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| parse(format!("{field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| parse(format!("{field} extent overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| parse(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

pub(crate) fn encode(image: &DynamicImage, format: ImageFormat, maximum: usize) -> Result<Vec<u8>> {
    let mut output = LimitedCursor::new(maximum)?;
    image
        .write_to(&mut output, format)
        .map_err(|error| parse(format!("failed to encode image: {error}")))?;
    Ok(output.into_inner())
}

struct LimitedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: u64,
}

impl LimitedCursor {
    fn new(maximum: usize) -> Result<Self> {
        Ok(Self {
            inner: Cursor::new(Vec::new()),
            maximum: u64::try_from(maximum)
                .map_err(|_| parse("encoded output limit does not fit u64"))?,
        })
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }
}

impl Write for LimitedCursor {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let length = u64::try_from(buffer.len())
            .map_err(|_| std::io::Error::other("encoded write length does not fit u64"))?;
        let end = self
            .inner
            .position()
            .checked_add(length)
            .ok_or_else(|| std::io::Error::other("encoded output position overflows"))?;
        if end > self.maximum {
            return Err(std::io::Error::other("encoded output limit exceeded"));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for LimitedCursor {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let previous = self.inner.position();
        let next = self.inner.seek(position)?;
        if next > self.maximum {
            self.inner.set_position(previous);
            return Err(std::io::Error::other("encoded output seek exceeds limit"));
        }
        Ok(next)
    }
}

fn parse(error: impl Into<String>) -> Error {
    Error::ParseError(error.into())
}

fn limit(resource: &str, maximum: usize) -> Error {
    parse(format!("{resource} exceeds the {maximum}-byte limit"))
}

fn bound(resource: &str, actual: u32, maximum: u32) -> Error {
    parse(format!("{resource} is {actual}; limit is {maximum}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::{
        Compression as Level,
        write::{DeflateEncoder, ZlibEncoder},
    };
    use litchi_odraw::image::Blip;

    fn empty_emf(width: i32, height: i32) -> Vec<u8> {
        let mut data = vec![0_u8; 88];
        data[0..4].copy_from_slice(&1_u32.to_le_bytes());
        data[4..8].copy_from_slice(&88_u32.to_le_bytes());
        data[16..20].copy_from_slice(&width.to_le_bytes());
        data[20..24].copy_from_slice(&height.to_le_bytes());
        data[40..44].copy_from_slice(&0x464D_4520_u32.to_le_bytes());
        data[44..48].copy_from_slice(&0x0001_0000_u32.to_le_bytes());
        append_emf_eof(&mut data);
        let length = u32::try_from(data.len()).unwrap();
        data[48..52].copy_from_slice(&length.to_le_bytes());
        data[52..56].copy_from_slice(&2_u32.to_le_bytes());
        data
    }

    fn rectangle_emf(width: i32, height: i32) -> Vec<u8> {
        let mut data = vec![0_u8; 88];
        data[0..4].copy_from_slice(&1_u32.to_le_bytes());
        data[4..8].copy_from_slice(&88_u32.to_le_bytes());
        data[16..20].copy_from_slice(&width.to_le_bytes());
        data[20..24].copy_from_slice(&height.to_le_bytes());
        data[40..44].copy_from_slice(&0x464D_4520_u32.to_le_bytes());
        data[44..48].copy_from_slice(&0x0001_0000_u32.to_le_bytes());
        data.extend_from_slice(&43_u32.to_le_bytes());
        data.extend_from_slice(&24_u32.to_le_bytes());
        data.extend_from_slice(&0_i32.to_le_bytes());
        data.extend_from_slice(&0_i32.to_le_bytes());
        data.extend_from_slice(&width.to_le_bytes());
        data.extend_from_slice(&height.to_le_bytes());
        append_emf_eof(&mut data);
        let length = u32::try_from(data.len()).unwrap();
        data[48..52].copy_from_slice(&length.to_le_bytes());
        data[52..56].copy_from_slice(&3_u32.to_le_bytes());
        data
    }

    fn emfplus_metafile(
        width: i32,
        height: i32,
        dual: bool,
        get_dc: bool,
        rendering_hint: bool,
    ) -> Vec<u8> {
        fn plus_record(kind: u16, flags: u16, body: &[u8]) -> Vec<u8> {
            let size = 12_u32 + u32::try_from(body.len()).unwrap();
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&kind.to_le_bytes());
            bytes.extend_from_slice(&flags.to_le_bytes());
            bytes.extend_from_slice(&size.to_le_bytes());
            bytes.extend_from_slice(&u32::try_from(body.len()).unwrap().to_le_bytes());
            bytes.extend_from_slice(body);
            bytes
        }

        let mut payload = plus_record(0x4001, u16::from(dual), &[0; 16]);
        payload.extend(plus_record(0x4009, 0, &0xff11_2233_u32.to_le_bytes()));
        if rendering_hint {
            payload.extend(plus_record(0x401e, 2, &[]));
        }
        if get_dc {
            payload.extend(plus_record(0x4004, 0, &[]));
        }
        payload.extend(plus_record(0x4002, 0, &[]));

        let data_size = 4_u32 + u32::try_from(payload.len()).unwrap();
        let mut comment_body = Vec::new();
        comment_body.extend_from_slice(&data_size.to_le_bytes());
        comment_body.extend_from_slice(b"EMF+");
        comment_body.extend_from_slice(&payload);
        while (comment_body.len() + 8) % 4 != 0 {
            comment_body.push(0);
        }

        let mut data = vec![0_u8; 88];
        data[0..4].copy_from_slice(&1_u32.to_le_bytes());
        data[4..8].copy_from_slice(&88_u32.to_le_bytes());
        data[16..20].copy_from_slice(&width.to_le_bytes());
        data[20..24].copy_from_slice(&height.to_le_bytes());
        data[40..44].copy_from_slice(&0x464D_4520_u32.to_le_bytes());
        data[44..48].copy_from_slice(&0x0001_0000_u32.to_le_bytes());
        data.extend_from_slice(&70_u32.to_le_bytes());
        data.extend_from_slice(&u32::try_from(comment_body.len() + 8).unwrap().to_le_bytes());
        data.extend_from_slice(&comment_body);
        append_emf_eof(&mut data);
        let length = u32::try_from(data.len()).unwrap();
        data[48..52].copy_from_slice(&length.to_le_bytes());
        data[52..56].copy_from_slice(&3_u32.to_le_bytes());
        data
    }

    fn append_emf_eof(data: &mut Vec<u8>) {
        data.extend_from_slice(&14_u32.to_le_bytes());
        data.extend_from_slice(&20_u32.to_le_bytes());
        data.extend_from_slice(&0_u32.to_le_bytes());
        data.extend_from_slice(&0_u32.to_le_bytes());
        data.extend_from_slice(&20_u32.to_le_bytes());
    }

    fn record(instance: u16, kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&(instance << 4).to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    fn meta(compressed: &[u8], uncompressed_size: u32) -> Vec<u8> {
        let mut body = vec![0; 16];
        body.extend_from_slice(&uncompressed_size.to_le_bytes());
        body.extend_from_slice(&[0; 24]);
        body.extend_from_slice(&(compressed.len() as u32).to_le_bytes());
        body.extend_from_slice(&[0x00, 0xFE]);
        body.extend_from_slice(compressed);
        record(0x3D4, 0xF01A, &body)
    }

    #[test]
    fn accepts_only_rfc1950_and_requires_exact_cb_size() {
        let mut zlib = ZlibEncoder::new(Vec::new(), Level::default());
        zlib.write_all(b"metafile").unwrap();
        let encoded = zlib.finish().unwrap();
        let bytes = meta(&encoded, 8);
        let blip = Blip::parse(&bytes).unwrap();
        assert_eq!(
            decode_data(&blip, &Limits::default()).unwrap(),
            b"metafile".as_slice()
        );

        let wrong = meta(&encoded, 9);
        let wrong = Blip::parse(&wrong).unwrap();
        assert!(decode_data(&wrong, &Limits::default()).is_err());

        let mut raw = DeflateEncoder::new(Vec::new(), Level::default());
        raw.write_all(b"metafile").unwrap();
        let raw = raw.finish().unwrap();
        let raw = meta(&raw, 8);
        let raw = Blip::parse(&raw).unwrap();
        assert!(decode_data(&raw, &Limits::default()).is_err());
    }

    #[test]
    fn enforces_decompression_limit_before_allocation() {
        let mut zlib = ZlibEncoder::new(Vec::new(), Level::default());
        zlib.write_all(&[0; 1024]).unwrap();
        let encoded = zlib.finish().unwrap();
        let bytes = meta(&encoded, 1024);
        let blip = Blip::parse(&bytes).unwrap();
        let limits = Limits {
            max_uncompressed_bytes: 32,
            ..Limits::default()
        };
        assert!(decode_data(&blip, &limits).is_err());
    }

    #[test]
    fn bitmap_decode_honors_pixel_and_output_limits() {
        let image = DynamicImage::new_rgba8(4, 4);
        let mut png = Cursor::new(Vec::new());
        image.write_to(&mut png, ImageFormat::Png).unwrap();
        let png = png.into_inner();
        let mut body = vec![0; 16];
        body.push(0xFF);
        body.extend_from_slice(&png);
        let bytes = record(0x6E0, 0xF01E, &body);
        let blip = Blip::parse(&bytes).unwrap();
        let limits = Limits {
            max_pixels: 15,
            ..Limits::default()
        };
        assert!(to_png(&blip, Options::default().limits(limits)).is_err());

        let limits = Limits {
            max_output_bytes: 4,
            ..Limits::default()
        };
        assert!(to_png(&blip, Options::default().limits(limits)).is_err());
    }

    #[test]
    fn raw_metafile_api_returns_typed_vector_first_output() {
        let converted = convert_metafile(
            &empty_emf(10, 20),
            InputFormat::Emf,
            OutputFormat::Auto,
            Options::default(),
        )
        .unwrap();
        assert_eq!(converted.format, ConvertedFormat::Svg);
        assert_eq!(converted.mime_type, "image/svg+xml");
        assert_eq!(converted.extension, "svg");
        assert_eq!(converted.report.diagnostics[0].code, "auto-vector-first");
        assert!(
            String::from_utf8(converted.bytes)
                .unwrap()
                .starts_with("<svg")
        );
    }

    #[test]
    fn raw_metafile_api_enforces_encoded_input_limit() {
        let limits = Limits {
            max_encoded_bytes: 8,
            ..Limits::default()
        };
        assert!(
            convert_metafile(
                &empty_emf(10, 20),
                InputFormat::Emf,
                OutputFormat::Svg,
                Options::default().limits(limits),
            )
            .is_err()
        );
    }

    #[test]
    fn raw_metafile_png_rasterizes_svg_content() {
        let converted = convert_metafile(
            &rectangle_emf(10, 20),
            InputFormat::Emf,
            OutputFormat::Png,
            Options::default(),
        )
        .unwrap();
        assert_eq!(converted.format, ConvertedFormat::Png);
        let rendered =
            image::load_from_memory_with_format(&converted.bytes, ImageFormat::Png).unwrap();
        assert_eq!((rendered.width(), rendered.height()), (10, 20));
    }

    #[test]
    fn public_convert_metafile_renders_emfplus_dual_once() {
        let converted = convert_metafile(
            &emfplus_metafile(10, 20, true, false, false),
            InputFormat::Emf,
            OutputFormat::Svg,
            Options::default(),
        )
        .unwrap();
        assert_eq!(converted.format, ConvertedFormat::Svg);
        let svg = String::from_utf8(converted.bytes).unwrap();
        assert!(svg.contains("<rect"));
        assert!(svg.contains("#112233"));
        assert_eq!(svg.matches("<rect").count(), 1);
    }

    #[test]
    fn public_convert_metafile_handles_emfplus_mux_strictly() {
        let converted = convert_metafile(
            &emfplus_metafile(10, 20, false, false, false),
            InputFormat::Emf,
            OutputFormat::Svg,
            Options::default(),
        )
        .unwrap();
        assert!(
            String::from_utf8(converted.bytes)
                .unwrap()
                .contains("<rect")
        );

        let error = convert_metafile(
            &emfplus_metafile(10, 20, false, true, false),
            InputFormat::Emf,
            OutputFormat::Svg,
            Options::default(),
        )
        .unwrap_err();
        assert!(matches!(error, Error::Unsupported(message) if message.contains("GetDC")));

        let dual_error = convert_metafile(
            &emfplus_metafile(10, 20, true, true, false),
            InputFormat::Emf,
            OutputFormat::Svg,
            Options::default(),
        )
        .unwrap_err();
        assert!(matches!(dual_error, Error::Unsupported(message) if message.contains("GetDC")));
    }

    #[test]
    fn typed_api_reports_approximations_and_convenience_api_is_strict() {
        let bytes = emfplus_metafile(10, 20, true, false, true);
        let converted = convert_metafile(
            &bytes,
            InputFormat::Emf,
            OutputFormat::Svg,
            Options::default(),
        )
        .unwrap();
        assert!(
            converted
                .report
                .diagnostics
                .iter()
                .any(|diagnostic| { diagnostic.code == "rendering_property_not_represented" })
        );
        assert!(crate::emf::convert_emf_to_svg(&bytes).is_err());
        assert!(
            rasterize_raw_metafile(
                &bytes,
                InputFormat::Emf,
                ImageFormat::WebP,
                Options::default()
            )
            .is_err()
        );
    }
}
