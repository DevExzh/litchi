//! EMF raster conversion through the SVG playback pipeline.

use std::io::Cursor;

use image::{DynamicImage, ImageFormat, Rgba};
use litchi_core::error::{Error, Result};

use super::parser::EmfParser;

/// Options for EMF raster conversion.
#[derive(Debug, Clone)]
pub struct EmfToRasterOptions {
    /// Target width (None = use source dimensions).
    pub width: Option<u32>,
    /// Target height (None = use source dimensions).
    pub height: Option<u32>,
    /// Background used while flattening transparency.
    pub background_color: Rgba<u8>,
}

impl Default for EmfToRasterOptions {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            background_color: Rgba([255, 255, 255, 255]),
        }
    }
}

/// Compatibility converter for callers that already own an [`EmfParser`].
///
/// New code should use [`crate::convert_metafile`], which accepts explicit
/// resource limits and returns typed output metadata.
pub struct EmfConverter {
    parser: EmfParser,
    options: EmfToRasterOptions,
}

impl EmfConverter {
    /// Creates a converter.
    pub fn new(parser: EmfParser, options: EmfToRasterOptions) -> Self {
        Self { parser, options }
    }

    fn dimensions(&self) -> (u32, u32) {
        let source_width = u32::try_from(self.parser.width()).unwrap_or(1).max(1);
        let source_height = u32::try_from(self.parser.height()).unwrap_or(1).max(1);
        match (self.options.width, self.options.height) {
            (Some(width), Some(height)) => (width.max(1), height.max(1)),
            (Some(width), None) => (
                width.max(1),
                proportional(source_height, width.max(1), source_width),
            ),
            (None, Some(height)) => (
                proportional(source_width, height.max(1), source_height),
                height.max(1),
            ),
            (None, None) => (source_width.min(4096), source_height.min(4096)),
        }
    }

    fn svg(&self) -> Result<String> {
        crate::emf::EmfSvgConverter::new(&self.parser).convert()
    }

    /// Renders EMF content to an in-memory image without using a placeholder.
    pub fn convert_to_image(&self) -> Result<DynamicImage> {
        let png = self.convert_to_format(ImageFormat::Png)?;
        image::load_from_memory_with_format(&png, ImageFormat::Png).map_err(|error| {
            Error::ParseError(format!("failed to decode rendered EMF PNG: {error}"))
        })
    }

    /// Renders EMF content to PNG, JPEG, or WebP.
    pub fn convert_to_format(&self, format: ImageFormat) -> Result<Vec<u8>> {
        let (width, height) = self.dimensions();
        let svg = self.svg()?;
        match format {
            ImageFormat::Png | ImageFormat::Jpeg => crate::raster::rasterize_svg(
                &svg,
                width,
                height,
                self.options.background_color,
                format,
                &crate::raster::RasterLimits::default(),
            ),
            ImageFormat::WebP => {
                let png = crate::raster::rasterize_svg(
                    &svg,
                    width,
                    height,
                    self.options.background_color,
                    ImageFormat::Png,
                    &crate::raster::RasterLimits::default(),
                )?;
                let image = image::load_from_memory_with_format(&png, ImageFormat::Png).map_err(
                    |error| {
                        Error::ParseError(format!("failed to decode rendered EMF PNG: {error}"))
                    },
                )?;
                let mut output = Cursor::new(Vec::new());
                image
                    .write_to(&mut output, ImageFormat::WebP)
                    .map_err(|error| {
                        Error::ParseError(format!("failed to encode rendered EMF WebP: {error}"))
                    })?;
                Ok(output.into_inner())
            },
            _ => Err(Error::Unsupported(format!(
                "unsupported EMF output format: {format:?}"
            ))),
        }
    }

    /// Renders EMF content to PNG.
    pub fn convert_to_png(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Png)
    }

    /// Renders EMF content to JPEG.
    pub fn convert_to_jpeg(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Jpeg)
    }

    /// Renders EMF content to WebP.
    pub fn convert_to_webp(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::WebP)
    }
}

fn proportional(other: u32, target: u32, source: u32) -> u32 {
    u32::try_from(u64::from(other) * u64::from(target) / u64::from(source.max(1)))
        .unwrap_or(u32::MAX)
        .max(1)
}
