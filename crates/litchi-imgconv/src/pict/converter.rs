// PICT to raster image converter
//
// Plays a parsed picture back to SVG and rasterizes it. Like the other
// strict helpers, a picture that can only be drawn approximately is
// rejected instead of being rendered differently from QuickDraw.

use image::{DynamicImage, ImageFormat, Rgba};
use litchi_core::error::{Error, Result};

use super::parser::PictParser;
use crate::codec::{ConversionReport, ConvertedFormat, reject_lossy_diagnostics};
use crate::raster::{RasterLimits, rasterize_svg};
use crate::{InputFormat, OutputFormat};

/// Options for PICT to raster conversion
#[derive(Debug, Clone)]
pub struct PictToRasterOptions {
    /// Target width (None = use source dimensions)
    pub width: Option<u32>,
    /// Target height (None = use source dimensions)
    pub height: Option<u32>,
    /// Background color for rendering
    pub background_color: Rgba<u8>,
    /// Maximum width of the output or of an embedded bitmap.
    pub max_width: u32,
    /// Maximum height of the output or of an embedded bitmap.
    pub max_height: u32,
    /// Maximum pixels of the output or of an embedded bitmap.
    pub max_pixels: u64,
    /// Maximum encoded output bytes.
    pub max_output_bytes: usize,
}

impl Default for PictToRasterOptions {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            background_color: Rgba([255, 255, 255, 255]),
            max_width: 8192,
            max_height: 8192,
            max_pixels: 32 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

/// PICT to raster converter
pub struct PictConverter {
    parser: PictParser,
    options: PictToRasterOptions,
}

impl PictConverter {
    /// Create a new PICT converter
    pub fn new(parser: PictParser, options: PictToRasterOptions) -> Self {
        Self { parser, options }
    }

    fn limits(&self) -> crate::Limits {
        crate::Limits {
            max_width: self.options.max_width,
            max_height: self.options.max_height,
            max_pixels: self.options.max_pixels,
            max_output_bytes: self.options.max_output_bytes,
            ..crate::Limits::default()
        }
    }

    /// Calculates output dimensions, keeping the aspect ratio when only one
    /// dimension is requested.
    fn dimensions(&self, natural: (u32, u32)) -> Result<(u32, u32)> {
        let (natural_width, natural_height) = (natural.0.max(1), natural.1.max(1));
        let scaled = |source_other: u32, target: u32, source_axis: u32| {
            let value = u64::from(source_other) * u64::from(target) / u64::from(source_axis);
            u32::try_from(value.max(1)).unwrap_or(u32::MAX)
        };
        let (width, height) = match (self.options.width, self.options.height) {
            (Some(width), Some(height)) => (width, height),
            (Some(width), None) => (width, scaled(natural_height, width, natural_width)),
            (None, Some(height)) => (scaled(natural_width, height, natural_height), height),
            (None, None) => (natural_width, natural_height),
        };
        let pixels = u64::from(width) * u64::from(height);
        if width == 0
            || height == 0
            || width > self.options.max_width
            || height > self.options.max_height
            || pixels > self.options.max_pixels
        {
            return Err(Error::ParseError(
                "PICT target image exceeds configured dimensions".into(),
            ));
        }
        Ok((width, height))
    }

    fn rasterize(&self, format: ImageFormat) -> Result<Vec<u8>> {
        let limits = self.limits();
        let rendered = super::render(&self.parser, &limits)?;
        let selected = match format {
            ImageFormat::Jpeg => ConvertedFormat::Jpeg,
            _ => ConvertedFormat::Png,
        };
        reject_lossy_diagnostics(&ConversionReport {
            input: InputFormat::Pict,
            requested: OutputFormat::Png,
            selected,
            diagnostics: rendered.diagnostics,
        })?;
        let (width, height) = self.dimensions((rendered.width, rendered.height))?;
        rasterize_svg(
            &rendered.svg,
            width,
            height,
            self.options.background_color,
            if format == ImageFormat::Jpeg {
                ImageFormat::Jpeg
            } else {
                ImageFormat::Png
            },
            &RasterLimits {
                max_height: limits.max_height,
                max_input_bytes: limits.max_uncompressed_bytes,
                max_output_bytes: limits.max_output_bytes,
                max_pixels: limits.max_pixels,
                max_width: limits.max_width,
            },
        )
    }

    /// Convert PICT to a raster image
    pub fn convert_to_image(&self) -> Result<DynamicImage> {
        let png = self.rasterize(ImageFormat::Png)?;
        image::load_from_memory_with_format(&png, ImageFormat::Png)
            .map_err(|error| Error::ParseError(format!("failed to decode rendered PICT: {error}")))
    }

    /// Convert PICT to specified image format
    pub fn convert_to_format(&self, format: ImageFormat) -> Result<Vec<u8>> {
        match format {
            ImageFormat::Png | ImageFormat::Jpeg => self.rasterize(format),
            _ => {
                let image = self.convert_to_image()?;
                crate::codec::encode(&image, format, self.options.max_output_bytes)
            },
        }
    }

    /// Convert PICT to PNG bytes
    pub fn convert_to_png(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Png)
    }

    /// Convert PICT to JPEG bytes
    pub fn convert_to_jpeg(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Jpeg)
    }

    /// Convert PICT to WebP bytes
    pub fn convert_to_webp(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::WebP)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pict::parser::tests::version2_picture;

    fn picture(records: &[u8]) -> PictParser {
        let data = version2_picture([0, 0, 4, 6], records);
        PictParser::new(&data, &crate::Limits::default()).unwrap()
    }

    #[test]
    fn converter_rasterizes_vector_content_at_the_frame_size() {
        let mut records = 0x0031u16.to_be_bytes().to_vec(); // paintRect
        for value in [0i16, 0, 2, 6] {
            records.extend_from_slice(&value.to_be_bytes());
        }
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let image = PictConverter::new(picture(&records), PictToRasterOptions::default())
            .convert_to_image()
            .unwrap()
            .to_rgba8();
        assert_eq!(image.dimensions(), (6, 4));
        assert_eq!(image.get_pixel(3, 0).0, [0, 0, 0, 255]);
        assert_eq!(image.get_pixel(3, 3).0, [255, 255, 255, 255]);
    }

    #[test]
    fn converter_rejects_approximations_instead_of_drawing_placeholders() {
        let mut records = 0x0033u16.to_be_bytes().to_vec(); // invertRect
        for value in [0i16, 0, 2, 6] {
            records.extend_from_slice(&value.to_be_bytes());
        }
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let converter = PictConverter::new(picture(&records), PictToRasterOptions::default());
        assert!(converter.convert_to_png().is_err());
    }

    #[test]
    fn converter_checks_target_dimensions() {
        let converter = PictConverter::new(
            picture(&0x00FFu16.to_be_bytes()),
            PictToRasterOptions {
                width: Some(10_000),
                ..PictToRasterOptions::default()
            },
        );
        assert!(converter.convert_to_png().is_err());
    }
}
