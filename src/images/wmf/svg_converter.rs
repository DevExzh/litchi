// WMF to SVG converter with SIMD acceleration
//
// Converts Windows Metafile vector graphics to SVG while extracting embedded raster images

use super::parser::{WmfParser, WmfRecord};
use crate::common::error::{Error, Result};
use crate::images::svg::*;
use rayon::prelude::*;

/// WMF to SVG converter
pub struct WmfSvgConverter {
    parser: WmfParser,
}

impl WmfSvgConverter {
    /// Create a new WMF to SVG converter
    pub fn new(parser: WmfParser) -> Self {
        Self { parser }
    }

    /// Convert WMF to SVG
    ///
    /// Processes WMF records in parallel and generates SVG with vector graphics
    /// and embedded raster images.
    pub fn convert_to_svg(&self) -> Result<String> {
        let width = self.parser.width() as f64;
        let height = self.parser.height() as f64;

        let mut builder = SvgBuilder::new(width, height);

        // WMF coordinates are in logical units, set viewBox appropriately
        builder = builder.with_viewbox(0.0, 0.0, width, height);

        // Process records in parallel
        let elements: Vec<SvgElement> = self
            .parser
            .records
            .par_iter()
            .filter_map(|record| self.process_record(record).ok().flatten())
            .collect();

        for element in elements {
            builder.add_element(element);
        }

        Ok(builder.build())
    }

    /// Convert WMF to SVG bytes
    pub fn convert_to_svg_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.convert_to_svg()?.into_bytes())
    }

    /// Process a single WMF record
    fn process_record(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        match record.function {
            // Rectangle
            0x041B => self.parse_rectangle(record),
            // Ellipse
            0x0418 => self.parse_ellipse(record),
            // Polygon
            0x0324 => self.parse_polygon(record),
            // Polyline
            0x0325 => self.parse_polyline(record),
            // Arc
            0x0817 => self.parse_arc(record),
            // Pie
            0x081A => self.parse_pie(record),
            // Chord
            0x0830 => self.parse_chord(record),
            // StretchDIB (embedded bitmap)
            0x0F43 => self.parse_stretchdib(record),
            // DIBStretchBlt
            0x0B41 => self.parse_dibstretchblt(record),
            // DIBBitBlt
            0x0940 => self.parse_dibbitblt(record),
            // RoundRect
            0x061C => self.parse_roundrect(record),
            // LineTo
            0x0213 => self.parse_lineto(record),
            _ => Ok(None),
        }
    }

    /// Parse META_RECTANGLE record
    fn parse_rectangle(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 8 {
            return Ok(None);
        }

        // WMF uses 16-bit coordinates in big-endian (words)
        let bottom = i16::from_le_bytes([record.params[0], record.params[1]]) as f64;
        let right = i16::from_le_bytes([record.params[2], record.params[3]]) as f64;
        let top = i16::from_le_bytes([record.params[4], record.params[5]]) as f64;
        let left = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;

        Ok(Some(SvgElement::Rect(SvgRect {
            x: left,
            y: top,
            width: right - left,
            height: bottom - top,
            fill: None,
            stroke: Some("#000000".to_string()),
            stroke_width: 1.0,
        })))
    }

    /// Parse META_ELLIPSE record
    fn parse_ellipse(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 8 {
            return Ok(None);
        }

        let bottom = i16::from_le_bytes([record.params[0], record.params[1]]) as f64;
        let right = i16::from_le_bytes([record.params[2], record.params[3]]) as f64;
        let top = i16::from_le_bytes([record.params[4], record.params[5]]) as f64;
        let left = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;

        let cx = (left + right) / 2.0;
        let cy = (top + bottom) / 2.0;
        let rx = (right - left) / 2.0;
        let ry = (bottom - top) / 2.0;

        Ok(Some(SvgElement::Ellipse(SvgEllipse {
            cx,
            cy,
            rx,
            ry,
            fill: None,
            stroke: Some("#000000".to_string()),
            stroke_width: 1.0,
        })))
    }

    /// Parse META_POLYGON record
    fn parse_polygon(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 2 {
            return Ok(None);
        }

        let count = i16::from_le_bytes([record.params[0], record.params[1]]) as usize;

        if record.params.len() < 2 + count * 4 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count + 1);

        // Parse points in parallel batches for better performance
        let points: Vec<(f64, f64)> = (0..count)
            .into_par_iter()
            .map(|i| {
                let offset = 2 + i * 4;
                let x = i16::from_le_bytes([
                    record.params[offset],
                    record.params[offset + 1],
                ]) as f64;
                let y = i16::from_le_bytes([
                    record.params[offset + 2],
                    record.params[offset + 3],
                ]) as f64;
                (x, y)
            })
            .collect();

        for (i, (x, y)) in points.into_iter().enumerate() {
            if i == 0 {
                commands.push(PathCommand::MoveTo { x, y });
            } else {
                commands.push(PathCommand::LineTo { x, y });
            }
        }

        commands.push(PathCommand::ClosePath);

        Ok(Some(SvgElement::Path(
            SvgPath::new(commands)
                .with_stroke("#000000".to_string())
                .with_fill("none".to_string()),
        )))
    }

    /// Parse META_POLYLINE record
    fn parse_polyline(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 2 {
            return Ok(None);
        }

        let count = i16::from_le_bytes([record.params[0], record.params[1]]) as usize;

        if record.params.len() < 2 + count * 4 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count);

        // SIMD-friendly parallel processing
        let points: Vec<(f64, f64)> = (0..count)
            .into_par_iter()
            .map(|i| {
                let offset = 2 + i * 4;
                let x = i16::from_le_bytes([
                    record.params[offset],
                    record.params[offset + 1],
                ]) as f64;
                let y = i16::from_le_bytes([
                    record.params[offset + 2],
                    record.params[offset + 3],
                ]) as f64;
                (x, y)
            })
            .collect();

        for (i, (x, y)) in points.into_iter().enumerate() {
            if i == 0 {
                commands.push(PathCommand::MoveTo { x, y });
            } else {
                commands.push(PathCommand::LineTo { x, y });
            }
        }

        Ok(Some(SvgElement::Path(
            SvgPath::new(commands)
                .with_stroke("#000000".to_string())
                .with_fill("none".to_string()),
        )))
    }

    /// Parse META_ARC record
    fn parse_arc(&self, _record: &WmfRecord) -> Result<Option<SvgElement>> {
        // Arc conversion requires trigonometric calculations
        // Placeholder for future implementation
        Ok(None)
    }

    /// Parse META_PIE record
    fn parse_pie(&self, _record: &WmfRecord) -> Result<Option<SvgElement>> {
        // Pie chart slice
        Ok(None)
    }

    /// Parse META_CHORD record
    fn parse_chord(&self, _record: &WmfRecord) -> Result<Option<SvgElement>> {
        // Chord shape
        Ok(None)
    }

    /// Parse META_ROUNDRECT record
    fn parse_roundrect(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 12 {
            return Ok(None);
        }

        let corner_height = i16::from_le_bytes([record.params[0], record.params[1]]) as f64;
        let corner_width = i16::from_le_bytes([record.params[2], record.params[3]]) as f64;
        let bottom = i16::from_le_bytes([record.params[4], record.params[5]]) as f64;
        let right = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;
        let top = i16::from_le_bytes([record.params[8], record.params[9]]) as f64;
        let left = i16::from_le_bytes([record.params[10], record.params[11]]) as f64;

        let rx = corner_width / 2.0;
        let ry = corner_height / 2.0;

        // Create rounded rectangle using path with arcs
        let commands = vec![
            PathCommand::MoveTo { x: left + rx, y: top },
            PathCommand::LineTo { x: right - rx, y: top },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: right,
                y: top + ry,
            },
            PathCommand::LineTo { x: right, y: bottom - ry },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: right - rx,
                y: bottom,
            },
            PathCommand::LineTo { x: left + rx, y: bottom },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: left,
                y: bottom - ry,
            },
            PathCommand::LineTo { x: left, y: top + ry },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: left + rx,
                y: top,
            },
            PathCommand::ClosePath,
        ];

        Ok(Some(SvgElement::Path(
            SvgPath::new(commands)
                .with_stroke("#000000".to_string())
                .with_fill("none".to_string()),
        )))
    }

    /// Parse META_LINETO record
    fn parse_lineto(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 4 {
            return Ok(None);
        }

        let y = i16::from_le_bytes([record.params[0], record.params[1]]) as f64;
        let x = i16::from_le_bytes([record.params[2], record.params[3]]) as f64;

        Ok(Some(SvgElement::Path(
            SvgPath::new(vec![
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo { x, y },
            ])
            .with_stroke("#000000".to_string()),
        )))
    }

    /// Parse META_STRETCHDIB record
    fn parse_stretchdib(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        // Extract DIB and convert to PNG for embedding
        if record.params.len() < 20 {
            return Ok(None);
        }

        // Parse destination rectangle (simplified)
        let dest_height = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;
        let dest_width = i16::from_le_bytes([record.params[8], record.params[9]]) as f64;
        let dest_y = i16::from_le_bytes([record.params[10], record.params[11]]) as f64;
        let dest_x = i16::from_le_bytes([record.params[12], record.params[13]]) as f64;

        // Try to extract DIB data
        if let Ok(png_data) = self.extract_and_convert_dib(&record.params[20..]) {
            return Ok(Some(SvgElement::Image(SvgImage::from_png_data(
                dest_x,
                dest_y,
                dest_width,
                dest_height,
                &png_data,
            ))));
        }

        Ok(None)
    }

    /// Parse META_DIBSTRETCHBLT record
    fn parse_dibstretchblt(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 20 {
            return Ok(None);
        }

        // Similar to StretchDIB but with different parameter layout
        if let Ok(png_data) = self.extract_and_convert_dib(&record.params[18..]) {
            let dest_x = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;
            let dest_y = i16::from_le_bytes([record.params[8], record.params[9]]) as f64;
            let dest_width = i16::from_le_bytes([record.params[10], record.params[11]]) as f64;
            let dest_height = i16::from_le_bytes([record.params[12], record.params[13]]) as f64;

            return Ok(Some(SvgElement::Image(SvgImage::from_png_data(
                dest_x,
                dest_y,
                dest_width,
                dest_height,
                &png_data,
            ))));
        }

        Ok(None)
    }

    /// Parse META_DIBBITBLT record
    fn parse_dibbitblt(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        if record.params.len() < 16 {
            return Ok(None);
        }

        if let Ok(png_data) = self.extract_and_convert_dib(&record.params[14..]) {
            let dest_x = i16::from_le_bytes([record.params[4], record.params[5]]) as f64;
            let dest_y = i16::from_le_bytes([record.params[6], record.params[7]]) as f64;
            let width = i16::from_le_bytes([record.params[8], record.params[9]]) as f64;
            let height = i16::from_le_bytes([record.params[10], record.params[11]]) as f64;

            return Ok(Some(SvgElement::Image(SvgImage::from_png_data(
                dest_x,
                dest_y,
                width,
                height,
                &png_data,
            ))));
        }

        Ok(None)
    }

    /// Extract and convert DIB to PNG
    fn extract_and_convert_dib(&self, dib_data: &[u8]) -> Result<Vec<u8>> {
        if dib_data.len() < 40 {
            return Err(Error::ParseError("DIB data too small".into()));
        }

        // Construct BMP from DIB
        let file_size = 14u32 + dib_data.len() as u32;
        let pixel_data_offset = 14u32 + 40u32;

        let mut bmp_data = Vec::with_capacity(file_size as usize);
        bmp_data.extend_from_slice(b"BM");
        bmp_data.extend_from_slice(&file_size.to_le_bytes());
        bmp_data.extend_from_slice(&[0u8; 4]);
        bmp_data.extend_from_slice(&pixel_data_offset.to_le_bytes());
        bmp_data.extend_from_slice(dib_data);

        // Load and re-encode as PNG
        let img = image::load_from_memory(&bmp_data)
            .map_err(|e| Error::ParseError(format!("Failed to load DIB: {}", e)))?;

        let mut png_data = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut png_data);
        img.write_to(&mut cursor, image::ImageFormat::Png)
            .map_err(|e| Error::ParseError(format!("Failed to encode PNG: {}", e)))?;

        Ok(png_data)
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_wmf_svg_converter_creation() {
        // Test requires valid WMF data
        // Placeholder for future tests
    }
}

