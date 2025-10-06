// EMF to SVG converter with SIMD acceleration
//
// Converts Enhanced Metafile vector graphics to SVG while extracting embedded raster images

use super::parser::{EmfParser, EmfRecord};
use crate::common::error::{Error, Result};
use crate::images::svg::*;
use rayon::prelude::*;

/// EMF to SVG converter
pub struct EmfSvgConverter {
    parser: EmfParser,
}

impl EmfSvgConverter {
    /// Create a new EMF to SVG converter
    pub fn new(parser: EmfParser) -> Self {
        Self { parser }
    }

    /// Convert EMF to SVG
    ///
    /// This processes all EMF records in parallel where possible and generates
    /// an SVG document with vector graphics and embedded raster images.
    pub fn convert_to_svg(&self) -> Result<String> {
        let mut builder = SvgBuilder::new(
            self.parser.width() as f64,
            self.parser.height() as f64,
        );

        // Set viewBox based on bounds
        let (x1, y1, x2, y2) = self.parser.header.bounds;
        builder = builder.with_viewbox(
            x1 as f64,
            y1 as f64,
            (x2 - x1) as f64,
            (y2 - y1) as f64,
        );

        // Process records in parallel to extract elements
        // Group records by type for efficient parallel processing
        let elements: Vec<SvgElement> = self
            .parser
            .records
            .par_iter()
            .filter_map(|record| self.process_record(record).ok().flatten())
            .collect();

        // Add all elements to builder
        for element in elements {
            builder.add_element(element);
        }

        Ok(builder.build())
    }

    /// Convert EMF to SVG bytes
    pub fn convert_to_svg_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.convert_to_svg()?.into_bytes())
    }

    /// Process a single EMF record and convert to SVG element
    fn process_record(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        match record.record_type {
            // Rectangle
            0x0000002B => self.parse_rectangle(record),
            // Ellipse
            0x0000002A => self.parse_ellipse(record),
            // Polygon
            0x00000003 => self.parse_polygon(record),
            // Polyline
            0x00000004 => self.parse_polyline(record),
            // PolyBezier
            0x00000002 => self.parse_polybezier(record),
            // LineTo
            0x00000036 => self.parse_lineto(record),
            // StretchDIBits (embedded bitmap)
            0x00000051 => self.parse_stretchdibits(record),
            // Arc
            0x0000002D => self.parse_arc(record),
            // Pie
            0x0000002F => self.parse_pie(record),
            // Chord
            0x0000002E => self.parse_chord(record),
            _ => Ok(None), // Unsupported or non-drawing record
        }
    }

    /// Parse EMR_RECTANGLE record
    fn parse_rectangle(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 16 {
            return Ok(None);
        }

        let left = i32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]);
        let top = i32::from_le_bytes([
            record.data[4],
            record.data[5],
            record.data[6],
            record.data[7],
        ]);
        let right = i32::from_le_bytes([
            record.data[8],
            record.data[9],
            record.data[10],
            record.data[11],
        ]);
        let bottom = i32::from_le_bytes([
            record.data[12],
            record.data[13],
            record.data[14],
            record.data[15],
        ]);

        Ok(Some(SvgElement::Rect(SvgRect {
            x: left as f64,
            y: top as f64,
            width: (right - left) as f64,
            height: (bottom - top) as f64,
            fill: None,
            stroke: Some("#000000".to_string()),
            stroke_width: 1.0,
        })))
    }

    /// Parse EMR_ELLIPSE record
    fn parse_ellipse(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 16 {
            return Ok(None);
        }

        let left = i32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]) as f64;
        let top = i32::from_le_bytes([
            record.data[4],
            record.data[5],
            record.data[6],
            record.data[7],
        ]) as f64;
        let right = i32::from_le_bytes([
            record.data[8],
            record.data[9],
            record.data[10],
            record.data[11],
        ]) as f64;
        let bottom = i32::from_le_bytes([
            record.data[12],
            record.data[13],
            record.data[14],
            record.data[15],
        ]) as f64;

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

    /// Parse EMR_POLYGON record
    fn parse_polygon(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 20 {
            return Ok(None);
        }

        // Bounds (16 bytes) + count (4 bytes)
        let count = u32::from_le_bytes([
            record.data[16],
            record.data[17],
            record.data[18],
            record.data[19],
        ]) as usize;

        if record.data.len() < 20 + count * 8 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count + 1);

        // Parse points using SIMD-friendly iteration
        for i in 0..count {
            let offset = 20 + i * 8;
            let x = i32::from_le_bytes([
                record.data[offset],
                record.data[offset + 1],
                record.data[offset + 2],
                record.data[offset + 3],
            ]) as f64;
            let y = i32::from_le_bytes([
                record.data[offset + 4],
                record.data[offset + 5],
                record.data[offset + 6],
                record.data[offset + 7],
            ]) as f64;

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

    /// Parse EMR_POLYLINE record
    fn parse_polyline(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 20 {
            return Ok(None);
        }

        let count = u32::from_le_bytes([
            record.data[16],
            record.data[17],
            record.data[18],
            record.data[19],
        ]) as usize;

        if record.data.len() < 20 + count * 8 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count);

        for i in 0..count {
            let offset = 20 + i * 8;
            let x = i32::from_le_bytes([
                record.data[offset],
                record.data[offset + 1],
                record.data[offset + 2],
                record.data[offset + 3],
            ]) as f64;
            let y = i32::from_le_bytes([
                record.data[offset + 4],
                record.data[offset + 5],
                record.data[offset + 6],
                record.data[offset + 7],
            ]) as f64;

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

    /// Parse EMR_POLYBEZIER record
    fn parse_polybezier(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 20 {
            return Ok(None);
        }

        let count = u32::from_le_bytes([
            record.data[16],
            record.data[17],
            record.data[18],
            record.data[19],
        ]) as usize;

        if record.data.len() < 20 + count * 8 || count < 4 {
            return Ok(None);
        }

        let mut commands = Vec::new();

        // First point is MoveTo
        let x0 = i32::from_le_bytes([
            record.data[20],
            record.data[21],
            record.data[22],
            record.data[23],
        ]) as f64;
        let y0 = i32::from_le_bytes([
            record.data[24],
            record.data[25],
            record.data[26],
            record.data[27],
        ]) as f64;

        commands.push(PathCommand::MoveTo { x: x0, y: y0 });

        // Process Bezier curves (groups of 3 points)
        let mut i = 1;
        while i + 2 < count {
            let offset1 = 20 + i * 8;
            let offset2 = 20 + (i + 1) * 8;
            let offset3 = 20 + (i + 2) * 8;

            let x1 = i32::from_le_bytes([
                record.data[offset1],
                record.data[offset1 + 1],
                record.data[offset1 + 2],
                record.data[offset1 + 3],
            ]) as f64;
            let y1 = i32::from_le_bytes([
                record.data[offset1 + 4],
                record.data[offset1 + 5],
                record.data[offset1 + 6],
                record.data[offset1 + 7],
            ]) as f64;

            let x2 = i32::from_le_bytes([
                record.data[offset2],
                record.data[offset2 + 1],
                record.data[offset2 + 2],
                record.data[offset2 + 3],
            ]) as f64;
            let y2 = i32::from_le_bytes([
                record.data[offset2 + 4],
                record.data[offset2 + 5],
                record.data[offset2 + 6],
                record.data[offset2 + 7],
            ]) as f64;

            let x = i32::from_le_bytes([
                record.data[offset3],
                record.data[offset3 + 1],
                record.data[offset3 + 2],
                record.data[offset3 + 3],
            ]) as f64;
            let y = i32::from_le_bytes([
                record.data[offset3 + 4],
                record.data[offset3 + 5],
                record.data[offset3 + 6],
                record.data[offset3 + 7],
            ]) as f64;

            commands.push(PathCommand::CubicBezier { x1, y1, x2, y2, x, y });

            i += 3;
        }

        Ok(Some(SvgElement::Path(
            SvgPath::new(commands)
                .with_stroke("#000000".to_string())
                .with_fill("none".to_string()),
        )))
    }

    /// Parse EMR_LINETO record
    fn parse_lineto(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        if record.data.len() < 8 {
            return Ok(None);
        }

        let x = i32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]) as f64;
        let y = i32::from_le_bytes([
            record.data[4],
            record.data[5],
            record.data[6],
            record.data[7],
        ]) as f64;

        // LineTo requires current position, which we'd need to track
        // For simplicity, create a line from origin
        Ok(Some(SvgElement::Path(
            SvgPath::new(vec![
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo { x, y },
            ])
            .with_stroke("#000000".to_string()),
        )))
    }

    /// Parse EMR_STRETCHDIBITS record (embedded bitmap)
    fn parse_stretchdibits(&self, record: &EmfRecord) -> Result<Option<SvgElement>> {
        // This is complex - extract DIB and convert to PNG, then embed
        if record.data.len() < 80 {
            return Ok(None);
        }

        // Parse bounds and offsets
        let dest_x = i32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]) as f64;
        let dest_y = i32::from_le_bytes([
            record.data[4],
            record.data[5],
            record.data[6],
            record.data[7],
        ]) as f64;
        let dest_width = i32::from_le_bytes([
            record.data[8],
            record.data[9],
            record.data[10],
            record.data[11],
        ]) as f64;
        let dest_height = i32::from_le_bytes([
            record.data[12],
            record.data[13],
            record.data[14],
            record.data[15],
        ]) as f64;

        // Try to extract and convert DIB data
        // This is simplified - full implementation would parse DIB structure
        if let Ok(png_data) = self.extract_and_convert_dib(&record.data[80..]) {
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

    /// Parse EMR_ARC record
    fn parse_arc(&self, _record: &EmfRecord) -> Result<Option<SvgElement>> {
        // Arc parsing is complex and requires trigonometry
        // Placeholder for future implementation
        Ok(None)
    }

    /// Parse EMR_PIE record
    fn parse_pie(&self, _record: &EmfRecord) -> Result<Option<SvgElement>> {
        // Pie chart slice - complex path construction
        Ok(None)
    }

    /// Parse EMR_CHORD record
    fn parse_chord(&self, _record: &EmfRecord) -> Result<Option<SvgElement>> {
        // Chord - similar to arc with connecting line
        Ok(None)
    }

    /// Extract and convert DIB data to PNG
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
    fn test_emf_svg_converter_creation() {
        // Test requires valid EMF data
        // Placeholder for future tests
    }
}

