/// Comprehensive EMF to SVG Converter with In-Place Optimizations
///
/// Production-ready converter supporting all major EMF record types
///
/// Optimizations applied during conversion:
/// - Merge consecutive lines into polylines
/// - Group elements with same styles
/// - Eliminate redundant attributes
/// - Optimize number precision (2 decimal places)
/// - Reuse styles via grouping
use super::{
    buffer::ElementBuffer,
    path::PathBuilder,
    state::{DeviceContext, RenderState},
};
use crate::common::error::Result;
use crate::common::xml::escape::escape_xml;
use crate::images::emf::parser::EmfParser;
use crate::images::emf::records::*;
use crate::images::svg_utils::write_num;
use std::fmt::Write;
use zerocopy::FromBytes;

/// EMF to SVG Converter with in-place optimization
pub struct EmfSvgConverter<'a> {
    parser: &'a EmfParser,
}

impl<'a> EmfSvgConverter<'a> {
    /// Create new converter
    pub fn new(parser: &'a EmfParser) -> Self {
        Self { parser }
    }

    /// Convert EMF to SVG with in-place optimizations
    pub fn convert(&self) -> Result<String> {
        let mut state = RenderState::new();
        let mut buffer = ElementBuffer::new();

        // Process all records with buffering for optimization
        for record in &self.parser.records {
            if let Some(elements) = self.process_record(record, &mut state)? {
                for element in elements {
                    buffer.add_element(element, &state.dc);
                }
            }
        }

        // Flush any pending buffered elements
        buffer.flush();

        // Build final SVG
        self.build_svg(&buffer.elements, &state)
    }

    /// Process a single EMF record
    fn process_record(
        &self,
        record: &super::super::parser::EmfRecord,
        state: &mut RenderState,
    ) -> Result<Option<Vec<String>>> {
        let record_type = EmrType::from_u32(record.record_type);

        match record_type {
            // State management
            Some(EmrType::SaveDc) => {
                state.push_dc();
                Ok(None)
            },
            Some(EmrType::RestoreDc) => {
                if record.data.len() >= 4 {
                    let index = i32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]);
                    state.pop_dc(index);
                }
                Ok(None)
            },

            // Transform operations
            Some(EmrType::SetWorldTransform) => {
                if let Ok((xform, _)) = XForm::read_from_prefix(&record.data) {
                    state.dc.world_transform = xform;
                }
                Ok(None)
            },
            Some(EmrType::ModifyWorldTransform) => {
                if record.data.len() >= 28
                    && let Ok((xform, rest)) = XForm::read_from_prefix(&record.data)
                {
                    let mode = u32::from_le_bytes([rest[0], rest[1], rest[2], rest[3]]);
                    match mode {
                        2 => state.dc.world_transform = state.dc.world_transform.multiply(&xform), // Left multiply
                        3 => state.dc.world_transform = xform.multiply(&state.dc.world_transform), // Right multiply
                        _ => state.dc.world_transform = xform, // Set
                    }
                }
                Ok(None)
            },

            // Window/Viewport mapping
            Some(EmrType::SetWindowExtEx) => {
                if let Ok((extent, _)) = SizeL::read_from_prefix(&record.data) {
                    state.dc.window_ext = (extent.cx, extent.cy);
                }
                Ok(None)
            },
            Some(EmrType::SetWindowOrgEx) => {
                if let Ok((origin, _)) = PointL::read_from_prefix(&record.data) {
                    state.dc.window_org = (origin.x, origin.y);
                }
                Ok(None)
            },
            Some(EmrType::SetViewportExtEx) => {
                if let Ok((extent, _)) = SizeL::read_from_prefix(&record.data) {
                    state.dc.viewport_ext = (extent.cx, extent.cy);
                }
                Ok(None)
            },
            Some(EmrType::SetViewportOrgEx) => {
                if let Ok((origin, _)) = PointL::read_from_prefix(&record.data) {
                    state.dc.viewport_org = (origin.x, origin.y);
                }
                Ok(None)
            },

            // Colors and modes
            Some(EmrType::SetTextColor) | Some(EmrType::SetBkColor) => {
                if let Ok((color, _)) = ColorRef::read_from_prefix(&record.data) {
                    if record_type == Some(EmrType::SetTextColor) {
                        state.dc.text_color = color;
                    } else {
                        state.dc.bg_color = color;
                    }
                }
                Ok(None)
            },
            Some(EmrType::SetBkMode) => {
                if record.data.len() >= 4 {
                    state.dc.bg_mode = u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]);
                }
                Ok(None)
            },
            Some(EmrType::SetPolyFillMode) => {
                if record.data.len() >= 4 {
                    state.dc.poly_fill_mode = u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]);
                }
                Ok(None)
            },
            Some(EmrType::SetTextAlign) => {
                if record.data.len() >= 4 {
                    state.dc.text_align = u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]);
                }
                Ok(None)
            },

            // Path operations
            Some(EmrType::BeginPath) => {
                state.begin_path();
                Ok(None)
            },
            Some(EmrType::EndPath) => {
                state.end_path();
                Ok(None)
            },
            Some(EmrType::CloseFigure) => {
                if let Some(ref mut builder) = state.path_builder {
                    builder.close();
                }
                Ok(None)
            },
            Some(EmrType::MoveToEx) => {
                if let Ok((point, _)) = PointL::read_from_prefix(&record.data) {
                    let (x, y) = state.dc.transform_point(point.x as f64, point.y as f64);
                    state.dc.current_pos = (x, y);

                    if let Some(ref mut builder) = state.path_builder {
                        builder.move_to(x, y);
                    }
                }
                Ok(None)
            },
            Some(EmrType::LineTo) => {
                if let Ok((point, _)) = PointL::read_from_prefix(&record.data) {
                    let (x, y) = state.dc.transform_point(point.x as f64, point.y as f64);

                    if let Some(ref mut builder) = state.path_builder {
                        builder.line_to(x, y);
                    } else {
                        // Direct line rendering
                        return Ok(Some(vec![self.render_line(
                            state.dc.current_pos.0,
                            state.dc.current_pos.1,
                            x,
                            y,
                            &state.dc,
                        )]));
                    }

                    state.dc.current_pos = (x, y);
                }
                Ok(None)
            },
            Some(EmrType::FillPath)
            | Some(EmrType::StrokePath)
            | Some(EmrType::StrokeAndFillPath) => {
                if let Some(mut builder) = state.take_path() {
                    builder.optimize();
                    let path_str = builder.build();

                    let fill = if record_type == Some(EmrType::FillPath)
                        || record_type == Some(EmrType::StrokeAndFillPath)
                    {
                        state.dc.get_fill_attr()
                    } else {
                        "fill=\"none\"".to_string()
                    };

                    let stroke = if record_type == Some(EmrType::StrokePath)
                        || record_type == Some(EmrType::StrokeAndFillPath)
                    {
                        state.dc.get_stroke_attrs()
                    } else {
                        "stroke=\"none\"".to_string()
                    };

                    let mut svg = format!("<path d=\"{}\" {} {}", path_str, fill, stroke);
                    if let Some(fill_rule) = state.dc.get_fill_rule() {
                        write!(&mut svg, " {}", fill_rule).ok();
                    }
                    svg.push_str("/>");

                    Ok(Some(vec![svg]))
                } else {
                    Ok(None)
                }
            },

            // Shape rendering
            Some(EmrType::Rectangle) | Some(EmrType::Ellipse) => {
                if let Ok((rect, _)) = RectL::read_from_prefix(&record.data) {
                    let (x1, y1) = state.dc.transform_point(rect.left as f64, rect.top as f64);
                    let (x2, y2) = state
                        .dc
                        .transform_point(rect.right as f64, rect.bottom as f64);

                    let svg = if record_type == Some(EmrType::Rectangle) {
                        self.render_rectangle(x1, y1, x2 - x1, y2 - y1, &state.dc)
                    } else {
                        self.render_ellipse(
                            (x1 + x2) / 2.0,
                            (y1 + y2) / 2.0,
                            (x2 - x1) / 2.0,
                            (y2 - y1) / 2.0,
                            &state.dc,
                        )
                    };

                    Ok(Some(vec![svg]))
                } else {
                    Ok(None)
                }
            },
            Some(EmrType::RoundRect) => {
                if let Ok((hdr, _)) = EmrRoundRect::read_from_prefix(&record.data) {
                    let (x1, y1) = state
                        .dc
                        .transform_point(hdr.rect.left as f64, hdr.rect.top as f64);
                    let (x2, y2) = state
                        .dc
                        .transform_point(hdr.rect.right as f64, hdr.rect.bottom as f64);
                    let rx = hdr.corner.cx as f64 / 2.0;
                    let ry = hdr.corner.cy as f64 / 2.0;

                    Ok(Some(vec![self.render_rounded_rectangle(
                        (x1, y1, x2 - x1, y2 - y1),
                        (rx, ry),
                        &state.dc,
                    )]))
                } else {
                    Ok(None)
                }
            },

            // Polygon rendering
            Some(EmrType::Polygon) | Some(EmrType::Polyline) | Some(EmrType::PolyBezier) => {
                self.render_polygon(record, state, record_type)
            },
            Some(EmrType::Polygon16) | Some(EmrType::Polyline16) | Some(EmrType::PolyBezier16) => {
                self.render_polygon16(record, state, record_type)
            },

            // Object creation (store in object table)
            Some(EmrType::CreatePen) => {
                if let Ok((pen, _)) = EmrCreatePen::read_from_prefix(&record.data) {
                    state.dc.pen =
                        super::state::Pen::from_create_pen(pen.pen_style, pen.width, pen.color);
                }
                Ok(None)
            },
            Some(EmrType::CreateBrushIndirect) => {
                if let Ok((brush, _)) = EmrCreateBrushIndirect::read_from_prefix(&record.data) {
                    state.dc.brush.style = brush.brush_style;
                    state.dc.brush.color = brush.color;
                    state.dc.brush.hatch = Some(brush.brush_hatch);
                }
                Ok(None)
            },

            // Text rendering
            Some(EmrType::ExtTextOutA) | Some(EmrType::ExtTextOutW) => {
                self.render_text(record, state, record_type == Some(EmrType::ExtTextOutW))
            },

            // Unimplemented records - log but don't error
            _ => Ok(None),
        }
    }

    /// Render a line (optimized format - minimal whitespace, no trailing zeros)
    fn render_line(&self, x1: f64, y1: f64, x2: f64, y2: f64, dc: &DeviceContext) -> String {
        let mut s = String::with_capacity(128);
        s.push_str("<line x1=\"");
        write_num(&mut s, x1);
        s.push_str("\" y1=\"");
        write_num(&mut s, y1);
        s.push_str("\" x2=\"");
        write_num(&mut s, x2);
        s.push_str("\" y2=\"");
        write_num(&mut s, y2);
        s.push_str("\" ");
        s.push_str(&dc.get_stroke_attrs());
        s.push_str("/>");
        s
    }

    /// Render a rectangle (optimized)
    fn render_rectangle(
        &self,
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        dc: &DeviceContext,
    ) -> String {
        let mut s = String::with_capacity(128);
        s.push_str("<rect x=\"");
        write_num(&mut s, x);
        s.push_str("\" y=\"");
        write_num(&mut s, y);
        s.push_str("\" width=\"");
        write_num(&mut s, width);
        s.push_str("\" height=\"");
        write_num(&mut s, height);
        s.push_str("\" ");
        s.push_str(&dc.get_fill_attr());
        s.push(' ');
        s.push_str(&dc.get_stroke_attrs());
        s.push_str("/>");
        s
    }

    /// Render a rounded rectangle (optimized)
    fn render_rounded_rectangle(
        &self,
        rect: (f64, f64, f64, f64), // (x, y, width, height)
        corners: (f64, f64),        // (rx, ry)
        dc: &DeviceContext,
    ) -> String {
        let (x, y, width, height) = rect;
        let (rx, ry) = corners;
        let mut s = String::with_capacity(128);
        s.push_str("<rect x=\"");
        write_num(&mut s, x);
        s.push_str("\" y=\"");
        write_num(&mut s, y);
        s.push_str("\" width=\"");
        write_num(&mut s, width);
        s.push_str("\" height=\"");
        write_num(&mut s, height);
        s.push_str("\" rx=\"");
        write_num(&mut s, rx);
        s.push_str("\" ry=\"");
        write_num(&mut s, ry);
        s.push_str("\" ");
        s.push_str(&dc.get_fill_attr());
        s.push(' ');
        s.push_str(&dc.get_stroke_attrs());
        s.push_str("/>");
        s
    }

    /// Render an ellipse (optimized)
    fn render_ellipse(&self, cx: f64, cy: f64, rx: f64, ry: f64, dc: &DeviceContext) -> String {
        let mut s = String::with_capacity(128);
        s.push_str("<ellipse cx=\"");
        write_num(&mut s, cx);
        s.push_str("\" cy=\"");
        write_num(&mut s, cy);
        s.push_str("\" rx=\"");
        write_num(&mut s, rx);
        s.push_str("\" ry=\"");
        write_num(&mut s, ry);
        s.push_str("\" ");
        s.push_str(&dc.get_fill_attr());
        s.push(' ');
        s.push_str(&dc.get_stroke_attrs());
        s.push_str("/>");
        s
    }

    /// Render polygon (32-bit coordinates)
    fn render_polygon(
        &self,
        record: &super::super::parser::EmfRecord,
        state: &RenderState,
        record_type: Option<EmrType>,
    ) -> Result<Option<Vec<String>>> {
        if let Ok((hdr, rest)) = EmrPolyHeader::read_from_prefix(&record.data) {
            if let Some(poly_data) = PolygonData::from_poly32(rest, 0, hdr.count as usize) {
                let mut builder = PathBuilder::new();

                for (i, (x, y)) in poly_data.iter_points().enumerate() {
                    let (px, py) = state.dc.transform_point(x as f64, y as f64);
                    if i == 0 {
                        builder.move_to(px, py);
                    } else {
                        builder.line_to(px, py);
                    }
                }

                if record_type == Some(EmrType::Polygon) {
                    builder.close();
                }

                builder.optimize();
                let path_str = builder.build();

                let is_filled = record_type == Some(EmrType::Polygon);
                let fill = if is_filled {
                    state.dc.get_fill_attr()
                } else {
                    "fill=\"none\"".to_string()
                };

                Ok(Some(vec![format!(
                    "<path d=\"{}\" {} {}/>",
                    path_str,
                    fill,
                    state.dc.get_stroke_attrs()
                )]))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }

    /// Render polygon (16-bit coordinates)
    fn render_polygon16(
        &self,
        record: &super::super::parser::EmfRecord,
        state: &RenderState,
        record_type: Option<EmrType>,
    ) -> Result<Option<Vec<String>>> {
        if let Ok((hdr, rest)) = EmrPoly16Header::read_from_prefix(&record.data) {
            if let Some(poly_data) = PolygonData::from_poly16(rest, 0, hdr.count as usize) {
                let mut builder = PathBuilder::new();

                for (i, (x, y)) in poly_data.iter_points().enumerate() {
                    let (px, py) = state.dc.transform_point(x as f64, y as f64);
                    if i == 0 {
                        builder.move_to(px, py);
                    } else {
                        builder.line_to(px, py);
                    }
                }

                if record_type == Some(EmrType::Polygon16) {
                    builder.close();
                }

                builder.optimize();
                let path_str = builder.build();

                let is_filled = record_type == Some(EmrType::Polygon16);
                let fill = if is_filled {
                    state.dc.get_fill_attr()
                } else {
                    "fill=\"none\"".to_string()
                };

                Ok(Some(vec![format!(
                    "<path d=\"{}\" {} {}/>",
                    path_str,
                    fill,
                    state.dc.get_stroke_attrs()
                )]))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }

    /// Render text
    fn render_text(
        &self,
        record: &super::super::parser::EmfRecord,
        state: &RenderState,
        is_unicode: bool,
    ) -> Result<Option<Vec<String>>> {
        if let Ok((hdr, _)) = EmrExtTextOutHeader::read_from_prefix(&record.data) {
            let (x, y) = state
                .dc
                .transform_point(hdr.text.reference.x as f64, hdr.text.reference.y as f64);

            // The off_string is relative to the start of the EMF record (including type+size),
            // but record.data starts AFTER type+size (8 bytes). So we need to subtract 8.
            let string_offset = if hdr.text.off_string >= 8 {
                (hdr.text.off_string - 8) as usize
            } else {
                hdr.text.off_string as usize
            };

            // Extract text string
            let text = if is_unicode {
                // Unicode (UTF-16LE)
                self.extract_unicode_string(
                    &record.data,
                    string_offset,
                    hdr.text.num_chars as usize,
                )
            } else {
                // ANSI
                self.extract_ansi_string(&record.data, string_offset, hdr.text.num_chars as usize)
            };

            if !text.is_empty() {
                let mut svg = String::with_capacity(128 + text.len());
                svg.push_str("<text x=\"");
                write_num(&mut svg, x);
                svg.push_str("\" y=\"");
                write_num(&mut svg, y);
                svg.push_str("\" fill=\"");
                svg.push_str(&state.dc.text_color.to_svg_color());
                svg.push_str("\" ");
                svg.push_str(&state.dc.font.to_svg_attrs());
                svg.push('>');
                svg.push_str(&escape_xml(&text));
                svg.push_str("</text>");
                Ok(Some(vec![svg]))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }

    /// Extract Unicode string
    fn extract_unicode_string(&self, data: &[u8], offset: usize, count: usize) -> String {
        let byte_count = count * 2;
        if offset + byte_count > data.len() {
            return String::new();
        }

        let mut chars = Vec::with_capacity(count);
        for i in 0..count {
            let idx = offset + i * 2;
            if idx + 2 <= data.len() {
                let ch = u16::from_le_bytes([data[idx], data[idx + 1]]);
                if ch != 0 {
                    chars.push(ch);
                }
            }
        }

        String::from_utf16_lossy(&chars)
    }

    /// Extract ANSI string
    fn extract_ansi_string(&self, data: &[u8], offset: usize, count: usize) -> String {
        if offset + count > data.len() {
            return String::new();
        }

        String::from_utf8_lossy(&data[offset..offset + count])
            .trim_end_matches('\0')
            .to_string()
    }

    /// Build final SVG document
    fn build_svg(&self, elements: &[String], _state: &RenderState) -> Result<String> {
        let header = &self.parser.header;
        let width = header.width();
        let height = header.height();

        let mut svg = format!(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"{} {} {} {}\">",
            width, height, header.bounds.0, header.bounds.1, width, height
        );

        for element in elements {
            svg.push_str(element);
        }

        svg.push_str("</svg>");

        Ok(svg)
    }
}
