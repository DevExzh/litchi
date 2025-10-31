//! SVG element rendering from WMF records
//!
//! Processes WMF records sequentially, updating graphics state and generating
//! minimal SVG elements for drawing operations. Matches libwmf behavior while
//! producing compact, optimized SVG output.

use super::super::constants::record;
use super::super::parser::WmfRecord;
use super::state::{Brush, Font, GdiObject, GraphicsState, Pen};
use super::style::{fill_attr, map_font_family, stroke_attrs};
use super::transform::CoordinateTransform;
use crate::common::binary::{read_i16_le, read_u16_le};
use crate::images::svg_utils::{write_color_hex, write_num};

/// Type of arc rendering
#[derive(Debug, Clone, Copy)]
enum ArcType {
    Open,  // Arc only
    Pie,   // Arc + lines to center
    Chord, // Arc + closing line
}

/// SVG renderer that processes WMF records
pub struct SvgRenderer {
    transform: CoordinateTransform,
    state: GraphicsState,
}

impl SvgRenderer {
    pub fn new(transform: CoordinateTransform) -> Self {
        Self {
            transform,
            state: GraphicsState::new(),
        }
    }

    /// Render a WMF record to SVG element (or None if no output)
    pub fn render_record(&mut self, rec: &WmfRecord) -> Option<String> {
        // Update state first
        self.update_state(rec);

        // Then render if this record produces output
        match rec.function {
            record::RECTANGLE => self.render_rectangle(rec),
            record::ELLIPSE => self.render_ellipse(rec),
            record::POLYGON => self.render_polygon(rec),
            record::POLYLINE => self.render_polyline(rec),
            record::LINE_TO => self.render_line_to(rec),
            record::TEXT_OUT => self.render_text_out(rec),
            record::EXT_TEXT_OUT => self.render_ext_text_out(rec),
            record::ARC => self.render_arc(rec),
            record::PIE => self.render_pie(rec),
            record::CHORD => self.render_chord(rec),
            record::ROUND_RECT => self.render_round_rect(rec),
            record::POLYPOLYGON => self.render_polypolygon(rec),
            _ => None,
        }
    }

    /// Update graphics state from record
    fn update_state(&mut self, rec: &WmfRecord) {
        match rec.function {
            record::MOVE_TO | record::SET_PIXEL_V if rec.params.len() >= 4 => {
                self.state.position = (
                    read_i16_le(&rec.params, 2).unwrap_or(0),
                    read_i16_le(&rec.params, 0).unwrap_or(0),
                );
            },
            record::SET_TEXT_COLOR if rec.params.len() >= 4 => {
                self.state.text_color = u32::from_le_bytes([
                    rec.params[0],
                    rec.params[1],
                    rec.params[2],
                    rec.params[3],
                ]);
            },
            record::SET_BK_COLOR if rec.params.len() >= 4 => {
                self.state.bk_color = u32::from_le_bytes([
                    rec.params[0],
                    rec.params[1],
                    rec.params[2],
                    rec.params[3],
                ]);
            },
            record::CREATE_PEN_INDIRECT if rec.params.len() >= 10 => {
                let pen = Pen {
                    style: read_u16_le(&rec.params, 0).unwrap_or(0),
                    width: read_u16_le(&rec.params, 2).unwrap_or(1),
                    color: u32::from_le_bytes([
                        rec.params[6],
                        rec.params[7],
                        rec.params[8],
                        rec.params[9],
                    ]),
                };
                self.state.objects.push(Some(GdiObject::Pen(pen)));
            },
            record::CREATE_BRUSH_INDIRECT if rec.params.len() >= 8 => {
                let brush = Brush {
                    style: read_u16_le(&rec.params, 0).unwrap_or(1),
                    color: u32::from_le_bytes([
                        rec.params[2],
                        rec.params[3],
                        rec.params[4],
                        rec.params[5],
                    ]),
                };
                self.state.objects.push(Some(GdiObject::Brush(brush)));
            },
            record::CREATE_FONT_INDIRECT if rec.params.len() >= 18 => {
                let height = read_i16_le(&rec.params, 0).unwrap_or(12);
                let escapement = read_u16_le(&rec.params, 4).unwrap_or(0);
                let weight = read_u16_le(&rec.params, 8).unwrap_or(400);
                let attrs = read_u16_le(&rec.params, 10).unwrap_or(0);
                let italic = (attrs & 0xFF) != 0;
                let underline = ((attrs >> 8) & 0xFF) != 0;
                let strike_out = rec.params.get(12).map(|&b| b != 0).unwrap_or(false);

                let name_bytes = &rec.params[18..];
                let name_end = name_bytes
                    .iter()
                    .position(|&b| b == 0)
                    .unwrap_or(name_bytes.len());
                let name = if name_end > 0 {
                    String::from_utf8_lossy(&name_bytes[..name_end]).to_string()
                } else {
                    "serif".to_string()
                };

                let font = Font {
                    height,
                    escapement,
                    weight,
                    italic,
                    underline,
                    strike_out,
                    name: map_font_family(&name).to_string(),
                };
                self.state.objects.push(Some(GdiObject::Font(font)));
            },
            record::SELECT_OBJECT if rec.params.len() >= 2 => {
                let idx = u16::from_le_bytes([rec.params[0], rec.params[1]]) as usize;
                if let Some(Some(obj)) = self.state.objects.get(idx) {
                    match obj {
                        GdiObject::Pen(p) => self.state.pen = *p,
                        GdiObject::Brush(b) => self.state.brush = *b,
                        GdiObject::Font(f) => self.state.font = f.clone(),
                    }
                }
            },
            record::DELETE_OBJECT if rec.params.len() >= 2 => {
                let idx = u16::from_le_bytes([rec.params[0], rec.params[1]]) as usize;
                if idx < self.state.objects.len() {
                    self.state.objects[idx] = None;
                }
            },
            record::SET_POLY_FILL_MODE if rec.params.len() >= 2 => {
                self.state.poly_fill_mode = read_u16_le(&rec.params, 0).unwrap_or(1);
            },
            _ => {},
        }
    }

    fn render_rectangle(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 8 {
            return None;
        }

        let bottom = read_i16_le(&rec.params, 0).unwrap_or(0);
        let right = read_i16_le(&rec.params, 2).unwrap_or(0);
        let top = read_i16_le(&rec.params, 4).unwrap_or(0);
        let left = read_i16_le(&rec.params, 6).unwrap_or(0);

        let (x, y) = self.transform.point(left, top);
        let (x2, y2) = self.transform.point(right, bottom);

        let mut s = String::with_capacity(128);
        s.push_str(r#"<rect x=""#);
        write_num(&mut s, x);
        s.push_str(r#"" y=""#);
        write_num(&mut s, y);
        s.push_str(r#"" width=""#);
        write_num(&mut s, (x2 - x).abs());
        s.push_str(r#"" height=""#);
        write_num(&mut s, (y2 - y).abs());
        s.push('"');

        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_ellipse(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 8 {
            return None;
        }

        let bottom = read_i16_le(&rec.params, 0).unwrap_or(0);
        let right = read_i16_le(&rec.params, 2).unwrap_or(0);
        let top = read_i16_le(&rec.params, 4).unwrap_or(0);
        let left = read_i16_le(&rec.params, 6).unwrap_or(0);

        let (x1, y1) = self.transform.point(left, top);
        let (x2, y2) = self.transform.point(right, bottom);

        let cx = (x1 + x2) / 2.0;
        let cy = (y1 + y2) / 2.0;
        let rx = (x2 - x1).abs() / 2.0;
        let ry = (y2 - y1).abs() / 2.0;

        let mut s = String::with_capacity(128);
        s.push_str(r#"<ellipse cx=""#);
        write_num(&mut s, cx);
        s.push_str(r#"" cy=""#);
        write_num(&mut s, cy);
        s.push_str(r#"" rx=""#);
        write_num(&mut s, rx);
        s.push_str(r#"" ry=""#);
        write_num(&mut s, ry);
        s.push('"');

        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_polygon(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 2 {
            return None;
        }

        let count = read_i16_le(&rec.params, 0).unwrap_or(0) as usize;
        if count < 3 || rec.params.len() < 2 + count * 4 {
            return None;
        }

        let mut s = String::with_capacity(128 + count * 12);
        s.push_str(r#"<polygon points=""#);

        // Extract x and y coordinates separately for SIMD processing
        let mut xs = Vec::with_capacity(count);
        let mut ys = Vec::with_capacity(count);
        for i in 0..count {
            xs.push(read_i16_le(&rec.params, 2 + i * 4).unwrap_or(0));
            ys.push(read_i16_le(&rec.params, 4 + i * 4).unwrap_or(0));
        }

        // Use SIMD-accelerated batch transformation
        self.transform
            .transform_and_format_points(&xs, &ys, &mut s, ' ');

        s.push('"');
        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_polyline(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 2 {
            return None;
        }

        let count = read_i16_le(&rec.params, 0).unwrap_or(0) as usize;
        if count < 2 || rec.params.len() < 2 + count * 4 {
            return None;
        }

        let mut s = String::with_capacity(128 + count * 12);
        s.push_str(r#"<polyline points=""#);

        // Extract x and y coordinates separately for SIMD processing
        let mut xs = Vec::with_capacity(count);
        let mut ys = Vec::with_capacity(count);
        for i in 0..count {
            xs.push(read_i16_le(&rec.params, 2 + i * 4).unwrap_or(0));
            ys.push(read_i16_le(&rec.params, 4 + i * 4).unwrap_or(0));
        }

        // Use SIMD-accelerated batch transformation
        self.transform
            .transform_and_format_points(&xs, &ys, &mut s, ' ');

        s.push_str(r#"" fill="none""#);
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_line_to(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 4 {
            return None;
        }

        let y2 = read_i16_le(&rec.params, 0).unwrap_or(0);
        let x2 = read_i16_le(&rec.params, 2).unwrap_or(0);
        let (x1, y1) = self
            .transform
            .point(self.state.position.0, self.state.position.1);
        let (x2, y2) = self.transform.point(x2, y2);

        let mut s = String::with_capacity(128);
        s.push_str(r#"<line x1=""#);
        write_num(&mut s, x1);
        s.push_str(r#"" y1=""#);
        write_num(&mut s, y1);
        s.push_str(r#"" x2=""#);
        write_num(&mut s, x2);
        s.push_str(r#"" y2=""#);
        write_num(&mut s, y2);
        s.push('"');
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_text_out(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 6 {
            return None;
        }

        let len = read_u16_le(&rec.params, 0).unwrap_or(0) as usize;
        if len == 0 {
            return None;
        }

        let text_end = (2 + len).min(rec.params.len());
        let text = String::from_utf8_lossy(&rec.params[2..text_end]);

        let off = 2 + len.div_ceil(2) * 2;
        if rec.params.len() < off + 4 {
            return None;
        }

        let y = read_i16_le(&rec.params, off).unwrap_or(0);
        let x = read_i16_le(&rec.params, off + 2).unwrap_or(0);

        self.render_text(&text, x, y)
    }

    fn render_ext_text_out(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 8 {
            return None;
        }

        let y = read_i16_le(&rec.params, 0).unwrap_or(0);
        let x = read_i16_le(&rec.params, 2).unwrap_or(0);
        let len = read_u16_le(&rec.params, 4).unwrap_or(0) as usize;
        let opts = read_u16_le(&rec.params, 6).unwrap_or(0);

        if len == 0 {
            return None;
        }

        let off = if (opts & 0x06) != 0 { 16 } else { 8 };
        if rec.params.len() < off + len {
            return None;
        }

        let text = String::from_utf8_lossy(&rec.params[off..off + len]);

        self.render_text(&text, x, y)
    }

    fn render_text(&self, text: &str, x: i16, y: i16) -> Option<String> {
        let (tx, ty) = self.transform.point(x, y);
        let font_size = self.transform.height(self.state.font.height.abs() as f64);

        let mut s = String::with_capacity(128 + text.len());
        s.push_str(r#"<text x=""#);
        write_num(&mut s, tx);
        s.push_str(r#"" y=""#);
        write_num(&mut s, ty);
        s.push_str(r#"" font-size=""#);
        write_num(&mut s, font_size);
        s.push_str("\" fill=\"");
        write_color_hex(&mut s, self.state.text_color);
        s.push('"');

        // Non-default font attributes
        if self.state.font.name != "serif" {
            s.push_str(r#" font-family=""#);
            s.push_str(&self.state.font.name);
            s.push('"');
        }
        if self.state.font.italic {
            s.push_str(r#" font-style="italic""#);
        }
        if self.state.font.weight >= 700 {
            s.push_str(r#" font-weight="bold""#);
        }
        if self.state.font.underline {
            s.push_str(r#" text-decoration="underline""#);
        } else if self.state.font.strike_out {
            s.push_str(r#" text-decoration="line-through""#);
        }

        // Rotation transform if escapement is non-zero
        if self.state.font.escapement != 0 {
            let angle = -(self.state.font.escapement as f64 / 10.0);
            s.push_str(r#" transform="rotate("#);
            write_num(&mut s, angle);
            s.push(' ');
            write_num(&mut s, tx);
            s.push(' ');
            write_num(&mut s, ty);
            s.push_str(r#")""#);
        }

        s.push('>');

        // XML escape
        for c in text.chars() {
            match c {
                '<' => s.push_str("&lt;"),
                '>' => s.push_str("&gt;"),
                '&' => s.push_str("&amp;"),
                '"' => s.push_str("&quot;"),
                _ => s.push(c),
            }
        }

        s.push_str("</text>");
        Some(s)
    }

    fn render_arc(&self, rec: &WmfRecord) -> Option<String> {
        self.render_arc_common(rec, ArcType::Open)
    }

    fn render_pie(&self, rec: &WmfRecord) -> Option<String> {
        self.render_arc_common(rec, ArcType::Pie)
    }

    fn render_chord(&self, rec: &WmfRecord) -> Option<String> {
        self.render_arc_common(rec, ArcType::Chord)
    }

    fn render_arc_common(&self, rec: &WmfRecord, arc_type: ArcType) -> Option<String> {
        if rec.params.len() < 16 {
            return None;
        }

        let y_end = read_i16_le(&rec.params, 0).unwrap_or(0);
        let x_end = read_i16_le(&rec.params, 2).unwrap_or(0);
        let y_start = read_i16_le(&rec.params, 4).unwrap_or(0);
        let x_start = read_i16_le(&rec.params, 6).unwrap_or(0);
        let bottom = read_i16_le(&rec.params, 8).unwrap_or(0);
        let right = read_i16_le(&rec.params, 10).unwrap_or(0);
        let top = read_i16_le(&rec.params, 12).unwrap_or(0);
        let left = read_i16_le(&rec.params, 14).unwrap_or(0);

        // If start and end are the same, render as ellipse
        if x_start == x_end && y_start == y_end {
            return self.render_ellipse_at(left, top, right, bottom);
        }

        let (tl_x, tl_y) = self.transform.point(left, top);
        let (br_x, br_y) = self.transform.point(right, bottom);
        let (start_x, start_y) = self.transform.point(x_start, y_start);
        let (end_x, end_y) = self.transform.point(x_end, y_end);

        let rx = (br_x - tl_x).abs() / 2.0;
        let ry = (br_y - tl_y).abs() / 2.0;
        let cx = (tl_x + br_x) / 2.0;
        let cy = (tl_y + br_y) / 2.0;

        let mut s = String::with_capacity(128);
        s.push_str(r#"<path d="M"#);

        // Start at the start point
        write_num(&mut s, start_x);
        s.push(',');
        write_num(&mut s, start_y);

        // Arc to end point (large-arc-flag=0, sweep-flag=1 for clockwise)
        s.push('A');
        write_num(&mut s, rx);
        s.push(',');
        write_num(&mut s, ry);
        s.push_str(" 0 0,1 ");
        write_num(&mut s, end_x);
        s.push(',');
        write_num(&mut s, end_y);

        match arc_type {
            ArcType::Pie => {
                // Line to center, then close
                s.push('L');
                write_num(&mut s, cx);
                s.push(',');
                write_num(&mut s, cy);
                s.push('Z');
            },
            ArcType::Chord => {
                // Close path (straight line to start)
                s.push('Z');
            },
            ArcType::Open => {
                // No close for open arc
            },
        }

        s.push('"');

        if matches!(arc_type, ArcType::Pie | ArcType::Chord) {
            if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
                s.push_str(&fill);
            }
        } else {
            s.push_str(r#" fill="none""#);
        }

        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_ellipse_at(&self, left: i16, top: i16, right: i16, bottom: i16) -> Option<String> {
        let (x1, y1) = self.transform.point(left, top);
        let (x2, y2) = self.transform.point(right, bottom);

        let cx = (x1 + x2) / 2.0;
        let cy = (y1 + y2) / 2.0;
        let rx = (x2 - x1).abs() / 2.0;
        let ry = (y2 - y1).abs() / 2.0;

        let mut s = String::with_capacity(128);
        s.push_str(r#"<ellipse cx=""#);
        write_num(&mut s, cx);
        s.push_str(r#"" cy=""#);
        write_num(&mut s, cy);
        s.push_str(r#"" rx=""#);
        write_num(&mut s, rx);
        s.push_str(r#"" ry=""#);
        write_num(&mut s, ry);
        s.push('"');

        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_polypolygon(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 2 {
            return None;
        }

        let num_polys = read_u16_le(&rec.params, 0).unwrap_or(0) as usize;
        if num_polys == 0 {
            return None;
        }

        // Read polygon counts
        let mut offset = 2;
        let mut poly_counts = Vec::with_capacity(num_polys);
        for _ in 0..num_polys {
            if offset + 2 > rec.params.len() {
                return None;
            }
            let count = read_u16_le(&rec.params, offset).unwrap_or(0) as usize;
            poly_counts.push(count);
            offset += 2;
        }

        // Build path data
        let mut path_data = String::with_capacity(256);
        for count in poly_counts {
            if count < 3 {
                continue;
            }

            if offset + count * 4 > rec.params.len() {
                break;
            }

            // Extract all points for this polygon for SIMD processing
            let mut xs = Vec::with_capacity(count);
            let mut ys = Vec::with_capacity(count);
            for i in 0..count {
                xs.push(read_i16_le(&rec.params, offset + i * 4).unwrap_or(0));
                ys.push(read_i16_le(&rec.params, offset + i * 4 + 2).unwrap_or(0));
            }
            offset += count * 4;

            // Transform all points at once using SIMD
            let mut out_x = vec![0.0; count];
            let mut out_y = vec![0.0; count];
            self.transform
                .transform_points_batch(&xs, &ys, &mut out_x, &mut out_y);

            // First point - MoveTo
            path_data.push('M');
            write_num(&mut path_data, out_x[0]);
            path_data.push(',');
            write_num(&mut path_data, out_y[0]);
            path_data.push('L');

            // Remaining points - LineTo
            for i in 1..count {
                if i > 1 {
                    path_data.push(' ');
                }
                write_num(&mut path_data, out_x[i]);
                path_data.push(',');
                write_num(&mut path_data, out_y[i]);
            }

            path_data.push('Z');
        }

        if path_data.is_empty() {
            return None;
        }

        let mut s = format!(r#"<path d="{}""#, path_data);
        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }

    fn render_round_rect(&self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 12 {
            return None;
        }

        let h = read_i16_le(&rec.params, 0).unwrap_or(0);
        let w = read_i16_le(&rec.params, 2).unwrap_or(0);
        let bottom = read_i16_le(&rec.params, 4).unwrap_or(0);
        let right = read_i16_le(&rec.params, 6).unwrap_or(0);
        let top = read_i16_le(&rec.params, 8).unwrap_or(0);
        let left = read_i16_le(&rec.params, 10).unwrap_or(0);

        let (x, y) = self.transform.point(left, top);
        let (x2, y2) = self.transform.point(right, bottom);
        let rx = self.transform.width(w as f64 / 2.0);
        let ry = self.transform.height(h as f64 / 2.0);

        let mut s = String::with_capacity(128);
        s.push_str(r#"<rect x=""#);
        write_num(&mut s, x);
        s.push_str(r#"" y=""#);
        write_num(&mut s, y);
        s.push_str(r#"" width=""#);
        write_num(&mut s, (x2 - x).abs());
        s.push_str(r#"" height=""#);
        write_num(&mut s, (y2 - y).abs());
        s.push_str(r#"" rx=""#);
        write_num(&mut s, rx);
        s.push_str(r#"" ry=""#);
        write_num(&mut s, ry);
        s.push('"');

        if let Some(fill) = fill_attr(&self.state.brush, self.state.poly_fill_mode) {
            s.push_str(&fill);
        }
        s.push_str(&stroke_attrs(&self.state.pen, &self.transform));
        s.push_str("/>");

        Some(s)
    }
}
