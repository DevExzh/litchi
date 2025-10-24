// WMF to SVG converter with comprehensive object management
//
// Converts Windows Metafile vector graphics to SVG while maintaining full graphics state

use super::parser::{WmfParser, WmfRecord};
use crate::common::binary::{read_i16_le, read_u16_le};
use crate::common::error::{Error, Result};
use crate::images::svg::*;

/// WMF pen styles
#[derive(Debug, Clone, Copy)]
#[allow(dead_code)]
struct PenStyle {
    /// Line style bits [0-3]: PS_SOLID, PS_DASH, PS_DOT, etc.
    style: u16,
    /// Line width in logical units
    width: u16,
    /// Pen color (COLORREF)
    color: u32,
}

impl Default for PenStyle {
    fn default() -> Self {
        Self {
            style: 0, // PS_SOLID
            width: 1,
            color: 0x000000, // Black
        }
    }
}

impl PenStyle {
    /// Get SVG stroke-dasharray based on pen style
    fn dasharray(&self) -> Option<String> {
        let width = self.width.max(1) as f64;
        match self.style & 0x000F {
            1 => Some(format!("{} {}", width * 3.0, width * 3.0)), // PS_DASH
            2 => Some(format!("{} {}", width, width * 2.0)),       // PS_DOT
            3 => Some(format!("{} {} {} {}", width * 3.0, width, width, width)), // PS_DASHDOT
            4 => Some(format!(
                "{} {} {} {} {} {}",
                width * 3.0,
                width,
                width,
                width,
                width,
                width
            )), // PS_DASHDOTDOT
            5 => None,                                             // PS_NULL - no stroke
            _ => None,                                             // PS_SOLID and others
        }
    }

    /// Get SVG line cap style
    fn linecap(&self) -> &str {
        match (self.style >> 8) & 0x000F {
            0x01 => "square", // PS_ENDCAP_SQUARE
            0x02 => "butt",   // PS_ENDCAP_FLAT
            _ => "round",     // PS_ENDCAP_ROUND (default)
        }
    }

    /// Get SVG line join style
    fn linejoin(&self) -> &str {
        match (self.style >> 12) & 0x000F {
            0x01 => "bevel", // PS_JOIN_BEVEL
            0x02 => "round", // PS_JOIN_ROUND
            _ => "miter",    // PS_JOIN_MITER (default)
        }
    }

    /// Check if pen should not draw (PS_NULL)
    fn is_null(&self) -> bool {
        (self.style & 0x000F) == 5
    }
}

/// WMF brush styles
#[derive(Debug, Clone, Copy)]
#[allow(dead_code)]
struct BrushStyle {
    /// Brush style: BS_SOLID, BS_NULL, BS_HATCHED, etc.
    style: u16,
    /// Brush color (COLORREF)
    color: u32,
    /// Hatch pattern (if BS_HATCHED)
    hatch: u16,
}

impl Default for BrushStyle {
    fn default() -> Self {
        Self {
            style: 0,        // BS_SOLID
            color: 0xFFFFFF, // White
            hatch: 0,
        }
    }
}

impl BrushStyle {
    /// Check if brush should not fill (BS_NULL)
    fn is_null(&self) -> bool {
        self.style == 1 // BS_NULL
    }
}

/// WMF font object
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct FontStyle {
    /// Font height in logical units
    height: i16,
    /// Font width
    width: u16,
    /// Text rotation in tenths of degrees
    escapement: u16,
    /// Font weight (400 = normal, 700 = bold)
    weight: u16,
    /// Italic flag
    italic: bool,
    /// Underline flag
    underline: bool,
    /// Strikeout flag
    strikeout: bool,
    /// Character set
    charset: u8,
    /// Font family name
    name: String,
}

impl Default for FontStyle {
    fn default() -> Self {
        Self {
            height: 12,
            width: 0,
            escapement: 0,
            weight: 400,
            italic: false,
            underline: false,
            strikeout: false,
            charset: 0,
            name: "Times New Roman".to_string(),
        }
    }
}

/// WMF graphics object (pen, brush, or font)
#[derive(Debug, Clone)]
enum WmfObject {
    Pen(PenStyle),
    Brush(BrushStyle),
    Font(FontStyle),
}

/// Graphics state for WMF rendering with full object management
#[derive(Debug, Clone)]
struct GraphicsState {
    /// Current drawing position
    current_pos: (i16, i16),

    /// Object table (indexed by handle)
    objects: Vec<Option<WmfObject>>,

    /// Currently selected pen
    current_pen: PenStyle,

    /// Currently selected brush
    current_brush: BrushStyle,

    /// Currently selected font
    current_font: FontStyle,

    /// Text color
    text_color: u32,

    /// Background color
    bg_color: u32,

    /// Background mode (1=TRANSPARENT, 2=OPAQUE)
    bg_mode: u16,

    /// Polygon fill mode (1=ALTERNATE/evenodd, 2=WINDING/nonzero)
    poly_fill_mode: u16,
}

impl Default for GraphicsState {
    fn default() -> Self {
        Self {
            current_pos: (0, 0),
            objects: Vec::new(),
            current_pen: PenStyle::default(),
            current_brush: BrushStyle::default(),
            current_font: FontStyle::default(),
            text_color: 0x000000,
            bg_color: 0xFFFFFF,
            bg_mode: 1,        // TRANSPARENT
            poly_fill_mode: 1, // ALTERNATE
        }
    }
}

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
    /// Processes WMF records sequentially to maintain graphics state and generates SVG
    /// with vector graphics and embedded raster images.
    pub fn convert_to_svg(&self) -> Result<String> {
        let width = self.parser.width() as f64;
        let height = self.parser.height() as f64;

        let mut builder = SvgBuilder::new(width, height);

        // WMF coordinates are in logical units, set viewBox appropriately
        builder = builder.with_viewbox(0.0, 0.0, width, height);

        // Process records sequentially to maintain state
        // We can't parallelize this because graphics state depends on previous records
        let mut state = GraphicsState::default();

        for record in &self.parser.records {
            // Update state based on record type
            self.update_state(record, &mut state);

            // Try to convert record to SVG element
            if let Ok(Some(element)) = self.process_record(record, &mut state) {
                builder.add_element(element);
            }
        }

        Ok(builder.build())
    }

    /// Convert WMF to SVG bytes
    pub fn convert_to_svg_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.convert_to_svg()?.into_bytes())
    }

    /// Update graphics state based on record
    fn update_state(&self, record: &WmfRecord, state: &mut GraphicsState) {
        match record.function {
            // MoveTo (0x020D/0x0214) - Update current position without drawing
            0x020D | 0x0214 if record.params.len() >= 4 => {
                let y = read_i16_le(&record.params, 0).unwrap_or(0);
                let x = read_i16_le(&record.params, 2).unwrap_or(0);
                state.current_pos = (x, y);
            },
            // LineTo (0x0213) - Position update handled in parse_lineto
            // Don't update here to avoid double update
            // SetTextColor (0x020A/0x0209)
            0x020A | 0x0209 if record.params.len() >= 4 => {
                state.text_color = u32::from_le_bytes([
                    record.params[0],
                    record.params[1],
                    record.params[2],
                    record.params[3],
                ]);
            },
            // SetBkColor (0x0201)
            0x0201 if record.params.len() >= 4 => {
                state.bg_color = u32::from_le_bytes([
                    record.params[0],
                    record.params[1],
                    record.params[2],
                    record.params[3],
                ]);
            },
            // SetBkMode (0x0102)
            0x0102 if record.params.len() >= 2 => {
                state.bg_mode = u16::from_le_bytes([record.params[0], record.params[1]]);
            },
            // SetPolyFillMode (0x0106)
            0x0106 if record.params.len() >= 2 => {
                state.poly_fill_mode = u16::from_le_bytes([record.params[0], record.params[1]]);
            },
            // CreatePenIndirect (0x02FA)
            0x02FA => {
                if let Some(pen) = self.parse_create_pen(record) {
                    state.objects.push(Some(WmfObject::Pen(pen)));
                }
            },
            // CreateBrushIndirect (0x02FC)
            0x02FC => {
                if let Some(brush) = self.parse_create_brush(record) {
                    state.objects.push(Some(WmfObject::Brush(brush)));
                }
            },
            // CreateFontIndirect (0x02FE/0x02FB)
            0x02FE | 0x02FB => {
                if let Some(font) = self.parse_create_font(record) {
                    state.objects.push(Some(WmfObject::Font(font)));
                }
            },
            // SelectObject (0x012D)
            0x012D if record.params.len() >= 2 => {
                let index = u16::from_le_bytes([record.params[0], record.params[1]]) as usize;
                if let Some(Some(obj)) = state.objects.get(index) {
                    match obj {
                        WmfObject::Pen(pen) => state.current_pen = *pen,
                        WmfObject::Brush(brush) => state.current_brush = *brush,
                        WmfObject::Font(font) => state.current_font = font.clone(),
                    }
                }
            },
            // DeleteObject (0x01F0)
            0x01F0 if record.params.len() >= 2 => {
                let index = u16::from_le_bytes([record.params[0], record.params[1]]) as usize;
                if index < state.objects.len() {
                    state.objects[index] = None;
                }
            },
            _ => {},
        }
    }

    /// Parse CreatePenIndirect record
    fn parse_create_pen(&self, record: &WmfRecord) -> Option<PenStyle> {
        if record.params.len() < 10 {
            return None;
        }

        let style = read_u16_le(&record.params, 0).ok()?;
        let width = read_u16_le(&record.params, 2).ok()?;
        let color = u32::from_le_bytes([
            record.params[6],
            record.params[7],
            record.params[8],
            record.params[9],
        ]);

        Some(PenStyle {
            style,
            width,
            color,
        })
    }

    /// Parse CreateBrushIndirect record
    fn parse_create_brush(&self, record: &WmfRecord) -> Option<BrushStyle> {
        if record.params.len() < 8 {
            return None;
        }

        let style = read_u16_le(&record.params, 0).ok()?;
        let color = u32::from_le_bytes([
            record.params[2],
            record.params[3],
            record.params[4],
            record.params[5],
        ]);
        let hatch = read_u16_le(&record.params, 6).ok()?;

        Some(BrushStyle {
            style,
            color,
            hatch,
        })
    }

    /// Parse CreateFontIndirect record
    fn parse_create_font(&self, record: &WmfRecord) -> Option<FontStyle> {
        if record.params.len() < 18 {
            return None;
        }

        let height = read_i16_le(&record.params, 0).ok()?;
        let width = read_u16_le(&record.params, 2).ok()?;
        let escapement = read_u16_le(&record.params, 4).ok()?;
        let weight = read_u16_le(&record.params, 8).ok()?;
        let italic_byte = read_u16_le(&record.params, 10).ok()?;
        let italic = (italic_byte & 0xFF) != 0;
        let underline = ((italic_byte >> 8) & 0xFF) != 0;
        let strikeout_byte = read_u16_le(&record.params, 12).ok()?;
        let strikeout = (strikeout_byte & 0xFF) != 0;
        let charset = ((strikeout_byte >> 8) & 0xFF) as u8;

        // Extract font name (starts at offset 18, null-terminated ASCII)
        let name_bytes = &record.params[18..];
        let name_end = name_bytes
            .iter()
            .position(|&b| b == 0)
            .unwrap_or(name_bytes.len());
        let name = if name_end > 0 {
            String::from_utf8_lossy(&name_bytes[..name_end]).to_string()
        } else {
            "Times New Roman".to_string()
        };

        Some(FontStyle {
            height,
            width,
            escapement,
            weight,
            italic,
            underline,
            strikeout,
            charset,
            name,
        })
    }

    /// Process a single WMF record
    fn process_record(
        &self,
        record: &WmfRecord,
        state: &mut GraphicsState,
    ) -> Result<Option<SvgElement>> {
        match record.function {
            // Rectangle
            0x041B => self.parse_rectangle(record, state),
            // Ellipse
            0x0418 => self.parse_ellipse(record, state),
            // Polygon
            0x0324 => self.parse_polygon(record, state),
            // Polyline
            0x0325 => self.parse_polyline(record, state),
            // Arc
            0x0817 => self.parse_arc(record, state),
            // Pie
            0x081A => self.parse_pie(record, state),
            // Chord
            0x0830 => self.parse_chord(record, state),
            // StretchDIB (embedded bitmap)
            0x0F43 => self.parse_stretchdib(record),
            // DIBStretchBlt
            0x0B41 => self.parse_dibstretchblt(record),
            // DIBBitBlt
            0x0940 => self.parse_dibbitblt(record),
            // RoundRect
            0x061C => self.parse_roundrect(record, state),
            // LineTo (0x0213)
            0x0213 => self.parse_lineto(record, state),
            // TextOut (0x0521)
            0x0521 => self.parse_textout(record, state),
            // ExtTextOut (0x0A32)
            0x0A32 => self.parse_exttextout(record, state),
            _ => Ok(None),
        }
    }

    /// Parse META_RECTANGLE record
    fn parse_rectangle(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 8 {
            return Ok(None);
        }

        // WMF uses 16-bit coordinates in little-endian
        let bottom = read_i16_le(&record.params, 0).unwrap_or(0) as f64;
        let right = read_i16_le(&record.params, 2).unwrap_or(0) as f64;
        let top = read_i16_le(&record.params, 4).unwrap_or(0) as f64;
        let left = read_i16_le(&record.params, 6).unwrap_or(0) as f64;

        let mut rect = SvgRect {
            x: left,
            y: top,
            width: (right - left).abs(),
            height: (bottom - top).abs(),
            fill: None,
            stroke: None,
            stroke_width: state.current_pen.width.max(1) as f64,
        };

        // Apply brush fill
        if !state.current_brush.is_null() {
            rect.fill = Some(color::colorref_to_hex(state.current_brush.color));
        }

        // Apply pen stroke
        if !state.current_pen.is_null() {
            rect.stroke = Some(color::colorref_to_hex(state.current_pen.color));
        }

        Ok(Some(SvgElement::Rect(rect)))
    }

    /// Parse META_ELLIPSE record
    fn parse_ellipse(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 8 {
            return Ok(None);
        }

        let bottom = read_i16_le(&record.params, 0).unwrap_or(0) as f64;
        let right = read_i16_le(&record.params, 2).unwrap_or(0) as f64;
        let top = read_i16_le(&record.params, 4).unwrap_or(0) as f64;
        let left = read_i16_le(&record.params, 6).unwrap_or(0) as f64;

        let cx = (left + right) / 2.0;
        let cy = (top + bottom) / 2.0;
        let rx = (right - left) / 2.0;
        let ry = (bottom - top) / 2.0;

        let mut ellipse = SvgEllipse {
            cx,
            cy,
            rx: rx.abs(),
            ry: ry.abs(),
            fill: None,
            stroke: None,
            stroke_width: state.current_pen.width.max(1) as f64,
        };

        // Apply brush fill
        if !state.current_brush.is_null() {
            ellipse.fill = Some(color::colorref_to_hex(state.current_brush.color));
        }

        // Apply pen stroke
        if !state.current_pen.is_null() {
            ellipse.stroke = Some(color::colorref_to_hex(state.current_pen.color));
        }

        Ok(Some(SvgElement::Ellipse(ellipse)))
    }

    /// Parse META_POLYGON record
    fn parse_polygon(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 2 {
            return Ok(None);
        }

        let count = read_i16_le(&record.params, 0).unwrap_or(0) as usize;

        if record.params.len() < 2 + count * 4 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count + 1);

        // Parse points
        let points: Vec<(f64, f64)> = (0..count)
            .map(|i| {
                let offset = 2 + i * 4;
                let x = read_i16_le(&record.params, offset).unwrap_or(0) as f64;
                let y = read_i16_le(&record.params, offset + 2).unwrap_or(0) as f64;
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

        let mut path = SvgPath::new(commands);

        // Apply brush fill
        if !state.current_brush.is_null() {
            path = path.with_fill(color::colorref_to_hex(state.current_brush.color));
        } else {
            path = path.with_fill("none".to_string());
        }

        // Apply pen stroke
        if !state.current_pen.is_null() {
            path = path.with_stroke(color::colorref_to_hex(state.current_pen.color));
            path.stroke_width = state.current_pen.width.max(1) as f64;
        }

        Ok(Some(SvgElement::Path(path)))
    }

    /// Parse META_POLYLINE record
    fn parse_polyline(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 2 {
            return Ok(None);
        }

        let count = read_i16_le(&record.params, 0).unwrap_or(0) as usize;

        if record.params.len() < 2 + count * 4 {
            return Ok(None);
        }

        let mut commands = Vec::with_capacity(count);

        // Parse points
        let points: Vec<(f64, f64)> = (0..count)
            .map(|i| {
                let offset = 2 + i * 4;
                let x = read_i16_le(&record.params, offset).unwrap_or(0) as f64;
                let y = read_i16_le(&record.params, offset + 2).unwrap_or(0) as f64;
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

        let mut path = SvgPath::new(commands);
        path.fill = Some("none".to_string());

        // Apply pen stroke styles
        path = self.apply_pen_styles(path, &state.current_pen);

        Ok(Some(SvgElement::Path(path)))
    }

    /// Parse META_ARC record (0x0817)
    fn parse_arc(&self, record: &WmfRecord, state: &GraphicsState) -> Result<Option<SvgElement>> {
        self.parse_arc_like(record, state, false, false)
    }

    /// Parse META_PIE record (0x081A)
    fn parse_pie(&self, record: &WmfRecord, state: &GraphicsState) -> Result<Option<SvgElement>> {
        self.parse_arc_like(record, state, true, false)
    }

    /// Parse META_CHORD record (0x0830)
    fn parse_chord(&self, record: &WmfRecord, state: &GraphicsState) -> Result<Option<SvgElement>> {
        self.parse_arc_like(record, state, false, true)
    }

    /// Parse arc-like records (arc, pie, chord)
    fn parse_arc_like(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
        is_pie: bool,
        is_chord: bool,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 16 {
            return Ok(None);
        }

        // WMF arc format: EndY, EndX, StartY, StartX, Bottom, Right, Top, Left
        let end_y = read_i16_le(&record.params, 0).unwrap_or(0) as f64;
        let end_x = read_i16_le(&record.params, 2).unwrap_or(0) as f64;
        let start_y = read_i16_le(&record.params, 4).unwrap_or(0) as f64;
        let start_x = read_i16_le(&record.params, 6).unwrap_or(0) as f64;
        let bottom = read_i16_le(&record.params, 8).unwrap_or(0) as f64;
        let right = read_i16_le(&record.params, 10).unwrap_or(0) as f64;
        let top = read_i16_le(&record.params, 12).unwrap_or(0) as f64;
        let left = read_i16_le(&record.params, 14).unwrap_or(0) as f64;

        self.create_arc_path(
            left, top, right, bottom, start_x, start_y, end_x, end_y, state, is_pie, is_chord,
        )
    }

    /// Create SVG path for arc, pie, or chord
    #[allow(clippy::too_many_arguments)]
    fn create_arc_path(
        &self,
        left: f64,
        top: f64,
        right: f64,
        bottom: f64,
        start_x: f64,
        start_y: f64,
        end_x: f64,
        end_y: f64,
        state: &GraphicsState,
        is_pie: bool,
        is_chord: bool,
    ) -> Result<Option<SvgElement>> {
        // Calculate ellipse center and radii
        let cx = (left + right) / 2.0;
        let cy = (top + bottom) / 2.0;
        let rx = (right - left).abs() / 2.0;
        let ry = (bottom - top).abs() / 2.0;

        if rx == 0.0 || ry == 0.0 {
            return Ok(None);
        }

        // Calculate angles for start and end points
        let start_angle = ((start_y - cy) / ry).atan2((start_x - cx) / rx);
        let end_angle = ((end_y - cy) / ry).atan2((end_x - cx) / rx);

        // Calculate actual start and end points on the ellipse
        let actual_start_x = cx + rx * start_angle.cos();
        let actual_start_y = cy + ry * start_angle.sin();
        let actual_end_x = cx + rx * end_angle.cos();
        let actual_end_y = cy + ry * end_angle.sin();

        // Determine if large arc (> 180 degrees)
        let mut angle_diff = end_angle - start_angle;
        if angle_diff < 0.0 {
            angle_diff += 2.0 * std::f64::consts::PI;
        }
        let large_arc = angle_diff > std::f64::consts::PI;

        // Build path commands
        let mut commands = vec![PathCommand::MoveTo {
            x: actual_start_x,
            y: actual_start_y,
        }];

        // Arc command (always counter-clockwise in WMF)
        commands.push(PathCommand::Arc {
            rx,
            ry,
            x_axis_rotation: 0.0,
            large_arc,
            sweep: true, // Counter-clockwise
            x: actual_end_x,
            y: actual_end_y,
        });

        if is_pie {
            // Pie: connect to center and close
            commands.push(PathCommand::LineTo { x: cx, y: cy });
            commands.push(PathCommand::ClosePath);
        } else if is_chord {
            // Chord: connect end to start with straight line
            commands.push(PathCommand::ClosePath);
        }

        let mut path = SvgPath::new(commands);

        // Apply brush fill for pie and chord
        if (is_pie || is_chord) && !state.current_brush.is_null() {
            path.fill = Some(color::colorref_to_hex(state.current_brush.color));
        } else {
            path.fill = Some("none".to_string());
        }

        // Apply pen stroke styles
        path = self.apply_pen_styles(path, &state.current_pen);

        Ok(Some(SvgElement::Path(path)))
    }

    /// Parse META_ROUNDRECT record
    fn parse_roundrect(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 12 {
            return Ok(None);
        }

        let corner_height = read_i16_le(&record.params, 0).unwrap_or(0) as f64;
        let corner_width = read_i16_le(&record.params, 2).unwrap_or(0) as f64;
        let bottom = read_i16_le(&record.params, 4).unwrap_or(0) as f64;
        let right = read_i16_le(&record.params, 6).unwrap_or(0) as f64;
        let top = read_i16_le(&record.params, 8).unwrap_or(0) as f64;
        let left = read_i16_le(&record.params, 10).unwrap_or(0) as f64;

        let rx = corner_width / 2.0;
        let ry = corner_height / 2.0;

        // Create rounded rectangle using path with arcs
        let commands = vec![
            PathCommand::MoveTo {
                x: left + rx,
                y: top,
            },
            PathCommand::LineTo {
                x: right - rx,
                y: top,
            },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: right,
                y: top + ry,
            },
            PathCommand::LineTo {
                x: right,
                y: bottom - ry,
            },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: right - rx,
                y: bottom,
            },
            PathCommand::LineTo {
                x: left + rx,
                y: bottom,
            },
            PathCommand::Arc {
                rx,
                ry,
                x_axis_rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: left,
                y: bottom - ry,
            },
            PathCommand::LineTo {
                x: left,
                y: top + ry,
            },
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

        let mut path = SvgPath::new(commands);

        // Apply brush fill
        if !state.current_brush.is_null() {
            path = path.with_fill(color::colorref_to_hex(state.current_brush.color));
        } else {
            path = path.with_fill("none".to_string());
        }

        // Apply pen stroke
        if !state.current_pen.is_null() {
            path = path.with_stroke(color::colorref_to_hex(state.current_pen.color));
            path.stroke_width = state.current_pen.width.max(1) as f64;
        }

        Ok(Some(SvgElement::Path(path)))
    }

    /// Parse META_LINETO record
    fn parse_lineto(
        &self,
        record: &WmfRecord,
        state: &mut GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 4 {
            return Ok(None);
        }

        let y = read_i16_le(&record.params, 0).unwrap_or(0);
        let x = read_i16_le(&record.params, 2).unwrap_or(0);

        // Use current position from state as start
        let (start_x, start_y) = state.current_pos;

        // Create path from current position to new position
        let path = SvgPath::new(vec![
            PathCommand::MoveTo {
                x: start_x as f64,
                y: start_y as f64,
            },
            PathCommand::LineTo {
                x: x as f64,
                y: y as f64,
            },
        ]);

        // Apply pen stroke styles
        let path = self.apply_pen_styles(path, &state.current_pen);

        // Update current position to end of line
        state.current_pos = (x, y);

        Ok(Some(SvgElement::Path(path)))
    }

    /// Parse META_STRETCHDIB record
    fn parse_stretchdib(&self, record: &WmfRecord) -> Result<Option<SvgElement>> {
        // Extract DIB and convert to PNG for embedding
        if record.params.len() < 20 {
            return Ok(None);
        }

        // Parse destination rectangle (simplified)
        let dest_height = read_i16_le(&record.params, 6).unwrap_or(0) as f64;
        let dest_width = read_i16_le(&record.params, 8).unwrap_or(0) as f64;
        let dest_y = read_i16_le(&record.params, 10).unwrap_or(0) as f64;
        let dest_x = read_i16_le(&record.params, 12).unwrap_or(0) as f64;

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
            let dest_x = read_i16_le(&record.params, 6).unwrap_or(0) as f64;
            let dest_y = read_i16_le(&record.params, 8).unwrap_or(0) as f64;
            let dest_width = read_i16_le(&record.params, 10).unwrap_or(0) as f64;
            let dest_height = read_i16_le(&record.params, 12).unwrap_or(0) as f64;

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
            let dest_x = read_i16_le(&record.params, 4).unwrap_or(0) as f64;
            let dest_y = read_i16_le(&record.params, 6).unwrap_or(0) as f64;
            let width = read_i16_le(&record.params, 8).unwrap_or(0) as f64;
            let height = read_i16_le(&record.params, 10).unwrap_or(0) as f64;

            return Ok(Some(SvgElement::Image(SvgImage::from_png_data(
                dest_x, dest_y, width, height, &png_data,
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

    /// Parse META_TEXTOUT record (0x0521)
    fn parse_textout(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 6 {
            return Ok(None);
        }

        // META_TEXTOUT format:
        // - string length (word)
        // - string data (variable, word-aligned)
        // - y position (word)
        // - x position (word)

        let string_length = read_u16_le(&record.params, 0).unwrap_or(0) as usize;

        if string_length == 0 || record.params.len() < 2 + string_length + 4 {
            return Ok(None);
        }

        // Extract text string (after length, before coordinates)
        let text_bytes = &record.params[2..2 + string_length];
        let text = String::from_utf8_lossy(text_bytes).into_owned();

        // Coordinates are after the string (word-aligned)
        let coord_offset = 2 + (string_length + 1).div_ceil(2) * 2; // Word align
        let y = read_i16_le(&record.params, coord_offset).unwrap_or(0) as f64;
        let x = read_i16_le(&record.params, coord_offset + 2).unwrap_or(0) as f64;

        self.create_text_element(x, y, text, state)
    }

    /// Parse META_EXTTEXTOUT record (0x0A32)
    fn parse_exttextout(
        &self,
        record: &WmfRecord,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        if record.params.len() < 8 {
            return Ok(None);
        }

        // META_EXTTEXTOUT format:
        // - y position (word)
        // - x position (word)
        // - string length (word)
        // - options (word)
        // - optional rectangle (4 words)
        // - string data (variable)

        let y = read_i16_le(&record.params, 0).unwrap_or(0) as f64;
        let x = read_i16_le(&record.params, 2).unwrap_or(0) as f64;
        let string_length = read_u16_le(&record.params, 4).unwrap_or(0) as usize;
        let options = read_u16_le(&record.params, 6).unwrap_or(0);

        if string_length == 0 {
            return Ok(None);
        }

        // Check if rectangle is present (ETO_CLIPPED or ETO_OPAQUE)
        let has_rect = (options & 0x0006) != 0;
        let text_offset = if has_rect { 16 } else { 8 };

        if record.params.len() < text_offset + string_length {
            return Ok(None);
        }

        // Extract text string
        let text_bytes = &record.params[text_offset..text_offset + string_length];
        let text = String::from_utf8_lossy(text_bytes).into_owned();

        self.create_text_element(x, y, text, state)
    }

    /// Apply pen stroke styles to SVG path
    fn apply_pen_styles(&self, mut path: SvgPath, pen: &PenStyle) -> SvgPath {
        // Apply stroke color and width
        if !pen.is_null() {
            path.stroke = Some(color::colorref_to_hex(pen.color));
            path.stroke_width = pen.width.max(1) as f64;

            // Apply dasharray pattern
            if let Some(dasharray) = pen.dasharray() {
                path = path.with_stroke_dasharray(dasharray);
            }

            // Apply linecap style
            path = path.with_stroke_linecap(pen.linecap().to_string());

            // Apply linejoin style
            path = path.with_stroke_linejoin(pen.linejoin().to_string());
        }

        path
    }

    /// Create SVG text element from WMF text data
    fn create_text_element(
        &self,
        x: f64,
        y: f64,
        text: String,
        state: &GraphicsState,
    ) -> Result<Option<SvgElement>> {
        let font = &state.current_font;

        // Calculate font size (WMF height is negative for baseline alignment)
        let font_size = font.height.abs() as f64;

        // Normalize font family name
        let font_family = if font.name.is_empty() {
            "Times New Roman".to_string()
        } else {
            // Map common WMF fonts to SVG-safe names
            match font.name.as_str() {
                "MS Sans Serif" | "Microsoft Sans Serif" => "sans-serif".to_string(),
                "MS Serif" => "serif".to_string(),
                "Courier New" | "Courier" => "monospace".to_string(),
                "Symbol" => "Symbol".to_string(),
                _ => font.name.clone(),
            }
        };

        let mut text_elem = SvgText::new(x, y, text, font_size)
            .with_font_family(font_family)
            .with_fill(color::colorref_to_hex(state.text_color));

        // Apply font weight (bold)
        if font.weight >= 700 {
            text_elem = text_elem.with_font_weight(font.weight);
        }

        // Apply font styles
        if font.italic {
            text_elem = text_elem.with_italic(true);
        }

        if font.underline {
            text_elem = text_elem.with_underline(true);
        }

        if font.strikeout {
            text_elem = text_elem.with_strikethrough(true);
        }

        // Apply rotation if escapement is set
        // WMF escapement is in tenths of degrees
        if font.escapement != 0 {
            let rotation_degrees = -(font.escapement as f64) / 10.0;

            // Calculate rotation matrix for proper text rotation
            let rad = rotation_degrees.to_radians();
            let cos_val = rad.cos();
            let sin_val = rad.sin();

            // Transform matrix: [a, b, c, d, e, f]
            // Rotate around (x, y): matrix(cos, sin, -sin, cos, x, y)
            text_elem = text_elem.with_transform([cos_val, sin_val, -sin_val, cos_val, x, y]);
        }

        Ok(Some(SvgElement::Text(text_elem)))
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
