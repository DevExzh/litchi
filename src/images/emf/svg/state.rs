/// SVG Rendering State Management
///
/// Manages device context stack, transforms, clipping, and graphics state
use super::path::PathBuilder;
use crate::images::emf::records::*;

/// Complete rendering state
pub struct RenderState {
    /// Device context stack
    pub dc_stack: Vec<DeviceContext>,
    /// Current device context
    pub dc: DeviceContext,
    /// Current path being built
    pub path_builder: Option<PathBuilder>,
    /// Whether we're recording a path
    pub in_path: bool,
}

impl RenderState {
    /// Create new rendering state
    pub fn new() -> Self {
        Self {
            dc_stack: Vec::new(),
            dc: DeviceContext::default(),
            path_builder: None,
            in_path: false,
        }
    }

    /// Push current DC to stack (SaveDC)
    pub fn push_dc(&mut self) {
        self.dc_stack.push(self.dc.clone());
    }

    /// Pop DC from stack (RestoreDC)
    pub fn pop_dc(&mut self, index: i32) {
        if index < 0 {
            // Relative index (-1 = most recent)
            let idx = (-index as usize).saturating_sub(1);
            if let Some(dc) = self
                .dc_stack
                .get(self.dc_stack.len().saturating_sub(idx + 1))
                .cloned()
            {
                self.dc_stack
                    .truncate(self.dc_stack.len().saturating_sub(idx + 1));
                self.dc = dc;
            }
        } else if index > 0 {
            // Absolute index (1-based)
            let idx = (index as usize).saturating_sub(1);
            if let Some(dc) = self.dc_stack.get(idx).cloned() {
                self.dc_stack.truncate(idx);
                self.dc = dc;
            }
        }
    }

    /// Start building a path
    pub fn begin_path(&mut self) {
        self.path_builder = Some(PathBuilder::new());
        self.in_path = true;
    }

    /// End path building
    pub fn end_path(&mut self) {
        self.in_path = false;
    }

    /// Get current path and reset builder
    pub fn take_path(&mut self) -> Option<PathBuilder> {
        self.path_builder.take()
    }
}

impl Default for RenderState {
    fn default() -> Self {
        Self::new()
    }
}

/// Device Context state
#[derive(Debug, Clone)]
pub struct DeviceContext {
    /// Current position
    pub current_pos: (f64, f64),

    /// Transform state
    pub world_transform: XForm,
    pub window_org: (i32, i32),
    pub window_ext: (i32, i32),
    pub viewport_org: (i32, i32),
    pub viewport_ext: (i32, i32),
    pub map_mode: u32,

    /// Drawing state
    pub pen: Pen,
    pub brush: Brush,
    pub font: Font,

    /// Colors
    pub text_color: ColorRef,
    pub bg_color: ColorRef,

    /// Modes
    pub bg_mode: u32, // BackgroundMode
    pub poly_fill_mode: u32, // PolyFillMode
    pub text_align: u32,     // TextAlign flags
    pub rop2: u32,           // ROP2 mode
    pub arc_direction: bool, // true = clockwise
    pub miter_limit: f64,

    /// Clipping
    pub clip_region: Option<Vec<(f64, f64, f64, f64)>>, // Rectangles
}

impl Default for DeviceContext {
    fn default() -> Self {
        Self {
            current_pos: (0.0, 0.0),
            world_transform: XForm::default(),
            window_org: (0, 0),
            window_ext: (1, 1),
            viewport_org: (0, 0),
            viewport_ext: (1, 1),
            map_mode: 1, // MM_TEXT
            pen: Pen::default(),
            brush: Brush::default(),
            font: Font::default(),
            text_color: ColorRef::from_rgb(0, 0, 0), // Black
            bg_color: ColorRef::from_rgb(255, 255, 255), // White
            bg_mode: 2,                              // OPAQUE
            poly_fill_mode: 1,                       // ALTERNATE
            text_align: 0,
            rop2: 13,             // COPYPEN
            arc_direction: false, // Counter-clockwise
            miter_limit: 10.0,
            clip_region: None,
        }
    }
}

impl DeviceContext {
    /// Transform a point from logical to device coordinates
    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        // Apply window/viewport mapping
        let (wx, wy) = (x - self.window_org.0 as f64, y - self.window_org.1 as f64);
        let scale_x = self.viewport_ext.0 as f64 / self.window_ext.0 as f64;
        let scale_y = self.viewport_ext.1 as f64 / self.window_ext.1 as f64;
        let vx = wx * scale_x + self.viewport_org.0 as f64;
        let vy = wy * scale_y + self.viewport_org.1 as f64;

        // Apply world transform
        self.world_transform.transform_point(vx, vy)
    }

    /// Get SVG stroke attributes
    pub fn get_stroke_attrs(&self) -> String {
        let mut attrs = String::new();

        // Stroke color
        if self.pen.style != pen_style::NULL {
            attrs.push_str(&format!("stroke=\"{}\" ", self.pen.color.to_svg_color()));

            // Stroke width
            if self.pen.width > 1.0 {
                attrs.push_str(&format!("stroke-width=\"{}\" ", self.pen.width));
            }

            // Stroke dash array
            if let Some(ref dash) = self.pen.dash_pattern {
                attrs.push_str(&format!("stroke-dasharray=\"{}\" ", dash));
            }

            // Line cap
            if self.pen.end_cap != pen_style::ENDCAP_FLAT {
                let cap = match self.pen.end_cap {
                    pen_style::ENDCAP_ROUND => "round",
                    pen_style::ENDCAP_SQUARE => "square",
                    _ => "butt",
                };
                attrs.push_str(&format!("stroke-linecap=\"{}\" ", cap));
            }

            // Line join
            if self.pen.line_join != pen_style::JOIN_MITER {
                let join = match self.pen.line_join {
                    pen_style::JOIN_ROUND => "round",
                    pen_style::JOIN_BEVEL => "bevel",
                    _ => "miter",
                };
                attrs.push_str(&format!("stroke-linejoin=\"{}\" ", join));
            }

            // Miter limit
            if self.miter_limit != 10.0 {
                attrs.push_str(&format!("stroke-miterlimit=\"{}\" ", self.miter_limit));
            }
        } else {
            attrs.push_str("stroke=\"none\" ");
        }

        attrs
    }

    /// Get SVG fill attribute
    pub fn get_fill_attr(&self) -> String {
        if self.brush.style == brush_style::NULL {
            "fill=\"none\"".to_string()
        } else {
            format!("fill=\"{}\"", self.brush.color.to_svg_color())
        }
    }

    /// Get SVG fill-rule attribute
    pub fn get_fill_rule(&self) -> Option<String> {
        if self.poly_fill_mode == 1 {
            Some("fill-rule=\"evenodd\"".to_string()) // ALTERNATE
        } else {
            None // WINDING is default
        }
    }

    /// Get SVG transform attribute
    pub fn get_transform_attr(&self) -> Option<String> {
        self.world_transform.to_svg_transform()
    }
}

/// Pen state
#[derive(Debug, Clone)]
pub struct Pen {
    pub style: u32,
    pub width: f64,
    pub color: ColorRef,
    pub end_cap: u32,
    pub line_join: u32,
    pub dash_pattern: Option<String>,
}

impl Default for Pen {
    fn default() -> Self {
        Self {
            style: pen_style::SOLID,
            width: 1.0,
            color: ColorRef::from_rgb(0, 0, 0), // Black
            end_cap: pen_style::ENDCAP_FLAT,
            line_join: pen_style::JOIN_MITER,
            dash_pattern: None,
        }
    }
}

impl Pen {
    /// Create from EMR_CREATEPEN record
    pub fn from_create_pen(pen_style: u32, width: u32, color: ColorRef) -> Self {
        let base_style = pen_style & 0xFF;
        let dash_pattern = match base_style {
            pen_style::DASH => Some("5 2".to_string()),
            pen_style::DOT => Some("1 1".to_string()),
            pen_style::DASHDOT => Some("5 2 1 2".to_string()),
            pen_style::DASHDOTDOT => Some("5 2 1 2 1 2".to_string()),
            _ => None,
        };

        Self {
            style: base_style,
            width: width as f64,
            color,
            end_cap: pen_style & 0x0F00,
            line_join: pen_style & 0xF000,
            dash_pattern,
        }
    }
}

/// Brush state
#[derive(Debug, Clone)]
pub struct Brush {
    pub style: u32,
    pub color: ColorRef,
    pub hatch: Option<u32>,
}

impl Default for Brush {
    fn default() -> Self {
        Self {
            style: brush_style::SOLID,
            color: ColorRef::from_rgb(255, 255, 255), // White
            hatch: None,
        }
    }
}

/// Font state
#[derive(Debug, Clone)]
pub struct Font {
    pub height: f64,
    pub width: f64,
    pub escapement: f64,
    pub weight: i32,
    pub italic: bool,
    pub underline: bool,
    pub strike_out: bool,
    pub face_name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: 12.0,
            width: 0.0,
            escapement: 0.0,
            weight: font_weight::NORMAL,
            italic: false,
            underline: false,
            strike_out: false,
            face_name: "Arial".to_string(),
        }
    }
}

impl Font {
    /// Get SVG font attributes
    pub fn to_svg_attrs(&self) -> String {
        let mut attrs = String::new();

        // Font size (convert from logical height)
        let size = self.height.abs();
        if size > 0.0 {
            attrs.push_str(&format!("font-size=\"{}\" ", size));
        }

        // Font family
        if !self.face_name.is_empty() {
            attrs.push_str(&format!("font-family=\"{}\" ", self.face_name));
        }

        // Font weight
        if self.weight != font_weight::NORMAL {
            attrs.push_str(&format!("font-weight=\"{}\" ", self.weight));
        }

        // Font style
        if self.italic {
            attrs.push_str("font-style=\"italic\" ");
        }

        // Text decoration
        let mut decorations = Vec::new();
        if self.underline {
            decorations.push("underline");
        }
        if self.strike_out {
            decorations.push("line-through");
        }
        if !decorations.is_empty() {
            attrs.push_str(&format!("text-decoration=\"{}\" ", decorations.join(" ")));
        }

        attrs
    }
}
