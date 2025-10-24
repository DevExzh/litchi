// Device Context (DC) State Management for EMF rendering
//
// The device context maintains the current graphics state including selected
// objects, transforms, colors, and modes.

use super::gdi_objects::{Brush, Font, Pen};
use crate::images::svg::color::colorref_to_hex;

/// Text alignment modes
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextAlign(pub u16);

impl TextAlign {
    pub const NOUPDATECP: u16 = 0x0000;
    pub const UPDATECP: u16 = 0x0001;
    pub const LEFT: u16 = 0x0000;
    pub const RIGHT: u16 = 0x0002;
    pub const CENTER: u16 = 0x0006;
    pub const TOP: u16 = 0x0000;
    pub const BOTTOM: u16 = 0x0008;
    pub const BASELINE: u16 = 0x0018;
    pub const RTLREADING: u16 = 0x0100;

    pub fn is_center(&self) -> bool {
        (self.0 & 0x0006) == Self::CENTER
    }

    pub fn is_right(&self) -> bool {
        (self.0 & 0x0002) == Self::RIGHT
    }

    pub fn is_bottom(&self) -> bool {
        (self.0 & 0x0008) == Self::BOTTOM
    }

    pub fn is_baseline(&self) -> bool {
        (self.0 & 0x0018) == Self::BASELINE
    }

    pub fn to_svg_anchor(&self) -> &str {
        if self.is_center() {
            "middle"
        } else if self.is_right() {
            "end"
        } else {
            "start"
        }
    }
}

/// Background mix mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum BackgroundMode {
    Transparent = 1,
    Opaque = 2,
}

impl BackgroundMode {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Transparent),
            2 => Some(Self::Opaque),
            _ => None,
        }
    }
}

/// Polygon fill mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum PolyFillMode {
    Alternate = 1,
    Winding = 2,
}

impl PolyFillMode {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Alternate),
            2 => Some(Self::Winding),
            _ => None,
        }
    }

    pub fn to_svg_fill_rule(&self) -> &str {
        match self {
            Self::Alternate => "evenodd",
            Self::Winding => "nonzero",
        }
    }
}

/// World transform (2D affine transformation matrix)
#[derive(Debug, Clone, Copy)]
pub struct WorldTransform {
    pub m11: f32, // Horizontal scaling
    pub m12: f32, // Horizontal shearing
    pub m21: f32, // Vertical shearing
    pub m22: f32, // Vertical scaling
    pub dx: f32,  // Horizontal translation
    pub dy: f32,  // Vertical translation
}

impl Default for WorldTransform {
    fn default() -> Self {
        Self::identity()
    }
}

impl WorldTransform {
    /// Identity transform
    pub fn identity() -> Self {
        Self {
            m11: 1.0,
            m12: 0.0,
            m21: 0.0,
            m22: 1.0,
            dx: 0.0,
            dy: 0.0,
        }
    }

    /// Apply transform to a point
    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        let new_x = (self.m11 as f64) * x + (self.m21 as f64) * y + (self.dx as f64);
        let new_y = (self.m12 as f64) * x + (self.m22 as f64) * y + (self.dy as f64);
        (new_x, new_y)
    }

    /// Convert to SVG transform matrix
    pub fn to_svg_matrix(&self) -> String {
        format!(
            "matrix({},{},{},{},{},{})",
            self.m11, self.m12, self.m21, self.m22, self.dx, self.dy
        )
    }

    /// Check if this is identity transform
    pub fn is_identity(&self) -> bool {
        (self.m11 - 1.0).abs() < 1e-6
            && self.m12.abs() < 1e-6
            && self.m21.abs() < 1e-6
            && (self.m22 - 1.0).abs() < 1e-6
            && self.dx.abs() < 1e-6
            && self.dy.abs() < 1e-6
    }
}

/// Device Context state
///
/// Represents the current graphics state including selected objects,
/// transforms, colors, and various modes.
#[derive(Debug, Clone)]
pub struct DeviceContext {
    // Selected objects
    pub pen: Pen,
    pub brush: Brush,
    pub font: Font,

    // Current position
    pub current_x: f64,
    pub current_y: f64,

    // Colors
    pub text_color: String,
    pub background_color: String,

    // Modes
    pub background_mode: BackgroundMode,
    pub poly_fill_mode: PolyFillMode,
    pub text_align: TextAlign,

    // Transforms
    pub world_transform: WorldTransform,

    // Viewport and Window
    pub viewport_org_x: i32,
    pub viewport_org_y: i32,
    pub viewport_ext_x: i32,
    pub viewport_ext_y: i32,
    pub window_org_x: i32,
    pub window_org_y: i32,
    pub window_ext_x: i32,
    pub window_ext_y: i32,

    // Map mode
    pub map_mode: u32,

    // Clipping
    pub clip_id: Option<String>,
}

impl Default for DeviceContext {
    fn default() -> Self {
        Self {
            pen: Pen::default(),
            brush: Brush::default(),
            font: Font::default(),
            current_x: 0.0,
            current_y: 0.0,
            text_color: "#000000".to_string(),
            background_color: "#FFFFFF".to_string(),
            background_mode: BackgroundMode::Opaque,
            poly_fill_mode: PolyFillMode::Alternate,
            text_align: TextAlign(TextAlign::LEFT | TextAlign::TOP),
            world_transform: WorldTransform::identity(),
            viewport_org_x: 0,
            viewport_org_y: 0,
            viewport_ext_x: 1,
            viewport_ext_y: 1,
            window_org_x: 0,
            window_org_y: 0,
            window_ext_x: 1,
            window_ext_y: 1,
            map_mode: 1, // MM_TEXT
            clip_id: None,
        }
    }
}

impl DeviceContext {
    /// Transform point from logical to device coordinates
    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        // First apply window-to-viewport mapping
        let vp_x = if self.window_ext_x != 0 {
            (x - self.window_org_x as f64) * (self.viewport_ext_x as f64 / self.window_ext_x as f64)
                + self.viewport_org_x as f64
        } else {
            x
        };

        let vp_y = if self.window_ext_y != 0 {
            (y - self.window_org_y as f64) * (self.viewport_ext_y as f64 / self.window_ext_y as f64)
                + self.viewport_org_y as f64
        } else {
            y
        };

        // Then apply world transform
        self.world_transform.transform_point(vp_x, vp_y)
    }

    /// Set text color from COLORREF
    pub fn set_text_color(&mut self, colorref: u32) {
        self.text_color = colorref_to_hex(colorref);
    }

    /// Set background color from COLORREF
    pub fn set_background_color(&mut self, colorref: u32) {
        self.background_color = colorref_to_hex(colorref);
    }

    /// Get current pen stroke attributes for SVG
    pub fn get_stroke_attrs(&self) -> Vec<(String, String)> {
        self.pen.to_svg_attrs()
    }

    /// Get current brush fill attribute for SVG
    pub fn get_fill_attr(&self) -> String {
        self.brush.to_svg_fill()
    }

    /// Get current text style attributes for SVG
    pub fn get_text_attrs(&self) -> Vec<(String, String)> {
        let mut attrs = self.font.to_svg_attrs();
        attrs.push(("fill".to_string(), self.text_color.clone()));
        attrs
    }
}

/// Device Context stack for SaveDC/RestoreDC
pub struct DeviceContextStack {
    stack: Vec<DeviceContext>,
}

impl DeviceContextStack {
    /// Create new empty stack
    pub fn new() -> Self {
        Self { stack: Vec::new() }
    }

    /// Push current DC onto stack
    pub fn push(&mut self, dc: DeviceContext) {
        self.stack.push(dc);
    }

    /// Pop DC from stack
    pub fn pop(&mut self) -> Option<DeviceContext> {
        self.stack.pop()
    }

    /// Pop multiple DCs (for RestoreDC with negative index)
    pub fn pop_to(&mut self, index: isize) -> Option<DeviceContext> {
        if index < 0 {
            // Negative index means relative from top
            let abs_index = (-index) as usize;
            if abs_index <= self.stack.len() {
                // Pop abs_index items
                for _ in 1..abs_index {
                    self.stack.pop();
                }
                return self.stack.pop();
            }
        } else if index > 0 {
            // Positive index means absolute position
            let abs_index = (index - 1) as usize;
            if abs_index < self.stack.len() {
                self.stack.truncate(abs_index + 1);
                return self.stack.pop();
            }
        }
        None
    }

    /// Get current stack depth
    pub fn depth(&self) -> usize {
        self.stack.len()
    }
}

impl Default for DeviceContextStack {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_align() {
        let left_top = TextAlign(TextAlign::LEFT | TextAlign::TOP);
        assert_eq!(left_top.to_svg_anchor(), "start");

        let center = TextAlign(TextAlign::CENTER);
        assert!(center.is_center());
        assert_eq!(center.to_svg_anchor(), "middle");

        let right = TextAlign(TextAlign::RIGHT);
        assert!(right.is_right());
        assert_eq!(right.to_svg_anchor(), "end");
    }

    #[test]
    fn test_world_transform() {
        let identity = WorldTransform::identity();
        assert!(identity.is_identity());

        let (x, y) = identity.transform_point(10.0, 20.0);
        assert_eq!(x, 10.0);
        assert_eq!(y, 20.0);

        let translate = WorldTransform {
            m11: 1.0,
            m12: 0.0,
            m21: 0.0,
            m22: 1.0,
            dx: 5.0,
            dy: 10.0,
        };
        let (x, y) = translate.transform_point(10.0, 20.0);
        assert_eq!(x, 15.0);
        assert_eq!(y, 30.0);
    }

    #[test]
    fn test_dc_stack() {
        let mut stack = DeviceContextStack::new();
        assert_eq!(stack.depth(), 0);

        let dc1 = DeviceContext::default();
        stack.push(dc1.clone());
        assert_eq!(stack.depth(), 1);

        let dc2 = DeviceContext::default();
        stack.push(dc2);
        assert_eq!(stack.depth(), 2);

        let popped = stack.pop();
        assert!(popped.is_some());
        assert_eq!(stack.depth(), 1);

        let popped = stack.pop();
        assert!(popped.is_some());
        assert_eq!(stack.depth(), 0);

        let popped = stack.pop();
        assert!(popped.is_none());
    }

    #[test]
    fn test_poly_fill_mode() {
        assert_eq!(PolyFillMode::Alternate.to_svg_fill_rule(), "evenodd");
        assert_eq!(PolyFillMode::Winding.to_svg_fill_rule(), "nonzero");
    }
}
