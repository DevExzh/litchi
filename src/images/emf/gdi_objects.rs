// GDI Object Management for EMF rendering
//
// This module manages GDI objects (pens, brushes, fonts) that are created,
// selected, and deleted during EMF playback.

use crate::images::svg::color::colorref_to_hex;
use std::collections::HashMap;

/// Pen styles from GDI
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum PenStyle {
    Solid = 0,
    Dash = 1,
    Dot = 2,
    DashDot = 3,
    DashDotDot = 4,
    Null = 5,
    InsideFrame = 6,
}

impl PenStyle {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value & 0xFF {
            0 => Some(Self::Solid),
            1 => Some(Self::Dash),
            2 => Some(Self::Dot),
            3 => Some(Self::DashDot),
            4 => Some(Self::DashDotDot),
            5 => Some(Self::Null),
            6 => Some(Self::InsideFrame),
            _ => None,
        }
    }

    /// Convert to SVG stroke-dasharray attribute
    pub fn to_dasharray(&self, width: f64) -> Option<String> {
        match self {
            Self::Solid | Self::InsideFrame | Self::Null => None,
            Self::Dash => Some(format!("{},{}", width * 3.0, width)),
            Self::Dot => Some(format!("{},{}", width, width)),
            Self::DashDot => Some(format!("{},{},{},{}", width * 3.0, width, width, width)),
            Self::DashDotDot => Some(format!(
                "{},{},{},{},{},{}",
                width * 3.0,
                width,
                width,
                width,
                width,
                width
            )),
        }
    }
}

/// Brush styles from GDI
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum BrushStyle {
    Solid = 0,
    Null = 1,
    Hatched = 2,
    Pattern = 3,
    Indexed = 4,
    DibPattern = 5,
    DibPatternPt = 6,
    Pattern8x8 = 7,
    DibPattern8x8 = 8,
    MonoPattern = 9,
}

impl BrushStyle {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::Solid),
            1 => Some(Self::Null),
            2 => Some(Self::Hatched),
            3 => Some(Self::Pattern),
            4 => Some(Self::Indexed),
            5 => Some(Self::DibPattern),
            6 => Some(Self::DibPatternPt),
            7 => Some(Self::Pattern8x8),
            8 => Some(Self::DibPattern8x8),
            9 => Some(Self::MonoPattern),
            _ => None,
        }
    }
}

/// Pen object for drawing
#[derive(Debug, Clone)]
pub struct Pen {
    pub style: PenStyle,
    pub width: f64,
    pub color: String, // RGB hex color
}

impl Default for Pen {
    fn default() -> Self {
        Self {
            style: PenStyle::Solid,
            width: 1.0,
            color: "#000000".to_string(),
        }
    }
}

impl Pen {
    /// Create pen from EMR_CREATEPEN data
    pub fn from_emr_data(style: u32, width: i32, colorref: u32) -> Self {
        Self {
            style: PenStyle::from_u32(style).unwrap_or(PenStyle::Solid),
            width: width.max(1) as f64,
            color: colorref_to_hex(colorref),
        }
    }

    /// Convert to SVG stroke attributes
    pub fn to_svg_attrs(&self) -> Vec<(String, String)> {
        let mut attrs = vec![
            ("stroke".to_string(), self.color.clone()),
            ("stroke-width".to_string(), self.width.to_string()),
        ];

        if let Some(dasharray) = self.style.to_dasharray(self.width) {
            attrs.push(("stroke-dasharray".to_string(), dasharray));
        }

        attrs
    }
}

/// Brush object for filling
#[derive(Debug, Clone)]
pub struct Brush {
    pub style: BrushStyle,
    pub color: String, // RGB hex color
    pub hatch: u32,    // Hatch pattern if hatched
}

impl Default for Brush {
    fn default() -> Self {
        Self {
            style: BrushStyle::Solid,
            color: "#FFFFFF".to_string(),
            hatch: 0,
        }
    }
}

impl Brush {
    /// Create brush from EMR_CREATEBRUSHINDIRECT data
    pub fn from_emr_data(style: u32, colorref: u32, hatch: u32) -> Self {
        Self {
            style: BrushStyle::from_u32(style).unwrap_or(BrushStyle::Solid),
            color: colorref_to_hex(colorref),
            hatch,
        }
    }

    /// Convert to SVG fill attribute
    pub fn to_svg_fill(&self) -> String {
        match self.style {
            BrushStyle::Null => "none".to_string(),
            BrushStyle::Solid => self.color.clone(),
            _ => self.color.clone(), // Simplified for now
        }
    }
}

/// Font object for text rendering
#[derive(Debug, Clone)]
pub struct Font {
    pub height: i32,
    pub width: i32,
    pub escapement: i32,
    pub orientation: i32,
    pub weight: i32,
    pub italic: bool,
    pub underline: bool,
    pub strike_out: bool,
    pub charset: u8,
    pub face_name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: 12,
            width: 0,
            escapement: 0,
            orientation: 0,
            weight: 400,
            italic: false,
            underline: false,
            strike_out: false,
            charset: 1, // DEFAULT_CHARSET
            face_name: "Arial".to_string(),
        }
    }
}

impl Font {
    /// Calculate font size for SVG
    pub fn svg_font_size(&self) -> f64 {
        self.height.abs() as f64 * 0.75 // Convert from logical units
    }

    /// Get font weight for SVG
    pub fn svg_font_weight(&self) -> u16 {
        if self.weight >= 700 {
            700
        } else if self.weight >= 400 {
            400
        } else {
            self.weight as u16
        }
    }

    /// Convert to SVG text style attributes
    pub fn to_svg_attrs(&self) -> Vec<(String, String)> {
        let mut attrs = vec![
            ("font-family".to_string(), self.face_name.clone()),
            (
                "font-size".to_string(),
                format!("{}px", self.svg_font_size()),
            ),
            (
                "font-weight".to_string(),
                self.svg_font_weight().to_string(),
            ),
        ];

        if self.italic {
            attrs.push(("font-style".to_string(), "italic".to_string()));
        }

        let mut decorations = Vec::new();
        if self.underline {
            decorations.push("underline");
        }
        if self.strike_out {
            decorations.push("line-through");
        }
        if !decorations.is_empty() {
            attrs.push(("text-decoration".to_string(), decorations.join(" ")));
        }

        attrs
    }
}

/// GDI Object types
#[derive(Debug, Clone)]
pub enum GdiObject {
    Pen(Pen),
    Brush(Brush),
    Font(Font),
    Palette, // Placeholder
    Region,  // Placeholder
}

/// GDI Object table
///
/// Manages the object table for EMF playback. Objects are created with
/// EMR_CREATE* records, selected with EMR_SELECTOBJECT, and deleted
/// with EMR_DELETEOBJECT.
pub struct ObjectTable {
    objects: HashMap<u32, GdiObject>,
    next_free: u32,
}

impl ObjectTable {
    /// Create new empty object table
    pub fn new() -> Self {
        Self {
            objects: HashMap::new(),
            next_free: 0,
        }
    }

    /// Create object and return its handle
    pub fn create_object(&mut self, object: GdiObject) -> u32 {
        let handle = self.next_free;
        self.objects.insert(handle, object);
        self.next_free += 1;
        handle
    }

    /// Get object by handle
    pub fn get(&self, handle: u32) -> Option<&GdiObject> {
        self.objects.get(&handle)
    }

    /// Delete object by handle
    pub fn delete(&mut self, handle: u32) -> bool {
        self.objects.remove(&handle).is_some()
    }

    /// Check if handle exists
    pub fn exists(&self, handle: u32) -> bool {
        self.objects.contains_key(&handle)
    }
}

impl Default for ObjectTable {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pen_creation() {
        let pen = Pen::from_emr_data(0, 2, 0x0000FF); // Red, 2px wide
        assert_eq!(pen.style, PenStyle::Solid);
        assert_eq!(pen.width, 2.0);
        assert_eq!(pen.color, "#FF0000");
    }

    #[test]
    fn test_brush_creation() {
        let brush = Brush::from_emr_data(0, 0x00FF00, 0); // Green
        assert_eq!(brush.style, BrushStyle::Solid);
        assert_eq!(brush.color, "#00FF00");
    }

    #[test]
    fn test_object_table() {
        let mut table = ObjectTable::new();

        let pen = Pen::default();
        let handle = table.create_object(GdiObject::Pen(pen));

        assert!(table.exists(handle));
        assert!(matches!(table.get(handle), Some(GdiObject::Pen(_))));

        assert!(table.delete(handle));
        assert!(!table.exists(handle));
    }

    #[test]
    fn test_pen_dasharray() {
        let dash_pen = Pen {
            style: PenStyle::Dash,
            width: 2.0,
            color: "#000000".to_string(),
        };
        assert_eq!(dash_pen.style.to_dasharray(2.0), Some("6,2".to_string()));

        let dot_pen = Pen {
            style: PenStyle::Dot,
            width: 1.0,
            color: "#000000".to_string(),
        };
        assert_eq!(dot_pen.style.to_dasharray(1.0), Some("1,1".to_string()));
    }
}
