// GDI Object Management for EMF rendering
//
// This module manages GDI objects (pens, brushes, fonts) that are created,
// selected, and deleted during EMF playback.

use crate::images::svg::color::colorref_to_hex;
use std::collections::HashMap;
use xml_minifier::minified_xml_format;

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

/// Hatch styles from GDI (for BS_HATCHED brushes)
/// Reference: [MS-EMF] Section 2.1.25 HatchStyle Enumeration
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum HatchStyle {
    Horizontal = 0, // -----
    Vertical = 1,   // |||||
    FDiagonal = 2,  // \\\\\
    BDiagonal = 3,  // /////
    Cross = 4,      // +++++
    DiagCross = 5,  // xxxxx
}

impl HatchStyle {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::Horizontal),
            1 => Some(Self::Vertical),
            2 => Some(Self::FDiagonal),
            3 => Some(Self::BDiagonal),
            4 => Some(Self::Cross),
            5 => Some(Self::DiagCross),
            _ => None,
        }
    }

    /// Generate SVG pattern definition for this hatch style
    /// Returns (pattern_id, pattern_svg) tuple
    pub fn to_svg_pattern(&self, pattern_id: &str, color: &str, bg_color: &str) -> String {
        let pattern_size = 8.0; // Standard hatch pattern size

        match self {
            Self::Horizontal => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="0" y1="4" x2="8" y2="4" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color
                )
            },
            Self::Vertical => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="4" y1="0" x2="4" y2="8" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color
                )
            },
            Self::FDiagonal => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="0" y1="0" x2="8" y2="8" stroke="{}" stroke-width="1"/>
  <line x1="-2" y1="6" x2="2" y2="10" stroke="{}" stroke-width="1"/>
  <line x1="6" y1="-2" x2="10" y2="2" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color,
                    color,
                    color
                )
            },
            Self::BDiagonal => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="0" y1="8" x2="8" y2="0" stroke="{}" stroke-width="1"/>
  <line x1="-2" y1="2" x2="2" y2="-2" stroke="{}" stroke-width="1"/>
  <line x1="6" y1="10" x2="10" y2="6" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color,
                    color,
                    color
                )
            },
            Self::Cross => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="0" y1="4" x2="8" y2="4" stroke="{}" stroke-width="1"/>
  <line x1="4" y1="0" x2="4" y2="8" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color,
                    color
                )
            },
            Self::DiagCross => {
                minified_xml_format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}">
  <rect width="{}" height="{}" fill="{}"/>
  <line x1="0" y1="0" x2="8" y2="8" stroke="{}" stroke-width="1"/>
  <line x1="0" y1="8" x2="8" y2="0" stroke="{}" stroke-width="1"/>
  <line x1="-2" y1="6" x2="2" y2="10" stroke="{}" stroke-width="1"/>
  <line x1="6" y1="-2" x2="10" y2="2" stroke="{}" stroke-width="1"/>
  <line x1="-2" y1="2" x2="2" y2="-2" stroke="{}" stroke-width="1"/>
  <line x1="6" y1="10" x2="10" y2="6" stroke="{}" stroke-width="1"/>
</pattern>"#,
                    pattern_id,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    pattern_size,
                    bg_color,
                    color,
                    color,
                    color,
                    color,
                    color,
                    color
                )
            },
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
    pub color: String,                   // RGB hex color (foreground for hatched)
    pub hatch_style: Option<HatchStyle>, // Hatch pattern style if BS_HATCHED
    pub pattern_id: Option<String>,      // Pattern reference ID for SVG
}

impl Default for Brush {
    fn default() -> Self {
        Self {
            style: BrushStyle::Solid,
            color: "#FFFFFF".to_string(),
            hatch_style: None,
            pattern_id: None,
        }
    }
}

impl Brush {
    /// Create brush from EMR_CREATEBRUSHINDIRECT data
    ///
    /// # Arguments
    /// * `style` - Brush style (BS_SOLID, BS_HATCHED, etc.)
    /// * `colorref` - COLORREF value (RGB color)
    /// * `hatch` - Hatch style value (only used if style is BS_HATCHED)
    pub fn from_emr_data(style: u32, colorref: u32, hatch: u32) -> Self {
        let brush_style = BrushStyle::from_u32(style).unwrap_or(BrushStyle::Solid);
        let hatch_style = if brush_style == BrushStyle::Hatched {
            HatchStyle::from_u32(hatch)
        } else {
            None
        };

        Self {
            style: brush_style,
            color: colorref_to_hex(colorref),
            hatch_style,
            pattern_id: None,
        }
    }

    /// Convert to SVG fill attribute
    ///
    /// Returns either a color value, "none", or a pattern reference.
    /// For hatched brushes, returns a pattern reference that must be
    /// defined in the SVG <defs> section.
    pub fn to_svg_fill(&self) -> String {
        match self.style {
            BrushStyle::Null => "none".to_string(),
            BrushStyle::Solid => self.color.clone(),
            BrushStyle::Hatched => {
                // For hatched brushes, return pattern reference if available
                if let Some(ref pattern_id) = self.pattern_id {
                    format!("url(#{})", pattern_id)
                } else {
                    // Fallback to solid color if pattern not generated
                    self.color.clone()
                }
            },
            // For other pattern types, use solid color as fallback
            // Full pattern support would require bitmap handling
            BrushStyle::Pattern
            | BrushStyle::Indexed
            | BrushStyle::DibPattern
            | BrushStyle::DibPatternPt
            | BrushStyle::Pattern8x8
            | BrushStyle::DibPattern8x8
            | BrushStyle::MonoPattern => {
                // TODO: Full pattern/bitmap brush support
                // For now, use the brush color as a reasonable fallback
                self.color.clone()
            },
        }
    }

    /// Generate SVG pattern definition for hatched brush
    ///
    /// This should be called to generate the pattern definition that goes
    /// in the <defs> section of the SVG document.
    ///
    /// # Arguments
    /// * `pattern_id` - Unique ID for this pattern
    /// * `bg_color` - Background color for the pattern
    ///
    /// # Returns
    /// SVG pattern definition string, or None if not a hatched brush
    pub fn generate_svg_pattern(&mut self, pattern_id: String, bg_color: &str) -> Option<String> {
        if self.style == BrushStyle::Hatched {
            if let Some(hatch_style) = self.hatch_style {
                self.pattern_id = Some(pattern_id.clone());
                Some(hatch_style.to_svg_pattern(&pattern_id, &self.color, bg_color))
            } else {
                None
            }
        } else {
            None
        }
    }

    /// Check if this brush requires a pattern definition
    pub fn needs_pattern(&self) -> bool {
        matches!(self.style, BrushStyle::Hatched) && self.hatch_style.is_some()
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
        let brush = Brush::from_emr_data(0, 0x00FF00, 0); // Green solid brush
        assert_eq!(brush.style, BrushStyle::Solid);
        assert_eq!(brush.color, "#00FF00");
        assert!(brush.hatch_style.is_none());
    }

    #[test]
    fn test_hatched_brush_creation() {
        let brush = Brush::from_emr_data(2, 0xFF0000, 4); // Red cross-hatched brush
        assert_eq!(brush.style, BrushStyle::Hatched);
        assert_eq!(brush.color, "#FF0000");
        assert_eq!(brush.hatch_style, Some(HatchStyle::Cross));
        assert!(brush.needs_pattern());
    }

    #[test]
    fn test_brush_svg_fill() {
        // Solid brush
        let solid = Brush::from_emr_data(0, 0x00FF00, 0);
        assert_eq!(solid.to_svg_fill(), "#00FF00");

        // Null brush
        let null = Brush::from_emr_data(1, 0x000000, 0);
        assert_eq!(null.to_svg_fill(), "none");

        // Hatched brush without pattern ID (fallback)
        let hatched = Brush::from_emr_data(2, 0xFF0000, 0);
        assert_eq!(hatched.to_svg_fill(), "#FF0000");
    }

    #[test]
    fn test_hatch_pattern_generation() {
        let mut brush = Brush::from_emr_data(2, 0xFF0000, 0); // Horizontal hatch
        let pattern = brush.generate_svg_pattern("pattern-1".to_string(), "#FFFFFF");
        assert!(pattern.is_some());

        let pattern_svg = pattern.unwrap();
        assert!(pattern_svg.contains("pattern-1"));
        assert!(pattern_svg.contains("#FF0000")); // Foreground color
        assert!(pattern_svg.contains("#FFFFFF")); // Background color

        // Now the brush should have pattern_id set
        assert_eq!(brush.pattern_id, Some("pattern-1".to_string()));
        assert_eq!(brush.to_svg_fill(), "url(#pattern-1)");
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
