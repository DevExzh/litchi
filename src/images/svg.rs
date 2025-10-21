// SVG generation module
//
// Provides high-performance SVG generation from vector graphics operations

use std::fmt::Write;

/// SVG path commands
#[derive(Debug, Clone)]
pub enum PathCommand {
    /// Move to absolute position
    MoveTo { x: f64, y: f64 },
    /// Line to absolute position
    LineTo { x: f64, y: f64 },
    /// Cubic Bezier curve
    CubicBezier {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    /// Quadratic Bezier curve
    QuadraticBezier { x1: f64, y1: f64, x: f64, y: f64 },
    /// Arc
    Arc {
        rx: f64,
        ry: f64,
        x_axis_rotation: f64,
        large_arc: bool,
        sweep: bool,
        x: f64,
        y: f64,
    },
    /// Close path
    ClosePath,
}

impl PathCommand {
    /// Convert to SVG path string
    pub fn to_svg(&self) -> String {
        match self {
            Self::MoveTo { x, y } => format!("M {} {}", x, y),
            Self::LineTo { x, y } => format!("L {} {}", x, y),
            Self::CubicBezier {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                format!("C {} {} {} {} {} {}", x1, y1, x2, y2, x, y)
            },
            Self::QuadraticBezier { x1, y1, x, y } => {
                format!("Q {} {} {} {}", x1, y1, x, y)
            },
            Self::Arc {
                rx,
                ry,
                x_axis_rotation,
                large_arc,
                sweep,
                x,
                y,
            } => {
                format!(
                    "A {} {} {} {} {} {} {}",
                    rx,
                    ry,
                    x_axis_rotation,
                    if *large_arc { 1 } else { 0 },
                    if *sweep { 1 } else { 0 },
                    x,
                    y
                )
            },
            Self::ClosePath => "Z".to_string(),
        }
    }
}

/// SVG path element
#[derive(Debug, Clone)]
pub struct SvgPath {
    /// Path commands
    pub commands: Vec<PathCommand>,
    /// Stroke color (RGB hex)
    pub stroke: Option<String>,
    /// Stroke width
    pub stroke_width: f64,
    /// Fill color (RGB hex)
    pub fill: Option<String>,
    /// Fill opacity
    pub fill_opacity: f64,
    /// Stroke opacity
    pub stroke_opacity: f64,
}

impl Default for SvgPath {
    fn default() -> Self {
        Self {
            commands: Vec::new(),
            stroke: Some("#000000".to_string()),
            stroke_width: 1.0,
            fill: None,
            fill_opacity: 1.0,
            stroke_opacity: 1.0,
        }
    }
}

impl SvgPath {
    /// Create new path with commands
    pub fn new(commands: Vec<PathCommand>) -> Self {
        Self {
            commands,
            ..Default::default()
        }
    }

    /// Set stroke color
    pub fn with_stroke(mut self, color: String) -> Self {
        self.stroke = Some(color);
        self
    }

    /// Set stroke width
    pub fn with_stroke_width(mut self, width: f64) -> Self {
        self.stroke_width = width;
        self
    }

    /// Set fill color
    pub fn with_fill(mut self, color: String) -> Self {
        self.fill = Some(color);
        self
    }

    /// Set fill opacity
    pub fn with_fill_opacity(mut self, opacity: f64) -> Self {
        self.fill_opacity = opacity;
        self
    }

    /// Set stroke opacity
    pub fn with_stroke_opacity(mut self, opacity: f64) -> Self {
        self.stroke_opacity = opacity;
        self
    }

    /// Generate SVG path string
    pub fn to_svg(&self) -> String {
        let mut path_data = String::new();
        for cmd in &self.commands {
            if !path_data.is_empty() {
                path_data.push(' ');
            }
            path_data.push_str(&cmd.to_svg());
        }

        let mut attrs = format!(r#"d="{}""#, path_data);

        if let Some(ref stroke) = self.stroke {
            write!(attrs, r#" stroke="{}""#, stroke).unwrap();
        } else {
            attrs.push_str(r#" stroke="none""#);
        }

        write!(attrs, r#" stroke-width="{}""#, self.stroke_width).unwrap();

        if let Some(ref fill) = self.fill {
            write!(attrs, r#" fill="{}""#, fill).unwrap();
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if self.fill_opacity < 1.0 {
            write!(attrs, r#" fill-opacity="{}""#, self.fill_opacity).unwrap();
        }

        if self.stroke_opacity < 1.0 {
            write!(attrs, r#" stroke-opacity="{}""#, self.stroke_opacity).unwrap();
        }

        format!(r#"<path {} />"#, attrs)
    }
}

/// SVG rectangle element
#[derive(Debug, Clone)]
pub struct SvgRect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub fill: Option<String>,
    pub stroke: Option<String>,
    pub stroke_width: f64,
}

impl SvgRect {
    /// Generate SVG rect string
    pub fn to_svg(&self) -> String {
        let mut attrs = format!(
            r#"x="{}" y="{}" width="{}" height="{}""#,
            self.x, self.y, self.width, self.height
        );

        if let Some(ref fill) = self.fill {
            write!(attrs, r#" fill="{}""#, fill).unwrap();
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if let Some(ref stroke) = self.stroke {
            write!(attrs, r#" stroke="{}""#, stroke).unwrap();
            write!(attrs, r#" stroke-width="{}""#, self.stroke_width).unwrap();
        }

        format!(r#"<rect {} />"#, attrs)
    }
}

/// SVG ellipse element
#[derive(Debug, Clone)]
pub struct SvgEllipse {
    pub cx: f64,
    pub cy: f64,
    pub rx: f64,
    pub ry: f64,
    pub fill: Option<String>,
    pub stroke: Option<String>,
    pub stroke_width: f64,
}

impl SvgEllipse {
    /// Generate SVG ellipse string
    pub fn to_svg(&self) -> String {
        let mut attrs = format!(
            r#"cx="{}" cy="{}" rx="{}" ry="{}""#,
            self.cx, self.cy, self.rx, self.ry
        );

        if let Some(ref fill) = self.fill {
            write!(attrs, r#" fill="{}""#, fill).unwrap();
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if let Some(ref stroke) = self.stroke {
            write!(attrs, r#" stroke="{}""#, stroke).unwrap();
            write!(attrs, r#" stroke-width="{}""#, self.stroke_width).unwrap();
        }

        format!(r#"<ellipse {} />"#, attrs)
    }
}

/// SVG text element
#[derive(Debug, Clone)]
pub struct SvgText {
    pub x: f64,
    pub y: f64,
    pub text: String,
    pub font_size: f64,
    pub font_family: Option<String>,
    pub fill: Option<String>,
}

impl SvgText {
    /// Generate SVG text string
    pub fn to_svg(&self) -> String {
        let mut attrs = format!(
            r#"x="{}" y="{}" font-size="{}""#,
            self.x, self.y, self.font_size
        );

        if let Some(ref family) = self.font_family {
            write!(attrs, r#" font-family="{}""#, family).unwrap();
        }

        if let Some(ref fill) = self.fill {
            write!(attrs, r#" fill="{}""#, fill).unwrap();
        }

        format!(r#"<text {}>{}</text>"#, attrs, self.text)
    }
}

/// SVG image element (for embedded raster images)
#[derive(Debug, Clone)]
pub struct SvgImage {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    /// Base64-encoded image data with data URI scheme
    pub href: String,
}

impl SvgImage {
    /// Create from PNG data
    pub fn from_png_data(x: f64, y: f64, width: f64, height: f64, png_data: &[u8]) -> Self {
        use base64::Engine;
        let base64_engine = base64::engine::general_purpose::STANDARD;
        let encoded = base64_engine.encode(png_data);
        let href = format!("data:image/png;base64,{}", encoded);

        Self {
            x,
            y,
            width,
            height,
            href,
        }
    }

    /// Create from JPEG data
    pub fn from_jpeg_data(x: f64, y: f64, width: f64, height: f64, jpeg_data: &[u8]) -> Self {
        use base64::Engine;
        let base64_engine = base64::engine::general_purpose::STANDARD;
        let encoded = base64_engine.encode(jpeg_data);
        let href = format!("data:image/jpeg;base64,{}", encoded);

        Self {
            x,
            y,
            width,
            height,
            href,
        }
    }

    /// Generate SVG image string
    pub fn to_svg(&self) -> String {
        format!(
            r#"<image x="{}" y="{}" width="{}" height="{}" href="{}" />"#,
            self.x, self.y, self.width, self.height, self.href
        )
    }
}

/// SVG element types
#[derive(Debug, Clone)]
pub enum SvgElement {
    Path(SvgPath),
    Rect(SvgRect),
    Ellipse(SvgEllipse),
    Text(SvgText),
    Image(SvgImage),
}

impl SvgElement {
    /// Convert to SVG string
    pub fn to_svg(&self) -> String {
        match self {
            Self::Path(p) => p.to_svg(),
            Self::Rect(r) => r.to_svg(),
            Self::Ellipse(e) => e.to_svg(),
            Self::Text(t) => t.to_svg(),
            Self::Image(i) => i.to_svg(),
        }
    }
}

/// SVG document builder
#[derive(Debug, Clone)]
pub struct SvgBuilder {
    /// Document width
    pub width: f64,
    /// Document height
    pub height: f64,
    /// ViewBox (x, y, width, height)
    pub viewbox: Option<(f64, f64, f64, f64)>,
    /// SVG elements
    pub elements: Vec<SvgElement>,
}

impl SvgBuilder {
    /// Create new SVG builder
    pub fn new(width: f64, height: f64) -> Self {
        Self {
            width,
            height,
            viewbox: None,
            elements: Vec::new(),
        }
    }

    /// Set viewBox
    pub fn with_viewbox(mut self, x: f64, y: f64, width: f64, height: f64) -> Self {
        self.viewbox = Some((x, y, width, height));
        self
    }

    /// Add an element
    pub fn add_element(&mut self, element: SvgElement) {
        self.elements.push(element);
    }

    /// Add a path
    pub fn add_path(&mut self, path: SvgPath) {
        self.elements.push(SvgElement::Path(path));
    }

    /// Add a rectangle
    pub fn add_rect(&mut self, rect: SvgRect) {
        self.elements.push(SvgElement::Rect(rect));
    }

    /// Add an ellipse
    pub fn add_ellipse(&mut self, ellipse: SvgEllipse) {
        self.elements.push(SvgElement::Ellipse(ellipse));
    }

    /// Add text
    pub fn add_text(&mut self, text: SvgText) {
        self.elements.push(SvgElement::Text(text));
    }

    /// Add an embedded image
    pub fn add_image(&mut self, image: SvgImage) {
        self.elements.push(SvgElement::Image(image));
    }

    /// Generate complete SVG document
    pub fn build(&self) -> String {
        let mut svg = String::new();

        // XML declaration
        svg.push_str(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        svg.push('\n');

        // SVG opening tag
        svg.push_str(r#"<svg xmlns="http://www.w3.org/2000/svg" "#);
        write!(svg, r#"width="{}" height="{}""#, self.width, self.height).unwrap();

        if let Some((x, y, w, h)) = self.viewbox {
            write!(svg, r#" viewBox="{} {} {} {}""#, x, y, w, h).unwrap();
        }

        svg.push_str(">\n");

        // Add elements
        for element in &self.elements {
            svg.push_str("  ");
            svg.push_str(&element.to_svg());
            svg.push('\n');
        }

        // SVG closing tag
        svg.push_str("</svg>");

        svg
    }

    /// Build and return as bytes
    pub fn build_bytes(&self) -> Vec<u8> {
        self.build().into_bytes()
    }
}

/// Color conversion utilities
pub mod color {
    /// Convert RGB color to hex string
    pub fn rgb_to_hex(r: u8, g: u8, b: u8) -> String {
        format!("#{:02X}{:02X}{:02X}", r, g, b)
    }

    /// Convert COLORREF (Windows color format) to hex string
    pub fn colorref_to_hex(colorref: u32) -> String {
        let r = (colorref & 0xFF) as u8;
        let g = ((colorref >> 8) & 0xFF) as u8;
        let b = ((colorref >> 16) & 0xFF) as u8;
        rgb_to_hex(r, g, b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_svg_path() {
        let path = SvgPath::new(vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo { x: 100.0, y: 100.0 },
        ])
        .with_stroke("#000000".to_string());

        let svg = path.to_svg();
        assert!(svg.contains("M 0 0 L 100 100"));
    }

    #[test]
    fn test_svg_builder() {
        let mut builder = SvgBuilder::new(100.0, 100.0);
        builder.add_rect(SvgRect {
            x: 10.0,
            y: 10.0,
            width: 80.0,
            height: 80.0,
            fill: Some("#FF0000".to_string()),
            stroke: None,
            stroke_width: 0.0,
        });

        let svg = builder.build();
        assert!(svg.contains("<svg"));
        assert!(svg.contains("</svg>"));
        assert!(svg.contains("<rect"));
    }

    #[test]
    fn test_color_conversion() {
        assert_eq!(color::rgb_to_hex(255, 0, 0), "#FF0000");
        assert_eq!(color::colorref_to_hex(0x0000FF), "#FF0000");
    }
}
