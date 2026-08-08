// SVG generation module
//
// Provides high-performance SVG generation from vector graphics operations

const INERT_IMAGE_HREF: &str = "data:,";

fn write_compact_xml_escaped(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{10ffff}' => output.push(ch),
            _ => output.push('\u{fffd}'),
        }
    }
}

fn push_string_attr(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    write_compact_xml_escaped(output, value);
    output.push('"');
}

fn safe_paint(value: &str) -> &str {
    if value.contains('\\') || contains_css_reference_function(value) {
        let trimmed = value.trim();
        if is_local_fragment_paint(trimmed) {
            trimmed
        } else {
            "none"
        }
    } else {
        value
    }
}

fn contains_css_reference_function(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index + 3 <= bytes.len() {
        if bytes[index..index + 3].eq_ignore_ascii_case(b"url")
            || bytes[index..index + 3].eq_ignore_ascii_case(b"var")
        {
            let mut after = index + 3;
            while after < bytes.len() && bytes[after].is_ascii_whitespace() {
                after += 1;
            }
            if bytes.get(after) == Some(&b'(') {
                return true;
            }
        }
        index += 1;
    }
    false
}

fn is_local_fragment_paint(value: &str) -> bool {
    let Some(fragment) = value
        .strip_prefix("url(#")
        .and_then(|value| value.strip_suffix(')'))
    else {
        return false;
    };
    !fragment.is_empty()
        && fragment
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':'))
}

fn safe_image_href(value: &str) -> &str {
    let payload = value
        .strip_prefix("data:image/png;base64,")
        .or_else(|| value.strip_prefix("data:image/jpeg;base64,"));
    if payload.is_some_and(valid_base64_payload) {
        value
    } else {
        INERT_IMAGE_HREF
    }
}

fn valid_base64_payload(payload: &str) -> bool {
    if payload.len() % 4 != 0 || !payload.is_ascii() {
        return false;
    }
    let padding = payload
        .bytes()
        .rev()
        .take_while(|&byte| byte == b'=')
        .count();
    padding <= 2
        && payload[..payload.len().saturating_sub(padding)]
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/'))
        && payload[payload.len().saturating_sub(padding)..]
            .bytes()
            .all(|byte| byte == b'=')
}

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
    /// Stroke dasharray pattern
    pub stroke_dasharray: Option<String>,
    /// Stroke linecap style
    pub stroke_linecap: Option<String>,
    /// Stroke linejoin style
    pub stroke_linejoin: Option<String>,
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
            stroke_dasharray: None,
            stroke_linecap: None,
            stroke_linejoin: None,
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

    /// Set stroke dasharray pattern
    pub fn with_stroke_dasharray(mut self, dasharray: String) -> Self {
        self.stroke_dasharray = Some(dasharray);
        self
    }

    /// Set stroke linecap style
    pub fn with_stroke_linecap(mut self, linecap: String) -> Self {
        self.stroke_linecap = Some(linecap);
        self
    }

    /// Set stroke linejoin style
    pub fn with_stroke_linejoin(mut self, linejoin: String) -> Self {
        self.stroke_linejoin = Some(linejoin);
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
            push_string_attr(&mut attrs, "stroke", safe_paint(stroke));
        } else {
            attrs.push_str(r#" stroke="none""#);
        }

        attrs.push_str(&format!(r#" stroke-width="{}""#, self.stroke_width));

        if let Some(ref fill) = self.fill {
            push_string_attr(&mut attrs, "fill", safe_paint(fill));
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if self.fill_opacity < 1.0 {
            attrs.push_str(&format!(r#" fill-opacity="{}""#, self.fill_opacity));
        }

        if self.stroke_opacity < 1.0 {
            attrs.push_str(&format!(r#" stroke-opacity="{}""#, self.stroke_opacity));
        }

        if let Some(ref dasharray) = self.stroke_dasharray {
            push_string_attr(&mut attrs, "stroke-dasharray", dasharray);
        }

        if let Some(ref linecap) = self.stroke_linecap {
            push_string_attr(&mut attrs, "stroke-linecap", linecap);
        }

        if let Some(ref linejoin) = self.stroke_linejoin {
            push_string_attr(&mut attrs, "stroke-linejoin", linejoin);
        }

        format!(r#"<path {}/>"#, attrs)
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
            push_string_attr(&mut attrs, "fill", safe_paint(fill));
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if let Some(ref stroke) = self.stroke {
            push_string_attr(&mut attrs, "stroke", safe_paint(stroke));
            attrs.push_str(&format!(r#" stroke-width="{}""#, self.stroke_width));
        }

        format!(r#"<rect {}/>"#, attrs)
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
            push_string_attr(&mut attrs, "fill", safe_paint(fill));
        } else {
            attrs.push_str(r#" fill="none""#);
        }

        if let Some(ref stroke) = self.stroke {
            push_string_attr(&mut attrs, "stroke", safe_paint(stroke));
            attrs.push_str(&format!(r#" stroke-width="{}""#, self.stroke_width));
        }

        format!(r#"<ellipse {}/>"#, attrs)
    }
}

/// SVG text element with full WMF support
#[derive(Debug, Clone)]
pub struct SvgText {
    pub x: f64,
    pub y: f64,
    pub text: String,
    pub font_size: f64,
    pub font_family: Option<String>,
    pub fill: Option<String>,
    /// Font weight (400 = normal, 700 = bold)
    pub font_weight: Option<u16>,
    /// Italic style
    pub italic: bool,
    /// Underline decoration
    pub underline: bool,
    /// Strikethrough decoration
    pub strikethrough: bool,
    /// Rotation angle in degrees
    pub rotation: Option<f64>,
    /// Transform matrix (6 values: a b c d e f)
    pub transform: Option<[f64; 6]>,
}

impl Default for SvgText {
    fn default() -> Self {
        Self {
            x: 0.0,
            y: 0.0,
            text: String::new(),
            font_size: 12.0,
            font_family: None,
            fill: Some("#000000".to_string()),
            font_weight: None,
            italic: false,
            underline: false,
            strikethrough: false,
            rotation: None,
            transform: None,
        }
    }
}

impl SvgText {
    /// Create new text element
    pub fn new(x: f64, y: f64, text: String, font_size: f64) -> Self {
        Self {
            x,
            y,
            text,
            font_size,
            ..Default::default()
        }
    }

    /// Set font family
    pub fn with_font_family(mut self, family: String) -> Self {
        self.font_family = Some(family);
        self
    }

    /// Set fill color
    pub fn with_fill(mut self, fill: String) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set font weight
    pub fn with_font_weight(mut self, weight: u16) -> Self {
        self.font_weight = Some(weight);
        self
    }

    /// Set italic style
    pub fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Set underline
    pub fn with_underline(mut self, underline: bool) -> Self {
        self.underline = underline;
        self
    }

    /// Set strikethrough
    pub fn with_strikethrough(mut self, strikethrough: bool) -> Self {
        self.strikethrough = strikethrough;
        self
    }

    /// Set rotation angle in degrees
    pub fn with_rotation(mut self, degrees: f64) -> Self {
        self.rotation = Some(degrees);
        self
    }

    /// Set transform matrix
    pub fn with_transform(mut self, matrix: [f64; 6]) -> Self {
        self.transform = Some(matrix);
        self
    }

    /// Generate SVG text string
    pub fn to_svg(&self) -> String {
        let mut attrs = format!(r#"x="{}" y="{}""#, self.x, self.y);

        if let Some(ref family) = self.font_family {
            push_string_attr(&mut attrs, "font-family", family);
        }

        attrs.push_str(&format!(r#" font-size="{}px""#, self.font_size));

        if let Some(weight) = self.font_weight {
            attrs.push_str(&format!(r#" font-weight="{}""#, weight));
        }

        if self.italic {
            attrs.push_str(r#" font-style="italic""#);
        }

        if let Some(ref fill) = self.fill {
            push_string_attr(&mut attrs, "fill", safe_paint(fill));
        }

        // Build text decorations
        let mut decorations = Vec::new();
        if self.underline {
            decorations.push("underline");
        }
        if self.strikethrough {
            decorations.push("line-through");
        }
        if !decorations.is_empty() {
            push_string_attr(&mut attrs, "text-decoration", &decorations.join(" "));
        }

        // Build transform attribute
        if let Some(matrix) = self.transform {
            // Transform matrix: matrix(a b c d e f)
            attrs.push_str(&format!(
                r#" transform="matrix({} {} {} {} {} {})""#,
                matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]
            ));
        } else if let Some(rotation) = self.rotation {
            // Simple rotation around (x, y)
            attrs.push_str(&format!(
                r#" transform="rotate({} {} {})""#,
                rotation, self.x, self.y
            ));
        }

        let mut escaped_text = String::with_capacity(self.text.len());
        write_compact_xml_escaped(&mut escaped_text, &self.text);

        format!(r#"<text {}>{}</text>"#, attrs, escaped_text)
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
        let mut output = format!(
            r#"<image x="{}" y="{}" width="{}" height="{}""#,
            self.x, self.y, self.width, self.height
        );
        push_string_attr(&mut output, "href", safe_image_href(&self.href));
        output.push_str("/>");
        output
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

        // SVG opening tag
        svg.push_str(r#"<svg xmlns="http://www.w3.org/2000/svg" "#);
        svg.push_str(&format!(
            r#"width="{}" height="{}""#,
            self.width, self.height
        ));

        if let Some((x, y, w, h)) = self.viewbox {
            svg.push_str(&format!(r#" viewBox="{} {} {} {}""#, x, y, w, h));
        }

        svg.push('>');

        // Add elements
        for element in &self.elements {
            svg.push_str(&element.to_svg());
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
    fn element_helpers_emit_compact_exact_bytes() {
        let path = SvgPath::new(vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo { x: 100.0, y: 100.0 },
        ])
        .with_stroke("#000000".to_string());

        assert_eq!(
            path.to_svg().as_bytes(),
            br##"<path d="M 0 0 L 100 100" stroke="#000000" stroke-width="1" fill="none"/>"##
        );

        let ellipse = SvgEllipse {
            cx: 10.0,
            cy: 20.0,
            rx: 3.0,
            ry: 4.0,
            fill: None,
            stroke: Some("#00FF00".to_string()),
            stroke_width: 2.0,
        };
        assert_eq!(
            ellipse.to_svg().as_bytes(),
            br##"<ellipse cx="10" cy="20" rx="3" ry="4" fill="none" stroke="#00FF00" stroke-width="2"/>"##
        );

        let image = SvgImage::from_png_data(1.0, 2.0, 3.0, 4.0, &[0, 1, 2]);
        assert_eq!(
            image.to_svg().as_bytes(),
            br#"<image x="1" y="2" width="3" height="4" href="data:image/png;base64,AAEC"/>"#
        );

        let text = SvgText::new(1.0, 2.0, "a<b&c".to_string(), 12.0);
        assert_eq!(
            text.to_svg().as_bytes(),
            br##"<text x="1" y="2" font-size="12px" fill="#000000">a&lt;b&amp;c</text>"##
        );
    }

    #[test]
    fn adversarial_strings_are_escaped_and_external_references_are_inert() {
        let path = SvgPath::new(vec![PathCommand::MoveTo { x: 0.0, y: 0.0 }])
            .with_stroke("\" onload=\"alert(1)&<>\n".to_string())
            .with_fill("URL ( https://example.invalid/pixel )".to_string())
            .with_stroke_dasharray("1\"/><script>&\n".to_string())
            .with_stroke_linecap("round\0\"".to_string());
        assert_eq!(
            path.to_svg().as_bytes(),
            r##"<path d="M 0 0" stroke="&quot; onload=&quot;alert(1)&amp;&lt;&gt;&#xA;" stroke-width="1" fill="none" stroke-dasharray="1&quot;/&gt;&lt;script&gt;&amp;&#xA;" stroke-linecap="round�&quot;"/>"##.as_bytes()
        );

        let text = SvgText::new(1.0, 2.0, "<&\0\n".to_string(), 12.0)
            .with_font_family("x';fill:url(https://example.invalid);\"&\n".to_string())
            .with_fill("var(--host-paint)".to_string());
        assert_eq!(
            text.to_svg().as_bytes(),
            r##"<text x="1" y="2" font-family="x&apos;;fill:url(https://example.invalid);&quot;&amp;&#xA;" font-size="12px" fill="none">&lt;&amp;�&#xA;</text>"##.as_bytes()
        );

        let image = SvgImage {
            x: 1.0,
            y: 2.0,
            width: 3.0,
            height: 4.0,
            href: "https://example.invalid/x\" onload=\"alert(1)".to_string(),
        };
        assert_eq!(
            image.to_svg().as_bytes(),
            br#"<image x="1" y="2" width="3" height="4" href="data:,"/>"#
        );

        assert_eq!(safe_paint("url(#safe-id)"), "url(#safe-id)");
        assert_eq!(safe_paint("u\\72l(https://example.invalid)"), "none");
    }

    #[test]
    fn adversarial_document_is_well_formed_single_line_xml() {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut builder = SvgBuilder::new(10.0, 10.0);
        builder.add_text(
            SvgText::new(
                0.0,
                1.0,
                "</text><script>alert(1)</script>\n".to_string(),
                1.0,
            )
            .with_font_family("\" onload=\"alert(1)".to_string()),
        );
        builder.add_image(SvgImage {
            x: 0.0,
            y: 0.0,
            width: 1.0,
            height: 1.0,
            href: "javascript:alert(1)".to_string(),
        });
        let svg = builder.build();
        assert!(!svg.as_bytes().contains(&b'\n'));
        assert!(!svg.contains("javascript:"));

        let mut reader = Reader::from_str(&svg);
        let mut starts = Vec::new();
        loop {
            match reader.read_event().expect("adversarial SVG must parse") {
                Event::Start(event) | Event::Empty(event) => {
                    starts.push(String::from_utf8_lossy(event.name().as_ref()).into_owned());
                },
                Event::Eof => break,
                _ => {},
            }
        }
        assert_eq!(starts, ["svg", "text", "image"]);
    }

    #[test]
    fn builder_emits_single_line_compact_exact_bytes() {
        let mut builder = SvgBuilder::new(100.0, 100.0).with_viewbox(0.0, 0.0, 100.0, 100.0);
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
        assert_eq!(
            svg.as_bytes(),
            br##"<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><rect x="10" y="10" width="80" height="80" fill="#FF0000"/></svg>"##
        );
        assert!(!svg.as_bytes().contains(&b'\n'));
        assert!(!svg.as_bytes().windows(3).any(|window| window == b" />"));
    }

    #[test]
    fn test_color_conversion() {
        assert_eq!(color::rgb_to_hex(255, 0, 0), "#FF0000");
        assert_eq!(color::colorref_to_hex(0x0000FF), "#FF0000");
    }
}
