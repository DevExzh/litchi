/// Drawing objects support for DOCX documents.
///
/// This module provides structures and functions for extracting drawing objects
/// (shapes, text boxes, diagrams) from Word documents. Drawing objects are embedded
/// within `<w:drawing>` elements in paragraphs, using DrawingML (DML) markup.
///
/// # Architecture
///
/// - `DrawingObject`: Base type for all drawing objects
/// - `Shape`: A shape with geometry, fill, and text
/// - `TextBox`: A text box containing formatted text
/// - `ShapeType`: Enumeration of standard shape types
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Extract all drawing objects from the document
/// for para in doc.paragraphs()? {
///     for drawing in para.drawing_objects()? {
///         match drawing.shape_type() {
///             ShapeType::TextBox => {
///                 println!("Text box: {}", drawing.text());
///             }
///             ShapeType::Rectangle => {
///                 println!("Rectangle: {}x{} EMUs",
///                     drawing.width_emu(),
///                     drawing.height_emu()
///                 );
///             }
///             _ => {}
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;

/// Type of shape.
///
/// DrawingML supports numerous predefined shape types. This enum covers
/// the most commonly used shapes in Word documents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeType {
    /// Rectangle shape
    Rectangle,
    /// Rounded rectangle
    RoundRectangle,
    /// Ellipse/circle shape
    Ellipse,
    /// Triangle shape
    Triangle,
    /// Right triangle
    RightTriangle,
    /// Parallelogram
    Parallelogram,
    /// Trapezoid
    Trapezoid,
    /// Diamond shape
    Diamond,
    /// Pentagon
    Pentagon,
    /// Hexagon
    Hexagon,
    /// Octagon
    Octagon,
    /// Star (5-pointed)
    Star5,
    /// Star (6-pointed)
    Star6,
    /// Arrow (right)
    ArrowRight,
    /// Arrow (left)
    ArrowLeft,
    /// Arrow (up)
    ArrowUp,
    /// Arrow (down)
    ArrowDown,
    /// Text box (rectangular with text)
    TextBox,
    /// Line
    Line,
    /// Custom or unknown shape type
    Custom(String),
}

impl ShapeType {
    /// Parse shape type from preset geometry string.
    ///
    /// # Arguments
    ///
    /// * `prst` - The preset geometry value (e.g., "rect", "ellipse")
    #[inline]
    pub fn from_preset(prst: &str) -> Self {
        match prst {
            "rect" => Self::Rectangle,
            "roundRect" => Self::RoundRectangle,
            "ellipse" => Self::Ellipse,
            "triangle" => Self::Triangle,
            "rtTriangle" => Self::RightTriangle,
            "parallelogram" => Self::Parallelogram,
            "trapezoid" => Self::Trapezoid,
            "diamond" => Self::Diamond,
            "pentagon" => Self::Pentagon,
            "hexagon" => Self::Hexagon,
            "octagon" => Self::Octagon,
            "star5" => Self::Star5,
            "star6" => Self::Star6,
            "rightArrow" => Self::ArrowRight,
            "leftArrow" => Self::ArrowLeft,
            "upArrow" => Self::ArrowUp,
            "downArrow" => Self::ArrowDown,
            "textBox" => Self::TextBox,
            "line" => Self::Line,
            _ => Self::Custom(prst.to_string()),
        }
    }

    /// Get the preset string for this shape type.
    #[inline]
    pub fn to_preset(&self) -> &str {
        match self {
            Self::Rectangle => "rect",
            Self::RoundRectangle => "roundRect",
            Self::Ellipse => "ellipse",
            Self::Triangle => "triangle",
            Self::RightTriangle => "rtTriangle",
            Self::Parallelogram => "parallelogram",
            Self::Trapezoid => "trapezoid",
            Self::Diamond => "diamond",
            Self::Pentagon => "pentagon",
            Self::Hexagon => "hexagon",
            Self::Octagon => "octagon",
            Self::Star5 => "star5",
            Self::Star6 => "star6",
            Self::ArrowRight => "rightArrow",
            Self::ArrowLeft => "leftArrow",
            Self::ArrowUp => "upArrow",
            Self::ArrowDown => "downArrow",
            Self::TextBox => "textBox",
            Self::Line => "line",
            Self::Custom(s) => s,
        }
    }
}

/// A drawing object (shape, text box, or diagram) in a Word document.
///
/// Represents a DrawingML object within a `<w:drawing>` element. Drawing objects
/// can contain geometry, fill, outline, effects, and text.
///
/// # Performance
///
/// Drawing metadata is stored inline for fast access. Text content is parsed
/// on demand to minimize memory usage.
///
/// # Field Ordering
///
/// Fields are ordered to maximize CPU cache line utilization:
/// - Strings (24 bytes each on 64-bit systems)
/// - i64/u64 values (8 bytes each)
/// - Enums and smaller types
#[derive(Debug, Clone)]
pub struct DrawingObject {
    /// Shape name/title
    name: String,

    /// Shape description/alt text
    description: String,

    /// Text content within the shape (if any)
    text: String,

    /// Width in EMUs (English Metric Units, 1 inch = 914400 EMUs)
    width_emu: i64,

    /// Height in EMUs
    height_emu: i64,

    /// X position in EMUs (distance from anchor)
    x_emu: i64,

    /// Y position in EMUs (distance from anchor)
    y_emu: i64,

    /// Shape type
    shape_type: ShapeType,

    /// Whether this is an inline shape (vs. anchored/floating)
    is_inline: bool,
}

impl DrawingObject {
    /// Create a new DrawingObject.
    ///
    /// # Arguments
    ///
    /// * `name` - Shape name/title
    /// * `description` - Shape description/alt text
    /// * `width_emu` - Width in EMUs
    /// * `height_emu` - Height in EMUs
    /// * `shape_type` - Type of shape
    #[inline]
    pub fn new(
        name: String,
        description: String,
        width_emu: i64,
        height_emu: i64,
        shape_type: ShapeType,
    ) -> Self {
        Self {
            name,
            description,
            text: String::new(),
            width_emu,
            height_emu,
            x_emu: 0,
            y_emu: 0,
            shape_type,
            is_inline: true,
        }
    }

    /// Get the shape name/title.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Get the shape description/alt text.
    #[inline]
    pub fn description(&self) -> &str {
        &self.description
    }

    /// Get the text content within the shape.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content within the shape.
    #[inline]
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Get the width in EMUs (English Metric Units).
    #[inline]
    pub fn width_emu(&self) -> i64 {
        self.width_emu
    }

    /// Get the height in EMUs.
    #[inline]
    pub fn height_emu(&self) -> i64 {
        self.height_emu
    }

    /// Get the X position in EMUs.
    #[inline]
    pub fn x_emu(&self) -> i64 {
        self.x_emu
    }

    /// Get the Y position in EMUs.
    #[inline]
    pub fn y_emu(&self) -> i64 {
        self.y_emu
    }

    /// Set the position in EMUs.
    #[inline]
    pub fn set_position(&mut self, x_emu: i64, y_emu: i64) {
        self.x_emu = x_emu;
        self.y_emu = y_emu;
    }

    /// Get the width in pixels (assuming 96 DPI).
    #[inline]
    pub fn width_px(&self) -> u32 {
        ((self.width_emu as f64) * 96.0 / 914400.0) as u32
    }

    /// Get the height in pixels (assuming 96 DPI).
    #[inline]
    pub fn height_px(&self) -> u32 {
        ((self.height_emu as f64) * 96.0 / 914400.0) as u32
    }

    /// Get the width in points.
    #[inline]
    pub fn width_pt(&self) -> f64 {
        (self.width_emu as f64) / 12700.0
    }

    /// Get the height in points.
    #[inline]
    pub fn height_pt(&self) -> f64 {
        (self.height_emu as f64) / 12700.0
    }

    /// Get the shape type.
    #[inline]
    pub fn shape_type(&self) -> &ShapeType {
        &self.shape_type
    }

    /// Check if this is an inline shape (vs. anchored/floating).
    #[inline]
    pub fn is_inline(&self) -> bool {
        self.is_inline
    }

    /// Set whether this is an inline shape.
    #[inline]
    pub fn set_inline(&mut self, is_inline: bool) {
        self.is_inline = is_inline;
    }
}

/// Parse drawing objects from paragraph XML.
///
/// Extracts all `<w:drawing>` elements containing shapes, text boxes, and other
/// DrawingML objects from the paragraph XML.
///
/// # Arguments
///
/// * `xml_bytes` - The raw XML bytes of the paragraph
///
/// # Performance
///
/// Uses streaming XML parsing with pre-allocated SmallVec for efficient
/// storage of typically small drawing collections.
///
/// # Example XML Structure
///
/// ```xml
/// <w:drawing>
///   <wp:anchor>  <!-- or wp:inline for inline shapes -->
///     <wp:extent cx="914400" cy="914400"/>
///     <wp:docPr name="Shape 1" descr="Description"/>
///     <a:graphic>
///       <a:graphicData>
///         <wps:wsp>
///           <wps:cNvSpPr/>
///           <wps:spPr>
///             <a:xfrm>
///               <a:off x="0" y="0"/>
///               <a:ext cx="914400" cy="914400"/>
///             </a:xfrm>
///             <a:prstGeom prst="rect"/>
///           </wps:spPr>
///           <wps:txbx>
///             <w:txbxContent>
///               <w:p><w:r><w:t>Text content</w:t></w:r></w:p>
///             </w:txbxContent>
///           </wps:txbx>
///         </wps:wsp>
///       </a:graphicData>
///     </a:graphic>
///   </wp:anchor>
/// </w:drawing>
/// ```
pub(crate) fn parse_drawing_objects(xml_bytes: &[u8]) -> Result<SmallVec<[DrawingObject; 4]>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    // Use SmallVec for efficient storage of typically small drawing collections
    let mut drawings = SmallVec::new();

    // State tracking for parsing
    let mut in_drawing = false;
    let mut in_shape = false;
    let mut in_text_box = false;
    let mut in_text_content = false;

    // Drawing attributes being built
    let mut width_emu: i64 = 914400; // Default 1 inch
    let mut height_emu: i64 = 914400;
    let mut x_emu: i64 = 0;
    let mut y_emu: i64 = 0;
    let mut description = String::new();
    let mut name = String::new();
    let mut shape_type = ShapeType::Rectangle;
    let mut is_inline = true;
    let mut text_content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"drawing" => {
                        in_drawing = true;
                        // Reset state for new drawing
                        width_emu = 914400;
                        height_emu = 914400;
                        x_emu = 0;
                        y_emu = 0;
                        description.clear();
                        name.clear();
                        shape_type = ShapeType::Rectangle;
                        is_inline = true;
                        text_content.clear();
                    },
                    b"inline" if in_drawing => {
                        is_inline = true;
                    },
                    b"anchor" if in_drawing => {
                        is_inline = false;
                    },
                    b"extent" if in_drawing => {
                        // Parse width and height from extent element
                        // <wp:extent cx="914400" cy="914400"/>
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"cx" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        width_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                b"cy" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        height_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"off" if in_drawing => {
                        // Parse position from offset element
                        // <a:off x="0" y="0"/>
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"x" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        x_emu = s.parse().unwrap_or(0);
                                    }
                                },
                                b"y" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        y_emu = s.parse().unwrap_or(0);
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"ext" if in_drawing => {
                        // Alternative extent element (in xfrm)
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"cx" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        width_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                b"cy" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        height_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"docPr" if in_drawing => {
                        // Parse name and description from docPr element
                        // <wp:docPr id="1" name="Shape" descr="Description"/>
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        name = s.to_string();
                                    }
                                },
                                b"descr" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        description = s.to_string();
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"wsp" if in_drawing => {
                        // WordprocessingShape element
                        in_shape = true;
                    },
                    b"prstGeom" if in_shape => {
                        // Parse preset geometry (shape type)
                        // <a:prstGeom prst="rect"/>
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"prst"
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                shape_type = ShapeType::from_preset(s);
                            }
                        }
                    },
                    b"txbx" if in_shape => {
                        // Text box element
                        in_text_box = true;
                    },
                    b"txbxContent" if in_text_box => {
                        // Text box content
                        in_text_content = true;
                    },
                    b"t" if in_text_content => {
                        // Text element - will be captured in Text event
                    },
                    _ => {},
                }
            },
            Ok(Event::Text(e)) if in_text_content => {
                // Extract text content from text box
                if let Ok(text) = std::str::from_utf8(e.as_ref()) {
                    text_content.push_str(text);
                }
            },
            Ok(Event::End(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"drawing" => {
                        // Finished parsing a drawing object
                        in_drawing = false;

                        // Create and add the drawing object
                        let mut drawing = DrawingObject::new(
                            name.clone(),
                            description.clone(),
                            width_emu,
                            height_emu,
                            shape_type.clone(),
                        );
                        drawing.set_position(x_emu, y_emu);
                        drawing.set_inline(is_inline);
                        drawing.set_text(text_content.clone());

                        drawings.push(drawing);
                    },
                    b"wsp" => {
                        in_shape = false;
                    },
                    b"txbx" => {
                        in_text_box = false;
                    },
                    b"txbxContent" => {
                        in_text_content = false;
                    },
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
    }

    Ok(drawings)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shape_type_parsing() {
        assert_eq!(ShapeType::from_preset("rect"), ShapeType::Rectangle);
        assert_eq!(ShapeType::from_preset("ellipse"), ShapeType::Ellipse);
        assert_eq!(ShapeType::from_preset("rightArrow"), ShapeType::ArrowRight);
        assert_eq!(ShapeType::from_preset("textBox"), ShapeType::TextBox);

        if let ShapeType::Custom(s) = ShapeType::from_preset("customShape") {
            assert_eq!(s, "customShape");
        } else {
            panic!("Expected Custom shape type");
        }
    }

    #[test]
    fn test_drawing_object_dimensions() {
        let drawing = DrawingObject::new(
            "Shape 1".to_string(),
            "Test shape".to_string(),
            914400,  // 1 inch
            1828800, // 2 inches
            ShapeType::Rectangle,
        );

        assert_eq!(drawing.width_emu(), 914400);
        assert_eq!(drawing.height_emu(), 1828800);
        assert_eq!(drawing.width_px(), 96); // 1 inch at 96 DPI
        assert_eq!(drawing.height_px(), 192); // 2 inches at 96 DPI
        assert!((drawing.width_pt() - 72.0).abs() < 0.1);
        assert!((drawing.height_pt() - 144.0).abs() < 0.1);
    }

    #[test]
    fn test_parse_drawing_objects_empty() {
        let xml = b"<w:p><w:r><w:t>Text only</w:t></w:r></w:p>";
        let drawings = parse_drawing_objects(xml).unwrap();
        assert_eq!(drawings.len(), 0);
    }

    #[test]
    fn test_parse_drawing_object_simple() {
        let xml = br#"<w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="1000000" cy="2000000"/>
                        <wp:docPr id="1" name="MyShape" descr="Test shape"/>
                        <a:graphic>
                            <a:graphicData>
                                <wps:wsp>
                                    <wps:spPr>
                                        <a:prstGeom prst="rect"/>
                                    </wps:spPr>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>"#;

        let drawings = parse_drawing_objects(xml).unwrap();
        assert_eq!(drawings.len(), 1);

        let drawing = &drawings[0];
        assert_eq!(drawing.name(), "MyShape");
        assert_eq!(drawing.description(), "Test shape");
        assert_eq!(drawing.width_emu(), 1000000);
        assert_eq!(drawing.height_emu(), 2000000);
        assert_eq!(drawing.shape_type(), &ShapeType::Rectangle);
        assert!(drawing.is_inline());
    }

    #[test]
    fn test_parse_text_box() {
        let xml = br#"<w:p>
            <w:r>
                <w:drawing>
                    <wp:anchor>
                        <wp:extent cx="1000000" cy="1000000"/>
                        <wp:docPr name="TextBox1"/>
                        <wps:wsp>
                            <wps:spPr>
                                <a:prstGeom prst="textBox"/>
                            </wps:spPr>
                            <wps:txbx>
                                <w:txbxContent>
                                    <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
                                </w:txbxContent>
                            </wps:txbx>
                        </wps:wsp>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>"#;

        let drawings = parse_drawing_objects(xml).unwrap();
        assert_eq!(drawings.len(), 1);

        let drawing = &drawings[0];
        assert_eq!(drawing.name(), "TextBox1");
        assert_eq!(drawing.shape_type(), &ShapeType::TextBox);
        assert_eq!(drawing.text(), "Hello World");
        assert!(!drawing.is_inline()); // anchor = not inline
    }
}
