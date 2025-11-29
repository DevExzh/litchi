/// Base shape types for PowerPoint presentations.
use crate::ooxml::error::Result;
use crate::ooxml::pptx::shapes::textframe::TextFrame;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Shape type enumeration.
///
/// Indicates what kind of shape this is.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeType {
    /// A text shape (p:sp)
    Shape,
    /// A picture shape (p:pic)
    Picture,
    /// A graphic frame containing a table or chart (p:graphicFrame)
    GraphicFrame,
    /// A group shape (p:grpSp)
    GroupShape,
    /// A connector shape (p:cxnSp)
    Connector,
    /// Unknown or unsupported shape type
    Unknown,
}

/// Base shape containing common properties.
///
/// Provides access to position, size, name, and other properties
/// common to all shapes.
///
/// # Examples
///
/// ```rust,ignore
/// if let Some(shape) = shapes.get(0) {
///     println!("Shape: {}", shape.name());
///     println!("Position: ({}, {})", shape.left(), shape.top());
///     println!("Size: {}x{}", shape.width(), shape.height());
/// }
/// ```
#[derive(Debug, Clone)]
pub struct BaseShape {
    /// Raw XML bytes for this shape
    xml_bytes: Vec<u8>,
    /// Shape type
    shape_type: ShapeType,
    /// Shape name (cached)
    name: Option<String>,
    /// Position and size (cached)
    geometry: Option<ShapeGeometry>,
}

/// Shape geometry (position and size).
#[derive(Debug, Clone, Copy)]
struct ShapeGeometry {
    /// X position in EMUs
    x: i64,
    /// Y position in EMUs
    y: i64,
    /// Width in EMUs
    cx: i64,
    /// Height in EMUs
    cy: i64,
}

impl BaseShape {
    /// Create a new BaseShape from XML bytes and shape type.
    pub fn new(xml_bytes: Vec<u8>, shape_type: ShapeType) -> Self {
        Self {
            xml_bytes,
            shape_type,
            name: None,
            geometry: None,
        }
    }

    /// Get the shape type.
    #[inline]
    pub fn shape_type(&self) -> &ShapeType {
        &self.shape_type
    }

    /// Get the shape name.
    ///
    /// Returns the name from the `<p:cNvPr>` element.
    pub fn name(&mut self) -> Result<String> {
        if let Some(ref name) = self.name {
            return Ok(name.clone());
        }

        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"cNvPr" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                let name =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                self.name = Some(name.clone());
                                return Ok(name);
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Get the X position (left edge) in EMUs.
    pub fn left(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self.geometry.unwrap().x)
    }

    /// Get the Y position (top edge) in EMUs.
    pub fn top(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self.geometry.unwrap().y)
    }

    /// Get the width in EMUs.
    pub fn width(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self.geometry.unwrap().cx)
    }

    /// Get the height in EMUs.
    pub fn height(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self.geometry.unwrap().cy)
    }

    /// Check if this shape is a placeholder.
    pub fn is_placeholder(&self) -> bool {
        // Look for <p:ph> element
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"ph" {
                        return true;
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        false
    }

    /// Get the placeholder type if this shape is a placeholder.
    ///
    /// Returns the type attribute value from the `<p:ph>` element,
    /// such as "title", "body", "ctrTitle", "subTitle", "dt", "ftr", "sldNum", etc.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// if shape.is_placeholder() {
    ///     if let Ok(ph_type) = shape.placeholder_type() {
    ///         println!("Placeholder type: {}", ph_type);
    ///     }
    /// }
    /// ```
    pub fn placeholder_type(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"ph" {
                        // Look for the type attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"type" {
                                return std::str::from_utf8(&attr.value)
                                    .map(|s| s.to_string())
                                    .map_err(|e| {
                                        crate::ooxml::error::OoxmlError::Xml(e.to_string())
                                    });
                            }
                        }
                        // If no type attribute, it's usually a body placeholder
                        return Ok("body".to_string());
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::ooxml::error::OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Check if this shape has a text frame.
    pub fn has_text_frame(&self) -> bool {
        self.shape_type == ShapeType::Shape
    }

    /// Check if this shape contains a table.
    pub fn has_table(&self) -> bool {
        self.shape_type == ShapeType::GraphicFrame && self.contains_table_marker()
    }

    /// Internal helper to check for table marker in XML.
    fn contains_table_marker(&self) -> bool {
        let xml_str = String::from_utf8_lossy(&self.xml_bytes);
        xml_str.contains("a:tbl") || xml_str.contains("<a:tbl")
    }

    /// Ensure geometry is parsed and cached.
    fn ensure_geometry(&mut self) -> Result<()> {
        if self.geometry.is_some() {
            return Ok(());
        }

        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut x = 0;
        let mut y = 0;
        let mut cx = 0;
        let mut cy = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"off" {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"x" => {
                                    x = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                b"y" => {
                                    y = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                _ => {},
                            }
                        }
                    } else if tag_name.as_ref() == b"ext" {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"cx" => {
                                    cx = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                b"cy" => {
                                    cy = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                _ => {},
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        self.geometry = Some(ShapeGeometry { x, y, cx, cy });
        Ok(())
    }

    /// Get raw XML bytes.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml_bytes
    }

    /// Extract text content from this shape if it has any.
    ///
    /// Returns None if the shape doesn't contain text (e.g., pictures without text).
    pub fn text(&self) -> Result<Option<String>> {
        // Only text shapes have text frames
        if !self.has_text_frame() {
            return Ok(None);
        }

        // Parse text from the shape using TextFrame
        match TextFrame::from_xml(&self.xml_bytes) {
            Ok(tf) => Ok(Some(tf.text()?)),
            Err(_) => Ok(None),
        }
    }
}

/// A shape containing text (p:sp).
///
/// Provides access to text content through a text frame.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Base shape properties
    base: BaseShape,
}

impl Shape {
    /// Create a new Shape from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            base: BaseShape::new(xml_bytes, ShapeType::Shape),
        }
    }

    /// Get the base shape.
    #[inline]
    pub fn base(&mut self) -> &mut BaseShape {
        &mut self.base
    }

    /// Get the text frame for this shape.
    ///
    /// Returns a TextFrame that provides access to the text content.
    pub fn text_frame(&self) -> Result<TextFrame> {
        TextFrame::from_xml(&self.base.xml_bytes)
    }

    /// Quick access to get all text from this shape.
    ///
    /// This is a convenience method that extracts all text content.
    pub fn text(&self) -> Result<String> {
        let tf = self.text_frame()?;
        tf.text()
    }
}
