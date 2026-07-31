/// Base shape types for PowerPoint presentations.
use litchi_ooxml_common::xml::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, unqualified_attribute_value,
};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use crate::pptx::shapes::textframe::TextFrame;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

fn is_shape_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    expected_prefix: &[u8],
    transitional_namespace: &[u8],
    strict_namespace: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == transitional_namespace || *value == strict_namespace
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == expected_prefix,
        ResolveResult::Unbound => false,
    }
}

fn is_drawingml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    is_shape_name(
        namespace,
        name,
        local_name,
        b"a",
        DRAWINGML_NAMESPACE,
        STRICT_DRAWINGML_NAMESPACE,
    )
}

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

        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Empty(element) | Event::Start(element)
                    if is_presentationml_name(&namespace, element.name(), b"cNvPr") =>
                {
                    if let Some(name) = unqualified_attribute_value(&element, b"name", decoder)? {
                        self.name = Some(name.clone());
                        return Ok(name);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Get the non-visual shape ID.
    ///
    /// Returns the optional numeric ID from the shape's PresentationML cNvPr
    /// element. This ID is used by timing and animation records to refer to a
    /// shape.
    pub fn shape_id(&self) -> Result<Option<u32>> {
        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Empty(element) | Event::Start(element)
                    if is_presentationml_name(&namespace, element.name(), b"cNvPr") =>
                {
                    let Some(id) = unqualified_attribute_value(&element, b"id", decoder)? else {
                        return Ok(None);
                    };
                    return id.parse::<u32>().map(Some).map_err(|_| {
                        OoxmlError::InvalidFormat(format!("invalid non-visual shape ID '{id}'"))
                    });
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the X position (left edge) in EMUs.
    pub fn left(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self
            .geometry
            .ok_or_else(|| OoxmlError::InvalidFormat("shape geometry is missing".to_string()))?
            .x)
    }

    /// Get the Y position (top edge) in EMUs.
    pub fn top(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self
            .geometry
            .ok_or_else(|| OoxmlError::InvalidFormat("shape geometry is missing".to_string()))?
            .y)
    }

    /// Get the width in EMUs.
    pub fn width(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self
            .geometry
            .ok_or_else(|| OoxmlError::InvalidFormat("shape geometry is missing".to_string()))?
            .cx)
    }

    /// Get the height in EMUs.
    pub fn height(&mut self) -> Result<i64> {
        self.ensure_geometry()?;
        Ok(self
            .geometry
            .ok_or_else(|| OoxmlError::InvalidFormat("shape geometry is missing".to_string()))?
            .cy)
    }

    /// Check if this shape is a placeholder.
    pub fn is_placeholder(&self) -> bool {
        // Look for <p:ph> element
        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        loop {
            match reader.read_resolved_event() {
                Ok((namespace, Event::Empty(element) | Event::Start(element)))
                    if is_presentationml_name(&namespace, element.name(), b"ph") =>
                {
                    return true;
                },
                Ok((_, Event::Eof)) => break,
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
        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Empty(element) | Event::Start(element)
                    if is_presentationml_name(&namespace, element.name(), b"ph") =>
                {
                    if let Some(placeholder_type) =
                        unqualified_attribute_value(&element, b"type", decoder)?
                    {
                        return Ok(placeholder_type);
                    }
                    // If no type attribute, it's usually a body placeholder
                    return Ok("body".to_string());
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Get the placeholder index for this shape.
    ///
    /// Returns None when the shape is not a placeholder. When a placeholder
    /// omits the idx attribute, this returns its OOXML default of Some(0).
    pub fn placeholder_index(&self) -> Result<Option<u32>> {
        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Empty(element) | Event::Start(element)
                    if is_presentationml_name(&namespace, element.name(), b"ph") =>
                {
                    let Some(index) = unqualified_attribute_value(&element, b"idx", decoder)?
                    else {
                        return Ok(Some(0));
                    };
                    return index.parse::<u32>().map(Some).map_err(|_| {
                        OoxmlError::InvalidFormat(format!("invalid placeholder index '{index}'"))
                    });
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(None)
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
        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);
        loop {
            match reader.read_resolved_event() {
                Ok((namespace, Event::Start(element) | Event::Empty(element)))
                    if is_drawingml_name(&namespace, element.name(), b"tbl") =>
                {
                    return true;
                },
                Ok((_, Event::Eof)) | Err(_) => return false,
                _ => {},
            }
        }
    }

    /// Ensure geometry is parsed and cached.
    fn ensure_geometry(&mut self) -> Result<()> {
        if self.geometry.is_some() {
            return Ok(());
        }

        let mut reader = NsReader::from_reader(&self.xml_bytes[..]);

        let mut x = None;
        let mut y = None;
        let mut cx = None;
        let mut cy = None;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Empty(element) | Event::Start(element)
                    if is_drawingml_name(&namespace, element.name(), b"off")
                        && (x.is_none() || y.is_none()) =>
                {
                    x = Some(parse_coordinate(&element, b"x", decoder)?);
                    y = Some(parse_coordinate(&element, b"y", decoder)?);
                },
                Event::Empty(element) | Event::Start(element)
                    if is_drawingml_name(&namespace, element.name(), b"ext")
                        && (cx.is_none() || cy.is_none()) =>
                {
                    cx = Some(parse_positive_coordinate(&element, b"cx", decoder)?);
                    cy = Some(parse_positive_coordinate(&element, b"cy", decoder)?);
                },
                Event::Eof => break,
                _ => {},
            }
        }

        self.geometry = Some(ShapeGeometry {
            x: x.unwrap_or(0),
            y: y.unwrap_or(0),
            cx: cx.unwrap_or(0),
            cy: cy.unwrap_or(0),
        });
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

fn parse_coordinate(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "DrawingML coordinate is missing {}",
            String::from_utf8_lossy(name)
        ))
    })?;
    value
        .parse::<i64>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid DrawingML coordinate '{}'", value)))
}

fn parse_positive_coordinate(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<i64> {
    let value = parse_coordinate(element, name, decoder)?;
    if value < 0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "DrawingML extent {} cannot be negative",
            String::from_utf8_lossy(name)
        )));
    }
    Ok(value)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shape_properties_resolve_namespaces_and_decode_attributes() {
        let xml = br#"<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:false="urn:not-office">
            <q:nvSpPr><false:cNvPr name="ignored"/><q:cNvPr id="7" name="A &amp; B"/><q:nvPr><false:ph type="title"/><q:ph type="ctrTitle" idx="4"/></q:nvPr></q:nvSpPr>
            <q:spPr><d:xfrm><false:off x="999" y="999"/><d:off x="-10" y="20"/><d:ext cx="30" cy="40"/></d:xfrm></q:spPr>
        </q:sp>"#;
        let mut shape = BaseShape::new(xml.to_vec(), ShapeType::Shape);
        assert_eq!(shape.name().unwrap(), "A & B");
        assert_eq!(shape.shape_id().unwrap(), Some(7));
        assert!(shape.is_placeholder());
        assert_eq!(shape.placeholder_type().unwrap(), "ctrTitle");
        assert_eq!(shape.placeholder_index().unwrap(), Some(4));
        assert_eq!(shape.left().unwrap(), -10);
        assert_eq!(shape.top().unwrap(), 20);
        assert_eq!(shape.width().unwrap(), 30);
        assert_eq!(shape.height().unwrap(), 40);
    }

    #[test]
    fn shape_properties_accept_strict_and_inherited_prefixes() {
        let strict = br#"<s:sp xmlns:s="http://purl.oclc.org/ooxml/presentationml/main" xmlns:d="http://purl.oclc.org/ooxml/drawingml/main"><s:cNvPr name="Strict"/><d:off x="1" y="2"/><d:ext cx="3" cy="4"/></s:sp>"#;
        let mut shape = BaseShape::new(strict.to_vec(), ShapeType::Shape);
        assert_eq!(shape.name().unwrap(), "Strict");
        assert_eq!(shape.shape_id().unwrap(), None);
        assert_eq!(shape.placeholder_index().unwrap(), None);
        assert_eq!(shape.width().unwrap(), 3);

        let inherited = br#"<p:sp><p:cNvPr name="Inherited"/><a:off x="5" y="6"/><a:ext cx="7" cy="8"/></p:sp>"#;
        let mut shape = BaseShape::new(inherited.to_vec(), ShapeType::Shape);
        assert_eq!(shape.name().unwrap(), "Inherited");
        assert_eq!(shape.height().unwrap(), 8);
    }

    #[test]
    fn table_detection_requires_a_drawingml_element() {
        let spoof = br#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:false="urn:not-drawingml"><p:meta value="a:tbl"/><false:tbl/></p:graphicFrame>"#;
        assert!(!BaseShape::new(spoof.to_vec(), ShapeType::GraphicFrame).has_table());

        let table = br#"<p:graphicFrame><a:graphic><a:graphicData><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>"#;
        assert!(BaseShape::new(table.to_vec(), ShapeType::GraphicFrame).has_table());
    }

    #[test]
    fn malformed_shape_attributes_are_errors() {
        let duplicate_name = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cNvPr name="one" name="two"/></p:sp>"#;
        assert!(
            BaseShape::new(duplicate_name.to_vec(), ShapeType::Shape)
                .name()
                .is_err()
        );

        let invalid_geometry = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="NaN" y="2"/><a:ext cx="-1" cy="4"/></p:sp>"#;
        let mut shape = BaseShape::new(invalid_geometry.to_vec(), ShapeType::Shape);
        assert!(shape.left().is_err());

        let invalid_shape_id = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cNvPr id="NaN"/></p:sp>"#;
        assert!(
            BaseShape::new(invalid_shape_id.to_vec(), ShapeType::Shape)
                .shape_id()
                .is_err()
        );

        let invalid_placeholder_index = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ph idx="NaN"/></p:sp>"#;
        assert!(
            BaseShape::new(invalid_placeholder_index.to_vec(), ShapeType::Shape)
                .placeholder_index()
                .is_err()
        );

        let missing_extent = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext cx="3"/></p:sp>"#;
        assert!(
            BaseShape::new(missing_extent.to_vec(), ShapeType::Shape)
                .width()
                .is_err()
        );
    }

    #[test]
    fn placeholder_index_defaults_to_zero() {
        let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ph type="body"/></p:sp>"#;
        let shape = BaseShape::new(xml.to_vec(), ShapeType::Shape);
        assert_eq!(shape.placeholder_index().unwrap(), Some(0));
    }
}
