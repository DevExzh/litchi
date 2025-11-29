/// Slide parts and related types.
///
/// This module contains parts for slides, slide layouts, and slide masters.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use crate::ooxml::pptx::shapes::base::{BaseShape, ShapeType};
use quick_xml::Reader;
use quick_xml::events::Event;

/// A slide part.
///
/// Corresponds to `/ppt/slides/slideN.xml` in the package.
pub struct SlidePart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlidePart<'a> {
    /// Create a SlidePart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the slide.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the slide name.
    ///
    /// Returns the name attribute from the <p:cSld> element.
    pub fn name(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"cSld" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                let name = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return Ok(name.to_string());
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Extract all text content from the slide.
    ///
    /// This extracts text from all `<a:t>` elements in the slide (DrawingML text).
    pub fn extract_text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    // Check if this is an a:t element (DrawingML text)
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Extract text content
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    if !text.is_empty() && !text.ends_with('\n') {
                        text.push('\n');
                    }
                    text.push_str(t);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(text)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Parse and return all shapes on this slide.
    ///
    /// Returns a vector of BaseShape objects that can be checked for type
    /// and converted to specific shape types.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut shapes = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) => {
                    let local_name = e.local_name();
                    let tag_name_bytes = local_name.as_ref();

                    // Extract individual shape elements
                    let shape_type = match tag_name_bytes {
                        b"sp" => Some(ShapeType::Shape),                  // Text shape
                        b"pic" => Some(ShapeType::Picture),               // Picture
                        b"graphicFrame" => Some(ShapeType::GraphicFrame), // Table/Chart
                        b"grpSp" => Some(ShapeType::GroupShape),          // Group
                        b"cxnSp" => Some(ShapeType::Connector),           // Connector
                        _ => None,
                    };

                    if let Some(st) = shape_type {
                        // Create a new buffer for extracting shape XML
                        let mut shape_buf = Vec::new();
                        // Extract the complete shape XML
                        if let Ok(shape_xml) =
                            Self::extract_shape_xml(&mut reader, tag_name_bytes, &mut shape_buf)
                        {
                            shapes.push(BaseShape::new(shape_xml, st));
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        Ok(shapes)
    }

    /// Helper to extract complete shape XML.
    fn extract_shape_xml(
        reader: &mut Reader<&[u8]>,
        tag_name: &[u8],
        _buf: &mut Vec<u8>,
    ) -> Result<Vec<u8>> {
        let mut shape_xml = Vec::new();
        let mut depth = 1;

        // Start tag already consumed, write it
        shape_xml.push(b'<');
        shape_xml.extend_from_slice(tag_name);
        shape_xml.push(b'>');

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    depth += 1;
                    shape_xml.push(b'<');
                    shape_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        shape_xml.push(b' ');
                        shape_xml.extend_from_slice(attr.key.as_ref());
                        shape_xml.extend_from_slice(b"=\"");
                        shape_xml.extend_from_slice(&attr.value);
                        shape_xml.push(b'"');
                    }
                    shape_xml.push(b'>');
                },
                Ok(Event::End(e)) => {
                    shape_xml.extend_from_slice(b"</");
                    shape_xml.extend_from_slice(e.name().as_ref());
                    shape_xml.push(b'>');

                    depth -= 1;
                    if depth == 0 {
                        return Ok(shape_xml);
                    }
                },
                Ok(Event::Text(e)) => {
                    shape_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) => {
                    shape_xml.push(b'<');
                    shape_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        shape_xml.push(b' ');
                        shape_xml.extend_from_slice(attr.key.as_ref());
                        shape_xml.extend_from_slice(b"=\"");
                        shape_xml.extend_from_slice(&attr.value);
                        shape_xml.push(b'"');
                    }
                    shape_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Err(OoxmlError::Xml("Unexpected end of shape XML".to_string()))
    }

    /// Get the transition effect for this slide.
    ///
    /// Parses the `<p:transition>` element from the slide XML.
    /// Returns `None` if no transition is defined.
    pub fn transition(&self) -> Result<Option<crate::ooxml::pptx::transitions::SlideTransition>> {
        crate::ooxml::pptx::transitions::SlideTransition::from_xml(self.xml_bytes())
    }

    /// Get the background for this slide.
    ///
    /// Parses the `<p:bg>` element from the slide XML.
    /// Returns `None` if no background is defined.
    pub fn background(&self) -> Result<Option<crate::ooxml::pptx::backgrounds::SlideBackground>> {
        crate::ooxml::pptx::backgrounds::SlideBackground::from_xml(self.xml_bytes())
    }
}

/// A slide layout part.
///
/// Corresponds to `/ppt/slideLayouts/slideLayoutN.xml` in the package.
pub struct SlideLayoutPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlideLayoutPart<'a> {
    /// Create a SlideLayoutPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the layout.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the layout name.
    pub fn name(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"cSld" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                let name = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return Ok(name.to_string());
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

/// A slide master part.
///
/// Corresponds to `/ppt/slideMasters/slideMasterN.xml` in the package.
pub struct SlideMasterPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlideMasterPart<'a> {
    /// Create a SlideMasterPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the master.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the master name.
    pub fn name(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"cSld" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                let name = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return Ok(name.to_string());
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(String::new())
    }

    /// Get the relationship IDs of all slide layouts in this master.
    pub fn slide_layout_rids(&self) -> Result<Vec<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut rids = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldLayoutId" {
                        for attr in e.attributes().flatten() {
                            // Look for r:id attribute (can be r:id or just id with relationships namespace)
                            let key = attr.key.as_ref();
                            // Check if this is the relationship ID attribute
                            if key == b"r:id"
                                || (key.starts_with(b"r:")
                                    && attr.key.local_name().as_ref() == b"id")
                                || attr.key.local_name().as_ref() == b"id"
                            {
                                let rid = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                // Only push if it looks like a relationship ID (starts with "rId")
                                if rid.starts_with("rId") {
                                    rids.push(rid.to_string());
                                    break;
                                }
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(rids)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}
