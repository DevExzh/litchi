use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::PackURI;
use crate::ooxml::opc::rel::Relationships;
use memchr::memmem;
use quick_xml::events::Event;
use quick_xml::Reader;
/// Open Packaging Convention (OPC) objects related to package parts.
///
/// This module provides the Part trait and XmlPart implementation for representing
/// parts within an OPC package. Parts are the fundamental units of content in an
/// OPC package, each with a unique partname, content type, and optional relationships.
use std::collections::HashMap;
use std::sync::Arc;

/// Trait representing a part in an OPC package.
///
/// Parts are the fundamental units of content in an OPC package. Each part
/// has a unique partname (PackURI), a content type, and may have relationships
/// to other parts.
pub trait Part {
    /// Get the partname of this part.
    fn partname(&self) -> &PackURI;

    /// Get the content type of this part.
    fn content_type(&self) -> &str;

    /// Get the binary content of this part.
    /// Returns a reference to the blob data for efficient access.
    fn blob(&self) -> &[u8];

    /// Get the relationships for this part.
    fn rels(&self) -> &Relationships;

    /// Get mutable access to the relationships for this part.
    fn rels_mut(&mut self) -> &mut Relationships;

    /// Add or get a relationship to another part.
    ///
    /// If a relationship of the given type to the target already exists,
    /// returns its rId. Otherwise, creates a new relationship and returns
    /// the new rId.
    fn relate_to(&mut self, target_partname: &str, reltype: &str) -> String {
        let rel = self.rels_mut().get_or_add(reltype, target_partname);
        rel.r_id().to_string()
    }

    /// Add or get an external relationship.
    fn relate_to_ext(&mut self, target_url: &str, reltype: &str) -> String {
        self.rels_mut().get_or_add_ext_rel(reltype, target_url)
    }

    /// Get the target reference for a relationship ID.
    fn target_ref(&self, r_id: &str) -> Result<&str> {
        self.rels()
            .get(r_id)
            .map(|rel| rel.target_ref())
            .ok_or_else(|| OpcError::RelationshipNotFound(format!("rId: {}", r_id)))
    }

    /// Count references to a relationship ID in the part content.
    ///
    /// Uses memchr for fast byte searching. For non-XML parts, returns 0.
    fn rel_ref_count(&self, r_id: &str) -> usize {
        // Fast byte-level search for r:id attribute references
        let blob = self.blob();
        let pattern = format!(r#"r:id="{}""#, r_id);

        // Use memmem from memchr for fast substring searching
        let finder = memmem::Finder::new(pattern.as_bytes());
        finder.find_iter(blob).count()
    }
}

/// A basic implementation of a Part that stores binary content.
///
/// This is the default part type for non-XML content. It stores the
/// content as a byte vector and manages relationships. Uses Arc for
/// efficient sharing of blob data.
#[derive(Debug)]
pub struct BlobPart {
    /// The partname (URI) of this part
    partname: PackURI,

    /// The content type of this part
    content_type: String,

    /// The binary content of this part (shared via Arc for efficiency)
    blob: Arc<Vec<u8>>,

    /// Relationships from this part to other parts
    rels: Relationships,
}

impl BlobPart {
    /// Create a new BlobPart.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of this part
    /// * `content_type` - The content type of this part
    /// * `blob` - The binary content of this part
    pub fn new(partname: PackURI, content_type: String, blob: Vec<u8>) -> Self {
        let rels = Relationships::new(partname.base_uri().to_string());
        Self {
            partname,
            content_type,
            blob: Arc::new(blob),
            rels,
        }
    }

    /// Load a part from raw data.
    pub fn load(partname: PackURI, content_type: String, blob: Vec<u8>) -> Self {
        Self::new(partname, content_type, blob)
    }
}

impl Part for BlobPart {
    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn blob(&self) -> &[u8] {
        &self.blob
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }
}

/// An XML part that provides parsed access to its XML content.
///
/// XmlPart extends the basic Part functionality with XML parsing capabilities.
/// It stores the raw XML as bytes and provides methods for efficient XML processing
/// using quick-xml with zero-copy parsing where possible. Uses Arc for efficient
/// sharing of XML data.
#[derive(Debug)]
pub struct XmlPart {
    /// The partname (URI) of this part
    partname: PackURI,

    /// The content type of this part
    content_type: String,

    /// The XML content as raw bytes (UTF-8 encoded, shared via Arc)
    xml_bytes: Arc<Vec<u8>>,

    /// Relationships from this part to other parts
    rels: Relationships,

    /// Cached parsed elements (optional, for frequently accessed data)
    /// Maps element paths to their string values for quick lookup
    element_cache: HashMap<String, String>,
}

impl XmlPart {
    /// Create a new XmlPart.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of this part
    /// * `content_type` - The content type of this part
    /// * `xml_bytes` - The XML content as raw bytes
    pub fn new(partname: PackURI, content_type: String, xml_bytes: Vec<u8>) -> Self {
        let rels = Relationships::new(partname.base_uri().to_string());
        Self {
            partname,
            content_type,
            xml_bytes: Arc::new(xml_bytes),
            rels,
            element_cache: HashMap::new(),
        }
    }

    /// Load an XML part from raw data.
    pub fn load(partname: PackURI, content_type: String, xml_bytes: Vec<u8>) -> Result<Self> {
        // Validate that it's valid UTF-8 XML
        std::str::from_utf8(&xml_bytes)
            .map_err(|e| OpcError::XmlError(format!("Invalid UTF-8 in XML: {}", e)))?;

        Ok(Self::new(partname, content_type, xml_bytes))
    }

    /// Get a reader for parsing the XML content.
    ///
    /// Returns a quick-xml Reader configured for efficient parsing.
    /// The reader uses zero-copy parsing where possible.
    pub fn reader(&self) -> Reader<&[u8]> {
        let mut reader = Reader::from_reader(&**self.xml_bytes);
        reader.config_mut().trim_text(true);
        reader
    }

    /// Extract text content from a specific XML element.
    ///
    /// Uses efficient event-based parsing with quick-xml to find and extract
    /// text from the first occurrence of the specified element.
    ///
    /// # Arguments
    /// * `element_name` - The local name of the element to find (e.g., "text")
    pub fn extract_text(&mut self, element_name: &str) -> Result<Option<String>> {
        // Check cache first
        if let Some(cached) = self.element_cache.get(element_name) {
            return Ok(Some(cached.clone()));
        }

        let mut reader = self.reader();
        let mut buf = Vec::new();
        let mut in_target_element = false;
        let mut text_content = String::new();

        // Use memchr for fast element name matching
        let element_name_bytes = element_name.as_bytes();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    // Fast byte-level comparison
                    if e.local_name().as_ref() == element_name_bytes {
                        in_target_element = true;
                    }
                }
                Ok(Event::Text(e)) if in_target_element => {
                    // Efficiently decode text without unnecessary allocation
                    let text = std::str::from_utf8(e.as_ref())?;
                    text_content.push_str(text);
                }
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == element_name_bytes {
                        in_target_element = false;
                        if !text_content.is_empty() {
                            // Cache the result
                            self.element_cache
                                .insert(element_name.to_string(), text_content.clone());
                            return Ok(Some(text_content));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OpcError::XmlError(format!("XML parse error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Find all elements matching a tag name and extract their attributes.
    ///
    /// Returns a vector of HashMaps, where each HashMap contains the attributes
    /// of one matching element. Uses efficient streaming parsing.
    pub fn find_elements_with_attrs(
        &self,
        element_name: &str,
    ) -> Result<Vec<HashMap<String, String>>> {
        let mut reader = self.reader();
        let mut buf = Vec::new();
        let mut results = Vec::new();
        let element_name_bytes = element_name.as_bytes();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == element_name_bytes {
                        let mut attrs = HashMap::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = std::str::from_utf8(attr.key.as_ref())?;
                            let value = attr.unescape_value()?;
                            attrs.insert(key.to_string(), value.to_string());
                        }
                        results.push(attrs);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OpcError::XmlError(format!("XML parse error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        Ok(results)
    }

    /// Get the XML content as a UTF-8 string.
    ///
    /// Performs zero-copy conversion if possible.
    pub fn xml_str(&self) -> Result<&str> {
        std::str::from_utf8(&self.xml_bytes).map_err(Into::into)
    }
}

impl Part for XmlPart {
    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn blob(&self) -> &[u8] {
        &self.xml_bytes
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }
}

/// Factory for creating Part instances based on content type.
///
/// The factory uses a type-based dispatch system to create the appropriate
/// Part implementation (BlobPart for binary content, XmlPart for XML content).
pub struct PartFactory;

impl PartFactory {
    /// Load a part from raw data, selecting the appropriate Part type based on content type.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of the part
    /// * `content_type` - The content type of the part
    /// * `blob` - The raw binary content (consumed by this function)
    ///
    /// # Returns
    /// A boxed Part trait object
    pub fn load(partname: PackURI, content_type: String, blob: Vec<u8>) -> Result<Box<dyn Part>> {
        // Determine if this is an XML part based on content type
        if Self::is_xml_content_type(&content_type) {
            Ok(Box::new(XmlPart::load(partname, content_type, blob)?))
        } else {
            Ok(Box::new(BlobPart::load(partname, content_type, blob)))
        }
    }

    /// Check if a content type represents XML content.
    ///
    /// Uses fast string searching with memchr to check for "+xml" suffix.
    #[inline]
    fn is_xml_content_type(content_type: &str) -> bool {
        // Fast check for "+xml" or "xml" in content type
        content_type.ends_with("+xml") || content_type.ends_with("/xml")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_blob_part() {
        let partname = PackURI::new("/word/media/image1.png").unwrap();
        let content = vec![0x89, 0x50, 0x4E, 0x47]; // PNG header
        let part = BlobPart::new(partname, "image/png".to_string(), content.clone());

        assert_eq!(part.content_type(), "image/png");
        assert_eq!(part.blob(), content.as_slice());
    }

    #[test]
    fn test_xml_part() {
        let partname = PackURI::new("/word/document.xml").unwrap();
        let xml = b"<root><text>Hello</text></root>".to_vec();
        let mut part = XmlPart::new(partname, "application/xml".to_string(), xml);

        let text = part.extract_text("text").unwrap();
        assert_eq!(text, Some("Hello".to_string()));
    }

    #[test]
    fn test_is_xml_content_type() {
        assert!(PartFactory::is_xml_content_type("application/xml"));
        assert!(PartFactory::is_xml_content_type(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        ));
        assert!(!PartFactory::is_xml_content_type("image/png"));
    }
}
