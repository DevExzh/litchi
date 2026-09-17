use crate::error::{OpcError, Result};
use crate::packuri::PackURI;
use crate::payload::PartPayload;
pub use crate::payload::PayloadHandle;
use crate::rel::{Relationship, Relationships};
use memchr::memmem;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
/// Open Packaging Convention (OPC) objects related to package parts.
///
/// This module provides the Part trait and `XmlPart` implementation for representing
/// parts within an OPC package. Parts are the fundamental units of content in an
/// OPC package, each with a unique partname, content type, and optional relationships.
use std::collections::HashMap;
use std::sync::Arc;

/// Object-safe cloning support for package parts.
///
/// This is public only because it is a supertrait of [`Part`]. Concrete part
/// users implement [`Clone`]; the blanket implementation supplies the erased
/// clone operation.
#[doc(hidden)]
#[allow(
    clippy::module_name_repetitions,
    reason = "supertrait name intentionally mirrors the Part trait it supports"
)]
pub trait PartClone {
    fn clone_part(&self) -> Box<dyn Part + Send + Sync>;
}

impl<T> PartClone for T
where
    T: Part + Clone + 'static,
{
    fn clone_part(&self) -> Box<dyn Part + Send + Sync> {
        Box::new(self.clone())
    }
}

impl Clone for Box<dyn Part + Send + Sync> {
    fn clone(&self) -> Self {
        self.clone_part()
    }
}

/// Trait representing a part in an OPC package.
///
/// Parts are the fundamental units of content in an OPC package. Each part
/// has a unique partname (`PackURI`), a content type, and may have relationships
/// to other parts.
pub trait Part: PartClone + Send + Sync {
    /// Get the binary content of this part.
    /// Returns a reference to the blob data for efficient access.
    fn blob(&self) -> &[u8];

    /// Get the binary content as a shared Arc (zero-copy sharing).
    /// This allows creating sub-slices that share the same allocation.
    fn blob_arc(&self) -> Arc<Vec<u8>>;

    /// Get the content type of this part.
    fn content_type(&self) -> &str;

    /// Get the partname of this part.
    fn partname(&self) -> &PackURI;

    /// Count references to a relationship ID in the part content.
    ///
    /// Uses memchr for fast byte searching. For non-XML parts, returns 0.
    fn rel_ref_count(&self, r_id: &str) -> usize {
        // Fast byte-level search for r:id attribute references
        let blob = self.blob();
        let pattern = format!(r#"r:id="{r_id}""#);

        // Use memmem from memchr for fast substring searching
        let finder = memmem::Finder::new(pattern.as_bytes());
        finder.find_iter(blob).count()
    }

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

    /// Decode this part's payload if it is still held by a source archive.
    ///
    /// A package opened from an owned source keeps each part's payload in the
    /// retained archive until something asks for it (ADR 0030). Every public
    /// route that hands a part to a caller outside `litchi-opc` calls this
    /// first, so a caller never receives a part whose payload has not been
    /// decoded and never sees a decode refusal swallowed.
    ///
    /// Parts that already hold their bytes — every part this crate's callers
    /// can build — implement this as a no-op.
    ///
    /// # Errors
    ///
    /// Returns the refusal the decode produced: [`OpcError::ReadLimit`] for
    /// `PartBytes` or `TotalPartBytes` when the central directory
    /// under-declared the member, an allocation failure, a cancellation, an
    /// I/O error, or a ZIP error for a corrupt Deflate stream or a CRC
    /// mismatch. The refusal is recorded, so it is the same value on every
    /// later access.
    fn ensure_payload(&self) -> Result<()> {
        Ok(())
    }

    /// The payload if it is already available, without decoding anything.
    ///
    /// `None` means the part still holds the payload its source member
    /// carries, so it cannot have been replaced. This is an implementation
    /// detail of publication planning and is not part of the supported
    /// surface.
    #[doc(hidden)]
    fn decoded_blob(&self) -> Option<&[u8]> {
        Some(self.blob())
    }

    /// Capture this part's payload storage without decoding it.
    ///
    /// This is an implementation detail of publication planning and is not
    /// part of the supported surface.
    #[doc(hidden)]
    fn payload_handle(&self) -> PayloadHandle {
        PayloadHandle(PartPayload::ready(self.blob_arc()))
    }

    /// Get the relationships for this part.
    fn rels(&self) -> &Relationships;

    /// Get mutable access to the relationships for this part.
    fn rels_mut(&mut self) -> &mut Relationships;

    /// Set the binary content of this part.
    ///
    /// This allows for modification of part content.
    fn set_blob(&mut self, blob: Vec<u8>);

    /// Replace content with an already shared immutable allocation.
    ///
    /// Custom part implementations need not override this method; the default
    /// may copy when the allocation has other owners. Built-in parts adopt the
    /// allocation directly.
    fn set_blob_shared(&mut self, blob: Arc<Vec<u8>>) {
        match Arc::try_unwrap(blob) {
            Ok(owned_blob) => self.set_blob(owned_blob),
            Err(shared_blob) => self.set_blob(shared_blob.as_ref().clone()),
        }
    }

    /// Replace the content type of this part.
    ///
    /// Callers are responsible for selecting a content type permitted by the
    /// owning package format. The package writer will emit the updated value
    /// into `[Content_Types].xml`.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::InvalidContentType`] if the part does not support
    /// changing its content type (the default behavior).
    fn set_content_type(&mut self, content_type: String) -> Result<()> {
        Err(OpcError::InvalidContentType {
            value: content_type,
            reason: format!(
                "part '{}' does not support changing its content type",
                self.partname().as_str()
            ),
        })
    }

    /// Get the target reference for a relationship ID.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::RelationshipNotFound`] if no relationship with the
    /// given `r_id` exists on this part.
    fn target_ref(&self, r_id: &str) -> Result<&str> {
        self.rels()
            .get(r_id)
            .map(Relationship::target_ref)
            .ok_or_else(|| OpcError::RelationshipNotFound(format!("rId: {r_id}")))
    }
}

/// A basic implementation of a Part that stores binary content.
///
/// This is the default part type for non-XML content. It stores the
/// content as a byte vector and manages relationships. Uses Arc for
/// efficient sharing of blob data.
#[derive(Debug, Clone)]
#[allow(
    clippy::module_name_repetitions,
    reason = "name mirrors the Part trait role; renaming would break the public API"
)]
pub struct BlobPart {
    /// The partname (URI) of this part
    partname: PackURI,

    /// The content type of this part
    content_type: String,

    /// The binary content of this part. A part read from an owned source
    /// archive holds a deferred payload until something asks for its bytes.
    blob: PartPayload,

    /// Relationships from this part to other parts
    rels: Relationships,
}

impl BlobPart {
    /// Create a new `BlobPart`.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of this part
    /// * `content_type` - The content type of this part
    /// * `blob` - The binary content of this part
    #[must_use]
    pub fn new(partname: PackURI, content_type: String, blob: Vec<u8>) -> Self {
        Self::new_shared(partname, content_type, Arc::new(blob))
    }

    /// Create a new `BlobPart` from an already shared allocation.
    ///
    /// The part adopts `blob` without copying its contents.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of this part
    /// * `content_type` - The content type of this part
    /// * `blob` - The shared binary content of this part
    #[must_use]
    pub fn new_shared(partname: PackURI, content_type: String, blob: Arc<Vec<u8>>) -> Self {
        Self::with_payload(partname, content_type, PartPayload::ready(blob))
    }

    /// Create a `BlobPart` over payload storage that may still be deferred.
    pub(crate) fn with_payload(partname: PackURI, content_type: String, blob: PartPayload) -> Self {
        let rels = Relationships::for_source(&partname);
        Self {
            partname,
            content_type,
            blob,
            rels,
        }
    }

    /// Load a part from raw data.
    #[must_use]
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

    fn set_content_type(&mut self, content_type: String) -> Result<()> {
        self.content_type = content_type;
        Ok(())
    }

    fn blob(&self) -> &[u8] {
        self.blob.bytes()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        self.blob.arc()
    }

    fn ensure_payload(&self) -> Result<()> {
        self.blob.force().map(|_| ())
    }

    fn decoded_blob(&self) -> Option<&[u8]> {
        self.blob.decoded().map(|blob| blob.as_slice())
    }

    fn payload_handle(&self) -> PayloadHandle {
        PayloadHandle(self.blob.clone())
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.blob = PartPayload::ready(Arc::new(blob));
    }

    fn set_blob_shared(&mut self, blob: Arc<Vec<u8>>) {
        self.blob = PartPayload::ready(blob);
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
/// `XmlPart` extends the basic Part functionality with XML parsing capabilities.
/// It stores the raw XML as bytes and provides methods for efficient XML processing
/// using quick-xml with zero-copy parsing where possible. Uses Arc for efficient
/// sharing of XML data.
#[derive(Debug)]
#[allow(
    clippy::module_name_repetitions,
    reason = "name mirrors the Part trait role; renaming would break the public API"
)]
pub struct XmlPart {
    /// The partname (URI) of this part
    partname: PackURI,

    /// The content type of this part
    content_type: String,

    /// The XML content as raw bytes (UTF-8 encoded). A part read from an
    /// owned source archive holds a deferred payload until something asks for
    /// its bytes.
    xml_bytes: PartPayload,

    /// Relationships from this part to other parts
    rels: Relationships,

    /// Cached parsed elements (optional, for frequently accessed data)
    /// Maps element paths to their string values for quick lookup
    element_cache: HashMap<String, String>,
}

impl Clone for XmlPart {
    fn clone(&self) -> Self {
        Self {
            partname: self.partname.clone(),
            content_type: self.content_type.clone(),
            xml_bytes: self.xml_bytes.clone(),
            rels: self.rels.clone(),
            // Parsed lookups are disposable derived state. Cloning them would
            // inflate edit snapshots without preserving additional semantics.
            element_cache: HashMap::new(),
        }
    }
}

impl XmlPart {
    /// Create a new `XmlPart`.
    ///
    /// # Arguments
    /// * `partname` - The partname (URI) of this part
    /// * `content_type` - The content type of this part
    /// * `xml_bytes` - The XML content as raw bytes
    #[must_use]
    pub fn new(partname: PackURI, content_type: String, xml_bytes: Vec<u8>) -> Self {
        Self::new_shared(partname, content_type, Arc::new(xml_bytes))
    }

    /// Create a new `XmlPart` from an already shared allocation.
    ///
    /// The part adopts `xml_bytes` without copying its contents. This is used
    /// by the eager package reader when the ZIP reader already owns a shared
    /// decompression buffer.
    #[must_use]
    pub fn new_shared(partname: PackURI, content_type: String, xml_bytes: Arc<Vec<u8>>) -> Self {
        Self::with_payload(partname, content_type, PartPayload::ready(xml_bytes))
    }

    /// Create an `XmlPart` over payload storage that may still be deferred.
    pub(crate) fn with_payload(
        partname: PackURI,
        content_type: String,
        xml_bytes: PartPayload,
    ) -> Self {
        let rels = Relationships::for_source(&partname);
        Self {
            partname,
            content_type,
            xml_bytes,
            rels,
            element_cache: HashMap::new(),
        }
    }

    /// Load an XML part from raw data.
    ///
    /// Note: UTF-8 validation is deferred until actual parsing/access for performance.
    /// Invalid UTF-8 will be caught by quick-xml during parsing or by `xml_str()`.
    #[inline]
    #[must_use]
    pub fn load(partname: PackURI, content_type: String, xml_bytes: Vec<u8>) -> Self {
        Self::new(partname, content_type, xml_bytes)
    }

    /// Get a reader for parsing the XML content.
    ///
    /// Returns a quick-xml Reader configured for efficient parsing.
    /// The reader uses zero-copy parsing where possible.
    #[must_use]
    pub fn reader(&self) -> Reader<&[u8]> {
        let mut reader = Reader::from_reader(self.xml_bytes.bytes());
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
    ///
    /// # Errors
    ///
    /// Returns an error if the XML content cannot be parsed or the extracted
    /// text is not valid UTF-8.
    pub fn extract_text(&mut self, element_name: &str) -> Result<Option<String>> {
        // Check cache first
        if let Some(cached) = self.element_cache.get(element_name) {
            return Ok(Some(cached.clone()));
        }

        let mut reader = self.reader();
        let mut target_depth = 0usize;
        let mut text_content = String::new();

        // Use memchr for fast element name matching
        let element_name_bytes = element_name.as_bytes();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e))
                    // Fast byte-level comparison
                    if target_depth == 0 && e.local_name().as_ref() == element_name_bytes => {
                        target_depth = 1;
                    },
                Ok(Event::Start(_)) if target_depth > 0 => {
                    target_depth = target_depth.saturating_add(1);
                },
                Ok(Event::Text(e)) if target_depth > 0 => {
                    // Efficiently decode text without unnecessary allocation
                    let text = std::str::from_utf8(e.as_ref())?;
                    text_content.push_str(text);
                },
                Ok(Event::End(_)) if target_depth > 0 => {
                    target_depth -= 1;
                    if target_depth == 0 && !text_content.is_empty() {
                        // Cache the result
                        self.element_cache
                            .insert(element_name.to_string(), text_content.clone());
                        return Ok(Some(text_content));
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OpcError::XmlError(format!("XML parse error: {e}"))),
                // A self-closing target has no text and must not make unrelated
                // following text appear to belong to it.
                _ => {},
            }
        }

        Ok(None)
    }

    /// Find all elements matching a tag name and extract their attributes.
    ///
    /// Returns a vector of `HashMaps`, where each `HashMap` contains the attributes
    /// of one matching element. Uses efficient streaming parsing.
    ///
    /// # Errors
    ///
    /// Returns an error if the XML content cannot be parsed or an attribute
    /// name or value is not valid UTF-8.
    pub fn find_elements_with_attrs(
        &self,
        element_name: &str,
    ) -> Result<Vec<HashMap<String, String>>> {
        let mut reader = self.reader();
        let mut results = Vec::new();
        let element_name_bytes = element_name.as_bytes();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e) | Event::Empty(ref e))
                    if e.local_name().as_ref() == element_name_bytes =>
                {
                    let mut attrs = HashMap::new();
                    for attr in e.attributes() {
                        let attribute = attr?;
                        let key = std::str::from_utf8(attribute.key.as_ref())?;
                        let value = attribute.decoded_and_normalized_value(
                            XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )?;
                        attrs.insert(key.to_string(), value.to_string());
                    }
                    results.push(attrs);
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OpcError::XmlError(format!("XML parse error: {e}"))),
                _ => {},
            }
        }

        Ok(results)
    }

    /// Get the XML content as a UTF-8 string.
    ///
    /// Performs zero-copy conversion if possible.
    ///
    /// # Errors
    ///
    /// Returns an error if the XML content is not valid UTF-8.
    pub fn xml_str(&self) -> Result<&str> {
        std::str::from_utf8(self.xml_bytes.bytes()).map_err(Into::into)
    }
}

impl Part for XmlPart {
    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn set_content_type(&mut self, content_type: String) -> Result<()> {
        self.content_type = content_type;
        Ok(())
    }

    fn blob(&self) -> &[u8] {
        self.xml_bytes.bytes()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        self.xml_bytes.arc()
    }

    fn ensure_payload(&self) -> Result<()> {
        self.xml_bytes.force().map(|_| ())
    }

    fn decoded_blob(&self) -> Option<&[u8]> {
        self.xml_bytes.decoded().map(|blob| blob.as_slice())
    }

    fn payload_handle(&self) -> PayloadHandle {
        PayloadHandle(self.xml_bytes.clone())
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.xml_bytes = PartPayload::ready(Arc::new(blob));
        // Clear cache when blob is updated
        self.element_cache.clear();
    }

    fn set_blob_shared(&mut self, blob: Arc<Vec<u8>>) {
        self.xml_bytes = PartPayload::ready(blob);
        self.element_cache.clear();
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }
}

/// A part's name and metadata, without a route to its payload.
///
/// [`OpcPackage::iter_parts`](crate::package::OpcPackage::iter_parts) is the
/// only infallible route to the parts of a package, so it must not be able to
/// reach a payload: a package opened from an owned source decodes a part's
/// payload on first access and that decode can fail, and an infallible
/// iterator has nowhere to report the refusal (ADR 0030). Its item type is
/// therefore this view, which carries everything a metadata pass needs and
/// nothing a byte pass does.
///
/// Use [`OpcPackage::try_iter_parts`](crate::package::OpcPackage::try_iter_parts)
/// when the iteration needs payloads: it yields `Result<&dyn Part>` and forces
/// each part's decode as it yields it.
#[derive(Clone, Copy)]
pub struct PartMetadata<'part> {
    part: &'part (dyn Part + 'part),
}

impl std::fmt::Debug for PartMetadata<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PartMetadata")
            .field("partname", &self.part.partname())
            .field("content_type", &self.part.content_type())
            .finish()
    }
}

impl<'part> PartMetadata<'part> {
    pub(crate) fn new(part: &'part (dyn Part + 'part)) -> Self {
        Self { part }
    }

    /// Get the partname of this part.
    #[must_use]
    pub fn partname(&self) -> &'part PackURI {
        self.part.partname()
    }

    /// Get the content type of this part.
    #[must_use]
    pub fn content_type(&self) -> &'part str {
        self.part.content_type()
    }

    /// Get the relationships for this part.
    #[must_use]
    pub fn rels(&self) -> &'part Relationships {
        self.part.rels()
    }

    /// Get the target reference for a relationship ID.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::RelationshipNotFound`] if no relationship with the
    /// given `r_id` exists on this part.
    pub fn target_ref(&self, r_id: &str) -> Result<&'part str> {
        self.part.target_ref(r_id)
    }

    /// Whether this part's payload is already available, without decoding it.
    ///
    /// `false` means the part still holds the payload its source member
    /// carries. This reports what a lazy open has paid and exists for
    /// measurement and regression tests; it is not part of the supported
    /// surface.
    #[doc(hidden)]
    #[must_use]
    pub fn payload_is_decoded(&self) -> bool {
        self.part.decoded_blob().is_some()
    }
}

/// Factory for creating Part instances based on content type.
///
/// The factory uses a type-based dispatch system to create the appropriate
/// Part implementation (`BlobPart` for binary content, `XmlPart` for XML content).
#[allow(
    clippy::module_name_repetitions,
    reason = "factory name mirrors the Part trait it constructs; renaming would break the public API"
)]
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
    ///
    /// # Errors
    ///
    /// This function currently never fails; the `Result` return type is kept
    /// for future fallible part construction.
    pub fn load(
        partname: PackURI,
        content_type: String,
        blob: Vec<u8>,
    ) -> Result<Box<dyn Part + Send + Sync>> {
        Self::load_shared(partname, content_type, Arc::new(blob))
    }

    /// Load a part from an already shared payload, selecting the appropriate
    /// Part implementation without copying the payload.
    ///
    /// This is the ingress used by the eager OPC reader. The `Vec<u8>`-based
    /// [`Self::load`] method remains the compatibility surface for callers
    /// that own an ordinary vector.
    pub fn load_shared(
        partname: PackURI,
        content_type: String,
        blob: Arc<Vec<u8>>,
    ) -> Result<Box<dyn Part + Send + Sync>> {
        Self::load_payload(partname, content_type, PartPayload::ready(blob))
    }

    /// Load a part over payload storage that may still be deferred.
    ///
    /// The part type is selected from the content type exactly as
    /// [`Self::load_shared`] selects it, so a deferred part and an eager part
    /// of the same content type are the same concrete type.
    pub(crate) fn load_payload(
        partname: PackURI,
        content_type: String,
        blob: PartPayload,
    ) -> Result<Box<dyn Part + Send + Sync>> {
        // Determine if this is an XML part based on content type
        if Self::is_xml_content_type(&content_type) {
            Ok(Box::new(XmlPart::with_payload(
                partname,
                content_type,
                blob,
            )))
        } else {
            Ok(Box::new(BlobPart::with_payload(
                partname,
                content_type,
                blob,
            )))
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
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    #[test]
    fn test_blob_part() {
        let partname = PackURI::new("/word/media/image1.png").unwrap();
        let content = vec![0x89, 0x50, 0x4E, 0x47]; // PNG header
        let mut part = BlobPart::new(partname, "image/png".to_string(), content.clone());

        assert_eq!(part.content_type(), "image/png");
        assert_eq!(part.blob(), content.as_slice());
        part.set_content_type("image/webp".to_string()).unwrap();
        assert_eq!(part.content_type(), "image/webp");
    }

    #[test]
    fn blob_part_new_shared_preserves_allocation() {
        let partname = PackURI::new("/word/media/image1.png").unwrap();
        let content = Arc::new(vec![0x89, 0x50, 0x4E, 0x47]);
        let part = BlobPart::new_shared(partname, "image/png".to_string(), Arc::clone(&content));

        let stored = part.blob_arc();
        assert!(Arc::ptr_eq(&content, &stored));
        assert_eq!(stored.as_slice(), content.as_slice());
    }

    #[test]
    fn part_factory_shared_load_preserves_allocation_for_xml_and_binary() {
        let xml_content = Arc::new(b"<root/>".to_vec());
        let xml = PartFactory::load_shared(
            PackURI::new("/word/document.xml").unwrap(),
            "application/xml".to_string(),
            Arc::clone(&xml_content),
        )
        .unwrap();
        let stored_xml = xml.blob_arc();
        assert!(Arc::ptr_eq(&xml_content, &stored_xml));

        let binary_content = Arc::new(vec![0x89, 0x50, 0x4E, 0x47]);
        let binary = PartFactory::load_shared(
            PackURI::new("/word/media/image1.png").unwrap(),
            "image/png".to_string(),
            Arc::clone(&binary_content),
        )
        .unwrap();
        let stored_binary = binary.blob_arc();
        assert!(Arc::ptr_eq(&binary_content, &stored_binary));
    }

    #[test]
    fn test_xml_part() {
        let partname = PackURI::new("/word/document.xml").unwrap();
        let xml = b"<root><text>Hello</text></root>".to_vec();
        let mut part = XmlPart::new(partname, "application/xml".to_string(), xml);

        let text = part.extract_text("text").unwrap();
        assert_eq!(text, Some("Hello".to_string()));
        part.set_content_type("application/example+xml".to_string())
            .unwrap();
        assert_eq!(part.content_type(), "application/example+xml");
    }

    #[test]
    fn extract_text_ignores_self_closing_targets_and_nested_elements() {
        let partname = PackURI::new("/word/document.xml").unwrap();
        let xml = b"<root><text/><other>ignored</other><text><run>Kept</run></text></root>";
        let mut part = XmlPart::new(partname, "application/xml".to_string(), xml.to_vec());

        assert_eq!(part.extract_text("text").unwrap(), Some("Kept".to_string()));
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
