//! Provenance-bearing publication of source-preserving OPC XML parts.
//!
//! Source XML is deliberately a different publication class from authored
//! XML.  The source bytes are retained exactly, while only the small fragments
//! supplied by a caller are checked against the compact authored-XML contract.
//! The resulting value remains tied to the positional source that issued it.

use crate::content_type::ContentType;
use crate::error::{OpcError, Result};
use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::PackURI;
use crate::pkgreader::is_xml_id;
use crate::source_backed::{PartData, SourceLineage, SourceSnapshot};
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource, SourceVersion};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesStart, Event, attributes::Attribute};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::fmt;
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

const MAX_FRAGMENT_BYTES: usize = 256 * 1024;
const MAX_SOURCE_XML_ATTRIBUTES_PER_ELEMENT: usize = 4096;
const MAX_SOURCE_XML_NAMESPACE_DECLARATIONS_PER_ELEMENT: usize = 256;
const XML_VALIDATION_FIXED_WORKING_BYTES: usize = 4096;
const XML_VALIDATION_NAMESPACE_BINDING_BYTES: usize = 64;
const XML_VALIDATION_OPEN_ELEMENT_BYTES: usize = 2 * size_of::<usize>();
// Raw keys, expanded names, and quick-xml duplicate-check ranges, with
// capacity growth allowance for all three vectors.
const XML_VALIDATION_ATTRIBUTE_BYTES: usize =
    2 * (size_of::<&[u8]>() + size_of::<(Option<&[u8]>, &[u8])>() + size_of::<Range<usize>>());
const XML_MIN_NAMESPACE_DECLARATION_BYTES: usize = 7;
const XML_MIN_OPEN_ELEMENT_BYTES: usize = 3;
const XML_MIN_ATTRIBUTE_BYTES: usize = 4;
const QUICK_XML_MAX_SAFE_DEPTH: usize = (u16::MAX as usize) - 1;
// `verify_authored` parses a caller-owned slice, so it does not need the
// streaming auditor's retained input windows. Keep a conservative envelope
// for its temporary event/attribute scratch and state vectors anyway.
const AUTHORED_AUDIT_WORKING_BYTES: usize = 4096;
const AUTHORED_AUDIT_MEMORY_MULTIPLIER: usize = 32;
const FRAGMENT_ROOT_OPEN: &[u8] = b"<litchi-opc-fragment>";
const FRAGMENT_ROOT_CLOSE: &[u8] = b"</litchi-opc-fragment>";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Exact XML bytes loaded from one source-backed OPC Part.
///
/// This value has no public constructor.  It is issued by [`crate::PartView::source_xml`]
/// after OPC classification, source checks, and bounded well-formedness
/// validation.  Its source bytes and managed [`PartData`] handle remain alive
/// until the value is dropped.
#[derive(Clone)]
pub struct SourceXmlPart {
    pub(crate) source: SourceSnapshot,
    pub(crate) source_lineage: SourceLineage,
    pub(crate) source_version: SourceVersion,
    pub(crate) source_partname: Arc<PackURI>,
    pub(crate) source_content_type: Arc<ContentType>,
    pub(crate) original: PartData,
    pub(crate) payload: Arc<Vec<u8>>,
    pub(crate) payload_reservation: Option<Arc<Reservation>>,
    _metadata_reservation: Option<Arc<Reservation>>,
    pub(crate) limits: ReadLimits,
}

impl fmt::Debug for SourceXmlPart {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceXmlPart")
            .field("source_lineage", &self.source_lineage)
            .field("source_version", &self.source_version)
            .field("source_partname", &self.source_partname)
            .field("source_content_type", &self.source_content_type)
            .field("source_bytes", &self.original.as_bytes().len())
            .field("payload_bytes", &self.payload.len())
            .field("derived_payload", &self.payload_reservation.is_some())
            .finish()
    }
}

impl SourceXmlPart {
    pub(crate) fn new(
        source: SourceSnapshot,
        limits: ReadLimits,
        source_partname: &PackURI,
        source_content_type: &str,
        original: PartData,
    ) -> Result<Self> {
        source.ensure_current_public()?;
        if let Some(context) = source.context_ref() {
            context.check().map_err(map_execution_error)?;
        }
        // Reserve retained metadata before cloning the catalog strings. Token
        // clones share both metadata and this reservation.
        let metadata_bytes = size_of::<Self>()
            .checked_add(source_partname.as_str().len())
            .and_then(|bytes| bytes.checked_add(source_content_type.len()))
            .and_then(|bytes| bytes.checked_add(size_of::<[usize; 8]>()))
            .ok_or_else(|| invalid_source("source XML metadata size overflows"))?;
        let metadata_reservation = reserve_output_memory(&source, metadata_bytes)?;
        let source_content_type = ContentType::new(source_content_type)?;
        let bytes = original.as_bytes();
        limits.check(
            ReadResource::PartBytes,
            bytes.len() as u64,
            limits.max_part_bytes(),
        )?;
        validate_source_xml(source_partname, bytes, limits, source.context_ref())?;
        source.ensure_current_public()?;
        if let Some(context) = source.context_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let payload = original.shared_bytes();
        Ok(Self {
            source_lineage: source.lineage().clone(),
            source_version: source.version(),
            source,
            source_partname: Arc::new(source_partname.clone()),
            source_content_type: Arc::new(source_content_type),
            original,
            payload,
            payload_reservation: None,
            _metadata_reservation: metadata_reservation,
            limits,
        })
    }

    /// Return the exact source or assembled XML bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.payload.as_slice()
    }

    /// Return the Part URI from which the source bytes were captured.
    #[must_use]
    pub fn partname(&self) -> &PackURI {
        self.source_partname.as_ref()
    }

    /// Return the source Part content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        self.source_content_type.as_str()
    }

    /// Issue a proof for an exact source byte range.
    ///
    /// The expected bytes are compared before the proof is issued.  Empty
    /// ranges are valid and are used for bounded insertion at an exact source
    /// position.
    pub fn checked_range(
        &self,
        range: Range<usize>,
        expected_source: &[u8],
    ) -> Result<XmlSourceRange> {
        self.check_source_state()?;
        if !self.is_original_payload() {
            return Err(invalid_source(
                "derived source XML cannot issue a second range proof",
            ));
        }
        if range.end < range.start {
            return Err(invalid_source("XML source range is reversed"));
        }
        let range_len = range
            .end
            .checked_sub(range.start)
            .ok_or_else(|| invalid_source("XML source range length overflows"))?;
        self.limits.check(
            ReadResource::XmlAttributeBytes,
            range_len as u64,
            limits_range_max(self.limits),
        )?;
        if range_len != expected_source.len() {
            return Err(invalid_source(
                "XML source range does not match expected byte length",
            ));
        }
        let actual = self
            .original
            .as_bytes()
            .get(range.clone())
            .ok_or_else(|| invalid_source("XML source range is outside the source Part"))?;
        if actual != expected_source {
            return Err(invalid_source("stale XML source range"));
        }
        consume_work(&self.source, expected_source.len())?;
        let expected_metadata = size_of::<XmlSourceRange>()
            .checked_add(self.source_partname.as_str().len())
            .ok_or_else(|| invalid_source("XML source range metadata size overflows"))?;
        let expected_reservation =
            reserve_edit_memory(&self.source, expected_source.len(), expected_metadata)?;
        let mut expected = Vec::new();
        expected
            .try_reserve_exact(expected_source.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC XML range proof",
                source,
            })?;
        expected.extend_from_slice(expected_source);
        self.check_source_state()?;
        Ok(XmlSourceRange {
            source_lineage: self.source_lineage.clone(),
            source_version: self.source_version,
            source_partname: self.source_partname.as_ref().clone(),
            range,
            expected,
            _reservation: expected_reservation,
        })
    }

    /// Begin a source-preserving edit transaction.
    pub fn into_publication(self) -> Result<XmlSplicePublication> {
        self.check_source_state()?;
        if !self.is_original_payload() {
            return Err(invalid_source(
                "derived source XML cannot be edited a second time",
            ));
        }
        Ok(XmlSplicePublication {
            source: self,
            edits: Vec::new(),
        })
    }

    pub(crate) fn payload_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.payload)
    }

    pub(crate) fn source_snapshot(&self) -> SourceSnapshot {
        self.source.clone()
    }

    fn is_original_payload(&self) -> bool {
        Arc::ptr_eq(&self.payload, &self.original.shared_bytes())
    }

    pub(crate) fn check_for_publication(
        &self,
        destination_content_type: &str,
        destination_limits: ReadLimits,
    ) -> Result<()> {
        if self.source_content_type.as_str() != destination_content_type {
            return Err(invalid_source(
                "source XML content type differs from destination content type",
            ));
        }
        if self.source.snapshot_lineage() != self.source_lineage {
            return Err(invalid_source(
                "source XML lineage authority is inconsistent",
            ));
        }
        self.check_source_state()?;
        destination_limits.check(
            ReadResource::PartBytes,
            self.payload.len() as u64,
            destination_limits.max_part_bytes(),
        )?;
        if destination_limits == self.limits {
            // Both capture and splice finish validate the immutable payload
            // under these exact limits. Reuse that proof without allocating a
            // second parser workspace; keep publication Work bounded even
            // when the payload does not need to be scanned again.
            consume_work_from_context(self.source.context_ref(), self.payload.len())?;
        } else {
            validate_source_xml(
                &self.source_partname,
                self.payload.as_slice(),
                destination_limits,
                self.source.context_ref(),
            )?;
        }
        self.check_source_state()?;
        Ok(())
    }

    pub(crate) fn check_for_replacement(
        &self,
        destination: &SourceSnapshot,
        destination_partname: &PackURI,
        destination_content_type: &str,
        destination_original: &[u8],
        destination_limits: ReadLimits,
    ) -> Result<()> {
        destination.ensure_current_public()?;
        if self.source_lineage != destination.lineage().clone()
            || self.source_version != destination.version()
        {
            return Err(invalid_source(
                "source XML replacement belongs to a different destination source",
            ));
        }
        if !self.source_partname.is_equivalent_to(destination_partname) {
            return Err(invalid_source(
                "source XML replacement Part identity differs from destination",
            ));
        }
        if self.original.as_bytes() != destination_original {
            return Err(invalid_source(
                "source XML replacement does not retain the destination source bytes",
            ));
        }
        self.check_for_publication(destination_content_type, destination_limits)
    }

    pub(crate) fn check_source_state(&self) -> Result<()> {
        self.source.ensure_current_public()?;
        if let Some(context) = self.source.context_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let current = self.source.version();
        if current != self.source_version {
            return Err(OpcError::SourceChanged {
                expected: self.source_version,
                actual: current,
            });
        }
        Ok(())
    }
}

/// An opaque proof that exact bytes at one source range were observed.
pub struct XmlSourceRange {
    source_lineage: SourceLineage,
    source_version: SourceVersion,
    source_partname: PackURI,
    range: Range<usize>,
    expected: Vec<u8>,
    _reservation: Option<Arc<Reservation>>,
}

impl fmt::Debug for XmlSourceRange {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("XmlSourceRange")
            .field("source_lineage", &self.source_lineage)
            .field("source_version", &self.source_version)
            .field("source_partname", &self.source_partname)
            .field("range", &self.range)
            .field("expected_bytes", &self.expected.len())
            .finish()
    }
}

/// Individually audited bytes that may replace a checked source range.
pub struct AuthoredXmlFragment {
    bytes: Vec<u8>,
}

impl fmt::Debug for AuthoredXmlFragment {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("AuthoredXmlFragment")
            .field("bytes", &self.bytes.len())
            .finish()
    }
}

impl AuthoredXmlFragment {
    /// Audit one or more compact authored markup nodes.
    pub fn markup(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let bytes = bytes.into();
        if bytes.is_empty() || bytes.first() != Some(&b'<') {
            return Err(invalid_source(
                "authored XML markup fragment is unclassified",
            ));
        }
        audit_wrapped_fragment(&bytes)?;
        Ok(Self { bytes })
    }

    /// Audit compact authored markup while charging its temporary wrapper and
    /// parser workspace to an explicit execution context before allocation.
    ///
    /// The ordinary [`Self::markup`] constructor remains unmanaged for
    /// compatibility. Source-backed callers must use this constructor when
    /// they already hold a managed execution context. The caller-owned input
    /// vector and the returned fragment's retained bytes are outside this
    /// temporary reservation; the caller admits that retained payload through
    /// its surrounding operation budget.
    pub fn markup_with_execution_context(
        bytes: Vec<u8>,
        context: &ExecutionContext,
    ) -> Result<Self> {
        if bytes.is_empty() || bytes.first() != Some(&b'<') {
            return Err(invalid_source(
                "authored XML markup fragment is unclassified",
            ));
        }
        audit_wrapped_fragment_with_context(&bytes, context)?;
        Ok(Self { bytes })
    }

    /// Audit compact XML character data.
    pub fn text(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let bytes = bytes.into();
        if bytes.is_empty() || bytes.contains(&b'<') {
            return Err(invalid_source("authored XML text fragment is unclassified"));
        }
        audit_wrapped_fragment(&bytes)?;
        Ok(Self { bytes })
    }

    /// Audit compact XML character data while charging its temporary wrapper
    /// and parser workspace to an explicit execution context before
    /// allocation. The caller-owned input vector and returned fragment bytes
    /// are outside this scoped temporary reservation.
    pub fn text_with_execution_context(bytes: Vec<u8>, context: &ExecutionContext) -> Result<Self> {
        if bytes.is_empty() || bytes.contains(&b'<') {
            return Err(invalid_source("authored XML text fragment is unclassified"));
        }
        audit_wrapped_fragment_with_context(&bytes, context)?;
        Ok(Self { bytes })
    }
}

/// A checked set of bounded edits to one exact source XML Part.
pub struct XmlSplicePublication {
    source: SourceXmlPart,
    edits: Vec<XmlSpliceEdit>,
}

impl fmt::Debug for XmlSplicePublication {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("XmlSplicePublication")
            .field("source", &self.source)
            .field("edit_count", &self.edits.len())
            .finish()
    }
}

#[derive(Debug)]
struct XmlSpliceEdit {
    range: Range<usize>,
    fragment: AuthoredXmlFragment,
    _reservation: Option<Arc<Reservation>>,
}

impl XmlSplicePublication {
    /// Stage one checked replacement or insertion.
    pub fn replace(&mut self, proof: XmlSourceRange, fragment: AuthoredXmlFragment) -> Result<()> {
        self.source.check_source_state()?;
        if !self.source.is_original_payload() {
            return Err(invalid_source(
                "derived source XML cannot receive another source splice",
            ));
        }
        if proof.source_lineage != self.source.source_lineage
            || proof.source_version != self.source.source_version
            || !proof
                .source_partname
                .is_equivalent_to(&self.source.source_partname)
        {
            return Err(invalid_source("XML source range has different provenance"));
        }
        let actual = self
            .source
            .original
            .as_bytes()
            .get(proof.range.clone())
            .ok_or_else(|| invalid_source("invalid XML source range"))?;
        if actual != proof.expected.as_slice() {
            return Err(invalid_source("stale XML source range"));
        }
        if actual == fragment.bytes.as_slice() {
            return Ok(());
        }
        if self
            .edits
            .iter()
            .any(|edit| ranges_overlap_or_conflict(&edit.range, &proof.range))
        {
            return Err(invalid_source("overlapping XML source ranges"));
        }
        consume_work(&self.source.source, fragment.bytes.len())?;
        let capacity_metadata = if self.edits.len() == self.edits.capacity() {
            self.edits
                .capacity()
                .max(4)
                .checked_mul(size_of::<XmlSpliceEdit>())
                .ok_or_else(|| invalid_source("XML source splice capacity overflows"))?
        } else {
            0
        };
        let edit_metadata = size_of::<XmlSpliceEdit>()
            .checked_add(capacity_metadata)
            .and_then(|bytes| {
                bytes.checked_add(if self.edits.is_empty() {
                    size_of::<Vec<XmlSpliceEdit>>()
                } else {
                    0
                })
            })
            .ok_or_else(|| invalid_source("XML source splice metadata size overflows"))?;
        let edit_reservation = reserve_edit_memory(
            &self.source.source,
            fragment.bytes.capacity(),
            edit_metadata,
        )?;
        self.edits
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC XML splice edits",
                source,
            })?;
        self.edits.push(XmlSpliceEdit {
            range: proof.range,
            fragment,
            _reservation: edit_reservation,
        });
        self.source.check_source_state()?;
        Ok(())
    }

    /// Finish the checked transaction while preserving every untouched source
    /// byte and validating the assembled XML with source-compatible rules.
    pub fn finish(mut self) -> Result<SourceXmlPart> {
        self.source.check_source_state()?;
        if self.edits.is_empty() {
            return Ok(self.source);
        }
        self.edits.sort_unstable_by_key(|edit| edit.range.start);
        let original = self.source.original.as_bytes();
        let mut removed = 0usize;
        let mut added = 0usize;
        for (index, edit) in self.edits.iter().enumerate() {
            if index & 0x3f == 0 {
                self.source.check_source_state()?;
            }
            let range_len = edit
                .range
                .end
                .checked_sub(edit.range.start)
                .ok_or_else(|| invalid_source("XML source range is reversed"))?;
            removed = removed
                .checked_add(range_len)
                .ok_or_else(|| invalid_source("XML source splice size overflows"))?;
            added = added
                .checked_add(edit.fragment.bytes.len())
                .ok_or_else(|| invalid_source("XML source splice size overflows"))?;
            if edit.range.end > original.len() {
                return Err(invalid_source(
                    "XML source range is outside the source Part",
                ));
            }
        }
        let output_len = original
            .len()
            .checked_sub(removed)
            .and_then(|length| length.checked_add(added))
            .ok_or_else(|| invalid_source("XML source splice size overflows"))?;
        self.source.limits.check(
            ReadResource::PartBytes,
            output_len as u64,
            self.source.limits.max_part_bytes(),
        )?;
        consume_work(&self.source.source, output_len)?;
        let output_reservation = reserve_output_memory(&self.source.source, output_len)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC XML splice output",
                source,
            })?;
        let mut cursor = 0usize;
        for (index, edit) in self.edits.iter().enumerate() {
            if index & 0x3f == 0 {
                self.source.check_source_state()?;
            }
            if edit.range.start < cursor || edit.range.end < edit.range.start {
                return Err(invalid_source("XML source splice ranges overlap"));
            }
            output.extend_from_slice(&original[cursor..edit.range.start]);
            output.extend_from_slice(&edit.fragment.bytes);
            cursor = edit.range.end;
        }
        output.extend_from_slice(&original[cursor..]);
        self.source.check_source_state()?;
        validate_source_xml(
            &self.source.source_partname,
            &output,
            self.source.limits,
            self.source.source.context_ref(),
        )?;
        self.source.check_source_state()?;
        self.source.payload = Arc::new(output);
        self.source.payload_reservation = output_reservation;
        Ok(self.source)
    }
}

/// Classify an XML payload at the OPC authored publication boundary.
///
/// `Ok(false)` means the existing authored compactness contract accepts the
/// payload. `Ok(true)` means the audit stopped at a compactness violation;
/// it does not prove that the remainder is well-formed. The caller must obtain
/// a [`SourceXmlPart`] to validate the complete source before publication.
/// Other audit failures remain errors. This predicate never authorizes bypassing
/// the ordinary authoring gate.
pub fn authored_xml_requires_source_proof(
    partname: &PackURI,
    content_type: &str,
    bytes: &[u8],
) -> Result<bool> {
    if !xml_minifier::audit::package::is_xml_part(partname.as_str(), content_type) {
        return Ok(false);
    }
    match xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default()) {
        Ok(_) => Ok(false),
        Err(xml_minifier::audit::Error::NotCompact(_)) => Ok(true),
        Err(source) => Err(OpcError::XmlPublication {
            part: partname.to_string(),
            source,
        }),
    }
}

fn reserve_output_memory(
    source: &SourceSnapshot,
    bytes: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = source.context_ref() else {
        return Ok(None);
    };
    if bytes == 0 {
        return Ok(None);
    }
    let bytes =
        u64::try_from(bytes).map_err(|_| invalid_source("XML source output size exceeds u64"))?;
    context
        .reserve(Resource::Memory, bytes)
        .map(Arc::new)
        .map(Some)
        .map_err(map_execution_error)
}

fn consume_work(source: &SourceSnapshot, bytes: usize) -> Result<()> {
    let Some(context) = source.context_ref() else {
        return Ok(());
    };
    let amount = u64::try_from(bytes)
        .map_err(|_| invalid_source("source-backed OPC XML work size exceeds u64"))?;
    context
        .consume(Resource::Work, amount)
        .map_err(map_execution_error)
}

fn reserve_xml_validation_memory(
    context: Option<&ExecutionContext>,
    bytes: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    // The source payload reservation covers the input bytes. This separate
    // admission bound covers quick-xml 0.41's input buffer, namespace
    // resolver binding Vec, opened-element state Vec, and our per-element
    // duplicate-key bookkeeping before NsReader is constructed. The concrete
    // quick-xml element types are private, so each input-derived slot uses a
    // conservative fixed footprint. The fixed term covers the predefined
    // namespace bindings and the non-empty Vec/reader state for a tiny `<a/>`.
    let source_bytes = bytes.max(1);
    let namespace_slots = source_bytes
        .checked_div(XML_MIN_NAMESPACE_DECLARATION_BYTES)
        .and_then(|slots| slots.checked_add(1))
        .ok_or_else(|| invalid_source("source-backed OPC namespace bound overflows"))?;
    let open_element_slots = source_bytes
        .checked_div(XML_MIN_OPEN_ELEMENT_BYTES)
        .and_then(|slots| slots.checked_add(1))
        .ok_or_else(|| invalid_source("source-backed OPC depth bound overflows"))?;
    let attribute_slots = source_bytes
        .checked_div(XML_MIN_ATTRIBUTE_BYTES)
        .and_then(|slots| slots.checked_add(1))
        .ok_or_else(|| invalid_source("source-backed OPC attribute bound overflows"))?;
    let amount = XML_VALIDATION_FIXED_WORKING_BYTES
        .checked_add(
            source_bytes
                .checked_mul(2)
                .ok_or_else(|| invalid_source("source-backed OPC input bound overflows"))?,
        )
        .and_then(|amount| {
            namespace_slots
                .checked_mul(XML_VALIDATION_NAMESPACE_BINDING_BYTES)
                .and_then(|bound| amount.checked_add(bound))
        })
        .and_then(|amount| {
            open_element_slots
                .checked_mul(XML_VALIDATION_OPEN_ELEMENT_BYTES)
                .and_then(|bound| amount.checked_add(bound))
        })
        .and_then(|amount| {
            attribute_slots
                .checked_mul(XML_VALIDATION_ATTRIBUTE_BYTES)
                .and_then(|bound| amount.checked_add(bound))
        })
        .ok_or_else(|| invalid_source("source-backed OPC XML parser memory overflows"))?;
    let amount = u64::try_from(amount)
        .map_err(|_| invalid_source("source-backed OPC XML parser memory exceeds u64"))?;
    context
        .reserve(Resource::Memory, amount)
        .map(Arc::new)
        .map(Some)
        .map_err(map_execution_error)
}

fn audit_wrapped_fragment(fragment: &[u8]) -> Result<()> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid_source(
            "authored XML fragment exceeds the size limit",
        ));
    }
    let length = wrapped_fragment_len(fragment.len())?;
    audit_wrapped_fragment_with_limits(fragment, length, xml_minifier::audit::Limits::default())
}

fn audit_wrapped_fragment_with_limits(
    fragment: &[u8],
    length: usize,
    limits: xml_minifier::audit::Limits,
) -> Result<()> {
    let mut document = Vec::new();
    document
        .try_reserve_exact(length)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC authored XML fragment",
            source,
        })?;
    document.extend_from_slice(FRAGMENT_ROOT_OPEN);
    document.extend_from_slice(fragment);
    document.extend_from_slice(FRAGMENT_ROOT_CLOSE);
    audit_document_with_limits(&document, limits)
}

fn audit_wrapped_fragment_with_context(fragment: &[u8], context: &ExecutionContext) -> Result<()> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid_source(
            "authored XML fragment exceeds the size limit",
        ));
    }
    let length = wrapped_fragment_len(fragment.len())?;
    let limits = authored_fragment_audit_limits(fragment, length);
    let _admission = reserve_fragment_audit(context, length, limits)?;
    context.check().map_err(map_execution_error)?;
    let result = audit_wrapped_fragment_with_limits(fragment, length, limits);
    context.check().map_err(map_execution_error)?;
    result
}

fn wrapped_fragment_len(fragment_len: usize) -> Result<usize> {
    FRAGMENT_ROOT_OPEN
        .len()
        .checked_add(fragment_len)
        .and_then(|value| value.checked_add(FRAGMENT_ROOT_CLOSE.len()))
        .ok_or_else(|| invalid_source("authored XML fragment size overflows"))
}

fn authored_fragment_audit_limits(fragment: &[u8], bytes: usize) -> xml_minifier::audit::Limits {
    let bytes = bytes.max(1);
    // A valid element can only increase nesting at a `<` marker. Counting all
    // markers, including closing tags and markup inside quoted values, is
    // conservative while avoiding a byte-length-based depth charge for a
    // large flat text node.
    let depth = fragment
        .iter()
        .filter(|byte| **byte == b'<')
        .count()
        .saturating_add(1)
        .min(xml_minifier::audit::Limits::ceiling(
            xml_minifier::audit::Resource::Depth,
        ))
        .max(1);
    xml_minifier::audit::Limits::default()
        .narrow(xml_minifier::audit::Resource::Bytes, bytes)
        .narrow(xml_minifier::audit::Resource::Depth, depth)
        .narrow(
            xml_minifier::audit::Resource::Events,
            bytes.saturating_add(1),
        )
        .narrow(xml_minifier::audit::Resource::Attributes, bytes)
        .narrow(xml_minifier::audit::Resource::TokenBytes, bytes)
        .narrow(xml_minifier::audit::Resource::TextBytes, bytes)
}

fn reserve_fragment_audit(
    context: &ExecutionContext,
    bytes: usize,
    limits: xml_minifier::audit::Limits,
) -> Result<FragmentAuditAdmission> {
    // `verify_authored` is the slice auditor, so it does not allocate the
    // streaming reader's token windows. The wrapper Vec itself is retained
    // only for this call. The multiplier covers parser scratch, the open-space
    // stack, transient decoded attributes, and allocator capacity slack.
    let parser_memory = bytes
        .checked_mul(AUTHORED_AUDIT_MEMORY_MULTIPLIER)
        .and_then(|amount| amount.checked_add(AUTHORED_AUDIT_WORKING_BYTES))
        .ok_or_else(|| invalid_source("authored XML audit memory overflows"))?;
    let memory = bytes
        .checked_add(parser_memory)
        .ok_or_else(|| invalid_source("authored XML audit memory overflows"))?;
    // The XML audit profile maps aggregate attributes to Objects and events
    // to cumulative Work. Add the retained open-element space stack and a
    // small fixed allowance for the reader and State.
    let objects = limits
        .max_attributes()
        .checked_add(limits.max_depth())
        .and_then(|amount| amount.checked_add(8))
        .ok_or_else(|| invalid_source("authored XML audit object count overflows"))?;
    let depth = limits.max_depth();
    let work = limits.max_events();
    let memory = reserve_fragment_audit_resource(context, Resource::Memory, memory)?;
    let objects = reserve_fragment_audit_resource(context, Resource::Objects, objects)?;
    let depth = reserve_fragment_audit_resource(context, Resource::Depth, depth)?;
    let work =
        u64::try_from(work).map_err(|_| invalid_source("authored XML audit work exceeds u64"))?;
    context
        .consume(Resource::Work, work)
        .map_err(map_execution_error)?;
    Ok(FragmentAuditAdmission {
        _memory: memory,
        _objects: objects,
        _depth: depth,
    })
}

fn reserve_fragment_audit_resource(
    context: &ExecutionContext,
    resource: Resource,
    amount: usize,
) -> Result<Reservation> {
    let amount = u64::try_from(amount)
        .map_err(|_| invalid_source("authored XML audit resource amount exceeds u64"))?;
    context
        .reserve(resource, amount)
        .map_err(map_execution_error)
}

fn audit_document_with_limits(document: &[u8], limits: xml_minifier::audit::Limits) -> Result<()> {
    xml_minifier::audit::verify_authored(document, limits)
        .map(|_| ())
        .map_err(|source| OpcError::XmlPublication {
            part: "<source-backed XML fragment>".to_string(),
            source,
        })
}

struct FragmentAuditAdmission {
    _memory: Reservation,
    _objects: Reservation,
    _depth: Reservation,
}

fn ranges_overlap_or_conflict(left: &Range<usize>, right: &Range<usize>) -> bool {
    if left.start == left.end {
        return right.start <= left.start && left.start <= right.end;
    }
    if right.start == right.end {
        return left.start <= right.start && right.start <= left.end;
    }
    left.start < right.end && right.start < left.end
}

fn reserve_edit_memory(
    source: &SourceSnapshot,
    bytes: usize,
    metadata: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = source.context_ref() else {
        return Ok(None);
    };
    let bytes = bytes
        .checked_add(metadata)
        .ok_or_else(|| invalid_source("XML source proof allocation overflows"))?;
    let bytes = u64::try_from(bytes)
        .map_err(|_| invalid_source("XML source proof allocation exceeds u64"))?;
    if bytes == 0 {
        return Ok(None);
    }
    context
        .reserve(Resource::Memory, bytes)
        .map(Arc::new)
        .map(Some)
        .map_err(map_execution_error)
}

fn limits_range_max(limits: ReadLimits) -> u64 {
    limits.max_xml_attribute_bytes() as u64
}

fn validate_source_xml(
    partname: &PackURI,
    bytes: &[u8],
    limits: ReadLimits,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    limits.check(
        ReadResource::PartBytes,
        bytes.len() as u64,
        limits.max_part_bytes(),
    )?;
    std::str::from_utf8(bytes).map_err(|error| {
        OpcError::XmlError(format!("source XML '{}' is not UTF-8: {error}", partname))
    })?;
    consume_work_from_context(context, bytes.len())?;
    let _validation_memory = reserve_xml_validation_memory(context, bytes.len())?;
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.config_mut().trim_markup_names_in_closing_tags = false;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_SOURCE_XML_NAMESPACE_DECLARATIONS_PER_ELEMENT);
    // quick-xml's namespace resolver stores its nesting level as `u16` and
    // increments it before returning a start event. Clamp one slot below its
    // maximum so the next event can be observed and refused without wrapping
    // the resolver counter before our own depth accounting sees it.
    let xml_depth_limit = limits.max_xml_depth().min(QUICK_XML_MAX_SAFE_DEPTH);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut events = 0u64;
    let mut declaration_seen = false;
    loop {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_source("source XML event count overflows"))?;
        limits.check(
            ReadResource::XmlEvents,
            events,
            limits.max_xml_events() as u64,
        )?;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            OpcError::XmlError(format!("source XML '{partname}' is malformed: {error}"))
        })?;
        if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(OpcError::XmlError(format!(
                "source XML '{partname}' contains an unbound namespace prefix"
            )));
        }
        match event {
            Event::Start(element) => {
                validate_source_element(&reader, &element, limits, partname)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid_source("source XML root count overflows"))?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_source("source XML depth overflows"))?;
                limits.check(ReadResource::XmlDepth, depth as u64, xml_depth_limit as u64)?;
            },
            Event::Empty(element) => {
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_source("source XML depth overflows"))?;
                limits.check(
                    ReadResource::XmlDepth,
                    element_depth as u64,
                    xml_depth_limit as u64,
                )?;
                validate_source_element(&reader, &element, limits, partname)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid_source("source XML root count overflows"))?;
                }
            },
            Event::End(element) => {
                validate_source_qname(&reader, element.name().as_ref(), "end element")?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_source("source XML has an unexpected end tag"))?;
            },
            Event::Text(text) => {
                if text.as_ref().windows(3).any(|window| window == b"]]>") {
                    return Err(invalid_source(
                        "source XML text contains the forbidden ']]>' sequence",
                    ));
                }
                let text = text.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_source(format!("source XML text is not decodable: {error}"))
                })?;
                validate_xml_string(&text)?;
                if depth == 0 && !text.chars().all(is_xml_whitespace) {
                    return Err(invalid_source("source XML has text outside its root"));
                }
            },
            Event::CData(data) => {
                if data.as_ref().windows(3).any(|window| window == b"]]>") {
                    return Err(invalid_source(
                        "source XML CDATA contains the forbidden ']]>' sequence",
                    ));
                }
                let data = data.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_source(format!("source XML CDATA is not decodable: {error}"))
                })?;
                validate_xml_string(&data)?;
                if depth == 0 {
                    return Err(invalid_source("source XML has CDATA outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                validate_general_reference(&reference)?;
                if depth == 0 {
                    return Err(invalid_source(
                        "source XML has a reference outside its root",
                    ));
                }
            },
            Event::Decl(declaration) => {
                if declaration_seen || roots != 0 || events != 1 {
                    return Err(invalid_source("source XML declaration is out of place"));
                }
                validate_source_declaration(&declaration, partname.as_str())?;
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(invalid_source("source XML DTDs are not permitted"));
            },
            Event::Eof => break,
            Event::Comment(comment) => {
                validate_comment(&comment)?;
            },
            Event::PI(instruction) => {
                validate_processing_instruction(&reader, &instruction)?;
            },
        }
    }
    if depth != 0 || roots != 1 {
        return Err(invalid_source(format!(
            "source XML '{partname}' must contain one closed root element"
        )));
    }
    Ok(())
}

fn consume_work_from_context(context: Option<&ExecutionContext>, bytes: usize) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    let amount = u64::try_from(bytes)
        .map_err(|_| invalid_source("source-backed OPC XML work size exceeds u64"))?;
    context
        .consume(Resource::Work, amount)
        .map_err(map_execution_error)
}

pub(crate) fn validate_source_declaration(
    declaration: &BytesDecl<'_>,
    partname: &str,
) -> Result<()> {
    let content = std::str::from_utf8(declaration.as_ref()).map_err(|error| {
        invalid_source(format!(
            "source XML '{partname}' declaration is not UTF-8: {error}"
        ))
    })?;
    // BytesDecl intentionally exposes convenience accessors but does not
    // enforce the XML declaration grammar. Reparse its borrowed content as a
    // BytesStart so duplicate attributes and ordering are checked before the
    // source token is issued.
    let start = BytesStart::from_content(content, 3);
    let mut seen_version = false;
    let mut seen_encoding = false;
    let mut seen_standalone = false;
    let mut attribute_count = 0usize;
    for attribute_result in start.attributes().with_checks(true) {
        attribute_count = attribute_count
            .checked_add(1)
            .ok_or_else(|| invalid_source("source XML declaration attribute count overflows"))?;
        if attribute_count > 3 {
            return Err(invalid_source(
                "source XML declaration has too many attributes",
            ));
        }
        let attribute = attribute_result.map_err(|error| {
            invalid_source(format!(
                "source XML declaration attribute is invalid: {error}"
            ))
        })?;
        let key = attribute.key.0;
        let raw_value = attribute.value.as_ref();
        if raw_value.contains(&b'<') || raw_value.contains(&b'&') {
            return Err(invalid_source(
                "source XML declaration contains a forbidden raw attribute character",
            ));
        }
        let value = std::str::from_utf8(raw_value).map_err(|error| {
            invalid_source(format!(
                "source XML declaration attribute value is invalid: {error}"
            ))
        })?;
        validate_xml_string(value)?;
        match key {
            b"version" if !seen_version && !seen_encoding && !seen_standalone => {
                seen_version = true;
                if value != "1.0" {
                    return Err(invalid_source("source XML declaration version must be 1.0"));
                }
            },
            b"encoding" if seen_version && !seen_encoding && !seen_standalone => {
                seen_encoding = true;
                if !value.eq_ignore_ascii_case("UTF-8") {
                    return Err(invalid_source(
                        "source XML declaration encoding must be UTF-8",
                    ));
                }
            },
            b"standalone" if seen_version && !seen_standalone => {
                seen_standalone = true;
                if value != "yes" && value != "no" {
                    return Err(invalid_source(
                        "source XML declaration standalone must be yes or no",
                    ));
                }
            },
            _ => {
                return Err(invalid_source(
                    "source XML declaration has a duplicate or out-of-order attribute",
                ));
            },
        }
    }
    if !seen_version {
        return Err(invalid_source(
            "source XML declaration is missing its version",
        ));
    }
    Ok(())
}

fn validate_source_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: ReadLimits,
    partname: &PackURI,
) -> Result<()> {
    validate_source_qname(reader, element.name().as_ref(), "element")?;
    let mut attribute_count = 0usize;
    let mut seen_keys: Vec<&[u8]> = Vec::new();
    let mut seen_expanded: Vec<(Option<&[u8]>, &[u8])> = Vec::new();
    for attribute_result in element.attributes().with_checks(true) {
        attribute_count = attribute_count
            .checked_add(1)
            .ok_or_else(|| invalid_source("source XML attribute count overflows"))?;
        if attribute_count > MAX_SOURCE_XML_ATTRIBUTES_PER_ELEMENT {
            return Err(invalid_source(format!(
                "source XML element in '{}' has too many attributes",
                partname
            )));
        }
        let attribute: Attribute<'_> = attribute_result.map_err(|error| {
            invalid_source(format!("source XML attribute is malformed: {error}"))
        })?;
        let key = attribute.key.0;
        if seen_keys.contains(&key) {
            return Err(invalid_source("source XML contains duplicate attributes"));
        }
        seen_keys
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC XML attribute names",
                source,
            })?;
        seen_keys.push(key);
        let raw_attribute_bytes = key
            .len()
            .checked_add(attribute.value.as_ref().len())
            .ok_or_else(|| invalid_source("source XML attribute size overflows"))?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            raw_attribute_bytes as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        let key_name = reader.decoder().decode(key).map_err(|error| {
            invalid_source(format!("source XML attribute name is invalid: {error}"))
        })?;
        validate_xml_qname(&key_name, "attribute")?;
        if attribute.value.as_ref().contains(&b'<') {
            return Err(invalid_source(
                "source XML attribute contains a literal '<'",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                invalid_source(format!("source XML attribute value is invalid: {error}"))
            })?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            value.len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        validate_xml_string(&value)?;
        if key_name == "xmlns" || key_name.starts_with("xmlns:") {
            // quick-xml's resolver retains raw namespace URI bytes. Refuse
            // bindings requiring normalization so expanded-name comparisons
            // cannot miss aliases expressed with character references.
            if attribute.value.as_ref() != value.as_bytes() {
                return Err(invalid_source(
                    "source XML requires literal namespace bindings",
                ));
            }
            validate_namespace_binding(&key_name, &value)?;
        } else {
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = match namespace {
                ResolveResult::Bound(quick_xml::name::Namespace(uri)) => Some(uri),
                ResolveResult::Unbound => None,
                ResolveResult::Unknown(_) => {
                    return Err(invalid_source(
                        "source XML attribute uses an unbound namespace prefix",
                    ));
                },
            };
            if seen_expanded.iter().any(|(known_namespace, known_local)| {
                *known_namespace == namespace && *known_local == local.as_ref()
            }) {
                return Err(invalid_source(
                    "source XML contains duplicate expanded attributes",
                ));
            }
            seen_expanded
                .try_reserve(1)
                .map_err(|source| OpcError::Allocation {
                    resource: "source-backed OPC expanded attribute names",
                    source,
                })?;
            seen_expanded.push((namespace, local.into_inner()));
        }
    }
    Ok(())
}

fn validate_source_qname(reader: &NsReader<&[u8]>, raw: &[u8], kind: &str) -> Result<()> {
    let name = reader
        .decoder()
        .decode(raw)
        .map_err(|error| invalid_source(format!("source XML {kind} name is invalid: {error}")))?;
    validate_xml_qname(&name, kind)?;
    if name
        .split_once(':')
        .is_some_and(|(prefix, _)| prefix == "xmlns")
    {
        return Err(invalid_source(format!(
            "source XML {kind} uses the reserved xmlns prefix"
        )));
    }
    if matches!(
        reader
            .resolver()
            .resolve_element(quick_xml::name::QName(raw))
            .0,
        ResolveResult::Unknown(_)
    ) {
        return Err(invalid_source(format!(
            "source XML {kind} uses an unbound namespace prefix"
        )));
    }
    Ok(())
}

fn validate_xml_qname(value: &str, kind: &str) -> Result<()> {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return Err(invalid_source(format!("source XML {kind} QName is empty")));
    };
    let valid = if !is_xml_id(first) {
        false
    } else {
        match parts.next() {
            None => true,
            Some(second) => is_xml_id(second) && parts.next().is_none(),
        }
    };
    if valid {
        Ok(())
    } else {
        Err(invalid_source(format!(
            "source XML {kind} name is not a valid QName"
        )))
    }
}

fn validate_namespace_binding(name: &str, value: &str) -> Result<()> {
    let prefix = name.strip_prefix("xmlns:");
    if name != "xmlns" && prefix.is_none_or(|prefix| !is_xml_id(prefix)) {
        return Err(invalid_source(
            "source XML namespace declaration name is invalid",
        ));
    }
    if prefix == Some("xmlns")
        || value == XMLNS_NAMESPACE
        || (prefix == Some("xml")) != (value == XML_NAMESPACE)
        || (prefix.is_some() && value.is_empty())
    {
        return Err(invalid_source(
            "source XML has an invalid reserved namespace binding",
        ));
    }
    Ok(())
}

fn validate_xml_string(value: &str) -> Result<()> {
    if value.chars().all(xml10_character) {
        Ok(())
    } else {
        Err(invalid_source(
            "source XML contains a forbidden XML 1.0 character",
        ))
    }
}

fn xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&value)
        || ('\u{e000}'..='\u{fffd}').contains(&value)
        || ('\u{10000}'..='\u{10ffff}').contains(&value)
}

fn is_xml_whitespace(value: char) -> bool {
    matches!(value, ' ' | '\t' | '\r' | '\n')
}

fn validate_general_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
    let name = reference
        .decode()
        .map_err(|error| invalid_source(format!("source XML reference is invalid: {error}")))?;
    match name.as_ref() {
        "amp" | "lt" | "gt" | "apos" | "quot" => Ok(()),
        value if value.starts_with('#') => {
            let character = reference
                .resolve_char_ref()
                .map_err(|error| {
                    invalid_source(format!(
                        "source XML character reference is invalid: {error}"
                    ))
                })?
                .ok_or_else(|| invalid_source("source XML character reference is invalid"))?;
            if xml10_character(character) {
                Ok(())
            } else {
                Err(invalid_source(
                    "source XML character reference uses a forbidden character",
                ))
            }
        },
        _ => Err(invalid_source(
            "source XML contains a reference that requires a DTD",
        )),
    }
}

fn validate_comment(comment: &quick_xml::events::BytesText<'_>) -> Result<()> {
    let bytes = comment.as_ref();
    if bytes.windows(2).any(|pair| pair == b"--") || bytes.last() == Some(&b'-') {
        return Err(invalid_source(
            "source XML comment contains a forbidden '--'",
        ));
    }
    let value = comment
        .decode()
        .map_err(|error| invalid_source(format!("source XML comment is invalid: {error}")))?;
    validate_xml_string(&value)
}

fn validate_processing_instruction(
    reader: &NsReader<&[u8]>,
    instruction: &quick_xml::events::BytesPI<'_>,
) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| invalid_source(format!("source XML PI target is invalid: {error}")))?;
    validate_xml_qname(&target, "processing-instruction")?;
    if target.eq_ignore_ascii_case("xml") {
        return Err(invalid_source(
            "source XML processing-instruction target is reserved",
        ));
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| invalid_source(format!("source XML PI content is invalid: {error}")))?;
    validate_xml_string(&content)
}

fn invalid_source(reason: impl Into<String>) -> OpcError {
    OpcError::SourceBackedOverlayUnavailable {
        reason: reason.into(),
    }
}

fn map_execution_error(error: ExecutionError) -> OpcError {
    match error {
        ExecutionError::Cancelled => OpcError::Cancelled,
        other => OpcError::Execution(other),
    }
}

impl SourceXmlPart {
    pub(crate) fn from_source_parts(
        source: SourceSnapshot,
        limits: ReadLimits,
        source_partname: &PackURI,
        source_content_type: &str,
        original: PartData,
    ) -> Result<Self> {
        Self::new(
            source,
            limits,
            source_partname,
            source_content_type,
            original,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits as BudgetLimits};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
        managed_context_with_depth(memory, u64::MAX)
    }

    fn managed_context_with_depth(
        memory: u64,
        depth: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "xml-splice-test",
            BudgetLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, depth, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("non-zero worker count"),
            NonZeroUsize::new(1).expect("non-zero task count"),
            NonZeroU64::new(memory.max(1)).expect("non-zero in-flight bytes"),
            0,
        )
        .expect("valid execution limits");
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    fn test_partname() -> PackURI {
        PackURI::new("/word/document.xml").expect("valid test Part URI")
    }

    fn source_is_valid(bytes: &[u8]) -> bool {
        validate_source_xml(&test_partname(), bytes, ReadLimits::default(), None).is_ok()
    }

    #[test]
    fn source_validator_allows_unprefixed_xml_and_utf8_declaration() {
        assert!(source_is_valid(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root><child/></root>"#
        ));
        assert!(source_is_valid(br#"<root><child/></root>"#));
    }

    #[test]
    fn source_validator_rejects_namespace_and_lexical_ambiguities() {
        let invalid = [
            br#"<bad?:root/>"#.as_slice(),
            br#"<p:root/>"#.as_slice(),
            br#"<xmlns:root/>"#.as_slice(),
            br#"<root xmlns:p="urn:test" xmlns:q="urn:&#116;est" p:x="1" q:x="2"/>"#.as_slice(),
            br#"<root xmlns:p="urn:test" xmlns:q="urn:test" p:id="1" q:id="2"/>"#.as_slice(),
            br#"<root value="literal < bracket"/>"#.as_slice(),
            br#"<root><!--bad--comment--></root>"#.as_slice(),
            br#"<root>text]]>more</root>"#.as_slice(),
            br#"<?xml version="1.0" encoding="UTF-16"?><root/>"#.as_slice(),
            br#"<?xml version="1.1"?><root/>"#.as_slice(),
            br#"<?xml version="1.&#48;"?><root/>"#.as_slice(),
            br#"<?xml encoding="UTF-8" version="1.0"?><root/>"#.as_slice(),
            br#"<root>&external;</root>"#.as_slice(),
        ];
        for bytes in invalid {
            assert!(!source_is_valid(bytes), "unexpectedly admitted {bytes:?}");
        }
    }

    #[test]
    fn authored_classifier_keeps_compact_fast_path_and_requires_source_for_whitespace() {
        let partname = test_partname();
        let compact = authored_xml_requires_source_proof(
            &partname,
            "application/xml",
            br#"<root><child/></root>"#,
        )
        .expect("compact authored XML should classify");
        assert!(!compact);
        let formatted = authored_xml_requires_source_proof(
            &partname,
            "application/xml",
            b"<root>\n  <child/>\n</root>",
        )
        .expect("well-formed formatted XML should require source proof");
        assert!(formatted);
        assert!(
            authored_xml_requires_source_proof(&partname, "application/xml", b"<root").is_err()
        );
        let malformed_after_whitespace = b"<root>\n <unclosed>";
        assert!(
            authored_xml_requires_source_proof(
                &partname,
                "application/xml",
                malformed_after_whitespace,
            )
            .unwrap()
        );
        assert!(!source_is_valid(malformed_after_whitespace));
    }

    #[test]
    fn empty_elements_obey_depth_and_attributes_are_limited_individually() {
        let shallow = ReadLimits::builder()
            .max_xml_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            validate_source_xml(&test_partname(), b"<root><leaf/></root>", shallow, None).is_err()
        );
        assert!(validate_source_xml(&test_partname(), b"<root/>", shallow, None).is_ok());
        let attributes = ReadLimits::builder()
            .max_xml_attribute_bytes(5)
            .unwrap()
            .max_relationship_target_bytes(5)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            validate_source_xml(
                &test_partname(),
                br#"<root a="1234" b="5678"/>"#,
                attributes,
                None
            )
            .is_ok()
        );
        assert!(
            validate_source_xml(&test_partname(), br#"<root a="12345"/>"#, attributes, None)
                .is_err()
        );
    }

    #[test]
    fn insertion_and_replacement_ranges_touching_are_conflicts() {
        assert!(ranges_overlap_or_conflict(&(4..4), &(4..8)));
        assert!(ranges_overlap_or_conflict(&(4..8), &(4..4)));
        assert!(ranges_overlap_or_conflict(&(4..4), &(4..4)));
        assert!(!ranges_overlap_or_conflict(&(0..4), &(4..8)));
    }

    #[test]
    fn contextual_authored_audit_releases_temporary_resources_on_success_and_error() {
        let (budget, _cancellation, context) = managed_context(1024 * 1024);

        AuthoredXmlFragment::markup_with_execution_context(b"<tag/>".to_vec(), &context)
            .expect("compact markup should pass the managed audit");
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);
        assert!(budget.used(Resource::Work) > 0);

        AuthoredXmlFragment::text_with_execution_context(b"text".to_vec(), &context)
            .expect("compact text should pass the managed audit");
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);

        assert!(
            AuthoredXmlFragment::markup_with_execution_context(b"<unclosed>".to_vec(), &context,)
                .is_err()
        );
        assert!(
            AuthoredXmlFragment::text_with_execution_context(b"a & b".to_vec(), &context,).is_err()
        );
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);
    }

    #[test]
    fn contextual_authored_audit_rejects_low_memory_before_wrapper_allocation() {
        let (budget, _cancellation, context) = managed_context(1);
        let error =
            AuthoredXmlFragment::markup_with_execution_context(b"<tag/>".to_vec(), &context)
                .expect_err("the temporary audit envelope must exceed one byte");
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        let text_error =
            AuthoredXmlFragment::text_with_execution_context(b"text".to_vec(), &context)
                .expect_err("the temporary text audit envelope must exceed one byte");
        assert!(matches!(
            text_error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);
        assert_eq!(budget.used(Resource::Work), 0);
    }

    #[test]
    fn contextual_authored_audit_honors_cancellation_without_leaking_reservations() {
        let (budget, cancellation, context) = managed_context(1024 * 1024);
        cancellation.cancel();

        let markup =
            AuthoredXmlFragment::markup_with_execution_context(b"<tag/>".to_vec(), &context);
        assert!(matches!(markup, Err(OpcError::Cancelled)));
        let text = AuthoredXmlFragment::text_with_execution_context(b"text".to_vec(), &context);
        assert!(matches!(text, Err(OpcError::Cancelled)));
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);
        assert_eq!(budget.used(Resource::Work), 0);
    }

    #[test]
    fn contextual_authored_audit_charges_depth_from_markup_shape() {
        let (budget, _cancellation, context) = managed_context_with_depth(2 * 1024 * 1024, 64);
        let flat_text = vec![b'x'; 10_000];
        AuthoredXmlFragment::text_with_execution_context(flat_text, &context)
            .expect("a large flat text node only needs the wrapper depth");
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);

        let mut nested = Vec::new();
        for _ in 0..70 {
            nested.extend_from_slice(b"<a>");
        }
        for _ in 0..70 {
            nested.extend_from_slice(b"</a>");
        }
        let error = AuthoredXmlFragment::markup_with_execution_context(nested, &context)
            .expect_err("nested markup must exceed the managed depth budget");
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Depth
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        assert_eq!(budget.used(Resource::Depth), 0);
    }
}

#[cfg(test)]
#[path = "publication_proof_tests.rs"]
mod publication_proof_tests;
