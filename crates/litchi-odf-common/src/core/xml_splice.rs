//! Provenance-bearing publication of bounded edits to source-loaded XML parts.

use litchi_core::{Error, Result};
use quick_xml::name::PrefixDeclaration;
use quick_xml::{Reader, XmlVersion, events::Event};
use std::{collections::HashSet, io::Write, ops::Range, sync::Arc};

use super::binding_tracker::BindingTracker;
use super::stream_xml::{
    validate_decl_attributes, validate_end_element, validate_event_bytes, validate_name,
    validate_start_element,
};
use super::{OwnedPackage, PackageWriter};

const MAX_PART_BYTES: usize = 256 * 1024 * 1024;
const MAX_SOURCE_EVENTS: usize = 4_000_000;
const MAX_SOURCE_DEPTH: usize = 4_096;
const MAX_SOURCE_ATTRIBUTES: usize = 1_000_000;
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";
const FRAGMENT_ROOT_OPEN: &[u8] = b"<litchi-fragment>";
const FRAGMENT_ROOT_CLOSE: &[u8] = b"</litchi-fragment>";

/// Exact XML bytes loaded from one entry in one owned ODF package.
#[derive(Clone, Debug)]
pub struct XmlSourcePart {
    archive: Arc<Vec<u8>>,
    bytes: Arc<Vec<u8>>,
    media_type: String,
    path: String,
}

/// An opaque byte-range proof issued by an [`XmlSourcePart`].
#[derive(Clone, Debug)]
pub struct XmlSourceRange {
    archive: Arc<Vec<u8>>,
    expected: Vec<u8>,
    path: String,
    range: Range<usize>,
}

/// Individually audited bytes that may replace a checked source range.
#[derive(Clone, Debug)]
pub struct AuthoredXmlFragment {
    bytes: Vec<u8>,
}

/// A checked set of byte-minimal edits to one exact source-loaded XML part.
#[derive(Debug)]
#[allow(
    clippy::module_name_repetitions,
    reason = "The public name distinguishes a publishable transaction from source parts and ranges."
)]
pub struct XmlSplicePublication {
    edits: Vec<Edit>,
    source: XmlSourcePart,
    source_candidate: Option<Vec<u8>>,
}

#[derive(Debug)]
struct Edit {
    fragment: AuthoredXmlFragment,
    range: Range<usize>,
}

impl XmlSourcePart {
    /// Load one exact XML-classified part from `source`.
    ///
    /// XML classification includes conventional XML/RDF paths, signature
    /// paths, and entries whose manifest media type is XML or ends in `+xml`.
    ///
    /// # Errors
    ///
    /// Returns an error when the part is absent, is not XML-classified, its
    /// declared materialized size exceeds the hard bound, or it is not a
    /// well-formed XML document. The declared member size is checked before
    /// archive payload extraction.
    pub fn load(source: &OwnedPackage, path: impl Into<String>) -> Result<Self> {
        let part_path = path.into();
        let package = source.package()?;
        let media_type = package
            .manifest()
            .get_media_type(&part_path)
            .unwrap_or_else(|| guess_media_type(&part_path))
            .to_string();
        if !xml_minifier::audit::package::is_xml_part(&part_path, &media_type) {
            return invalid(format!(
                "ODF splice source '{part_path}' is not an XML part"
            ));
        }
        let declared_size = package.member_materialized_size(&part_path)?;
        if declared_size.is_some_and(|size| size > MAX_PART_BYTES as u64) {
            return invalid(format!(
                "ODF splice source '{part_path}' exceeds the size limit"
            ));
        }
        let bytes = package.get_file(&part_path)?;
        if bytes.len() > MAX_PART_BYTES {
            return invalid(format!(
                "ODF splice source '{part_path}' exceeds the size limit"
            ));
        }
        verify_well_formed(&bytes, &part_path)?;
        Ok(Self {
            archive: source.shared_bytes(),
            bytes: Arc::new(bytes),
            media_type,
            path: part_path,
        })
    }

    /// Return the exact source-loaded bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Return this part's package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Issue a range proof after comparing the caller's expected source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for reversed/out-of-bounds ranges or stale expected
    /// bytes.
    pub fn checked_range(
        &self,
        range: Range<usize>,
        expected_source: &[u8],
    ) -> Result<XmlSourceRange> {
        let actual = self
            .bytes
            .get(range.clone())
            .ok_or_else(|| Error::InvalidFormat("invalid XML splice source range".to_string()))?;
        if actual != expected_source {
            return invalid("stale XML splice source range");
        }
        Ok(XmlSourceRange {
            archive: Arc::clone(&self.archive),
            expected: expected_source.to_vec(),
            path: self.path.clone(),
            range,
        })
    }
}

impl AuthoredXmlFragment {
    /// Return the audited authored bytes.
    ///
    /// The returned slice is the exact compact fragment supplied to the
    /// constructor.  It is intended for bounded streaming composition; the
    /// source document remains free to retain its original lexical bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the capacity retained by the audited fragment allocation.
    ///
    /// This crate-private measure is used by bounded source-publication plans
    /// because a plan retains the backing allocation until it is dropped.
    #[must_use]
    pub(crate) fn retention_bytes(&self) -> usize {
        self.bytes.capacity()
    }

    /// Audit one or more balanced markup nodes as compact authored XML.
    ///
    /// # Errors
    ///
    /// Returns an error when `bytes` are empty, not markup, malformed, contain
    /// a doctype, or violate the compact XML publication contract.
    pub fn markup(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let fragment = bytes.into();
        if fragment.first() != Some(&b'<') {
            return invalid("authored XML markup fragment is unclassified");
        }
        audit_wrapped_fragment(&fragment)?;
        Ok(Self { bytes: fragment })
    }

    /// Audit a non-empty XML start tag as compact authored XML.
    ///
    /// This classification is intended for lexical attribute edits that
    /// replace an entire source start tag.
    ///
    /// # Errors
    ///
    /// Returns an error when `bytes` are not exactly one compact non-empty
    /// start tag.
    pub fn start_tag(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let fragment = bytes.into();
        let name = start_tag_name(&fragment)?;
        let mut document = Vec::with_capacity(fragment.len() + name.len() + 3);
        document.extend_from_slice(&fragment);
        document.extend_from_slice(b"</");
        document.extend_from_slice(name);
        document.push(b'>');
        audit_document(&document)?;
        Ok(Self { bytes: fragment })
    }

    /// Audit one compact XML end tag.
    ///
    /// This classification is intended for lexical splices whose token diff
    /// isolates the closing half of an otherwise balanced authored element.
    /// The assembled source part is still verified as a complete document.
    ///
    /// # Errors
    ///
    /// Returns an error when `bytes` are not exactly one compact end tag.
    pub fn end_tag(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let fragment = bytes.into();
        let name = end_tag_name(&fragment)?;
        let mut document = Vec::with_capacity(fragment.len() + name.len() + 2);
        document.push(b'<');
        document.extend_from_slice(name);
        document.push(b'>');
        document.extend_from_slice(&fragment);
        audit_document(&document)?;
        Ok(Self { bytes: fragment })
    }

    /// Audit escaped character data as compact authored XML.
    ///
    /// # Errors
    ///
    /// Returns an error when `bytes` are empty, include markup, are malformed,
    /// or contain unclassifiable whitespace-only content.
    pub fn text(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let fragment = bytes.into();
        if fragment.is_empty() || fragment.contains(&b'<') {
            return invalid("authored XML text fragment is unclassified");
        }
        audit_wrapped_fragment(&fragment)?;
        Ok(Self { bytes: fragment })
    }

    /// Create the explicitly classified empty fragment used for deletion.
    #[must_use]
    pub const fn deletion() -> Self {
        Self { bytes: Vec::new() }
    }
}

impl XmlSplicePublication {
    /// Begin a publication transaction over one exact source-loaded part.
    #[must_use]
    pub const fn new(source: XmlSourcePart) -> Self {
        Self {
            edits: Vec::new(),
            source,
            source_candidate: None,
        }
    }

    /// Create a source-qualified replacement for one root opening tag.
    ///
    /// `proof` must cover exactly one source `Start` or `Empty` token issued
    /// by this same [`XmlSourcePart`]. The complete candidate may change that
    /// opening token, but every byte before and after it must remain identical
    /// to the source. Its root qualified name and start/empty shape are also
    /// retained. This is the narrow source-preserving seam used by metadata
    /// owners; arbitrary complete-document replacements are intentionally not
    /// admitted here.
    ///
    /// The candidate is fully validated before publication, including XML
    /// declaration grammar, namespace bindings, qualified attributes,
    /// references, roots, comments, processing instructions, and CDATA.
    /// `maximum` bounds both the candidate length and the `Vec` capacity that
    /// this publication retains until it is assembled.
    ///
    /// # Errors
    ///
    /// Returns an error when the proof is foreign or stale, does not cover one
    /// opening token, the candidate changes bytes outside that token, the root
    /// shape changes, the candidate is oversized, or XML validation fails.
    pub fn from_source_start_tag_candidate_with_limit(
        source: XmlSourcePart,
        proof: XmlSourceRange,
        candidate: Vec<u8>,
        maximum: usize,
    ) -> Result<Self> {
        let maximum = maximum.min(MAX_PART_BYTES);
        if candidate.len() > maximum || candidate.capacity() > maximum {
            return invalid("source XML candidate exceeds the size limit");
        }
        let original_range = validate_opening_proof(&source, &proof)?;
        let candidate_range =
            validate_candidate_opening(&source, original_range.clone(), &candidate)?;
        validate_candidate_outside_opening(&source, original_range, candidate_range, &candidate)?;
        verify_source_document(&candidate, source.path())?;
        Ok(Self {
            edits: Vec::new(),
            source,
            source_candidate: Some(candidate),
        })
    }

    /// Create a source-qualified expansion of one empty root element.
    ///
    /// `proof` must cover one source `Empty` token. The candidate replaces
    /// that token with the same QName and exact opening attributes as a
    /// non-empty `Start` token, followed by bounded child markup and its
    /// matching `End` token. Bytes outside the proven empty token remain
    /// identical to the source. This is the only source seam that may turn a
    /// self-closing metadata container into a child-bearing container.
    /// `maximum` bounds both the candidate length and the `Vec` capacity that
    /// this publication retains until it is assembled.
    pub fn from_source_empty_expansion_candidate_with_limit(
        source: XmlSourcePart,
        proof: XmlSourceRange,
        candidate: Vec<u8>,
        maximum: usize,
    ) -> Result<Self> {
        let maximum = maximum.min(MAX_PART_BYTES);
        if candidate.len() > maximum || candidate.capacity() > maximum {
            return invalid("source XML candidate exceeds the size limit");
        }
        let original_range = validate_opening_proof(&source, &proof)?;
        let original = read_opening_token(
            source
                .bytes
                .get(original_range.clone())
                .ok_or_else(|| invalid_error("invalid XML source opening proof range"))?,
            source.path(),
        )?;
        if !original.empty {
            return invalid("XML source expansion proof must cover an empty element");
        }
        let candidate_range = candidate_replacement_range(&source, &original_range, &candidate)?;
        let expansion = read_empty_expansion(
            candidate
                .get(candidate_range.clone())
                .ok_or_else(|| invalid_error("invalid XML source expansion range"))?,
            source.path(),
        )?;
        if expansion.name != original.name || expansion.opening != original.content {
            return invalid("XML source expansion changed the root QName or attributes");
        }
        validate_candidate_outside_opening(&source, original_range, candidate_range, &candidate)?;
        verify_source_document(&candidate, source.path())?;
        Ok(Self {
            edits: Vec::new(),
            source,
            source_candidate: Some(candidate),
        })
    }

    /// Stage one checked replacement.
    ///
    /// # Errors
    ///
    /// Returns an error when the proof came from another package or part, its
    /// expected bytes are stale, or its range overlaps an earlier edit.
    pub fn replace(&mut self, proof: XmlSourceRange, fragment: AuthoredXmlFragment) -> Result<()> {
        if self.source_candidate.is_some() {
            return invalid("source candidate publication cannot mix XML splice edits");
        }
        if !Arc::ptr_eq(&self.source.archive, &proof.archive) || self.source.path != proof.path {
            return invalid("XML splice range has different source provenance");
        }
        let actual =
            self.source.bytes.get(proof.range.clone()).ok_or_else(|| {
                Error::InvalidFormat("invalid XML splice source range".to_string())
            })?;
        if actual != proof.expected {
            return invalid("stale XML splice source range");
        }
        if self
            .edits
            .iter()
            .any(|edit| ranges_overlap_or_conflict(&edit.range, &proof.range))
        {
            return invalid("overlapping XML splice ranges");
        }
        self.edits.push(Edit {
            fragment,
            range: proof.range,
        });
        Ok(())
    }

    /// Publish the checked splice through the normal ODF package writer.
    ///
    /// # Errors
    ///
    /// Returns an error when the assembled part is oversized or malformed, or
    /// when the package writer cannot emit it.
    pub fn publish<W: Write>(self, writer: &mut PackageWriter<W>) -> Result<()> {
        writer.add_spliced_xml(self)
    }

    pub(crate) fn belongs_to(&self, source: &OwnedPackage) -> bool {
        Arc::ptr_eq(&self.source.archive, &source.shared_bytes())
    }

    pub(crate) fn assemble(mut self) -> Result<(String, Vec<u8>, String)> {
        if let Some(candidate) = self.source_candidate.take() {
            if !self.edits.is_empty() {
                return invalid("source candidate publication cannot mix XML splice edits");
            }
            return Ok((self.source.path, candidate, self.source.media_type));
        }
        self.edits.sort_by_key(|edit| edit.range.start);
        let removed = self.edits.iter().try_fold(0usize, |total, edit| {
            total
                .checked_add(edit.range.end - edit.range.start)
                .ok_or_else(|| Error::InvalidFormat("XML splice size overflow".to_string()))
        })?;
        let added = self.edits.iter().try_fold(0usize, |total, edit| {
            total
                .checked_add(edit.fragment.bytes.len())
                .ok_or_else(|| Error::InvalidFormat("XML splice size overflow".to_string()))
        })?;
        let capacity = self
            .source
            .bytes
            .len()
            .checked_sub(removed)
            .and_then(|length| length.checked_add(added))
            .ok_or_else(|| Error::InvalidFormat("XML splice size overflow".to_string()))?;
        if capacity > MAX_PART_BYTES {
            return invalid("spliced XML part exceeds the size limit");
        }
        let mut output = Vec::new();
        output.try_reserve_exact(capacity).map_err(|error| {
            Error::InvalidFormat(format!("spliced XML allocation failed: {error}"))
        })?;
        let mut cursor = 0usize;
        for edit in self.edits {
            output.extend_from_slice(&self.source.bytes[cursor..edit.range.start]);
            output.extend_from_slice(&edit.fragment.bytes);
            cursor = edit.range.end;
        }
        output.extend_from_slice(&self.source.bytes[cursor..]);
        verify_well_formed(&output, &self.source.path)?;
        Ok((self.source.path, output, self.source.media_type))
    }
}

/// Rebuild `source` with checked XML splice publications and a bounded output.
///
/// Untouched members are copied as exact-source payloads, formatting outside
/// checked splice ranges remains byte-identical, stale signatures are omitted,
/// and the manifest is regenerated by [`PackageWriter`].
///
/// # Errors
///
/// Returns an error for foreign or duplicate publications, unsupported source
/// encryption, publication defects, or output beyond `output_limit`.
pub fn rebuild_package_with_xml_splices(
    source: &OwnedPackage,
    publications: Vec<XmlSplicePublication>,
    output_limit: usize,
) -> Result<Vec<u8>> {
    let mut paths = HashSet::with_capacity(publications.len());
    for publication in &publications {
        if !publication.belongs_to(source) {
            return invalid("XML splice publication has different package provenance");
        }
        if !paths.insert(publication.source.path.clone()) {
            return invalid("duplicate XML splice publication path");
        }
    }

    let mut writer = PackageWriter::new_bounded(output_limit);
    writer.set_mimetype(&source.mimetype()?)?;
    for publication in publications {
        publication.publish(&mut writer)?;
    }
    writer.copy_source_files_from_except(source, &paths)?;
    writer.finish_to_bounded_bytes()
}

fn audit_wrapped_fragment(fragment: &[u8]) -> Result<()> {
    let length = FRAGMENT_ROOT_OPEN
        .len()
        .checked_add(fragment.len())
        .and_then(|value| value.checked_add(FRAGMENT_ROOT_CLOSE.len()))
        .ok_or_else(|| Error::InvalidFormat("authored XML fragment size overflow".to_string()))?;
    let mut document = Vec::new();
    document.try_reserve_exact(length).map_err(|error| {
        Error::InvalidFormat(format!("authored XML fragment allocation failed: {error}"))
    })?;
    document.extend_from_slice(FRAGMENT_ROOT_OPEN);
    document.extend_from_slice(fragment);
    document.extend_from_slice(FRAGMENT_ROOT_CLOSE);
    audit_document(&document)
}

fn audit_document(document: &[u8]) -> Result<()> {
    xml_minifier::audit::verify_authored(document, xml_minifier::audit::Limits::default())
        .map(|_report| ())
        .map_err(|source| Error::InvalidFormat(format!("authored XML fragment rejected: {source}")))
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct OpeningToken {
    name: Vec<u8>,
    empty: bool,
    content: Vec<u8>,
}

fn validate_opening_proof(source: &XmlSourcePart, proof: &XmlSourceRange) -> Result<Range<usize>> {
    if !Arc::ptr_eq(&source.archive, &proof.archive) || source.path != proof.path {
        return invalid("XML source opening proof has different provenance");
    }
    let range = proof.range.clone();
    let bytes = source.bytes.get(range.clone()).ok_or_else(|| {
        Error::InvalidFormat("invalid XML source opening proof range".to_string())
    })?;
    if bytes != proof.expected.as_slice() {
        return invalid("stale XML source opening proof");
    }
    let _ = read_opening_token(bytes, source.path())?;
    Ok(range)
}

fn validate_candidate_opening(
    source: &XmlSourcePart,
    original_range: Range<usize>,
    candidate: &[u8],
) -> Result<Range<usize>> {
    let candidate_range = candidate_replacement_range(source, &original_range, candidate)?;
    let source_token = read_opening_token(
        source.bytes.get(original_range.clone()).ok_or_else(|| {
            Error::InvalidFormat("invalid XML source opening proof range".to_string())
        })?,
        source.path(),
    )?;
    let candidate_token = read_opening_token(
        candidate
            .get(candidate_range.clone())
            .ok_or_else(|| invalid_error("XML source candidate opening range is invalid"))?,
        source.path(),
    )?;
    if candidate_token.name != source_token.name || candidate_token.empty != source_token.empty {
        return invalid("XML source candidate changed the root QName or empty-element shape");
    }
    Ok(candidate_range)
}

fn candidate_replacement_range(
    source: &XmlSourcePart,
    original_range: &Range<usize>,
    candidate: &[u8],
) -> Result<Range<usize>> {
    let source_len = source.bytes.len();
    let candidate_len = candidate.len();
    let candidate_end = if candidate_len >= source_len {
        original_range.end.checked_add(candidate_len - source_len)
    } else {
        original_range.end.checked_sub(source_len - candidate_len)
    }
    .ok_or_else(|| {
        Error::InvalidFormat("XML source candidate opening range overflow".to_string())
    })?;
    if candidate_end < original_range.start || candidate_end > candidate_len {
        return invalid("XML source candidate opening range is invalid");
    }
    Ok(original_range.start..candidate_end)
}

fn validate_candidate_outside_opening(
    source: &XmlSourcePart,
    original_range: Range<usize>,
    candidate_range: Range<usize>,
    candidate: &[u8],
) -> Result<()> {
    let source_prefix = source
        .bytes
        .get(..original_range.start)
        .ok_or_else(|| invalid_error("invalid XML source opening proof prefix"))?;
    let candidate_prefix = candidate
        .get(..candidate_range.start)
        .ok_or_else(|| invalid_error("invalid XML source candidate opening prefix"))?;
    if source_prefix != candidate_prefix {
        return invalid("XML source candidate changed bytes before the proven opening tag");
    }
    let source_suffix = source
        .bytes
        .get(original_range.end..)
        .ok_or_else(|| invalid_error("invalid XML source opening proof suffix"))?;
    let candidate_suffix = candidate
        .get(candidate_range.end..)
        .ok_or_else(|| invalid_error("invalid XML source candidate opening suffix"))?;
    if source_suffix != candidate_suffix {
        return invalid("XML source candidate changed bytes after the proven opening tag");
    }
    Ok(())
}

fn read_opening_token(bytes: &[u8], path: &str) -> Result<OpeningToken> {
    let text = std::str::from_utf8(bytes).map_err(|error| {
        invalid_error(format!(
            "XML source opening token for '{path}' is not UTF-8: {error}"
        ))
    })?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let event = reader.read_event_into(&mut buffer).map_err(|error| {
        invalid_error(format!("XML source opening token is malformed: {error}"))
    })?;
    let (name, empty, content) = match &event {
        Event::Start(element) => (
            copy_opening_qname(element.name().as_ref())?,
            false,
            copy_opening_content(element)?,
        ),
        Event::Empty(element) => (
            copy_opening_qname(element.name().as_ref())?,
            true,
            copy_opening_content(element)?,
        ),
        _ => return invalid("XML source proof must cover one start or empty element token"),
    };
    buffer.clear();
    match reader
        .read_event_into(&mut buffer)
        .map_err(|error| invalid_error(format!("XML source opening token is malformed: {error}")))?
    {
        Event::Eof => Ok(OpeningToken {
            name,
            empty,
            content,
        }),
        _ => invalid("XML source proof must cover exactly one opening token"),
    }
}

fn copy_opening_qname(name: &[u8]) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|source| Error::Allocation {
            resource: "XML source opening QName",
            source,
        })?;
    owned.extend_from_slice(name);
    Ok(owned)
}

fn copy_opening_content(element: &quick_xml::events::BytesStart<'_>) -> Result<Vec<u8>> {
    let raw: &[u8] = element.as_ref();
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(raw.len())
        .map_err(|source| Error::Allocation {
            resource: "XML source opening attributes",
            source,
        })?;
    owned.extend_from_slice(raw);
    Ok(owned)
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct EmptyExpansion {
    name: Vec<u8>,
    opening: Vec<u8>,
}

fn read_empty_expansion(bytes: &[u8], path: &str) -> Result<EmptyExpansion> {
    let text = std::str::from_utf8(bytes).map_err(|error| {
        invalid_error(format!(
            "XML source expansion for '{path}' is not UTF-8: {error}"
        ))
    })?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let event = reader
        .read_event_into(&mut buffer)
        .map_err(|error| invalid_error(format!("XML source expansion is malformed: {error}")))?;
    let (name, opening) = match &event {
        Event::Start(element) => (
            copy_opening_qname(element.name().as_ref())?,
            copy_opening_content(element)?,
        ),
        _ => return invalid("XML source expansion must begin with a start element"),
    };
    buffer.clear();
    let mut depth = 1usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            invalid_error(format!("XML source expansion is malformed: {error}"))
        })?;
        match &event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("XML source expansion depth overflow"))?;
            },
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
            Event::End(end) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("XML source expansion has an unmatched end"))?;
                if depth == 0 {
                    if end.name().as_ref() != name.as_slice() {
                        return invalid("XML source expansion closing QName does not match");
                    }
                    buffer.clear();
                    return match reader.read_event_into(&mut buffer).map_err(|error| {
                        invalid_error(format!("XML source expansion is malformed: {error}"))
                    })? {
                        Event::Eof => Ok(EmptyExpansion { name, opening }),
                        _ => invalid("XML source expansion has content after its root"),
                    };
                }
            },
            Event::Decl(_) | Event::DocType(_) => {
                return invalid("XML source expansion contains a declaration or doctype");
            },
            Event::Eof => return invalid("XML source expansion has no matching end element"),
        }
        buffer.clear();
    }
}

fn start_tag_name(bytes: &[u8]) -> Result<&[u8]> {
    if bytes.len() < 3
        || bytes.first() != Some(&b'<')
        || bytes.last() != Some(&b'>')
        || bytes.starts_with(b"</")
        || bytes.starts_with(b"<!")
        || bytes.starts_with(b"<?")
        || bytes.ends_with(b"/>")
    {
        return invalid("authored XML start-tag fragment is unclassified");
    }
    let end = bytes[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'>')
        .map_or(bytes.len() - 1, |offset| offset + 1);
    if end == 1 {
        return invalid("authored XML start-tag fragment has no name");
    }
    Ok(&bytes[1..end])
}

fn end_tag_name(bytes: &[u8]) -> Result<&[u8]> {
    if bytes.len() < 4
        || !bytes.starts_with(b"</")
        || bytes.last() != Some(&b'>')
        || bytes[2..bytes.len() - 1]
            .iter()
            .any(|byte| byte.is_ascii_whitespace() || matches!(byte, b'<' | b'>'))
    {
        return invalid("authored XML end-tag fragment is unclassified");
    }
    Ok(&bytes[2..bytes.len() - 1])
}

fn ranges_overlap_or_conflict(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
        || (left.start == left.end && right.start == right.end && left.start == right.start)
}

fn verify_well_formed(bytes: &[u8], path: &str) -> Result<()> {
    let xml = std::str::from_utf8(bytes).map_err(|error| {
        Error::InvalidFormat(format!("XML part '{path}' is not UTF-8: {error}"))
    })?;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                if depth == 0 {
                    roots += 1;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("XML part '{path}' depth overflow"))
                })?;
            },
            Ok(Event::Empty(_)) => {
                if depth == 0 {
                    roots += 1;
                }
            },
            Ok(Event::End(_)) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("XML part '{path}' has an unexpected end tag"))
                })?;
            },
            Ok(Event::Text(text)) => {
                let text_bytes: &[u8] = text.as_ref();
                if depth == 0 && !text_bytes.iter().all(u8::is_ascii_whitespace) {
                    return invalid(format!("XML part '{path}' has text outside its root"));
                }
            },
            Ok(Event::CData(_) | Event::GeneralRef(_)) if depth == 0 => {
                return invalid(format!("XML part '{path}' has content outside its root"));
            },
            Ok(Event::Eof) => break,
            Ok(
                Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_),
            ) => {},
            Err(error) => {
                return invalid(format!("XML part '{path}' is malformed: {error}"));
            },
        }
    }
    if depth != 0 || roots != 1 {
        return invalid(format!(
            "XML part '{path}' must contain one closed root element"
        ));
    }
    Ok(())
}

/// Validate a complete source-backed XML candidate without imposing the
/// compact lexical rules used by [`AuthoredXmlFragment`].
///
/// quick-xml's ordinary reader provides structural tokenization but leaves
/// namespace resolution to its resolver. The common binding tracker is used
/// here so inherited aliases, noncompact spacing, comments, processing
/// instructions, and CDATA remain legal while unbound prefixes, reserved
/// namespace misuse, malformed attributes, and invalid references fail before
/// package publication.
fn verify_source_document(bytes: &[u8], path: &str) -> Result<()> {
    let xml = std::str::from_utf8(bytes).map_err(|error| {
        Error::InvalidFormat(format!("XML part '{path}' is not UTF-8: {error}"))
    })?;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.config_mut().trim_text(false);
    let mut tracker = BindingTracker::new().map_err(|error| {
        error.into_litchi_error_with_context(|| format!("invalid XML part '{path}'"))
    })?;
    let mut buffer = Vec::new();
    let mut pending_pop = false;
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut events = 0usize;
    let mut saw_event = false;
    let mut declaration_seen = false;

    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event_start = reader.buffer_position();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("XML part '{path}' is malformed: {error}")))?;
        validate_event_bytes(&event, event_start)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("XML source event counter overflow"))?;
        if events > MAX_SOURCE_EVENTS {
            return invalid(format!("XML part '{path}' exceeds its event limit"));
        }
        match &event {
            Event::Start(element) => {
                if depth >= MAX_SOURCE_DEPTH {
                    return invalid(format!("XML part '{path}' exceeds its depth limit"));
                }
                validate_source_element(&mut tracker, &reader, element, path)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid_error("XML source root counter overflow"))?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("XML source depth overflow"))?;
            },
            Event::Empty(element) => {
                if depth >= MAX_SOURCE_DEPTH {
                    return invalid(format!("XML part '{path}' exceeds its depth limit"));
                }
                validate_source_element(&mut tracker, &reader, element, path)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid_error("XML source root counter overflow"))?;
                }
                pending_pop = true;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid_error(format!(
                        "XML part '{path}' has an unexpected end tag"
                    )));
                }
                validate_end_element(&tracker, element, 0)?;
                depth -= 1;
                pending_pop = true;
            },
            Event::Text(text) => {
                let raw: &[u8] = text.as_ref();
                let decoded = text.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_error(format!("XML part '{path}' has invalid text: {error}"))
                })?;
                validate_source_event_characters(raw, decoded.as_ref(), path, "text", true)?;
                if depth == 0 && !raw.iter().all(u8::is_ascii_whitespace) {
                    return invalid(format!("XML part '{path}' has text outside its root"));
                }
            },
            Event::CData(text) => {
                let raw: &[u8] = text.as_ref();
                let decoded = text.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_error(format!("XML part '{path}' has invalid CDATA: {error}"))
                })?;
                validate_source_event_characters(raw, decoded.as_ref(), path, "CDATA", true)?;
                if depth == 0 {
                    return invalid(format!("XML part '{path}' has CDATA outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                if !crate::validation::valid_xml_reference(reference) {
                    return invalid(format!(
                        "XML part '{path}' has an invalid character or entity reference"
                    ));
                }
                if depth == 0 {
                    return invalid(format!(
                        "XML part '{path}' has a reference outside its root"
                    ));
                }
            },
            Event::Decl(declaration) => {
                if declaration_seen || saw_event || depth != 0 {
                    return invalid(format!(
                        "XML part '{path}' has an XML declaration outside its prologue"
                    ));
                }
                let mut declaration_attributes = 0usize;
                validate_decl_attributes(
                    &tracker,
                    declaration,
                    &mut declaration_attributes,
                    MAX_SOURCE_ATTRIBUTES,
                    0,
                )?;
                let version = declaration.version().map_err(|error| {
                    invalid_error(format!(
                        "XML part '{path}' has an invalid XML version: {error}"
                    ))
                })?;
                if version.as_ref() != b"1.0" {
                    return invalid(format!("XML part '{path}' uses an unsupported XML version"));
                }
                if let Some(encoding) = declaration.encoding() {
                    let encoding = encoding.map_err(|error| {
                        invalid_error(format!(
                            "XML part '{path}' has an invalid XML encoding: {error}"
                        ))
                    })?;
                    if !encoding.eq_ignore_ascii_case(b"UTF-8") {
                        return invalid(format!(
                            "XML part '{path}' uses an unsupported XML encoding"
                        ));
                    }
                }
                if let Some(standalone) = declaration.standalone() {
                    let standalone = standalone.map_err(|error| {
                        invalid_error(format!(
                            "XML part '{path}' has an invalid standalone declaration: {error}"
                        ))
                    })?;
                    if standalone.as_ref() != b"yes" && standalone.as_ref() != b"no" {
                        return invalid(format!(
                            "XML part '{path}' has an invalid standalone declaration"
                        ));
                    }
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return invalid(format!("XML part '{path}' contains a doctype"));
            },
            Event::Comment(comment) => {
                let raw: &[u8] = comment.as_ref();
                let decoded = std::str::from_utf8(raw).map_err(|error| {
                    invalid_error(format!("XML part '{path}' has invalid comment: {error}"))
                })?;
                validate_source_event_characters(raw, decoded, path, "comment", false)?;
            },
            Event::PI(instruction) => {
                validate_name(instruction.target(), 0, "processing-instruction target")?;
                if instruction.target().eq_ignore_ascii_case(b"xml") {
                    return invalid(format!(
                        "XML part '{path}' has a reserved processing-instruction target"
                    ));
                }
                let raw = instruction.content();
                let decoded = std::str::from_utf8(raw).map_err(|error| {
                    invalid_error(format!(
                        "XML part '{path}' has invalid processing-instruction data: {error}"
                    ))
                })?;
                validate_source_event_characters(
                    raw,
                    decoded,
                    path,
                    "processing-instruction data",
                    false,
                )?;
            },
            Event::Eof => break,
        }
        if !matches!(&event, Event::Eof) {
            saw_event = true;
        }
        buffer.clear();
    }
    if depth != 0 || roots != 1 {
        return invalid(format!(
            "XML part '{path}' must contain one closed root element"
        ));
    }
    Ok(())
}

fn validate_source_element(
    tracker: &mut BindingTracker,
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<()> {
    tracker.push(element).map_err(|error| {
        error.into_litchi_error_with_context(|| format!("invalid XML part '{path}'"))
    })?;
    let mut attributes = 0usize;
    validate_start_element(
        tracker,
        element,
        &mut attributes,
        MAX_SOURCE_ATTRIBUTES,
        MAX_PART_BYTES,
        0,
    )?;
    validate_source_attribute_characters(reader, element, path)?;
    validate_source_namespace_bindings(reader, element, path)?;
    Ok(())
}

fn validate_source_attribute_characters(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<()> {
    for raw in element.attributes().with_checks(true) {
        let attribute = raw.map_err(|error| {
            invalid_error(format!(
                "XML part '{path}' has an invalid attribute: {error}"
            ))
        })?;
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                invalid_error(format!(
                    "XML part '{path}' has an invalid attribute value: {error}"
                ))
            })?;
        validate_source_xml_characters(decoded.as_ref(), path, "attribute value")?;
    }
    Ok(())
}

fn validate_source_namespace_bindings(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<()> {
    for raw in element.attributes().with_checks(true) {
        let attribute = raw.map_err(|error| {
            invalid_error(format!(
                "XML part '{path}' has an invalid attribute: {error}"
            ))
        })?;
        let Some(prefix) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                invalid_error(format!(
                    "XML part '{path}' has an invalid namespace binding: {error}"
                ))
            })?;
        let value = value.as_bytes();
        match prefix {
            PrefixDeclaration::Default if value == XML_NAMESPACE || value == XMLNS_NAMESPACE => {
                return invalid(format!(
                    "XML part '{path}' binds a reserved namespace as default"
                ));
            },
            PrefixDeclaration::Named(b"xml") if value != XML_NAMESPACE => {
                return invalid(format!(
                    "XML part '{path}' binds the xml prefix to the wrong namespace"
                ));
            },
            PrefixDeclaration::Named(b"xmlns") => {
                return invalid(format!(
                    "XML part '{path}' redeclares the reserved xmlns prefix"
                ));
            },
            PrefixDeclaration::Named(_) if value == XML_NAMESPACE || value == XMLNS_NAMESPACE => {
                return invalid(format!(
                    "XML part '{path}' binds a non-reserved prefix to a reserved namespace"
                ));
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_source_event_characters(
    raw: &[u8],
    decoded: &str,
    path: &str,
    kind: &str,
    reject_cdata_delimiter: bool,
) -> Result<()> {
    if reject_cdata_delimiter && raw.windows(3).any(|window| window == b"]]>") {
        return invalid(format!(
            "XML part '{path}' contains an invalid {kind} delimiter"
        ));
    }
    validate_source_xml_characters(decoded, path, kind)
}

fn validate_source_xml_characters(value: &str, path: &str, kind: &str) -> Result<()> {
    if value.chars().all(is_xml_character) {
        Ok(())
    } else {
        invalid(format!(
            "XML part '{path}' has an invalid character in {kind}"
        ))
    }
}

fn is_xml_character(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

fn guess_media_type(path: &str) -> &'static str {
    if path
        .rsplit_once('.')
        .is_some_and(|(_stem, extension)| extension.eq_ignore_ascii_case("rdf"))
    {
        "application/rdf+xml"
    } else if path
        .rsplit_once('.')
        .is_some_and(|(_stem, extension)| extension.eq_ignore_ascii_case("xml"))
    {
        "text/xml"
    } else {
        "application/octet-stream"
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
