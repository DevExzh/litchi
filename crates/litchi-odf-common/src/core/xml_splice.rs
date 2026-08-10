//! Provenance-bearing publication of bounded edits to source-loaded XML parts.

use litchi_core::{Error, Result};
use quick_xml::{Reader, events::Event};
use std::{collections::HashSet, io::Write, ops::Range, sync::Arc};

use super::{OwnedPackage, PackageWriter};

const MAX_PART_BYTES: usize = 256 * 1024 * 1024;
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
    /// Returns an error when the part is absent, is not XML-classified, is
    /// oversized, or is not a well-formed XML document.
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
        }
    }

    /// Stage one checked replacement.
    ///
    /// # Errors
    ///
    /// Returns an error when the proof came from another package or part, its
    /// expected bytes are stale, or its range overlaps an earlier edit.
    pub fn replace(&mut self, proof: XmlSourceRange, fragment: AuthoredXmlFragment) -> Result<()> {
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
