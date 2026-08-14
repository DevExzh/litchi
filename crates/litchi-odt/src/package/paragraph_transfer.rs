//! Checked cross-document transfer for dependency-free plain paragraphs.
//!
//! This module deliberately accepts only a complete direct `text:p` whose
//! children are character data.  The paragraph's source bytes are retained as
//! an authored fragment; no style, identifier, relationship, or package
//! resource is inferred or copied.

#![deny(
    clippy::expect_used,
    clippy::panic,
    clippy::todo,
    clippy::unimplemented,
    clippy::unwrap_used
)]

use crate::{
    constants::ODF_CONTENT,
    protection::Policy,
    transaction::{Edit, EnvelopeKind, Position, Snapshot},
};
use litchi_core::{BlobId, Error, Result};
use litchi_odf_common::core::{AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication};
use litchi_odf_common::package::{
    MAX_CONTENT_REPLACEMENT_BYTES, raw_identical_members, replace_content_xml_spliced,
};
use quick_xml::{
    events::{BytesStart, Event},
    name::{PrefixDeclaration, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

/// Maximum number of paragraphs accepted by one transfer plan.
pub const MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS: usize = 256;
/// Maximum bytes retained by one transferred paragraph fragment.
pub const MAX_PLAIN_PARAGRAPH_TRANSFER_FRAGMENT_BYTES: usize = 1024 * 1024;
/// Maximum combined lexical bytes retained by one transfer plan.
pub const MAX_PLAIN_PARAGRAPH_TRANSFER_BYTES: usize = 8 * 1024 * 1024;
const MAX_EVENTS: usize = 1_048_576;
const MAX_DEPTH: usize = 64;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MCE_NS: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// A checked, immutable donor-to-destination paragraph transfer.
///
/// The plan captures exact donor paragraph bytes and the exact destination
/// package digest observed during planning.  Applying it to any other target,
/// or after target bytes have changed, is refused before publication.
#[derive(Clone)]
pub struct ParagraphTransferPlan {
    pub(crate) donor: Snapshot,
    pub(crate) donor_fingerprint: String,
    pub(crate) destination_fingerprint: String,
    pub(crate) source_positions: Vec<usize>,
    pub(crate) target_position: usize,
    pub(crate) fragments: Vec<Vec<u8>>,
    pub(crate) fragment_digest: String,
}

impl ParagraphTransferPlan {
    /// Exact SHA-256 digest of the immutable donor package used for planning.
    #[must_use]
    pub fn donor_fingerprint(&self) -> String {
        self.donor_fingerprint.clone()
    }

    /// Alias for callers that use source terminology for the donor.
    #[must_use]
    pub fn source_fingerprint(&self) -> String {
        self.donor_fingerprint()
    }

    /// Exact SHA-256 digest of the destination package used for planning.
    #[must_use]
    pub fn destination_fingerprint(&self) -> String {
        self.destination_fingerprint.clone()
    }

    /// Exact SHA-256 digest of the concatenated lexical paragraph fragments.
    #[must_use]
    pub fn fragment_digest(&self) -> String {
        self.fragment_digest.clone()
    }

    /// Number of paragraphs inserted by this plan.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.fragments.len()
    }

    /// Checked donor positions captured by this plan.
    #[must_use]
    pub fn source_positions(&self) -> &[usize] {
        &self.source_positions
    }

    /// Checked destination insertion position captured by this plan.
    #[must_use]
    pub const fn target_position(&self) -> usize {
        self.target_position
    }

    /// Returns the immutable donor snapshot retained by this plan.
    #[must_use]
    pub const fn donor(&self) -> &Snapshot {
        &self.donor
    }

    /// Applies this plan to an edit whose source is the exact planned target.
    pub fn apply(&self, edit: &mut Edit) -> Result<()> {
        edit.stage_plain_paragraph_transfer(self).map(|_| ())
    }

    pub(crate) fn validate_integrity(&self, destination: &Snapshot) -> Result<()> {
        if BlobId::of(self.donor.as_bytes()).as_hex() != self.donor_fingerprint {
            return Err(Error::InvalidFormat(
                "ODT paragraph transfer donor changed after planning".to_string(),
            ));
        }
        if BlobId::of(destination.as_bytes()).as_hex() != self.destination_fingerprint {
            return Err(Error::InvalidFormat(
                "ODT paragraph transfer destination is stale or foreign".to_string(),
            ));
        }
        validate_positions_and_fragments(
            &self.source_positions,
            self.target_position,
            &self.fragments,
            &self.fragment_digest,
            None,
        )?;
        Ok(())
    }

    pub(crate) fn operation(&self) -> PlainParagraphTransferOperation {
        PlainParagraphTransferOperation {
            destination_fingerprint: self.destination_fingerprint.clone(),
            source_positions: self.source_positions.clone(),
            target_position: self.target_position,
            fragments: self.fragments.clone(),
            fragment_digest: self.fragment_digest.clone(),
        }
    }
}

/// Internal operation payload retained by the ODT transaction and durable
/// decoder. It contains no mutable document handle and never executes donor
/// content. Donor package provenance is intentionally retained only by the
/// in-memory plan: a durable patch does not carry the complete donor artifact,
/// so a donor digest alone could not be authenticated during replay.
#[derive(Clone)]
pub(crate) struct PlainParagraphTransferOperation {
    pub(crate) destination_fingerprint: String,
    pub(crate) source_positions: Vec<usize>,
    pub(crate) target_position: usize,
    pub(crate) fragments: Vec<Vec<u8>>,
    pub(crate) fragment_digest: String,
}

/// Plan one direct paragraph transfer.
pub(crate) fn plan_one(
    destination: &Snapshot,
    donor: &Snapshot,
    source: Position,
    target: Position,
) -> Result<ParagraphTransferPlan> {
    plan_many(destination, donor, &[source], target)
}

/// Plan a bounded ordered batch of direct paragraph transfers.
pub(crate) fn plan_many(
    destination: &Snapshot,
    donor: &Snapshot,
    sources: &[Position],
    target: Position,
) -> Result<ParagraphTransferPlan> {
    validate_envelope(donor, "donor")?;
    validate_envelope(destination, "destination")?;

    let donor_document = donor.document()?;
    let destination_document = destination.document()?;
    let donor_content = donor_document.transaction_content_xml();
    let destination_content = destination_document.transaction_content_xml();
    if donor_content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("ODT paragraph transfer donor content.xml exceeds the limit");
    }
    if destination_content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("ODT paragraph transfer destination content.xml exceeds the limit");
    }

    let donor_body = scan_plain_content(donor_content)?;
    let destination_body = scan_plain_content(destination_content)?;
    if sources.len() > MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS {
        return invalid("ODT paragraph transfer paragraph limit exceeded");
    }
    if target.get() > MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS {
        return invalid("ODT paragraph transfer target limit exceeded");
    }
    if target.get() > destination_body.paragraphs.len() {
        return invalid("ODT paragraph transfer destination position is out of bounds");
    }

    let mut source_positions = Vec::new();
    source_positions
        .try_reserve_exact(sources.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer source positions",
            source,
        })?;
    let mut fragments = Vec::new();
    fragments
        .try_reserve_exact(sources.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer fragments",
            source,
        })?;
    let mut total_bytes = 0usize;
    for source in sources {
        let index = source.get();
        let range = donor_body.paragraphs.get(index).ok_or_else(|| {
            invalid_error("ODT paragraph transfer donor position is out of bounds")
        })?;
        if source_positions.contains(&index) {
            return Err(Error::InvalidFormat(
                "ODT paragraph transfer donor positions must be unique".to_string(),
            ));
        }
        let bytes = donor_content
            .as_bytes()
            .get(range.clone())
            .ok_or_else(|| invalid_error("ODT paragraph transfer donor range is invalid"))?;
        if bytes.len() > MAX_PLAIN_PARAGRAPH_TRANSFER_FRAGMENT_BYTES {
            return invalid("ODT paragraph transfer fragment exceeds the limit");
        }
        total_bytes = total_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| invalid_error("ODT paragraph transfer byte count overflow"))?;
        if total_bytes > MAX_PLAIN_PARAGRAPH_TRANSFER_BYTES {
            return invalid("ODT paragraph transfer byte limit exceeded");
        }
        let mut fragment = Vec::new();
        fragment
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT paragraph transfer lexical fragment",
                source,
            })?;
        fragment.extend_from_slice(bytes);
        source_positions.push(index);
        fragments.push(fragment);
    }

    let fragment_digest = digest_fragments(&fragments)?;
    let fragment_digest_hex = fragment_digest.as_hex();
    validate_positions_and_fragments(
        &source_positions,
        target.get(),
        &fragments,
        &fragment_digest_hex,
        Some(&destination_body.text_prefixes),
    )?;
    Ok(ParagraphTransferPlan {
        donor: donor.clone(),
        donor_fingerprint: BlobId::of(donor.as_bytes()).as_hex(),
        destination_fingerprint: BlobId::of(destination.as_bytes()).as_hex(),
        source_positions,
        target_position: target.get(),
        fragments,
        fragment_digest: fragment_digest_hex,
    })
}

/// Apply one already-authenticated transaction operation to the target.
pub(crate) fn apply_operation(
    destination: &Snapshot,
    operation: &PlainParagraphTransferOperation,
) -> Result<Snapshot> {
    validate_envelope(destination, "destination")?;
    if BlobId::of(destination.as_bytes()).as_hex() != operation.destination_fingerprint {
        return Err(Error::InvalidFormat(
            "ODT paragraph transfer destination is stale or foreign".to_string(),
        ));
    }
    if operation.fragments.is_empty() {
        return Ok(destination.clone());
    }

    let document = destination.document()?;
    let content = document.transaction_content_xml();
    if content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("ODT paragraph transfer destination content.xml exceeds the limit");
    }
    let body = scan_plain_content(content)?;
    validate_positions_and_fragments(
        &operation.source_positions,
        operation.target_position,
        &operation.fragments,
        &operation.fragment_digest,
        Some(&body.text_prefixes),
    )?;
    if operation.target_position > body.paragraphs.len() {
        return invalid("ODT paragraph transfer destination position is stale");
    }
    let insertion = insertion_offset(&body, operation.target_position)?;
    let replacement = assembled_content(content, insertion, &operation.fragments)?;
    let source_part = XmlSourcePart::load(document.transaction_package(), ODF_CONTENT)?;
    let proof_source = source_part.clone();
    let proof = proof_source.checked_range(insertion..insertion, &[])?;
    let fragment_bytes = joined_fragments(&operation.fragments)?;
    let authored = AuthoredXmlFragment::markup(fragment_bytes)?;
    let mut publication = XmlSplicePublication::new(source_part);
    publication.replace(proof, authored)?;
    let target_bytes =
        replace_content_xml_spliced(document.transaction_package(), &replacement, publication)?;
    let target = Snapshot::from_bytes(target_bytes)?;
    let target_document = target.document()?;
    require_raw_untouched_members(
        document.transaction_package(),
        target_document.transaction_package(),
    )?;
    Ok(target)
}

/// Validate a decoded durable operation before it is admitted to replay.
pub(crate) fn validate_operation_for_durable(
    operation: &PlainParagraphTransferOperation,
) -> Result<()> {
    if operation.destination_fingerprint.len() != 64
        || !operation
            .destination_fingerprint
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return invalid("ODT paragraph transfer durable destination digest is invalid");
    }
    if operation.fragment_digest.len() != 64
        || !operation
            .fragment_digest
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return invalid("ODT paragraph transfer durable fragment digest is invalid");
    }
    if operation.source_positions.is_empty() || operation.fragments.is_empty() {
        return invalid("ODT paragraph transfer durable operation is empty");
    }
    validate_positions_and_fragments(
        &operation.source_positions,
        operation.target_position,
        &operation.fragments,
        &operation.fragment_digest,
        None,
    )
}

fn validate_envelope(snapshot: &Snapshot, role: &str) -> Result<()> {
    match snapshot.envelope_kind()? {
        EnvelopeKind::Plain => {},
        EnvelopeKind::Signed => {
            return Err(Error::Unsupported(format!(
                "ODT paragraph transfer refuses signed {role} packages"
            )));
        },
        EnvelopeKind::Encrypted => {
            return Err(Error::Unsupported(format!(
                "ODT paragraph transfer refuses encrypted {role} packages"
            )));
        },
    }
    let document = snapshot.document()?;
    if document.protection()? != Policy::default() {
        return Err(Error::Unsupported(format!(
            "ODT paragraph transfer refuses protected {role} documents"
        )));
    }
    if document
        .document_scripts()?
        .is_some_and(|scripts| !scripts.scripts.is_empty() || !scripts.event_listeners.is_empty())
        || !document.script_resources()?.is_empty()
    {
        return Err(Error::Unsupported(format!(
            "ODT paragraph transfer refuses scripted {role} documents"
        )));
    }
    let declarations = document.variable_declarations()?;
    if !declarations.dde_connections.is_empty() || !declarations.dde_connection_uses.is_empty() {
        return Err(Error::Unsupported(format!(
            "ODT paragraph transfer refuses DDE-bearing {role} documents"
        )));
    }
    Ok(())
}

fn validate_positions_and_fragments(
    source_positions: &[usize],
    target_position: usize,
    fragments: &[Vec<u8>],
    expected_digest: &str,
    expected_text_prefixes: Option<&[Vec<u8>]>,
) -> Result<()> {
    if source_positions.len() != fragments.len()
        || source_positions.len() > MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS
    {
        return invalid("ODT paragraph transfer operation exceeds its bounds");
    }
    if target_position > MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS {
        return invalid("ODT paragraph transfer target exceeds its bounds");
    }
    let mut total_bytes = 0usize;
    for (position, fragment) in source_positions.iter().zip(fragments) {
        if *position >= MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS
            || fragment.len() > MAX_PLAIN_PARAGRAPH_TRANSFER_FRAGMENT_BYTES
        {
            return invalid("ODT paragraph transfer fragment or position exceeds its bounds");
        }
        total_bytes = total_bytes
            .checked_add(fragment.len())
            .ok_or_else(|| invalid_error("ODT paragraph transfer byte count overflow"))?;
        if total_bytes > MAX_PLAIN_PARAGRAPH_TRANSFER_BYTES {
            return invalid("ODT paragraph transfer byte limit exceeded");
        }
        let prefix = validate_fragment_shape(fragment)?;
        if let Some(expected_text_prefixes) = expected_text_prefixes {
            if !expected_text_prefixes
                .iter()
                .any(|expected| expected.as_slice() == prefix)
            {
                return unsupported(
                    "paragraph fragment text namespace is not bound by destination",
                );
            }
        }
    }
    if digest_fragments(fragments)?.as_hex() != expected_digest {
        return Err(Error::InvalidFormat(
            "ODT paragraph transfer fragment digest does not match".to_string(),
        ));
    }
    for (index, position) in source_positions.iter().enumerate() {
        if source_positions[index + 1..]
            .iter()
            .any(|other| other == position)
        {
            return Err(Error::InvalidFormat(
                "ODT paragraph transfer donor positions must be unique".to_string(),
            ));
        }
    }
    Ok(())
}

fn validate_fragment_shape(fragment: &[u8]) -> Result<&[u8]> {
    let tag = fragment
        .strip_prefix(b"<")
        .ok_or_else(|| invalid_error("ODT paragraph transfer fragment is not markup"))?;
    let tag_end = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))
        .ok_or_else(|| invalid_error("ODT paragraph transfer fragment start tag is invalid"))?;
    let name = &tag[..tag_end];
    let (prefix, local) = match name.iter().position(|byte| *byte == b':') {
        Some(separator) => (&name[..separator], &name[separator + 1..]),
        None => (&[][..], name),
    };
    if local != b"p" || prefix == b"office" || prefix.iter().any(|byte| !is_xml_name_byte(*byte)) {
        return unsupported("non-text paragraph fragment namespace or name");
    }
    let mut wrapper = Vec::new();
    let namespace_bytes = if prefix.is_empty() {
        b" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"".as_slice()
    } else {
        &[][..]
    };
    let prefix_capacity = if prefix.is_empty() {
        0
    } else {
        prefix.len().saturating_add(64)
    };
    wrapper
        .try_reserve_exact(
            fragment
                .len()
                .saturating_add(192)
                .saturating_add(prefix_capacity),
        )
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer fragment validation",
            source,
        })?;
    wrapper.extend_from_slice(b"<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"");
    if prefix.is_empty() {
        wrapper.extend_from_slice(namespace_bytes);
    } else {
        wrapper.extend_from_slice(b" xmlns:");
        wrapper.extend_from_slice(prefix);
        wrapper.extend_from_slice(b"=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"");
    }
    wrapper.extend_from_slice(b"><office:body><office:text>");
    wrapper.extend_from_slice(fragment);
    wrapper.extend_from_slice(b"</office:text></office:body></office:document-content>");
    let xml = std::str::from_utf8(&wrapper)
        .map_err(|error| Error::InvalidFormat(format!("invalid ODT paragraph UTF-8: {error}")))?;
    let body = scan_plain_content(xml)?;
    if body.paragraphs.len() != 1 {
        return unsupported("paragraph transfer fragment is not one direct plain paragraph");
    }
    Ok(prefix)
}

fn is_xml_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')
}

fn is_allowed_content_preamble(local: &[u8]) -> bool {
    matches!(
        local,
        b"scripts" | b"font-face-decls" | b"automatic-styles" | b"body"
    )
}

fn joined_fragments(fragments: &[Vec<u8>]) -> Result<Vec<u8>> {
    let length = fragments.iter().try_fold(0usize, |total, fragment| {
        total
            .checked_add(fragment.len())
            .ok_or_else(|| invalid_error("ODT paragraph transfer fragment size overflow"))
    })?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer fragment assembly",
            source,
        })?;
    for fragment in fragments {
        output.extend_from_slice(fragment);
    }
    Ok(output)
}

fn digest_fragments(fragments: &[Vec<u8>]) -> Result<BlobId> {
    Ok(BlobId::of(&joined_fragments(fragments)?))
}

fn assembled_content(content: &str, insertion: usize, fragments: &[Vec<u8>]) -> Result<String> {
    let joined = joined_fragments(fragments)?;
    let capacity = content
        .len()
        .checked_add(joined.len())
        .ok_or_else(|| invalid_error("ODT paragraph transfer content size overflow"))?;
    if capacity > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("ODT paragraph transfer output content.xml exceeds the limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer output content.xml",
            source,
        })?;
    output.push_str(
        content.get(..insertion).ok_or_else(|| {
            invalid_error("ODT paragraph transfer insertion is not UTF-8 aligned")
        })?,
    );
    output.push_str(
        std::str::from_utf8(&joined).map_err(|error| {
            Error::InvalidFormat(format!("invalid ODT paragraph UTF-8: {error}"))
        })?,
    );
    output.push_str(
        content
            .get(insertion..)
            .ok_or_else(|| invalid_error("ODT paragraph transfer suffix is not UTF-8 aligned"))?,
    );
    Ok(output)
}

fn insertion_offset(body: &PlainContent, position: usize) -> Result<usize> {
    if let Some(paragraph) = body.paragraphs.get(position) {
        return Ok(paragraph.start);
    }
    if position == body.paragraphs.len() {
        return Ok(body.text_end);
    }
    invalid("ODT paragraph transfer destination insertion is out of bounds")
}

struct PlainContent {
    paragraphs: Vec<Range<usize>>,
    text_prefixes: Vec<Vec<u8>>,
    text_end: usize,
}

fn scan_plain_content(xml: &str) -> Result<PlainContent> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    // The event buffer is reused for every event. Reserving the complete
    // bounded input up front keeps parser scratch growth fallible and bounded;
    // events are processed borrowed, so no per-event `into_owned` allocation is
    // needed.
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer XML event buffer",
            source,
        })?;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut body_depth = None;
    let mut text_depth = None;
    let mut paragraph = None::<(usize, usize)>;
    let mut paragraphs = Vec::new();
    let mut text_prefixes = Vec::new();
    let mut text_end = None;
    let mut preamble_depth = None;
    let mut saw_root = false;
    let mut saw_body = false;
    let mut saw_text = false;

    loop {
        let event_start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODT content.xml: {error}")))?;
        let mce = is_bound(&namespace, MCE_NS);
        let office = is_bound(&namespace, OFFICE_NS);
        let text = is_bound(&namespace, TEXT_NS);
        if matches!(&event, Event::Start(_) | Event::Empty(_))
            && match &event {
                Event::Start(element) | Event::Empty(element) => {
                    has_mce_attribute(&reader, element)?
                },
                _ => false,
            }
        {
            return unsupported("markup-compatibility attributes");
        }
        let event_end = position(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODT paragraph transfer event count overflow"))?;
        if events > MAX_EVENTS {
            return invalid("ODT paragraph transfer event limit exceeded");
        }
        if mce {
            return unsupported("markup-compatibility content");
        }

        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if depth == 0 {
                    if saw_root || !office || local.as_ref() != b"document-content" {
                        return unsupported("noncanonical content.xml root ownership");
                    }
                    saw_root = true;
                } else if preamble_depth.is_some() {
                    return unsupported("unknown content.xml preamble descendants");
                } else if office && local.as_ref() == b"scripts" {
                    return unsupported("documents containing scripts or events");
                } else if depth == 1 && body_depth.is_none() {
                    if !office || !is_allowed_content_preamble(local.as_ref()) {
                        return unsupported("foreign or unknown content.xml preamble markup");
                    }
                    if matches!(local.as_ref(), b"font-face-decls" | b"automatic-styles") {
                        if has_non_namespace_attributes(&reader, &element)? {
                            return unsupported("preamble placeholder attributes");
                        }
                        // Only empty, inert preamble placeholders are accepted.
                        // Any child element is rejected on the next event rather
                        // than being silently carried through this narrow scan.
                        preamble_depth = Some(depth);
                    }
                }
                if paragraph.is_some() {
                    return unsupported("paragraphs containing nested markup");
                }
                if office && local.as_ref() == b"body" {
                    if depth != 1 || saw_body || body_depth.is_some() {
                        return unsupported("multiple or nested office:body elements");
                    }
                    if has_non_namespace_attributes(&reader, &element)? {
                        return unsupported("office:body attributes");
                    }
                    saw_body = true;
                    body_depth = Some(depth);
                } else if office && local.as_ref() == b"text" {
                    if body_depth.and_then(|value| value.checked_add(1)) != Some(depth)
                        || saw_text
                        || text_depth.is_some()
                    {
                        return unsupported("noncanonical office:text ownership");
                    }
                    if has_non_namespace_attributes(&reader, &element)? {
                        return unsupported("office:text attributes");
                    }
                    text_prefixes = text_namespace_prefixes(&reader)?;
                    saw_text = true;
                    text_depth = Some(depth);
                } else if body_depth.is_some() && text_depth.is_none() {
                    return unsupported("non-text office:body children");
                } else if let Some(owner_depth) = text_depth {
                    if depth != owner_depth.saturating_add(1) || !text || local.as_ref() != b"p" {
                        return unsupported("non-plain office:text children");
                    }
                    if has_namespace_declaration(&element)
                        || has_non_namespace_attributes(&reader, &element)?
                    {
                        return unsupported("plain paragraph attributes, styles, or identifiers");
                    }
                    let prefix = element
                        .name()
                        .prefix()
                        .map_or(&[][..], |value| value.into_inner());
                    if !text_prefixes
                        .iter()
                        .any(|expected| expected.as_slice() == prefix)
                    {
                        return unsupported("plain paragraph has an unsafe text namespace binding");
                    }
                    paragraph = Some((depth, event_start));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODT paragraph transfer XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid("ODT paragraph transfer XML depth limit exceeded");
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if paragraph.is_some() {
                    return unsupported("paragraphs containing nested markup");
                }
                if depth == 0 {
                    return unsupported("empty or noncanonical content.xml root ownership");
                }
                if preamble_depth.is_some() {
                    return unsupported("unknown content.xml preamble descendants");
                }
                if office && matches!(local.as_ref(), b"body" | b"text") {
                    return unsupported("empty or duplicate body/text ownership");
                } else if office && local.as_ref() == b"scripts" {
                    if depth != 1
                        || body_depth.is_some()
                        || text_depth.is_some()
                        || has_non_namespace_attributes(&reader, &element)?
                    {
                        return unsupported("documents containing scripts or events");
                    }
                    // The empty placeholder emitted by ordinary ODT writers
                    // carries no script or event payload and is inert.
                } else if depth == 1 && body_depth.is_none() {
                    if !office || !is_allowed_content_preamble(local.as_ref()) {
                        return unsupported("foreign or unknown content.xml preamble markup");
                    }
                    if matches!(local.as_ref(), b"font-face-decls" | b"automatic-styles")
                        && has_non_namespace_attributes(&reader, &element)?
                    {
                        return unsupported("preamble placeholder attributes");
                    }
                } else if body_depth.is_some() && text_depth.is_none() {
                    return unsupported("non-text office:body children");
                } else if let Some(owner_depth) = text_depth {
                    if depth != owner_depth.saturating_add(1) || !text || local.as_ref() != b"p" {
                        return unsupported("non-plain office:text children");
                    }
                    if has_namespace_declaration(&element)
                        || has_non_namespace_attributes(&reader, &element)?
                    {
                        return unsupported("plain paragraph attributes, styles, or identifiers");
                    }
                    let prefix = element
                        .name()
                        .prefix()
                        .map_or(&[][..], |value| value.into_inner());
                    if !text_prefixes
                        .iter()
                        .any(|expected| expected.as_slice() == prefix)
                    {
                        return unsupported("plain paragraph has an unsafe text namespace binding");
                    }
                    push_paragraph(&mut paragraphs, event_start..event_end)?;
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODT paragraph transfer XML depth underflow"))?;
                let local = element.local_name();
                if preamble_depth == Some(depth) {
                    preamble_depth = None;
                }
                if let Some((paragraph_depth, start)) = paragraph {
                    if depth == paragraph_depth {
                        if !text || local.as_ref() != b"p" {
                            return invalid("ODT plain paragraph has a mismatched close tag");
                        }
                        push_paragraph(&mut paragraphs, start..event_end)?;
                        paragraph = None;
                    }
                } else if text_depth == Some(depth) {
                    if !office || local.as_ref() != b"text" {
                        return invalid("ODT office:text has a mismatched close tag");
                    }
                    text_end = Some(event_start);
                    text_depth = None;
                } else if body_depth == Some(depth) {
                    if !office || local.as_ref() != b"body" {
                        return invalid("ODT office:body has a mismatched close tag");
                    }
                    body_depth = None;
                }
            },
            Event::Text(value) => {
                if paragraph.is_none()
                    && value
                        .as_ref()
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace())
                {
                    return unsupported("text outside direct paragraphs");
                }
            },
            Event::CData(_) if paragraph.is_some() => {},
            Event::GeneralRef(reference) if paragraph.is_some() => {
                let name = reference.as_ref();
                if !is_allowed_character_reference(name) {
                    return unsupported("unknown entity references");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                return unsupported("opaque content outside direct paragraphs");
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {
                if preamble_depth.is_some()
                    || paragraph.is_some()
                    || text_depth.is_some()
                    || (body_depth.is_some() && text_depth.is_none())
                {
                    return unsupported("opaque markup in ODT content");
                }
            },
            Event::DocType(_) => return unsupported("documents containing a doctype"),
            Event::Eof => break,
        }
        buffer.clear();
    }
    if depth != 0
        || paragraph.is_some()
        || text_depth.is_some()
        || body_depth.is_some()
        || preamble_depth.is_some()
    {
        return invalid("ODT paragraph transfer XML is truncated");
    }
    if !saw_root || !saw_body || !saw_text {
        return unsupported("missing content.xml/office:body/office:text ownership");
    }
    let text_end =
        text_end.ok_or_else(|| invalid_error("ODT paragraph transfer text end missing"))?;
    Ok(PlainContent {
        paragraphs,
        text_prefixes,
        text_end,
    })
}

fn require_raw_untouched_members(
    source: &crate::core::OwnedPackage,
    target: &crate::core::OwnedPackage,
) -> Result<()> {
    let identical = raw_identical_members(source.as_bytes(), target.as_bytes())
        .ok_or_else(|| invalid_error("ODT paragraph transfer cannot audit raw ZIP members"))?;
    let mut source_paths = source.package()?.files()?;
    let mut target_paths = target.package()?.files()?;
    source_paths.sort();
    target_paths.sort();
    if source_paths != target_paths {
        return unsupported("packages whose member set changes during publication");
    }
    for path in source_paths {
        if path != ODF_CONTENT && !identical.contains(&path) {
            return unsupported("packages that cannot raw-preserve untouched members");
        }
    }
    Ok(())
}

fn push_paragraph(paragraphs: &mut Vec<Range<usize>>, range: Range<usize>) -> Result<()> {
    if paragraphs.len() >= MAX_PLAIN_PARAGRAPH_TRANSFER_PARAGRAPHS {
        return invalid("ODT paragraph transfer paragraph limit exceeded");
    }
    paragraphs
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "ODT paragraph transfer paragraph index",
            source,
        })?;
    paragraphs.push(range);
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("ODT paragraph transfer XML position overflow"))
}

fn is_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}

fn text_namespace_prefixes(reader: &NsReader<&[u8]>) -> Result<Vec<Vec<u8>>> {
    let mut prefixes = Vec::new();
    for (declaration, namespace) in reader.resolver().bindings() {
        if namespace.as_ref() != TEXT_NS {
            continue;
        }
        let prefix = match declaration {
            PrefixDeclaration::Default => &[][..],
            PrefixDeclaration::Named(prefix) => prefix,
        };
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(prefix.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT paragraph transfer text namespace prefix",
                source,
            })?;
        owned.extend_from_slice(prefix);
        prefixes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT paragraph transfer text namespace prefixes",
                source,
            })?;
        prefixes.push(owned);
    }
    Ok(prefixes)
}

fn has_mce_attribute(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODT content.xml attribute: {error}"))
        })?;
        if is_bound(
            &reader.resolver().resolve_attribute(attribute.key).0,
            MCE_NS,
        ) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn has_non_namespace_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODT content.xml attribute: {error}"))
        })?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        // Namespace declarations are not yielded by current quick-xml
        // resolver versions, but the explicit check keeps this boundary
        // conservative across resolver implementations.
        let _ = reader.resolver().resolve_attribute(attribute.key);
        return Ok(true);
    }
    Ok(false)
}

fn has_namespace_declaration(element: &BytesStart<'_>) -> bool {
    element.as_ref().windows(5).any(|window| window == b"xmlns")
}

fn is_allowed_character_reference(reference: &[u8]) -> bool {
    matches!(reference, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported<T>(what: &str) -> Result<T> {
    Err(Error::Unsupported(format!(
        "ODT paragraph transfer refuses {what}"
    )))
}
