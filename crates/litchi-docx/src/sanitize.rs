//! Source-bound external main-document hyperlink-wrapper detachment.
//!
//! This module owns only detachment of `w:hyperlink` wrappers whose `r:id`
//! resolves to an external hyperlink relationship owned by the main document.
//! Visible child markup is retained. The relationship records are not removed
//! because the current source-backed OPC publisher owns same-topology payload
//! overlays, not relationship-part or member deletion. The effect report makes
//! that retained relationship state explicit.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;
use std::sync::Arc;

const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const DEFAULT_MAX_DOCUMENT_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_DEPTH: usize = 128;
const DEFAULT_MAX_ELEMENTS: usize = 1_000_000;
const DEFAULT_MAX_EXTERNAL_HYPERLINKS: usize = 4_096;

/// Resource limits for external relationship-backed hyperlink-wrapper
/// detachment in the main document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_document_bytes: usize,
    max_depth: usize,
    max_elements: usize,
    max_external_hyperlinks: usize,
}

impl Limits {
    /// Return the maximum accepted main-document XML size.
    #[must_use]
    pub const fn max_document_bytes(self) -> usize {
        self.max_document_bytes
    }

    /// Return the maximum accepted XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Return the maximum number of XML elements scanned.
    #[must_use]
    pub const fn max_elements(self) -> usize {
        self.max_elements
    }

    /// Return the maximum number of external hyperlinks detached.
    #[must_use]
    pub const fn max_external_hyperlinks(self) -> usize {
        self.max_external_hyperlinks
    }

    /// Set the maximum accepted main-document XML size.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero.
    pub fn with_max_document_bytes(mut self, value: usize) -> Result<Self> {
        require_nonzero("main-document byte", value)?;
        self.max_document_bytes = value;
        Ok(self)
    }

    /// Set the maximum accepted XML nesting depth.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero.
    pub fn with_max_depth(mut self, value: usize) -> Result<Self> {
        require_nonzero("XML depth", value)?;
        self.max_depth = value;
        Ok(self)
    }

    /// Set the maximum number of XML elements scanned.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero.
    pub fn with_max_elements(mut self, value: usize) -> Result<Self> {
        require_nonzero("XML element", value)?;
        self.max_elements = value;
        Ok(self)
    }

    /// Set the maximum number of external hyperlinks detached.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero.
    pub fn with_max_external_hyperlinks(mut self, value: usize) -> Result<Self> {
        require_nonzero("external-hyperlink", value)?;
        self.max_external_hyperlinks = value;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_document_bytes: DEFAULT_MAX_DOCUMENT_BYTES,
            max_depth: DEFAULT_MAX_DEPTH,
            max_elements: DEFAULT_MAX_ELEMENTS,
            max_external_hyperlinks: DEFAULT_MAX_EXTERNAL_HYPERLINKS,
        }
    }
}

/// Deterministic effects of external relationship-backed hyperlink-wrapper
/// detachment in the main document.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EffectReport {
    detached_hyperlinks: usize,
    restored_hyperlinks: usize,
    referenced_external_relationships: usize,
    retained_relationships: usize,
}

impl EffectReport {
    /// Number of external `w:hyperlink` wrappers removed.
    #[must_use]
    pub const fn detached_hyperlinks(self) -> usize {
        self.detached_hyperlinks
    }

    /// Number of external relationship-backed wrappers restored by an inverse
    /// patch. A forward detachment report returns zero.
    #[must_use]
    pub const fn restored_hyperlinks(self) -> usize {
        self.restored_hyperlinks
    }

    /// Number of distinct external hyperlink relationship IDs referenced by
    /// the wrappers affected in either direction.
    #[must_use]
    pub const fn referenced_external_relationships(self) -> usize {
        self.referenced_external_relationships
    }

    /// Number of those relationship records retained by the same-topology
    /// source-backed publisher. A forward detachment leaves no selected
    /// wrapper for them; an inverse restores wrappers without changing the
    /// relationship records. This API makes no claim about unknown markup that
    /// may also carry the ID. Litchi never fetches or executes the targets.
    #[must_use]
    pub const fn retained_relationships(self) -> usize {
        self.retained_relationships
    }

    /// Return whether applying the plan changes no main-document bytes.
    #[must_use]
    pub const fn is_noop(self) -> bool {
        self.detached_hyperlinks == 0 && self.restored_hyperlinks == 0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipState {
    id: String,
    relationship_type: String,
    target: String,
    external: bool,
}

impl RelationshipState {
    pub(crate) fn new(
        id: String,
        relationship_type: String,
        target: String,
        external: bool,
    ) -> Self {
        Self {
            id,
            relationship_type,
            target,
            external,
        }
    }
}

/// Immutable exact-source snapshot for external relationship-backed wrapper
/// detachment in the main document.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    relationships: Arc<Vec<RelationshipState>>,
    limits: Limits,
    wrappers: Arc<Vec<Wrapper>>,
    report: EffectReport,
}

impl Snapshot {
    pub(crate) fn from_source(
        xml: Arc<Vec<u8>>,
        mut relationships: Vec<RelationshipState>,
        limits: Limits,
    ) -> Result<Self> {
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self::parse(xml, Arc::new(relationships), limits)
    }

    fn parse(
        xml: Arc<Vec<u8>>,
        relationships: Arc<Vec<RelationshipState>>,
        limits: Limits,
    ) -> Result<Self> {
        let (wrappers, report) = scan(&xml, &relationships, limits)?;
        Ok(Self {
            xml,
            relationships,
            limits,
            wrappers: Arc::new(wrappers),
            report,
        })
    }

    /// Borrow the exact main-document XML bytes.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return the number of external hyperlink wrappers currently selected by
    /// this exact detachment operation.
    #[must_use]
    pub const fn external_hyperlink_count(&self) -> usize {
        self.report.detached_hyperlinks
    }

    /// Return whether two snapshots share the same retained XML allocation.
    #[must_use]
    pub fn shares_xml_allocation_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.xml, &other.xml)
    }

    /// Build a non-mutating plan for detaching every selected external
    /// hyperlink wrapper.
    #[must_use]
    pub fn plan(&self) -> SanitizePlan {
        SanitizePlan {
            source: self.clone(),
        }
    }
}

/// A non-mutating plan for explicitly detaching external relationship-backed
/// hyperlink wrappers from the main document.
#[derive(Debug, Clone)]
pub struct SanitizePlan {
    source: Snapshot,
}

impl SanitizePlan {
    /// Borrow deterministic predicted effects without changing the snapshot.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.source.report
    }

    /// Apply the explicit wrapper-detachment plan and produce a source-checked
    /// commit. The source snapshot is never mutated.
    ///
    /// # Errors
    ///
    /// Returns an error if the surgical result does not parse, retain visible
    /// text, or remove exactly the selected wrappers.
    pub fn apply(self) -> Result<Commit> {
        if self.source.wrappers.is_empty() {
            return Ok(Commit {
                snapshot: self.source.clone(),
                patch: Patch {
                    before_xml: Arc::clone(&self.source.xml),
                    after_xml: Arc::clone(&self.source.xml),
                    relationships: Arc::clone(&self.source.relationships),
                    limits: self.source.limits,
                    report: self.source.report,
                },
                report: self.source.report,
            });
        }

        let rewritten = rewrite(&self.source.xml, &self.source.wrappers)?;
        let before_text = crate::paragraph::extract_word_text(&self.source.xml)?;
        let after_text = crate::paragraph::extract_word_text(&rewritten)?;
        if before_text != after_text {
            return Err(invalid(
                "external-hyperlink detachment changed visible document text",
            ));
        }
        let snapshot = Snapshot::parse(
            Arc::new(rewritten),
            Arc::clone(&self.source.relationships),
            self.source.limits,
        )?;
        if snapshot.report.detached_hyperlinks != 0 {
            return Err(invalid(
                "external-hyperlink detachment left a selected hyperlink wrapper",
            ));
        }
        let report = self.source.report;
        Ok(Commit {
            patch: Patch {
                before_xml: Arc::clone(&self.source.xml),
                after_xml: Arc::clone(&snapshot.xml),
                relationships: Arc::clone(&self.source.relationships),
                limits: self.source.limits,
                report,
            },
            snapshot,
            report,
        })
    }
}

/// A successful external main-document hyperlink-wrapper detachment result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    report: EffectReport,
}

impl Commit {
    /// Borrow the main-document snapshot after wrapper detachment.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return the deterministic effects applied by this commit.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.report
    }

    /// Move the snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-source-checked reversible patch for external main-document
/// hyperlink-wrapper detachment.
#[derive(Debug, Clone)]
pub struct Patch {
    before_xml: Arc<Vec<u8>>,
    after_xml: Arc<Vec<u8>>,
    relationships: Arc<Vec<RelationshipState>>,
    limits: Limits,
    report: EffectReport,
}

impl Patch {
    /// Return whether this patch changes no bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        Arc::ptr_eq(&self.before_xml, &self.after_xml)
            || self.before_xml.as_ref() == self.after_xml.as_ref()
    }

    /// Return deterministic effects captured by the forward detachment.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.report
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before_xml: Arc::clone(&self.after_xml),
            after_xml: Arc::clone(&self.before_xml),
            relationships: Arc::clone(&self.relationships),
            limits: self.limits,
            report: EffectReport {
                detached_hyperlinks: self.report.restored_hyperlinks,
                restored_hyperlinks: self.report.detached_hyperlinks,
                referenced_external_relationships: self.report.referenced_external_relationships,
                retained_relationships: self.report.retained_relationships,
            },
        }
    }

    /// Apply this patch only to its exact main-document and relationship
    /// closure.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied snapshot is foreign or stale.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_ref() != self.before_xml.as_ref()
            || source.relationships.as_ref() != self.relationships.as_ref()
        {
            return Err(Error::ExternalHyperlinkDetachmentConflict);
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::parse(
            Arc::clone(&self.after_xml),
            Arc::clone(&self.relationships),
            self.limits,
        )
    }

    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }
}

#[derive(Debug, Clone)]
struct Wrapper {
    open: Range<usize>,
    close: Option<Range<usize>>,
}

#[derive(Debug)]
struct ActiveHyperlink {
    depth: usize,
    open: Range<usize>,
    relationship_id: Option<String>,
}

fn scan(
    xml: &[u8],
    relationships: &[RelationshipState],
    limits: Limits,
) -> Result<(Vec<Wrapper>, EffectReport)> {
    if xml.len() > limits.max_document_bytes {
        return Err(limit(
            "main-document XML bytes",
            limits.max_document_bytes,
            xml.len(),
        ));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut depth = 0usize;
    let mut elements = 0usize;
    let mut roots = 0usize;
    let mut active: Option<ActiveHyperlink> = None;
    let mut wrappers = Vec::new();
    let mut distinct_relationships = Vec::new();

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = position(&reader)?;

        match event {
            Event::Start(element) => {
                elements = add_element(elements, limits)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("main-document XML depth overflows usize"))?;
                if depth > limits.max_depth {
                    return Err(limit("main-document XML depth", limits.max_depth, depth));
                }
                validate_root(&namespace, &element, depth, &mut roots)?;
                if is_hyperlink(&namespace, element.local_name().as_ref()) {
                    if active.is_some() {
                        return Err(invalid("nested WordprocessingML hyperlinks are ambiguous"));
                    }
                    active = Some(ActiveHyperlink {
                        depth,
                        open: start..end,
                        relationship_id: external_relationship_id(
                            &element,
                            decoder,
                            &resolver,
                            relationships,
                        )?,
                    });
                }
            },
            Event::Empty(element) => {
                elements = add_element(elements, limits)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("main-document XML depth overflows usize"))?;
                if child_depth > limits.max_depth {
                    return Err(limit(
                        "main-document XML depth",
                        limits.max_depth,
                        child_depth,
                    ));
                }
                validate_root(&namespace, &element, child_depth, &mut roots)?;
                if is_hyperlink(&namespace, element.local_name().as_ref()) {
                    if active.is_some() {
                        return Err(invalid("nested WordprocessingML hyperlinks are ambiguous"));
                    }
                    if let Some(relationship_id) =
                        external_relationship_id(&element, decoder, &resolver, relationships)?
                    {
                        record_relationship(&mut distinct_relationships, relationship_id)?;
                        push_wrapper(
                            &mut wrappers,
                            Wrapper {
                                open: start..end,
                                close: None,
                            },
                            limits,
                        )?;
                    }
                }
            },
            Event::End(element) => {
                if is_hyperlink(&namespace, element.local_name().as_ref()) {
                    let Some(open) = active.take() else {
                        return Err(invalid(
                            "WordprocessingML hyperlink end tag has no matching start tag",
                        ));
                    };
                    if open.depth != depth {
                        return Err(invalid(
                            "WordprocessingML hyperlink nesting changed before its end tag",
                        ));
                    }
                    if let Some(relationship_id) = open.relationship_id {
                        record_relationship(&mut distinct_relationships, relationship_id)?;
                        push_wrapper(
                            &mut wrappers,
                            Wrapper {
                                open: open.open,
                                close: Some(start..end),
                            },
                            limits,
                        )?;
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("main-document XML has an unexpected end tag"))?;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "external hyperlink-wrapper detachment refuses document type declarations",
                ));
            },
            Event::PI(_) => {
                return Err(invalid(
                    "external hyperlink-wrapper detachment refuses processing instructions",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if depth != 0 || active.is_some() {
        return Err(invalid("main-document XML has an unterminated element"));
    }
    if roots != 1 {
        return Err(invalid(
            "main-document XML must contain exactly one w:document root",
        ));
    }
    wrappers.sort_unstable_by_key(|wrapper| wrapper.open.start);
    for pair in wrappers.windows(2) {
        let left_end = pair[0]
            .close
            .as_ref()
            .map_or(pair[0].open.end, |range| range.end);
        if left_end > pair[1].open.start {
            return Err(invalid(
                "external-hyperlink wrapper ranges overlap unexpectedly",
            ));
        }
    }
    distinct_relationships.sort_unstable();
    distinct_relationships.dedup();
    let report = EffectReport {
        detached_hyperlinks: wrappers.len(),
        restored_hyperlinks: 0,
        referenced_external_relationships: distinct_relationships.len(),
        retained_relationships: distinct_relationships.len(),
    };
    Ok((wrappers, report))
}

fn record_relationship(relationships: &mut Vec<String>, relationship_id: String) -> Result<()> {
    relationships
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "external-hyperlink relationship effects",
            source,
        })?;
    relationships.push(relationship_id);
    Ok(())
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    roots: &mut usize,
) -> Result<()> {
    if depth != 1 {
        return Ok(());
    }
    *roots = roots
        .checked_add(1)
        .ok_or_else(|| invalid("main-document root counter overflows usize"))?;
    if !is_wordprocessing_namespace(namespace) || element.local_name().as_ref() != b"document" {
        return Err(invalid(
            "external hyperlink-wrapper detachment source root is not w:document",
        ));
    }
    Ok(())
}

fn is_hyperlink(namespace: &ResolveResult<'_>, local_name: &[u8]) -> bool {
    is_wordprocessing_namespace(namespace) && local_name == b"hyperlink"
}

fn external_relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    relationships: &[RelationshipState],
) -> Result<Option<String>> {
    let mut id = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_relationships_namespace(&namespace) {
            continue;
        }
        if id.is_some() {
            return Err(invalid(
                "WordprocessingML hyperlink has duplicate relationship IDs",
            ));
        }
        id = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }

    let Some(id) = id else {
        return Ok(None);
    };
    let relationship = relationships
        .binary_search_by(|candidate| candidate.id.as_str().cmp(&id))
        .ok()
        .map(|index| &relationships[index])
        .ok_or_else(|| {
            invalid(format!(
                "WordprocessingML hyperlink references missing relationship '{id}'"
            ))
        })?;
    if !matches!(
        relationship.relationship_type.as_str(),
        litchi_opc::constants::relationship_type::HYPERLINK
            | litchi_opc::constants::relationship_type::STRICT_HYPERLINK
    ) {
        return Err(invalid(format!(
            "WordprocessingML hyperlink relationship '{id}' has the wrong type"
        )));
    }
    if !relationship.external {
        return Ok(None);
    }
    validate_removable_wrapper_attributes(element, resolver)?;
    Ok(Some(id))
}

fn validate_removable_wrapper_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let mut seen = 0u8;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            return Err(unsafe_wrapper_attribute(
                "namespace declarations may scope retained descendants",
            ));
        }
        let (namespace, local_name) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == XML_NAMESPACE)
            || raw_name.starts_with(b"xml:")
        {
            return Err(unsafe_wrapper_attribute(
                "inheritable xml:* attributes may change retained descendants",
            ));
        }

        let bit = if is_relationships_namespace(&namespace) && local_name.as_ref() == b"id" {
            1
        } else if is_wordprocessing_namespace(&namespace) {
            match local_name.as_ref() {
                b"anchor" => 1 << 1,
                b"docLocation" => 1 << 2,
                b"history" => 1 << 3,
                b"tgtFrame" => 1 << 4,
                b"tooltip" => 1 << 5,
                _ => {
                    return Err(unsafe_wrapper_attribute(
                        "unknown Word hyperlink wrapper attribute",
                    ));
                },
            }
        } else {
            return Err(unsafe_wrapper_attribute(
                "unknown hyperlink wrapper attribute",
            ));
        };
        if seen & bit != 0 {
            return Err(invalid(
                "external WordprocessingML hyperlink has a duplicate wrapper attribute",
            ));
        }
        seen |= bit;
    }
    Ok(())
}

fn unsafe_wrapper_attribute(reason: &'static str) -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "external_hyperlink_wrapper_detachment",
        reason,
    }
}

fn is_relationships_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == RELATIONSHIPS_NAMESPACE || *value == STRICT_RELATIONSHIPS_NAMESPACE
    )
}

fn add_element(current: usize, limits: Limits) -> Result<usize> {
    let next = current
        .checked_add(1)
        .ok_or_else(|| invalid("main-document element counter overflows usize"))?;
    if next > limits.max_elements {
        return Err(limit(
            "main-document XML elements",
            limits.max_elements,
            next,
        ));
    }
    Ok(next)
}

fn push_wrapper(wrappers: &mut Vec<Wrapper>, wrapper: Wrapper, limits: Limits) -> Result<()> {
    if wrappers.len() >= limits.max_external_hyperlinks {
        return Err(limit(
            "external hyperlinks",
            limits.max_external_hyperlinks,
            wrappers.len().saturating_add(1),
        ));
    }
    wrappers
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "external hyperlink-wrapper detachment plan",
            source,
        })?;
    wrappers.push(wrapper);
    Ok(())
}

fn rewrite(xml: &[u8], wrappers: &[Wrapper]) -> Result<Vec<u8>> {
    let removed = wrappers.iter().try_fold(0usize, |total, wrapper| {
        let total = total
            .checked_add(wrapper.open.len())
            .ok_or_else(|| invalid("external-hyperlink removed-byte total overflows usize"))?;
        total
            .checked_add(wrapper.close.as_ref().map_or(0, Range::len))
            .ok_or_else(|| invalid("external-hyperlink removed-byte total overflows usize"))
    })?;
    let capacity = xml
        .len()
        .checked_sub(removed)
        .ok_or_else(|| invalid("external-hyperlink removal ranges exceed the source"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "detached-hyperlink main-document XML",
            source,
        })?;
    let mut cursor = 0usize;
    for wrapper in wrappers {
        append_before_and_skip(&mut output, xml, &mut cursor, &wrapper.open)?;
        if let Some(close) = &wrapper.close {
            output.extend_from_slice(
                xml.get(cursor..close.start).ok_or_else(|| {
                    invalid("external-hyperlink close range is outside the source")
                })?,
            );
            cursor = close.end;
        }
    }
    output.extend_from_slice(
        xml.get(cursor..)
            .ok_or_else(|| invalid("external-hyperlink rewrite cursor exceeds the source"))?,
    );
    Ok(output)
}

fn append_before_and_skip(
    output: &mut Vec<u8>,
    xml: &[u8],
    cursor: &mut usize,
    range: &Range<usize>,
) -> Result<()> {
    if range.start < *cursor || range.end < range.start || range.end > xml.len() {
        return Err(invalid(
            "external-hyperlink removal range is invalid or overlaps",
        ));
    }
    output.extend_from_slice(&xml[*cursor..range.start]);
    *cursor = range.end;
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| invalid("main-document XML offset does not fit usize"))
}

fn require_nonzero(resource: &str, value: usize) -> Result<()> {
    if value == 0 {
        return Err(invalid(format!(
            "{resource} external hyperlink-detachment limit must be nonzero"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn limit(resource: &'static str, maximum: usize, actual: usize) -> Error {
    Error::ExternalHyperlinkDetachmentLimit {
        resource,
        maximum,
        actual,
    }
}
