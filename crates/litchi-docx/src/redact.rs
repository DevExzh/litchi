//! Explicit irreversible redaction of external main-document hyperlinks.
//!
//! This API is deliberately separate from reversible hyperlink-wrapper
//! detachment. A caller first inventories the exact source closure, then
//! selects target URLs, applies a non-mutating plan, and finally invokes the
//! consuming source-backed publisher. Applying a redaction unwraps selected
//! `w:hyperlink` elements so their visible child markup remains, and removes
//! every selected external relationship record. There is intentionally no
//! inverse operation.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use crate::sanitize::{self, RelationshipState};
use litchi_core::SourceVersion;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::SourceArtifactFingerprint;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest as _, Sha256};
use std::sync::Arc;

const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_SELECTED_RELATIONSHIPS: usize = 64;
const MAX_SELECTOR_BYTES: usize = MAX_SELECTED_RELATIONSHIPS * 4 * 1024;

/// One deterministic external hyperlink relationship in the main document.
#[derive(Clone, PartialEq, Eq)]
pub struct ExternalHyperlinkRelationship {
    relationship_id: String,
    target_url: String,
    wrapper_count: usize,
}

impl ExternalHyperlinkRelationship {
    /// Relationship identifier used by `w:hyperlink` wrappers.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Inert external target. Litchi never fetches or executes this value.
    #[must_use]
    pub fn target_url(&self) -> &str {
        &self.target_url
    }

    /// Number of main-document wrappers referencing this relationship.
    #[must_use]
    pub const fn wrapper_count(&self) -> usize {
        self.wrapper_count
    }
}

impl std::fmt::Debug for ExternalHyperlinkRelationship {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ExternalHyperlinkRelationship")
            .field("relationship_id_bytes", &self.relationship_id.len())
            .field("target_url_bytes", &self.target_url.len())
            .field("wrapper_count", &self.wrapper_count)
            .finish_non_exhaustive()
    }
}

/// Deterministic effects predicted or applied by one redaction.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EffectReport {
    selected_targets: usize,
    removed_relationships: usize,
    unwrapped_hyperlinks: usize,
}

impl EffectReport {
    /// Number of distinct target URL values selected.
    #[must_use]
    pub const fn selected_targets(self) -> usize {
        self.selected_targets
    }

    /// Number of external relationship records removed.
    #[must_use]
    pub const fn removed_relationships(self) -> usize {
        self.removed_relationships
    }

    /// Number of `w:hyperlink` wrappers unwrapped while retaining children.
    #[must_use]
    pub const fn unwrapped_hyperlinks(self) -> usize {
        self.unwrapped_hyperlinks
    }

    /// Return whether publication reproduces the exact source artifact.
    #[must_use]
    pub const fn is_noop(self) -> bool {
        self.removed_relationships == 0 && self.unwrapped_hyperlinks == 0
    }
}

/// Exact-source immutable inventory used to plan irreversible redaction.
#[derive(Clone)]
pub struct Snapshot {
    closure: sanitize::Snapshot,
    source_version: SourceVersion,
    source_fingerprint: SourceArtifactFingerprint,
    inventory: Arc<Vec<ExternalHyperlinkRelationship>>,
}

impl Snapshot {
    pub(crate) fn from_source(
        xml: Arc<Vec<u8>>,
        relationships: Vec<RelationshipState>,
        source_version: SourceVersion,
        source_fingerprint: SourceArtifactFingerprint,
        limits: sanitize::Limits,
    ) -> Result<Self> {
        let closure = sanitize::Snapshot::from_source(xml, relationships, limits)?;
        validate_main_document_closure(&closure)?;
        let inventory = build_inventory(&closure, limits)?;
        Ok(Self {
            closure,
            source_version,
            source_fingerprint,
            inventory: Arc::new(inventory),
        })
    }

    /// Parse a standalone main-document redaction closure under an explicit
    /// bounded XML policy.
    pub fn from_xml(
        xml: Vec<u8>,
        relationships: Vec<sanitize::raw::Relationship>,
        limits: sanitize::Limits,
    ) -> Result<Self> {
        let fingerprint = fingerprint(&xml);
        let closure = sanitize::Snapshot::from_xml(xml, relationships, limits)?;
        validate_main_document_closure(&closure)?;
        let inventory = build_inventory(&closure, limits)?;
        Ok(Self {
            closure,
            source_version: SourceVersion::new(0, 0),
            source_fingerprint: fingerprint,
            inventory: Arc::new(inventory),
        })
    }

    /// Borrow the deterministic relationship inventory, ordered by ID.
    #[must_use]
    pub fn relationships(&self) -> &[ExternalHyperlinkRelationship] {
        &self.inventory
    }

    /// Borrow exact main-document bytes without copying them.
    #[must_use]
    pub fn document_xml(&self) -> &[u8] {
        self.closure.xml_bytes()
    }

    /// Plan removal of every relationship whose target exactly equals one of
    /// `target_urls`. Duplicate selectors are folded deterministically. An
    /// unknown target is rejected rather than becoming a silent no-op.
    pub fn plan_target_urls(&self, target_urls: &[&str]) -> Result<Plan> {
        if target_urls.len() > MAX_SELECTED_RELATIONSHIPS {
            return Err(Error::ExternalHyperlinkRedactionLimit {
                kind: LimitKind::TargetUrlSelectors,
                maximum: MAX_SELECTED_RELATIONSHIPS,
                actual: target_urls.len(),
            });
        }
        let selector_bytes = target_urls.iter().try_fold(0usize, |total, target| {
            total
                .checked_add(target.len())
                .ok_or(Error::ExternalHyperlinkRedactionLimit {
                    kind: LimitKind::TargetUrlSelectorBytes,
                    maximum: MAX_SELECTOR_BYTES,
                    actual: usize::MAX,
                })
        })?;
        if selector_bytes > MAX_SELECTOR_BYTES {
            return Err(Error::ExternalHyperlinkRedactionLimit {
                kind: LimitKind::TargetUrlSelectorBytes,
                maximum: MAX_SELECTOR_BYTES,
                actual: selector_bytes,
            });
        }
        let mut targets = Vec::new();
        targets
            .try_reserve_exact(target_urls.len())
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink redaction target selectors",
                source,
            })?;
        for target in target_urls {
            let mut owned = String::new();
            owned
                .try_reserve_exact(target.len())
                .map_err(|source| Error::Allocation {
                    resource: "external-hyperlink redaction target selector",
                    source,
                })?;
            owned.push_str(target);
            targets.push(owned);
        }
        targets.sort_unstable();
        targets.dedup();
        for target in &targets {
            if !self.inventory.iter().any(|item| item.target_url == *target) {
                return Err(Error::InvalidFormat(
                    "external-hyperlink redaction target is not present".into(),
                ));
            }
        }

        let mut selected_ids = Vec::new();
        let mut wrapper_count = 0usize;
        for item in self
            .inventory
            .iter()
            .filter(|item| targets.binary_search(&item.target_url).is_ok())
        {
            selected_ids
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "external-hyperlink redaction relationship selection",
                    source,
                })?;
            selected_ids.push(item.relationship_id.clone());
            wrapper_count = wrapper_count
                .checked_add(item.wrapper_count)
                .ok_or_else(|| {
                    Error::InvalidFormat("external-hyperlink wrapper count overflow".into())
                })?;
        }
        if selected_ids.len() > MAX_SELECTED_RELATIONSHIPS {
            return Err(Error::ExternalHyperlinkRedactionLimit {
                kind: LimitKind::SelectedRelationships,
                maximum: MAX_SELECTED_RELATIONSHIPS,
                actual: selected_ids.len(),
            });
        }
        selected_ids.sort_unstable();
        Ok(Plan {
            source: self.clone(),
            selected_ids,
            report: EffectReport {
                selected_targets: targets.len(),
                removed_relationships: self
                    .inventory
                    .iter()
                    .filter(|item| targets.binary_search(&item.target_url).is_ok())
                    .count(),
                unwrapped_hyperlinks: wrapper_count,
            },
        })
    }

    pub(crate) const fn limits(&self) -> sanitize::Limits {
        self.closure.limits()
    }

    fn exact_source_matches(&self, other: &Self) -> bool {
        self.source_version == other.source_version
            && self.source_fingerprint == other.source_fingerprint
            && self.closure.xml_bytes() == other.closure.xml_bytes()
            && self.closure.relationships() == other.closure.relationships()
    }
}

/// Non-mutating, exact-source plan for irreversible external-link redaction.
#[derive(Clone)]
pub struct Plan {
    source: Snapshot,
    selected_ids: Vec<String>,
    report: EffectReport,
}

impl Plan {
    /// Borrow predicted effects before applying the plan.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.report
    }

    /// Produce a sealed forward-only commit. This does not mutate the source.
    /// There is intentionally no inverse API.
    pub fn apply(self) -> Result<IrreversibleRedactionPatch> {
        if self.selected_ids.is_empty() {
            let source = self.source;
            return Ok(IrreversibleRedactionPatch {
                replacement_xml: source.closure.shared_xml(),
                source,
                removed_relationship_ids: self.selected_ids,
                report: self.report,
            });
        }

        let wrapper_count = self
            .source
            .closure
            .wrappers()
            .iter()
            .filter(|wrapper| {
                self.selected_ids
                    .binary_search(&wrapper.relationship_id)
                    .is_ok()
            })
            .count();
        let mut wrappers = Vec::new();
        wrappers
            .try_reserve_exact(wrapper_count)
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink redaction wrapper plan",
                source,
            })?;
        for wrapper in self.source.closure.wrappers().iter().filter(|wrapper| {
            self.selected_ids
                .binary_search(&wrapper.relationship_id)
                .is_ok()
        }) {
            wrappers.push(wrapper.clone());
        }
        let rewritten = sanitize::rewrite(self.source.document_xml(), &wrappers)?;
        if crate::paragraph::extract_word_text(self.source.document_xml())?
            != crate::paragraph::extract_word_text(&rewritten)?
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "external_hyperlink_redaction",
                "redaction changed visible main-document text",
            ));
        }
        let retained_count = self
            .source
            .closure
            .relationships()
            .iter()
            .filter(|relationship| {
                self.selected_ids
                    .binary_search_by(|id| id.as_str().cmp(relationship.id()))
                    .is_err()
            })
            .count();
        let mut retained = Vec::new();
        retained
            .try_reserve_exact(retained_count)
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink redaction retained relationships",
                source,
            })?;
        for relationship in self
            .source
            .closure
            .relationships()
            .iter()
            .filter(|relationship| {
                self.selected_ids
                    .binary_search_by(|id| id.as_str().cmp(relationship.id()))
                    .is_err()
            })
        {
            retained.push(relationship.clone());
        }
        let after = sanitize::Snapshot::from_source(
            Arc::new(rewritten.clone()),
            retained,
            self.source.limits(),
        )?;
        validate_main_document_closure(&after)?;
        Ok(IrreversibleRedactionPatch {
            source: self.source,
            replacement_xml: Arc::new(rewritten),
            removed_relationship_ids: self.selected_ids,
            report: self.report,
        })
    }
}

/// Exact-source-checked, sealed, forward-only redaction patch.
///
/// This type intentionally has no public apply or inverse method. Publication
/// is available only through the consuming source-backed package API.
#[derive(Clone)]
pub struct IrreversibleRedactionPatch {
    source: Snapshot,
    replacement_xml: Arc<Vec<u8>>,
    removed_relationship_ids: Vec<String>,
    report: EffectReport,
}

impl IrreversibleRedactionPatch {
    /// Return whether this patch changes no package bytes.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.report.is_noop()
    }

    /// Deterministic effects captured by the forward-only patch.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.report
    }

    pub(crate) fn validate_source(&self, current: &Snapshot) -> Result<()> {
        if self.source.exact_source_matches(current) {
            Ok(())
        } else {
            Err(Error::conflict(
                "external_hyperlink_redaction",
                "the patch base does not match the supplied source",
            ))
        }
    }
}

pub(crate) fn publication_parts(patch: &IrreversibleRedactionPatch) -> (Vec<u8>, Vec<String>) {
    (
        patch.replacement_xml.as_ref().clone(),
        patch.removed_relationship_ids.clone(),
    )
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("xml_bytes", &self.closure.xml_bytes().len())
            .field("relationships", &self.inventory.len())
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Plan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Plan")
            .field("effect_report", &self.effect_report())
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for IrreversibleRedactionPatch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("IrreversibleRedactionPatch")
            .field("noop", &self.is_noop())
            .field("effect_report", &self.effect_report())
            .field(
                "removed_relationships",
                &self.removed_relationship_ids.len(),
            )
            .finish_non_exhaustive()
    }
}

/// Standalone bounded forward-only redaction codec.
pub mod raw {
    use super::{EffectReport, IrreversibleRedactionPatch, Plan, Result, Snapshot};
    use crate::sanitize::Limits;

    /// Mutable raw state for one main-document redaction closure.
    #[derive(Clone)]
    pub struct Codec {
        snapshot: Snapshot,
        replacement_xml: Vec<u8>,
        removed_relationship_ids: Vec<String>,
        report: EffectReport,
        applied: bool,
    }

    impl std::fmt::Debug for Codec {
        fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            formatter
                .debug_struct("Codec")
                .field("xml_bytes", &self.replacement_xml.len())
                .field(
                    "removed_relationships",
                    &self.removed_relationship_ids.len(),
                )
                .field("applied", &self.applied)
                .finish_non_exhaustive()
        }
    }

    impl Codec {
        /// Parse a standalone closure under the required finite policy.
        pub fn from_xml(
            xml: Vec<u8>,
            relationships: Vec<crate::sanitize::raw::Relationship>,
            limits: Limits,
        ) -> Result<Self> {
            let snapshot = Snapshot::from_xml(xml, relationships, limits)?;
            Self::from_snapshot(snapshot, limits)
        }

        /// Adopt a parsed snapshot under an explicitly repeated policy.
        pub fn from_snapshot(snapshot: Snapshot, limits: Limits) -> Result<Self> {
            if snapshot.limits() != limits {
                return Err(super::Error::InvalidFormat(
                    "raw redactor limits do not match the snapshot".into(),
                ));
            }
            let replacement_xml = snapshot.document_xml().to_vec();
            Ok(Self {
                snapshot,
                replacement_xml,
                removed_relationship_ids: Vec::new(),
                report: EffectReport::default(),
                applied: false,
            })
        }

        /// Borrow the current inventory snapshot.
        #[must_use]
        pub const fn snapshot(&self) -> &Snapshot {
            &self.snapshot
        }

        /// Borrow the current XML bytes.
        #[must_use]
        pub fn xml_bytes(&self) -> &[u8] {
            &self.replacement_xml
        }

        /// Relationship identifiers removed from the current raw output.
        #[must_use]
        pub fn removed_relationship_ids(&self) -> &[String] {
            &self.removed_relationship_ids
        }

        /// Plan exact target URL removal over the current state.
        pub fn plan_target_urls(&self, target_urls: &[&str]) -> Result<Plan> {
            self.snapshot.plan_target_urls(target_urls)
        }

        /// Apply a forward-only patch to the owned raw state.
        pub fn apply(&mut self, patch: &IrreversibleRedactionPatch) -> Result<()> {
            if self.applied {
                return Err(super::Error::InvalidFormat(
                    "raw redactor accepts one forward patch".into(),
                ));
            }
            patch.validate_source(&self.snapshot)?;
            self.replacement_xml = patch.replacement_xml.to_vec();
            self.removed_relationship_ids = patch.removed_relationship_ids.clone();
            self.report = patch.report;
            self.applied = true;
            Ok(())
        }

        /// Plan and apply exact target URL removal in one mutable operation.
        pub fn redact(&mut self, target_urls: &[&str]) -> Result<IrreversibleRedactionPatch> {
            let patch = self.plan_target_urls(target_urls)?.apply()?;
            self.apply(&patch)?;
            Ok(patch)
        }

        /// Effects of the most recently applied patch.
        #[must_use]
        pub const fn effect_report(&self) -> EffectReport {
            self.report
        }
    }
}

fn fingerprint(xml: &[u8]) -> SourceArtifactFingerprint {
    let mut digest = Sha256::new();
    digest.update(xml);
    SourceArtifactFingerprint::from_sha256(digest.finalize().into())
}

fn build_inventory(
    closure: &sanitize::Snapshot,
    limits: sanitize::Limits,
) -> Result<Vec<ExternalHyperlinkRelationship>> {
    let external = closure
        .relationships()
        .iter()
        .filter(|relationship| relationship.is_external())
        .count();
    if external > limits.max_external_hyperlinks() {
        return Err(Error::ExternalHyperlinkRedactionLimit {
            kind: LimitKind::ExternalHyperlinkRelationships,
            maximum: limits.max_external_hyperlinks(),
            actual: external,
        });
    }
    let mut inventory = Vec::new();
    inventory
        .try_reserve_exact(external)
        .map_err(|source| Error::Allocation {
            resource: "external-hyperlink redaction inventory",
            source,
        })?;
    for relationship in closure
        .relationships()
        .iter()
        .filter(|relationship| relationship.is_external())
    {
        if !matches!(
            relationship.relationship_type(),
            litchi_opc::constants::relationship_type::HYPERLINK
                | litchi_opc::constants::relationship_type::STRICT_HYPERLINK
        ) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "external_hyperlink_redaction",
                "main document has an external relationship outside the hyperlink closure",
            ));
        }
        inventory.push(ExternalHyperlinkRelationship {
            relationship_id: relationship.id().to_owned(),
            target_url: relationship.target().to_owned(),
            wrapper_count: closure
                .wrappers()
                .iter()
                .filter(|wrapper| wrapper.relationship_id == relationship.id())
                .count(),
        });
    }
    inventory.sort_unstable_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(inventory)
}

fn validate_main_document_closure(closure: &sanitize::Snapshot) -> Result<()> {
    // The scan resolves names on namespace values as written; a binding that
    // spells the markup-compatibility namespace with references is invisible
    // to it, while readers apply it.
    if litchi_ooxml_common::mce::binds_namespace_by_reference(closure.xml_bytes()) {
        return unsafe_closure(
            "markup-compatibility owner syntax is outside the redaction closure",
        );
    }
    let mut reader = NsReader::from_reader(closure.xml_bytes());
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut external_hyperlink_depth = None;
    let mut depth = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let starts_external = inspect_owner_element(
                    closure,
                    &element,
                    &namespace,
                    &resolver,
                    decoder,
                    external_hyperlink_depth.is_some(),
                )?;
                depth = depth.saturating_add(1);
                if starts_external {
                    external_hyperlink_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                let _starts_external = inspect_owner_element(
                    closure,
                    &element,
                    &namespace,
                    &resolver,
                    decoder,
                    external_hyperlink_depth.is_some(),
                )?;
            },
            Event::End(_) => {
                if external_hyperlink_depth == Some(depth) {
                    external_hyperlink_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            Event::DocType(_) | Event::PI(_) => {
                return unsafe_closure(
                    "DTD and processing-instruction syntax is outside the redaction closure",
                );
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(())
}

fn inspect_owner_element(
    closure: &sanitize::Snapshot,
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    inside_external_hyperlink: bool,
) -> Result<bool> {
    if is_mce_namespace(namespace) {
        return unsafe_closure(
            "markup-compatibility owner syntax is outside the redaction closure",
        );
    }
    let is_word = is_wordprocessing_namespace(namespace);
    let local = element.local_name();
    if is_word && matches!(local.as_ref(), b"fldSimple" | b"fldChar" | b"instrText") {
        return unsafe_closure(
            "field, DDE, and external-field forms are outside the redaction closure",
        );
    }
    if inside_external_hyperlink && !is_word {
        return unsafe_closure("unknown markup occurs inside an external hyperlink owner");
    }
    let is_hyperlink = is_word && local.as_ref() == b"hyperlink";
    let mut external_id = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (attribute_namespace, attribute_local) = resolver.resolve_attribute(attribute.key);
        if is_mce_namespace(&attribute_namespace) {
            return unsafe_closure(
                "markup-compatibility attributes are outside the redaction closure",
            );
        }
        if is_relationships_namespace(&attribute_namespace) {
            let id = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            if closure
                .relationships()
                .iter()
                .any(|relationship| relationship.id() == id && relationship.is_external())
            {
                if !is_hyperlink || attribute_local.as_ref() != b"id" {
                    return unsafe_closure(
                        "an external relationship ID is referenced outside w:hyperlink r:id",
                    );
                }
                external_id = true;
            }
        }
    }
    Ok(external_id)
}

fn is_relationships_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == RELATIONSHIPS_NAMESPACE || *value == STRICT_RELATIONSHIPS_NAMESPACE
    )
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE_NAMESPACE)
}

fn unsafe_closure<T>(reason: &'static str) -> Result<T> {
    Err(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "external_hyperlink_redaction",
        reason,
    ))
}

/// The bounded resources of external-hyperlink redaction, carried by [`crate::Error::ExternalHyperlinkRedactionLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Changed story parts.
    ChangedStoryParts,
    /// External hyperlink relationships.
    ExternalHyperlinkRelationships,
    /// External story hyperlink relationships.
    ExternalStoryHyperlinkRelationships,
    /// Selected relationships.
    SelectedRelationships,
    /// Selected story hyperlink relationships.
    SelectedStoryHyperlinkRelationships,
    /// Target URL selector bytes.
    TargetUrlSelectorBytes,
    /// Target URL selectors.
    TargetUrlSelectors,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ChangedStoryParts => "changed story parts",
            Self::ExternalHyperlinkRelationships => "external hyperlink relationships",
            Self::ExternalStoryHyperlinkRelationships => "external story hyperlink relationships",
            Self::SelectedRelationships => "selected relationships",
            Self::SelectedStoryHyperlinkRelationships => "selected story hyperlink relationships",
            Self::TargetUrlSelectorBytes => "target URL selector bytes",
            Self::TargetUrlSelectors => "target URL selectors",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
