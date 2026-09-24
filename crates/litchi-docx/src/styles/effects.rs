//! Source-preserving ownership of the optional Word stylesWithEffects part.
//!
//! The part is a copy of the Word style definitions part. This module owns
//! its package topology and a bounded, read-only style projection; it does not
//! attempt to model the complete visual-effects vocabulary or rewrite style
//! definitions individually.

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{
    ContentTypeEdit, OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, ReadLimits,
    ReadResource, RelationshipEdit, TargetMode, validate_source_xml_bytes,
};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::styles::{Style, Styles, Type};
use crate::{Error, Result};

#[cfg(test)]
mod test_trace {
    use std::cell::Cell;

    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
    pub(super) struct Counts {
        pub(super) loads: usize,
        pub(super) candidate_clones: usize,
        pub(super) projection_builds: usize,
        pub(super) readbacks: usize,
        pub(super) graph_validations: usize,
    }

    thread_local! {
        static COUNTS: Cell<Counts> = const {
            Cell::new(Counts {
                loads: 0,
                candidate_clones: 0,
                projection_builds: 0,
                readbacks: 0,
                graph_validations: 0,
            })
        };
    }

    pub(super) fn reset() {
        COUNTS.with(|counts| counts.set(Counts::default()));
    }

    pub(super) fn take() -> Counts {
        COUNTS.with(|counts| counts.get())
    }

    pub(super) fn record_load() {
        COUNTS.with(|counts| {
            let mut value = counts.get();
            value.loads = value.loads.saturating_add(1);
            counts.set(value);
        });
    }

    pub(super) fn record_candidate_clone() {
        COUNTS.with(|counts| {
            let mut value = counts.get();
            value.candidate_clones = value.candidate_clones.saturating_add(1);
            counts.set(value);
        });
    }

    pub(super) fn record_projection_build() {
        COUNTS.with(|counts| {
            let mut value = counts.get();
            value.projection_builds = value.projection_builds.saturating_add(1);
            counts.set(value);
        });
    }

    pub(super) fn record_readback() {
        COUNTS.with(|counts| {
            let mut value = counts.get();
            value.readbacks = value.readbacks.saturating_add(1);
            counts.set(value);
        });
    }

    pub(super) fn record_graph_validation() {
        COUNTS.with(|counts| {
            let mut value = counts.get();
            value.graph_validations = value.graph_validations.saturating_add(1);
            counts.set(value);
        });
    }
}

/// Microsoft's stylesWithEffects relationship type.
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
/// Content type used by a Word stylesWithEffects part.
pub const CONTENT_TYPE: &str = "application/vnd.ms-word.stylesWithEffects+xml";

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const GLOSSARY_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_GLOSSARY_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_EVENTS: usize = 1_000_000;
const MAX_XML_DEPTH: usize = 256;
const MAX_EFFECTS_PARTS: usize = 2;
const MAX_PART_NAME_ATTEMPTS: usize = 4096;

/// The semantic package owner of one optional effects resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Owner {
    /// The main Word document part.
    MainDocument,
    /// The glossary document part, when the package has one.
    Glossary,
}

impl Owner {
    #[inline]
    const fn index(self) -> usize {
        match self {
            Self::MainDocument => 0,
            Self::Glossary => 1,
        }
    }
}

impl fmt::Display for Owner {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::MainDocument => formatter.write_str("main document"),
            Self::Glossary => formatter.write_str("glossary"),
        }
    }
}

/// Word XML conformance family carried by a style definitions document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Conformance {
    /// Transitional WordprocessingML namespace.
    Transitional,
    /// Strict WordprocessingML namespace.
    Strict,
}

impl Conformance {
    fn from_namespace(namespace: &str) -> Option<Self> {
        match namespace {
            WORD_NAMESPACE => Some(Self::Transitional),
            STRICT_WORD_NAMESPACE => Some(Self::Strict),
            _ => None,
        }
    }
}

/// A bounded, typed read projection of the style definitions in an effects
/// resource.
///
/// The projection intentionally contains only the fields already understood
/// by the ordinary styles reader. The source XML in Resource remains
/// authoritative for every unmodeled style child and attribute.
#[derive(Debug, Clone)]
pub struct StyleDefinitions {
    styles: Arc<[Style]>,
}

/// Short alias for the read-only style projection.
pub type Projection = StyleDefinitions;

impl StyleDefinitions {
    fn from_xml(xml: &Arc<Vec<u8>>) -> Result<Self> {
        #[cfg(test)]
        test_trace::record_projection_build();

        // Reuse the established style reader for the typed projection. The
        // temporary part adopts the same Arc allocation; it is not a second
        // source copy and is dropped after the projection is materialized.
        let part_name = PackURI::new("/word/stylesWithEffects.xml")
            .map_err(|error| Error::InvalidUri(error.to_string()))?;
        let part = BlobPart::new_shared(part_name, CONTENT_TYPE.to_owned(), Arc::clone(xml));
        let mut styles = Styles::from_part(&part);
        let values = styles.iter()?.cloned().collect::<Vec<_>>();
        Ok(Self {
            styles: Arc::from(values.into_boxed_slice()),
        })
    }

    /// Number of projected style definitions.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.styles.len()
    }

    /// Whether the projection contains no style definitions.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }

    /// Iterate over style definitions in source order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, Style> {
        self.styles.iter()
    }

    /// Look up one style by its stable style ID.
    #[must_use]
    pub fn get_by_id(&self, style_id: &str) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.style_id() == style_id)
    }

    /// Look up one style by its UI name.
    #[must_use]
    pub fn get_by_name(&self, name: &str) -> Option<&Style> {
        self.styles.iter().find(|style| style.name() == Some(name))
    }

    /// Return the default style for a style kind.
    #[must_use]
    pub fn get_default(&self, style_type: Type) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.is_default() && style.style_type() == style_type)
    }

    /// Resolve inherited paragraph numbering with the ordinary style bounds.
    ///
    /// This is a read projection only; no style inheritance or visual cascade
    /// is authored by this owner.
    pub fn resolved_numbering(
        &self,
        style_id: &str,
    ) -> Result<Option<crate::numbering::Paragraph>> {
        let mut current = style_id;
        let mut visited = HashSet::new();
        loop {
            if !visited.insert(current.to_owned()) {
                return Err(Error::InvalidFormat(format!(
                    "style basedOn cycle at '{current}'"
                )));
            }
            if visited.len() > self.styles.len().saturating_add(1) {
                return Err(Error::InvalidFormat(
                    "style inheritance exceeds the style table".to_owned(),
                ));
            }
            let Some(style) = self.get_by_id(current) else {
                return Ok(None);
            };
            if let Some(numbering) = style.numbering() {
                return Ok(Some(numbering));
            }
            let Some(parent) = style.based_on() else {
                return Ok(None);
            };
            current = parent;
        }
    }
}

/// One complete, validated stylesWithEffects resource.
#[derive(Debug, Clone)]
pub struct Resource {
    xml: Arc<Vec<u8>>,
    styles: StyleDefinitions,
    conformance: Conformance,
}

impl PartialEq for Resource {
    fn eq(&self, other: &Self) -> bool {
        self.conformance == other.conformance && self.xml.as_slice() == other.xml.as_slice()
    }
}

impl Eq for Resource {}

impl Resource {
    /// Validate and retain a caller-owned effects XML source.
    ///
    /// The source must have a w:styles root in either the Transitional or
    /// Strict WordprocessingML namespace. Unknown children and attributes are
    /// retained in the exact source bytes.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = Arc::new(xml.into());
        let conformance = validate_xml(&xml, None)?;
        Self::from_shared_with_conformance(xml, conformance)
    }

    /// Alias for Resource::new emphasizing its source-backed nature.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::new(xml)
    }

    /// Borrow the exact authored XML source.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Borrow the typed, read-only style projection.
    #[inline]
    #[must_use]
    pub const fn styles(&self) -> &StyleDefinitions {
        &self.styles
    }

    /// Alias for Resource::styles.
    #[inline]
    #[must_use]
    pub const fn definitions(&self) -> &StyleDefinitions {
        &self.styles
    }

    /// Alias for Resource::styles using the projection terminology.
    #[inline]
    #[must_use]
    pub const fn projection(&self) -> &StyleDefinitions {
        &self.styles
    }

    /// Return the Word XML conformance family.
    #[inline]
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    pub(crate) fn blob_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    pub(crate) fn from_part(
        part: &dyn Part,
        expected: Conformance,
        limits: ReadLimits,
    ) -> Result<Self> {
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        Self::from_shared_with_conformance_limits(part.blob_arc(), expected, limits)
    }

    pub(crate) fn from_shared_with_conformance(
        xml: Arc<Vec<u8>>,
        expected: Conformance,
    ) -> Result<Self> {
        let conformance = validate_xml(&xml, Some(expected))?;
        let styles = StyleDefinitions::from_xml(&xml)?;
        Ok(Self {
            xml,
            styles,
            conformance,
        })
    }

    pub(crate) fn from_shared_with_conformance_limits(
        xml: Arc<Vec<u8>>,
        expected: Conformance,
        limits: ReadLimits,
    ) -> Result<Self> {
        let conformance = validate_xml_with_limits(&xml, Some(expected), limits)?;
        let styles = StyleDefinitions::from_xml(&xml)?;
        Ok(Self {
            xml,
            styles,
            conformance,
        })
    }
}

/// An immutable owner-wide snapshot. Resource absence is valid and is exposed
/// through resource() returning None.
#[derive(Debug, Clone)]
pub struct Snapshot {
    owner: Owner,
    resource: Option<Resource>,
    conformance: Conformance,
    graph: Option<GraphToken>,
}

impl Snapshot {
    /// Construct an owner snapshot from an optional validated resource.
    pub fn new(owner: Owner, resource: Option<Resource>) -> Result<Self> {
        let conformance = resource
            .as_ref()
            .map_or(Conformance::Transitional, Resource::conformance);
        Ok(Self {
            owner,
            resource,
            conformance,
            graph: None,
        })
    }

    /// Construct an empty owner snapshot.
    #[must_use]
    pub const fn empty(owner: Owner) -> Self {
        Self {
            owner,
            resource: None,
            conformance: Conformance::Transitional,
            graph: None,
        }
    }

    /// Construct a snapshot from a standalone XML resource.
    pub fn from_xml(owner: Owner, xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::new(owner, Some(Resource::new(xml)?))
    }

    /// Semantic owner represented by this snapshot.
    #[inline]
    #[must_use]
    pub const fn owner(&self) -> Owner {
        self.owner
    }

    /// Optional effects resource, with exact source bytes when present.
    #[inline]
    #[must_use]
    pub const fn resource(&self) -> Option<&Resource> {
        self.resource.as_ref()
    }

    /// Whether this owner has no effects resource.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.resource.is_none()
    }

    /// Conformance of the resource, or the package-default Transitional value
    /// for an absent standalone snapshot.
    #[inline]
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Start an isolated edit over this optional resource.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.resource.clone(),
        }
    }

    pub(crate) fn from_package(
        owner: Owner,
        resource: Option<Resource>,
        conformance: Conformance,
        graph: Option<GraphToken>,
    ) -> Self {
        Self {
            owner,
            resource,
            conformance,
            graph,
        }
    }

    pub(crate) fn same_state(&self, other: &Self) -> bool {
        self.owner == other.owner
            && self.conformance == other.conformance
            && same_resource(self.resource.as_ref(), other.resource.as_ref())
            && self.graph == other.graph
    }
}

/// An isolated replacement transaction over one owner-wide resource state.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Option<Resource>,
}

impl Transaction {
    /// Borrow the currently staged optional resource.
    #[inline]
    #[must_use]
    pub const fn resource(&self) -> Option<&Resource> {
        self.next.as_ref()
    }

    /// Replace the complete resource, or pass None to stage removal.
    pub fn replace_resource(&mut self, resource: Option<Resource>) -> Result<&mut Self> {
        if let Some(resource) = &resource {
            if resource.conformance() != self.base.conformance {
                return Err(Error::InvalidFormat(
                    "stylesWithEffects resource conformance does not match its snapshot".to_owned(),
                ));
            }
        }
        self.next = resource;
        Ok(self)
    }

    /// Clear the staged resource.
    pub fn clear_resource(&mut self) -> &mut Self {
        self.next = None;
        self
    }

    /// Whether the staged owner state differs from its source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_resource(self.base.resource.as_ref(), self.next.as_ref())
    }

    /// Validate and produce a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.is_changed();
        if !changed {
            let patch = Patch {
                owner: self.base.owner,
                conformance: self.base.conformance,
                before: self.base.resource.clone(),
                after: self.base.resource.clone(),
                before_graph: self.base.graph.clone(),
                after_graph: self.base.graph.clone(),
                package_bound: self.base.graph.is_some(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
                changed: false,
            });
        }

        let after_graph = graph_after_transaction(&self.base, self.next.as_ref())?;
        let package_bound = self.base.graph.is_some();
        let next = self.next;
        let snapshot = Snapshot {
            owner: self.base.owner,
            conformance: self.base.conformance,
            resource: next,
            // A transaction does not invent a new physical target. Package
            // publication binds the after-state to its actual graph.
            graph: after_graph.clone().or_else(|| self.base.graph.clone()),
        };
        let patch = Patch {
            owner: self.base.owner,
            conformance: self.base.conformance,
            before: self.base.resource,
            after: snapshot.resource.clone(),
            before_graph: self.base.graph,
            after_graph,
            package_bound,
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful source-bound owner publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the transaction changed the owner resource.
    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed()
    }

    /// Borrow the owner snapshot produced by the edit.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the owner snapshot out of this commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible source patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible source patch out of this commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A source-checked reversible replacement/removal patch.
///
/// Package-bound patches retain the originating package's [`ReadLimits`]. Applying
/// one to a package opened with a different limit profile fails its source
/// precondition; create a fresh edit from that package instead.
#[derive(Debug, Clone)]
pub struct Patch {
    owner: Owner,
    conformance: Conformance,
    before: Option<Resource>,
    after: Option<Resource>,
    before_graph: Option<GraphToken>,
    after_graph: Option<GraphToken>,
    package_bound: bool,
}

impl Patch {
    /// Owner to which this patch is bound.
    #[inline]
    #[must_use]
    pub const fn owner(&self) -> Owner {
        self.owner
    }

    /// Optional resource expected before application.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> Option<&Resource> {
        self.before.as_ref()
    }

    /// Optional resource produced by application.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> Option<&Resource> {
        self.after.as_ref()
    }

    /// Return the inverse source-checked operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            owner: self.owner,
            conformance: self.conformance,
            before: self.after.clone(),
            after: self.before.clone(),
            before_graph: self.after_graph.clone(),
            after_graph: self.before_graph.clone(),
            package_bound: self.package_bound,
        }
    }

    /// Apply to an exact owner snapshot.
    pub fn apply_to_snapshot(&self, source: &Snapshot) -> Result<Snapshot> {
        let graph_matches = match (self.before_graph.as_ref(), source.graph.as_ref()) {
            (Some(expected), Some(actual)) => expected.matches_source_precondition(actual),
            (None, None) => !self.package_bound,
            _ => false,
        };
        if source.owner != self.owner
            || source.conformance != self.conformance
            || !same_resource(source.resource.as_ref(), self.before.as_ref())
            || !graph_matches
            || (!self.package_bound && source.graph.is_some())
        {
            return Err(Error::InvalidFormat(
                "stylesWithEffects patch source does not match its owner, resource, or graph precondition"
                    .to_owned(),
            ));
        }
        Ok(Snapshot {
            owner: self.owner,
            resource: self.after.clone(),
            conformance: source.conformance,
            graph: self.after_graph.clone().or_else(|| source.graph.clone()),
        })
    }

    /// Apply this source-checked patch atomically to its owner in a package.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        apply_patch(package, self.owner, self)
    }

    /// Apply the exact inverse of this patch atomically to a package.
    pub fn undo(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        self.inverse().apply(package)
    }

    /// Whether this patch is an exact no-op.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        same_resource(self.before.as_ref(), self.after.as_ref())
    }

    /// Alias for [`Self::is_empty`].
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }
}

#[derive(Debug, Clone)]
pub(crate) struct GraphToken {
    owner: Owner,
    source: Option<String>,
    target: Option<String>,
    relationship_id: Option<String>,
    planned_target: Option<String>,
    planned_relationship_id: Option<String>,
    relationships: Option<OwnedRelationships>,
    target_relationships: Option<OwnedRelationships>,
    content_types: OwnedContentTypes,
    aggregate: AggregateMetrics,
    limits: ReadLimits,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct AggregateMetrics {
    parts: usize,
    total_part_bytes: u64,
    total_relationships: usize,
    relationship_parts: usize,
    relationship_graph_nodes: usize,
    relationship_xml_bytes: u64,
    relationship_xml_events: usize,
}

#[derive(Debug, Clone, Copy)]
struct AggregateEdit {
    part_delta: i64,
    graph_node_delta: i64,
    removed_part_len: Option<usize>,
    added_part_len: Option<usize>,
    removed_relationship_parts: usize,
    added_relationship_parts: usize,
    removed_relationships: usize,
    added_relationships: usize,
    removed_relationship_events: usize,
    added_relationship_events: usize,
    removed_relationship_bytes: usize,
    added_relationship_bytes: usize,
    content_type_len: usize,
    content_type_mappings: usize,
    content_type_events: usize,
}

fn project_aggregate_metrics(
    current: AggregateMetrics,
    limits: ReadLimits,
    edit: AggregateEdit,
) -> Result<AggregateMetrics> {
    let parts = apply_signed_delta(current.parts, edit.part_delta, "part count")?;
    let relationship_graph_nodes = apply_signed_delta(
        current.relationship_graph_nodes,
        edit.graph_node_delta,
        "relationship graph node count",
    )?;
    let relationship_parts = current
        .relationship_parts
        .checked_sub(edit.removed_relationship_parts)
        .and_then(|value| value.checked_add(edit.added_relationship_parts))
        .ok_or_else(|| invalid("stylesWithEffects relationship part count underflow"))?;
    let total_relationships = current
        .total_relationships
        .checked_sub(edit.removed_relationships)
        .and_then(|value| value.checked_add(edit.added_relationships))
        .ok_or_else(|| invalid("stylesWithEffects relationship count underflow"))?;
    let relationship_xml_events = current
        .relationship_xml_events
        .checked_sub(edit.removed_relationship_events)
        .and_then(|value| value.checked_add(edit.added_relationship_events))
        .ok_or_else(|| invalid("stylesWithEffects relationship event count underflow"))?;
    let relationship_xml_bytes = replace_u64_metric(
        current.relationship_xml_bytes,
        edit.removed_relationship_bytes,
        edit.added_relationship_bytes,
        "relationship bytes",
    )?;
    let total_part_bytes = replace_u64_metric(
        current.total_part_bytes,
        edit.removed_part_len.unwrap_or(0),
        edit.added_part_len.unwrap_or(0),
        "part bytes",
    )?;
    if let Some(length) = edit.added_part_len {
        check_limit(
            limits,
            ReadResource::PartBytes,
            u64::try_from(length)
                .map_err(|_| invalid("stylesWithEffects part byte count overflows u64"))?,
            limits.max_part_bytes(),
        )?;
    }
    check_limit(
        limits,
        ReadResource::Parts,
        parts as u64,
        limits.max_parts() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalPartBytes,
        total_part_bytes,
        limits.max_total_part_bytes(),
    )?;
    check_limit(
        limits,
        ReadResource::RelationshipParts,
        relationship_parts as u64,
        limits.max_relationship_parts() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationships,
        total_relationships as u64,
        limits.max_total_relationships() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::RelationshipGraphNodes,
        relationship_graph_nodes as u64,
        limits.max_relationship_graph_nodes() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlBytes,
        relationship_xml_bytes,
        limits.max_total_relationship_xml_bytes() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlEvents,
        relationship_xml_events as u64,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::ContentTypesBytes,
        edit.content_type_len as u64,
        limits.max_content_types_bytes() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::ContentTypeMappings,
        edit.content_type_mappings as u64,
        limits.max_content_type_mappings() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::XmlEvents,
        edit.content_type_events as u64,
        limits.max_xml_events() as u64,
    )?;
    Ok(AggregateMetrics {
        parts,
        total_part_bytes,
        total_relationships,
        relationship_parts,
        relationship_graph_nodes,
        relationship_xml_bytes,
        relationship_xml_events,
    })
}

fn apply_signed_delta(current: usize, delta: i64, label: &str) -> Result<usize> {
    if delta >= 0 {
        current
            .checked_add(
                usize::try_from(delta)
                    .map_err(|_| invalid(format!("stylesWithEffects {label} overflows")))?,
            )
            .ok_or_else(|| invalid(format!("stylesWithEffects {label} overflows")))
    } else {
        current
            .checked_sub(
                usize::try_from(delta.unsigned_abs())
                    .map_err(|_| invalid(format!("stylesWithEffects {label} underflows")))?,
            )
            .ok_or_else(|| invalid(format!("stylesWithEffects {label} underflows")))
    }
}

fn replace_u64_metric(current: u64, removed: usize, added: usize, label: &str) -> Result<u64> {
    let removed = u64::try_from(removed)
        .map_err(|_| invalid(format!("stylesWithEffects {label} overflows u64")))?;
    let added = u64::try_from(added)
        .map_err(|_| invalid(format!("stylesWithEffects {label} overflows u64")))?;
    current
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| invalid(format!("stylesWithEffects {label} underflow")))
}

impl PartialEq for GraphToken {
    fn eq(&self, other: &Self) -> bool {
        self.owner == other.owner
            && self.source == other.source
            && self.target == other.target
            && self.relationship_id == other.relationship_id
            && self.planned_target == other.planned_target
            && self.planned_relationship_id == other.planned_relationship_id
            && self.relationships == other.relationships
            && self.target_relationships == other.target_relationships
            && self.content_types == other.content_types
            && self.limits == other.limits
    }
}

impl Eq for GraphToken {}

impl GraphToken {
    fn replaced(&self, old_part_len: usize, new_part_len: usize) -> Result<Self> {
        // A replacement keeps the owner edge, target part, content-type
        // manifest, and relationship members unchanged.  Still project the
        // selected part's byte delta against the package aggregate before a
        // replacement commit can expose a graph token.  Planning a no-op
        // content-types edit supplies the retained manifest metrics without
        // materializing a new XML buffer.
        let content_type_plan = self.content_types.plan_edit(&[], &[], self.limits)?;
        let aggregate = project_aggregate_metrics(
            self.aggregate,
            self.limits,
            AggregateEdit {
                part_delta: 0,
                graph_node_delta: 0,
                removed_part_len: Some(old_part_len),
                added_part_len: Some(new_part_len),
                removed_relationship_parts: 0,
                added_relationship_parts: 0,
                removed_relationships: 0,
                added_relationships: 0,
                removed_relationship_events: 0,
                added_relationship_events: 0,
                removed_relationship_bytes: 0,
                added_relationship_bytes: 0,
                content_type_len: content_type_plan.final_len(),
                content_type_mappings: content_type_plan.mapping_count(),
                content_type_events: content_type_plan.event_count(),
            },
        )?;
        let mut replaced = self.clone();
        replaced.aggregate = aggregate;
        Ok(replaced)
    }

    fn added(&self, new_part_len: usize) -> Result<Self> {
        let (Some(source), Some(target), Some(relationship_id), Some(relationships)) = (
            self.source.as_deref(),
            self.planned_target.as_deref(),
            self.planned_relationship_id.as_deref(),
            self.relationships.as_ref(),
        ) else {
            return Ok(self.clone());
        };
        let source = PackURI::new(source.to_owned()).map_err(Error::Uri)?;
        let target = PackURI::new(target.to_owned()).map_err(Error::Uri)?;
        let target_ref = target.relative_ref(source.base_uri());
        let addition = RelationshipEdit {
            id: relationship_id,
            reltype: RELATIONSHIP_TYPE,
            target: target_ref.as_str(),
            mode: TargetMode::Internal,
        };
        let relationship_plan = relationships.plan_edit(&[addition], &[], self.limits)?;
        let content_type_addition = ContentTypeEdit {
            part_name: &target,
            content_type: CONTENT_TYPE,
        };
        let content_type_plan =
            self.content_types
                .plan_edit(&[content_type_addition], &[], self.limits)?;
        let (old_events, old_count) = relationship_metrics(self.relationships_bytes())?;
        let new_events = relationship_metrics_after_edit(self.relationships_bytes(), 1, &[])?.0;
        let (target_events, target_count) = self
            .target_relationships
            .as_ref()
            .map_or(Ok((0, 0)), |target_relationships| {
                relationship_metrics(target_relationships.bytes())
            })?;
        let target_relationships_present = self
            .target_relationships
            .as_ref()
            .is_some_and(OwnedRelationships::member_present);
        let aggregate = project_aggregate_metrics(
            self.aggregate,
            self.limits,
            AggregateEdit {
                part_delta: 1,
                graph_node_delta: 1,
                removed_part_len: None,
                added_part_len: Some(new_part_len),
                removed_relationship_parts: usize::from(self.relationships_member_present()),
                added_relationship_parts: usize::from(relationship_plan.member_present())
                    + usize::from(target_relationships_present),
                removed_relationships: old_count,
                added_relationships: relationship_plan
                    .relationship_count()
                    .checked_add(target_count)
                    .ok_or_else(|| invalid("stylesWithEffects relationship count overflows"))?,
                removed_relationship_events: if self.relationships_member_present() {
                    old_events
                } else {
                    0
                },
                added_relationship_events: (if relationship_plan.member_present() {
                    new_events
                } else {
                    0
                })
                .checked_add(if target_relationships_present {
                    target_events
                } else {
                    0
                })
                .ok_or_else(|| invalid("stylesWithEffects relationship event count overflows"))?,
                removed_relationship_bytes: if self.relationships_member_present() {
                    self.relationships_bytes().len()
                } else {
                    0
                },
                added_relationship_bytes: (if relationship_plan.member_present() {
                    relationship_plan.final_len()
                } else {
                    0
                })
                .checked_add(if target_relationships_present {
                    self.target_relationships
                        .as_ref()
                        .map_or(0, |relationships| relationships.bytes().len())
                } else {
                    0
                })
                .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflows"))?,
                content_type_len: content_type_plan.final_len(),
                content_type_mappings: content_type_plan.mapping_count(),
                content_type_events: content_type_plan.event_count(),
            },
        )?;
        let relationships = relationship_plan.materialize(self.limits)?;
        let content_types = content_type_plan.materialize(self.limits)?;
        Ok(Self {
            owner: self.owner,
            source: self.source.clone(),
            target: Some(target.as_str().to_owned()),
            relationship_id: Some(relationship_id.to_owned()),
            planned_target: None,
            planned_relationship_id: None,
            relationships: Some(relationships),
            target_relationships: self.target_relationships.clone(),
            content_types,
            aggregate,
            limits: self.limits,
        })
    }

    fn removed(&self, removed_part_len: usize) -> Result<Self> {
        let (Some(target), Some(relationship_id), Some(relationships)) = (
            self.target.as_deref(),
            self.relationship_id.as_deref(),
            self.relationships.as_ref(),
        ) else {
            return Ok(self.clone());
        };
        let target = PackURI::new(target.to_owned()).map_err(Error::Uri)?;
        let relationship_plan = relationships.plan_edit(&[], &[relationship_id], self.limits)?;
        let content_type_plan =
            self.content_types
                .plan_edit(&[], std::slice::from_ref(&target), self.limits)?;
        let target_relationships = self
            .target_relationships
            .as_ref()
            .ok_or_else(|| invalid("stylesWithEffects removal graph lacks target provenance"))?;
        let (old_events, old_count) = relationship_metrics(self.relationships_bytes())?;
        let new_events =
            relationship_metrics_after_edit(self.relationships_bytes(), 0, &[relationship_id])?.0;
        let (target_events, target_count) = relationship_metrics(target_relationships.bytes())?;
        let removed_relationships = old_count
            .checked_add(target_count)
            .ok_or_else(|| invalid("stylesWithEffects relationship count overflows"))?;
        let removed_relationship_events = (if self.relationships_member_present() {
            old_events
        } else {
            0
        })
        .checked_add(if target_relationships.member_present() {
            target_events
        } else {
            0
        })
        .ok_or_else(|| invalid("stylesWithEffects relationship event count overflows"))?;
        let removed_relationship_bytes = (if self.relationships_member_present() {
            self.relationships_bytes().len()
        } else {
            0
        })
        .checked_add(if target_relationships.member_present() {
            target_relationships.bytes().len()
        } else {
            0
        })
        .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflows"))?;
        let aggregate = project_aggregate_metrics(
            self.aggregate,
            self.limits,
            AggregateEdit {
                part_delta: -1,
                graph_node_delta: -1,
                removed_part_len: Some(removed_part_len),
                added_part_len: None,
                removed_relationship_parts: usize::from(self.relationships_member_present())
                    + usize::from(target_relationships.member_present()),
                added_relationship_parts: usize::from(relationship_plan.member_present()),
                removed_relationships,
                added_relationships: relationship_plan.relationship_count(),
                removed_relationship_events,
                added_relationship_events: if relationship_plan.member_present() {
                    new_events
                } else {
                    0
                },
                removed_relationship_bytes,
                added_relationship_bytes: if relationship_plan.member_present() {
                    relationship_plan.final_len()
                } else {
                    0
                },
                content_type_len: content_type_plan.final_len(),
                content_type_mappings: content_type_plan.mapping_count(),
                content_type_events: content_type_plan.event_count(),
            },
        )?;
        let relationships = relationship_plan.materialize(self.limits)?;
        let content_types = content_type_plan.materialize(self.limits)?;
        Ok(Self {
            owner: self.owner,
            source: self.source.clone(),
            target: None,
            relationship_id: None,
            planned_target: Some(target.as_str().to_owned()),
            planned_relationship_id: Some(relationship_id.to_owned()),
            relationships: Some(relationships),
            target_relationships: self.target_relationships.clone(),
            content_types,
            aggregate,
            limits: self.limits,
        })
    }

    fn relationships_bytes(&self) -> &[u8] {
        self.relationships
            .as_ref()
            .map_or(&[], OwnedRelationships::bytes)
    }

    fn relationships_member_present(&self) -> bool {
        self.relationships
            .as_ref()
            .is_some_and(OwnedRelationships::member_present)
    }

    fn same_published_state(&self, other: &Self) -> bool {
        let target_relationships_match = match (
            self.target_relationships.as_ref(),
            other.target_relationships.as_ref(),
        ) {
            (Some(left), Some(right)) => left == right,
            _ => true,
        };
        self.owner == other.owner
            && self.source == other.source
            && self.target == other.target
            && self.relationship_id == other.relationship_id
            && self.relationships == other.relationships
            && target_relationships_match
            && self.content_types == other.content_types
            && self.limits == other.limits
    }

    fn same_published_state_with_aggregate(&self, other: &Self) -> bool {
        self.same_published_state(other) && self.aggregate == other.aggregate
    }

    fn matches_source_precondition(&self, actual: &Self) -> bool {
        if self.target.is_none()
            && self.relationship_id.is_none()
            && actual.target.is_none()
            && actual.relationship_id.is_none()
        {
            self.same_published_state(actual)
        } else {
            let target_relationships_match = match (
                self.target_relationships.as_ref(),
                actual.target_relationships.as_ref(),
            ) {
                (Some(expected), Some(actual)) => expected == actual,
                (None, Some(actual)) => !actual.member_present(),
                _ => false,
            };
            self.owner == actual.owner
                && self.source == actual.source
                && self.target == actual.target
                && self.relationship_id == actual.relationship_id
                && self.relationships == actual.relationships
                && target_relationships_match
                && self.content_types == actual.content_types
                && self.limits == actual.limits
        }
    }
}

fn graph_after_transaction(base: &Snapshot, next: Option<&Resource>) -> Result<Option<GraphToken>> {
    let Some(graph) = base.graph.as_ref() else {
        return Ok(None);
    };
    match (base.resource.is_some(), next.is_some()) {
        (true, false) => graph
            .removed(
                base.resource
                    .as_ref()
                    .map_or(0, |resource| resource.xml.len()),
            )
            .map(Some),
        (true, true) => graph
            .replaced(
                base.resource
                    .as_ref()
                    .map_or(0, |resource| resource.xml.len()),
                next.map_or(0, |resource| resource.xml.len()),
            )
            .map(Some),
        (false, false) => Ok(Some(graph.clone())),
        (false, true) => graph
            .added(next.map_or(0, |resource| resource.xml.len()))
            .map(Some),
    }
}

#[derive(Debug, Clone)]
struct Binding {
    source: PackURI,
    target: PackURI,
    relationship_id: String,
}

#[derive(Debug)]
struct GraphState {
    conformance: Conformance,
    main: PackURI,
    glossary: Option<PackURI>,
    bindings: [Option<Binding>; 2],
}

impl GraphState {
    fn binding(&self, owner: Owner) -> Option<&Binding> {
        self.bindings[owner.index()].as_ref()
    }
}

fn capture_aggregate_metrics(package: &OpcPackage) -> Result<AggregateMetrics> {
    let limits = package.read_limits();
    let parts = package.part_count();
    // The aggregate is over actual payload bytes, so every part is decoded
    // here (ADR 0030).
    let total_part_bytes = package.try_iter_parts().try_fold(0_u64, |total, part| {
        let part = part?;
        total
            .checked_add(
                u64::try_from(part.blob().len())
                    .map_err(|_| invalid("stylesWithEffects part byte count overflows u64"))?,
            )
            .ok_or_else(|| invalid("stylesWithEffects part bytes overflow"))
    })?;
    let total_relationships = package
        .rels()
        .iter()
        .count()
        .checked_add(package.iter_parts().try_fold(0_usize, |total, part| {
            total
                .checked_add(part.rels().iter().count())
                .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))
        })?)
        .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?;
    let root = PackURI::new("/").map_err(Error::Uri)?;
    let root_token = package.source_relationships_with_limits(&root, limits)?;
    let mut relationship_parts = usize::from(root_token.member_present());
    let mut relationship_xml_bytes = 0_u64;
    let mut relationship_xml_events = 0_usize;
    if root_token.member_present() {
        relationship_xml_bytes = u64::try_from(root_token.bytes().len())
            .map_err(|_| invalid("stylesWithEffects relationship byte count overflows u64"))?;
        relationship_xml_events = relationship_metrics(root_token.bytes())?.0;
    }
    for part in package.iter_parts() {
        let token = package.source_relationships_with_limits(part.partname(), limits)?;
        if token.member_present() {
            relationship_parts = relationship_parts
                .checked_add(1)
                .ok_or_else(|| invalid("stylesWithEffects relationship part count overflow"))?;
            relationship_xml_bytes = relationship_xml_bytes
                .checked_add(u64::try_from(token.bytes().len()).map_err(|_| {
                    invalid("stylesWithEffects relationship byte count overflows u64")
                })?)
                .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflow"))?;
            relationship_xml_events = relationship_xml_events
                .checked_add(relationship_metrics(token.bytes())?.0)
                .ok_or_else(|| invalid("stylesWithEffects relationship event count overflow"))?;
        }
    }
    let metrics = AggregateMetrics {
        parts,
        total_part_bytes,
        total_relationships,
        relationship_parts,
        relationship_graph_nodes: relationship_graph_nodes(package)?,
        relationship_xml_bytes,
        relationship_xml_events,
    };
    check_limit(
        limits,
        ReadResource::Parts,
        metrics.parts as u64,
        limits.max_parts() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalPartBytes,
        metrics.total_part_bytes,
        limits.max_total_part_bytes(),
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationships,
        metrics.total_relationships as u64,
        limits.max_total_relationships() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::RelationshipParts,
        metrics.relationship_parts as u64,
        limits.max_relationship_parts() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlBytes,
        metrics.relationship_xml_bytes,
        limits.max_total_relationship_xml_bytes() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlEvents,
        metrics.relationship_xml_events as u64,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    Ok(metrics)
}

fn capture_graph_token(
    package: &OpcPackage,
    graph: &GraphState,
    owner: Owner,
) -> Result<GraphToken> {
    let source = match owner {
        Owner::MainDocument => Some(graph.main.clone()),
        Owner::Glossary => graph.glossary.clone(),
    };
    let relationships = source
        .as_ref()
        .map(|source| package.source_relationships_with_limits(source, package.read_limits()))
        .transpose()?;
    let content_types = package.source_content_types_with_limits(package.read_limits())?;
    let aggregate = capture_aggregate_metrics(package)?;
    let binding = graph.binding(owner);
    let target_relationships = binding
        .map(|binding| {
            package.source_relationships_with_limits(&binding.target, package.read_limits())
        })
        .transpose()?;
    let (planned_target, planned_relationship_id) = if binding.is_none() {
        let planned_target = source
            .as_ref()
            .map(|_| next_part_name(package, owner))
            .transpose()?;
        let planned_relationship_id = source
            .as_ref()
            .map(|source| next_relationship_id(package, source))
            .transpose()?;
        (
            planned_target.map(|value| value.as_str().to_owned()),
            planned_relationship_id,
        )
    } else {
        (None, None)
    };
    Ok(GraphToken {
        owner,
        source: source.map(|value| value.as_str().to_owned()),
        target: binding.map(|value| value.target.as_str().to_owned()),
        relationship_id: binding.map(|value| value.relationship_id.clone()),
        planned_target,
        planned_relationship_id,
        relationships,
        target_relationships,
        content_types,
        aggregate,
        limits: package.read_limits(),
    })
}

/// Load a complete owner-wide effects snapshot. Absence is represented by a
/// present snapshot whose resource() is None.
pub fn load(package: &OpcPackage, owner: Owner) -> Result<Snapshot> {
    #[cfg(test)]
    test_trace::record_load();

    let graph = inspect_graph(package)?;
    let resource = graph
        .binding(owner)
        .map(|binding| {
            Resource::from_part(
                package.get_part(&binding.target)?,
                graph.conformance,
                package.read_limits(),
            )
        })
        .transpose()?;
    let graph_token = Some(capture_graph_token(package, &graph, owner)?);
    Ok(Snapshot::from_package(
        owner,
        resource,
        graph.conformance,
        graph_token,
    ))
}

/// Source-backed spelling of load.
pub fn load_snapshot(package: &OpcPackage, owner: Owner) -> Result<Snapshot> {
    load(package, owner)
}

/// Apply a source-checked patch atomically to one package owner.
pub fn apply_patch(package: &mut OpcPackage, owner: Owner, patch: &Patch) -> Result<Snapshot> {
    let current = load(package, owner)?;
    if patch.owner != owner {
        return Err(Error::InvalidFormat(
            "stylesWithEffects patch owner does not match the requested package owner".to_owned(),
        ));
    }
    let projected = patch.apply_to_snapshot(&current)?;
    if current.same_state(&projected) {
        return Ok(current);
    }

    let mut candidate = {
        #[cfg(test)]
        test_trace::record_candidate_clone();
        package.clone()
    };
    let published = apply_projected_patch(&mut candidate, owner, patch, projected, Some(package))?;
    *package = candidate;
    Ok(published)
}

/// Apply a patch to the exact candidate clone owned by the package facade.
///
/// The caller must have loaded `current` from the same source package immediately
/// before entering the outer semantic edit, and `package` must be the candidate
/// clone created by that edit. This private seam is intentionally not a general
/// candidate admission API: it requires the package-bound graph token, recomputes
/// the source patch, and performs the complete publication/readback validation.
pub(crate) fn apply_patch_staged(
    package: &mut OpcPackage,
    owner: Owner,
    patch: &Patch,
    current: Snapshot,
) -> Result<Snapshot> {
    if current.graph.is_none() {
        return Err(invalid(
            "staged stylesWithEffects patch requires a package-bound current snapshot",
        ));
    }
    if patch.owner != owner {
        return Err(Error::InvalidFormat(
            "stylesWithEffects patch owner does not match the requested package owner".to_owned(),
        ));
    }
    let projected = patch.apply_to_snapshot(&current)?;
    if current.same_state(&projected) {
        return Ok(current);
    }
    apply_projected_patch(package, owner, patch, projected, None)
}

fn apply_projected_patch(
    package: &mut OpcPackage,
    owner: Owner,
    patch: &Patch,
    projected: Snapshot,
    signature_source: Option<&OpcPackage>,
) -> Result<Snapshot> {
    let replacement = is_existing_resource_replacement(patch, &projected);
    publish_resource(
        package,
        owner,
        projected.resource.clone(),
        patch.after_graph.as_ref(),
        replacement.then_some(patch.before.as_ref()).flatten(),
    )?;
    if let Some(source) = signature_source {
        package.validate_signature_edit_from(source)?;
    }
    let published = if replacement {
        readback_existing_resource_replacement(
            package,
            owner,
            projected
                .resource
                .as_ref()
                .ok_or_else(|| invalid("stylesWithEffects replacement lacks a resource"))?,
            patch
                .after_graph
                .as_ref()
                .ok_or_else(|| invalid("stylesWithEffects replacement lacks graph provenance"))?,
        )?
    } else {
        load(package, owner)?
    };
    let graph_matches = match (&published.graph, &projected.graph) {
        (Some(published), Some(projected)) => {
            published.same_published_state_with_aggregate(projected)
        },
        (None, None) => true,
        _ => false,
    };
    if !same_resource(published.resource.as_ref(), projected.resource.as_ref()) {
        return Err(invalid(
            "staged stylesWithEffects resource did not round-trip",
        ));
    }
    if !graph_matches {
        return Err(invalid("staged stylesWithEffects graph did not round-trip"));
    }
    Ok(published)
}

fn is_existing_resource_replacement(patch: &Patch, projected: &Snapshot) -> bool {
    patch.before.is_some()
        && projected.resource.is_some()
        && patch
            .before_graph
            .as_ref()
            .is_some_and(|graph| graph.target.is_some() && graph.relationship_id.is_some())
        && patch
            .after_graph
            .as_ref()
            .is_some_and(|graph| graph.target.is_some() && graph.relationship_id.is_some())
}

fn readback_existing_resource_replacement(
    package: &OpcPackage,
    owner: Owner,
    expected_resource: &Resource,
    desired_graph: &GraphToken,
) -> Result<Snapshot> {
    #[cfg(test)]
    test_trace::record_readback();

    let graph = inspect_graph_with_validated_resource(package, Some((owner, expected_resource)))?;
    let source = match owner {
        Owner::MainDocument => graph.main.clone(),
        Owner::Glossary => graph.glossary.clone().ok_or_else(|| {
            Error::PartNotFound(
                "glossary document owner is required before reading stylesWithEffects".to_owned(),
            )
        })?,
    };
    if expected_resource.conformance != graph.conformance {
        return Err(invalid(
            "stylesWithEffects replacement conformance does not match the package",
        ));
    }
    if desired_graph.owner != owner || desired_graph.source.as_deref() != Some(source.as_str()) {
        return Err(invalid(
            "stylesWithEffects replacement graph belongs to a different owner source",
        ));
    }
    let binding = graph
        .binding(owner)
        .ok_or_else(|| invalid("stylesWithEffects replacement lost its existing owner binding"))?;
    if desired_graph.target.as_deref() != Some(binding.target.as_str())
        || desired_graph.relationship_id.as_deref() != Some(binding.relationship_id.as_str())
    {
        return Err(invalid(
            "stylesWithEffects replacement graph target differs from its source",
        ));
    }
    let limits = package.read_limits();
    let target_relationships = package.source_relationships_with_limits(&binding.target, limits)?;
    if desired_graph.target_relationships.as_ref() != Some(&target_relationships) {
        return Err(invalid(
            "stylesWithEffects replacement graph target metadata differs from its source",
        ));
    }
    let part = package.get_part(&binding.target)?;
    validate_effects_part_for_package(part, graph.conformance, limits)?;
    if part.blob() != expected_resource.xml_bytes() {
        return Err(invalid(
            "stylesWithEffects replacement source bytes did not round-trip",
        ));
    }
    let graph_token = capture_graph_token(package, &graph, owner)?;
    Ok(Snapshot::from_package(
        owner,
        Some(expected_resource.clone()),
        graph.conformance,
        Some(graph_token),
    ))
}

/// Apply a committed source patch to one package owner.
pub fn apply_commit(package: &mut OpcPackage, owner: Owner, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, owner, &commit.patch)
}

/// Add or replace one owner resource. The operation returns false for an
/// exact source/conformance no-op and therefore leaves signatures untouched.
pub fn put(package: &mut OpcPackage, owner: Owner, resource: Resource) -> Result<bool> {
    let graph = inspect_graph(package)?;
    if resource.conformance != graph.conformance {
        return Err(invalid(
            "stylesWithEffects resource conformance does not match the package",
        ));
    }
    let current = graph
        .binding(owner)
        .map(|binding| {
            Resource::from_part(
                package.get_part(&binding.target)?,
                graph.conformance,
                package.read_limits(),
            )
        })
        .transpose()?;
    if same_resource(current.as_ref(), Some(&resource)) {
        return Ok(false);
    }
    if owner == Owner::Glossary && graph.glossary.is_none() {
        return Err(Error::PartNotFound(
            "glossary document owner is required before publishing stylesWithEffects".to_owned(),
        ));
    }

    let mut candidate = package.clone();
    publish_resource(&mut candidate, owner, Some(resource), None, None)?;
    candidate.validate_signature_edit_from(package)?;
    let published = load(&candidate, owner)?;
    if published.resource().is_none() {
        return Err(invalid("staged stylesWithEffects resource disappeared"));
    }
    *package = candidate;
    Ok(true)
}

/// Remove one owner resource. Missing resources are an exact no-op.
pub fn remove(package: &mut OpcPackage, owner: Owner) -> Result<bool> {
    let graph = inspect_graph(package)?;
    if graph.binding(owner).is_none() {
        return Ok(false);
    }
    let mut candidate = package.clone();
    publish_resource(&mut candidate, owner, None, None, None)?;
    candidate.validate_signature_edit_from(package)?;
    let published = load(&candidate, owner)?;
    if published.resource().is_some() {
        return Err(invalid("staged stylesWithEffects resource was not removed"));
    }
    *package = candidate;
    Ok(true)
}

fn publish_resource(
    package: &mut OpcPackage,
    owner: Owner,
    resource: Option<Resource>,
    desired_graph: Option<&GraphToken>,
    validated_resource: Option<&Resource>,
) -> Result<()> {
    let graph = inspect_graph_with_validated_resource(
        package,
        validated_resource.map(|resource| (owner, resource)),
    )?;
    let source = match owner {
        Owner::MainDocument => graph.main.clone(),
        Owner::Glossary => graph.glossary.clone().ok_or_else(|| {
            Error::PartNotFound(
                "glossary document owner is required before publishing stylesWithEffects"
                    .to_owned(),
            )
        })?,
    };
    let binding = graph.binding(owner).cloned();
    let limits = package.read_limits();
    if let Some(desired_graph) = desired_graph {
        if desired_graph.owner != owner || desired_graph.source.as_deref() != Some(source.as_str())
        {
            return Err(invalid(
                "stylesWithEffects patch graph belongs to a different owner source",
            ));
        }
    }

    match (binding, resource) {
        (Some(binding), Some(resource)) => {
            validate_resource_for_package(&resource, package, graph.conformance)?;
            if let Some(desired) = desired_graph {
                let target_relationships =
                    package.source_relationships_with_limits(&binding.target, limits)?;
                if desired.target.as_deref() != Some(binding.target.as_str())
                    || desired.relationship_id.as_deref() != Some(binding.relationship_id.as_str())
                    || desired.target_relationships.as_ref() != Some(&target_relationships)
                {
                    return Err(invalid(
                        "stylesWithEffects replacement graph target differs from its source",
                    ));
                }
            }
            let relationship_token = package.source_relationships_with_limits(&source, limits)?;
            let relationship_len = relationship_token.bytes().len();
            let (relationship_events, relationship_count) =
                relationship_metrics(relationship_token.bytes())?;
            preflight_package_limits(
                package,
                &source,
                0,
                0,
                Some(binding.target.clone()),
                Some(resource.xml.len()),
                relationship_len,
                relationship_count,
                relationship_events,
                None,
            )?;
            let (content_type, expected) = {
                let part = package.get_part(&binding.target)?;
                (part.content_type().to_owned(), part.blob_arc())
            };
            if content_type != CONTENT_TYPE {
                return Err(Error::ContentType {
                    expected: CONTENT_TYPE.to_owned(),
                    actual: content_type,
                });
            }
            package.try_replace_owned_xml_part_bytes(
                &binding.target,
                expected.as_slice(),
                resource.blob_arc(),
            )?;
        },
        (Some(binding), None) => {
            if has_other_inbound(package, &binding.target, &source, &binding.relationship_id)? {
                return Err(invalid(format!(
                    "stylesWithEffects target '{}' has another inbound relationship",
                    binding.target
                )));
            }

            let current_relationships =
                package.source_relationships_with_limits(&source, limits)?;
            let current_content_types = package.source_content_types_with_limits(limits)?;
            let target_relationships =
                package.source_relationships_with_limits(&binding.target, limits)?;
            let relationship_part_delta = -i64::from(target_relationships.member_present());
            let (replacement_relationships, replacement_content_types) = if let Some(
                desired_graph,
            ) = desired_graph
            {
                if desired_graph.target.is_some() || desired_graph.relationship_id.is_some() {
                    return Err(invalid(
                        "stylesWithEffects removal graph still owns a target",
                    ));
                }
                let target_metadata_matches = match desired_graph.target_relationships.as_ref() {
                    Some(expected) => expected == &target_relationships,
                    None => !target_relationships.member_present(),
                };
                if !target_metadata_matches {
                    return Err(invalid(
                        "stylesWithEffects removal graph target metadata differs from its source",
                    ));
                }
                let relationships = desired_graph.relationships.as_ref().ok_or_else(|| {
                    invalid("stylesWithEffects removal graph lacks relationship provenance")
                })?;
                let content_types = &desired_graph.content_types;
                (relationships.clone(), content_types.clone())
            } else {
                let relationship_plan = current_relationships.plan_edit(
                    &[],
                    &[binding.relationship_id.as_str()],
                    limits,
                )?;
                let content_type_plan = current_content_types.plan_edit(
                    &[],
                    std::slice::from_ref(&binding.target),
                    limits,
                )?;
                let relationship_len = relationship_plan.final_len();
                let relationship_count = relationship_plan.relationship_count();
                let relationship_events = relationship_metrics_after_edit(
                    current_relationships.bytes(),
                    0,
                    &[binding.relationship_id.as_str()],
                )?
                .0;
                let content_type_len = content_type_plan.final_len();
                preflight_package_limits(
                    package,
                    &source,
                    -1,
                    relationship_part_delta,
                    Some(binding.target.clone()),
                    None,
                    relationship_len,
                    relationship_count,
                    relationship_events,
                    None,
                )?;
                check_limit(
                    limits,
                    ReadResource::ContentTypesBytes,
                    content_type_len as u64,
                    limits.max_content_types_bytes() as u64,
                )?;
                (
                    relationship_plan.materialize(limits)?,
                    content_type_plan.materialize(limits)?,
                )
            };
            let (relationship_events, relationship_count) =
                relationship_metrics(replacement_relationships.bytes())?;
            preflight_package_limits(
                package,
                &source,
                -1,
                relationship_part_delta,
                Some(binding.target.clone()),
                None,
                replacement_relationships.bytes().len(),
                relationship_count,
                relationship_events,
                None,
            )?;
            check_limit(
                limits,
                ReadResource::ContentTypesBytes,
                replacement_content_types.bytes().len() as u64,
                limits.max_content_types_bytes() as u64,
            )?;
            package.try_replace_content_types_with_limits(
                current_content_types.bytes(),
                &replacement_content_types,
                limits,
            )?;
            package.try_replace_relationships_with_limits(
                &current_relationships,
                &replacement_relationships,
                limits,
            )?;
            package.remove_part(&binding.target);
        },
        (None, Some(resource)) => {
            validate_resource_for_package(&resource, package, graph.conformance)?;
            let current_relationships =
                package.source_relationships_with_limits(&source, limits)?;
            let current_content_types = package.source_content_types_with_limits(limits)?;
            let (
                target,
                relationship_id,
                replacement_relationships,
                replacement_content_types,
                target_relationships,
            ) = if let Some(desired_graph) = desired_graph {
                let target = desired_graph
                    .target
                    .as_deref()
                    .ok_or_else(|| invalid("stylesWithEffects add graph lacks target"))
                    .and_then(|value| PackURI::new(value.to_owned()).map_err(Error::Uri))?;
                let relationship_id = desired_graph
                    .relationship_id
                    .clone()
                    .ok_or_else(|| invalid("stylesWithEffects add graph lacks relationship ID"))?;
                package.validate_new_part_name(&target)?;
                let relationships = desired_graph.relationships.as_ref().ok_or_else(|| {
                    invalid("stylesWithEffects add graph lacks relationship provenance")
                })?;
                let content_types = desired_graph.content_types.clone();
                let target_relationships = desired_graph.target_relationships.clone();
                (
                    target,
                    relationship_id,
                    relationships.clone(),
                    content_types,
                    target_relationships,
                )
            } else {
                let target = next_part_name(package, owner)?;
                let target_ref = target.relative_ref(source.base_uri());
                let relationship_id = next_relationship_id(package, &source)?;
                let additions = [RelationshipEdit {
                    id: relationship_id.as_str(),
                    reltype: RELATIONSHIP_TYPE,
                    target: target_ref.as_str(),
                    mode: TargetMode::Internal,
                }];
                let relationship_plan = current_relationships.plan_edit(&additions, &[], limits)?;
                let content_type_addition = ContentTypeEdit {
                    part_name: &target,
                    content_type: CONTENT_TYPE,
                };
                let content_type_plan = current_content_types.plan_edit(
                    std::slice::from_ref(&content_type_addition),
                    &[],
                    limits,
                )?;
                let relationship_len = relationship_plan.final_len();
                let relationship_count = relationship_plan.relationship_count();
                let relationship_events =
                    relationship_metrics_after_edit(current_relationships.bytes(), 1, &[])?.0;
                let content_type_len = content_type_plan.final_len();
                preflight_package_limits(
                    package,
                    &source,
                    1,
                    i64::from(!current_relationships.member_present()),
                    None,
                    Some(resource.xml.len()),
                    relationship_len,
                    relationship_count,
                    relationship_events,
                    None,
                )?;
                check_limit(
                    limits,
                    ReadResource::ContentTypesBytes,
                    content_type_len as u64,
                    limits.max_content_types_bytes() as u64,
                )?;
                (
                    target,
                    relationship_id,
                    relationship_plan.materialize(limits)?,
                    content_type_plan.materialize(limits)?,
                    None,
                )
            };
            let (relationship_events, relationship_count) =
                relationship_metrics(replacement_relationships.bytes())?;
            preflight_package_limits(
                package,
                &source,
                1,
                i64::from(replacement_relationships.member_present())
                    - i64::from(current_relationships.member_present()),
                None,
                Some(resource.xml.len()),
                replacement_relationships.bytes().len(),
                relationship_count,
                relationship_events,
                target_relationships.as_ref(),
            )?;
            check_limit(
                limits,
                ReadResource::ContentTypesBytes,
                replacement_content_types.bytes().len() as u64,
                limits.max_content_types_bytes() as u64,
            )?;
            if relationship_id.is_empty() {
                return Err(invalid(
                    "stylesWithEffects add graph has an empty relationship ID",
                ));
            }
            let target_name = target.clone();
            let part = Box::new(BlobPart::new_shared(
                target,
                CONTENT_TYPE.to_owned(),
                resource.blob_arc(),
            ));
            package.try_add_parts_with_source_tokens(
                current_content_types.bytes(),
                &replacement_content_types,
                &current_relationships,
                &replacement_relationships,
                vec![part],
            )?;
            if let Some(target_relationships) = target_relationships {
                let current_target_relationships =
                    package.source_relationships_with_limits(&target_name, limits)?;
                package.try_replace_relationships_with_limits(
                    &current_target_relationships,
                    &target_relationships,
                    limits,
                )?;
            }
        },
        (None, None) => {},
    }
    Ok(())
}

fn validate_resource_for_package(
    resource: &Resource,
    package: &OpcPackage,
    conformance: Conformance,
) -> Result<()> {
    if resource.conformance != conformance {
        return Err(invalid(
            "stylesWithEffects resource conformance does not match the package",
        ));
    }
    validate_xml_with_limits(&resource.xml, Some(conformance), package.read_limits()).map(|_| ())
}

fn validate_effects_part_for_package(
    part: &dyn Part,
    conformance: Conformance,
    limits: ReadLimits,
) -> Result<()> {
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    validate_xml_with_limits(part.blob(), Some(conformance), limits).map(|_| ())
}

fn validate_effects_part_for_graph(
    part: &dyn Part,
    conformance: Conformance,
    limits: ReadLimits,
    validated_resource: Option<&Resource>,
) -> Result<()> {
    if let Some(resource) = validated_resource {
        validate_effects_part_for_package(part, conformance, limits)?;
        if resource.conformance != conformance || part.blob() != resource.xml_bytes() {
            return Err(invalid(
                "stylesWithEffects validated resource does not match its package target",
            ));
        }
        return Ok(());
    }

    // Graph inspection needs the same typed MCE/style-parser refusal as a
    // normal resource load, but it does not need to retain a projection for
    // an unselected owner. Keep the package-level XML and caller-limit check
    // first, then consume the established lazy Styles parser to force all
    // events through its validation path without cloning every Style.
    validate_effects_part_for_package(part, conformance, limits)?;
    #[cfg(test)]
    test_trace::record_graph_validation();
    let mut styles = Styles::from_part(part);
    styles.iter()?.count();
    Ok(())
}

fn preflight_package_limits(
    package: &OpcPackage,
    source: &PackURI,
    part_delta: i64,
    relationship_part_delta: i64,
    old_target: Option<PackURI>,
    new_part_len: Option<usize>,
    replacement_relationship_len: usize,
    replacement_relationship_count: usize,
    replacement_relationship_events: usize,
    new_part_relationships: Option<&OwnedRelationships>,
) -> Result<()> {
    let limits = package.read_limits();
    let current_parts = i64::try_from(package.part_count())
        .map_err(|_| invalid("stylesWithEffects part count overflows i64"))?;
    let final_parts = current_parts
        .checked_add(part_delta)
        .ok_or_else(|| invalid("stylesWithEffects part count overflows"))?;
    if final_parts < 0 {
        return Err(invalid("stylesWithEffects part count became negative"));
    }
    check_limit(
        limits,
        ReadResource::Parts,
        final_parts as u64,
        limits.max_parts() as u64,
    )?;
    let current_graph_nodes = relationship_graph_nodes(package)?;
    let final_graph_nodes = i64::try_from(current_graph_nodes)
        .map_err(|_| invalid("stylesWithEffects graph node count overflows i64"))?
        .checked_add(part_delta)
        .ok_or_else(|| invalid("stylesWithEffects graph node count overflows"))?;
    if final_graph_nodes < 0 {
        return Err(invalid(
            "stylesWithEffects graph node count became negative",
        ));
    }
    check_limit(
        limits,
        ReadResource::RelationshipGraphNodes,
        final_graph_nodes as u64,
        limits.max_relationship_graph_nodes() as u64,
    )?;
    // The aggregate is over actual payload bytes, so every part is decoded
    // here (ADR 0030).
    let mut total_part_bytes = 0u64;
    for part in package.try_iter_parts() {
        let part = part?;
        total_part_bytes = total_part_bytes
            .checked_add(part.blob().len() as u64)
            .ok_or_else(|| invalid("stylesWithEffects part bytes overflow"))?;
    }
    if let Some(old_target) = old_target.as_ref() {
        let old_part = package.get_part(old_target)?;
        total_part_bytes = total_part_bytes
            .checked_sub(old_part.blob().len() as u64)
            .ok_or_else(|| invalid("stylesWithEffects source part bytes underflow"))?;
    }
    if let Some(new_part_len) = new_part_len {
        total_part_bytes = total_part_bytes
            .checked_add(new_part_len as u64)
            .ok_or_else(|| invalid("stylesWithEffects replacement part bytes overflow"))?;
        check_limit(
            limits,
            ReadResource::PartBytes,
            new_part_len as u64,
            limits.max_part_bytes(),
        )?;
    }
    check_limit(
        limits,
        ReadResource::TotalPartBytes,
        total_part_bytes,
        limits.max_total_part_bytes(),
    )?;

    let root = PackURI::new("/").map_err(Error::Uri)?;
    let mut relationship_bytes = 0u64;
    let mut relationship_count = package.rels().iter().count();
    for part in package.iter_parts() {
        relationship_count = relationship_count
            .checked_add(part.rels().iter().count())
            .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?;
    }
    let mut relationship_events = 0usize;
    let root_token = package.source_relationships_with_limits(&root, limits)?;
    let mut relationship_parts = usize::from(root_token.member_present());
    if root_token.member_present() {
        relationship_bytes = relationship_bytes
            .checked_add(root_token.bytes().len() as u64)
            .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflow"))?;
        relationship_events = relationship_events
            .checked_add(relationship_metrics(root_token.bytes())?.0)
            .ok_or_else(|| invalid("stylesWithEffects relationship event count overflow"))?;
    }
    for part in package.iter_parts() {
        let token = package.source_relationships_with_limits(part.partname(), limits)?;
        let removed_part = old_target
            .as_ref()
            .is_some_and(|target| same_uri(target, part.partname()));
        let (length, count, events) = if removed_part {
            (0, 0, 0)
        } else if part.partname() == source {
            (
                replacement_relationship_len,
                replacement_relationship_count,
                replacement_relationship_events,
            )
        } else {
            let (events, count) = relationship_metrics(token.bytes())?;
            (token.bytes().len(), count, events)
        };
        let current_relationship_count = relationship_count
            .checked_sub(part.rels().iter().count())
            .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?;
        relationship_count = if removed_part {
            current_relationship_count
        } else {
            current_relationship_count
                .checked_add(count)
                .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?
        };
        if (!removed_part && token.member_present()) || part.partname() == source {
            relationship_bytes = relationship_bytes
                .checked_add(length as u64)
                .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflow"))?;
            relationship_events = relationship_events
                .checked_add(events)
                .ok_or_else(|| invalid("stylesWithEffects relationship event count overflow"))?;
        }
        if token.member_present() {
            relationship_parts = relationship_parts
                .checked_add(1)
                .ok_or_else(|| invalid("stylesWithEffects relationship part count overflow"))?;
        }
    }
    if let Some(new_part_relationships) = new_part_relationships {
        if new_part_relationships.member_present() {
            let (events, count) = relationship_metrics(new_part_relationships.bytes())?;
            relationship_bytes = relationship_bytes
                .checked_add(new_part_relationships.bytes().len() as u64)
                .ok_or_else(|| invalid("stylesWithEffects relationship bytes overflow"))?;
            relationship_count = relationship_count
                .checked_add(count)
                .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?;
            relationship_events = relationship_events
                .checked_add(events)
                .ok_or_else(|| invalid("stylesWithEffects relationship event count overflow"))?;
        }
    }
    let new_relationship_part_delta =
        new_part_relationships.map_or(0, |relationships| i64::from(relationships.member_present()));
    let final_relationship_parts = i64::try_from(relationship_parts)
        .map_err(|_| invalid("stylesWithEffects relationship part count overflows i64"))?
        .checked_add(relationship_part_delta)
        .and_then(|value| value.checked_add(new_relationship_part_delta))
        .ok_or_else(|| invalid("stylesWithEffects relationship part count overflows"))?;
    if final_relationship_parts < 0 {
        return Err(invalid(
            "stylesWithEffects relationship part count became negative",
        ));
    }
    check_limit(
        limits,
        ReadResource::RelationshipParts,
        final_relationship_parts as u64,
        limits.max_relationship_parts() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationships,
        relationship_count as u64,
        limits.max_total_relationships() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlEvents,
        relationship_events as u64,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    check_limit(
        limits,
        ReadResource::TotalRelationshipXmlBytes,
        relationship_bytes,
        limits.max_total_relationship_xml_bytes() as u64,
    )?;
    Ok(())
}

fn relationship_graph_nodes(package: &OpcPackage) -> Result<usize> {
    let limits = package.read_limits();
    let mut visited = HashSet::<Vec<u8>>::new();
    let mut work_queue = Vec::new();
    for relationship in package.rels().iter().filter(|value| !value.is_external()) {
        let target = relationship.target_partname()?;
        enqueue_relationship_graph_target(target, &mut visited, &mut work_queue, limits)?;
    }
    while let Some(owner) = work_queue.pop() {
        let Ok(part) = package.get_part(&owner) else {
            continue;
        };
        for relationship in part.rels().iter().filter(|value| !value.is_external()) {
            let target = relationship.target_partname()?;
            enqueue_relationship_graph_target(target, &mut visited, &mut work_queue, limits)?;
        }
    }
    Ok(visited.len())
}

fn enqueue_relationship_graph_target(
    target: PackURI,
    visited: &mut HashSet<Vec<u8>>,
    work_queue: &mut Vec<PackURI>,
    limits: ReadLimits,
) -> Result<()> {
    let key = canonical_uri_key(&target)?;
    if visited.contains(&key) {
        return Ok(());
    }
    let next = visited
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("stylesWithEffects relationship graph node count overflows"))?;
    check_limit(
        limits,
        ReadResource::RelationshipGraphNodes,
        next as u64,
        limits.max_relationship_graph_nodes() as u64,
    )?;
    visited.try_reserve(1).map_err(|source| {
        Error::Opc(litchi_opc::OpcError::Allocation {
            resource: "stylesWithEffects relationship graph nodes",
            source,
        })
    })?;
    work_queue.try_reserve(1).map_err(|source| {
        Error::Opc(litchi_opc::OpcError::Allocation {
            resource: "stylesWithEffects relationship graph queue",
            source,
        })
    })?;
    visited.insert(key);
    work_queue.push(target);
    Ok(())
}

fn canonical_uri_key(target: &PackURI) -> Result<Vec<u8>> {
    let mut key = Vec::new();
    key.try_reserve_exact(target.as_str().len())
        .map_err(|source| {
            Error::Opc(litchi_opc::OpcError::Allocation {
                resource: "stylesWithEffects relationship graph keys",
                source,
            })
        })?;
    key.extend_from_slice(target.as_str().as_bytes());
    key.make_ascii_lowercase();
    Ok(key)
}

fn relationship_metrics(xml: &[u8]) -> Result<(usize, usize)> {
    relationship_metrics_after_edit(xml, 0, &[])
}

/// Compute the metrics of a source relationship member after removing the
/// selected direct relationship elements and appending authored relationship
/// elements.  The scan uses the same trimmed quick-xml event view as the
/// package aggregate metric, while avoiding allocation of the final XML
/// buffer.  Relationship edit plans intentionally retain lexical whitespace,
/// so their internal scan count (which includes whitespace events) cannot be
/// used for the package-wide trimmed event limit.
fn relationship_metrics_after_edit(
    xml: &[u8],
    additions: usize,
    removals: &[&str],
) -> Result<(usize, usize)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut events = 0usize;
    let mut relationships = 0usize;
    let mut depth = 0usize;
    let mut skipped_relationship_depth = None;
    let mut root_empty = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        if let Some(skip_depth) = skipped_relationship_depth {
            match event {
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("stylesWithEffects relationship depth overflow"))?;
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("stylesWithEffects relationship depth underflow"))?;
                    if depth < skip_depth {
                        skipped_relationship_depth = None;
                    }
                },
                Event::Eof => {
                    return Err(invalid(
                        "stylesWithEffects relationship removal reached EOF before close",
                    ));
                },
                _ => {},
            }
            continue;
        }
        match event {
            Event::Start(element) => {
                if depth == 1 && local_name(element.name().as_ref()) == b"Relationship" {
                    let id = relationship_element_id(&element, reader.decoder())?;
                    if removals.iter().any(|candidate| *candidate == id) {
                        skipped_relationship_depth = Some(depth + 1);
                    } else {
                        relationships = relationships.checked_add(1).ok_or_else(|| {
                            invalid("stylesWithEffects relationship count overflow")
                        })?;
                        events = events.checked_add(1).ok_or_else(|| {
                            invalid("stylesWithEffects relationship event count overflow")
                        })?;
                    }
                } else {
                    events = events.checked_add(1).ok_or_else(|| {
                        invalid("stylesWithEffects relationship event count overflow")
                    })?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("stylesWithEffects relationship depth overflow"))?;
            },
            Event::Empty(element) => {
                let is_root = depth == 0 && local_name(element.name().as_ref()) == b"Relationships";
                let is_relationship =
                    depth == 1 && local_name(element.name().as_ref()) == b"Relationship";
                let removed = if is_relationship {
                    let id = relationship_element_id(&element, reader.decoder())?;
                    if removals.iter().any(|candidate| *candidate == id) {
                        true
                    } else {
                        relationships = relationships.checked_add(1).ok_or_else(|| {
                            invalid("stylesWithEffects relationship count overflow")
                        })?;
                        false
                    }
                } else {
                    false
                };
                if is_root {
                    root_empty = true;
                }
                if !removed {
                    events = events.checked_add(1).ok_or_else(|| {
                        invalid("stylesWithEffects relationship event count overflow")
                    })?;
                }
            },
            Event::End(_) => {
                events = events.checked_add(1).ok_or_else(|| {
                    invalid("stylesWithEffects relationship event count overflow")
                })?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("stylesWithEffects relationship depth underflow"))?;
            },
            Event::Eof => {
                events = events.checked_add(1).ok_or_else(|| {
                    invalid("stylesWithEffects relationship event count overflow")
                })?;
                if depth != 0 {
                    return Err(invalid("stylesWithEffects relationship root is unclosed"));
                }
                let root_expansion = usize::from(root_empty && additions != 0);
                let events = events
                    .checked_add(additions)
                    .and_then(|count| count.checked_add(root_expansion))
                    .ok_or_else(|| {
                        invalid("stylesWithEffects relationship event count overflow")
                    })?;
                let relationships = relationships
                    .checked_add(additions)
                    .ok_or_else(|| invalid("stylesWithEffects relationship count overflow"))?;
                return Ok((events, relationships));
            },
            _ => {
                events = events.checked_add(1).ok_or_else(|| {
                    invalid("stylesWithEffects relationship event count overflow")
                })?;
            },
        }
    }
}

fn relationship_element_id(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<String> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if attribute.key.as_ref() == b"Id" {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| value.into_owned())
                .map_err(|error| invalid(error.to_string()));
        }
    }
    Err(invalid("relationship is missing Id"))
}

fn next_relationship_id(package: &OpcPackage, source: &PackURI) -> Result<String> {
    let relationships = package.get_part(source)?.rels();
    for index in 0..MAX_PART_NAME_ATTEMPTS {
        let candidate = if index == 0 {
            "rIdEffects".to_owned()
        } else {
            format!("rIdEffects{index}")
        };
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "no bounded stylesWithEffects relationship ID is available",
    ))
}

fn inspect_graph(package: &OpcPackage) -> Result<GraphState> {
    inspect_graph_with_validated_resource(package, None)
}

fn inspect_graph_with_validated_resource(
    package: &OpcPackage,
    validated: Option<(Owner, &Resource)>,
) -> Result<GraphState> {
    let conformance = package_conformance(package)?;
    let main_part = package.main_document_part()?;
    if !matches!(
        main_part.content_type(),
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid("main document is not a WordprocessingML document"));
    }
    let main = main_part.partname().clone();

    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "package root cannot own a stylesWithEffects relationship",
        ));
    }

    let glossary = locate_glossary(package, &main, conformance)?;
    let mut bindings: [Option<Binding>; 2] = [None, None];
    let mut effects_targets = Vec::<PackURI>::new();

    for part in package.iter_parts() {
        let owner = if same_uri(part.partname(), &main) {
            Some(Owner::MainDocument)
        } else if glossary
            .as_ref()
            .is_some_and(|value| same_uri(part.partname(), value))
        {
            Some(Owner::Glossary)
        } else {
            None
        };
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
        {
            let owner = owner.ok_or_else(|| {
                invalid(format!(
                    "stylesWithEffects relationship has invalid source '{}'",
                    part.partname()
                ))
            })?;
            if relationship.is_external() {
                return Err(invalid("stylesWithEffects relationship cannot be external"));
            }
            if bindings[owner.index()].is_some() {
                return Err(invalid(format!(
                    "{owner} has multiple stylesWithEffects relationships"
                )));
            }
            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?.partname().clone();
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != CONTENT_TYPE {
                return Err(Error::ContentType {
                    expected: CONTENT_TYPE.to_owned(),
                    actual: target_part.content_type().to_owned(),
                });
            }
            if effects_targets
                .iter()
                .any(|existing| same_uri(existing, &target))
            {
                return Err(invalid(
                    "multiple stylesWithEffects owners target the same part",
                ));
            }
            if target_part.rels().iter().next().is_some() {
                return Err(invalid(
                    "stylesWithEffects part must be a relationship leaf",
                ));
            }
            effects_targets.push(target.clone());
            bindings[owner.index()] = Some(Binding {
                source: part.partname().clone(),
                target,
                relationship_id: relationship.r_id().to_owned(),
            });
        }
    }

    if effects_targets.len() > MAX_EFFECTS_PARTS {
        return Err(invalid(
            "package contains more than two stylesWithEffects parts",
        ));
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
    {
        if !effects_targets
            .iter()
            .any(|target| same_uri(target, part.partname()))
        {
            return Err(invalid(format!(
                "orphan stylesWithEffects part '{}' has no typed owner",
                part.partname()
            )));
        }
    }
    for target in &effects_targets {
        let expected = bindings
            .iter()
            .flatten()
            .find(|binding| same_uri(&binding.target, target));
        let Some(expected) = expected else {
            return Err(invalid(format!(
                "stylesWithEffects target '{}' has no owner",
                target
            )));
        };
        validate_exclusive_inbound(package, target, expected)?;
        let part = package.get_part(target)?;
        let validated_resource = validated.and_then(|(owner, resource)| {
            bindings[owner.index()]
                .as_ref()
                .filter(|binding| same_uri(&binding.target, target))
                .map(|_| resource)
        });
        validate_effects_part_for_graph(
            part,
            conformance,
            package.read_limits(),
            validated_resource,
        )?;
    }

    Ok(GraphState {
        conformance,
        main,
        glossary,
        bindings,
    })
}

fn locate_glossary(
    package: &OpcPackage,
    main: &PackURI,
    conformance: Conformance,
) -> Result<Option<PackURI>> {
    let mut found: Option<PackURI> = None;
    for part in package.iter_parts() {
        for relationship in part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                GLOSSARY_RELATIONSHIP_TYPE | STRICT_GLOSSARY_RELATIONSHIP_TYPE
            )
        }) {
            if !same_uri(part.partname(), main) {
                return Err(invalid(format!(
                    "glossary relationship has invalid source '{}'",
                    part.partname()
                )));
            }
            let expected_relationship = match conformance {
                Conformance::Transitional => GLOSSARY_RELATIONSHIP_TYPE,
                Conformance::Strict => STRICT_GLOSSARY_RELATIONSHIP_TYPE,
            };
            if relationship.reltype() != expected_relationship {
                return Err(invalid(
                    "glossary relationship does not match package conformance",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("glossary relationship cannot be external"));
            }
            if found.is_some() {
                return Err(invalid("main document has multiple glossary relationships"));
            }
            let target = relationship.target_partname()?;
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != ct::WML_DOCUMENT_GLOSSARY {
                return Err(Error::ContentType {
                    expected: ct::WML_DOCUMENT_GLOSSARY.to_owned(),
                    actual: target_part.content_type().to_owned(),
                });
            }
            found = Some(target_part.partname().clone());
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == ct::WML_DOCUMENT_GLOSSARY)
    {
        if found
            .as_ref()
            .is_none_or(|target| !same_uri(target, part.partname()))
        {
            return Err(invalid(format!(
                "orphan glossary document part '{}' has no main owner",
                part.partname()
            )));
        }
    }
    Ok(found)
}

fn package_conformance(package: &OpcPackage) -> Result<Conformance> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(if relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT {
        Conformance::Strict
    } else {
        Conformance::Transitional
    })
}

fn validate_exclusive_inbound(
    package: &OpcPackage,
    target: &PackURI,
    expected: &Binding,
) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if same_uri(&relationship.target_partname()?, target) {
            return Err(invalid(format!(
                "stylesWithEffects part '{}' has an inbound package-root relationship",
                target
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if same_uri(&relationship.target_partname()?, target)
                && !(same_uri(source.partname(), &expected.source)
                    && relationship.r_id() == expected.relationship_id)
            {
                return Err(invalid(format!(
                    "stylesWithEffects part '{}' has another inbound relationship",
                    target
                )));
            }
        }
    }
    Ok(())
}

fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    source: &PackURI,
    relationship_id: &str,
) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if same_uri(&relationship.target_partname()?, target) {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if same_uri(&relationship.target_partname()?, target)
                && !(same_uri(part.partname(), source) && relationship.r_id() == relationship_id)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn next_part_name(package: &OpcPackage, owner: Owner) -> Result<PackURI> {
    let stem = match owner {
        Owner::MainDocument => "/word/stylesWithEffects.xml",
        Owner::Glossary => "/word/glossary/stylesWithEffects.xml",
    };
    let (prefix, _) = stem
        .rsplit_once(".xml")
        .ok_or_else(|| invalid("invalid stylesWithEffects target stem"))?;
    for index in 0..MAX_PART_NAME_ATTEMPTS {
        let value = if index == 0 {
            stem.to_owned()
        } else {
            format!("{prefix}{index}.xml")
        };
        let candidate = PackURI::new(value).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "no bounded stylesWithEffects package target is available",
    ))
}

fn same_resource(left: Option<&Resource>, right: Option<&Resource>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left == right,
        _ => false,
    }
}

fn same_uri(left: &PackURI, right: &PackURI) -> bool {
    left.as_str().eq_ignore_ascii_case(right.as_str())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn check_limit(
    _limits: ReadLimits,
    resource: ReadResource,
    actual: u64,
    maximum: u64,
) -> Result<()> {
    if actual > maximum {
        return Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
            resource,
            actual,
            maximum,
        }));
    }
    Ok(())
}

fn validate_xml(xml: &[u8], expected: Option<Conformance>) -> Result<Conformance> {
    let bounded_limits = ReadLimits::builder()
        .max_part_bytes(
            u64::try_from(MAX_XML_BYTES)
                .map_err(|_| invalid("stylesWithEffects XML byte bound overflows u64"))?,
        )?
        .max_xml_events(MAX_XML_EVENTS)?
        .max_xml_depth(MAX_XML_DEPTH)?
        .build()?;
    validate_xml_with_limits(xml, expected, bounded_limits)
}

fn validate_xml_with_limits(
    xml: &[u8],
    expected: Option<Conformance>,
    limits: ReadLimits,
) -> Result<Conformance> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "stylesWithEffects XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    check_limit(
        limits,
        ReadResource::PartBytes,
        xml.len() as u64,
        limits.max_part_bytes(),
    )?;
    let part_name = PackURI::new("/word/stylesWithEffects.xml")
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    validate_source_xml_bytes(&part_name, xml, limits)?;
    let mut reader = Reader::from_reader(xml);
    let mut root_conformance = None;
    let mut depth = 0usize;
    let mut root_closed = false;
    let mut events = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("stylesWithEffects XML event count overflow"))?;
        if events > MAX_XML_EVENTS {
            return Err(invalid(format!(
                "stylesWithEffects XML exceeds {MAX_XML_EVENTS} events"
            )));
        }
        check_limit(
            limits,
            ReadResource::XmlEvents,
            events as u64,
            limits.max_xml_events() as u64,
        )?;
        match event {
            Event::Start(start) => {
                if root_closed {
                    return Err(invalid("stylesWithEffects XML has multiple roots"));
                }
                if root_conformance.is_none() {
                    let name = start.name();
                    let local = local_name(name.as_ref());
                    if local != b"styles" {
                        return Err(invalid("stylesWithEffects XML root must be styles"));
                    }
                    let namespace = root_namespace(&start)?;
                    let conformance = Conformance::from_namespace(&namespace).ok_or_else(|| {
                        invalid("stylesWithEffects XML uses an unknown namespace")
                    })?;
                    if expected.is_some_and(|value| value != conformance) {
                        return Err(invalid(
                            "stylesWithEffects XML namespace does not match package conformance",
                        ));
                    }
                    root_conformance = Some(conformance);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("stylesWithEffects XML depth overflow"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "stylesWithEffects XML exceeds {MAX_XML_DEPTH} nesting levels"
                    )));
                }
                check_limit(
                    limits,
                    ReadResource::XmlDepth,
                    depth as u64,
                    limits.max_xml_depth() as u64,
                )?;
            },
            Event::Empty(empty) => {
                if root_closed {
                    return Err(invalid("stylesWithEffects XML has multiple roots"));
                }
                if root_conformance.is_none() {
                    let name = empty.name();
                    let local = local_name(name.as_ref());
                    if local != b"styles" {
                        return Err(invalid("stylesWithEffects XML root must be styles"));
                    }
                    let namespace = root_namespace(&empty)?;
                    let conformance = Conformance::from_namespace(&namespace).ok_or_else(|| {
                        invalid("stylesWithEffects XML uses an unknown namespace")
                    })?;
                    if expected.is_some_and(|value| value != conformance) {
                        return Err(invalid(
                            "stylesWithEffects XML namespace does not match package conformance",
                        ));
                    }
                    root_conformance = Some(conformance);
                    root_closed = true;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("stylesWithEffects XML has an unmatched end tag"));
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if depth == 0
                    && !text.is_empty()
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "stylesWithEffects XML has non-whitespace content outside its root",
                    ));
                }
            },
            Event::CData(data) => {
                if depth == 0 && !data.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("stylesWithEffects XML has CDATA outside its root"));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !root_closed {
        return Err(invalid("stylesWithEffects XML root is incomplete"));
    }
    root_conformance.ok_or_else(|| invalid("stylesWithEffects XML has no root"))
}

fn local_name(value: &[u8]) -> &[u8] {
    value.rsplit(|byte| *byte == b':').next().unwrap_or(value)
}

fn root_namespace(element: &quick_xml::events::BytesStart<'_>) -> Result<String> {
    let name = element.name();
    let name = name.as_ref();
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |index| &name[..index]);
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        let matches = if prefix.is_empty() {
            key == b"xmlns"
        } else {
            key == [b"xmlns:".as_slice(), prefix].concat().as_slice()
        };
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if matches {
            return Ok(value);
        }
    }
    Err(invalid("stylesWithEffects XML root has no Word namespace"))
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    const FIXTURE: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/document/Bug54849.docx"
    ));

    fn changed_resource(snapshot: &Snapshot) -> Resource {
        let mut xml = snapshot
            .resource()
            .expect("effects fixture has a main resource")
            .xml_bytes()
            .to_vec();
        let closing = b"</w:styles>";
        let offset = xml
            .windows(closing.len())
            .position(|window| window == closing)
            .expect("effects fixture has a styles root");
        xml.splice(
            offset..offset,
            br#"<w:style w:type="paragraph" w:styleId="EffectsCountProbe"><w:name w:val="Effects Count Probe"/></w:style>"#.iter().copied(),
        );
        Resource::from_xml(xml).expect("changed effects resource")
    }

    fn changed_patch(snapshot: &Snapshot) -> (Patch, Resource) {
        let resource = changed_resource(snapshot);
        let mut edit = snapshot.edit();
        edit.replace_resource(Some(resource.clone()))
            .expect("stage changed effects resource");
        (
            edit.commit()
                .expect("commit changed effects resource")
                .into_patch(),
            resource,
        )
    }

    #[test]
    fn staged_facade_reuses_snapshot_and_skips_nested_effects_clone() {
        let mut package = crate::Package::from_reader(Cursor::new(FIXTURE.to_vec()))
            .expect("open effects fixture");
        let snapshot = package
            .styles_with_effects(Owner::MainDocument)
            .expect("load effects fixture");
        let (patch, expected) = changed_patch(&snapshot);

        test_trace::reset();
        let applied = package
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .expect("apply staged effects patch");
        let counts = test_trace::take();

        assert_eq!(counts.loads, 1);
        assert_eq!(counts.candidate_clones, 0);
        assert_eq!(counts.projection_builds, 1);
        assert_eq!(counts.readbacks, 1);
        assert_eq!(counts.graph_validations, 4);
        assert_eq!(
            applied
                .resource()
                .expect("published effects resource")
                .xml_bytes(),
            expected.xml_bytes()
        );
    }

    #[test]
    fn direct_patch_retains_nested_effects_candidate_clone() {
        let mut package = OpcPackage::from_vec(FIXTURE.to_vec()).expect("open effects fixture");
        let snapshot = load(&package, Owner::MainDocument).expect("load effects fixture");
        let (patch, expected) = changed_patch(&snapshot);

        test_trace::reset();
        let applied = patch
            .apply(&mut package)
            .expect("apply direct effects patch");
        let counts = test_trace::take();

        assert_eq!(counts.loads, 1);
        assert_eq!(counts.candidate_clones, 1);
        assert_eq!(counts.projection_builds, 1);
        assert_eq!(counts.readbacks, 1);
        assert_eq!(counts.graph_validations, 4);
        assert_eq!(
            applied
                .resource()
                .expect("published effects resource")
                .xml_bytes(),
            expected.xml_bytes()
        );
    }

    #[test]
    fn staged_patch_rejects_unbound_standalone_snapshot() {
        let current = Snapshot::from_xml(
            Owner::MainDocument,
            br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
                .to_vec(),
        )
        .expect("standalone effects snapshot");
        let replacement = Resource::from_xml(
            br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="probe"/></w:styles>"#
                .to_vec(),
        )
        .expect("standalone replacement resource");
        let mut edit = current.edit();
        edit.replace_resource(Some(replacement))
            .expect("stage standalone replacement");
        let patch = edit
            .commit()
            .expect("commit standalone replacement")
            .into_patch();
        let mut package = OpcPackage::new();

        let error = apply_patch_staged(&mut package, Owner::MainDocument, &patch, current)
            .expect_err("unbound staged patch must be rejected");
        assert!(matches!(
            error,
            Error::Invalid(message) if message.contains("package-bound current snapshot")
        ));
        assert_eq!(package.part_count(), 0);
    }
}
