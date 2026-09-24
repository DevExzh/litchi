//! Semantic values for slide-owned PowerPoint InkAction content parts.

use std::sync::Arc;

use litchi_drawingml::ink::actions::Profile;
use litchi_opc::{OwnedContentTypes, OwnedRelationships, PackURI, TargetMode};

use super::codec;
use crate::{Error, Result};

const MAX_ANCHORS: usize = 4096;
const MAX_TARGET_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_TARGET_BYTES: usize = 256 * 1024 * 1024;
const MAX_TARGET_RELATIONSHIPS: usize = 4096;

/// Bounded resource policy for one slide-owned InkAction graph.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Limits {
    /// Maximum action anchors retained from one slide.
    pub anchors: usize,
    /// Maximum bytes retained for one action target.
    pub target_bytes: usize,
    /// Maximum aggregate bytes retained for unique action targets.
    pub total_target_bytes: usize,
    /// Maximum inbound/outbound relationship records retained per target.
    pub target_relationships: usize,
}

impl Limits {
    /// Conservative bounded limits for ordinary PPTX slides.
    pub const DEFAULT: Self = Self {
        anchors: MAX_ANCHORS,
        target_bytes: MAX_TARGET_BYTES,
        total_target_bytes: MAX_TOTAL_TARGET_BYTES,
        target_relationships: MAX_TARGET_RELATIONSHIPS,
    };

    /// Construct finite limits within this owner's hard ceilings.
    #[must_use]
    pub const fn new(
        anchors: usize,
        target_bytes: usize,
        total_target_bytes: usize,
        target_relationships: usize,
    ) -> Option<Self> {
        if anchors == 0
            || anchors > MAX_ANCHORS
            || target_bytes == 0
            || target_bytes > MAX_TARGET_BYTES
            || total_target_bytes == 0
            || total_target_bytes > MAX_TOTAL_TARGET_BYTES
            || target_relationships == 0
            || target_relationships > MAX_TARGET_RELATIONSHIPS
        {
            None
        } else {
            Some(Self {
                anchors,
                target_bytes,
                total_target_bytes,
                target_relationships,
            })
        }
    }

    /// Check a caller-supplied value before any package scan or allocation.
    pub(crate) fn validate(self) -> Result<()> {
        if self.anchors == 0 || self.anchors > MAX_ANCHORS {
            return Err(Error::Limit {
                resource: "ink-action anchor count",
                limit: MAX_ANCHORS,
            });
        }
        if self.target_bytes == 0 || self.target_bytes > MAX_TARGET_BYTES {
            return Err(Error::Limit {
                resource: "ink-action target bytes",
                limit: MAX_TARGET_BYTES,
            });
        }
        if self.total_target_bytes == 0 || self.total_target_bytes > MAX_TOTAL_TARGET_BYTES {
            return Err(Error::Limit {
                resource: "ink-action aggregate target bytes",
                limit: MAX_TOTAL_TARGET_BYTES,
            });
        }
        if self.target_relationships == 0 || self.target_relationships > MAX_TARGET_RELATIONSHIPS {
            return Err(Error::Limit {
                resource: "ink-action relationship edges",
                limit: MAX_TARGET_RELATIONSHIPS,
            });
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Package conformance dialect resolved from the owning slide root.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum Dialect {
    Transitional,
    Strict,
}

/// A slide selector is a zero-based position in the immutable presentation
/// snapshot.  It is deliberately not a relationship ID or a package URI.
pub type SlideSelector = usize;

/// Which source order is used to resolve an action anchor.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum AnchorSelector {
    /// Active, contract-valid anchors in depth-first document order.
    Semantic {
        /// Owning slide position.
        owner: SlideSelector,
        /// Zero-based semantic anchor ordinal.
        ordinal: usize,
        /// Source context used for stale/ambiguous checks.
        context: AnchorFingerprint,
    },
    /// Raw action MCE anchor order in the owner source.
    Source {
        /// Owning slide position.
        owner: SlideSelector,
        /// Zero-based source anchor ordinal.
        ordinal: usize,
        /// Source context used for stale/ambiguous checks.
        context: AnchorFingerprint,
    },
}

impl AnchorSelector {
    /// Construct a semantic selector.
    #[must_use]
    pub const fn semantic(
        owner: SlideSelector,
        ordinal: usize,
        context: AnchorFingerprint,
    ) -> Self {
        Self::Semantic {
            owner,
            ordinal,
            context,
        }
    }

    /// Construct a source selector.
    #[must_use]
    pub const fn source(owner: SlideSelector, ordinal: usize, context: AnchorFingerprint) -> Self {
        Self::Source {
            owner,
            ordinal,
            context,
        }
    }

    pub(crate) const fn owner(self) -> SlideSelector {
        match self {
            Self::Semantic { owner, .. } | Self::Source { owner, .. } => owner,
        }
    }

    pub(crate) const fn context(self) -> AnchorFingerprint {
        match self {
            Self::Semantic { context, .. } | Self::Source { context, .. } => context,
        }
    }
}

/// Stable source fingerprint for one selected action anchor.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct AnchorFingerprint(u64);

impl AnchorFingerprint {
    pub(crate) fn from_closure(
        owner: &[u8],
        branch: &[u8],
        relationship_id: &[u8],
        relationship_type: &[u8],
        target_ref: &[u8],
        target_part: &[u8],
        target: &[u8],
    ) -> Self {
        let mut value = 0xcbf29ce484222325u64;
        for bytes in [
            owner,
            branch,
            relationship_id,
            relationship_type,
            target_ref,
            target_part,
            target,
        ] {
            for byte in bytes {
                value ^= u64::from(*byte);
                value = value.wrapping_mul(0x100000001b3);
            }
            value ^= 0xff;
            value = value.wrapping_mul(0x100000001b3);
        }
        Self(value)
    }

    pub(crate) fn with_graph(
        self,
        content_type: &[u8],
        target_mode: TargetMode,
        inbound: &[InboundReference],
        outbound: &[InboundReference],
    ) -> Self {
        let mut value = self.0;
        for bytes in [content_type] {
            for byte in bytes {
                value ^= u64::from(*byte);
                value = value.wrapping_mul(0x100000001b3);
            }
            value ^= 0xff;
            value = value.wrapping_mul(0x100000001b3);
        }
        value ^= target_mode as u8 as u64;
        value = value.wrapping_mul(0x100000001b3);
        for reference in inbound.iter().chain(outbound.iter()) {
            for bytes in [
                reference.source_part.as_str().as_bytes(),
                reference.relationship_id.as_bytes(),
                reference.relationship_type.as_bytes(),
                reference.target_ref.as_bytes(),
            ] {
                for byte in bytes {
                    value ^= u64::from(*byte);
                    value = value.wrapping_mul(0x100000001b3);
                }
                value ^= 0xff;
                value = value.wrapping_mul(0x100000001b3);
            }
            value ^= reference.target_mode as u8 as u64;
            value = value.wrapping_mul(0x100000001b3);
        }
        Self(value)
    }

    /// Return the compact fingerprint value for diagnostics and persistence.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// Selected MCE branch kind retained for diagnostics.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum Branch {
    /// Active `mc:Choice` branch.
    Choice,
    /// Inactive `mc:Fallback` branch.
    Fallback,
}

/// Internal raw candidate produced by the owner scanner.
#[derive(Clone, Debug)]
pub(crate) struct Candidate {
    pub(crate) source_ordinal: usize,
    pub(crate) span: std::ops::Range<usize>,
    pub(crate) choice_span: std::ops::Range<usize>,
    pub(crate) fallback_span: std::ops::Range<usize>,
    pub(crate) content_span: std::ops::Range<usize>,
    pub(crate) relationship_id: Option<String>,
    pub(crate) fingerprint: AnchorFingerprint,
    pub(crate) typed: bool,
    pub(crate) dialect: Dialect,
}

/// A raw action-capable MCE span retained for source selectors, including a
/// branch that is currently unsupported by the typed profile.
///
/// A resolvable internal target contributes its complete relationship closure
/// to the fingerprint. Unresolved, malformed, or query/fragment targets stay
/// opaque and are never edit authority; changes confined to such an untouched
/// target are therefore irrelevant to publication, while owner XML and owner
/// relationship changes remain stale-source inputs.
#[derive(Clone, Debug)]
pub struct RawAnchor {
    pub(crate) source_ordinal: usize,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) span: std::ops::Range<usize>,
    pub(crate) fingerprint: AnchorFingerprint,
}

impl RawAnchor {
    #[must_use]
    pub const fn source_ordinal(&self) -> usize {
        self.source_ordinal
    }

    #[must_use]
    pub const fn fingerprint(&self) -> AnchorFingerprint {
        self.fingerprint
    }

    /// Exact source bytes of this raw action-capable MCE span.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        self.source_xml.get(self.span.clone()).unwrap_or_default()
    }
}

/// One resolved inbound relationship to an action target.  Physical IDs are
/// retained for diagnostics and stale authorization, never used as semantic
/// selectors.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct InboundReference {
    pub(crate) source_part: PackURI,
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

/// Exact package-root relationship record retained by a source snapshot.
///
/// Package relationships have no owning part URI, so they are kept in a
/// separate read-set type rather than being folded into [`InboundReference`].
/// The complete record is compared during patch publication; the compact
/// snapshot revision remains diagnostic only.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct PackageRelationship {
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

impl InboundReference {
    #[must_use]
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part
    }

    /// Alias exposing the physical source part in graph diagnostics.
    #[must_use]
    pub fn source_part(&self) -> &PackURI {
        self.source_part_name()
    }

    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    #[must_use]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }
}

/// A classified slide-owned action anchor and its complete package closure.
#[derive(Clone, Debug)]
pub struct ActionAnchor {
    pub(crate) slide_index: usize,
    pub(crate) semantic_ordinal: usize,
    pub(crate) source_ordinal: usize,
    pub(crate) fingerprint: AnchorFingerprint,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) owner_span: std::ops::Range<usize>,
    pub(crate) choice_span: std::ops::Range<usize>,
    pub(crate) fallback_span: std::ops::Range<usize>,
    pub(crate) anchor_span: std::ops::Range<usize>,
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
    pub(crate) target_part_name: PackURI,
    pub(crate) content_type: String,
    pub(crate) target_bytes: Arc<[u8]>,
    pub(crate) profile: Arc<Profile>,
    pub(crate) inbound: Arc<[InboundReference]>,
    pub(crate) outbound: Arc<[InboundReference]>,
}

impl ActionAnchor {
    #[must_use]
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    #[must_use]
    pub const fn semantic_ordinal(&self) -> usize {
        self.semantic_ordinal
    }

    #[must_use]
    pub const fn source_ordinal(&self) -> usize {
        self.source_ordinal
    }

    #[must_use]
    pub const fn fingerprint(&self) -> AnchorFingerprint {
        self.fingerprint
    }

    /// Exact enclosing `mc:AlternateContent` bytes.
    #[must_use]
    pub fn owner_xml(&self) -> &[u8] {
        self.source_xml
            .get(self.owner_span.clone())
            .unwrap_or_default()
    }

    /// Exact selected `mc:Choice` bytes.
    #[must_use]
    pub fn choice_xml(&self) -> &[u8] {
        self.source_xml
            .get(self.choice_span.clone())
            .unwrap_or_default()
    }

    /// Exact inactive `mc:Fallback` bytes.
    #[must_use]
    pub fn fallback_xml(&self) -> &[u8] {
        self.source_xml
            .get(self.fallback_span.clone())
            .unwrap_or_default()
    }

    /// Exact selected `p:contentPart` bytes.
    #[must_use]
    pub fn anchor_xml(&self) -> &[u8] {
        self.source_xml
            .get(self.anchor_span.clone())
            .unwrap_or_default()
    }

    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    #[must_use]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    #[must_use]
    pub fn target_part_name(&self) -> &PackURI {
        &self.target_part_name
    }

    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact action target bytes retained by the immutable snapshot.
    #[must_use]
    pub fn target_bytes(&self) -> &[u8] {
        &self.target_bytes
    }

    /// Shared strict action profile parsed only after package closure checks.
    #[must_use = "inspect the shared action profile"]
    pub fn profile(&self) -> &Profile {
        &self.profile
    }

    /// All inbound package edges to the resolved target.
    #[must_use]
    pub fn inbound_references(&self) -> &[InboundReference] {
        &self.inbound
    }

    /// Outbound edges retained on the target part for diagnostics.
    #[must_use]
    pub fn outbound_references(&self) -> &[InboundReference] {
        &self.outbound
    }

    pub(crate) fn same_value(&self, other: &Self) -> bool {
        self.slide_index == other.slide_index
            && self.semantic_ordinal == other.semantic_ordinal
            && self.source_ordinal == other.source_ordinal
            && self.fingerprint == other.fingerprint
            && self.owner_span == other.owner_span
            && self.choice_span == other.choice_span
            && self.fallback_span == other.fallback_span
            && self.anchor_span == other.anchor_span
            && self.relationship_id == other.relationship_id
            && self.relationship_type == other.relationship_type
            && self.target_ref == other.target_ref
            && self.target_mode == other.target_mode
            && self.target_part_name == other.target_part_name
            && self.content_type == other.content_type
            && self.target_bytes == other.target_bytes
            && self.profile == other.profile
            && self.inbound == other.inbound
            && self.outbound == other.outbound
    }
}

/// An immutable source-backed view of one slide's InkAction graph.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(crate) slide_index: usize,
    pub(crate) slide_part_name: PackURI,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) anchors: Arc<[ActionAnchor]>,
    pub(crate) raw_anchors: Arc<[RawAnchor]>,
    pub(crate) owner_relationships: Arc<[InboundReference]>,
    pub(crate) content_types_source: OwnedContentTypes,
    pub(crate) owner_relationship_source: OwnedRelationships,
    pub(crate) package_relationship_source: OwnedRelationships,
    pub(crate) package_relationships: Arc<[PackageRelationship]>,
    pub(crate) read_limits: litchi_opc::ReadLimits,
    pub(crate) revision: u64,
    pub(crate) limits: Limits,
}

impl Snapshot {
    #[must_use]
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    #[must_use]
    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        &self.source_xml
    }

    #[must_use]
    pub fn anchors(&self) -> &[ActionAnchor] {
        &self.anchors
    }

    /// Alias matching the package vocabulary used by the design note.
    #[must_use]
    pub fn action_anchors(&self) -> &[ActionAnchor] {
        self.anchors()
    }

    /// All raw action-capable MCE spans in source order. Unsupported spans
    /// remain source-preservation locations and are never promoted by this
    /// catalog to typed action targets.
    #[must_use]
    pub fn source_anchors(&self) -> &[RawAnchor] {
        &self.raw_anchors
    }

    /// Complete relationship set retained from the owning slide.
    #[must_use]
    pub fn owner_relationships(&self) -> &[InboundReference] {
        &self.owner_relationships
    }

    /// Exact source-bound `[Content_Types].xml` bytes in this read set.
    #[must_use]
    pub fn content_types_source(&self) -> &[u8] {
        self.content_types_source.bytes()
    }

    /// Exact source-bound owner `.rels` bytes in this read set.
    #[must_use]
    pub fn owner_relationship_source(&self) -> &[u8] {
        self.owner_relationship_source.bytes()
    }

    /// Whether the owner `.rels` member was present in the source package.
    #[must_use]
    pub fn owner_relationship_member_present(&self) -> bool {
        self.owner_relationship_source.member_present()
    }

    /// Exact source-bound package-root `.rels` bytes in this read set.
    #[must_use]
    pub fn package_relationship_source(&self) -> &[u8] {
        self.package_relationship_source.bytes()
    }

    /// Whether the package-root `.rels` member was present in the source.
    #[must_use]
    pub fn package_relationship_member_present(&self) -> bool {
        self.package_relationship_source.member_present()
    }

    /// OPC read policy captured with this source snapshot.
    #[must_use]
    pub const fn read_limits(&self) -> litchi_opc::ReadLimits {
        self.read_limits
    }

    #[must_use]
    pub const fn revision(&self) -> u64 {
        self.revision
    }

    /// Bounded resource policy carried through edits, commits, and patches.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the semantic selector for an active anchor.
    pub fn selector(&self, ordinal: usize) -> Option<AnchorSelector> {
        self.anchors
            .get(ordinal)
            .map(|anchor| AnchorSelector::Semantic {
                owner: self.slide_index,
                ordinal,
                context: anchor.fingerprint,
            })
    }

    /// Return the source selector for an active anchor.
    pub fn source_selector(&self, ordinal: usize) -> Option<AnchorSelector> {
        self.anchors
            .iter()
            .map(|anchor| (anchor.source_ordinal, anchor.fingerprint))
            .chain(
                self.raw_anchors
                    .iter()
                    .map(|anchor| (anchor.source_ordinal, anchor.fingerprint)),
            )
            .find(|(source_ordinal, _)| *source_ordinal == ordinal)
            .map(|(_, context)| AnchorSelector::Source {
                owner: self.slide_index,
                ordinal,
                context,
            })
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.slide_index == other.slide_index
            && self.slide_part_name == other.slide_part_name
            && self.source_xml == other.source_xml
            && self.revision == other.revision
            && self.anchors.len() == other.anchors.len()
            && self.raw_anchors.len() == other.raw_anchors.len()
            && self.owner_relationships == other.owner_relationships
            && self.content_types_source == other.content_types_source
            && self.owner_relationship_source == other.owner_relationship_source
            && self.package_relationship_source == other.package_relationship_source
            && self.package_relationships == other.package_relationships
            && self.read_limits == other.read_limits
            && self
                .anchors
                .iter()
                .zip(other.anchors.iter())
                .all(|(left, right)| left.same_value(right))
            && self
                .raw_anchors
                .iter()
                .zip(other.raw_anchors.iter())
                .all(|(left, right)| {
                    left.source_ordinal == right.source_ordinal
                        && left.span == right.span
                        && left.fingerprint == right.fingerprint
                })
    }

    pub(crate) fn find(&self, selector: AnchorSelector) -> Result<usize> {
        if selector.owner() != self.slide_index {
            return Err(Error::SlideIndexOutOfBounds {
                index: selector.owner(),
                len: self.slide_index,
            });
        }
        let index = match selector {
            AnchorSelector::Semantic { ordinal, .. } => ordinal,
            AnchorSelector::Source {
                ordinal, context, ..
            } => {
                let index = self
                    .anchors
                    .iter()
                    .position(|anchor| anchor.source_ordinal == ordinal);
                if let Some(index) = index {
                    index
                } else {
                    let raw = self
                        .raw_anchors
                        .iter()
                        .find(|anchor| anchor.source_ordinal == ordinal)
                        .ok_or_else(|| {
                            Error::Invalid("ink-action source selector is absent".into())
                        })?;
                    if raw.fingerprint != context {
                        return Err(Error::StaleSource);
                    }
                    return Err(Error::Invalid(
                        "ink-action source selector names an unsupported branch".into(),
                    ));
                }
            },
        };
        let anchor = self.anchors.get(index).ok_or(Error::IndexOutOfBounds {
            index,
            len: self.anchors.len(),
        })?;
        if anchor.fingerprint != selector.context() {
            return Err(Error::StaleSource);
        }
        Ok(index)
    }

    pub(crate) fn from_parts(
        slide_index: usize,
        slide_part_name: PackURI,
        source_xml: Arc<Vec<u8>>,
        anchors: Vec<ActionAnchor>,
        raw_anchors: Vec<RawAnchor>,
        owner_relationships: Vec<InboundReference>,
        content_types_source: OwnedContentTypes,
        owner_relationship_source: OwnedRelationships,
        package_relationship_source: OwnedRelationships,
        package_relationships: Vec<PackageRelationship>,
        read_limits: litchi_opc::ReadLimits,
        limits: Limits,
    ) -> Self {
        let revision = revision(
            &source_xml,
            &anchors,
            &raw_anchors,
            &owner_relationships,
            &content_types_source,
            &owner_relationship_source,
            &package_relationship_source,
            &package_relationships,
        );
        Self {
            slide_index,
            slide_part_name,
            source_xml,
            anchors: Arc::from(anchors),
            raw_anchors: Arc::from(raw_anchors),
            owner_relationships: Arc::from(owner_relationships),
            content_types_source,
            owner_relationship_source,
            package_relationship_source,
            package_relationships: Arc::from(package_relationships),
            read_limits,
            revision,
            limits,
        }
    }

    /// Begin a source-backed existing-target edit.
    #[must_use]
    pub fn edit(&self) -> super::transaction::Edit {
        super::transaction::Edit::new(self.clone())
    }
}

fn revision(
    source: &[u8],
    anchors: &[ActionAnchor],
    raw_anchors: &[RawAnchor],
    owner_relationships: &[InboundReference],
    content_types_source: &OwnedContentTypes,
    owner_relationship_source: &OwnedRelationships,
    package_relationship_source: &OwnedRelationships,
    package_relationships: &[PackageRelationship],
) -> u64 {
    let mut value = 0xcbf29ce484222325u64;
    for bytes in [source] {
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x100000001b3);
        }
    }
    for anchor in anchors {
        value ^= anchor.fingerprint.value();
        value = value.wrapping_mul(0x100000001b3);
        for byte in anchor.target_bytes() {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x100000001b3);
        }
    }
    for anchor in raw_anchors {
        value ^= anchor.source_ordinal as u64;
        value = value.wrapping_mul(0x100000001b3);
        value ^= anchor.fingerprint.value();
        value = value.wrapping_mul(0x100000001b3);
    }
    for relationship in owner_relationships {
        for bytes in [
            relationship.source_part.as_str().as_bytes(),
            relationship.relationship_id.as_bytes(),
            relationship.relationship_type.as_bytes(),
            relationship.target_ref.as_bytes(),
        ] {
            for byte in bytes {
                value ^= u64::from(*byte);
                value = value.wrapping_mul(0x100000001b3);
            }
        }
        value ^= relationship.target_mode as u8 as u64;
        value = value.wrapping_mul(0x100000001b3);
    }
    for byte in content_types_source.bytes() {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x100000001b3);
    }
    value ^= u64::from(!content_types_source.bytes().is_empty());
    value = value.wrapping_mul(0x100000001b3);
    for source in [owner_relationship_source, package_relationship_source] {
        for byte in source.bytes() {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x100000001b3);
        }
        value ^= u64::from(source.member_present());
        value = value.wrapping_mul(0x100000001b3);
    }
    for relationship in package_relationships {
        for bytes in [
            relationship.relationship_id.as_bytes(),
            relationship.relationship_type.as_bytes(),
            relationship.target_ref.as_bytes(),
        ] {
            for byte in bytes {
                value ^= u64::from(*byte);
                value = value.wrapping_mul(0x100000001b3);
            }
        }
        value ^= relationship.target_mode as u8 as u64;
        value = value.wrapping_mul(0x100000001b3);
    }
    value
}

// Keep the codec module in this file's dependency graph even when callers only
// use the model re-exports.  This avoids leaking the scanner's physical types.
#[allow(dead_code)]
fn _codec_namespace_marker() -> &'static [u8] {
    codec::ACTION
}
