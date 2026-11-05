//! The staged part-level engine behind the XLSB package edit families
//! (unified-ops records 0322 and 0360b).
//!
//! [`Staged`] owns one immutable [`Package`] (a reference-count clone, with
//! its cached workbook projection) and, from the first change on, a staged
//! copy of its OPC graph: starting an edit copies nothing. A family stages
//! whole-part replacements, additions and removals, or runs typed graph
//! operations with [`Staged::update`], and commits: [`PartPatch`] captures
//! the exact before/after state (content type, payload and relationships) of
//! each changed part and of the package-level relationships, with the
//! publication manifests of both sides, and an unchanged staging returns the
//! source package itself, so an exact no-op keeps every byte. The source
//! package never changes.
//!
//! Worksheet streams are edited at the record level: [`SheetLayout`] splits
//! the records after the cell table into top-level units (one record, or a
//! begin/end collection), so a family replaces, inserts or removes exactly
//! its own units and copies every other byte. A rewritten worksheet keeps its
//! binary index current ([`crate::worksheet_index::maintain`]).
//!
//! The families built on it ([`crate::page_breaks::edit`],
//! [`crate::auto_filter::edit`], [`crate::hyperlinks::edit`],
//! [`crate::data_validation::edit`], [`crate::conditional_formatting::edit`],
//! [`crate::comments::edit`], [`crate::drawing::edit`], [`crate::opc`] and the
//! package-level families) expose the ADR 0034 transaction shape through
//! [`family!`].

use std::collections::{BTreeSet, HashSet};
use std::ops::Range;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};
use sha2::{Digest as _, Sha256};

use crate::Workbook;
use crate::package::Package;
use crate::package::error::{Error, Result};
use crate::raw::{Kind, Records, kind};

pub(crate) fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

fn relationships(graph: &Relationships) -> Vec<Relationship> {
    let mut states: Vec<_> = graph
        .iter()
        .map(|relationship| Relationship {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        })
        .collect();
    states.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    states
}

fn add_relationships(graph: &mut Relationships, states: &[Relationship]) -> Result<()> {
    for relationship in states {
        graph.try_add_relationship(
            relationship.kind.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            if relationship.external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

/// Whether two relationship collections hold the same relationships, without
/// allocating.
fn relationships_equal(left: &Relationships, right: &Relationships) -> bool {
    left.len() == right.len()
        && left.iter().all(|relationship| {
            right.get(relationship.r_id()).is_some_and(|candidate| {
                candidate.reltype() == relationship.reltype()
                    && candidate.target_ref() == relationship.target_ref()
                    && candidate.is_external() == relationship.is_external()
            })
        })
}

/// Whether a typed relationship graph has the captured entries, without
/// allocating a temporary sorted vector during patch replay.
fn relationships_match(graph: &Relationships, states: &[Relationship]) -> bool {
    graph.len() == states.len()
        && states.iter().all(|state| {
            graph.get(&state.id).is_some_and(|candidate| {
                candidate.reltype() == state.kind
                    && candidate.target_ref() == state.target
                    && candidate.is_external() == state.external
            })
        })
}

/// One relationship owner's exact semantic and publication state.
///
/// OPC relationship members are not ordinary parts: their XML spelling lives
/// beside the typed edge graph. Keep both views in a patch so a raw edit that
/// changes only whitespace, attribute order, or another admitted lexical
/// detail remains an observable, reversible edit.
/// The admitted source member, or `None` for a deterministic authored
/// relationship graph. Keeping this as an optional Arc preserves the compact
/// payload-only state layout while deferring canonical serialization.
type RelationshipXml = Option<Arc<Vec<u8>>>;

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    entries: Vec<Relationship>,
    xml: RelationshipXml,
    member_present: bool,
}

impl RelationshipState {
    fn read(package: &OpcPackage, owner: &PackURI, graph: &Relationships) -> Result<Self> {
        let (xml, member_present) =
            if let Some((xml, member_present)) = package.retained_relationships_xml(owner)? {
                (Some(xml), member_present)
            } else if owner.as_str() != "/" && graph.is_empty() {
                (None, false)
            } else {
                (None, owner.as_str() == "/" || !graph.is_empty())
            };
        Ok(Self {
            entries: relationships(graph),
            xml,
            member_present,
        })
    }

    fn matches(
        &self,
        package: &OpcPackage,
        owner: &PackURI,
        graph: &Relationships,
    ) -> Result<bool> {
        if !relationships_match(graph, &self.entries) {
            return Ok(false);
        }
        let current = package.retained_relationships_xml(owner)?;
        let member_present = current.as_ref().map_or_else(
            || owner.as_str() == "/" || !graph.is_empty(),
            |(_, member_present)| *member_present,
        );
        if self.member_present != member_present {
            return Ok(false);
        }
        if !self.member_present {
            return Ok(true);
        }
        if let Some((bytes, _)) = current {
            return match &self.xml {
                Some(expected) => Ok(expected.as_slice() == bytes.as_slice()),
                None => {
                    let authored =
                        package.authored_relationships_xml(owner, package.read_limits())?;
                    Ok(authored.as_slice() == bytes.as_slice())
                },
            };
        }
        match &self.xml {
            Some(expected) => {
                let authored = package.authored_relationships_xml(owner, package.read_limits())?;
                Ok(expected.as_slice() == authored.as_slice())
            },
            None => Ok(true),
        }
    }
}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: RelationshipState,
}

impl PartialEq for State {
    fn eq(&self, other: &Self) -> bool {
        self.content_type == other.content_type
            && (Arc::ptr_eq(&self.blob, &other.blob) || self.blob == other.blob)
            && self.relationships == other.relationships
    }
}

impl Eq for State {}

impl State {
    fn read(package: &OpcPackage, part: &dyn Part) -> Result<Self> {
        Ok(Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: RelationshipState::read(package, part.partname(), part.rels())?,
        })
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Result<Option<Self>> {
        match package.get_part(name) {
            Ok(part) => Ok(Some(Self::read(package, part)?)),
            Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Whether `part` holds exactly this state, without copying its payload.
    fn matches(&self, package: &OpcPackage, part: &dyn Part) -> Result<bool> {
        Ok(self.content_type == part.content_type()
            && {
                let blob = part.blob();
                std::ptr::eq(self.blob.as_slice(), blob) || self.blob.as_slice() == blob
            }
            && self
                .relationships
                .matches(package, part.partname(), part.rels())?)
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        add_relationships(part.rels_mut(), &self.relationships.entries)?;
        Ok(part)
    }
}

fn optional_part<'a>(package: &'a OpcPackage, name: &PackURI) -> Result<Option<&'a dyn Part>> {
    match package.get_part(name) {
        Ok(part) => Ok(Some(part)),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Delta {
    name: PackURI,
    before: Option<State>,
    after: Option<State>,
}

/// A compact retained content-types image for one side of a patch.
///
/// The byte image is enough for bounded exact replacement; retaining the
/// whole parsed package graph just to defer this capture would pin every
/// unchanged part while callers keep a patch alive.
#[derive(Debug, Clone, PartialEq, Eq)]
enum ContentTypesXml {
    /// The admitted source member, kept for exact lexical replay.
    Retained(Arc<Vec<u8>>),
    /// The deterministic authored form of the typed content-type graph.
    /// Serialize it only when replay needs to compare or publish it.
    Authored,
}

impl ContentTypesXml {
    fn bytes(
        &self,
        package: &OpcPackage,
        limits: litchi_opc::SourceLimits,
    ) -> Result<Arc<Vec<u8>>> {
        match self {
            Self::Retained(bytes) => Ok(Arc::clone(bytes)),
            Self::Authored => package
                .authored_content_types_xml(limits)
                .map_err(Into::into),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ContentTypesImage {
    /// The exact replacement image. Authored images are serialized only when
    /// replay needs to publish them.
    bytes: ContentTypesXml,
    /// The publication view for this graph. This can be authored XML even
    /// when `bytes` is an admitted source image whose binding is temporarily
    /// incompatible with the typed graph.
    publication: ContentTypesXml,
    source: bool,
    explicit: bool,
    /// A canonical publication guard for the authored side of an explicit
    /// lexical replacement.  It keeps replay exact without serializing the
    /// authored manifest during every ordinary structural commit.
    exact_canonical: Option<[u8; 32]>,
}

impl ContentTypesImage {
    fn bind_exact(&mut self, package: &OpcPackage) -> Result<()> {
        if matches!(self.publication, ContentTypesXml::Authored) {
            let bytes = package.authored_content_types_xml(package.read_limits())?;
            self.exact_canonical = Some(Sha256::digest(&*bytes).into());
        }
        Ok(())
    }

    fn current_publication(package: &OpcPackage) -> Result<ContentTypesXml> {
        Ok(match package.retained_content_types_xml()? {
            Some(bytes) => ContentTypesXml::Retained(bytes),
            None => ContentTypesXml::Authored,
        })
    }

    fn publication_matches_current(&self, package: &OpcPackage) -> Result<bool> {
        let current = Self::current_publication(package)?;
        if let Some(expected) = self.exact_canonical {
            let bytes = current.bytes(package, package.read_limits())?;
            let digest: [u8; 32] = Sha256::digest(&*bytes).into();
            return Ok(digest == expected);
        }
        match (&self.publication, current) {
            (ContentTypesXml::Retained(expected), ContentTypesXml::Retained(actual)) => {
                Ok(expected.as_slice() == actual.as_slice())
            },
            (ContentTypesXml::Authored, ContentTypesXml::Authored) => Ok(true),
            (expected, actual) => {
                let limits = package.read_limits();
                let expected = expected.bytes(package, limits)?;
                let actual = actual.bytes(package, limits)?;
                Ok(expected.as_slice() == actual.as_slice())
            },
        }
    }

    /// Whether this image still describes the destination before the patch's
    /// typed deltas are applied. Explicit raw replacements need both the
    /// publication view and the admitted source allocation: comparing only
    /// source bytes would accept an unrelated typed manifest overlay.
    fn matches_base(&self, package: &OpcPackage) -> Result<bool> {
        let retained = package.retained_content_types_xml()?;
        if !self.publication_matches_current(package)? {
            return Ok(false);
        }
        if self.source {
            let source = package.content_types_source_bytes();
            let source_matches = source
                .as_ref()
                .is_some_and(|bytes| {
                    matches!(&self.bytes, ContentTypesXml::Retained(expected) if bytes.as_slice() == expected.as_slice())
                });
            // A source member that still binds to the typed graph is the
            // publication already checked above. Reopening can regenerate
            // that publication and discard an invalid, non-published source
            // image, so provenance is optional when the publication matches.
            return Ok(source_matches || source.is_none() || retained.is_some());
        }
        Ok(true)
    }
}

/// A compact retained relationship image for one changed owner.
#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipImage {
    owner: PackURI,
    bytes: RelationshipXml,
    member_present: bool,
}

/// The publication manifests of one side of a patch: the content-types view
/// and the relationship views of the changed owners, so a replay publishes
/// the same manifest spelling as the commit it was captured from.
///
/// Only the views a patch can change are captured: the relationships of a
/// part that is added, removed or whose relationships change, the
/// content-types view when a part is added or removed or changes its content
/// type, and the package relationships when they change. A part whose
/// payload alone changes is replayed in place, so its own views are kept.
#[derive(Debug, Clone, PartialEq, Eq)]
struct Manifests {
    content_types: Option<ContentTypesImage>,
    relationships: Vec<RelationshipImage>,
}

impl Delta {
    /// Whether only the payload changes: the part exists on both sides with
    /// the same content type and relationships.
    fn payload_only(&self) -> bool {
        match (&self.before, &self.after) {
            (Some(before), Some(after)) => {
                before.content_type == after.content_type
                    && before.relationships == after.relationships
            },
            _ => false,
        }
    }
}

impl Manifests {
    /// Whether a patch of `deltas` changes any manifest view.
    fn needed(deltas: &[Delta], root: bool, content_types_changed: bool) -> bool {
        root || content_types_changed || deltas.iter().any(|delta| !delta.payload_only())
    }

    fn capture(
        deltas: &[Delta],
        root: Option<&(RelationshipState, RelationshipState)>,
        after: bool,
        content_types: Option<ContentTypesImage>,
    ) -> Result<Self> {
        let mut relationships = Vec::new();
        for delta in deltas.iter().filter(|delta| !delta.payload_only()) {
            let state = if after {
                delta.after.as_ref()
            } else {
                delta.before.as_ref()
            };
            if let Some(state) = state.filter(|state| state.relationships.member_present) {
                relationships.push(RelationshipImage {
                    owner: delta.name.clone(),
                    bytes: state.relationships.xml.clone(),
                    member_present: true,
                });
            }
        }
        if let Some((before, after_state)) = root {
            let state = if after { after_state } else { before };
            let owner = PackURI::new("/")?;
            relationships.push(RelationshipImage {
                owner,
                bytes: state.xml.clone(),
                member_present: state.member_present,
            });
        }
        Ok(Self {
            content_types,
            relationships,
        })
    }

    /// Publish these manifest views on `package`, whose typed graph is the
    /// one they were captured from.
    fn restore(
        &self,
        package: &mut OpcPackage,
        expected: Option<&ContentTypesImage>,
    ) -> Result<()> {
        let limits = package.read_limits();
        if let Some(content_types) = &self.content_types {
            let current = ContentTypesImage::current_publication(package)?;
            // An authored view is already the package's lazy fallback after
            // a structural graph change invalidates retained coverage. Avoid
            // serializing the same canonical XML twice during replay; the
            // writer will materialize it only when publication needs bytes.
            if !matches!(
                (&current, &content_types.bytes),
                (ContentTypesXml::Authored, ContentTypesXml::Authored)
            ) {
                let replacement = content_types.bytes.bytes(package, limits)?;
                let current_bytes = current.bytes(package, limits)?;
                let source_differs = content_types.explicit
                    && content_types.source
                    && package
                        .content_types_source_bytes()
                        .is_none_or(|bytes| {
                            !matches!(&content_types.bytes, ContentTypesXml::Retained(expected) if bytes.as_slice() == expected.as_slice())
                        });
                if current_bytes.as_slice() != replacement.as_slice() || source_differs {
                    let current_matches_base = match expected {
                        Some(expected) => expected.publication_matches_current(package)?,
                        None => false,
                    };
                    if content_types.explicit || current_matches_base {
                        package.try_replace_content_types_bytes(
                            current_bytes.as_slice(),
                            replacement,
                            limits,
                        )?;
                    }
                }
            }
        }
        for relationship in &self.relationships {
            let (current, current_member_present) = if let Some((bytes, member_present)) =
                package.retained_relationships_xml(&relationship.owner)?
            {
                (bytes, member_present)
            } else {
                let member_present = if relationship.owner.as_str() == "/" {
                    true
                } else {
                    match package.get_part(&relationship.owner) {
                        Ok(part) => !part.rels().is_empty(),
                        Err(litchi_opc::OpcError::PartNotFound(_)) => false,
                        Err(error) => return Err(error.into()),
                    }
                };
                (
                    package.authored_relationships_xml(&relationship.owner, limits)?,
                    member_present,
                )
            };
            let replacement = match &relationship.bytes {
                Some(bytes) => Arc::clone(bytes),
                None => package.authored_relationships_xml(&relationship.owner, limits)?,
            };
            if current.as_slice() != replacement.as_slice()
                || current_member_present != relationship.member_present
            {
                package.try_replace_relationships_source_bytes(
                    &relationship.owner,
                    current.as_slice(),
                    current_member_present,
                    replacement,
                    relationship.member_present,
                    limits,
                )?;
            }
        }
        Ok(())
    }
}

/// The semantic retained buffers of a native patch.
///
/// The patch keeps images, relationship data and publication-manifest proof,
/// not Rust container capacity or parser caches.  The semantic budget charges
/// each before/after endpoint, even when package clones share an allocation;
/// duplicate views of one endpoint (for example a relationship image and its
/// state) are accounted only once.
#[derive(Debug, Default)]
struct RetainedAccounting {
    bytes: usize,
}

impl RetainedAccounting {
    fn add_bytes(&mut self, bytes: usize) {
        self.bytes = self.bytes.saturating_add(bytes);
    }

    fn add_str(&mut self, value: &str) {
        self.add_bytes(value.len());
    }

    fn add_arc(&mut self, bytes: &Arc<Vec<u8>>) {
        self.add_bytes(bytes.len());
    }

    fn add_storage(&mut self, bytes: usize) {
        self.add_bytes(bytes);
    }

    fn add_relationship_state(&mut self, state: &RelationshipState) {
        if let Some(bytes) = &state.xml {
            self.add_arc(bytes);
        }
        for relationship in &state.entries {
            self.add_str(&relationship.id);
            self.add_str(&relationship.kind);
            self.add_str(&relationship.target);
        }
    }

    fn add_state(&mut self, state: &State) {
        self.add_str(&state.content_type);
        self.add_arc(&state.blob);
        self.add_relationship_state(&state.relationships);
    }

    fn add_metadata_state(
        &mut self,
        package: &OpcPackage,
        name: &PackURI,
        metadata: litchi_opc::PartMetadata<'_>,
    ) -> Result<()> {
        self.add_str(metadata.content_type());
        self.add_storage(metadata.payload_len()?);
        self.add_relationship_state(&RelationshipState::read(package, name, metadata.rels())?);
        Ok(())
    }

    fn add_content_types(&mut self, image: &ContentTypesImage) {
        // `bytes` and `publication` can be two views of the same endpoint
        // allocation.  Count that endpoint once, while separate before and
        // after images are charged independently by their callers.
        let mut endpoint_allocations = HashSet::new();
        for xml in [&image.bytes, &image.publication] {
            if let ContentTypesXml::Retained(bytes) = xml {
                if endpoint_allocations.insert(Arc::as_ptr(bytes) as usize) {
                    self.add_arc(bytes);
                }
            }
        }
        if image.exact_canonical.is_some()
            || (image.explicit && matches!(&image.publication, ContentTypesXml::Authored))
        {
            // An explicit authored replacement binds this digest during
            // manifest capture.  Admission has not materialized it yet, so
            // account for the same 32-byte endpoint that final capture will
            // retain without serializing canonical XML here.
            self.add_bytes(32);
        }
    }

    fn add_manifests(&mut self, manifests: &Manifests) {
        if let Some(content_types) = &manifests.content_types {
            self.add_content_types(content_types);
        }
        for relationship in &manifests.relationships {
            self.add_str(relationship.owner.as_str());
        }
    }
}

/// The exact, reversible part-level change of one staged edit: the changed
/// parts and, when they changed, the package-level relationships, with the
/// publication manifests of both sides.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
    root: Option<(RelationshipState, RelationshipState)>,
    manifests: Option<LazyManifests>,
}

/// The compact manifest views of both sides of a changed patch. The changed
/// part states already retain every relationship image needed for replay, so
/// this value does not retain either whole package graph.
#[derive(Debug, Clone)]
struct LazyManifests {
    captured: Arc<(Manifests, Manifests)>,
    content_types_changed: bool,
    /// Whether this is the inverse of the captured direction.
    inverse: bool,
}

impl LazyManifests {
    /// The `(before, after)` manifest views of this direction.
    fn get(&self) -> (&Manifests, &Manifests) {
        let pair = self.captured.as_ref();
        if self.inverse {
            (&pair.1, &pair.0)
        } else {
            (&pair.0, &pair.1)
        }
    }

    fn inverse(&self) -> Self {
        Self {
            inverse: !self.inverse,
            ..self.clone()
        }
    }
}

/// Two patches are equal when they change the same parts and manifest images
/// the same way.
impl PartialEq for LazyManifests {
    fn eq(&self, other: &Self) -> bool {
        self.content_types_changed == other.content_types_changed
            && self.inverse == other.inverse
            && self.captured == other.captured
    }
}

impl Eq for LazyManifests {}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none() && self.manifests.is_none()
    }

    /// The number of parts (and the package-level relationship set) the
    /// patch changes.
    pub(crate) fn parts(&self) -> usize {
        let parts = self.deltas.len() + usize::from(self.root.is_some());
        if parts == 0 {
            usize::from(
                self.manifests
                    .as_ref()
                    .is_some_and(|manifests| manifests.content_types_changed),
            )
        } else {
            parts
        }
    }

    /// The semantic bytes the patch retains: endpoint payloads, relationship
    /// proof and publication-manifest images.
    pub(crate) fn retained_bytes(&self) -> usize {
        let mut accounting = RetainedAccounting::default();
        for delta in &self.deltas {
            accounting.add_str(delta.name.as_str());
            if let Some(state) = &delta.before {
                accounting.add_state(state);
            }
            if let Some(state) = &delta.after {
                accounting.add_state(state);
            }
        }
        if let Some((before, after)) = &self.root {
            accounting.add_relationship_state(before);
            accounting.add_relationship_state(after);
        }
        if let Some(manifests) = &self.manifests {
            let (before, after) = manifests.get();
            accounting.add_manifests(before);
            accounting.add_manifests(after);
        }
        accounting.bytes
    }

    /// Calculate the semantic charge for the named part states from package
    /// metadata before retaining or inflating any deferred payload.
    fn retained_bytes_for_names(
        before: &OpcPackage,
        after: &OpcPackage,
        names: &[PackURI],
        changed_parts: &[PackURI],
    ) -> Result<(usize, usize)> {
        let mut accounting = RetainedAccounting::default();
        let mut changed = 0usize;
        for name in names {
            let typed_changed = changed_parts
                .binary_search_by(|current| current.as_str().cmp(name.as_str()))
                .is_ok();
            let lexical_changed = !before.same_retained_relationships(after, name)?;
            if !typed_changed && !lexical_changed {
                continue;
            }
            changed = changed.saturating_add(1);
            accounting.add_str(name.as_str());
            let before_metadata = before.part_metadata(name);
            let after_metadata = after.part_metadata(name);
            if let Some(metadata) = before_metadata {
                accounting.add_metadata_state(before, name, metadata)?;
            }
            if let Some(metadata) = after_metadata {
                accounting.add_metadata_state(after, name, metadata)?;
            }

            // A relationship image is retained for each non-payload-only
            // endpoint.  Count its owner string here, before the payloads are
            // inflated, so the admission check covers the complete semantic
            // proof rather than only the part blobs.
            let relationship_changed = match (before_metadata, after_metadata) {
                (Some(left), Some(right)) => {
                    left.content_type() != right.content_type()
                        || !relationships_equal(left.rels(), right.rels())
                        || !before.same_retained_relationships(after, name)?
                },
                _ => true,
            };
            if relationship_changed {
                let before_member = before_metadata
                    .map(|metadata| RelationshipState::read(before, name, metadata.rels()))
                    .transpose()?
                    .is_some_and(|state| state.member_present);
                let after_member = after_metadata
                    .map(|metadata| RelationshipState::read(after, name, metadata.rels()))
                    .transpose()?
                    .is_some_and(|state| state.member_present);
                if before_member {
                    accounting.add_str(name.as_str());
                }
                if after_member {
                    accounting.add_str(name.as_str());
                }
            }
        }

        // Root relationships and the content-types image are manifest proof,
        // not part payloads.  They are metadata-only to inspect, so account
        // them before any State::of call can inflate a source member.
        if let Some((root_before, root_after)) = Self::root(before, after)? {
            accounting.add_relationship_state(&root_before);
            accounting.add_relationship_state(&root_after);
            accounting.add_str("/");
            accounting.add_str("/");
        }
        if Self::content_types_changed_metadata(before, after, names)? {
            let mut before_image = Self::content_types_image(before)?;
            let mut after_image = Self::content_types_image(after)?;
            let source_changed =
                before.content_types_source_bytes() != after.content_types_source_bytes();
            let explicit = source_changed && (before_image.source || after_image.source);
            before_image.explicit = explicit;
            after_image.explicit = explicit;
            if !explicit {
                before_image.bytes = before_image.publication.clone();
                after_image.bytes = after_image.publication.clone();
            }
            accounting.add_content_types(&before_image);
            accounting.add_content_types(&after_image);
        }
        Ok((accounting.bytes, changed))
    }

    /// The content-types graph can be classified without reading any part
    /// payload.  This mirrors [`Self::content_types_changed`] for the
    /// metadata-only admission pass.
    fn content_types_changed_metadata(
        before: &OpcPackage,
        after: &OpcPackage,
        names: &[PackURI],
    ) -> Result<bool> {
        let source_changed = match (
            before.content_types_source_bytes(),
            after.content_types_source_bytes(),
        ) {
            (None, None) => false,
            (Some(left), Some(right)) => {
                !(Arc::ptr_eq(&left, &right) || left.as_slice() == right.as_slice())
            },
            (Some(_), None) | (None, Some(_)) => true,
        };
        if source_changed {
            return Ok(true);
        }
        for name in names {
            match (before.part_metadata(name), after.part_metadata(name)) {
                (Some(left), Some(right)) if left.content_type() == right.content_type() => {},
                (Some(_), Some(_)) | (Some(_), None) | (None, Some(_)) => return Ok(true),
                (None, None) => {},
            }
        }
        Ok(false)
    }

    fn enforce_retained_limit(retained: usize, maximum: Option<usize>) -> Result<()> {
        if let Some(maximum) = maximum.filter(|maximum| retained > *maximum) {
            return Err(Error::limit(
                crate::opc::LimitKind::PatchBytes,
                retained,
                maximum,
            ));
        }
        Ok(())
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            deltas: self
                .deltas
                .iter()
                .map(|delta| Delta {
                    name: delta.name.clone(),
                    before: delta.after.clone(),
                    after: delta.before.clone(),
                })
                .collect(),
            root: self
                .root
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
            manifests: self.manifests.as_ref().map(LazyManifests::inverse),
        }
    }

    /// Whether the patch leaves the workbook projection's dependency closure
    /// unchanged.
    ///
    /// A raw edit may add, remove, or rewrite an OPC part that is not reachable
    /// from any package relationship.  Such a change still needs the normal
    /// patch-state, manifest, signature, and source-limit checks, but it cannot
    /// alter the typed workbook.  Keeping the projection lazy for that narrow
    /// case avoids parsing the same workbook graph twice when a caller only
    /// needs to publish an opaque part.  The recipient's before graph is the
    /// only graph that needs to be walked. With no root delta, every
    /// relationship owner that can differ after apply is already in
    /// `self.deltas`: an unchanged reachable owner cannot acquire a new edge
    /// to a changed part. Therefore proving the changed set unreachable before
    /// apply also proves it unreachable afterward, while still catching a
    /// recipient's unrelated relationship to a part that was inert in the
    /// recorded base.
    ///
    /// A malformed or dangling relationship is treated conservatively as
    /// requiring a reparse.  The full workbook/OPC validation remains the
    /// authority for that case.
    pub(crate) fn projection_unchanged(&self, before: &OpcPackage) -> bool {
        // A root relationship edit can select a different workbook or expose
        // a previously unreachable dependency.  Keep the existing full
        // validation path even when the lexical relationship bytes happen to
        // describe the same typed edges.
        if self.root.is_some() {
            return false;
        }

        let mut changed = BTreeSet::new();
        for delta in &self.deltas {
            changed.insert(delta.name.as_str().to_ascii_lowercase());
        }
        // A manifest-only patch changes source publication bytes, not the
        // already typed part graph.  The replacement is still restored and
        // bounded by `apply_graph`; no workbook dependency was changed.
        if changed.is_empty() {
            return true;
        }

        closure_is_disjoint(before, &changed)
    }

    /// Attach the manifest views of both sides to a changed patch.
    fn with_manifests(
        mut self,
        before: &Arc<OpcPackage>,
        after: &Arc<OpcPackage>,
        content_types_changed: bool,
    ) -> Result<Self> {
        if Manifests::needed(&self.deltas, self.root.is_some(), content_types_changed) {
            let (before_content_types, after_content_types) = if content_types_changed {
                let mut before_content_types = Self::content_types_image(before)?;
                let mut after_content_types = Self::content_types_image(after)?;
                let source_changed =
                    before.content_types_source_bytes() != after.content_types_source_bytes();
                // A source allocation change is the provenance of an explicit
                // raw manifest replacement. Mark both directions so an
                // inverse cannot silently discard that lexical image on a
                // divergent target; a structural source-to-fallback change
                // keeps the allocation and leaves both sides composable.
                let explicit =
                    source_changed && (before_content_types.source || after_content_types.source);
                before_content_types.explicit = explicit;
                after_content_types.explicit = explicit;
                if explicit {
                    before_content_types.bind_exact(before)?;
                    after_content_types.bind_exact(after)?;
                }
                if !explicit {
                    before_content_types.bytes = before_content_types.publication.clone();
                    after_content_types.bytes = after_content_types.publication.clone();
                }
                (Some(before_content_types), Some(after_content_types))
            } else {
                (None, None)
            };
            let before_manifests = Manifests::capture(
                &self.deltas,
                self.root.as_ref(),
                false,
                before_content_types,
            )?;
            let after_manifests =
                Manifests::capture(&self.deltas, self.root.as_ref(), true, after_content_types)?;
            self.manifests = Some(LazyManifests {
                captured: Arc::new((before_manifests, after_manifests)),
                content_types_changed,
                inverse: false,
            });
        }
        Ok(self)
    }

    fn content_types_image(package: &OpcPackage) -> Result<ContentTypesImage> {
        let source_bytes = package.content_types_source_bytes();
        let source = source_bytes.is_some();
        let publication = match package.retained_content_types_xml()? {
            Some(bytes) => ContentTypesXml::Retained(bytes),
            None => ContentTypesXml::Authored,
        };
        let bytes = match source_bytes {
            Some(bytes) => ContentTypesXml::Retained(bytes),
            None => publication.clone(),
        };
        Ok(ContentTypesImage {
            bytes,
            publication,
            source,
            explicit: false,
            exact_canonical: None,
        })
    }

    /// Content-types publication is affected by a part being added or
    /// removed, by a part's declared type changing, or by an explicit source
    /// manifest replacement. Payload and relationship-only deltas do not
    /// change the content-types graph. Checking the delta shape first avoids
    /// rebuilding a complete source-coverage index for every ordinary raw
    /// structural commit.
    fn content_types_changed(before: &OpcPackage, after: &OpcPackage, deltas: &[Delta]) -> bool {
        let source_changed = match (
            before.content_types_source_bytes(),
            after.content_types_source_bytes(),
        ) {
            (None, None) => false,
            (Some(left), Some(right)) => {
                !(Arc::ptr_eq(&left, &right) || left.as_slice() == right.as_slice())
            },
            (Some(_), None) | (None, Some(_)) => true,
        };
        source_changed
            || deltas
                .iter()
                .any(|delta| match (&delta.before, &delta.after) {
                    (None, Some(_)) | (Some(_), None) => true,
                    (Some(before), Some(after)) => before.content_type != after.content_type,
                    (None, None) => false,
                })
    }

    fn root(
        before: &OpcPackage,
        after: &OpcPackage,
    ) -> Result<Option<(RelationshipState, RelationshipState)>> {
        let owner = PackURI::new("/")?;
        if relationships_equal(before.rels(), after.rels())
            && before.same_retained_relationships(after, &owner)?
        {
            Ok(None)
        } else {
            Ok(Some((
                RelationshipState::read(before, &owner, before.rels())?,
                RelationshipState::read(after, &owner, after.rels())?,
            )))
        }
    }

    /// Capture the parts named in `names` that differ between `before` and
    /// `after`, and the package-level relationships when they differ.
    fn capture(
        before: &Arc<OpcPackage>,
        after: &Arc<OpcPackage>,
        names: &BTreeSet<String>,
        max_snapshot_bytes: Option<usize>,
    ) -> Result<Self> {
        let names: Vec<PackURI> = names
            .iter()
            .map(|name| PackURI::new(name.as_str()))
            .collect::<std::result::Result<_, _>>()
            .map_err(Error::from)?;
        let changed_parts = after.changed_parts(before);
        let (retained, _) = Self::retained_bytes_for_names(before, after, &names, &changed_parts)?;
        Self::enforce_retained_limit(retained, max_snapshot_bytes)?;
        let mut deltas = Vec::new();
        for name in names {
            let before = State::of(before, &name)?;
            let after = State::of(after, &name)?;
            if before != after {
                deltas.push(Delta {
                    name,
                    before,
                    after,
                });
            }
        }
        let content_types_changed = Self::content_types_changed(before, after, &deltas);
        let patch = Self {
            deltas,
            root: Self::root(before, after)?,
            manifests: None,
        }
        .with_manifests(before, after, content_types_changed)?;
        Self::enforce_retained_limit(patch.retained_bytes(), max_snapshot_bytes)?;
        Ok(patch)
    }

    /// Capture every part that differs between `before` and `after`.
    ///
    /// The staged graph is a clone of the source, so an untouched part shares
    /// its payload allocation and compares by pointer; only changed parts are
    /// compared by value.
    pub(crate) fn capture_all(
        before: &Arc<OpcPackage>,
        after: &Arc<OpcPackage>,
        max_snapshot_bytes: Option<usize>,
    ) -> Result<Self> {
        // One commit changes at most as many parts as the package's source
        // limits admit.
        let maximum = before.read_limits().max_parts();
        // `changed_parts` compares the package-owned payload handles and
        // relationship graphs without decoding retained source members. The
        // staged graph is a clone of `before`, so every untouched payload is
        // one allocation and drops out here. Only names in this set need the
        // fallible `State::read` path below.
        let mut changed_names = after.changed_parts(before);

        // A relationship member's lexical bytes are outside the typed graph.
        // Add owners whose semantic entries stayed equal but whose retained
        // source capture changed. Metadata iteration cannot reach payloads.
        if !before.same_relationship_source_captures(after) {
            for metadata in before.iter_parts() {
                let name = metadata.partname();
                if changed_names
                    .binary_search_by(|current| current.as_str().cmp(name.as_str()))
                    .is_ok()
                {
                    continue;
                }
                if let Some(other) = after.part_metadata(name)
                    && relationships_equal(metadata.rels(), other.rels())
                    && !before.same_retained_relationships(after, name)?
                {
                    changed_names.push(name.clone());
                }
            }
        }
        changed_names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        changed_names.dedup();

        let (retained, changed) =
            Self::retained_bytes_for_names(before, after, &changed_names, &changed_names)?;
        if changed > maximum {
            return Err(Error::limit(crate::opc::LimitKind::Parts, changed, maximum));
        }
        Self::enforce_retained_limit(retained, max_snapshot_bytes)?;

        let mut deltas = Vec::new();
        let mut push = |delta: Delta| -> Result<()> {
            if delta.before == delta.after {
                return Ok(());
            }
            if deltas.len() >= maximum {
                return Err(Error::limit(
                    crate::opc::LimitKind::Parts,
                    deltas.len() + 1,
                    maximum,
                ));
            }
            deltas.push(delta);
            Ok(())
        };

        for name in changed_names {
            let before_state = State::of(before, &name)?;
            let after_state = State::of(after, &name)?;
            push(Delta {
                name,
                before: before_state,
                after: after_state,
            })?;
        }
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let content_types_changed = Self::content_types_changed(before, after, &deltas);
        let patch = Self {
            deltas,
            root: Self::root(before, after)?,
            manifests: None,
        }
        .with_manifests(before, after, content_types_changed)?;
        Self::enforce_retained_limit(patch.retained_bytes(), max_snapshot_bytes)?;
        Ok(patch)
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state. Dependency-disjoint changes keep the workbook projection
    /// lazy; changes in the relationship closure are reparsed. An empty patch
    /// returns `package`.
    pub(crate) fn apply(&self, package: &Package) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        let target = self.apply_graph(package.opc_package())?;
        if self.projection_unchanged(package.opc_package()) {
            Workbook::validate_edit_candidate(&target)?;
            Ok(package.with_unparsed(target))
        } else {
            package.reparse_candidate_retaining_projection(target)
        }
    }

    /// Replay onto `package` and publish the result without validating its
    /// workbook graph again: for parts the workbook projection never reads
    /// (signatures, ribbon, task panes, VBA, theme).
    pub(crate) fn apply_unparsed(&self, package: &Package) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        let target = self.apply_graph(package.opc_package())?;
        Ok(package.with_unparsed(target))
    }

    /// Replay onto a copy of the OPC graph `source`, refusing a graph that is
    /// not the patch's before state; `source` is unchanged.
    pub(crate) fn apply_graph(&self, source: &OpcPackage) -> Result<OpcPackage> {
        if let Some((before, _)) = &self.root {
            let owner = PackURI::new("/")?;
            if !before.matches(source, &owner, source.rels())? {
                return Err(Error::PatchConflict {
                    part: "/_rels/.rels",
                });
            }
        }
        for delta in &self.deltas {
            let current = optional_part(source, &delta.name)?;
            let matches = match (&delta.before, current) {
                (None, None) => true,
                (Some(state), Some(part)) => state.matches(source, part)?,
                _ => false,
            };
            if !matches {
                return Err(Error::PatchConflict {
                    part: "a changed part",
                });
            }
        }
        crate::signature_policy::ensure_unsigned(source)?;
        if let Some(manifests) = &self.manifests {
            let (before, _) = manifests.get();
            if let Some(content_types) = before.content_types.as_ref()
                && content_types.explicit
                && !content_types.matches_base(source)?
            {
                return Err(Error::PatchConflict {
                    part: "content types",
                });
            }
        }
        let mut target = source.clone();
        for delta in &self.deltas {
            if delta.payload_only()
                && let Some(after) = &delta.after
            {
                // Keep the part's own relationship and content-type views.
                target
                    .get_part_mut(&delta.name)?
                    .set_blob_shared(Arc::clone(&after.blob));
                continue;
            }
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        if let Some((before, after)) = &self.root {
            let graph = target.rels_mut();
            for relationship in &before.entries {
                graph.remove(&relationship.id);
            }
            add_relationships(graph, &after.entries)?;
        }
        if let Some(manifests) = &self.manifests {
            let (before, after) = manifests.get();
            after.restore(&mut target, before.content_types.as_ref())?;
        }
        Ok(target)
    }
}

/// Prove that no part in `changed` is reachable from the package's
/// relationships.
///
/// This is deliberately metadata-only: relationship closure is enough for the
/// projection dependency proof and avoids decoding opaque payloads while a
/// patch is being published.  Returning `false` means that the graph cannot
/// be proved safe to leave unparsed, either because a changed part is reachable
/// or because a malformed/dangling relationship prevents a complete proof.
/// The walk stops as soon as a changed part is found, which keeps the common
/// closure-affecting edit on the existing full-reparse path without paying for
/// the rest of the graph.
fn closure_is_disjoint(package: &OpcPackage, changed: &BTreeSet<String>) -> bool {
    let mut reachable = BTreeSet::new();
    let mut queue = Vec::new();
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let Ok(target) = relationship.target_partname() else {
            return false;
        };
        let key = target.as_str().to_ascii_lowercase();
        if changed.contains(&key) {
            return false;
        }
        if reachable.insert(key) {
            queue.push(target);
        }
    }

    while let Some(name) = queue.pop() {
        if changed.contains(&name.as_str().to_ascii_lowercase()) {
            return false;
        }
        let Some(part) = package.part_metadata(&name) else {
            // The source reader tolerates a dangling relationship and records
            // the repair.  A changed candidate must still go through the full
            // reparse so that the retained tolerance policy remains correct.
            return false;
        };
        for relationship in part.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let Ok(target) = relationship.target_partname() else {
                return false;
            };
            let key = target.as_str().to_ascii_lowercase();
            if reachable.insert(key) {
                queue.push(target);
            }
        }
    }
    true
}

/// How a changed commit publishes its staged graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Publication {
    /// Validate the staged graph as a complete workbook and keep the new
    /// projection (every part the workbook projection reads).
    Reparse,
    /// Publish without validating the workbook graph again: the change
    /// touches only parts the projection never reads.
    Unparsed,
}

/// A staged part-level edit of one [`Package`].
///
/// Starting one clones only the package owner (a reference-count increment);
/// the workbook projection is borrowed from that owner on the first operation
/// that needs it, and the OPC graph is copied on the first change only.
#[derive(Debug, Clone)]
pub(crate) struct Staged {
    source: Package,
    staged: Option<OpcPackage>,
    touched: BTreeSet<String>,
    /// Whether a whole-graph operation ran, so the commit compares every
    /// part instead of the touched ones.
    all: bool,
    publication: Publication,
}

impl Staged {
    /// Start staging over `package` without materializing its projection.
    pub(crate) fn of_package(package: &Package) -> Result<Self> {
        Ok(Self {
            source: package.clone(),
            staged: None,
            touched: BTreeSet::new(),
            all: false,
            publication: Publication::Reparse,
        })
    }

    /// Start staging a change that touches only parts the workbook
    /// projection never reads; the commit publishes without a reparse.
    pub(crate) fn unparsed(package: &Package) -> Result<Self> {
        let mut staged = Self::of_package(package)?;
        staged.publication = Publication::Unparsed;
        Ok(staged)
    }

    /// The source workbook projection, materialized only when needed.
    pub(crate) fn workbook(&self) -> Result<Arc<Workbook>> {
        self.source.projection()
    }

    /// The raw BIFF12 policy retained by the package being edited.
    pub(crate) fn raw_limits(&self) -> crate::raw::Limits {
        self.source.raw_limits_value()
    }

    /// The staged OPC graph.
    pub(crate) fn opc(&self) -> &OpcPackage {
        self.staged
            .as_ref()
            .unwrap_or_else(|| self.source.opc_package())
    }

    fn detach(&mut self) -> &mut OpcPackage {
        let source = &self.source;
        self.staged
            .get_or_insert_with(|| source.opc_package().clone())
    }

    /// The staged OPC graph for a change: refuses a signed package first.
    /// The caller records every part it changes with [`Self::touch`].
    pub(crate) fn opc_mut(&mut self) -> Result<&mut OpcPackage> {
        crate::signature_policy::ensure_unsigned(self.opc())?;
        Ok(self.detach())
    }

    pub(crate) fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    /// Run one whole-graph operation on a copy of the staged graph; the copy
    /// replaces the staged graph only when the operation succeeds, so an
    /// error leaves the staging as it was. The commit compares every part.
    /// The operation itself refuses a changing operation on a signed graph.
    pub(crate) fn update<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        self.all = true;
        Ok(value)
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.staged = None;
        self.touched.clear();
        self.all = false;
    }

    pub(crate) fn worksheet(&self, sheet: usize) -> Result<PackURI> {
        let workbook = self.workbook()?;
        worksheet_uri(&workbook, sheet)
    }

    pub(crate) fn blob(&self, name: &PackURI) -> Result<&[u8]> {
        Ok(self.opc().get_part(name)?.blob())
    }

    pub(crate) fn replace_blob(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        if self.opc().get_part(name)?.blob() == bytes.as_slice() {
            return Ok(false);
        }
        self.opc_mut()?.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        Ok(true)
    }

    pub(crate) fn replace_worksheet(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        let original = self.opc().get_part(name)?.blob_arc();
        if original.as_slice() == bytes.as_slice() {
            return Ok(false);
        }
        let index = crate::worksheet_index::find_index_target(self.opc(), name)?;
        let package = self.opc_mut()?;
        crate::worksheet_index::maintain(package, name, &original, &bytes)?;
        package.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        if let Some(index) = index {
            self.touch(&index);
        }
        Ok(true)
    }

    pub(crate) fn add_part(&mut self, part: BlobPart) -> Result<()> {
        let name = part.partname().clone();
        self.opc_mut()?.add_part(Box::new(part));
        self.touch(&name);
        Ok(())
    }

    pub(crate) fn remove_part(&mut self, name: &PackURI) -> Result<bool> {
        if self.opc().get_part(name).is_err() {
            return Ok(false);
        }
        let removed = self.opc_mut()?.remove_part(name);
        if removed {
            self.touch(name);
        }
        Ok(removed)
    }

    pub(crate) fn free_name(&self, prefix: &str, suffix: &str) -> Result<PackURI> {
        for index in 1_u32..=1_000_000 {
            let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))?;
            if self.opc().get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free part name is left"))
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging returns the source package. A changed staging of a signed
    /// package is refused with [`Error::Signed`] (ADR 0006).
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        self.commit_with_snapshot_limit(None)
    }

    /// Capture and publish with an exact native retained-payload ceiling.
    pub(crate) fn commit_with_snapshot_limit(
        self,
        max_snapshot_bytes: Option<usize>,
    ) -> Result<(Package, PartPatch)> {
        let Some(staged) = self.staged else {
            return Ok((self.source, PartPatch::default()));
        };
        let staged = Arc::new(staged);
        let patch = if self.all {
            PartPatch::capture_all(self.source.opc_shared(), &staged, max_snapshot_bytes)?
        } else {
            PartPatch::capture(
                self.source.opc_shared(),
                &staged,
                &self.touched,
                max_snapshot_bytes,
            )?
        };
        if patch.is_noop() {
            return Ok((self.source, patch));
        }
        crate::signature_policy::ensure_unsigned(&staged)?;
        let package = match self.publication {
            Publication::Reparse if patch.projection_unchanged(self.source.opc_package()) => {
                Workbook::validate_edit_candidate(&staged)?;
                self.source.with_unparsed(staged)
            },
            Publication::Reparse => self.source.reparse_candidate_retaining_projection(staged)?,
            Publication::Unparsed => self.source.with_unparsed(staged),
        };
        Ok((package, patch))
    }
}

pub(crate) fn worksheet_uri(workbook: &Workbook, sheet: usize) -> Result<PackURI> {
    let worksheet = workbook
        .worksheet_for_catalog_position(sheet)
        .ok_or_else(|| {
            if sheet < workbook.sheets().len() {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "edit",
                    "worksheet features belong to worksheets",
                )
            } else {
                Error::WorksheetNotFound(Some(sheet))
            }
        })?;
    workbook.worksheet_uri(worksheet)
}

/// The semantic worksheet features one feature pass decodes: everything
/// the full worksheet reader decodes after the cell table, without decoding
/// the cells.
pub(crate) struct FeaturePass {
    pub(crate) validations: Vec<crate::package::data_validation::Validation>,
    pub(crate) formats: Vec<crate::conditional_formatting::Formatting>,
}

/// Run the full reader's feature decoding over a worksheet stream, with the
/// formula context of complete-catalog position `sheet`.
pub(crate) fn feature_pass(
    bytes: &[u8],
    context: &crate::package::formula::Context,
    sheet: usize,
    raw_limits: crate::raw::Limits,
    leniency: crate::leniency::Leniency,
) -> Result<FeaturePass> {
    let context = context.for_sheet(sheet);
    let stream = crate::package::records::Stream::new(std::io::Cursor::new(bytes), raw_limits);
    let mut reader = crate::package::cells_reader::CellsReader::new(
        stream,
        &[],
        &context,
        usize::MAX,
        raw_limits,
        leniency,
    )?;
    reader.skip_to_features()?;
    Ok(FeaturePass {
        validations: reader.data_validations,
        formats: reader.conditional_formattings,
    })
}

impl Staged {
    /// The feature pass of the staged worksheet `name` at catalog position
    /// `sheet`.
    pub(crate) fn feature_pass(&self, name: &PackURI, sheet: usize) -> Result<FeaturePass> {
        let workbook = self.workbook()?;
        feature_pass(
            self.blob(name)?,
            &workbook.formula_context,
            sheet,
            workbook.raw_limits_value(),
            workbook.leniency(),
        )
    }
}

// ---------------------------------------------------------------------------
// Worksheet stream layout

/// The stream rank of a worksheet record after the cell table, in
/// [MS-XLSB] section 2.1.7.62 order. Authored units are inserted before the
/// first top-level unit of a higher rank.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum Rank {
    SheetCalcProp,
    SheetProtectionIso,
    SheetProtection,
    RangeProtectionIso,
    RangeProtection,
    ScenarioManager,
    AutoFilter,
    SortState,
    DataConsolidation,
    UserSheetViews,
    MergeCells,
    PhoneticInfo,
    ConditionalFormatting,
    DataValidations,
    Hyperlink,
    PrintOptions,
    Margins,
    PageSetup,
    HeaderFooter,
    RowBreaks,
    ColumnBreaks,
    BigName,
    CellWatches,
    IgnoredErrors,
    SmartTags,
    Drawing,
    LegacyDrawing,
    LegacyDrawingHeaderFooter,
    BackgroundPicture,
    OleObjects,
    ActiveXControls,
    WebPublishItems,
    ListParts,
    FutureRecords,
}

/// The collections (begin and end kinds) and single records the layout
/// knows, with their stream rank; `None` ranks an alternate-content block.
const BLOCKS: &[(Kind, Kind, Option<Rank>)] = &[
    (
        kind::BEGIN_SCEN_MAN,
        kind::END_SCEN_MAN,
        Some(Rank::ScenarioManager),
    ),
    (
        kind::BEGIN_A_FILTER,
        kind::END_A_FILTER,
        Some(Rank::AutoFilter),
    ),
    (
        kind::BEGIN_SORT_STATE,
        kind::END_SORT_STATE,
        Some(Rank::SortState),
    ),
    (
        kind::BEGIN_D_CON,
        kind::END_D_CON,
        Some(Rank::DataConsolidation),
    ),
    (
        kind::BEGIN_USER_SH_VIEWS,
        kind::END_USER_SH_VIEWS,
        Some(Rank::UserSheetViews),
    ),
    (
        kind::BEGIN_MERGE_CELLS,
        kind::END_MERGE_CELLS,
        Some(Rank::MergeCells),
    ),
    (
        kind::BEGIN_COND_FORMATTING,
        kind::END_COND_FORMATTING,
        Some(Rank::ConditionalFormatting),
    ),
    (
        kind::BEGIN_D_VALS,
        kind::END_D_VALS,
        Some(Rank::DataValidations),
    ),
    (
        kind::BEGIN_HEADER_FOOTER,
        kind::END_HEADER_FOOTER,
        Some(Rank::HeaderFooter),
    ),
    (kind::BEGIN_RW_BRK, kind::END_RW_BRK, Some(Rank::RowBreaks)),
    (
        kind::BEGIN_COL_BRK,
        kind::END_COL_BRK,
        Some(Rank::ColumnBreaks),
    ),
    (
        kind::BEGIN_CELL_WATCHES,
        kind::END_CELL_WATCHES,
        Some(Rank::CellWatches),
    ),
    (
        kind::BEGIN_CELL_IGNORE_ECS,
        kind::END_CELL_IGNORE_ECS,
        Some(Rank::IgnoredErrors),
    ),
    (
        kind::BEGIN_SMART_TAGS,
        kind::END_SMART_TAGS,
        Some(Rank::SmartTags),
    ),
    (
        kind::BEGIN_OLE_OBJECTS,
        kind::END_OLE_OBJECTS,
        Some(Rank::OleObjects),
    ),
    (
        kind::BEGIN_ACTIVE_X_CONTROLS,
        kind::END_ACTIVE_X_CONTROLS,
        Some(Rank::ActiveXControls),
    ),
    (
        kind::BEGIN_WEB_PUB_ITEMS,
        kind::END_WEB_PUB_ITEMS,
        Some(Rank::WebPublishItems),
    ),
    (
        kind::BEGIN_LIST_PARTS,
        kind::END_LIST_PARTS,
        Some(Rank::ListParts),
    ),
    (kind::FRT_BEGIN, kind::FRT_END, Some(Rank::FutureRecords)),
    (kind::AC_BEGIN, kind::AC_END, None),
];

const SINGLES: &[(Kind, Rank)] = &[
    (kind::SHEET_CALC_PROP, Rank::SheetCalcProp),
    (kind::SHEET_PROTECTION_ISO, Rank::SheetProtectionIso),
    (kind::SHEET_PROTECTION, Rank::SheetProtection),
    (kind::RANGE_PROTECTION_ISO, Rank::RangeProtectionIso),
    (kind::RANGE_PROTECTION, Rank::RangeProtection),
    (kind::PHONETIC_INFO, Rank::PhoneticInfo),
    (kind::H_LINK, Rank::Hyperlink),
    (kind::PRINT_OPTIONS, Rank::PrintOptions),
    (kind::MARGINS, Rank::Margins),
    (kind::PAGE_SETUP, Rank::PageSetup),
    (kind::BIG_NAME, Rank::BigName),
    (kind::DRAWING, Rank::Drawing),
    (kind::LEGACY_DRAWING, Rank::LegacyDrawing),
    (kind::LEGACY_DRAWING_HF, Rank::LegacyDrawingHeaderFooter),
    (kind::BK_HIM, Rank::BackgroundPicture),
];

/// One top-level unit after the cell table: a single record, or a begin/end
/// collection with every record it nests.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Unit {
    /// The kind of the unit's first record.
    pub(crate) kind: Kind,
    /// The unit's stream rank, when the layout knows its kind.
    pub(crate) rank: Option<Rank>,
    /// The unit's byte span.
    pub(crate) span: Range<usize>,
}

/// The top-level record units of a worksheet stream after its cell table.
#[derive(Debug, Clone)]
pub(crate) struct SheetLayout {
    pub(crate) units: Vec<Unit>,
    /// The offset of `BrtEndSheet`.
    pub(crate) end_sheet: usize,
}

impl SheetLayout {
    /// Scan a worksheet stream under the owning workbook's raw policy.
    pub(crate) fn scan_with_limits(data: &[u8], limits: crate::raw::Limits) -> Result<Self> {
        let mut records = Records::new(data, limits);
        let mut after_cells = false;
        let mut units = Vec::new();
        let mut stack: Vec<Kind> = Vec::new();
        let mut open: Option<(Kind, Option<Rank>, usize)> = None;
        while let Some(record) = records.next() {
            let record = record?;
            let start = record.offset();
            let end = records.offset();
            let kind = record.kind();
            if !after_cells {
                if kind == kind::END_SHEET_DATA {
                    after_cells = true;
                }
                continue;
            }
            if let Some(top) = stack.last().copied() {
                if kind == top {
                    stack.pop();
                    if stack.is_empty()
                        && let Some((first, rank, begin)) = open.take()
                    {
                        units.push(Unit {
                            kind: first,
                            rank,
                            span: begin..end,
                        });
                    }
                } else if let Some((_, close, _)) =
                    BLOCKS.iter().find(|(begin, _, _)| *begin == kind)
                {
                    stack.push(*close);
                } else if BLOCKS.iter().any(|(_, close, _)| *close == kind) {
                    return Err(invalid(
                        "worksheet record closes a collection it did not open",
                    ));
                }
                continue;
            }
            if kind == kind::END_SHEET {
                return Ok(Self {
                    units,
                    end_sheet: start,
                });
            }
            if let Some((_, close, rank)) = BLOCKS.iter().find(|(begin, _, _)| *begin == kind) {
                stack.push(*close);
                open = Some((kind, *rank, start));
                continue;
            }
            if BLOCKS.iter().any(|(_, close, _)| *close == kind) {
                return Err(invalid(
                    "worksheet record closes a collection it did not open",
                ));
            }
            let rank = SINGLES
                .iter()
                .find(|(single, _)| *single == kind)
                .map(|(_, rank)| *rank);
            units.push(Unit {
                kind,
                rank,
                span: start..end,
            });
        }
        if !after_cells {
            return Err(invalid("worksheet stream has no cell table"));
        }
        Err(invalid("worksheet stream has no BrtEndSheet"))
    }

    /// The units whose first record is `kind`, in stream order.
    pub(crate) fn units_of(&self, kind: Kind) -> impl Iterator<Item = &Unit> + '_ {
        self.units.iter().filter(move |unit| unit.kind == kind)
    }

    /// The offset a new unit of `rank` is inserted at: before the first
    /// top-level unit of a higher known rank, else before `BrtEndSheet`.
    pub(crate) fn insertion(&self, rank: Rank) -> usize {
        self.units
            .iter()
            .find(|unit| unit.rank.is_some_and(|other| other > rank))
            .map_or(self.end_sheet, |unit| unit.span.start)
    }
}

/// Replace `span` of `data` with `replacement`.
pub(crate) fn splice(data: &[u8], span: Range<usize>, replacement: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(data.len() - span.len() + replacement.len());
    out.extend_from_slice(&data[..span.start]);
    out.extend_from_slice(replacement);
    out.extend_from_slice(&data[span.end..]);
    out
}

/// Encode one record under an owning workbook's raw policy.
pub(crate) fn record_with_limits(
    kind: Kind,
    payload: &[u8],
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(payload.len() + 6);
    crate::raw::Encoder::new(&mut out, limits).write_record(kind, payload)?;
    Ok(out)
}

/// The `Patch`, `Commit` and `Diagnostics` types of one family, and its
/// `Error` re-export (ADR 0034 §3).
macro_rules! family {
    ($what:literal) => {
        pub use crate::package::PackageError as Error;

        crate::part_edit::family!($what, shared_error);
    };
    ($what:literal, shared_error) => {
        #[doc = concat!("The exact, reversible part-level change of a ", $what, " edit.")]
        #[derive(Debug, Clone, Default, PartialEq, Eq)]
        pub struct Patch(pub(crate) crate::part_edit::PartPatch);

        impl Patch {
            /// Whether the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The reverse patch.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }

            /// The payload bytes the patch retains (every before and after
            /// part image).
            #[must_use]
            pub fn retained_bytes(&self) -> usize {
                self.0.retained_bytes()
            }
        }

        #[doc = concat!("What a committed ", $what, " edit changed.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct Diagnostics {
            parts: usize,
        }

        impl Diagnostics {
            /// Whether the edit changed the package.
            #[must_use]
            pub const fn changed(&self) -> bool {
                self.parts > 0
            }

            /// The number of package parts the edit added, changed or
            /// removed.
            #[must_use]
            pub const fn parts(&self) -> usize {
                self.parts
            }
        }

        #[doc = concat!("The result of a ", $what, " edit's `commit`.")]
        #[derive(Debug)]
        pub struct Commit {
            package: crate::Package,
            patch: Patch,
            diagnostics: Diagnostics,
        }

        impl Commit {
            pub(crate) fn new(
                (package, patch): (crate::Package, crate::part_edit::PartPatch),
            ) -> Self {
                let diagnostics = Diagnostics {
                    parts: patch.parts(),
                };
                Self {
                    package,
                    patch: Patch(patch),
                    diagnostics,
                }
            }

            /// The package after the edit.
            #[must_use]
            pub const fn package(&self) -> &crate::Package {
                &self.package
            }

            /// Consume the commit into the package after the edit.
            #[must_use]
            pub fn into_package(self) -> crate::Package {
                self.package
            }

            /// The exact, reversible patch.
            #[must_use]
            pub const fn patch(&self) -> &Patch {
                &self.patch
            }

            /// What the edit changed.
            #[must_use]
            pub const fn diagnostics(&self) -> &Diagnostics {
                &self.diagnostics
            }

            /// Consume the commit into the package and the patch.
            #[must_use]
            pub fn into_parts(self) -> (crate::Package, Patch) {
                (self.package, self.patch)
            }
        }
    };
}

pub(crate) use family;

#[cfg(test)]
mod tests {
    use super::*;

    fn stream(records: &[(Kind, &[u8])]) -> Vec<u8> {
        let mut out = Vec::new();
        let mut encoder = crate::raw::Encoder::new(&mut out, crate::raw::Limits::DEFAULT);
        for (kind, payload) in records {
            encoder.write_record(*kind, payload).unwrap();
        }
        out
    }

    #[test]
    fn layout_splits_units_and_ranks_insertions() {
        let data = stream(&[
            (kind::BEGIN_SHEET, &[]),
            (kind::BEGIN_SHEET_DATA, &[]),
            (kind::ROW_HDR, &[0; 17]),
            (kind::END_SHEET_DATA, &[]),
            (kind::SHEET_PROTECTION, &[0; 66]),
            (kind::BEGIN_A_FILTER, &[0; 16]),
            (kind::BEGIN_FILTER_COLUMN, &[0; 6]),
            (kind::END_FILTER_COLUMN, &[]),
            (kind::END_A_FILTER, &[]),
            (kind::MARGINS, &[0; 48]),
            (kind::FRT_BEGIN, &[0; 4]),
            (kind::FRT_END, &[]),
            (kind::END_SHEET, &[]),
        ]);
        let layout = SheetLayout::scan_with_limits(&data, crate::raw::Limits::DEFAULT).unwrap();
        let kinds: Vec<_> = layout.units.iter().map(|unit| unit.kind).collect();
        assert_eq!(
            kinds,
            [
                kind::SHEET_PROTECTION,
                kind::BEGIN_A_FILTER,
                kind::MARGINS,
                kind::FRT_BEGIN
            ]
        );
        let margins = &layout.units[2];
        assert_eq!(layout.insertion(Rank::Hyperlink), margins.span.start);
        assert_eq!(
            layout.insertion(Rank::RowBreaks),
            layout.units[3].span.start
        );
        assert_eq!(
            layout.insertion(Rank::AutoFilter),
            layout.units[2].span.start
        );
        assert_eq!(
            layout.insertion(Rank::SheetCalcProp),
            layout.units[0].span.start
        );
    }

    #[test]
    fn layout_refuses_unbalanced_collections() {
        let data = stream(&[
            (kind::BEGIN_SHEET, &[]),
            (kind::BEGIN_SHEET_DATA, &[]),
            (kind::END_SHEET_DATA, &[]),
            (kind::END_D_VALS, &[]),
            (kind::END_SHEET, &[]),
        ]);
        assert!(SheetLayout::scan_with_limits(&data, crate::raw::Limits::DEFAULT).is_err());
    }

    #[test]
    fn staged_constructor_and_drop_do_not_materialize_cold_projection() {
        let cold = Package::from(Package::new().unwrap().into_opc());
        assert!(!cold.projection_cached());
        let staged = Staged::of_package(&cold).unwrap();
        drop(staged);
        assert!(!cold.projection_cached());
    }

    #[test]
    fn package_edit_constructors_and_drop_do_not_materialize_cold_projection() {
        let cold = Package::from(Package::new().unwrap().into_opc());
        assert!(!cold.projection_cached());

        drop(cold.edit_calculation_chain().unwrap());
        drop(cold.edit_volatile_dependencies().unwrap());
        drop(cold.edit_data_model().unwrap());
        drop(cold.edit_xml_maps().unwrap());
        drop(cold.edit_connections().unwrap());
        drop(cold.edit_shared_workbook().unwrap());
        drop(cold.edit_theme().unwrap());
        drop(cold.edit_custom_data().unwrap());
        drop(cold.edit_threaded_comments().unwrap());
        drop(cold.edit_task_panes().unwrap());
        drop(cold.edit_cell_values("missing").unwrap());

        assert!(!cold.projection_cached());
    }
}

/// The `EditOptions` of a family whose only configuration is its finite
/// semantic policy: `limits(L)` selects it and `limits_value()` reads it;
/// the default is the package's [`crate::SemanticLimits`] entry.
macro_rules! limits_options {
    ($limits:ty, $what:literal) => {
        #[doc = concat!("The configuration of a ", $what, " edit.")]
        #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
        pub struct EditOptions {
            limits: Option<$limits>,
        }

        impl EditOptions {
            #[doc = concat!("Select the finite policy of the ", $what, " edit; the default is the package's own [`crate::SemanticLimits`] entry.")]
            #[must_use]
            pub const fn limits(mut self, limits: $limits) -> Self {
                self.limits = Some(limits);
                self
            }

            /// The configured finite policy, if any.
            #[must_use]
            pub const fn limits_value(&self) -> Option<$limits> {
                self.limits
            }
        }
    };
}

pub(crate) use limits_options;

/// The native retained-snapshot option shared by the worksheet feature
/// families. Each family expands this into its own `EditOptions` type so the
/// limit enters at `Package::edit_x_with_options` and `Edit::commit()` keeps
/// the common family shape.
macro_rules! snapshot_options {
    ($what:literal) => {
        #[doc = concat!("The configuration of a ", $what, " edit.")]
        #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
        pub struct EditOptions {
            max_snapshot_bytes: Option<usize>,
        }

        impl EditOptions {
            /// Bound the exact before/after payload bytes retained by the
            /// native patch.
            #[must_use]
            pub const fn max_snapshot_bytes(mut self, maximum: usize) -> Self {
                self.max_snapshot_bytes = Some(maximum);
                self
            }

            /// The configured native retained-payload ceiling, if any.
            #[must_use]
            pub const fn max_snapshot_bytes_value(&self) -> Option<usize> {
                self.max_snapshot_bytes
            }
        }
    };
}

pub(crate) use snapshot_options;
