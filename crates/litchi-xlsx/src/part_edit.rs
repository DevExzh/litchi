//! The staged part-level engine behind the package-scoped edit families
//! (unified-ops records 0320 and 0360).
//!
//! [`Staged`] owns one immutable [`Package`] (a reference-count clone) and,
//! from the first change on, a staged copy of its OPC graph: starting an edit
//! copies nothing. A family stages whole-part replacements, additions and
//! removals, or runs typed graph operations with [`Staged::update`], and
//! commits: [`PartPatch`] captures the exact before/after state (content
//! type, payload and relationships) of each changed part and of the
//! package-level relationships, and an unchanged staging returns the source
//! package itself, so an exact no-op keeps every byte. The source package
//! never changes.
//!
//! The families built on it ([`crate::conditional_formatting::edit`],
//! [`crate::data_validation::edit`], [`crate::auto_filter::edit`],
//! [`crate::workbook::comments::edit`], [`crate::drawing::edit`],
//! [`crate::opc`] and the package-level families) expose the ADR 0034 shape
//! through [`family!`]; the families whose replay payload is their own record
//! patch use [`record_family!`].

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use crate::Package;
use crate::error::{Error, Result, invalid};
use crate::workbook::EditOptions;

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

/// Exact publication bytes without retaining the package's unrelated parts.
#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipImage {
    bytes: Arc<Vec<u8>>,
    member_present: bool,
    /// Canonical fallback bytes must not become retained source provenance
    /// merely because a patch was replayed.
    source: bool,
}

impl RelationshipImage {
    fn new(package: &OpcPackage, owner: &PackURI) -> Result<Self> {
        let (bytes, member_present) = match package.retained_relationships_xml(owner)? {
            Some((bytes, member_present)) => {
                return Ok(Self {
                    bytes,
                    member_present,
                    source: true,
                });
            },
            None => {
                let graph = if owner.as_str() == "/" {
                    package.rels()
                } else {
                    package.get_part(owner)?.rels()
                };
                (
                    package.authored_relationships_xml(owner, package.read_limits())?,
                    owner.as_str() == "/" || !graph.is_empty(),
                )
            },
        };
        Ok(Self {
            bytes,
            member_present,
            source: false,
        })
    }

    fn same_view(&self, other: &Self) -> bool {
        self.member_present == other.member_present && self.bytes == other.bytes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RootState {
    relationships: Vec<Relationship>,
    image: RelationshipImage,
}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
    /// Exact source or authored relationship-member XML for this owner.
    /// Semantic relationship equality alone cannot preserve lexical source
    /// bytes across a replay.
    relationship_xml: RelationshipImage,
}

impl PartialEq for State {
    fn eq(&self, other: &Self) -> bool {
        self.content_type == other.content_type
            && (Arc::ptr_eq(&self.blob, &other.blob) || self.blob == other.blob)
            && self.relationships == other.relationships
            && self.relationship_xml == other.relationship_xml
    }
}

impl Eq for State {}

impl State {
    fn read(package: &OpcPackage, part: &dyn Part) -> Result<Self> {
        Ok(Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: relationships(part.rels()),
            relationship_xml: RelationshipImage::new(package, part.partname())?,
        })
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Result<Option<Self>> {
        package
            .get_part(name)
            .ok()
            .map(|part| Self::read(package, part))
            .transpose()
    }

    /// Whether `part` holds exactly this state, without copying its payload.
    fn matches(&self, package: &OpcPackage, part: &dyn Part) -> Result<bool> {
        if self.content_type != part.content_type() {
            return Ok(false);
        }
        let blob = part.blob();
        if !std::ptr::eq(self.blob.as_slice(), blob) && self.blob.as_slice() != blob {
            return Ok(false);
        }
        if self.relationships != relationships(part.rels()) {
            return Ok(false);
        }
        let current = RelationshipImage::new(package, part.partname())?;
        Ok(if self.relationship_xml.source {
            current == self.relationship_xml
        } else {
            current.same_view(&self.relationship_xml)
        })
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        add_relationships(part.rels_mut(), &self.relationships)?;
        Ok(part)
    }
}

fn same_part(
    before_package: &OpcPackage,
    before: &dyn Part,
    after_package: &OpcPackage,
    after: &dyn Part,
) -> Result<bool> {
    Ok(before.content_type() == after.content_type()
        && {
            let (left, right) = (before.blob(), after.blob());
            std::ptr::eq(left, right) || left == right
        }
        && relationships_equal(before.rels(), after.rels())
        // The retained-token comparison is pointer-fast for the usual clone
        // case.  Only a changed token falls through to source capture below.
        && before_package.same_retained_relationships(after_package, before.partname())?)
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

/// Compact manifest state. Canonical structural output is reconstructible
/// from the typed graph. Only an explicit lexical patch needs an exact digest
/// of a canonical base; producer source bytes always remain exact images.
#[derive(Debug, Clone, PartialEq, Eq)]
struct ContentTypesImage {
    retained: Option<Arc<Vec<u8>>>,
    exact_canonical: Option<[u8; 32]>,
}

impl ContentTypesImage {
    fn new(package: &OpcPackage) -> Result<Self> {
        Ok(Self {
            retained: package.retained_content_types_xml()?,
            exact_canonical: None,
        })
    }

    fn bind_exact(&mut self, package: &OpcPackage) -> Result<()> {
        use sha2::Digest as _;
        if self.retained.is_none() {
            let bytes = package.authored_content_types_xml(package.read_limits())?;
            self.exact_canonical = Some(sha2::Sha256::digest(&*bytes).into());
        }
        Ok(())
    }

    fn publication_bytes(package: &OpcPackage) -> Result<Arc<Vec<u8>>> {
        match package.retained_content_types_xml()? {
            Some(bytes) => Ok(bytes),
            None => Ok(package.authored_content_types_xml(package.read_limits())?),
        }
    }

    fn matches_view(&self, package: &OpcPackage) -> Result<bool> {
        use sha2::Digest as _;
        if let Some(expected) = &self.retained {
            return Ok(*expected == Self::publication_bytes(package)?);
        }
        if let Some(expected) = self.exact_canonical {
            let bytes = Self::publication_bytes(package)?;
            let digest: [u8; 32] = sha2::Sha256::digest(&*bytes).into();
            return Ok(digest == expected);
        }
        // A structural patch composes with unrelated typed graph changes.
        // A reopened canonical artifact is equivalent to an authored view.
        match package.retained_content_types_xml()? {
            None => Ok(true),
            Some(bytes) => Ok(bytes == package.authored_content_types_xml(package.read_limits())?),
        }
    }

    fn bytes_for(&self, package: &OpcPackage) -> Result<Arc<Vec<u8>>> {
        use sha2::Digest as _;
        if let Some(bytes) = &self.retained {
            return Ok(Arc::clone(bytes));
        }
        let bytes = package.authored_content_types_xml(package.read_limits())?;
        if let Some(expected) = self.exact_canonical {
            let digest: [u8; 32] = sha2::Sha256::digest(&*bytes).into();
            if digest != expected {
                return Err(Error::PatchConflict {
                    part: "[Content_Types].xml".to_owned(),
                });
            }
        }
        Ok(bytes)
    }
}

/// The exact, reversible part-level change of one staged edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
    root: Option<(RootState, RootState)>,
    content_types: Option<(ContentTypesImage, ContentTypesImage)>,
    exact_content_types: bool,
}

/// Metadata-only admission for one part patch.  Whole-graph callers retain
/// the changed names so capture can materialize only those deltas after the
/// census, rather than walking every part a second time.
#[derive(Debug, Default)]
struct RetainedEstimate {
    bytes: usize,
    changed_names: BTreeSet<String>,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none() && self.content_types.is_none()
    }

    /// The number of parts (and the package-level relationship set) the
    /// patch changes.
    pub(crate) fn parts(&self) -> usize {
        let changed = self.deltas.len() + usize::from(self.root.is_some());
        if changed == 0 {
            usize::from(self.content_types.is_some())
        } else {
            changed
        }
    }

    /// The semantic payload retained by this patch.
    ///
    /// This charge covers the exact part images, relationship fields and
    /// manifest bytes held by the reversible patch. It deliberately excludes
    /// allocator overhead and parser/cache state; callers charge package
    /// endpoints separately when they retain them alongside this patch.
    pub(crate) fn retained_bytes(&self) -> usize {
        // These fields are canonical strings copied with `str::to_owned` by
        // `State::read`; charge their logical bytes so the metadata-only
        // estimator below has the same definition without rebuilding them.
        // Blob and exact relationship-member images retain their Vec
        // capacities because those buffers are the native patch payload.
        fn relationship_bytes(relationship: &Relationship) -> usize {
            relationship
                .id
                .len()
                .saturating_add(relationship.kind.len())
                .saturating_add(relationship.target.len())
        }

        fn image_bytes(image: &RelationshipImage) -> usize {
            image.bytes.capacity()
        }

        fn state_bytes(state: &State) -> usize {
            state
                .content_type
                .len()
                .saturating_add(state.blob.capacity())
                .saturating_add(
                    state
                        .relationships
                        .iter()
                        .map(relationship_bytes)
                        .sum::<usize>(),
                )
                .saturating_add(image_bytes(&state.relationship_xml))
        }

        fn root_bytes(root: &RootState) -> usize {
            root.relationships
                .iter()
                .map(relationship_bytes)
                .sum::<usize>()
                .saturating_add(image_bytes(&root.image))
        }

        fn content_types_bytes(image: &ContentTypesImage) -> usize {
            image.retained.as_ref().map_or(0, |bytes| bytes.capacity())
        }

        self.deltas
            .iter()
            .fold(0usize, |total, delta| {
                total
                    .saturating_add(delta.name.as_str().len())
                    .saturating_add(delta.before.as_ref().map_or(0, state_bytes))
                    .saturating_add(delta.after.as_ref().map_or(0, state_bytes))
            })
            .saturating_add(self.root.as_ref().map_or(0, |(before, after)| {
                root_bytes(before).saturating_add(root_bytes(after))
            }))
            .saturating_add(self.content_types.as_ref().map_or(0, |(before, after)| {
                content_types_bytes(before).saturating_add(content_types_bytes(after))
            }))
    }

    /// Estimate the patch payload without constructing any patch-owned state.
    /// The estimate follows the same semantic fields as [`Self::retained_bytes`]
    /// and is used as the admission step immediately before capture.
    fn estimate_retained_bytes(
        before: &OpcPackage,
        after: &OpcPackage,
        names: &BTreeSet<String>,
        all: bool,
    ) -> Result<RetainedEstimate> {
        fn relationship_bytes(relationship: &litchi_opc::Relationship) -> usize {
            relationship
                .r_id()
                .len()
                .saturating_add(relationship.reltype().len())
                .saturating_add(relationship.target_ref().len())
        }

        fn relationship_image_bytes(package: &OpcPackage, owner: &PackURI) -> Result<usize> {
            if let Some((bytes, _)) = package.retained_relationships_xml(owner)? {
                return Ok(bytes.capacity());
            }
            // The OPC relationship planner exposes the exact canonical member
            // length without materializing the XML. This keeps admission
            // bounded by existing graph metadata and avoids both a second
            // source copy and an archive-size-based false refusal.
            Ok(package
                .plan_source_relationships(owner, package.read_limits())?
                .final_len())
        }

        fn state_bytes(package: &OpcPackage, part: &dyn Part) -> Result<usize> {
            Ok(part
                .content_type()
                .len()
                .saturating_add(part.blob_arc().capacity())
                .saturating_add(part.rels().iter().map(relationship_bytes).sum::<usize>())
                .saturating_add(relationship_image_bytes(package, part.partname())?))
        }

        fn optional_part<'a>(
            package: &'a OpcPackage,
            name: &PackURI,
        ) -> Result<Option<&'a dyn Part>> {
            match package.get_part(name) {
                Ok(part) => Ok(Some(part)),
                Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
                Err(error) => Err(error.into()),
            }
        }

        let mut observed = 0usize;
        let mut changed_names = BTreeSet::new();
        let mut changes_types = false;
        {
            let mut add_delta = |name: &PackURI,
                                 before_part: Option<&dyn Part>,
                                 after_part: Option<&dyn Part>|
             -> Result<()> {
                let changed = match (before_part, after_part) {
                    (Some(before_part), Some(after_part)) => {
                        !same_part(before, before_part, after, after_part)?
                    },
                    (None, None) => false,
                    _ => true,
                };
                if changed {
                    changed_names.insert(name.as_str().to_owned());
                    if match (before_part, after_part) {
                        (Some(before_part), Some(after_part)) => {
                            before_part.content_type() != after_part.content_type()
                        },
                        (None, None) => false,
                        _ => true,
                    } {
                        changes_types = true;
                    }
                    observed = observed
                        .saturating_add(name.as_str().len())
                        .saturating_add(
                            before_part
                                .map(|part| state_bytes(before, part))
                                .transpose()?
                                .unwrap_or(0),
                        )
                        .saturating_add(
                            after_part
                                .map(|part| state_bytes(after, part))
                                .transpose()?
                                .unwrap_or(0),
                        );
                }
                Ok(())
            };

            if all {
                for part in before.try_iter_parts() {
                    let part = part?;
                    add_delta(
                        part.partname(),
                        Some(part),
                        optional_part(after, part.partname())?,
                    )?;
                }
                for part in after.try_iter_parts() {
                    let part = part?;
                    if optional_part(before, part.partname())?.is_none() {
                        add_delta(part.partname(), None, Some(part))?;
                    }
                }
            } else {
                for name in names {
                    let name =
                        PackURI::new(name.as_str()).map_err(|error| invalid(error.to_string()))?;
                    add_delta(
                        &name,
                        optional_part(before, &name)?,
                        optional_part(after, &name)?,
                    )?;
                }
            }
        }

        let root = PackURI::new("/").map_err(|error| invalid(error.to_string()))?;
        if !before.same_retained_relationships(after, &root)?
            || before.rels().len() != after.rels().len()
            || before.rels().iter().any(|relationship| {
                after
                    .rels()
                    .get(relationship.r_id())
                    .is_none_or(|candidate| {
                        candidate.reltype() != relationship.reltype()
                            || candidate.target_ref() != relationship.target_ref()
                            || candidate.is_external() != relationship.is_external()
                    })
            })
        {
            let before_fields = before.rels().iter().map(relationship_bytes).sum::<usize>();
            let after_fields = after.rels().iter().map(relationship_bytes).sum::<usize>();
            observed = observed
                .saturating_add(before_fields)
                .saturating_add(after_fields)
                .saturating_add(relationship_image_bytes(before, &root)?)
                .saturating_add(relationship_image_bytes(after, &root)?);
        }

        // Mirror `with_manifests`: differing provenance alone does not retain
        // a manifest pair when the old image already matches the resulting
        // publication view. Charge only buffers the actual patch retains.
        if changes_types || !before.same_retained_content_types(after)? {
            let before_image = ContentTypesImage::new(before)?;
            if changes_types || !before_image.matches_view(after)? {
                let after_image = ContentTypesImage::new(after)?;
                observed = observed
                    .saturating_add(
                        before_image
                            .retained
                            .as_ref()
                            .map_or(0, |bytes| bytes.capacity()),
                    )
                    .saturating_add(
                        after_image
                            .retained
                            .as_ref()
                            .map_or(0, |bytes| bytes.capacity()),
                    );
            }
        }
        Ok(RetainedEstimate {
            bytes: observed,
            changed_names,
        })
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
            exact_content_types: self.exact_content_types,
            content_types: self
                .content_types
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
        }
    }

    /// Attach the manifest views of both sides to a changed patch.
    fn with_manifests(mut self, before: &OpcPackage, after: &OpcPackage) -> Result<Self> {
        let changes_types = self
            .deltas
            .iter()
            .any(|delta| match (&delta.before, &delta.after) {
                (Some(before), Some(after)) => before.content_type != after.content_type,
                _ => true,
            });
        let mut content_types = if changes_types || !before.same_retained_content_types(after)? {
            let pair = (
                ContentTypesImage::new(before)?,
                ContentTypesImage::new(after)?,
            );
            (changes_types || !pair.0.matches_view(after)?).then_some(pair)
        } else {
            None
        };
        let exact_content_types = content_types.as_ref().is_some_and(|(_, image)| {
            image.retained.is_some()
                && before.content_types_source_bytes() != after.content_types_source_bytes()
        });
        if exact_content_types && let Some((old, new)) = &mut content_types {
            old.bind_exact(before)?;
            new.bind_exact(after)?;
        }
        self.content_types = content_types;
        self.exact_content_types = exact_content_types;
        Ok(self)
    }

    fn root(before: &OpcPackage, after: &OpcPackage) -> Result<Option<(RootState, RootState)>> {
        let owner = PackURI::new("/").map_err(|error| invalid(error.to_string()))?;
        if relationships_equal(before.rels(), after.rels())
            && before.same_retained_relationships(after, &owner)?
        {
            Ok(None)
        } else {
            let pair = (
                RootState {
                    relationships: relationships(before.rels()),
                    image: RelationshipImage::new(before, &owner)?,
                },
                RootState {
                    relationships: relationships(after.rels()),
                    image: RelationshipImage::new(after, &owner)?,
                },
            );
            Ok((pair.0 != pair.1).then_some(pair))
        }
    }

    /// Capture the parts named in `names` that differ between `before` and
    /// `after`.
    pub(crate) fn capture(
        before: &OpcPackage,
        after: &OpcPackage,
        names: &BTreeSet<String>,
    ) -> Result<Self> {
        let mut deltas = Vec::new();
        for name in names {
            let name = PackURI::new(name.as_str()).map_err(|error| invalid(error.to_string()))?;
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
        Self {
            deltas,
            root: Self::root(before, after)?,
            content_types: None,
            exact_content_types: false,
        }
        .with_manifests(before, after)
    }

    /// Capture every part that differs between `before` and `after`.
    ///
    /// The staged graph is a clone of the source, so an untouched part shares
    /// its payload allocation and compares by pointer; only changed parts are
    /// compared by value.
    pub(crate) fn capture_all(before: &OpcPackage, after: &OpcPackage) -> Result<Self> {
        // One commit changes at most as many parts as the package's source
        // limits admit.
        let maximum = before.read_limits().max_parts();
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
        for part in before.try_iter_parts() {
            let part = part?;
            let name = part.partname();
            match optional_part(after, name)? {
                Some(other) if same_part(before, part, after, other)? => {},
                other => push(Delta {
                    name: name.clone(),
                    before: Some(State::read(before, part)?),
                    after: other.map(|part| State::read(after, part)).transpose()?,
                })?,
            }
        }
        for part in after.try_iter_parts() {
            let part = part?;
            if optional_part(before, part.partname())?.is_none() {
                push(Delta {
                    name: part.partname().clone(),
                    before: None,
                    after: Some(State::read(after, part)?),
                })?;
            }
        }
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        Self {
            deltas,
            root: Self::root(before, after)?,
            content_types: None,
            exact_content_types: false,
        }
        .with_manifests(before, after)
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package, operation: &'static str) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        let target = self.apply_graph(package.opc())?;
        package.derive(target)
    }

    /// Replay onto a copy of the OPC graph `source`, refusing a graph that is
    /// not the patch's before state; `source` is unchanged.
    pub(crate) fn apply_graph(&self, source: &OpcPackage) -> Result<OpcPackage> {
        if let Some((before, _)) = &self.root {
            let owner = PackURI::new("/").map_err(|error| invalid(error.to_string()))?;
            if relationships(source.rels()) != before.relationships
                || (before.image.source && RelationshipImage::new(source, &owner)? != before.image)
            {
                return Err(Error::PatchConflict {
                    part: "/_rels/.rels".to_owned(),
                });
            }
        }
        if self.exact_content_types
            && let Some((before, _)) = &self.content_types
            && !before.matches_view(source)?
        {
            return Err(Error::PatchConflict {
                part: "[Content_Types].xml".to_owned(),
            });
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
                    part: delta.name.as_str().to_owned(),
                });
            }
        }
        crate::signature_policy::ensure_unsigned(source)?;
        let mut target = source.clone();
        for delta in &self.deltas {
            if let (Some(before), Some(after)) = (&delta.before, &delta.after)
                && before.content_type == after.content_type
                && before.relationship_xml == after.relationship_xml
            {
                target
                    .get_part_mut(&delta.name)?
                    .set_blob_shared(Arc::clone(&after.blob));
                continue;
            }
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
                let limits = target.read_limits();
                let current = RelationshipImage::new(&target, &delta.name)?;
                if after.relationship_xml.source && current != after.relationship_xml {
                    let expected = target.source_relationships(&delta.name, limits)?;
                    target.try_replace_relationships_bytes(
                        &expected,
                        Arc::clone(&after.relationship_xml.bytes),
                        after.relationship_xml.member_present,
                        limits,
                    )?;
                }
            }
        }
        if let Some((before, after)) = &self.root {
            let graph = target.rels_mut();
            for relationship in &before.relationships {
                graph.remove(&relationship.id);
            }
            add_relationships(graph, &after.relationships)?;
            if after.image.source {
                let limits = target.read_limits();
                let owner = PackURI::new("/").map_err(|error| invalid(error.to_string()))?;
                let current = target.source_relationships(&owner, limits)?;
                target.try_replace_relationships_bytes(
                    &current,
                    Arc::clone(&after.image.bytes),
                    after.image.member_present,
                    limits,
                )?;
            }
        }
        if let Some((before, after)) = &self.content_types {
            let limits = target.read_limits();
            // Ordinary structural patches compose with unrelated changes.
            // An explicit raw manifest replacement has an exact byte guard
            // in both directions, even when parts changed in the same edit.
            let needs_restore = self.exact_content_types
                || after.retained.as_ref().is_some_and(|expected| {
                    target.content_types_source_bytes().as_ref() != Some(expected)
                });
            if needs_restore {
                if before.matches_view(source)? {
                    let current = ContentTypesImage::publication_bytes(&target)?;
                    let replacement = after.bytes_for(&target)?;
                    if current != replacement {
                        target.try_replace_content_types_bytes(&current, replacement, limits)?;
                    }
                } else if self.exact_content_types {
                    return Err(Error::PatchConflict {
                        part: "[Content_Types].xml".to_owned(),
                    });
                }
            }
        }
        Ok(target)
    }
}

/// A staged part-level edit of one [`Package`].
///
/// Starting one clones the package (a reference-count increment); the OPC
/// graph is copied on the first change only.
#[derive(Debug, Clone)]
pub(crate) struct Staged {
    operation: &'static str,
    source: Package,
    staged: Option<OpcPackage>,
    touched: BTreeSet<String>,
    max_retained_bytes: Option<usize>,
    /// Whether a whole-graph operation ran, so the commit compares every
    /// part instead of the touched ones.
    all: bool,
}

impl Staged {
    /// Start staging; refuses packages whose mutations are not allowed.
    pub(crate) fn new(package: &Package, operation: &'static str) -> Result<Self> {
        Self::new_with_options(package, operation, EditOptions::default())
    }

    /// Start staging with an explicit patch-retention ceiling.
    pub(crate) fn new_with_options(
        package: &Package,
        operation: &'static str,
        options: EditOptions,
    ) -> Result<Self> {
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        Ok(Self {
            operation,
            source: package.clone(),
            staged: None,
            touched: BTreeSet::new(),
            max_retained_bytes: options.max_retained_bytes_value(),
            all: false,
        })
    }

    /// The staged OPC graph.
    pub(crate) fn opc(&self) -> &OpcPackage {
        self.staged.as_ref().unwrap_or_else(|| self.source.opc())
    }

    fn detach(&mut self) -> &mut OpcPackage {
        let source = &self.source;
        self.staged.get_or_insert_with(|| source.opc().clone())
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

    /// Run one operation that changes at most the part `name` on a copy of
    /// the staged graph; the copy replaces the staged graph only when the
    /// operation succeeds. The commit compares only the touched parts.
    pub(crate) fn update_part<T>(
        &mut self,
        name: &PackURI,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        self.touch(name);
        Ok(value)
    }

    /// Whether no change detached the staged graph yet.
    #[cfg(test)]
    pub(crate) const fn staged_is_shared(&self) -> bool {
        self.staged.is_none()
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.staged = None;
        self.touched.clear();
        self.all = false;
    }

    /// The part name of the worksheet at catalog position `sheet`; refuses a
    /// missing sheet and a chart or dialog sheet.
    pub(crate) fn worksheet(&self, sheet: usize) -> Result<PackURI> {
        let worksheet = self
            .source
            .workbook_ref()
            .sheet(sheet)?
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                self.operation,
                "the worksheet does not exist",
            ))?;
        if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: worksheet.name().to_owned(),
            });
        }
        Ok(worksheet.part_uri().clone())
    }

    /// The staged payload of `name`.
    pub(crate) fn blob(&self, name: &PackURI) -> Result<&[u8]> {
        Ok(self.opc().get_part(name)?.blob())
    }

    /// Replace the payload of an existing part; identical bytes are a no-op.
    pub(crate) fn replace_blob(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        if self.opc().get_part(name)?.blob() == bytes.as_slice() {
            return Ok(false);
        }
        self.opc_mut()?.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        Ok(true)
    }

    /// The first free part name `{prefix}{N}{suffix}`.
    pub(crate) fn free_name(&self, prefix: &str, suffix: &str) -> Result<PackURI> {
        for index in 1_u32..=1_000_000 {
            let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
                .map_err(|error| invalid(error.to_string()))?;
            if self.opc().get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free part name is left"))
    }

    /// The staged styles part, when the workbook has one.
    pub(crate) fn styles(&self) -> Option<PackURI> {
        self.source.workbook_ref().styles_part_uri().cloned()
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging returns the source package. A changed staging of a signed
    /// package is refused with [`Error::Signed`] (ADR 0006).
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        let Some(staged) = self.staged else {
            return Ok((self.source, PartPatch::default()));
        };
        let source = self.source;
        let mut endpoint_charge = None;
        let mut estimated_names = None;
        let target = if let Some(maximum) = self.max_retained_bytes {
            let estimate =
                PartPatch::estimate_retained_bytes(source.opc(), &staged, &self.touched, self.all)?;
            // Build the exact target owner once, before patch capture, so the
            // admission charge covers the real source and target endpoints
            // rather than assuming the edit preserves package size. A zero
            // retained-patch estimate is the no-op path: keep the original
            // owner so its retained source image is charged exactly as the
            // package returned below.
            let target = if estimate.bytes == 0 {
                source.clone()
            } else {
                source.derive(staged.clone())?
            };
            let source_bytes = source.retained_bytes();
            let target_bytes = target.retained_bytes();
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(estimate.bytes);
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX part edit retained package".into(),
                }));
            }
            endpoint_charge = Some((source_bytes, target_bytes));
            if self.all {
                estimated_names = Some(estimate.changed_names);
            }
            Some(target)
        } else {
            None
        };
        let patch = if let Some(names) = estimated_names {
            PartPatch::capture(source.opc(), &staged, &names)?
        } else if self.all {
            PartPatch::capture_all(source.opc(), &staged)?
        } else {
            PartPatch::capture(source.opc(), &staged, &self.touched)?
        };
        if patch.is_noop() {
            return Ok((source, patch));
        }
        crate::signature_policy::ensure_unsigned(&staged)?;
        let package = match target {
            Some(target) => target,
            None => source.derive(staged)?,
        };
        if let Some(maximum) = self.max_retained_bytes {
            let (source_bytes, target_bytes) =
                endpoint_charge.expect("retained endpoint charge established with the budget");
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(patch.retained_bytes());
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX part edit retained package".into(),
                }));
            }
        }
        Ok((package, patch))
    }
}

#[cfg(test)]
mod tests {
    use super::PartPatch;
    use litchi_opc::{BlobPart, OpcPackage, PackURI, TargetMode};
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::sync::Arc;

    const CONTENT_TYPES_A: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
"#;

    const CONTENT_TYPES_B: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/xml" Extension="xml"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/><Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/></Types>"#;

    const ROOT_RELS_A: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    const ROOT_RELS_B: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Target="xl/workbook.xml" TargetMode="Internal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/></Relationships>"#;

    fn package(content_types: &[u8], root_relationships: &[u8]) -> OpcPackage {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("[Content_Types].xml", content_types)
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        writer
            .write_stored(
                "xl/workbook.xml",
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            )
            .unwrap();
        OpcPackage::from_bytes(
            &writer.finish_to_bytes().unwrap(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap()
    }

    #[test]
    fn lexical_publication_manifests_are_replayed_and_inverted_exactly() {
        let before = package(CONTENT_TYPES_A, ROOT_RELS_A);
        let after = package(CONTENT_TYPES_B, ROOT_RELS_B);
        let patch = PartPatch::capture_all(&before, &after).unwrap();

        assert!(!patch.is_noop());
        assert_eq!(patch.parts(), 1);

        let forward = patch.apply_graph(&before).unwrap();
        assert_eq!(
            forward
                .source_content_types(litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes(),
            after
                .source_content_types(litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes()
        );
        let root = PackURI::new("/").unwrap();
        assert_eq!(
            forward
                .source_relationships(&root, litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes(),
            after
                .source_relationships(&root, litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes()
        );

        let restored = patch.inverse().apply_graph(&forward).unwrap();
        assert_eq!(
            restored
                .source_content_types(litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes(),
            before
                .source_content_types(litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes()
        );
        assert_eq!(
            restored
                .source_relationships(&root, litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes(),
            before
                .source_relationships(&root, litchi_opc::SourceLimits::default())
                .unwrap()
                .bytes()
        );
    }

    #[test]
    fn relationship_patch_keeps_unrelated_content_type_additions() {
        let before = package(CONTENT_TYPES_A, ROOT_RELS_A);
        let mut after = before.clone();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        after
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test".to_owned(),
                "https://example.test/".to_owned(),
                "rIdExternal".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let patch = PartPatch::capture_all(&before, &after).unwrap();

        let mut composed = before.clone();
        composed.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/octet-stream".to_owned(),
            vec![1, 2, 3],
        )));
        let forward = patch.apply_graph(&composed).unwrap();
        assert!(
            forward
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
        assert!(
            forward
                .get_part(&workbook)
                .unwrap()
                .rels()
                .get("rIdExternal")
                .is_some()
        );

        let restored = patch.inverse().apply_graph(&forward).unwrap();
        assert!(
            restored
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
        assert!(
            restored
                .get_part(&workbook)
                .unwrap()
                .rels()
                .get("rIdExternal")
                .is_none()
        );
    }

    #[test]
    fn structural_content_type_patch_composes_with_unrelated_parts_both_directions() {
        let before = package(CONTENT_TYPES_A, ROOT_RELS_A);
        let mut after = before.clone();
        after.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/changed.bin").unwrap(),
            "application/x-litchi-changed".to_owned(),
            vec![4, 5, 6],
        )));
        let patch = PartPatch::capture_all(&before, &after).unwrap();

        let mut composed = before.clone();
        composed.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/x-litchi-unrelated".to_owned(),
            vec![7, 8, 9],
        )));
        let forward = patch.apply_graph(&composed).unwrap();
        assert!(
            forward
                .get_part(&PackURI::new("/custom/changed.bin").unwrap())
                .is_ok()
        );
        assert!(
            forward
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );

        let restored = patch.inverse().apply_graph(&forward).unwrap();
        assert!(
            restored
                .get_part(&PackURI::new("/custom/changed.bin").unwrap())
                .is_err()
        );
        assert!(
            restored
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn mixed_lexical_and_structural_content_types_are_exact_both_directions() -> crate::Result<()> {
        let before = package(CONTENT_TYPES_A, ROOT_RELS_A);
        let mut after = before.clone();
        let changed = PackURI::new("/custom/changed.bin").unwrap();
        after.add_part(Box::new(BlobPart::new(
            changed.clone(),
            "application/x-litchi-changed".to_owned(),
            vec![4, 5, 6],
        )));
        let current = after
            .source_content_types(after.read_limits())?
            .bytes()
            .to_vec();
        let lexical = current
            .windows(b"</Types>".len())
            .position(|window| window == b"</Types>")
            .map(|at| {
                let mut bytes = current.clone();
                bytes.insert(at, b'\n');
                bytes
            })
            .expect("authored content-types root");
        after.try_replace_content_types_bytes(&current, Arc::new(lexical), after.read_limits())?;
        let patch = PartPatch::capture_all(&before, &after).unwrap();

        let mut unrelated = before.clone();
        unrelated.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/x-litchi-unrelated".to_owned(),
            vec![7, 8, 9],
        )));
        assert!(matches!(
            patch.apply_graph(&unrelated),
            Err(crate::Error::PatchConflict { .. })
        ));

        let forward = patch.apply_graph(&before).unwrap();
        assert_eq!(
            forward
                .source_content_types(forward.read_limits())
                .unwrap()
                .bytes(),
            after
                .source_content_types(after.read_limits())
                .unwrap()
                .bytes()
        );
        let restored = patch.inverse().apply_graph(&forward).unwrap();
        assert_eq!(
            restored
                .source_content_types(restored.read_limits())
                .unwrap()
                .bytes(),
            before
                .source_content_types(before.read_limits())
                .unwrap()
                .bytes()
        );
        assert!(restored.get_part(&changed).is_err());
        Ok(())
    }
}

/// The direct children of `worksheet`, in ECMA-376 `CT_Worksheet` order.
const WORKSHEET_CHILDREN: &[&[u8]] = &[
    b"sheetPr",
    b"dimension",
    b"sheetViews",
    b"sheetFormatPr",
    b"cols",
    b"sheetData",
    b"sheetCalcPr",
    b"sheetProtection",
    b"protectedRanges",
    b"scenarios",
    b"autoFilter",
    b"sortState",
    b"dataConsolidate",
    b"customSheetViews",
    b"mergeCells",
    b"phoneticPr",
    b"conditionalFormatting",
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];

const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_MAIN_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

/// One direct child of a worksheet root: its local name, its byte span and
/// its `r:id` (any prefix bound to a relationships namespace).
#[derive(Debug, Clone)]
pub(crate) struct WorksheetChild {
    pub(crate) name: Vec<u8>,
    pub(crate) span: std::ops::Range<usize>,
    pub(crate) relationship_id: Option<String>,
}

/// The layout of a worksheet part's direct children.
#[derive(Debug)]
pub(crate) struct WorksheetLayout {
    /// The root element's prefix with its colon (`x:`), or empty.
    pub(crate) prefix: String,
    /// Whether the root binds the Strict main namespace.
    pub(crate) strict: bool,
    pub(crate) children: Vec<WorksheetChild>,
    /// The offset of the root's end tag.
    pub(crate) root_close: usize,
}

impl WorksheetLayout {
    /// Scan a worksheet part. Markup-compatibility branches among the direct
    /// children are refused: their effective content cannot be spliced.
    pub(crate) fn scan(xml: &[u8]) -> Result<Self> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut prefix = String::new();
        let mut strict = false;
        let mut children = Vec::new();
        let mut open: Option<(Vec<u8>, usize, Option<String>)> = None;
        loop {
            let start = position(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| invalid(format!("worksheet markup: {error}")))?;
            let end = position(&reader)?;
            let (element, empty) = match event {
                Event::Start(element) => (Some(element), false),
                Event::Empty(element) => (Some(element), true),
                other => {
                    match other {
                        Event::End(_) => {
                            depth = depth.saturating_sub(1);
                            if depth == 1
                                && let Some((name, begin, relationship_id)) = open.take()
                            {
                                children.push(WorksheetChild {
                                    name,
                                    span: begin..end,
                                    relationship_id,
                                });
                            }
                            if depth == 0 {
                                return Ok(Self {
                                    prefix,
                                    strict,
                                    children,
                                    root_close: start,
                                });
                            }
                        },
                        Event::Eof => return Err(invalid("worksheet part has no closed root")),
                        _ => {},
                    }
                    (None, false)
                },
            };
            if let Some(element) = element {
                {
                    if depth == 0 {
                        let name = element.name();
                        if let Some(at) = name.as_ref().iter().position(|byte| *byte == b':') {
                            prefix = String::from_utf8_lossy(&name.as_ref()[..=at]).into_owned();
                        }
                        strict = litchi_ooxml_common::xml::attributes::first_wins(&element)
                            .flatten()
                            .any(|attribute| {
                                attribute.key.as_ref().starts_with(b"xmlns")
                                    && attribute.value.as_ref() == STRICT_MAIN_NS
                            });
                    } else if depth == 1 {
                        let local = element.local_name().as_ref().to_vec();
                        if local == b"AlternateContent" {
                            return Err(invalid(
                                "worksheet markup-compatibility branches cannot be edited",
                            ));
                        }
                        let relationship_id =
                            litchi_ooxml_common::xml::attributes::first_wins(&element)
                                .flatten()
                                .find_map(|attribute| {
                                    (attribute.key.local_name().as_ref() == b"id"
                                        && attribute.key.prefix().is_some())
                                    .then(|| String::from_utf8_lossy(&attribute.value).into_owned())
                                });
                        if empty {
                            children.push(WorksheetChild {
                                name: local,
                                span: start..end,
                                relationship_id,
                            });
                        } else {
                            open = Some((local, start, relationship_id));
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                }
            }
        }
    }

    /// The first direct child named `name`.
    pub(crate) fn child(&self, name: &str) -> Option<&WorksheetChild> {
        self.children
            .iter()
            .find(|child| child.name == name.as_bytes())
    }

    /// The relationships namespace this worksheet's flavor uses.
    pub(crate) const fn relationship_namespace(&self) -> &'static str {
        if self.strict {
            STRICT_RELATIONSHIP_NS
        } else {
            RELATIONSHIP_NS
        }
    }

    /// Insert a direct relationship child (`<drawing r:id/>`,
    /// `<legacyDrawing r:id/>`) at its schema position.
    pub(crate) fn insert_relationship_child(
        &self,
        xml: &[u8],
        name: &str,
        relationship_id: &str,
    ) -> Result<Vec<u8>> {
        if self.child(name).is_some() {
            return Err(invalid(format!(
                "the worksheet already has a {name} element"
            )));
        }
        let rank = WORKSHEET_CHILDREN
            .iter()
            .position(|child| *child == name.as_bytes())
            .ok_or_else(|| invalid("unknown worksheet child"))?;
        let at = self
            .children
            .iter()
            .find(|child| {
                WORKSHEET_CHILDREN
                    .iter()
                    .position(|known| *known == child.name.as_slice())
                    .is_some_and(|position| position > rank)
            })
            .map_or(self.root_close, |child| child.span.start);
        let fragment = format!(
            "<{}{name} xmlns:r=\"{}\" r:id=\"{}\"/>",
            self.prefix,
            self.relationship_namespace(),
            litchi_core::xml::escape_xml(relationship_id)
        );
        let mut out = Vec::with_capacity(xml.len() + fragment.len());
        out.extend_from_slice(&xml[..at]);
        out.extend_from_slice(fragment.as_bytes());
        out.extend_from_slice(&xml[at..]);
        Ok(out)
    }

    /// Rewrite a direct child `name` written in its own `x:` prefix with a
    /// local declaration into the root's prefix, so every worksheet writer
    /// reads it as a direct child. Other children keep their bytes.
    pub(crate) fn in_root_prefix(&self, xml: &[u8], name: &str) -> Vec<u8> {
        let Some(child) = self.child(name) else {
            return xml.to_vec();
        };
        let fragment = &xml[child.span.clone()];
        let namespace = if self.strict {
            STRICT_MAIN_NS
        } else {
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main".as_slice()
        };
        let mut declaration = b" xmlns:x=\"".to_vec();
        declaration.extend_from_slice(namespace);
        declaration.push(b'"');
        if !fragment.starts_with(b"<x:") || self.prefix == "x:" {
            return xml.to_vec();
        }
        let prefix = self.prefix.as_bytes();
        let mut out = Vec::with_capacity(xml.len());
        out.extend_from_slice(&xml[..child.span.start]);
        let mut index = 0;
        while index < fragment.len() {
            let rest = &fragment[index..];
            if rest.starts_with(&declaration) {
                index += declaration.len();
            } else if rest.starts_with(b"</x:") {
                out.extend_from_slice(b"</");
                out.extend_from_slice(prefix);
                index += 4;
            } else if rest.starts_with(b"<x:") {
                out.push(b'<');
                out.extend_from_slice(prefix);
                index += 3;
            } else {
                out.push(fragment[index]);
                index += 1;
            }
        }
        out.extend_from_slice(&xml[child.span.end..]);
        out
    }

    /// Rewrite a direct child `name` written unprefixed, in a default
    /// namespace declared on itself, into the root's prefix (declared on the
    /// child itself instead), so a prefixed worksheet keeps every direct
    /// child in its own prefix (record 0320 gap 7, fixed in record 0322).
    /// Prefixed descendants and every other child keep their bytes.
    pub(crate) fn unprefixed_in_root_prefix(&self, xml: &[u8], name: &str) -> Vec<u8> {
        let Some(child) = self.child(name) else {
            return xml.to_vec();
        };
        let fragment = &xml[child.span.clone()];
        let open = format!("<{name}");
        if self.prefix.is_empty()
            || !fragment.starts_with(open.as_bytes())
            || !matches!(
                fragment.get(open.len()),
                Some(b' ' | b'>' | b'/' | b'\t' | b'\r' | b'\n')
            )
        {
            return xml.to_vec();
        }
        let namespace = if self.strict {
            STRICT_MAIN_NS
        } else {
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main".as_slice()
        };
        let mut declaration = b" xmlns=\"".to_vec();
        declaration.extend_from_slice(namespace);
        declaration.push(b'"');
        let prefix = self.prefix.as_bytes();
        // The owner declares the root's prefix itself, so a reader that
        // parses the owner on its own still resolves it.
        let mut local = b" xmlns:".to_vec();
        local.extend_from_slice(&prefix[..prefix.len() - 1]);
        local.extend_from_slice(b"=\"");
        local.extend_from_slice(namespace);
        local.push(b'"');
        let unprefixed = |rest: &[u8]| {
            rest.first()
                .is_some_and(|byte| byte.is_ascii_alphabetic() || *byte == b'_')
                && !rest
                    .iter()
                    .take_while(|byte| !matches!(byte, b' ' | b'>' | b'/' | b'\t' | b'\r' | b'\n'))
                    .any(|byte| *byte == b':')
        };
        let mut out = Vec::with_capacity(xml.len() + 64);
        out.extend_from_slice(&xml[..child.span.start]);
        let mut index = 0;
        while index < fragment.len() {
            let rest = &fragment[index..];
            if rest.starts_with(&declaration) {
                out.extend_from_slice(&local);
                index += declaration.len();
            } else if rest.starts_with(b"</") && unprefixed(&rest[2..]) {
                out.extend_from_slice(b"</");
                out.extend_from_slice(prefix);
                index += 2;
            } else if rest.starts_with(b"<") && unprefixed(&rest[1..]) {
                out.push(b'<');
                out.extend_from_slice(prefix);
                index += 1;
            } else {
                out.push(fragment[index]);
                index += 1;
            }
        }
        out.extend_from_slice(&xml[child.span.end..]);
        out
    }

    /// Remove the first direct child named `name`; `None` when it is absent.
    pub(crate) fn remove_child(&self, xml: &[u8], name: &str) -> Option<Vec<u8>> {
        let child = self.child(name)?;
        let mut out = xml.to_vec();
        out.drain(child.span.clone());
        Some(out)
    }
}

/// The text of a predefined or character entity reference (`amp`, `#x41`).
pub(crate) fn entity(name: &[u8]) -> Option<String> {
    let name = std::str::from_utf8(name).ok()?;
    Some(match name {
        "amp" => "&".to_owned(),
        "lt" => "<".to_owned(),
        "gt" => ">".to_owned(),
        "apos" => "'".to_owned(),
        "quot" => "\"".to_owned(),
        _ => {
            let code = name.strip_prefix('#')?;
            let value = match code.strip_prefix('x') {
                Some(hex) => u32::from_str_radix(hex, 16).ok()?,
                None => code.parse().ok()?,
            };
            char::from_u32(value)?.to_string()
        },
    })
}

fn position(reader: &quick_xml::Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))
}

/// The `Patch`, `Commit` and `Diagnostics` types of one family, and its
/// `Error` re-export (ADR 0034 §3).
macro_rules! family {
    ($what:literal) => {
        pub use crate::Error;

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

            /// The semantic payload retained by this reversible patch.
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
