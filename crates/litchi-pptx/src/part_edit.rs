//! The staged part-level engine behind the package-scoped edit families
//! (record 0351 §2.3).
//!
//! [`Staged`] holds one immutable [`Package`] and a staged copy of its OPC
//! graph. A family runs its typed graph operations on the staged copy and
//! commits: [`PartPatch`] captures the exact before and after state (content
//! type, payload and relationships) of every part the operations added,
//! changed or removed, plus the package-level relationships, and an
//! unchanged staging publishes the source package itself, so an exact no-op
//! keeps every byte and every signature. The source package never changes.
//!
//! The families built on it expose the ADR 0034 shape through [`family!`]:
//! `x::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`,
//! `Package::edit_x()` and `Package::apply_x(&Patch)`.

use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use crate::{Error, Package, Result};

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

/// Exact publication bytes without retaining the package's unrelated parts.
#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipImage {
    bytes: Arc<Vec<u8>>,
    member_present: bool,
}

impl RelationshipImage {
    fn new(package: &OpcPackage, owner: &PackURI) -> Result<Self> {
        let (bytes, member_present) = match package.retained_relationships_xml(owner)? {
            Some(retained) => retained,
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
        })
    }
}

/// Compact manifest state. Canonical structural output is reconstructible
/// from the typed graph. Only an explicit lexical patch needs an exact digest
/// of a canonical base; producer source bytes always remain exact images.
#[derive(Debug, Clone)]
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
                return Err(Error::PatchConflict);
            }
        }
        Ok(bytes)
    }
}

impl PartialEq for ContentTypesImage {
    fn eq(&self, other: &Self) -> bool {
        self.retained == other.retained && self.exact_canonical == other.exact_canonical
    }
}
impl Eq for ContentTypesImage {}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
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
    fn read(package: &Arc<OpcPackage>, part: &dyn Part) -> Result<Self> {
        Ok(Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: relationships(part.rels()),
            relationship_xml: RelationshipImage::new(package, part.partname())?,
        })
    }

    /// Whether `part` holds exactly this state, without copying its payload.
    fn matches(&self, package: &OpcPackage, part: &dyn Part) -> Result<bool> {
        Ok(self.content_type == part.content_type()
            && {
                let blob = part.blob();
                std::ptr::eq(self.blob.as_slice(), blob) || self.blob.as_slice() == blob
            }
            && self.relationships == relationships(part.rels())
            && self.relationship_xml == RelationshipImage::new(package, part.partname())?)
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        add_relationships(part.rels_mut(), &self.relationships)?;
        Ok(part)
    }
}

fn same_part(before: &dyn Part, after: &dyn Part) -> bool {
    before.content_type() == after.content_type()
        && {
            let (left, right) = (before.blob(), after.blob());
            std::ptr::eq(left, right) || left == right
        }
        && crate::opened::relationships_equal(before.rels(), after.rels())
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

/// The exact, reversible part-level change of one staged edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
    root: Option<(RelationshipImage, RelationshipImage)>,
    content_types: Option<(ContentTypesImage, ContentTypesImage)>,
    exact_content_types: bool,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none() && self.content_types.is_none()
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

    /// Capture every part that differs between `before` and `after`.
    ///
    /// The staged graph is a clone of the source, so an untouched part shares
    /// its payload allocation and compares by pointer; only touched parts are
    /// compared by value.
    pub(crate) fn capture(before: &Arc<OpcPackage>, after: &Arc<OpcPackage>) -> Result<Self> {
        // One commit changes at most as many parts as the package's source
        // limits admit.
        let maximum = before.read_limits().max_parts();
        let mut deltas = Vec::new();
        let mut push = |delta: Delta| -> Result<()> {
            if delta.before == delta.after {
                return Ok(());
            }
            if deltas.len() >= maximum {
                return Err(Error::EditLimit {
                    kind: crate::opened::LimitKind::Parts,
                    maximum,
                });
            }
            deltas.push(delta);
            Ok(())
        };
        for part in before.try_iter_parts() {
            let part = part?;
            let name = part.partname();
            match optional_part(after, name)? {
                Some(other)
                    if same_part(part, other)
                        && before.same_retained_relationships(after, name)? => {},
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
        let owner = PackURI::new("/").map_err(Error::Uri)?;
        let root = if crate::opened::relationships_equal(before.rels(), after.rels())
            && before.same_retained_relationships(after, &owner)?
        {
            None
        } else {
            let pair = (
                RelationshipImage::new(before, &owner)?,
                RelationshipImage::new(after, &owner)?,
            );
            (!crate::opened::relationships_equal(before.rels(), after.rels()) || pair.0 != pair.1)
                .then_some(pair)
        };
        let changes_types = deltas
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
        Ok(Self {
            deltas,
            root,
            content_types,
            exact_content_types,
        })
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package, operation: &'static str) -> Result<Package> {
        self.apply_inner(package, operation, None)
    }

    /// Replay with a family-owned semantic validator for changed payloads.
    ///
    /// Generic replay still owns before-state matching, exact patch scope and
    /// atomic target construction. A typed family may validate only the
    /// changed payloads against the recipient's retained semantic policy.
    pub(crate) fn apply_with_replay_validation(
        &self,
        package: &Package,
        operation: &'static str,
        mut validate: impl FnMut(&Package, &OpcPackage, &Self) -> Result<()>,
    ) -> Result<Package> {
        self.apply_inner(package, operation, Some(&mut validate))
    }

    fn apply_inner(
        &self,
        package: &Package,
        operation: &'static str,
        mut validate: Option<&mut dyn FnMut(&Package, &OpcPackage, &Self) -> Result<()>>,
    ) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.duplicate());
        }
        package.ensure_plain_mutation(operation)?;
        let source = &package.opc;
        if let Some((before, _)) = &self.root {
            let owner = PackURI::new("/").map_err(Error::Uri)?;
            if RelationshipImage::new(source, &owner)? != *before {
                return Err(Error::PatchConflict);
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
                return Err(Error::PatchConflict);
            }
        }
        crate::package::signature::refuse_signed_change(source)?;
        let mut target = (**source).clone();
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
                if current != after.relationship_xml {
                    let token = target.source_relationships(&delta.name, limits)?;
                    target.try_replace_relationships_bytes(
                        &token,
                        Arc::clone(&after.relationship_xml.bytes),
                        after.relationship_xml.member_present,
                        limits,
                    )?;
                }
            }
        }
        if let Some((_, after)) = &self.root {
            let limits = target.read_limits();
            let owner = PackURI::new("/").map_err(Error::Uri)?;
            let current = target.source_relationships(&owner, limits)?;
            target.try_replace_relationships_bytes(
                &current,
                Arc::clone(&after.bytes),
                after.member_present,
                limits,
            )?;
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
                    return Err(Error::PatchConflict);
                }
            }
        }
        if let Some(validate) = validate {
            validate(package, &target, self)?;
        }
        crate::package::signature::settle_unsigned_change(&mut target)?;
        Ok(package.derive(target))
    }

    /// Visit only added or byte-replaced parts in an assembled replay target.
    /// Relationship-only deltas are skipped so a small patch does not turn
    /// replay into a package-wide semantic scan.
    pub(crate) fn validate_changed_parts(
        &self,
        target: &OpcPackage,
        mut validate: impl FnMut(&PackURI, &dyn Part) -> Result<()>,
    ) -> Result<()> {
        for delta in &self.deltas {
            let Some(after) = &delta.after else {
                continue;
            };
            let payload_changed = match &delta.before {
                Some(before) => {
                    before.content_type != after.content_type || before.blob != after.blob
                },
                None => true,
            };
            if !payload_changed {
                continue;
            }
            validate(&delta.name, target.get_part(&delta.name)?)?;
        }
        Ok(())
    }
}

/// A staged part-level edit of one [`Package`].
#[derive(Debug)]
pub(crate) struct Staged<'a> {
    package: &'a Package,
    source: &'a OpcPackage,
    staged: Option<OpcPackage>,
    changed: bool,
}

impl<'a> Staged<'a> {
    /// Start staging; refuses a package whose typed mutations are not allowed
    /// (a package with retained encryption provenance).
    pub(crate) fn new(package: &'a Package, operation: &'static str) -> Result<Self> {
        package.ensure_plain_mutation(operation)?;
        Ok(Self {
            package,
            source: &package.opc,
            staged: None,
            changed: false,
        })
    }

    /// Start staging a raw graph edit, which the encryption policy does not
    /// refuse: the raw OPC family declassifies its result instead.
    pub(crate) fn raw(package: &'a Package) -> Self {
        Self {
            package,
            source: &package.opc,
            staged: None,
            changed: false,
        }
    }

    /// The staged OPC graph.
    pub(crate) const fn opc(&self) -> &OpcPackage {
        match &self.staged {
            Some(staged) => staged,
            None => self.source,
        }
    }

    /// The owner-retained semantic codec policies for this staged family.
    pub(crate) fn semantic_limits(&self) -> crate::open_options::SemanticLimits {
        self.package.semantic_limits()
    }

    /// Run one typed graph operation on the staged graph. An error leaves
    /// the staged graph as it was before the operation.
    pub(crate) fn update<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        self.changed = true;
        Ok(value)
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.staged = None;
        self.changed = false;
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging publishes the source package.
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        if !self.changed {
            return Ok((self.package.duplicate(), PartPatch::default()));
        }
        let Some(staged) = self.staged else {
            return Ok((self.package.duplicate(), PartPatch::default()));
        };
        let staged = Arc::new(staged);
        let patch = PartPatch::capture(&self.package.opc, &staged)?;
        if patch.is_noop() {
            return Ok((self.package.duplicate(), patch));
        }
        Ok((self.package.derive(staged), patch))
    }
}

/// The `Patch`, `Commit`, `Diagnostics`, `Error` and `LimitKind` members of
/// a family built on [`Staged`].
macro_rules! family {
    ($what:literal) => {
        pub use crate::Error;
        /// The bounded dimension of this family: the number of parts one
        /// commit may change, at most the package's source-limit part count.
        pub use crate::opened::LimitKind;
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!("The result of a ", $what, " edit's `commit`: the new package, the patch and the diagnostics.")]
        pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

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
        }

        #[allow(dead_code, reason = "not every family replays through the shared engine")]
        pub(crate) fn commit_of(
            (package, patch): (crate::Package, crate::part_edit::PartPatch),
        ) -> Commit {
            let diagnostics = Diagnostics::new(!patch.is_noop());
            Commit::new(package, Patch(patch), diagnostics)
        }
    };
}

pub(crate) use family;
