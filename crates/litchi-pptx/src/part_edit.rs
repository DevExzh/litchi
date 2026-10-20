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

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
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
    fn read(part: &dyn Part) -> Self {
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: relationships(part.rels()),
        }
    }

    /// Whether `part` holds exactly this state, without copying its payload.
    fn matches(&self, part: &dyn Part) -> bool {
        self.content_type == part.content_type()
            && {
                let blob = part.blob();
                std::ptr::eq(self.blob.as_slice(), blob) || self.blob.as_slice() == blob
            }
            && self.relationships == relationships(part.rels())
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
    root: Option<(Vec<Relationship>, Vec<Relationship>)>,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none()
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
        }
    }

    /// Capture every part that differs between `before` and `after`.
    ///
    /// The staged graph is a clone of the source, so an untouched part shares
    /// its payload allocation and compares by pointer; only touched parts are
    /// compared by value.
    pub(crate) fn capture(before: &OpcPackage, after: &OpcPackage) -> Result<Self> {
        // One commit changes at most as many parts as the package's source
        // limits admit.
        let maximum = before.read_limits().max_parts();
        let mut deltas = Vec::new();
        let mut push = |delta: Delta| -> Result<()> {
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
                Some(other) if same_part(part, other) => {},
                other => push(Delta {
                    name: name.clone(),
                    before: Some(State::read(part)),
                    after: other.map(State::read),
                })?,
            }
        }
        for part in after.try_iter_parts() {
            let part = part?;
            if optional_part(before, part.partname())?.is_none() {
                push(Delta {
                    name: part.partname().clone(),
                    before: None,
                    after: Some(State::read(part)),
                })?;
            }
        }
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let root = if crate::opened::relationships_equal(before.rels(), after.rels()) {
            None
        } else {
            Some((relationships(before.rels()), relationships(after.rels())))
        };
        Ok(Self { deltas, root })
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package, operation: &'static str) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.duplicate());
        }
        package.ensure_plain_mutation(operation)?;
        let source = &package.opc;
        if let Some((before, _)) = &self.root
            && relationships(source.rels()) != *before
        {
            return Err(Error::PatchConflict);
        }
        for delta in &self.deltas {
            let current = optional_part(source, &delta.name)?;
            let matches = match (&delta.before, current) {
                (None, None) => true,
                (Some(state), Some(part)) => state.matches(part),
                _ => false,
            };
            if !matches {
                return Err(Error::PatchConflict);
            }
        }
        crate::package::signature::refuse_signed_change(source)?;
        let mut target = source.clone();
        for delta in &self.deltas {
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        if let Some((before, after)) = &self.root {
            let graph = target.rels_mut();
            for relationship in before {
                graph.remove(&relationship.id);
            }
            add_relationships(graph, after)?;
        }
        crate::package::signature::settle_unsigned_change(&mut target)?;
        Ok(package.derive(target))
    }
}

/// A staged part-level edit of one [`Package`].
#[derive(Debug)]
pub(crate) struct Staged<'a> {
    package: &'a Package,
    staged: OpcPackage,
    changed: bool,
}

impl<'a> Staged<'a> {
    /// Start staging; refuses a package whose typed mutations are not allowed
    /// (a package with retained encryption provenance).
    pub(crate) fn new(package: &'a Package, operation: &'static str) -> Result<Self> {
        package.ensure_plain_mutation(operation)?;
        Ok(Self {
            package,
            staged: package.opc.clone(),
            changed: false,
        })
    }

    /// Start staging a raw graph edit, which the encryption policy does not
    /// refuse: the raw OPC family declassifies its result instead.
    pub(crate) fn raw(package: &'a Package) -> Self {
        Self {
            package,
            staged: package.opc.clone(),
            changed: false,
        }
    }

    /// The staged OPC graph.
    pub(crate) const fn opc(&self) -> &OpcPackage {
        &self.staged
    }

    /// Run one typed graph operation on the staged graph. An error leaves
    /// the staged graph as it was before the operation.
    pub(crate) fn update<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.staged.clone();
        let value = operation(&mut candidate)?;
        self.staged = candidate;
        self.changed = true;
        Ok(value)
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.staged = self.package.opc.clone();
        self.changed = false;
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging publishes the source package.
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        if !self.changed {
            return Ok((self.package.duplicate(), PartPatch::default()));
        }
        let patch = PartPatch::capture(&self.package.opc, &self.staged)?;
        if patch.is_noop() {
            return Ok((self.package.duplicate(), patch));
        }
        Ok((self.package.derive(self.staged), patch))
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
