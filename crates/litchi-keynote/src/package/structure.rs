//! The structural-edit kernel shared by structural transactions (record 0298).
//!
//! Structural edits clone, create and remove native objects and whole
//! components, so they need what the focused value transactions never do:
//! schema-directed reference rewriting ([`refs`], generated [`schema`]),
//! registry edits ([`metadata`]) and a copy-on-write component view that
//! publishes one candidate artifact ([`workspace`]). Every published
//! candidate is reopened and validated by the transaction that staged it.

pub(in crate::package) mod metadata;
pub(in crate::package) mod refs;
mod schema;
pub(in crate::package) mod wire;
pub(in crate::package) mod workspace;

use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};
use litchi_iwa_core::{Archive, Limits as ArchiveLimits};

use super::{Package, PhysicalSource};

/// A bounded resource reported by the structural kernel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) enum StructureLimit {
    /// Bytes of one protobuf payload.
    WireBytes,
    /// Fields of one protobuf payload.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Rewrite work.
    Work,
    /// Bytes of one rewritten payload or member.
    OutputBytes,
    /// Bytes of one archive entry, object or package.
    EntryBytes,
}

/// A failure inside the structural kernel. Transactions map it to their
/// family error; the payloads are structural metadata only.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::package) enum StructureError {
    /// A referenced object is missing.
    MissingObject,
    /// A native message holds references the kernel cannot rewrite
    /// (its type is not modelled); the payload is the message type.
    UnsupportedContent(u32),
    /// A registry record uses a shape the kernel does not edit
    /// (versioned references, ambiguous object identifiers).
    UnsupportedRegistry,
    /// The package registry is malformed or inconsistent.
    Registry(&'static str),
    /// A native payload does not have the expected shape.
    Unexpected(&'static str),
    /// A payload is not well-formed protobuf.
    Malformed,
    /// The package object identifier space is exhausted.
    IdentifierOverflow,
    /// A physical or wire resource limit was exceeded.
    Limit {
        kind: StructureLimit,
        observed: u64,
        maximum: u64,
    },
    /// An allocation of `amount` units failed.
    Allocation(usize),
    /// The physical layer refused the candidate.
    Physical,
}

impl From<litchi_iwa_common::Error> for StructureError {
    fn from(error: litchi_iwa_common::Error) -> Self {
        match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Self::Limit {
                kind: match kind {
                    litchi_iwa_common::LimitKind::InputBytes => StructureLimit::WireBytes,
                    litchi_iwa_common::LimitKind::OutputBytes => StructureLimit::OutputBytes,
                    litchi_iwa_common::LimitKind::Nesting => StructureLimit::WireNesting,
                    litchi_iwa_common::LimitKind::RewriteWork => StructureLimit::Work,
                    _ => StructureLimit::WireFields,
                },
                observed: observed as u64,
                maximum: limit as u64,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => Self::Allocation(amount),
            _ => Self::Malformed,
        }
    }
}

impl From<litchi_iwa_text_wire::CharStyleError> for StructureError {
    fn from(error: litchi_iwa_text_wire::CharStyleError) -> Self {
        match error {
            litchi_iwa_text_wire::CharStyleError::Limit {
                kind,
                observed,
                maximum,
            } => Self::from(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit: maximum,
            }),
            litchi_iwa_text_wire::CharStyleError::Allocation { amount } => Self::Allocation(amount),
            _ => Self::Malformed,
        }
    }
}

impl From<litchi_iwa_core::Error> for StructureError {
    fn from(error: litchi_iwa_core::Error) -> Self {
        match error {
            litchi_iwa_core::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: StructureLimit::EntryBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_core::Error::Allocation { requested, .. } => Self::Allocation(requested),
            _ => Self::Physical,
        }
    }
}

impl From<litchi_iwa_archive::Error> for StructureError {
    fn from(error: litchi_iwa_archive::Error) -> Self {
        match error {
            litchi_iwa_archive::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: StructureLimit::EntryBytes,
                observed,
                maximum,
            },
            litchi_iwa_archive::Error::Allocation { amount, .. } => Self::Allocation(amount),
            litchi_iwa_archive::Error::Iwa(error) => error.into(),
            _ => Self::Physical,
        }
    }
}

impl From<refs::RefError> for StructureError {
    fn from(error: refs::RefError) -> Self {
        match error {
            refs::RefError::UnsupportedType(message_type) => Self::UnsupportedContent(message_type),
            refs::RefError::Malformed | refs::RefError::Depth => Self::Malformed,
        }
    }
}

impl Package {
    /// Whether this package retains an exact physical source that structural
    /// edits can reassemble.
    pub(in crate::package) fn structure_is_exact(&self) -> bool {
        matches!(&self.state.source, PhysicalSource::Package(catalog) if catalog.source_is_exact())
    }

    /// The parsed archive of one component member.
    pub(in crate::package) fn structure_component(&self, member: &str) -> Option<&Archive> {
        self.state
            .source
            .components()
            .get(member)
            .map(litchi_iwa_archive::Component::archive)
    }

    /// Whether the package has a member with this exact name.
    pub(in crate::package) fn structure_has_member(&self, member: &str) -> bool {
        self.structure_is_exact()
            && self
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == member)
    }

    /// The stored bytes of one package member, when present.
    pub(in crate::package) fn structure_member_bytes(&self, member: &str) -> Option<&[u8]> {
        if !self.structure_is_exact() {
            return None;
        }
        self.state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == member)
            .map(|entry| entry.data())
    }

    /// The component member holding one object.
    pub(in crate::package) fn structure_object_component(&self, identifier: u64) -> Option<&str> {
        self.object_with_component(identifier)
            .map(|(member, _object)| member)
    }

    /// The largest object identifier present in any component.
    pub(in crate::package) fn structure_max_object_identifier(&self) -> u64 {
        self.state
            .object_index
            .iter()
            .map(|locator| locator.identifier)
            .max()
            .unwrap_or(0)
    }

    /// The number of objects in the package.
    pub(in crate::package) fn structure_object_count(&self) -> usize {
        self.state.total_objects
    }

    /// The physical IWA limits this package was opened with.
    pub(in crate::package) fn structure_archive_limits(
        &self,
    ) -> Result<ArchiveLimits, StructureError> {
        Ok(self.state.options.archive().effective_archive_limits()?)
    }

    /// Reassemble the exact source with structural changes. The root
    /// preview members are always deleted: a structural edit changes what
    /// they render.
    pub(in crate::package) fn structure_reassemble(
        &self,
        insertions: &[EntryInsertion<'_>],
        edits: &[EntryEdit<'_>],
        deleted: &[&str],
    ) -> Result<Vec<u8>, StructureError> {
        if !self.structure_is_exact() {
            return Err(StructureError::Physical);
        }
        let catalog = self.state.source.package();
        let previews = super::rendering_invalidation::root_preview_deletions(catalog)
            .map_err(|_error| StructureError::Physical)?;
        let mut names = Vec::with_capacity(deleted.len() + previews.len());
        names.extend_from_slice(deleted);
        names.extend(
            previews
                .names()
                .iter()
                .copied()
                .filter(|name| !deleted.contains(name)),
        );
        Ok(catalog.reassemble_with_changes_to_bytes(
            insertions,
            edits,
            &names,
            self.state.options.archive(),
        )?)
    }

    /// Reopen a structural candidate under this package's limits and
    /// validate it.
    pub(in crate::package) fn structure_reopen(
        &self,
        bytes: Vec<u8>,
    ) -> Result<Self, super::ReadError> {
        // A structural candidate is an edited reassembly of this package:
        // unchanged IWA members reuse their parsed archives (the ZIP and
        // every budget are still checked in full).
        let candidate = match &self.state.source {
            PhysicalSource::Package(base) => {
                let catalog = base.reopen(Arc::from(bytes))?;
                Self::from_source_catalog(catalog, self.state.options)?
            },
            PhysicalSource::Semantic(_) => {
                Self::from_source_with_options(Arc::from(bytes), self.state.options)?
            },
        };
        candidate.validate()?;
        Ok(candidate)
    }
}

#[cfg(test)]
pub(super) mod test_support {
    use std::path::{Path, PathBuf};

    use crate::Package;

    #[test]
    fn character_style_resource_errors_survive_the_owner_boundary() {
        use super::{StructureError, StructureLimit};
        assert_eq!(
            StructureError::from(litchi_iwa_text_wire::CharStyleError::Limit {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: 11,
                maximum: 10,
            }),
            StructureError::Limit {
                kind: StructureLimit::Work,
                observed: 11,
                maximum: 10,
            }
        );
        assert_eq!(
            StructureError::from(litchi_iwa_text_wire::CharStyleError::Allocation { amount: 42 }),
            StructureError::Allocation(42)
        );
        assert_eq!(
            StructureError::from(litchi_iwa_text_wire::CharStyleError::Invalid),
            StructureError::Malformed
        );
    }

    /// Every Keynote package of the tracked corpus that opens.
    pub(in crate::package) fn corpus() -> Vec<(PathBuf, Package)> {
        let directory = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote");
        let mut paths = std::fs::read_dir(&directory)
            .expect("Keynote corpus directory")
            .map(|entry| entry.expect("corpus entry").path())
            .filter(|path| {
                path.is_file()
                    && path.extension().and_then(|extension| extension.to_str()) == Some("key")
            })
            .collect::<Vec<_>>();
        paths.sort();
        let packages = paths
            .into_iter()
            .filter_map(|path| Package::open(&path).ok().map(|package| (path, package)))
            .collect::<Vec<_>>();
        assert!(packages.len() >= 10, "corpus too small");
        packages
    }

    /// One corpus package by file name.
    pub(in crate::package) fn fixture(name: &str) -> Package {
        Package::open(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote")
                .join(name),
        )
        .expect("Keynote fixture")
    }

    /// The raw `TSP.PackageMetadata` payload of a package.
    pub(in crate::package) fn registry_payload(package: &Package) -> Vec<u8> {
        let component = package
            .structure_component("Index/Metadata.iwa")
            .expect("registry component");
        component
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .find(|message| message.type_ == 11_006)
            .expect("registry payload")
            .data
            .clone()
    }
}
