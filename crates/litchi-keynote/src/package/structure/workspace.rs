//! A staged, copy-on-write view of one package's components and registry.
//!
//! Structural edits read objects through the workspace, clone a component
//! the first time they change it, add new components, and finally publish
//! one candidate artifact: every changed component is re-encoded, the
//! registry is re-encoded once, and the exact source ZIP is reassembled with
//! the replacements and insertions in a single bounded pass (the root
//! previews, which a structural edit invalidates, are deleted).

use std::collections::{BTreeMap, BTreeSet, HashMap};

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{Archive, ArchiveObject, Limits as ArchiveLimits, RawMessage, SnappyStream};
use sha1::{Digest, Sha1};

use super::metadata::{Registry, Uuid};
use super::{StructureError, refs};
use crate::package::Package;

const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

/// The staged state of one structural transaction.
pub(in crate::package) struct Workspace<'p> {
    package: &'p Package,
    pub(in crate::package) wire: WireLimits,
    archive_limits: ArchiveLimits,
    /// Existing components changed so far, by member name.
    changed: BTreeMap<String, Archive>,
    /// Components created by this transaction, by member name.
    created: BTreeMap<String, Archive>,
    /// Where staged objects live now (overrides the package index).
    located: HashMap<u64, String>,
    pub(in crate::package) registry: Registry,
    registry_location: (u64, usize),
    /// Every object UUID in the registry, for collision checks.
    uuids: BTreeSet<(u64, u64)>,
    next_identifier: u64,
    /// Objects inserted by this transaction.
    inserted: usize,
    /// Non-archive package members created by this transaction (data).
    members: BTreeMap<String, Vec<u8>>,
    /// Source members deleted by this transaction (data culled with its
    /// final owner).
    deleted: BTreeSet<String>,
    /// Source objects removed by this transaction and not re-inserted.
    removed: BTreeSet<u64>,
}

impl<'p> Workspace<'p> {
    /// Stage a transaction over `package`.
    pub(in crate::package) fn new(package: &'p Package) -> Result<Self, StructureError> {
        let wire = package
            .wire_limits()
            .map_err(|_| StructureError::Registry("wire limits"))?;
        let archive_limits = package.structure_archive_limits()?;
        let component = package
            .structure_component(METADATA_MEMBER)
            .ok_or(StructureError::Registry("package metadata component"))?;
        let mut location = None;
        for object in &component.objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    let identifier = object
                        .archive_info
                        .identifier
                        .ok_or(StructureError::Registry("package metadata identity"))?;
                    if location.replace((identifier, index)).is_some() {
                        return Err(StructureError::Registry("duplicate package metadata"));
                    }
                }
            }
        }
        let registry_location = location.ok_or(StructureError::Registry("package metadata"))?;
        let payload = &component
            .object(registry_location.0)
            .ok_or(StructureError::Registry("package metadata"))?
            .messages[registry_location.1]
            .data;
        let registry = Registry::decode(payload, wire)?;
        let uuids = registry
            .components()
            .flat_map(|component| component.object_uuids().iter())
            .map(|(_, uuid)| (uuid.lower, uuid.upper))
            .collect();
        let next_identifier = registry
            .last_object_identifier()
            .max(package.structure_max_object_identifier())
            .checked_add(1)
            .ok_or(StructureError::IdentifierOverflow)?;
        Ok(Self {
            package,
            wire,
            archive_limits,
            changed: BTreeMap::new(),
            created: BTreeMap::new(),
            located: HashMap::new(),
            registry,
            registry_location,
            uuids,
            next_identifier,
            inserted: 0,
            members: BTreeMap::new(),
            deleted: BTreeSet::new(),
            removed: BTreeSet::new(),
        })
    }

    pub(in crate::package) const fn package(&self) -> &'p Package {
        self.package
    }

    /// Allocate a fresh package-wide object identifier.
    pub(in crate::package) fn allocate(&mut self) -> Result<u64, StructureError> {
        let identifier = self.next_identifier;
        self.next_identifier = identifier
            .checked_add(1)
            .ok_or(StructureError::IdentifierOverflow)?;
        self.registry.set_last_object_identifier(identifier);
        Ok(identifier)
    }

    /// The member name holding `identifier`, if it exists.
    pub(in crate::package) fn component_of(&self, identifier: u64) -> Option<String> {
        if let Some(name) = self.located.get(&identifier) {
            return Some(name.clone());
        }
        self.package
            .structure_object_component(identifier)
            .map(str::to_owned)
    }

    fn archive(&self, member: &str) -> Option<&Archive> {
        self.created
            .get(member)
            .or_else(|| self.changed.get(member))
            .or_else(|| self.package.structure_component(member))
    }

    /// Borrow one staged object.
    pub(in crate::package) fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        let member = self.component_of(identifier)?;
        self.archive(&member)?.object(identifier)
    }

    /// Borrow one staged object, or fail with a typed missing-object error.
    pub(in crate::package) fn require(
        &self,
        identifier: u64,
    ) -> Result<&ArchiveObject, StructureError> {
        self.object(identifier).ok_or(StructureError::MissingObject)
    }

    /// The unique payload of `message_type` held by `identifier`.
    pub(in crate::package) fn payload(
        &self,
        identifier: u64,
        message_type: u32,
    ) -> Result<&[u8], StructureError> {
        let object = self.require(identifier)?;
        let mut found = object
            .messages
            .iter()
            .filter(|message| message.type_ == message_type);
        let first = found
            .next()
            .ok_or(StructureError::Unexpected("message type"))?;
        if found.next().is_some() {
            return Err(StructureError::Unexpected("duplicate message"));
        }
        Ok(&first.data)
    }

    fn archive_mut(&mut self, member: &str) -> Result<&mut Archive, StructureError> {
        if self.created.contains_key(member) {
            return self
                .created
                .get_mut(member)
                .ok_or(StructureError::MissingObject);
        }
        if !self.changed.contains_key(member) {
            let source = self
                .package
                .structure_component(member)
                .ok_or(StructureError::MissingObject)?
                .clone();
            self.changed.insert(member.to_owned(), source);
        }
        self.changed
            .get_mut(member)
            .ok_or(StructureError::MissingObject)
    }

    /// Mutably borrow one staged object, cloning its component on first use.
    pub(in crate::package) fn object_mut(
        &mut self,
        identifier: u64,
    ) -> Result<&mut ArchiveObject, StructureError> {
        let member = self
            .component_of(identifier)
            .ok_or(StructureError::MissingObject)?;
        self.archive_mut(&member)?
            .object_mut(identifier)
            .ok_or(StructureError::MissingObject)
    }

    /// The physical archive limits of the staged package.
    pub(in crate::package) const fn archive_limits(&self) -> ArchiveLimits {
        self.archive_limits
    }

    /// How many objects this transaction inserted.
    pub(in crate::package) const fn created_objects(&self) -> usize {
        self.inserted
    }

    /// How many package members this transaction rewrites or creates
    /// (the registry member included).
    pub(in crate::package) fn touched_members(&self) -> usize {
        let registry = usize::from(!self.changed.contains_key(METADATA_MEMBER));
        self.changed.len() + self.created.len() + self.members.len() + registry
    }

    /// Insert a new object into an existing or created component.
    pub(in crate::package) fn insert(
        &mut self,
        member: &str,
        object: ArchiveObject,
    ) -> Result<(), StructureError> {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(StructureError::Unexpected("object identity"))?;
        if self.object(identifier).is_some() {
            return Err(StructureError::Unexpected("object identity collision"));
        }
        let limits = self.archive_limits;
        self.archive_mut(member)?.insert_object(object, limits)?;
        self.located.insert(identifier, member.to_owned());
        self.removed.remove(&identifier);
        self.inserted += 1;
        Ok(())
    }

    /// Remove one staged object from its component.
    pub(in crate::package) fn remove_object(
        &mut self,
        identifier: u64,
    ) -> Result<ArchiveObject, StructureError> {
        let member = self
            .component_of(identifier)
            .ok_or(StructureError::MissingObject)?;
        let removed = self
            .archive_mut(&member)?
            .remove_object(identifier)
            .ok_or(StructureError::MissingObject)?;
        self.located.remove(&identifier);
        self.removed.insert(identifier);
        Ok(removed)
    }

    /// Lower the package watermark as the legacy host does when a removal
    /// takes out the object that holds it: the new watermark is the largest
    /// surviving identifier. A watermark above every object (an unused
    /// metadata-only suffix) is kept, and so is one an allocation in this
    /// transaction raised.
    pub(in crate::package) fn release_removed_identifier_suffix(&mut self) {
        let last = self.registry.last_object_identifier();
        if !self.removed.contains(&last) {
            return;
        }
        let surviving = self
            .package
            .state
            .object_index
            .iter()
            .map(|locator| locator.identifier)
            .filter(|identifier| !self.removed.contains(identifier))
            .chain(self.located.keys().copied())
            .max()
            .unwrap_or(0);
        self.registry.release_last_object_identifier(surviving);
        self.next_identifier = surviving.saturating_add(1);
    }

    /// Stage the deletion of one source package member.
    pub(in crate::package) fn delete_member(&mut self, member: &str) -> Result<(), StructureError> {
        if !self.package.structure_has_member(member) || self.members.contains_key(member) {
            return Err(StructureError::Registry("package member to delete"));
        }
        self.deleted.insert(member.to_owned());
        Ok(())
    }

    /// Whether a package member of this exact name exists or is staged.
    pub(in crate::package) fn has_member(&self, member: &str) -> bool {
        self.members.contains_key(member)
            || self.created.contains_key(member)
            || self.package.structure_has_member(member)
    }

    /// Stage a new raw package member (for example a `Data/` file).
    pub(in crate::package) fn insert_member(
        &mut self,
        member: &str,
        bytes: Vec<u8>,
    ) -> Result<(), StructureError> {
        if member.is_empty() || self.has_member(member) {
            return Err(StructureError::Registry("package member exists"));
        }
        self.members.insert(member.to_owned(), bytes);
        Ok(())
    }

    /// Create a new, empty component member.
    pub(in crate::package) fn create_component(
        &mut self,
        member: &str,
    ) -> Result<(), StructureError> {
        if self.archive(member).is_some() || self.package.structure_has_member(member) {
            return Err(StructureError::Registry("component member exists"));
        }
        self.created.insert(member.to_owned(), Archive::new());
        Ok(())
    }

    /// Identifiers of every object in `member`, in archive order.
    pub(in crate::package) fn member_objects(&self, member: &str) -> Vec<u64> {
        self.archive(member)
            .map(|archive| {
                archive
                    .objects
                    .iter()
                    .filter_map(|object| object.archive_info.identifier)
                    .collect()
            })
            .unwrap_or_default()
    }

    /// A clone of `source` under `new_identifier`, every admitted payload
    /// reference rewritten through `remap`. Refuses a message whose type has
    /// no reference schema while it lists references, and a message whose
    /// listed references the schema walk does not find.
    pub(in crate::package) fn clone_object(
        &self,
        source: &ArchiveObject,
        new_identifier: u64,
        remap: &HashMap<u64, u64>,
    ) -> Result<ArchiveObject, StructureError> {
        let mut messages = Vec::with_capacity(source.messages.len());
        for (message, info) in source
            .messages
            .iter()
            .zip(&source.archive_info.message_infos)
        {
            if !refs::admitted(message.type_) {
                if info.object_references.is_empty() && info.data_references.is_empty() {
                    messages.push(message.clone());
                    continue;
                }
                return Err(StructureError::UnsupportedContent(message.type_));
            }
            let walked = refs::object_references(message.type_, &message.data, self.wire)?;
            if info
                .object_references
                .iter()
                .any(|listed| !walked.contains(listed))
            {
                return Err(StructureError::UnsupportedContent(message.type_));
            }
            let data = refs::rewrite(message.type_, &message.data, self.wire, &mut |kind, id| {
                (kind == refs::RefKind::Object)
                    .then(|| remap.get(&id).copied())
                    .flatten()
            })?
            .unwrap_or_else(|| message.data.clone());
            messages.push(RawMessage {
                type_: message.type_,
                data,
            });
        }
        let mut pairs = source
            .archive_info
            .identifier
            .into_iter()
            .chain(source.archive_info.message_infos.iter().flat_map(|info| {
                info.object_references.iter().copied().chain(
                    info.field_infos
                        .iter()
                        .flat_map(|field| field.object_references.iter().copied()),
                )
            }))
            .filter_map(|identifier| remap.get(&identifier).map(|target| (identifier, *target)))
            .collect::<Vec<_>>();
        pairs.sort_unstable();
        pairs.dedup();
        Ok(source.clone_with_identity_remap(
            new_identifier,
            &pairs,
            &messages,
            self.archive_limits,
        )?)
    }

    /// Derive a deterministic, collision-checked UUID for a clone.
    pub(in crate::package) fn clone_uuid(
        &mut self,
        source: Option<Uuid>,
        target: u64,
    ) -> Result<Uuid, StructureError> {
        let mut digest = Sha1::new();
        digest.update(b"litchi.keynote.structure.clone.v1");
        if let Some(source) = source {
            digest.update(source.lower.to_le_bytes());
            digest.update(source.upper.to_le_bytes());
        }
        digest.update(target.to_le_bytes());
        let digest = digest.finalize();
        let mut lower = [0_u8; 8];
        let mut upper = [0_u8; 8];
        lower.copy_from_slice(&digest[..8]);
        upper.copy_from_slice(&digest[8..16]);
        let uuid = Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper) | 1,
        };
        if !self.uuids.insert((uuid.lower, uuid.upper)) {
            return Err(StructureError::Registry("UUID collision"));
        }
        Ok(uuid)
    }

    /// Re-encode every staged component and the registry, and reassemble
    /// the exact source ZIP with the changes.
    pub(in crate::package) fn publish(mut self) -> Result<Vec<u8>, StructureError> {
        let registry = self.registry.encode()?;
        let (identifier, index) = self.registry_location;
        let limits = self.archive_limits;
        let snappy_limits = self
            .package
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(|_| StructureError::Registry("Snappy limits"))?;
        {
            let archive = self.archive_mut(METADATA_MEMBER)?;
            let object = archive
                .object_mut(identifier)
                .ok_or(StructureError::Registry("package metadata"))?;
            object.replace_message(
                index,
                RawMessage {
                    type_: PACKAGE_METADATA_MESSAGE_TYPE,
                    data: registry,
                },
                limits,
            )?;
        }
        let mut encoded = Vec::with_capacity(self.changed.len() + self.created.len());
        for (member, archive) in self.changed.iter().chain(self.created.iter()) {
            let bytes = archive.to_bytes(limits)?;
            encoded.push((
                member.clone(),
                SnappyStream::compress(&bytes, snappy_limits)?,
            ));
        }
        let edits = encoded
            .iter()
            .filter(|(member, _)| self.changed.contains_key(member))
            .map(|(member, bytes)| EntryEdit::new(member, bytes))
            .collect::<Vec<_>>();
        let insertions = encoded
            .iter()
            .filter(|(member, _)| self.created.contains_key(member))
            .map(|(member, bytes)| EntryInsertion::new(member, bytes))
            .chain(
                self.members
                    .iter()
                    .map(|(member, bytes)| EntryInsertion::new(member, bytes)),
            )
            .collect::<Vec<_>>();
        let deleted = self.deleted.iter().map(String::as_str).collect::<Vec<_>>();
        self.package
            .structure_reassemble(&insertions, &edits, &deleted)
    }
}

#[cfg(test)]
mod tests {
    use super::super::test_support::{fixture, registry_payload};
    use super::*;

    #[test]
    fn untouched_publish_keeps_the_registry_and_drops_previews() {
        let package = fixture("basic.key");
        let workspace = Workspace::new(&package).expect("workspace");
        assert_eq!(workspace.touched_members(), 1);
        assert_eq!(workspace.created_objects(), 0);
        let bytes = workspace.publish().expect("publishes");
        let candidate = package.structure_reopen(bytes).expect("reopens");
        assert_eq!(registry_payload(&candidate), registry_payload(&package));
        assert_eq!(
            candidate.slides().expect("slides").len(),
            package.slides().expect("slides").len()
        );
        assert!(!candidate.structure_has_member("preview.jpg"));
        assert!(package.structure_has_member("preview.jpg"));
    }

    #[test]
    fn identifiers_and_uuids_are_fresh_and_deterministic() {
        let package = fixture("basic.key");
        let mut workspace = Workspace::new(&package).expect("workspace");
        let watermark = workspace.registry.last_object_identifier();
        let first = workspace.allocate().expect("allocates");
        let second = workspace.allocate().expect("allocates");
        assert!(first > watermark && second == first + 1);
        assert_eq!(workspace.registry.last_object_identifier(), second);
        let source = Uuid { lower: 1, upper: 2 };
        let uuid = workspace.clone_uuid(Some(source), first).expect("derives");
        assert_eq!(
            workspace.clone_uuid(Some(source), first),
            Err(StructureError::Registry("UUID collision"))
        );
        let mut again = Workspace::new(&package).expect("workspace");
        assert_eq!(again.clone_uuid(Some(source), first), Ok(uuid));
        assert_ne!(again.clone_uuid(None, first).expect("derives"), uuid);
    }

    #[test]
    fn staged_objects_are_copy_on_write() {
        let package = fixture("basic.key");
        let mut workspace = Workspace::new(&package).expect("workspace");
        let slide = 2_652_150;
        let member = workspace.component_of(slide).expect("slide member");
        assert_eq!(member, "Index/Slide-2652150.iwa");
        let objects = workspace.member_objects(&member);
        assert!(objects.contains(&slide));
        workspace
            .create_component("Index/Slide-7.iwa")
            .expect("creates");
        assert!(workspace.create_component("Index/Slide-7.iwa").is_err());
        assert!(workspace.create_component(&member).is_err());
        let identifier = workspace.allocate().expect("allocates");
        let remap = HashMap::from([(slide, identifier)]);
        let source = workspace.require(slide).expect("slide").clone();
        let clone = workspace
            .clone_object(&source, identifier, &remap)
            .expect("clones");
        workspace
            .insert("Index/Slide-7.iwa", clone)
            .expect("inserts");
        assert_eq!(
            workspace.component_of(identifier).as_deref(),
            Some("Index/Slide-7.iwa")
        );
        assert_eq!(workspace.created_objects(), 1);
        assert!(package.object(identifier).is_none());
        workspace
            .object_mut(slide)
            .expect("staged")
            .archive_info
            .should_merge = None;
        assert_eq!(workspace.touched_members(), 3);
    }

    #[test]
    fn removed_objects_and_new_members_are_published() {
        let package = fixture("basic.key");
        let mut workspace = Workspace::new(&package).expect("workspace");
        let caption = 2_652_157;
        assert!(workspace.remove_object(caption).is_ok());
        assert!(workspace.object(caption).is_none());
        assert_eq!(
            workspace.remove_object(caption).err(),
            Some(StructureError::MissingObject)
        );
        assert!(workspace.has_member("Index/Document.iwa"));
        workspace
            .insert_member("Data/litchi-test.bin", b"bytes".to_vec())
            .expect("inserts");
        assert!(
            workspace
                .insert_member("Data/litchi-test.bin", Vec::new())
                .is_err()
        );
        assert_eq!(workspace.touched_members(), 3);
        // The publish carries the new member (the slide-drawable family
        // reopens and validates complete candidates).
        let bytes = workspace.publish().expect("publishes");
        assert!(
            bytes
                .windows(20)
                .any(|window| window == b"Data/litchi-test.bin")
        );
    }
}
