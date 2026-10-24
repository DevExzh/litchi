//! Raw OPC edits of a [`Package`] (record 0351 §2.1 rule 4).
//!
//! [`Package::edit_opc`] starts an [`Edit`] over the package's OPC graph;
//! [`Edit::update`] runs one low-level graph operation on a staged copy, and
//! [`Edit::commit`] validates the staged graph as a complete XLSB workbook
//! (the binary workbook main part, the inert VBA graph and a full reparse)
//! and publishes a new package with the reversible part-level [`Patch`] and
//! the [`Diagnostics`]. [`Package::apply_opc`] replays a patch onto an equal
//! package. The source package never changes.
//!
//! Callers that need a typed, semantic operation should prefer the package's
//! scoped edit families; this is the escape hatch for graph operations no
//! family models. A drawing-skipped projection stays drawing-skipped.
//!
//! Signature state is preserved (ADR 0006): a commit that changes a signed
//! graph is refused unless the operation removed the signature explicitly
//! with [`OpcPackage::unsign`] (or the caller called [`Package::unsign`]
//! first); deleting signature parts by hand does not authorize the change.
//! An operation that changes nothing keeps the signature byte for byte.
//!
//! ```
//! use litchi_opc::{BlobPart, PackURI};
//!
//! # fn main() -> litchi_xlsb::package::PackageResult<()> {
//! let package = litchi_xlsb::Package::new()?;
//! let name = PackURI::new("/custom/raw.bin")?;
//! let mut edit = package.edit_opc()?;
//! edit.update(|opc| {
//!     opc.add_part(Box::new(BlobPart::new(
//!         name.clone(),
//!         "application/octet-stream".into(),
//!         vec![1, 2, 3],
//!     )));
//!     Ok(())
//! })?;
//! let (after, patch) = edit.commit()?.into_parts();
//! assert!(after.opc_package().get_part(&name).is_ok());
//! let restored = after.apply_opc(&patch.inverse())?.into_package();
//! assert_eq!(restored.to_bytes()?, package.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use litchi_opc::OpcPackage;

use crate::Package;
use crate::package::error::Result;
use crate::part_edit::{PartPatch, family};

family!("raw OPC");

/// The bounded dimension of raw OPC edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Parts one commit may add, change or remove: at most the package's
    /// source-limit part count.
    Parts,
    /// Retained before and after payload bytes of one captured patch.
    PatchBytes,
}

/// A staged raw OPC edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Package,
    staged: Option<OpcPackage>,
}

impl Package {
    /// Start a raw OPC edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the family entry shape.
    pub fn edit_opc(&self) -> Result<Edit> {
        Ok(Edit {
            source: self.clone(),
            staged: None,
        })
    }

    /// Apply a raw OPC patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, [`Error::Signed`] for a changed patch of a signed
    /// package, and the workbook validation errors of the result.
    pub fn apply_opc(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// The staged OPC graph.
    #[must_use]
    pub fn opc(&self) -> &OpcPackage {
        self.staged
            .as_ref()
            .unwrap_or_else(|| self.source.opc_package())
    }

    /// Run one low-level graph operation on the staged graph.
    ///
    /// Any error leaves the staged graph as it was before the operation.
    ///
    /// # Errors
    ///
    /// Returns the operation's error.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        Ok(value)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged = None;
    }

    /// Validate the staged graph and publish it to a new package. A change
    /// outside the package relationship closure keeps the workbook projection
    /// lazy; a change that can affect that closure is reparsed before return.
    /// An unchanged graph returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the OPC `SignedSourceRequiresExplicitPolicy` error for a
    /// changed signed graph without an explicit signature disposition, the
    /// part-count limit of the patch, and the workbook validation errors of
    /// the staged graph.
    pub fn commit(self) -> Result<Commit> {
        let Some(staged) = self.staged else {
            return Ok(Commit::new((self.source, PartPatch::default())));
        };
        let staged = std::sync::Arc::new(staged);
        let patch = PartPatch::capture_all(self.source.opc_shared(), &staged, None)?;
        if patch.is_noop() {
            return Ok(Commit::new((self.source, patch)));
        }
        staged.validate_signature_edit_from(self.source.opc_package())?;
        crate::Workbook::validate_edit_candidate(&staged)?;
        let package = if patch.projection_unchanged(self.source.opc_package()) {
            self.source.with_unparsed(staged)
        } else {
            self.source.reparse_candidate_retaining_projection(staged)?
        };
        Ok(Commit::new((package, patch)))
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_opc::{BlobPart, PackURI, Relationships, TargetMode};

    use crate::Package;
    use crate::package::PackageError;

    fn raw_part(name: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )
    }

    /// Reopen a newly authored package before comparing replay results. A
    /// reopened package retains its content-types source bytes, whereas an
    /// in-memory builder package has no source provenance to retain. The
    /// projection tests compare the graph and provenance of like-for-like
    /// source-backed packages; their eager/deferred assertions remain about
    /// the workbook projection itself.
    fn source_backed_package() -> Package {
        let package = Package::new().unwrap();
        Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
    }

    type RelationshipSnapshot = (String, String, String, bool);

    #[derive(Debug, PartialEq, Eq)]
    struct PartSnapshot {
        name: String,
        content_type: String,
        blob: Vec<u8>,
        relationships: Vec<RelationshipSnapshot>,
        retained_relationships: Option<(Vec<u8>, bool)>,
    }

    #[derive(Debug, PartialEq, Eq)]
    struct GraphSnapshot {
        content_types_publication: Vec<u8>,
        content_types_source: Option<Vec<u8>>,
        root_relationships: Vec<RelationshipSnapshot>,
        retained_root_relationships: Option<(Vec<u8>, bool)>,
        parts: Vec<PartSnapshot>,
    }

    fn relationship_snapshot(relationships: &Relationships) -> Vec<RelationshipSnapshot> {
        let mut snapshot = relationships
            .iter()
            .map(|relationship| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        snapshot.sort();
        snapshot
    }

    fn retained_relationship_snapshot(
        opc: &litchi_opc::OpcPackage,
        owner: &PackURI,
    ) -> Option<(Vec<u8>, bool)> {
        opc.retained_relationships_xml(owner)
            .unwrap()
            .map(|(bytes, member_present)| (bytes.as_slice().to_vec(), member_present))
    }

    /// Compare semantic part contents and every retained source member while
    /// deliberately ignoring ZIP member order and writer metadata.
    fn graph_snapshot(package: &Package) -> GraphSnapshot {
        let opc = package.opc_package();
        let mut parts = opc
            .iter_parts()
            .map(|metadata| {
                let part = opc.get_part(metadata.partname()).unwrap();
                PartSnapshot {
                    name: metadata.partname().as_str().to_owned(),
                    content_type: metadata.content_type().to_owned(),
                    blob: part.blob().to_vec(),
                    relationships: relationship_snapshot(metadata.rels()),
                    retained_relationships: retained_relationship_snapshot(
                        opc,
                        metadata.partname(),
                    ),
                }
            })
            .collect::<Vec<_>>();
        parts.sort_by(|left, right| left.name.cmp(&right.name));
        let root = PackURI::new("/").unwrap();
        GraphSnapshot {
            content_types_publication: opc
                .source_content_types(opc.read_limits())
                .unwrap()
                .bytes()
                .to_vec(),
            content_types_source: opc
                .content_types_source_bytes()
                .map(|bytes| bytes.as_slice().to_vec()),
            root_relationships: relationship_snapshot(opc.rels()),
            retained_root_relationships: retained_relationship_snapshot(opc, &root),
            parts,
        }
    }

    /// Snapshot every published archive member by name and effective bytes.
    ///
    /// Targeted preservation is allowed to append a changed owner after the
    /// source archive's tail. That physical order is not part of the effective
    /// OPC archive, so compare the complete member set after sorting it while
    /// still requiring every member (including opaque ones) and its payload to
    /// match.
    fn effective_archive_snapshot(package: &Package) -> Vec<(String, Vec<u8>)> {
        let bytes = package.to_bytes().unwrap();
        let reader = soapberry_zip::office::ArchiveReader::new(
            &bytes,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        let mut members = reader
            .file_names()
            .map(|name| (name.to_owned(), reader.read(name).unwrap()))
            .collect::<Vec<_>>();
        members.sort_by(|left, right| left.0.cmp(&right.0));
        members
    }

    fn assert_bytes_equal(label: &str, left: &[u8], right: &[u8]) {
        if left == right {
            return;
        }
        let first_difference = left
            .iter()
            .zip(right)
            .position(|(left, right)| left != right);
        panic!(
            "{label} differs: left_len={}, right_len={}, first_difference={first_difference:?}",
            left.len(),
            right.len()
        );
    }

    fn assert_graph_equal(left: &Package, right: &Package) {
        assert_graph_equal_with_source_policy(left, right, false, &[]);
    }

    /// Compare a replayed graph with a directly committed graph. A structural
    /// content-types edit can leave source members from different earlier
    /// generations on the direct candidate and replay target. The publication
    /// and complete serialized archive must still be identical; source-member
    /// lineage itself is not required to equal the current publication.
    fn assert_replayed_graph_equal(left: &Package, right: &Package) {
        assert_graph_equal_with_source_policy(left, right, true, &[]);
        assert_eq!(left.to_bytes().unwrap(), right.to_bytes().unwrap());
    }

    /// The transitive test changes the package root and the relationship
    /// member of the newly reachable parent before the final replay. Those
    /// changed owners may carry different historical source captures, while
    /// untouched part and relationship members must remain byte-identical.
    fn assert_replayed_graph_equal_with_manifest_lineage(
        left: &Package,
        right: &Package,
        changed_source_owners: &[&str],
    ) {
        assert_graph_equal_with_source_policy(left, right, true, changed_source_owners);
        assert_eq!(
            effective_archive_snapshot(left),
            effective_archive_snapshot(right),
            "complete effective archive",
        );
    }

    fn assert_graph_equal_with_source_policy(
        left: &Package,
        right: &Package,
        allow_effective_content_types_source: bool,
        changed_source_owners: &[&str],
    ) {
        let left = graph_snapshot(left);
        let right = graph_snapshot(right);
        assert_bytes_equal(
            "content-types publication",
            &left.content_types_publication,
            &right.content_types_publication,
        );
        assert_eq!(
            left.content_types_source.is_some(),
            right.content_types_source.is_some(),
            "content-types source presence",
        );
        if !allow_effective_content_types_source
            && let (Some(left), Some(right)) = (
                left.content_types_source.as_ref(),
                right.content_types_source.as_ref(),
            )
        {
            assert_bytes_equal("content-types source", left, right);
        }
        assert_eq!(left.root_relationships, right.root_relationships);
        if !changed_source_owners.contains(&"/") {
            assert_eq!(
                left.retained_root_relationships.is_some(),
                right.retained_root_relationships.is_some(),
                "root relationships source presence",
            );
            if let (Some((left, left_present)), Some((right, right_present))) = (
                left.retained_root_relationships.as_ref(),
                right.retained_root_relationships.as_ref(),
            ) {
                assert_eq!(left_present, right_present);
                assert_bytes_equal("root relationships source", left, right);
            }
        }
        assert_eq!(left.parts.len(), right.parts.len());
        for (part_left, part_right) in left.parts.iter().zip(&right.parts) {
            assert_eq!(
                (part_left.name.as_str(), part_left.content_type.as_str()),
                (part_right.name.as_str(), part_right.content_type.as_str())
            );
            assert_bytes_equal(
                &format!("part {}", part_left.name),
                &part_left.blob,
                &part_right.blob,
            );
            assert_eq!(part_left.relationships, part_right.relationships);
            if !changed_source_owners.contains(&part_left.name.as_str()) {
                assert_eq!(
                    part_left.retained_relationships.is_some(),
                    part_right.retained_relationships.is_some(),
                    "{} relationships source presence",
                    part_left.name,
                );
                if let (Some((left, left_present)), Some((right, right_present))) = (
                    part_left.retained_relationships.as_ref(),
                    part_right.retained_relationships.as_ref(),
                ) {
                    assert_eq!(left_present, right_present);
                    assert_bytes_equal(
                        &format!("{} relationships source", part_left.name),
                        left,
                        right,
                    );
                }
            }
        }
    }

    fn lexical_content_types_edit(
        opc: &mut litchi_opc::OpcPackage,
    ) -> crate::package::error::Result<()> {
        let limits = opc.read_limits();
        let source = opc.source_content_types(limits)?;
        let marker = b"</Types>";
        let at = source
            .bytes()
            .windows(marker.len())
            .rposition(|window| window == marker)
            .expect("content-types root close");
        let mut replacement = source.bytes().to_vec();
        replacement.splice(at..at, [b'\n']);
        opc.try_replace_content_types_bytes(source.bytes(), Arc::new(replacement), limits)?;
        Ok(())
    }

    fn lexical_root_relationship_edit(
        opc: &mut litchi_opc::OpcPackage,
    ) -> crate::package::error::Result<()> {
        let owner = PackURI::new("/")?;
        let limits = opc.read_limits();
        let source = opc.authored_relationships_xml(&owner, limits)?;
        let marker = b"</Relationships>";
        let at = source
            .as_slice()
            .windows(marker.len())
            .rposition(|window| window == marker)
            .expect("relationships root close");
        let mut replacement = source.as_slice().to_vec();
        replacement.splice(at..at, [b'\n']);
        opc.try_replace_relationships_source_bytes(
            &owner,
            source.as_slice(),
            true,
            Arc::new(replacement),
            true,
            limits,
        )?;
        Ok(())
    }

    #[test]
    fn starting_an_edit_copies_nothing_and_a_no_op_returns_the_source() {
        let package = Package::new().unwrap();
        let edit = package.edit_opc().unwrap();
        assert!(edit.staged.is_none());
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.package().to_bytes().unwrap(),
            package.to_bytes().unwrap()
        );
    }

    #[test]
    fn root_relationships_are_captured_replayed_and_inverted() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_opc().unwrap();
        edit.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/raw.bin")));
            opc.rels_mut().try_add_relationship(
                "urn:litchi:test".into(),
                "custom/raw.bin".into(),
                "rIdRaw".into(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(after.opc_package().rels().get("rIdRaw").is_some());
        assert!(package.opc_package().rels().get("rIdRaw").is_none());
        let replayed = package.apply_opc(&patch).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert!(restored.opc_package().rels().get("rIdRaw").is_none());
        assert_eq!(restored.to_bytes().unwrap(), package.to_bytes().unwrap());
        assert!(matches!(
            after.apply_opc(&patch),
            Err(PackageError::PatchConflict { .. })
        ));
    }

    #[test]
    fn an_unreferenced_part_defers_projection_and_replays_after_reopen() {
        let package = source_backed_package();
        let before_bytes = package.to_bytes().unwrap();
        let mut edit = package.edit_opc().unwrap();
        edit.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/unreferenced.bin")));
            Ok(())
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(!after.projection_cached());

        let after_bytes = after.to_bytes().unwrap();
        assert!(after.workbook().is_ok());
        assert!(after.projection_cached());

        let reopened_before = Package::from_bytes(&before_bytes).unwrap();
        assert_graph_equal(&reopened_before, &package);
        let replayed = reopened_before.apply_opc(&patch).unwrap().into_package();
        assert!(!replayed.projection_cached());
        assert_replayed_graph_equal(&replayed, &after);
        assert!(replayed.workbook().is_ok());

        let reopened_after = Package::from_bytes(&after_bytes).unwrap();
        let restored = reopened_after
            .apply_opc(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(!restored.projection_cached());
        assert_replayed_graph_equal(&restored, &package);
        assert!(restored.workbook().is_ok());
    }

    #[test]
    fn a_recipient_edge_to_a_changed_part_forces_projection_reparse() {
        let package = Package::new().unwrap();
        let name = PackURI::new("/custom/retyped-unreferenced.bin").unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part(name.as_str())));
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();
        assert!(!source.projection_cached());

        let mut changed = source.edit_opc().unwrap();
        changed
            .update(|opc| {
                opc.get_part_mut(&name)?.set_blob(vec![4, 5, 6]);
                Ok(())
            })
            .unwrap();
        let (after, patch) = changed.commit().unwrap().into_parts();
        assert!(!after.projection_cached());

        // This recipient has an unrelated inbound edge to the patch's
        // changed part. The patch's recorded base did not, so only the
        // recipient-before closure can reveal that replay must reparse.
        let mut recipient_edit = source.edit_opc().unwrap();
        recipient_edit
            .update(|opc| {
                opc.rels_mut().try_add_relationship(
                    "urn:litchi:test".into(),
                    name.as_str().into(),
                    "rIdRecipientEdge".into(),
                    TargetMode::Internal,
                )?;
                Ok(())
            })
            .unwrap();
        let (recipient, _) = recipient_edit.commit().unwrap().into_parts();
        assert!(recipient.projection_cached());

        let replayed = recipient.apply_opc(&patch).unwrap().into_package();
        assert!(replayed.projection_cached());
        assert_eq!(
            replayed.opc_package().get_part(&name).unwrap().blob(),
            &[4, 5, 6]
        );
    }

    #[test]
    fn a_transitively_referenced_part_keeps_eager_projection_validation() {
        let package = source_backed_package();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            let name = PackURI::new("/custom/transitive-root.bin")?;
            opc.add_part(Box::new(raw_part(name.as_str())));
            opc.rels_mut().try_add_relationship(
                "urn:litchi:test".into(),
                name.as_str().into(),
                "rIdTransitiveRoot".into(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();
        assert!(source.projection_cached());

        let mut edit = source.edit_opc().unwrap();
        edit.update(|opc| {
            let child = PackURI::new("/custom/transitive-child.bin")?;
            opc.add_part(Box::new(raw_part(child.as_str())));
            let parent = PackURI::new("/custom/transitive-root.bin")?;
            opc.get_part_mut(&parent)?.rels_mut().try_add_relationship(
                "urn:litchi:test".into(),
                child.as_str().into(),
                "rIdTransitiveChild".into(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(after.projection_cached());

        let replayed = source.apply_opc(&patch).unwrap().into_package();
        assert!(replayed.projection_cached());
        assert_replayed_graph_equal_with_manifest_lineage(
            &replayed,
            &after,
            &["/", "/custom/transitive-root.bin"],
        );

        let reopened = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
        let reopened_replayed = reopened.apply_opc(&patch).unwrap().into_package();
        assert!(reopened_replayed.projection_cached());
        assert_replayed_graph_equal_with_manifest_lineage(
            &reopened_replayed,
            &after,
            &["/", "/custom/transitive-root.bin"],
        );

        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert!(restored.projection_cached());
        assert_replayed_graph_equal_with_manifest_lineage(
            &restored,
            &source,
            &["/", "/custom/transitive-root.bin"],
        );

        let reopened_after = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let reopened_restored = reopened_after
            .apply_opc(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(reopened_restored.projection_cached());
        assert_replayed_graph_equal_with_manifest_lineage(
            &reopened_restored,
            &source,
            &["/", "/custom/transitive-root.bin"],
        );
    }

    #[test]
    fn lexical_content_types_edit_is_replayed_and_inverted() {
        let package = Package::new().unwrap();
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit_opc().unwrap();
        edit.update(lexical_content_types_edit).unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(!patch.is_noop());

        let replayed = package.apply_opc(&patch).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert_eq!(restored.to_bytes().unwrap(), before);

        let mut divergent = after.edit_opc().unwrap();
        divergent
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/inverse-unrelated.bin")));
                Ok(())
            })
            .unwrap();
        let (divergent, _) = divergent.commit().unwrap().into_parts();
        assert!(matches!(
            divergent.apply_opc(&patch.inverse()),
            Err(PackageError::PatchConflict {
                part: "content types"
            })
        ));
    }

    #[test]
    fn lexical_root_relationship_edit_is_replayed_and_inverted_after_reopen() {
        let package = Package::new().unwrap();
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit_opc().unwrap();
        edit.update(lexical_root_relationship_edit).unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(!patch.is_noop());

        let replayed = package.apply_opc(&patch).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let reopened = Package::from_bytes(&before).unwrap();
        let reopened_replayed = reopened.apply_opc(&patch).unwrap().into_package();
        assert_eq!(
            reopened_replayed.to_bytes().unwrap(),
            after.to_bytes().unwrap()
        );

        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert_eq!(restored.to_bytes().unwrap(), before);
    }

    #[test]
    fn invalid_raw_content_types_source_is_retained_for_graph_replay() {
        let package = Package::new().unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/source.bin")));
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();
        let source = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();

        let mut edit = source.edit_opc().unwrap();
        let replacement = edit
            .update(|opc| {
                let limits = opc.read_limits();
                let current = opc.source_content_types(limits)?;
                let marker = b"application/octet-stream";
                let at = current
                    .bytes()
                    .windows(marker.len())
                    .position(|window| window == marker)
                    .expect("authored content type");
                let mut bytes = current.bytes().to_vec();
                bytes.splice(at..at + marker.len(), b"application/custom".iter().copied());
                opc.try_replace_content_types_bytes(
                    current.bytes(),
                    Arc::new(bytes.clone()),
                    limits,
                )?;
                Ok(bytes)
            })
            .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        let source_bytes = after
            .opc_package()
            .content_types_source_bytes()
            .expect("raw source allocation");
        assert_eq!(source_bytes.as_slice(), replacement.as_slice());

        let replayed = patch.0.apply_graph(source.opc_package()).unwrap();
        assert_eq!(
            replayed
                .content_types_source_bytes()
                .expect("replayed raw source allocation")
                .as_slice(),
            replacement.as_slice()
        );
        let reopened_source = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
        let reopened_replayed = patch
            .0
            .apply_graph(reopened_source.opc_package())
            .expect("reopened forward source replay");
        assert_eq!(
            Package::from_opc(reopened_replayed)
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        let restored = patch
            .0
            .inverse()
            .apply_graph(after.opc_package())
            .expect("inverse source replay");
        let restored_bytes = restored
            .content_types_source_bytes()
            .expect("restored source allocation");
        assert_eq!(
            restored_bytes.as_slice(),
            source
                .opc_package()
                .content_types_source_bytes()
                .expect("original source allocation")
                .as_slice()
        );
        let reopened_after = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let reopened_restored = patch
            .0
            .inverse()
            .apply_graph(reopened_after.opc_package())
            .expect("reopened inverse source replay");
        assert_eq!(
            Package::from_opc(reopened_restored)
                .unwrap()
                .to_bytes()
                .unwrap(),
            source.to_bytes().unwrap()
        );
    }

    #[test]
    fn content_types_replay_accepts_matching_publication_with_stale_source() {
        let package = Package::new().unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/source.bin")));
            Ok(())
        })
        .unwrap();
        let (valid, _) = seed.commit().unwrap().into_parts();

        // Leave a valid typed graph paired with a source manifest whose
        // declaration no longer covers it.  Publication therefore falls back
        // to the same authored bytes as `valid`, while provenance differs.
        let mut make_stale = valid.edit_opc().unwrap();
        make_stale
            .update(|opc| {
                let limits = opc.read_limits();
                let current = opc.source_content_types(limits)?;
                let marker = b"application/octet-stream";
                let at = current
                    .bytes()
                    .windows(marker.len())
                    .position(|window| window == marker)
                    .expect("authored content type");
                let mut bytes = current.bytes().to_vec();
                bytes.splice(at..at + marker.len(), b"application/custom".iter().copied());
                opc.try_replace_content_types_bytes(current.bytes(), Arc::new(bytes), limits)?;
                Ok(())
            })
            .unwrap();
        let (stale, _) = make_stale.commit().unwrap().into_parts();
        assert!(
            stale
                .opc_package()
                .retained_content_types_xml()
                .unwrap()
                .is_none()
        );

        // The patch starts from the stale-source package and changes only the
        // lexical publication.  Replaying it onto the independently retained
        // publication must use the recorded content base, not the stale source
        // allocation as an additional conflict key.
        let mut edit = stale.edit_opc().unwrap();
        edit.update(lexical_content_types_edit).unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        let replayed = valid
            .apply_opc(&patch)
            .expect("matching publication with stale source")
            .into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
    }

    #[test]
    fn mixed_lexical_content_types_edit_refuses_an_unrelated_manifest_change() {
        let package = Package::new().unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/source.bin")));
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();

        let mut edit = source.edit_opc().unwrap();
        edit.update(|opc| {
            let name = PackURI::new("/custom/source.bin")?;
            opc.get_part_mut(&name)?
                .set_content_type("application/custom".to_owned())?;
            lexical_content_types_edit(opc)
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        let replayed = source.apply_opc(&patch).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let reopened_source = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
        let reopened_replayed = reopened_source
            .apply_opc(&patch)
            .expect("reopened mixed forward replay")
            .into_package();
        assert_eq!(
            reopened_replayed.to_bytes().unwrap(),
            after.to_bytes().unwrap()
        );
        let reopened_after = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let reopened_restored = reopened_after
            .apply_opc(&patch.inverse())
            .expect("reopened mixed inverse replay")
            .into_package();
        assert_eq!(
            reopened_restored.to_bytes().unwrap(),
            source.to_bytes().unwrap()
        );

        let mut divergent = source.edit_opc().unwrap();
        divergent
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/unrelated.bin")));
                Ok(())
            })
            .unwrap();
        let (divergent, _) = divergent.commit().unwrap().into_parts();
        assert!(matches!(
            divergent.apply_opc(&patch),
            Err(PackageError::PatchConflict {
                part: "content types"
            })
        ));
    }

    #[test]
    fn structural_content_type_patch_composes_with_an_unrelated_part() {
        let package = Package::new().unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/source.bin")));
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();

        let mut edit = source.edit_opc().unwrap();
        edit.update(|opc| {
            let name = PackURI::new("/custom/source.bin")?;
            opc.get_part_mut(&name)?
                .set_content_type("application/custom".to_owned())?;
            Ok(())
        })
        .unwrap();
        let (_, patch) = edit.commit().unwrap().into_parts();

        let mut divergent = source.edit_opc().unwrap();
        divergent
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/unrelated.bin")));
                Ok(())
            })
            .unwrap();
        let (divergent, _) = divergent.commit().unwrap().into_parts();
        let applied = divergent.apply_opc(&patch).unwrap().into_package();
        assert_eq!(
            applied
                .opc_package()
                .get_part(&PackURI::new("/custom/source.bin").unwrap())
                .unwrap()
                .content_type(),
            "application/custom"
        );
        assert!(
            applied
                .opc_package()
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn structural_inverse_composes_with_an_unrelated_part_after_a_retained_source() {
        let package = Package::new().unwrap();
        let mut seed = package.edit_opc().unwrap();
        seed.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/source.bin")));
            Ok(())
        })
        .unwrap();
        let (source, _) = seed.commit().unwrap().into_parts();

        let mut lexical = source.edit_opc().unwrap();
        lexical.update(lexical_content_types_edit).unwrap();
        let (source, _) = lexical.commit().unwrap().into_parts();

        let mut edit = source.edit_opc().unwrap();
        edit.update(|opc| {
            let name = PackURI::new("/custom/source.bin")?;
            opc.get_part_mut(&name)?
                .set_content_type("application/custom".to_owned())?;
            Ok(())
        })
        .unwrap();
        let (changed, patch) = edit.commit().unwrap().into_parts();

        let mut divergent = changed.edit_opc().unwrap();
        divergent
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/unrelated.bin")));
                Ok(())
            })
            .unwrap();
        let (divergent, _) = divergent.commit().unwrap().into_parts();
        let restored = divergent
            .apply_opc(&patch.inverse())
            .unwrap()
            .into_package();
        assert_eq!(
            restored
                .opc_package()
                .get_part(&PackURI::new("/custom/source.bin").unwrap())
                .unwrap()
                .content_type(),
            "application/octet-stream"
        );
        assert!(
            restored
                .opc_package()
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn a_failed_update_and_reset_leave_the_staging_as_it_was() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_opc().unwrap();
        let error = edit
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/a.bin")));
                Err::<(), _>(PackageError::InvalidFormat("stop".into()))
            })
            .unwrap_err();
        assert!(matches!(error, PackageError::InvalidFormat(_)));
        assert!(
            edit.opc()
                .get_part(&PackURI::new("/custom/a.bin").unwrap())
                .is_err()
        );
        edit.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/b.bin")));
            Ok(())
        })
        .unwrap();
        edit.reset();
        assert!(edit.commit().unwrap().patch().is_noop());
    }
}
