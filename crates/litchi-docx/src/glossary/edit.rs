//! The glossary family (record 0355): the glossary document (`AutoText`
//! and other building blocks) and its auxiliary graph.
//!
//! [`Package::glossary`](crate::Package::glossary) reads the typed catalog
//! and [`Package::glossary_graph`](crate::Package::glossary_graph) the raw
//! graph; [`Package::edit_glossary`](crate::Package::edit_glossary) stages an
//! [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_glossary`](crate::Package::apply_glossary) replays its
//! [`Patch`].
//!
//! ```
//! use litchi_docx::{Package, glossary};
//!
//! let package = Package::new_template()?;
//! let mut catalog = glossary::Catalog::new();
//! catalog.add(glossary::Entry::new(
//!     "Signature",
//!     br#"<w:docPartBody xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:docPartBody>"#.to_vec(),
//!     &glossary::Limits::DEFAULT,
//! )?)?;
//! let mut edit = package.edit_glossary();
//! edit.put(catalog, glossary::Conformance::Transitional)?;
//! let commit = edit.commit()?;
//! let (catalog, _) = commit.package().glossary()?.unwrap();
//! assert_eq!(catalog.entries()[0].name(), Some("Signature"));
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use super::codec::invalid;
use super::graph::{self, ClosureLedger};
use super::transaction::Transaction as Draft;
use super::{Catalog, Conformance, Entry, MAX_VALUES, Name, raw};
use crate::package::family::{Staged, graph_family};

graph_family!("glossary", own_limits);

pub use super::limit::LimitKind;

/// A staged glossary change over one package snapshot. The source is never
/// changed: [`Self::commit`] publishes a new package, and dropping the edit
/// discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    draft: Option<Draft>,
    limits: super::Limits,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

macro_rules! draft_verbs {
    ($($(#[$meta:meta])* fn $name:ident(&mut self $(, $arg:ident: $ty:ty)*) -> $ret:ty;)*) => {$(
        $(#[$meta])*
        ///
        /// # Errors
        ///
        /// Returns an error for a change the glossary graph refuses or a
        /// glossary part that cannot be read.
        pub fn $name(&mut self $(, $arg: $ty)*) -> Result<$ret> {
            self.draft()?.$name($($arg),*)
        }
    )*};
}

impl Edit {
    /// The package with the changes staged before the pending catalog and
    /// graph verbs (those are published by [`Self::commit`]).
    #[must_use]
    pub const fn package(&self) -> &crate::Package {
        self.staged.working()
    }

    fn draft(&mut self) -> Result<&mut Draft> {
        if self.draft.is_none() {
            self.draft = Some(Draft::new_with_limits(
                self.staged.working().opc_package(),
                &self.limits,
            )?);
        }
        Ok(self
            .draft
            .as_mut()
            .expect("glossary draft was created above"))
    }

    fn flush(&mut self) -> Result<()> {
        if let Some(draft) = self.draft.take() {
            let mut working = self.staged.working().clone();
            working.edit_semantic_opc("glossary.edit", |opc| draft.commit(opc).map(|_| ()))?;
            *self.staged.working_mut() = working;
        }
        Ok(())
    }

    fn update<T>(&mut self, verb: impl FnOnce(&mut crate::Package) -> Result<T>) -> Result<T> {
        self.flush()?;
        let mut working = self.staged.working().clone();
        let value = verb(&mut working)?;
        *self.staged.working_mut() = working;
        Ok(value)
    }

    /// The staged typed catalog.
    ///
    /// # Errors
    ///
    /// Returns an error for a glossary part that cannot be read.
    pub fn catalog(&mut self) -> Result<Option<&Catalog>> {
        Ok(self.draft()?.catalog())
    }

    /// Apply an atomic mutation to the staged typed catalog.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or a catalog the glossary refuses; the
    /// staged catalog is then unchanged.
    pub fn edit_catalog(
        &mut self,
        edit: impl FnOnce(&mut Catalog) -> Result<()>,
    ) -> Result<&mut Self> {
        self.draft()?.edit_catalog(edit)?;
        Ok(self)
    }

    draft_verbs! {
        /// Replace the staged typed catalog; returns whether it changed.
        fn replace_catalog(&mut self, value: Catalog) -> bool;
        /// Append an entry; returns its position.
        fn add_entry(&mut self, value: Entry) -> usize;
        /// Insert or replace the entry with the same name; returns the
        /// replaced entry.
        fn put_entry(&mut self, value: Entry) -> Option<Entry>;
        /// Replace the entry named `name`; returns the replaced entry.
        fn replace_entry(&mut self, name: &str, value: Entry) -> Option<Entry>;
        /// Rename the entry named `name`; returns whether it existed.
        fn rename_entry(&mut self, name: &str, value: Name) -> bool;
        /// Remove the entry named `name`; returns it.
        fn remove_entry(&mut self, name: &str) -> Option<Entry>;
        /// Move an entry; returns whether the order changed.
        fn move_entry(&mut self, from: usize, to: usize) -> bool;
        /// Remove every entry; returns how many were removed.
        fn clear_entries(&mut self) -> usize;
        /// Replace the glossary document's relationships; returns whether
        /// they changed.
        fn set_relationships(&mut self, relationships: Vec<raw::Rel>) -> bool;
        /// Add a glossary document relationship; returns whether it was new.
        fn add_relationship(&mut self, relationship: raw::Rel) -> bool;
        /// Remove a glossary document relationship; returns it.
        fn remove_relationship(&mut self, id: &str) -> Option<raw::Rel>;
        /// Add an auxiliary part; returns whether it was new.
        fn add_part(&mut self, part: raw::Part) -> bool;
        /// Replace an auxiliary part's payload; returns whether it changed.
        fn replace_part_data(&mut self, name: &str, data: Vec<u8>) -> bool;
        /// Replace an auxiliary part's relationships; returns whether they
        /// changed.
        fn set_part_relationships(&mut self, name: &str, relationships: Vec<raw::Rel>) -> bool;
        /// Remove an auxiliary part; returns it.
        fn remove_part(&mut self, name: &str) -> Option<raw::Part>;
    }

    /// Replace an auxiliary part's content type, payload and relationships;
    /// returns whether it changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a change the glossary graph refuses.
    pub fn replace_part(
        &mut self,
        name: &str,
        content_type: impl AsRef<str>,
        data: Vec<u8>,
        relationships: Vec<raw::Rel>,
    ) -> Result<bool> {
        self.draft()?
            .replace_part(name, content_type, data, relationships)
    }

    /// Move a complete typed catalog into the package; returns whether the
    /// package changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid catalog and the signed-package refusal
    /// for a change to a signed package.
    pub fn put(&mut self, catalog: Catalog, conformance: Conformance) -> Result<bool> {
        self.update(|package| package.put_glossary(catalog, conformance))
    }

    /// Publish a complete low-level glossary graph; returns whether the
    /// package changed.
    ///
    /// # Errors
    ///
    /// As [`Self::put`].
    pub fn put_graph(&mut self, graph: &raw::Graph) -> Result<bool> {
        self.update(|package| package.put_glossary_graph(graph))
    }

    /// Remove and return the complete low-level glossary graph.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn take_graph(&mut self) -> Result<Option<raw::Graph>> {
        self.update(crate::Package::take_glossary_graph)
    }

    /// Remove the complete glossary-owned graph; returns whether it was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.update(crate::Package::remove_glossary)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Validate the staged changes and publish them as a new package.
    /// Nothing staged publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale or refused graph and the signed-package
    /// refusal for a change to a signed package.
    pub fn commit(mut self) -> Result<Commit> {
        self.flush()?;
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a glossary change over this package. O(1): the edit shares the
    /// package until a verb changes it.
    #[must_use]
    pub fn edit_glossary(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            draft: None,
            limits: self.semantic_limits().glossary_value(),
        }
    }

    /// Replay a glossary patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_glossary(&self, patch: &Patch) -> Result<Commit> {
        let limits = self.semantic_limits().glossary_value();
        patch.0.check(self.opc_package(), "apply_glossary")?;
        let source_owner = graph::locate(self.opc_package(), &limits)?;
        let main_name = self.opc_package().main_document_part()?.partname().clone();
        let source_owner_state = source_owner.as_ref().map(|owner| OwnerState {
            root: owner.root.as_str().to_owned(),
            id: owner.relationship_id.clone(),
            target: owner.relationship_target.clone(),
        });
        // Admit the current owner graph and check the patch base before any
        // graph allocation. The source closure ledger is retained so changed
        // parts can be applied as deltas without charging unrelated package
        // content or relationship sets.
        let (owned, mut ledger) =
            graph::admit_owner_closure(self.opc_package(), source_owner.as_ref(), &limits)?;
        let source_seed =
            owner_metadata_seed_for(main_name.as_str(), source_owner_state.as_ref(), &limits)?;
        let mut final_owner = None;
        let mut changed_main_owner = None;
        patch.0.for_each_part_delta(|name, _, after| {
            if name == &main_name {
                changed_main_owner = Some(
                    after
                        .map(|(_, _, relationships)| {
                            owner_state_from_relationships(relationships, &limits)
                        })
                        .transpose()?
                        .flatten(),
                );
            }
            Ok(())
        })?;
        patch
            .0
            .validate_after_relationships(|owner, relationships| {
                if owner == Some(&main_name) {
                    final_owner = Some(owner_state_from_relationships(relationships, &limits)?);
                }
                Ok(())
            })?;
        let final_owner = changed_main_owner
            .or(final_owner)
            .unwrap_or_else(|| source_owner_state.clone());
        patch.0.for_each_part_delta(|name, before, after| {
            if name == &main_name {
                return Ok(());
            }
            let was_owned = owned.contains(name);
            if was_owned {
                if let Some((content_type, data, relationships)) = before {
                    let old = patch_part_ledger(
                        name,
                        content_type,
                        data,
                        relationships,
                        source_owner_state
                            .as_ref()
                            .is_some_and(|owner| owner.root == name.as_str()),
                        &limits,
                        false,
                    )?;
                    subtract_ledger(&mut ledger, old)?;
                }
            }
            if let Some((content_type, data, relationships)) = after
                && (was_owned || before.is_none())
            {
                let new = patch_part_ledger(
                    name,
                    content_type,
                    data,
                    relationships,
                    final_owner
                        .as_ref()
                        .is_some_and(|owner| owner.root == name.as_str()),
                    &limits,
                    true,
                )?;
                add_ledger(&mut ledger, new, &limits)?;
            }
            Ok(())
        })?;
        let final_seed =
            owner_metadata_seed_for(main_name.as_str(), final_owner.as_ref(), &limits)?;
        ledger.metadata_bytes = ledger
            .metadata_bytes
            .checked_sub(source_seed)
            .ok_or_else(|| invalid("glossary source metadata ledger underflow"))?;
        ledger.metadata_bytes = ledger
            .metadata_bytes
            .checked_add(final_seed)
            .ok_or_else(|| invalid("glossary final metadata ledger overflow"))?;
        check_ledger(&ledger, &limits)?;
        let commit = apply_patch(self, patch, "apply_glossary")?;
        commit.package().glossary_graph()?;
        Ok(commit)
    }
}

fn admit_patch_relationships(
    relationships: &litchi_opc::rel::Relationships,
    limits: &super::Limits,
    count: &mut usize,
    metadata: &mut usize,
) -> Result<()> {
    if relationships.len() > limits.values() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::Relationship,
            actual: relationships.len(),
            maximum: limits.values(),
        });
    }
    *count = count
        .checked_add(relationships.len())
        .ok_or(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphWideRelationship,
            actual: usize::MAX,
            maximum: limits.values(),
        })?;
    if *count > limits.values() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphWideRelationship,
            actual: *count,
            maximum: limits.values(),
        });
    }
    for relationship in relationships.iter() {
        graph::validate_relationship_metadata(
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        )?;
        for value in [
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        ] {
            if value.len() > limits.string_bytes() {
                return Err(crate::Error::GlossaryLimit {
                    kind: LimitKind::StringBytes,
                    actual: value.len(),
                    maximum: limits.string_bytes(),
                });
            }
            *metadata = metadata
                .checked_add(value.len())
                .ok_or(crate::Error::GlossaryLimit {
                    kind: LimitKind::GraphMetadataBytes,
                    actual: usize::MAX,
                    maximum: limits.metadata_bytes(),
                })?;
        }
    }
    if *metadata > limits.metadata_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphMetadataBytes,
            actual: *metadata,
            maximum: limits.metadata_bytes(),
        });
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OwnerState {
    root: String,
    id: String,
    target: String,
}

fn owner_state_from_relationships(
    relationships: &litchi_opc::rel::Relationships,
    limits: &super::Limits,
) -> Result<Option<OwnerState>> {
    let mut owner = None;
    for relationship in relationships
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), super::REL | super::STRICT_REL))
    {
        if owner.is_some() {
            return Err(invalid("main document has multiple glossary relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("glossary relationship cannot be external"));
        }
        graph::validate_relationship_metadata(
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        )?;
        if relationship.r_id().len() > limits.string_bytes()
            || relationship.target_ref().len() > limits.string_bytes()
        {
            return Err(crate::Error::GlossaryLimit {
                kind: LimitKind::StringBytes,
                actual: relationship
                    .r_id()
                    .len()
                    .max(relationship.target_ref().len()),
                maximum: limits.string_bytes(),
            });
        }
        let root_part = relationship.target_partname()?;
        if root_part.as_str().len() > limits.name_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: LimitKind::StringBytes,
                actual: root_part.as_str().len(),
                maximum: limits.name_bytes(),
            });
        }
        let root = root_part.as_str().to_owned();
        owner = Some(OwnerState {
            root,
            id: relationship.r_id().to_owned(),
            target: relationship.target_ref().to_owned(),
        });
    }
    Ok(owner)
}

fn owner_metadata_seed_for(
    main_name: &str,
    owner: Option<&OwnerState>,
    limits: &super::Limits,
) -> Result<usize> {
    let Some(owner) = owner else {
        return Ok(0);
    };
    let mut metadata = 0usize;
    for value in [main_name, owner.root.as_str()] {
        if value.len() > limits.name_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: LimitKind::StringBytes,
                actual: value.len(),
                maximum: limits.name_bytes(),
            });
        }
        graph::add_selected_metadata(&mut metadata, value, limits)?;
    }
    for value in [owner.id.as_str(), owner.target.as_str()] {
        if value.len() > limits.string_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: LimitKind::StringBytes,
                actual: value.len(),
                maximum: limits.string_bytes(),
            });
        }
        graph::add_selected_metadata(&mut metadata, value, limits)?;
    }
    Ok(metadata)
}

fn patch_part_ledger(
    name: &litchi_opc::PackURI,
    content_type: &str,
    data: &[u8],
    relationships: &litchi_opc::rel::Relationships,
    is_root: bool,
    limits: &super::Limits,
    parse_xml: bool,
) -> Result<ClosureLedger> {
    if name.as_str().len() > limits.name_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::StringBytes,
            actual: name.as_str().len(),
            maximum: limits.name_bytes(),
        });
    }
    if content_type.len() > limits.string_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::StringBytes,
            actual: content_type.len(),
            maximum: limits.string_bytes(),
        });
    }
    if data.len() > limits.graph_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphBytes,
            actual: data.len(),
            maximum: limits.graph_bytes(),
        });
    }
    let mut metadata_bytes = 0usize;
    graph::add_selected_metadata(&mut metadata_bytes, name.as_str(), limits)?;
    graph::add_selected_metadata(&mut metadata_bytes, content_type, limits)?;
    let mut relationship_count = 0usize;
    admit_patch_relationships(
        relationships,
        limits,
        &mut relationship_count,
        &mut metadata_bytes,
    )?;
    if parse_xml {
        if content_type == super::CT {
            super::codec::read_with_limits(data, limits)?;
        } else if content_type.ends_with("+xml")
            || content_type.eq_ignore_ascii_case("application/xml")
            || content_type.eq_ignore_ascii_case("text/xml")
        {
            raw::admit_xml(data, limits)?;
        }
    }
    Ok(ClosureLedger {
        total_bytes: data.len(),
        metadata_bytes,
        relationship_count,
        part_count: usize::from(!is_root),
    })
}

fn add_ledger(
    target: &mut ClosureLedger,
    value: ClosureLedger,
    limits: &super::Limits,
) -> Result<()> {
    target.total_bytes = target
        .total_bytes
        .checked_add(value.total_bytes)
        .ok_or_else(|| invalid("glossary final graph ledger overflow"))?;
    target.metadata_bytes = target
        .metadata_bytes
        .checked_add(value.metadata_bytes)
        .ok_or_else(|| invalid("glossary final metadata ledger overflow"))?;
    target.relationship_count = target
        .relationship_count
        .checked_add(value.relationship_count)
        .ok_or_else(|| invalid("glossary final relationship ledger overflow"))?;
    target.part_count = target
        .part_count
        .checked_add(value.part_count)
        .ok_or_else(|| invalid("glossary final part ledger overflow"))?;
    check_ledger(target, limits)
}

fn subtract_ledger(target: &mut ClosureLedger, value: ClosureLedger) -> Result<()> {
    target.total_bytes = target
        .total_bytes
        .checked_sub(value.total_bytes)
        .ok_or_else(|| invalid("glossary source graph ledger underflow"))?;
    target.metadata_bytes = target
        .metadata_bytes
        .checked_sub(value.metadata_bytes)
        .ok_or_else(|| invalid("glossary source metadata ledger underflow"))?;
    target.relationship_count = target
        .relationship_count
        .checked_sub(value.relationship_count)
        .ok_or_else(|| invalid("glossary source relationship ledger underflow"))?;
    target.part_count = target
        .part_count
        .checked_sub(value.part_count)
        .ok_or_else(|| invalid("glossary source part ledger underflow"))?;
    Ok(())
}

fn check_ledger(ledger: &ClosureLedger, limits: &super::Limits) -> Result<()> {
    if ledger.total_bytes > limits.graph_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphBytes,
            actual: ledger.total_bytes,
            maximum: limits.graph_bytes(),
        });
    }
    if ledger.metadata_bytes > limits.metadata_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphMetadataBytes,
            actual: ledger.metadata_bytes,
            maximum: limits.metadata_bytes(),
        });
    }
    if ledger.relationship_count > limits.values() {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::GraphWideRelationship,
            actual: ledger.relationship_count,
            maximum: limits.values(),
        });
    }
    let part_limit = limits.parts().min(MAX_VALUES);
    if ledger.part_count > part_limit {
        return Err(crate::Error::GlossaryLimit {
            kind: LimitKind::AuxiliaryPart,
            actual: ledger.part_count,
            maximum: part_limit,
        });
    }
    Ok(())
}
