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
//! )?)?;
//! let mut edit = package.edit_glossary();
//! edit.put(catalog, glossary::Conformance::Transitional)?;
//! let commit = edit.commit()?;
//! let (catalog, _) = commit.package().glossary()?.unwrap();
//! assert_eq!(catalog.entries()[0].name(), Some("Signature"));
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use super::transaction::Transaction as Draft;
use super::{Catalog, Conformance, Entry, Name, raw};
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
            self.draft = Some(Draft::new(self.staged.working().opc_package())?);
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
        content_type: impl Into<String>,
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
        apply_patch(self, patch, "apply_glossary")
    }
}
