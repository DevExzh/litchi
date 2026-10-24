//! Immutable worksheet slicer edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_slicers`] starts an [`Edit`] that retains the
//! package and limits without loading the slicer graph; every operation
//! works on a private candidate graph, so an error cannot stage a partial
//! graph. [`Edit::commit`] validates the
//! complete owner graph and returns a [`Commit`] holding the new package,
//! the exact part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_slicers`] replays a patch onto an equal package.
//! The source package never changes.
//!
//! An operation that would change a signed package is refused with
//! [`Error::Signed`] and stages nothing (ADR 0006); an operation that changes
//! nothing (an absent name, a reorder into the stored bytes) keeps the
//! signature byte for byte. Remove the signature explicitly first
//! ([`crate::Package::unsign`]) to edit a signed workbook.

use super::model::{Cache, Definition, Part, Slicer};
use super::raw;
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};
use litchi_opc::{OpcPackage, PackURI, SourceLimits};

family!("slicer");

pub use crate::opc::LimitKind;

/// Configuration for one slicer edit.
///
/// The package's retained source policy is always authoritative.  A caller
/// may supply a tighter policy for this edit; the two policies are
/// intersected before any raw value is admitted.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<SourceLimits>,
}

impl EditOptions {
    /// Tighten this edit's source policy.
    #[must_use]
    pub const fn limits(mut self, limits: SourceLimits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// Return the optional per-edit policy.
    #[must_use]
    pub const fn limits_value(self) -> Option<SourceLimits> {
        self.limits
    }
}

/// A staged slicer edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    limits: SourceLimits,
    /// Whether the current staged graph has passed `graph_with_limits`.
    /// Validation is deferred until a graph operation or commit requires it.
    graph_validated: bool,
}

impl Package {
    /// Start a slicer edit without loading the existing slicer graph.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Existing graph errors are reported by the first operation or
    /// by [`Edit::commit`].
    pub fn edit_slicers(&self) -> Result<Edit> {
        self.edit_slicers_with_options(EditOptions::default())
    }

    /// Start a slicer edit with an optional tighter source policy.
    pub fn edit_slicers_with_options(&self, options: EditOptions) -> Result<Edit> {
        let staged = Staged::new(self, "edit_slicers")?;
        let limits = staged
            .opc()
            .read_limits()
            .min(options.limits.unwrap_or_else(|| staged.opc().read_limits()));
        Ok(Edit {
            staged,
            limits,
            graph_validated: false,
        })
    }

    /// Apply a slicer patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_slicers(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_slicers")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged OPC graph, for diagnostics only.
    #[must_use]
    pub fn opc(&self) -> &OpcPackage {
        self.staged.opc()
    }

    /// Find a cache by its case-insensitive semantic name.
    pub fn cache(&self, name: &str) -> Result<Option<Cache>> {
        crate::slicer_cache::crud::find_slicer_cache(self.staged.opc(), name, &self.limits)
    }

    /// Find a worksheet slicer by its case-insensitive semantic name.
    pub fn view(&self, worksheet: &PackURI, name: &str) -> Result<Option<Slicer>> {
        crate::slicer_cache::crud::find_slicer(self.staged.opc(), worksheet, name, &self.limits)
    }

    pub fn add_cache(&mut self, definition: Definition) -> Result<Cache> {
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_slicer_cache(candidate, definition, &limits)
        })
    }

    /// Insert a workbook-owned Slicer Cache with its caller-selected native
    /// relationship ID and part name.
    ///
    /// This is the raw identity escape for the slicer edit family. The
    /// complete candidate graph is checked against `limits`, and publication
    /// still happens only through [`Self::commit`].
    pub fn insert_cache(&mut self, value: raw::Cache) -> Result<()> {
        let limits = self.limits;
        crate::raw::check_new_part_count(self.staged.opc(), 1, &limits)?;
        let workbook = self.staged.opc().main_document_part()?.partname().clone();
        let uri = raw::validate_cache_identity(&value, workbook.base_uri(), &limits)?;
        crate::raw::check_new_relationship_identity(
            self.staged.opc(),
            &workbook,
            &value.relationship_id,
            &uri,
        )?;
        // Admit the existing package before serializing the caller's value.
        // In particular, an already-over-limit source must report its
        // aggregate refusal rather than spending the output budget on a new
        // cache first.  Keep the operation-local proof and reuse it on the
        // unchanged candidate so this path pays one package census.
        crate::slicer::validate(&value.definition)?;
        let additions = [(
            value.relationship_id.clone(),
            crate::slicer_cache::SLICER_CACHE_RELATIONSHIP_TYPE.to_owned(),
            uri.relative_ref(workbook.base_uri()),
        )];
        let budget =
            crate::raw::prepare_package_budget(self.staged.opc(), &workbook, &additions, &limits)?;
        let prepared = raw::prepare_cache(value, workbook.base_uri(), &limits)?;
        self.apply_raw(&limits, move |candidate| {
            crate::slicer_cache::package::store_slicer_cache_prepared_in_place_with_budget(
                candidate, &prepared, &limits, &budget,
            )
        })
    }

    pub fn update_cache<F>(&mut self, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Definition),
    {
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::update_slicer_cache(candidate, name, update, &limits)
        })
    }

    pub fn replace_cache(&mut self, name: &str, replacement: Definition) -> Result<bool> {
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::replace_slicer_cache(candidate, name, replacement, &limits)
        })
    }

    pub fn reorder_caches(&mut self, names: &[String]) -> Result<Vec<Cache>> {
        let names = names.to_vec();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::reorder_slicer_caches(candidate, &names, &limits)
        })
    }

    pub fn remove_cache(&mut self, name: &str) -> Result<bool> {
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::remove_slicer_cache(candidate, name, &limits)
        })
    }

    pub fn add_view(&mut self, worksheet: &PackURI, view: Slicer) -> Result<Part> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_slicer(candidate, &worksheet, view, &limits)
        })
    }

    /// Insert a worksheet Slicers part with its caller-selected native
    /// relationship ID and part name.
    pub fn insert_part(&mut self, worksheet: &PackURI, value: raw::Part) -> Result<()> {
        let limits = self.limits;
        crate::raw::check_new_part_count(self.staged.opc(), 1, &limits)?;
        let worksheet = worksheet.clone();
        let uri = raw::validate_part_identity(&value, worksheet.base_uri(), &limits)?;
        crate::raw::check_new_relationship_identity(
            self.staged.opc(),
            &worksheet,
            &value.relationship_id,
            &uri,
        )?;
        let additions = [(
            value.relationship_id.clone(),
            crate::slicer_cache::views::SLICERS_RELATIONSHIP_TYPE.to_owned(),
            uri.relative_ref(worksheet.base_uri()),
        )];
        // Admit the existing package before serializing the caller's value.
        let budget =
            crate::raw::prepare_package_budget(self.staged.opc(), &worksheet, &additions, &limits)?;
        let prepared = raw::prepare_part(value, worksheet.base_uri(), &limits)?;
        self.apply_raw(&limits, move |candidate| {
            crate::slicer_cache::views::store_slicer_part_prepared_limited_with_budget(
                candidate, &worksheet, &prepared, &limits, &budget,
            )
        })
    }

    pub fn update_view<F>(&mut self, worksheet: &PackURI, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Slicer),
    {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::update_slicer(candidate, &worksheet, name, update, &limits)
        })
    }

    pub fn replace_view(
        &mut self,
        worksheet: &PackURI,
        name: &str,
        replacement: Slicer,
    ) -> Result<bool> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::replace_slicer(
                candidate,
                &worksheet,
                name,
                replacement,
                &limits,
            )
        })
    }

    pub fn reorder_views(&mut self, worksheet: &PackURI, names: &[String]) -> Result<Vec<Slicer>> {
        let worksheet = worksheet.clone();
        let names = names.to_vec();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::reorder_slicers(candidate, &worksheet, &names, &limits)
        })
    }

    pub fn remove_view(&mut self, worksheet: &PackURI, name: &str) -> Result<bool> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::remove_slicer(candidate, &worksheet, name, &limits)
        })
    }

    /// Slicer filtering and rendering are intentionally outside this inert
    /// package owner.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::Unsupported`].
    pub fn apply_filter(&mut self) -> Result<()> {
        crate::slicer::unsupported_ui()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.graph_validated = false;
    }

    /// Validate the staged slicer graph and publish it to a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors, [`Error::Signed`] for a changed
    /// signed package and the workbook validation errors of the staged
    /// package.
    pub fn commit(self) -> Result<Commit> {
        if !self.graph_validated {
            super::validation::graph_with_limits(self.staged.opc(), &self.limits)?;
        }
        Ok(Commit::new(self.staged.commit()?))
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let limits = self.limits;
        self.graph_validated = false;
        let value = self.staged.update(|candidate| {
            let value = operation(candidate)?;
            super::validation::graph_with_limits(candidate, &limits)?;
            Ok(value)
        })?;
        self.graph_validated = true;
        Ok(value)
    }

    fn apply_raw<T>(
        &mut self,
        limits: &SourceLimits,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.graph_validated = false;
        let value = self.staged.update(|candidate| {
            let value = operation(candidate)?;
            super::validation::graph_with_limits(candidate, limits)?;
            Ok(value)
        })?;
        self.graph_validated = true;
        Ok(value)
    }
}
