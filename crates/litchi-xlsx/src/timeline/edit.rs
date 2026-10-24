//! Immutable worksheet timeline edits of a [`crate::Package`] (record
//! 0360).
//!
//! [`crate::Package::edit_timelines`] starts an [`Edit`] that retains the
//! package metadata and limits without loading the timeline graph; every
//! operation works on a private candidate graph, so an error cannot stage a
//! partial graph. Timeline cache references are
//! workbook-owned, so the workbook part identity is captured once when the
//! edit starts. [`Edit::commit`] returns a [`Commit`] holding the new
//! package, the exact part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_timelines`] replays a patch onto an equal
//! package. The source package never changes.
//!
//! An operation that would change a signed package is refused with
//! [`Error::Signed`] and stages nothing (ADR 0006); an operation that changes
//! nothing (an absent name, a reorder into the stored bytes) keeps the
//! signature byte for byte. Remove the signature explicitly first
//! ([`crate::Package::unsign`]) to edit a signed workbook.

use super::model::{Cache, CacheDefinition, Part, View};
use super::raw;
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};
use litchi_opc::{OpcPackage, PackURI, SourceLimits};
use std::collections::HashSet;

family!("timeline");

pub use crate::opc::LimitKind;

/// Configuration for one timeline edit.
///
/// The package's retained source policy remains authoritative. A supplied
/// policy can only tighten that policy for this edit.
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

/// A staged timeline edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    workbook: PackURI,
    limits: SourceLimits,
    /// Whether the current staged graph has passed `graph_with_limits`.
    /// Validation is deferred until a graph operation or commit requires it.
    graph_validated: bool,
}

impl Package {
    /// Start a timeline edit without loading the existing timeline graph.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Existing graph errors are reported by the first operation or
    /// by [`Edit::commit`].
    pub fn edit_timelines(&self) -> Result<Edit> {
        self.edit_timelines_with_options(EditOptions::default())
    }

    /// Start a timeline edit with an optional tighter source policy.
    pub fn edit_timelines_with_options(&self, options: EditOptions) -> Result<Edit> {
        let staged = Staged::new(self, "edit_timelines")?;
        let workbook = staged.opc().main_document_part()?.partname().clone();
        let limits = staged
            .opc()
            .read_limits()
            .min(options.limits.unwrap_or_else(|| staged.opc().read_limits()));
        Ok(Edit {
            staged,
            workbook,
            limits,
            graph_validated: false,
        })
    }

    /// Apply a timeline patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_timelines(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_timelines")?,
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
        crate::slicer_cache::crud::find_timeline_cache(self.staged.opc(), name, &self.limits)
    }

    /// Find a worksheet timeline by its case-insensitive semantic name.
    pub fn view(&self, worksheet: &PackURI, name: &str) -> Result<Option<View>> {
        crate::slicer_cache::crud::find_timeline(self.staged.opc(), worksheet, name, &self.limits)
    }

    pub fn add_cache(&mut self, definition: CacheDefinition) -> Result<Cache> {
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_timeline_cache(candidate, definition, &limits)
        })
    }

    /// Insert a complete workbook View Cache set with caller-selected native
    /// relationship IDs and part names.
    ///
    /// This is the raw identity escape for the timeline edit family. The
    /// complete candidate graph is checked against `limits`, and publication
    /// still happens only through [`Self::commit`].
    pub fn insert_caches(&mut self, values: &[raw::Cache]) -> Result<()> {
        let limits = self.limits;
        crate::raw::check_limit(
            litchi_opc::SourceLimitKind::RelationshipsPerPart,
            values.len() as u64,
            limits.max_relationships_per_part() as u64,
        )?;
        crate::raw::check_new_part_count(self.staged.opc(), values.len(), &limits)?;
        let mut ids = HashSet::with_capacity(values.len());
        let mut targets = HashSet::<String>::with_capacity(values.len());
        let mut additions = Vec::with_capacity(values.len());
        for value in values {
            let uri = raw::validate_cache_identity(value, self.workbook.base_uri(), &limits)?;
            if !ids.insert(value.relationship_id.as_str()) {
                return Err(crate::error::invalid(
                    "duplicate View Cache relationship ID",
                ));
            }
            if !targets.insert(uri.to_string()) {
                return Err(crate::error::invalid("duplicate View Cache part name"));
            }
            crate::raw::check_new_relationship_identity(
                self.staged.opc(),
                &self.workbook,
                &value.relationship_id,
                &uri,
            )?;
            additions.push((
                value.relationship_id.clone(),
                crate::timeline::CACHE_RELATIONSHIP_TYPE.to_owned(),
                uri.relative_ref(self.workbook.base_uri()),
            ));
        }
        let budget = crate::raw::prepare_package_budget(
            self.staged.opc(),
            &self.workbook,
            &additions,
            &limits,
        )?;
        let prepared = values
            .iter()
            .cloned()
            .map(|value| raw::prepare_cache(value, self.workbook.base_uri(), &limits))
            .collect::<Result<Vec<_>>>()?;
        let workbook = self.workbook.clone();
        self.apply_raw(&limits, move |candidate| {
            crate::timelines::store_timeline_caches_prepared_limited_with_budget(
                candidate, &workbook, &prepared, &limits, &budget,
            )
        })
    }

    pub fn update_cache<F>(&mut self, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut CacheDefinition),
    {
        let limits = self.limits;
        self.apply(|candidate| {
            crate::slicer_cache::crud::update_timeline_cache(candidate, name, update, &limits)
        })
    }

    pub fn replace_cache(&mut self, name: &str, replacement: CacheDefinition) -> Result<bool> {
        let limits = self.limits;
        self.apply(|candidate| {
            crate::slicer_cache::crud::replace_timeline_cache(candidate, name, replacement, &limits)
        })
    }

    pub fn reorder_caches(&mut self, names: &[String]) -> Result<Vec<Cache>> {
        let names = names.to_vec();
        let limits = self.limits;
        self.apply(|candidate| {
            crate::slicer_cache::crud::reorder_timeline_caches(candidate, &names, &limits)
        })
    }

    pub fn remove_cache(&mut self, name: &str) -> Result<bool> {
        let limits = self.limits;
        self.apply(|candidate| {
            crate::slicer_cache::crud::remove_timeline_cache(candidate, name, &limits)
        })
    }

    pub fn add_view(&mut self, worksheet: &PackURI, view: View) -> Result<Part> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_timeline(candidate, &worksheet, view, &limits)
        })
    }

    /// Insert a worksheet Timelines part with its caller-selected worksheet,
    /// relationship ID and part name.
    pub fn insert_part(&mut self, value: raw::Part) -> Result<()> {
        let limits = self.limits;
        crate::raw::check_new_part_count(self.staged.opc(), 1, &limits)?;
        let (uri, worksheet) = raw::validate_part_identity(&value, &limits)?;
        crate::raw::check_new_relationship_identity(
            self.staged.opc(),
            &worksheet,
            &value.relationship_id,
            &uri,
        )?;
        let additions = [(
            value.relationship_id.clone(),
            crate::timeline::RELATIONSHIP_TYPE.to_owned(),
            uri.relative_ref(worksheet.base_uri()),
        )];
        // Admit the existing package before serializing the caller's value.
        let budget =
            crate::raw::prepare_package_budget(self.staged.opc(), &worksheet, &additions, &limits)?;
        let prepared = raw::prepare_part(value, &limits)?;
        let workbook = self.workbook.clone();
        self.apply_raw(&limits, move |candidate| {
            crate::timelines::store_worksheet_timelines_prepared_limited_with_budget(
                candidate, &workbook, &prepared, &limits, &budget,
            )
        })
    }

    pub fn update_view<F>(&mut self, worksheet: &PackURI, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut View),
    {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::update_timeline(candidate, &worksheet, name, update, &limits)
        })
    }

    pub fn replace_view(
        &mut self,
        worksheet: &PackURI,
        name: &str,
        replacement: View,
    ) -> Result<bool> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::replace_timeline(
                candidate,
                &worksheet,
                name,
                replacement,
                &limits,
            )
        })
    }

    pub fn reorder_views(&mut self, worksheet: &PackURI, names: &[String]) -> Result<Vec<View>> {
        let worksheet = worksheet.clone();
        let names = names.to_vec();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::reorder_timelines(candidate, &worksheet, &names, &limits)
        })
    }

    pub fn remove_view(&mut self, worksheet: &PackURI, name: &str) -> Result<bool> {
        let worksheet = worksheet.clone();
        let limits = self.limits;
        self.apply(move |candidate| {
            crate::slicer_cache::crud::remove_timeline(candidate, &worksheet, name, &limits)
        })
    }

    /// Timeline filtering and rendering are intentionally outside this inert
    /// package owner.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::Unsupported`].
    pub fn apply_filter(&mut self) -> Result<()> {
        crate::timeline::unsupported_ui()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.graph_validated = false;
    }

    /// Validate the staged timeline graph and publish it to a new package.
    /// An unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors, [`Error::Signed`] for a changed
    /// signed package and the workbook validation errors of the staged
    /// package.
    pub fn commit(self) -> Result<Commit> {
        if !self.graph_validated {
            super::validation::graph_with_limits(self.staged.opc(), &self.workbook, &self.limits)?;
        }
        Ok(Commit::new(self.staged.commit()?))
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let workbook = &self.workbook;
        let limits = self.limits;
        self.graph_validated = false;
        let value = self.staged.update(|candidate| {
            let value = operation(candidate)?;
            super::validation::graph_with_limits(candidate, workbook, &limits)?;
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
        let workbook = &self.workbook;
        self.graph_validated = false;
        let value = self.staged.update(|candidate| {
            let value = operation(candidate)?;
            super::validation::graph_with_limits(candidate, workbook, limits)?;
            Ok(value)
        })?;
        self.graph_validated = true;
        Ok(value)
    }
}
