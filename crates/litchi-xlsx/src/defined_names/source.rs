//! Guarded defined-name publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_defined_names`] stages the workbook part without
//! materializing any other part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the workbook XML
//! is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_defined_names`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::Snapshot;
use crate::SourceBackedWorkbook;
use crate::error::{Result, invalid};
use crate::raw::{self, DefinedName};
use crate::workbook::source::LazySource;

crate::workbook::source::deferred::source_family!("defined-name", super::Patch);

impl Patch {
    /// The exact workbook state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact workbook state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }
}

impl crate::workbook::source::deferred::Change for super::Patch {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        super::Patch::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        let current = Snapshot::load_source_backed(package)?;
        if !current.same_source(self.before()) {
            return Err(Error::PatchConflict {
                part: current.workbook_part_name().to_string(),
            });
        }
        Ok(())
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        package.check_execution()?;
        crate::workbook::source::deferred::Change::check(self, package)?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().workbook_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().workbook_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }

    fn catalog(
        &self,
        package: &SourceBackedPackage,
    ) -> Option<crate::workbook::source::deferred::Catalog> {
        (package.read_limits() == self.after().limits).then(|| self.after().deferred_catalog())
    }

    fn catalog_xml(&self) -> Option<&[u8]> {
        Some(self.after().source_xml())
    }
}

impl SourceBackedWorkbook {
    /// Begin an isolated defined-name edit of the workbook part without
    /// materializing any other part body.
    ///
    /// # Errors
    ///
    /// Workbook/XML resolution errors are returned by the first accessor,
    /// staging operation, or commit that needs the source.
    pub fn edit_defined_names(&self) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone()))
    }

    /// Replay a defined-name patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the workbook part does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_defined_names(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated defined-name edit over the exact workbook part of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    names: Option<Vec<DefinedName>>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook) -> Self {
        Self {
            source: LazySource::new(workbook),
            names: None,
        }
    }

    fn resolve_before(&self) -> Result<&Snapshot> {
        // A committed workbook keeps its original owner until the first
        // source-backed read materializes the deferred output. Resolve that
        // owner before taking the shared catalog proof: otherwise a pending
        // defined-name commit would pair its new workbook XML with the base
        // owner's names and reject the exact output as a catalog mismatch.
        self.source.workbook().edit_package()?;
        let shared_catalog = self.source.workbook().shared_catalog_proof();
        self.source.resolve(|package| {
            Snapshot::load_source_backed_with_names(
                package,
                Some(shared_catalog.defined_names.clone()),
            )
        })
    }

    fn effective_names(&self) -> Result<&[DefinedName]> {
        match self.names.as_deref() {
            Some(names) => Ok(names),
            None => Ok(self.resolve_before()?.defined_names()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Currently staged defined names in workbook order.
    pub fn defined_names(&self) -> Result<&[DefinedName]> {
        self.effective_names()
    }

    /// Replace the complete staged inert defined-name catalog.
    pub fn replace(&mut self, names: Vec<DefinedName>) -> Result<bool> {
        if self.effective_names()? == names.as_slice() {
            return Ok(false);
        }
        let before = self.resolve_before()?;
        let candidate = raw::catalog_edit::replace_defined_names(before.source_xml(), &names)?;
        if raw::parse_catalog(&candidate, &before.limits)?.defined_names != names {
            return Err(invalid("defined-name authoring verification failed"));
        }
        self.names = Some(names);
        Ok(true)
    }

    /// Remove every direct workbook defined name.
    pub fn clear(&mut self) -> Result<bool> {
        self.replace(Vec::new())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.names = None;
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        let names = self
            .names
            .unwrap_or_else(|| before.defined_names().to_vec());
        before.check_execution()?;
        if before.defined_names() == names.as_slice() {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        let output = raw::catalog_edit::replace_defined_names(before.source_xml(), &names)?;
        let snapshot = Snapshot::from_rewritten_source(&before, output)?;
        snapshot.check_execution()?;
        if snapshot.defined_names() != names.as_slice() {
            return Err(invalid(
                "defined-name publication changed the staged semantic state",
            ));
        }
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
