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
    /// Returns an error when the workbook XML is invalid or the source cannot
    /// be read.
    pub fn edit_defined_names(&self) -> Result<Edit> {
        let package = self.edit_package()?;
        let shared_names = self.shared_defined_names();
        let before = Snapshot::load_source_backed_with_names(package, Some(shared_names))?;
        Ok(Edit::new(self.clone(), before))
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
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    names: Vec<DefinedName>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot) -> Self {
        Self {
            workbook,
            names: before.defined_names().to_vec(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged defined names in workbook order.
    #[must_use]
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.names
    }

    /// Replace the complete staged inert defined-name catalog.
    pub fn replace(&mut self, names: Vec<DefinedName>) -> Result<bool> {
        if self.names == names {
            return Ok(false);
        }
        let candidate = raw::catalog_edit::replace_defined_names(self.before.source_xml(), &names)?;
        if raw::parse_catalog(&candidate, &self.before.limits)?.defined_names != names {
            return Err(invalid("defined-name authoring verification failed"));
        }
        self.names = names;
        Ok(true)
    }

    /// Remove every direct workbook defined name.
    pub fn clear(&mut self) -> Result<bool> {
        self.replace(Vec::new())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.names = self.before.defined_names().to_vec();
    }

    fn changed(&self) -> bool {
        self.before.defined_names() != self.names
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.changed() {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }
        let output =
            raw::catalog_edit::replace_defined_names(self.before.source_xml(), &self.names)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        if snapshot.defined_names() != self.names {
            return Err(invalid(
                "defined-name publication changed the staged semantic state",
            ));
        }
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }
}
