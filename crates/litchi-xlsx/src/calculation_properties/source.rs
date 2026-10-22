//! Guarded calculation-metadata publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_calculation_metadata`] stages the workbook part without
//! materializing any other part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the workbook XML
//! is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_calculation_metadata`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::snapshot::{Snapshot, same_properties};
use super::{Features, Properties, inspect, rewrite};
use crate::SourceBackedWorkbook;
use crate::error::{Result, invalid};
use std::borrow::Cow;

crate::workbook::source::deferred::source_family!("calculation-metadata", super::Patch, own_limits);

pub use super::LimitKind;

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
        let current = Snapshot::load_source_backed_with_limits(package, &self.before().limits())?;
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
}

impl SourceBackedWorkbook {
    /// Begin an isolated calculation-metadata edit of the workbook part without
    /// materializing any other part body.
    ///
    /// The calculation-properties budgets are those of the
    /// [`crate::SemanticLimits`] the workbook was opened with.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook XML is invalid or exceeds its
    /// budgets, or the source cannot be read.
    pub fn edit_calculation_metadata(&self) -> Result<Edit> {
        let limits = self.semantic_limits().calculation_properties_value();
        let before = Snapshot::load_source_backed_with_limits(self.edit_package()?, &limits)?;
        Ok(Edit::new(self.clone(), before))
    }

    /// Replay a calculation-metadata patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the workbook part does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_calculation_metadata(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated calculation-metadata edit over the exact workbook part of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    properties: Option<Properties>,
    features: Option<Features>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot) -> Self {
        Self {
            workbook,
            properties: before.properties().cloned(),
            features: before.features().cloned(),
            before,
        }
    }

    /// Exact immutable source captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged exact-authored `calcPr` state.
    #[must_use]
    pub fn properties(&self) -> Option<&Properties> {
        self.properties.as_ref()
    }

    /// Currently staged ordered calculation features.
    #[must_use]
    pub fn features(&self) -> Option<&Features> {
        self.features.as_ref()
    }

    /// Replace the staged `calcPr` state.
    pub fn set_properties(&mut self, properties: Properties) -> bool {
        if same_properties(self.properties(), Some(&properties)) {
            return false;
        }
        self.properties = Some(properties);
        true
    }

    /// Clone-edit `calcPr`, creating an empty authored value if absent.
    pub fn edit_properties(
        &mut self,
        edit: impl FnOnce(&mut Properties) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self.properties.clone().unwrap_or_default();
        edit(&mut draft)?;
        Ok(self.set_properties(draft))
    }

    /// Remove `calcPr` from the staged workbook.
    pub fn remove_properties(&mut self) -> bool {
        self.properties.take().is_some()
    }

    /// Replace the staged calculation-feature collection.
    pub fn set_features(&mut self, features: Features) -> bool {
        if self.features.as_ref() == Some(&features) {
            return false;
        }
        self.features = Some(features);
        true
    }

    /// Clone-edit the existing calculation-feature collection.
    pub fn edit_features(
        &mut self,
        edit: impl FnOnce(&mut Features) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .features
            .clone()
            .ok_or_else(|| invalid("workbook has no calculation features to edit"))?;
        edit(&mut draft)?;
        Ok(self.set_features(draft))
    }

    /// Remove calculation features from the staged workbook.
    pub fn remove_features(&mut self) -> bool {
        self.features.take().is_some()
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.properties = self.before.properties().cloned();
        self.features = self.before.features().cloned();
    }

    fn changed(&self) -> bool {
        !same_properties(self.before.properties(), self.properties())
            || self.before.features() != self.features()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.changed() {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }

        let inspection = inspect(self.before.source_xml(), &self.before.limits())?;
        let output = match rewrite(
            &inspection,
            self.properties.as_ref(),
            self.features.as_ref(),
            &self.before.limits(),
        )? {
            Cow::Owned(output) => output,
            Cow::Borrowed(_) => {
                return Err(invalid(
                    "changed calculation metadata rewrite produced no output",
                ));
            },
        };
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        if !snapshot.same_semantics(self.properties.as_ref(), self.features.as_ref()) {
            return Err(invalid(
                "calculation metadata publication changed the staged semantics",
            ));
        }
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }
}
