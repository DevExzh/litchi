//! Guarded calculation-metadata publication over an immutable positional source.

use std::borrow::Cow;
use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{ReadLimits, SourceBackedPackage};

use super::patch::{Commit, Patch};
use super::snapshot::{Snapshot, same_properties};
use super::{Features, Limits, Properties, inspect, rewrite};
use crate::error::{Error, Result, invalid};

/// An owning source-backed editor for workbook calculation metadata.
///
/// This editor is intentionally not cloneable. Publication consumes its
/// deferred OPC source and may replace only the existing workbook XML payload;
/// worksheets, formulas, the calculation chain, relationships, and package
/// topology are outside this capability.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
    snapshot: Snapshot,
}

/// An isolated calculation-metadata edit over one exact workbook XML source.
pub struct SourceEdit {
    before: Snapshot,
    properties: Option<Properties>,
    features: Option<Features>,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC and calculation-metadata policies.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default(), Limits::default())
    }

    /// Open with explicit OPC and calculation-metadata resource policies.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        calculation_limits: Limits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits(source, read_limits)?,
            calculation_limits,
        )
    }

    fn from_source_backed_package(
        package: SourceBackedPackage,
        calculation_limits: Limits,
    ) -> Result<Self> {
        let snapshot = Snapshot::load_source_backed_with_limits(&package, &calculation_limits)?;
        Ok(Self { package, snapshot })
    }

    /// Exact source-bound calculation metadata captured at open.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Begin an isolated edit without materializing any other ordinary Part.
    #[must_use]
    pub fn edit(&self) -> SourceEdit {
        SourceEdit::new(self.snapshot.clone())
    }

    /// Return content-free payload-cache activity for the deferred package.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Publish an exact-source-checked commit to a sequential sink.
    ///
    /// The workbook XML is the only replaceable Part. Every other physical ZIP
    /// member is raw-copied. Exact no-ops reproduce the complete source bytes;
    /// changed signed sources and unsupported physical layouts are refused by
    /// the OPC owner before output begins.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        let current =
            Snapshot::load_source_backed_with_limits(&self.package, &self.snapshot.limits())?;
        if !current.same_source(commit.patch().before()) {
            return Err(Error::PatchConflict {
                part: current.workbook_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current
        } else {
            commit.patch().after().clone()
        };
        self.package.write_part_overlay_to_stream(
            writer,
            target.workbook_part_name(),
            target.source_xml().to_vec(),
        )?;
        Ok(target)
    }
}

impl SourceEdit {
    fn new(before: Snapshot) -> Self {
        Self {
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

    /// Whether the authored metadata differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_properties(self.before.properties(), self.properties())
            || self.before.features() != self.features()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
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
        if !snapshot.same_semantics(self.properties.as_ref(), self.features.as_ref()) {
            return Err(invalid(
                "calculation metadata publication changed the staged semantics",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
