//! Guarded page-setup publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{ReadLimits, SourceBackedPackage};

use super::{Commit, Patch, Setup, Snapshot};
use crate::Selector;
use crate::error::{Error, Result, invalid};
use crate::source_provenance::SourceProvenance;

/// An owning source-backed page-setup editor.
///
/// The editor is intentionally not cloneable. Publication consumes its
/// deferred OPC source and may replace only one existing worksheet payload;
/// relationships, printer settings, other Parts, and package topology remain
/// immutable under this capability.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// An isolated page-setup edit over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Option<Setup>,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with an explicit bounded OPC policy.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
    ) -> Result<Self> {
        Ok(Self {
            package: SourceBackedPackage::from_read_at_with_limits(source, read_limits)?,
        })
    }

    /// Capture exact source-bound page setup for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin an isolated edit without materializing any unselected Part body.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        Ok(SourceEdit::new(self.snapshot(selector)?))
    }

    /// Return content-free payload-cache activity for the deferred package.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Publish an exact-source-checked commit to a sequential sink.
    ///
    /// The selected worksheet XML is the only replaceable Part. Every other
    /// physical ZIP member is raw-copied. Exact no-ops reproduce the complete
    /// source bytes; changed signed sources and unsupported physical layouts
    /// are refused by the OPC owner before output begins.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        let before = commit.patch().before();
        let current = match before.matches_source_backed(&self.package)? {
            SourceProvenance::Matched => None,
            SourceProvenance::Mismatched => {
                return Err(Error::PatchConflict {
                    part: before.worksheet_part_name().to_string(),
                });
            },
            SourceProvenance::Unavailable => Some(Snapshot::load_source_backed(
                &self.package,
                before.sheet_position(),
            )?),
        };
        if let Some(current) = &current
            && !current.same_source(before)
        {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current.unwrap_or_else(|| before.clone())
        } else {
            commit.patch().after().clone()
        };
        self.package.write_part_overlay_to_stream(
            writer,
            target.worksheet_part_name(),
            target.source_xml().to_vec(),
        )?;
        Ok(target)
    }
}

impl SourceEdit {
    fn new(before: Snapshot) -> Self {
        Self {
            staged: before.page_setup().cloned(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-setup state.
    #[must_use]
    pub const fn page_setup(&self) -> Option<&Setup> {
        self.staged.as_ref()
    }

    /// Create or replace the complete staged relationship-free settings.
    ///
    /// An existing printer-settings relationship, if present, is retained
    /// exactly by the worksheet codec and is never exposed for mutation here.
    pub fn set(&mut self, value: Setup) -> bool {
        if self.staged.as_ref() == Some(&value) {
            return false;
        }
        self.staged = Some(value);
        true
    }

    /// Remove direct page setup.
    ///
    /// Commit retains the typed refusal when printer settings are attached.
    pub fn remove(&mut self) -> bool {
        self.staged.take().is_some()
    }

    /// Whether the exact authored settings differ from their source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.page_setup() != self.staged.as_ref()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let output =
            super::replace_worksheet_page_setup(self.before.source_xml(), self.staged.as_ref())?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        if snapshot.page_setup() != self.staged.as_ref() {
            return Err(invalid(
                "page-setup publication changed the staged semantic state",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
