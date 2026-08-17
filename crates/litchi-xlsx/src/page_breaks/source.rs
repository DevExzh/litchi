//! Guarded page-break publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};

use super::{Collection, Commit, PageBreaks, Patch, Snapshot};
use crate::Selector;
use crate::error::{Result, invalid};

/// An owning source-backed page-break editor.
///
/// The editor is intentionally not cloneable. Publication consumes its
/// deferred OPC source and may replace only one existing worksheet payload;
/// the workbook catalog, relationships, other worksheets, drawings, and
/// package topology are immutable under this capability.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// An isolated page-break edit over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: PageBreaks,
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
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source,
            read_limits,
        )?)
    }

    /// Open with an explicit finite deferred-payload cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open with explicit read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                read_limits,
                cache_limits,
            )?,
        )
    }

    /// Open with an explicit caller-owned execution context.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source,
            read_limits,
            context,
        )?)
    }

    /// Open with explicit read and execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, read_limits, context)
    }

    /// Open with explicit read, cache, and caller-owned execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                read_limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build an editor from an already opened deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        Ok(Self { package })
    }

    /// Capture exact source-bound page breaks for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        self.package.check_execution()?;
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin an isolated edit without materializing any unselected Part body.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        self.package.check_execution()?;
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
        self.package.check_execution()?;
        let current =
            Snapshot::load_source_backed(&self.package, commit.patch().before().sheet_position())?;
        if !current.same_source(commit.patch().before()) {
            return Err(crate::Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current
        } else {
            commit.patch().after().clone()
        };
        if commit.patch().is_empty() {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        } else {
            self.package.write_part_overlay_shared_to_stream(
                writer,
                target.worksheet_part_name(),
                target.source_arc()?,
            )?;
        }
        Ok(target)
    }
}

impl SourceEdit {
    fn new(before: Snapshot) -> Self {
        Self {
            staged: before.page_breaks().clone(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-break state.
    #[must_use]
    pub const fn page_breaks(&self) -> &PageBreaks {
        &self.staged
    }

    /// Replace the complete staged page-break state.
    pub fn set(&mut self, value: PageBreaks) -> Result<bool> {
        value.validate()?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Clone-edit the complete staged value; a failed closure changes nothing.
    pub fn edit(&mut self, edit: impl FnOnce(&mut PageBreaks) -> Result<()>) -> Result<bool> {
        let mut draft = self.staged.clone();
        edit(&mut draft)?;
        self.set(draft)
    }

    /// Replace horizontal row breaks.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_horizontal(collection))
    }

    /// Replace vertical column breaks.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_vertical(collection))
    }

    /// Remove horizontal row breaks.
    pub fn remove_horizontal(&mut self) -> bool {
        self.staged.remove_horizontal()
    }

    /// Remove vertical column breaks.
    pub fn remove_vertical(&mut self) -> bool {
        self.staged.remove_vertical()
    }

    /// Remove both authored collections.
    pub fn clear(&mut self) -> bool {
        self.staged.clear()
    }

    /// Whether the exact authored semantic state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.page_breaks() != &self.staged
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let output = super::replace(self.before.source_xml(), &self.staged)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        if snapshot.page_breaks() != &self.staged {
            return Err(invalid(
                "page-break publication changed the staged semantic state",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
