//! Guarded auto-filter publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};

use super::package::validate_definition;
use super::{Commit, Definition, Patch, Snapshot};
use crate::Selector;
use crate::error::{Error, Result, invalid};

/// An owning source-backed worksheet auto-filter editor.
///
/// The editor is intentionally not cloneable. Publication consumes its
/// deferred OPC source and may replace only one existing worksheet payload;
/// package topology and every relationship remain immutable.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// An isolated auto-filter edit over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Option<Definition>,
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

    /// Open with an explicit managed execution context.
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

    /// Open with explicit read and managed execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, read_limits, context)
    }

    /// Open with explicit read, cache, and managed execution policies.
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

    /// Capture exact source-bound auto-filter state for one worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        self.package.check_execution()?;
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin an isolated edit without materializing unselected Part bodies.
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
        if !commit
            .patch()
            .before()
            .matches_source_backed(&self.package)?
        {
            return Err(Error::PatchConflict {
                part: commit.patch().before().worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            commit.patch().before().clone()
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
            staged: before.auto_filter().cloned(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Complete currently staged filter and sort state.
    #[must_use]
    pub const fn auto_filter(&self) -> Option<&Definition> {
        self.staged.as_ref()
    }

    /// Create or replace the complete staged filter and sort state.
    pub fn set(&mut self, value: Definition) -> Result<bool> {
        validate_definition(&value)?;
        if self.staged.as_ref() == Some(&value) {
            return Ok(false);
        }
        self.staged = Some(value);
        Ok(true)
    }

    /// Edit a clone atomically, retaining the staged state on error.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Option<Definition>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.staged.clone();
        operation(&mut candidate)?;
        if let Some(value) = candidate.as_ref() {
            validate_definition(value)?;
        }
        if candidate == self.staged {
            return Ok(false);
        }
        self.staged = candidate;
        Ok(true)
    }

    /// Remove the direct worksheet auto-filter and its sort state.
    pub fn clear(&mut self) -> bool {
        self.staged.take().is_some()
    }

    /// Whether the authored semantic state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.auto_filter() != self.staged.as_ref()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.before.mutation_locked(self.staged.as_ref()) {
            return Err(invalid(
                "worksheet protection forbids the staged auto-filter or sort mutation",
            ));
        }
        if let Some(value) = self.staged.as_ref() {
            validate_definition(value)?;
        }
        let output = super::replace_auto_filter(self.before.source_xml(), self.staged.as_ref())?;
        let readback = super::parse_auto_filter(&output)?;
        if readback != self.staged {
            return Err(invalid(
                "auto-filter publication changed the staged semantic state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&self.before, output, readback)?;
        self.before.check_execution()?;
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
