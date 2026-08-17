//! Guarded conditional-formatting publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};

use super::{Commit, Formatting, Patch, Snapshot};
use crate::Selector;
use crate::error::{Error, Result, invalid};

/// Owning source-backed worksheet conditional-formatting editor.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// Isolated complete-collection edit over one exact worksheet source.
pub struct SourceEdit {
    before: Snapshot,
    staged: Vec<Formatting>,
}

impl SourceBackedEditor {
    /// Open a deferred OPC source with the standard bounded read policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open a deferred OPC source with an explicit bounded read policy.
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

    /// Capture the exact conditional-formatting state for one worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        self.package.check_execution()?;
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin an isolated complete-collection edit for one worksheet.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        self.package.check_execution()?;
        Ok(SourceEdit::new(self.snapshot(selector)?))
    }

    #[must_use]
    /// Return content-free payload-cache diagnostics for the deferred package.
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Publish by overlaying only the selected worksheet Part.
    ///
    /// Every unselected physical member is raw-copied. Publication checks the
    /// complete retained workbook, worksheet-relationship, and styles closure
    /// before writing; an exact no-op reproduces the source archive bytes.
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
            staged: before.collections().to_vec(),
            before,
        }
    }

    #[must_use]
    /// Exact source state captured when this edit began.
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    /// Complete ordered collection currently staged for publication.
    pub fn collections(&self) -> &[Formatting] {
        &self.staged
    }

    /// Replace the complete ordered owner collection atomically.
    pub fn set_collections(&mut self, value: Vec<Formatting>) -> Result<bool> {
        super::package::validate_authored(&value, self.before.differential_format_count())?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Mutate a clone and publish it to the edit only after complete validation.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Vec<Formatting>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.staged.clone();
        operation(&mut candidate)?;
        self.set_collections(candidate)
    }

    /// Remove all direct core conditional-formatting owners.
    pub fn clear(&mut self) -> bool {
        if self.staged.is_empty() {
            return false;
        }
        self.staged.clear();
        true
    }

    #[must_use]
    /// Whether the staged collection differs semantically from its source.
    pub fn is_changed(&self) -> bool {
        self.before.collections() != self.staged
    }

    /// Validate, rewrite, reopen, and freeze the isolated edit.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.before.mutation_locked() {
            return Err(invalid(
                "worksheet protection forbids conditional-formatting mutation",
            ));
        }
        let dxf_count = self.before.differential_format_count();
        super::package::validate_authored(&self.staged, dxf_count)?;
        let output = super::replace_conditional_formattings(
            self.before.source_xml(),
            &self.staged,
            dxf_count,
        )?;
        let readback = super::package::parse_editable_conditional_formattings(&output, dxf_count)?;
        if readback != self.staged {
            return Err(invalid(
                "conditional-formatting publication changed the staged state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&self.before, output, readback)?;
        self.before.check_execution()?;
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
