//! Guarded defined-name publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{ReadLimits, SourceBackedPackage};

use super::{Commit, Patch, Snapshot};
use crate::error::{Result, invalid};
use crate::raw::{self, DefinedName};

/// An owning source-backed editor for the workbook defined-name catalog.
///
/// This editor is intentionally not cloneable. Publication consumes its
/// deferred OPC source and may replace only the existing workbook XML payload;
/// worksheets, calculations, relationships, and package topology are outside
/// this capability.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
    snapshot: Snapshot,
}

/// An isolated defined-name edit over one exact workbook XML source.
pub struct SourceEdit {
    before: Snapshot,
    names: Vec<DefinedName>,
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
        let package = SourceBackedPackage::from_read_at_with_limits(source, read_limits)?;
        let snapshot = Snapshot::load_source_backed(&package)?;
        Ok(Self { package, snapshot })
    }

    /// Exact source-bound catalog captured at open.
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
        let current = Snapshot::load_source_backed(&self.package)?;
        if !current.same_source(commit.patch().before()) {
            return Err(crate::Error::PatchConflict {
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
        if raw::parse_catalog(&candidate)?.defined_names != names {
            return Err(invalid("defined-name authoring verification failed"));
        }
        self.names = names;
        Ok(true)
    }

    /// Remove every direct workbook defined name.
    pub fn clear(&mut self) -> Result<bool> {
        self.replace(Vec::new())
    }

    /// Whether the exact authored semantic state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.defined_names() != self.names
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let output =
            raw::catalog_edit::replace_defined_names(self.before.source_xml(), &self.names)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        if snapshot.defined_names() != self.names {
            return Err(invalid(
                "defined-name publication changed the staged semantic state",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}
