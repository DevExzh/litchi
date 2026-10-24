//! The source-backed cell-value and row-visibility families of record 0360
//! in the open/edit/publish shape these integration tests drive.
//!
//! The editors here are thin test drivers: they open a
//! [`SourceBackedWorkbook`], start family edits on it, and "publish" a commit
//! by replaying its patch onto their own workbook and writing the result, so
//! a commit made over another open of the same bytes, a stale source or a
//! foreign source exercises the families' replay checks.
#![allow(
    dead_code,
    unreachable_pub,
    reason = "each test target uses a subset of the drivers"
)]

use std::io::Write;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{OpcPackage, SourceBackedPackage};
use litchi_xlsx::cell_values::{self, MultiSnapshot, SheetChange, Snapshot};
use litchi_xlsx::{OpenOptions, Result, Selector, SourceBackedWorkbook, row_visibility};

/// Drive the cell-value families over one source-backed workbook.
pub struct SourceBackedEditor {
    workbook: SourceBackedWorkbook,
}

impl SourceBackedEditor {
    fn with(source: Arc<dyn ReadAt>, options: OpenOptions) -> Result<Self> {
        Ok(Self {
            workbook: SourceBackedWorkbook::from_read_at_with_options(source, options)?,
        })
    }

    pub fn workbook(&self) -> &SourceBackedWorkbook {
        &self.workbook
    }

    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::with(source, OpenOptions::new())
    }

    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: OpenOptions,
    ) -> Result<Self> {
        Self::with(source, options)
    }

    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Ok(Self {
            workbook: SourceBackedWorkbook::open(path)?,
        })
    }

    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self> {
        Ok(Self {
            workbook: SourceBackedWorkbook::open_with_options(path, options)?,
        })
    }

    #[cfg(any(unix, windows))]
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        Ok(Self {
            workbook: SourceBackedWorkbook::from_source_backed_package(package)?,
        })
    }

    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        Ok(self.workbook.edit_sheet_values(selector)?.before()?.clone())
    }

    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<cell_values::sheet::Edit> {
        self.workbook.edit_sheet_values(selector)
    }

    pub fn edit_sheets<'a, I>(&self, selectors: I) -> Result<cell_values::Edit>
    where
        I: IntoIterator<Item = Selector<'a>>,
    {
        self.workbook.edit_cell_values(selectors)
    }

    pub fn edit_many<'a, I>(&self, edits: I) -> Result<cell_values::Edit>
    where
        I: IntoIterator<Item = SheetChange<'a>>,
    {
        self.workbook.edit_cell_values_batch(edits)
    }

    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.workbook.cache_diagnostics()
    }

    /// Replay `commit`'s patch onto this editor's workbook and write it.
    pub fn publish_commit_to_stream<W: Write>(
        &self,
        writer: W,
        commit: &cell_values::sheet::Commit,
    ) -> Result<Snapshot> {
        let replayed = self.workbook.apply_sheet_values(commit.patch())?;
        replayed.package().write_to(writer)?;
        Ok(commit.patch().after().clone())
    }

    /// Replay `commit`'s patch onto this editor's workbook and write it.
    pub fn publish_multi_commit_to_stream<W: Write>(
        &self,
        writer: W,
        commit: &cell_values::Commit,
    ) -> Result<MultiSnapshot> {
        let replayed = self.workbook.apply_cell_values(commit.patch())?;
        replayed.package().write_to(writer)?;
        Ok(commit.patch().after().clone())
    }
}

/// Drive the row-visibility family over one source-backed workbook.
pub struct RowVisibilityEditor {
    inner: SourceBackedEditor,
}

impl RowVisibilityEditor {
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Ok(Self {
            inner: SourceBackedEditor::from_read_at(source)?,
        })
    }

    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: OpenOptions,
    ) -> Result<Self> {
        Ok(Self {
            inner: SourceBackedEditor::from_read_at_with_options(source, options)?,
        })
    }

    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        Ok(Self {
            inner: SourceBackedEditor::from_source_backed_package(package)?,
        })
    }

    pub fn snapshot<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> Result<row_visibility::Snapshot> {
        Ok(self.edit(selector)?.before().clone())
    }

    pub fn edit<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> Result<row_visibility::source::Edit> {
        self.inner.workbook.edit_row_visibility(selector)
    }

    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.inner.cache_diagnostics()
    }

    /// Replay `commit`'s patch onto this editor's workbook and write it.
    pub fn publish_commit_to_stream<W: Write>(
        &self,
        writer: W,
        commit: &row_visibility::source::Commit,
    ) -> Result<row_visibility::Snapshot> {
        let replayed = self.inner.workbook.apply_row_visibility(commit.patch())?;
        replayed.package().write_to(writer)?;
        Ok(commit.patch().after().clone())
    }
}

/// A source-backed patch that replays onto a source-backed workbook.
pub trait Replay {
    fn replay(&self, workbook: &SourceBackedWorkbook) -> Result<SourceBackedWorkbook>;
}

impl Replay for cell_values::sheet::Patch {
    fn replay(&self, workbook: &SourceBackedWorkbook) -> Result<SourceBackedWorkbook> {
        Ok(workbook.apply_sheet_values(self)?.into_package())
    }
}

impl Replay for cell_values::Patch {
    fn replay(&self, workbook: &SourceBackedWorkbook) -> Result<SourceBackedWorkbook> {
        Ok(workbook.apply_cell_values(self)?.into_package())
    }
}

impl Replay for row_visibility::source::Patch {
    fn replay(&self, workbook: &SourceBackedWorkbook) -> Result<SourceBackedWorkbook> {
        Ok(workbook.apply_row_visibility(self)?.into_package())
    }
}

/// Replay a source-backed patch onto `package` through a source-backed
/// workbook over its bytes.
pub fn replay_patch(package: &mut OpcPackage, patch: &impl Replay) -> Result<()> {
    let bytes = litchi_opc::PackageWriter::to_bytes(package)?;
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let next = patch.replay(&workbook)?.to_bytes()?;
    *package = OpcPackage::from_bytes(&next, litchi_opc::SourceLimits::default())?;
    Ok(())
}

/// Replay a source-backed patch onto a source-backed workbook over `bytes`.
pub fn replay_bytes(bytes: &[u8], patch: &impl Replay) -> Result<Vec<u8>> {
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(litchi_core::OwnedSource::new(
        bytes.to_vec(),
    )))?;
    patch.replay(&workbook)?.to_bytes()
}
