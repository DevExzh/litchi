//! Handle for Excel OOXML (`.xlsx`) inputs.
//!
//! [`Xlsx`] is one `Arc` over its state: the two-state owner of record 0003
//! §2.3 (`Source` for the deferred, source-backed catalog; `Eager` for the
//! owning workbook snapshot), the core-properties metadata and its repairs
//! read at open (§1.6), the lazily built worksheet-name list, and the
//! retention policy from [`OpenOptions`]. Cloning bumps one reference count.
//!
//! # Routes
//!
//! The constructors take what detection already produced, so an open does
//! one detection pass and one container index:
//!
//! | Facade route today | Detection result | Constructor |
//! |---|---|---|
//! | `Workbook::open` on `unix`/`windows` | `WorkbookSourcePathDetection::Xlsx` | `Xlsx::from_source_workbook` |
//! | `Workbook::from_bytes` | `WorkbookSourceBytesDetection::Xlsx` | `Xlsx::from_source_package` |
//! | `Workbook::from_bytes` fallback (`detect_format_smart`) | `DetectedFormat::Xlsx` | `Xlsx::from_opc_package` |
//! | `Workbook::from_prepared` | `PreparedOoxmlSource` | `Xlsx::from_prepared` |
//!
//! The facade has no reader route for workbooks; the `sheet::open_workbook*`
//! functions are path and byte routes onto the same two owners. On other
//! targets the path route reads the bytes and takes the byte route.
//!
//! # Source freshness
//!
//! `XlsxOwner::ensure_current` is the facade adapter's
//! `ensure_source_current`. Catalog reads (names, count, metadata, tolerance
//! report) check the source revision first, as the facade did. Reads that
//! touch worksheet payloads ([`Xlsx::text`], [`Xlsx::sheet`]) are bracketed by
//! `XlsxOwner::finish`, which checks before and after, so a changed source
//! wins over what the read produced. Cell reads through a selected or generic
//! worksheet are fenced by the source-backed leaf itself.
//!
//! # Generic sheet traits
//!
//! [`Xlsx`] implements [`WorkbookTrait`] directly, so the workbook view that
//! S6 adds delegates to it by a `match`, with no boxed workbook. The trait's
//! worksheet and error types stay boxed because the trait defines them so;
//! its errors are boxed [`crate::Error`] values.

use std::sync::{Arc, OnceLock};

use litchi_core::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as CoreWorksheet, WorksheetIterator,
};
use litchi_core::{FileFormat, Metadata};
use litchi_opc::{OpcPackage, PreparedOoxmlSource, SourceBackedPackage};
use litchi_xlsx as xlsx;

use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

#[allow(
    clippy::duplicate_mod,
    reason = "shared with the sibling workbook handle until S6 can declare it once in `format/mod.rs`"
)]
#[path = "cell_text.rs"]
mod cell_text;
mod selection;
mod sheet;

pub use selection::{SelectedCell, SelectedCellView, SelectedWorksheet, WorksheetSelector};
use sheet::{Sheet, boxed};

/// An opened Excel OOXML workbook.
///
/// Cheap to clone and `Send + Sync`. Read methods take `&self`; on a
/// source-backed workbook each read checks that the file has not changed
/// since it was opened and reports [`Code::IoError`] when it has.
///
/// ```no_run
/// # fn demo(workbook: &litchi_ops::format::xlsx::Xlsx) -> litchi_ops::Result<()> {
/// println!("{:?}", workbook.worksheet_names()?);
/// let text = workbook.text()?;
/// // Sheet names match case-insensitively, as in Excel.
/// let data = workbook.sheet("dAtA")?;
/// # let _ = (text, data);
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct Xlsx(Arc<XlsxState>);

/// Shared state behind an [`Xlsx`] handle.
struct XlsxState {
    owner: XlsxOwner,
    /// Worksheet names in catalog order, built on the first generic-trait
    /// request rather than at open.
    names: OnceLock<Box<[String]>>,
    /// Core properties, read at open as the facade does.
    metadata: Metadata,
    /// Repairs of that core-properties read, folded into
    /// [`Xlsx::tolerance_report`].
    metadata_repairs: MetadataRepairs,
    /// Whether write verbs may build an editor ([`OpenOptions::retain_source`]).
    /// Both owner states already hold everything an edit needs, so this
    /// releases nothing; it only refuses.
    retain_source: bool,
}

/// The two owner states of an XLSX handle (record 0003 §2.3).
pub(crate) enum XlsxOwner {
    /// Deferred catalog over the retained positional source.
    Source(xlsx::SourceBackedWorkbook),
    /// The owning workbook snapshot.
    Eager(xlsx::Workbook),
}

/// The edit entry an [`Xlsx`] handle hands to the write kernel (M2).
#[cfg_attr(
    not(test),
    expect(dead_code, reason = "consumed by the M2 write kernel")
)]
pub(crate) enum XlsxEditor {
    /// The package indexed at open, recovered without re-indexing (L-XLSX-1).
    /// Every per-feature `SourceBackedEditor` adopts it through
    /// `from_source_backed_package`.
    Source(Box<SourceBackedPackage>),
    /// The owning snapshot; [`xlsx::Workbook::edit`] starts a transaction.
    Eager(xlsx::Workbook),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Xlsx>();
    assert!(size_of::<Xlsx>() == size_of::<usize>());
};

impl XlsxOwner {
    /// The facade's `ensure_source_current`: refuse a source-backed owner
    /// whose file changed since it was opened.
    pub(crate) fn ensure_current(&self) -> Result<()> {
        match self {
            Self::Source(workbook) => workbook.source_version().map(drop).map_err(Error::from),
            Self::Eager(_) => Ok(()),
        }
    }

    /// Bracket `read` with freshness checks; a changed source wins over the
    /// read's own result.
    pub(crate) fn finish<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        self.ensure_current()?;
        let result = read();
        self.ensure_current()?;
        result
    }

    fn len(&self) -> usize {
        match self {
            Self::Source(workbook) => workbook.len(),
            Self::Eager(workbook) => workbook.len(),
        }
    }

    fn names(&self) -> Box<[String]> {
        match self {
            Self::Source(workbook) => workbook
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect(),
            Self::Eager(workbook) => workbook
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect(),
        }
    }

    fn select<'a>(&self, selector: impl Into<WorksheetSelector<'a>>) -> Result<Option<Sheet>> {
        Ok(match self {
            Self::Source(workbook) => workbook.sheet(selector)?.map(Sheet::source),
            Self::Eager(workbook) => workbook.sheet(selector)?.map(Sheet::owned),
        })
    }

    fn active(&self) -> Option<Sheet> {
        match self {
            Self::Source(workbook) => workbook.active_sheet().map(Sheet::source),
            Self::Eager(workbook) => workbook.active_sheet().map(Sheet::owned),
        }
    }
}

impl Xlsx {
    /// Adopt the source-backed catalog the filesystem detector built, with
    /// the metadata it read. No worksheet payload is read.
    ///
    /// Under [`SourceMode::Eager`] the catalog's package is materialized into
    /// the owning snapshot instead.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_source_workbook(
        workbook: xlsx::SourceBackedWorkbook,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
    ) -> Result<Self> {
        let owner = if options.source_mode == SourceMode::Eager {
            XlsxOwner::Eager(materialize(workbook.source_backed_package())?)
        } else {
            XlsxOwner::Source(workbook)
        };
        Ok(Self::with_owner(owner, metadata, metadata_repairs, options))
    }

    /// Adopt the source-backed package the byte detector validated: read its
    /// core properties, then build the deferred catalog over it.
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let (metadata, metadata_repairs) =
            read_source_backed_metadata(&package).map_err(metadata_error)?;
        let owner = if options.source_mode == SourceMode::Eager {
            XlsxOwner::Eager(materialize(&package)?)
        } else {
            XlsxOwner::Source(xlsx::SourceBackedWorkbook::from_source_backed_package(
                package,
            )?)
        };
        Ok(Self::with_owner(owner, metadata, metadata_repairs, options))
    }

    /// Adopt a package prepared by `detect_and_prepare` without indexing the
    /// container again.
    ///
    /// Returns [`Code::UnsupportedFormat`] naming the detected format when
    /// the prepared package is another OOXML family, and [`Code::IoError`]
    /// when the file changed after it was prepared.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_prepared(
        prepared: PreparedOoxmlSource,
        options: &OpenOptions,
    ) -> Result<Self> {
        let detected =
            crate::detect::ooxml::format_for_main_content_type(prepared.main_part_content_type())
                .ok_or(litchi_core::Error::NotOfficeFile)?;
        if detected != FileFormat::Xlsx {
            return Err(litchi_core::Error::UnexpectedFormat { detected }.into());
        }
        Self::from_source_package(prepared.into_package()?, options)
    }

    /// Build the owning snapshot from the OPC package the byte detector
    /// already parsed (`DetectedFormat::Xlsx`).
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        let mut metadata_repairs = MetadataRepairs::default();
        let metadata =
            litchi_ooxml_common::properties::read_with_report(&package, &mut metadata_repairs)
                .map_err(metadata_error)?
                .map(Metadata::from)
                .unwrap_or_default();
        let workbook = xlsx::Package::from_opc(package)?.into_workbook()?;
        Ok(Self::with_owner(
            XlsxOwner::Eager(workbook),
            metadata,
            metadata_repairs,
            options,
        ))
    }

    fn with_owner(
        owner: XlsxOwner,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
    ) -> Self {
        Self(Arc::new(XlsxState {
            owner,
            names: OnceLock::new(),
            metadata,
            metadata_repairs,
            retain_source: options.retain_source,
        }))
    }

    /// The two-state owner, for the `File` dispatch of S6.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the `File` dispatch from S6")
    )]
    pub(crate) fn owner(&self) -> &XlsxOwner {
        &self.0.owner
    }

    /// The format of this handle.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        FileFormat::Xlsx
    }

    /// Worksheet names in workbook order, every sheet kind included.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        self.0.owner.ensure_current()?;
        Ok(self.names().to_vec())
    }

    /// The number of sheets in the workbook catalog.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.0.owner.ensure_current()?;
        Ok(self.0.owner.len())
    }

    /// All cell text: each row of each sheet on its own line, cells
    /// separated by tabs. Non-grid sheets contribute nothing.
    ///
    /// # Errors
    ///
    /// Returns the XLSX error of the first worksheet that fails to read, or
    /// [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn text(&self) -> Result<String> {
        let owner = &self.0.owner;
        owner.finish(|| {
            let mut out = String::new();
            for index in 0..owner.len() {
                let sheet = owner
                    .select(index)
                    .and_then(|sheet| sheet.ok_or_else(|| sheet_index_out_of_bounds(index)))?;
                for row in 0..sheet.row_end()? {
                    cell_text::append_row(&mut out, &sheet.row_values(row as u32)?);
                }
            }
            Ok(out)
        })
    }

    /// The core document properties read at open.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn metadata(&self) -> Result<Metadata> {
        self.0.owner.ensure_current()?;
        Ok(self.0.metadata.clone())
    }

    /// Everything the reader repaired so far: the open (package, workbook
    /// part, styles, theme), the shared strings once read, every worksheet
    /// whose cells have been parsed, and the core-properties read at open.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn tolerance_report(&self) -> Result<xlsx::ToleranceReport> {
        self.0.owner.ensure_current()?;
        let mut report = match &self.0.owner {
            XlsxOwner::Source(workbook) => workbook.tolerance_report(),
            XlsxOwner::Eager(workbook) => workbook.tolerance_report(),
        };
        report.append_properties(&self.0.metadata_repairs);
        Ok(report)
    }

    /// Select one worksheet by name or zero-based position.
    ///
    /// Names match case-insensitively, as XLSX defines them; `"dAtA"` selects
    /// `Data`. Selecting is catalog-only on a source-backed workbook: cell and
    /// range reads are deferred until requested through the returned
    /// [`SelectedWorksheet`]. A missing name or position returns `Ok(None)`.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed,
    /// and the XLSX error for an unsupported selector form.
    pub fn sheet<'a>(
        &self,
        selector: impl Into<WorksheetSelector<'a>>,
    ) -> Result<Option<SelectedWorksheet>> {
        let selector = selector.into();
        self.0.owner.finish(|| {
            Ok(match &self.0.owner {
                XlsxOwner::Source(workbook) => workbook
                    .sheet(selector)?
                    .map(SelectedWorksheet::from_source),
                XlsxOwner::Eager(workbook) => {
                    workbook.sheet(selector)?.map(SelectedWorksheet::from_owned)
                },
            })
        })
    }

    /// Build the edit entry for the write kernel (M2) from what the handle
    /// retains, without detecting or indexing the input again.
    ///
    /// A source-backed handle that is the sole owner of its catalog returns
    /// the package indexed at open (L-XLSX-1). While another clone of the
    /// handle, or a [`SelectedWorksheet`], shares the catalog, the package is
    /// instead materialized through the same index into the owning snapshot.
    /// An eager handle returns its snapshot, which is `Clone`.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedCapability`] when the handle was opened with
    /// [`OpenOptions::retain_source`] `false`, [`Code::IoError`] when the file
    /// changed, and the XLSX error of a failed materialization.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "consumed by the M2 write kernel")
    )]
    pub(crate) fn into_editor(self) -> Result<XlsxEditor> {
        if !self.0.retain_source {
            return Err(source_released());
        }
        self.0.owner.ensure_current()?;
        let state = match Arc::try_unwrap(self.0) {
            Ok(state) => state,
            Err(shared) => {
                return Ok(match &shared.owner {
                    XlsxOwner::Source(workbook) => {
                        XlsxEditor::Eager(materialize(workbook.source_backed_package())?)
                    },
                    XlsxOwner::Eager(workbook) => XlsxEditor::Eager(workbook.clone()),
                });
            },
        };
        match state.owner {
            XlsxOwner::Source(workbook) => match workbook.try_into_source_backed_package() {
                Ok(package) => Ok(XlsxEditor::Source(Box::new(package))),
                Err(workbook) => Ok(XlsxEditor::Eager(materialize(
                    workbook.source_backed_package(),
                )?)),
            },
            XlsxOwner::Eager(workbook) => Ok(XlsxEditor::Eager(workbook)),
        }
    }

    fn names(&self) -> &[String] {
        self.0.names.get_or_init(|| self.0.owner.names())
    }

    fn sheet_at(&self, index: usize) -> Result<Sheet> {
        self.0
            .owner
            .select(index)?
            .ok_or_else(|| sheet_index_out_of_bounds(index))
    }
}

/// Build the owning snapshot from a source-backed package through its
/// existing index.
fn materialize(package: &SourceBackedPackage) -> Result<xlsx::Workbook> {
    Ok(xlsx::Package::from_opc(package.to_opc_package()?)?.into_workbook()?)
}

/// Classify a core-properties read failure. OPC failures keep their own
/// class (a changed source is [`Code::IoError`]); other failures are XLSX
/// package errors.
fn metadata_error(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Opc(error) => Error::from(error).with_format(FileFormat::Xlsx),
        error => Error::from(xlsx::Error::from(error)),
    }
}

fn sheet_index_out_of_bounds(index: usize) -> Error {
    Error::new(
        Code::OutOfRange,
        format!("XLSX worksheet index {index} is out of bounds"),
    )
    .with_format(FileFormat::Xlsx)
}

fn source_released() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "the workbook was opened without retaining its source for editing",
    )
    .with_format(FileFormat::Xlsx)
    .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit")
}

impl std::fmt::Debug for Xlsx {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Xlsx")
            .field("worksheet_count", &self.0.owner.len())
            .field(
                "source_backed",
                &matches!(self.0.owner, XlsxOwner::Source(_)),
            )
            .finish_non_exhaustive()
    }
}

impl WorkbookTrait for Xlsx {
    fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let sheet = self.0.owner.active().ok_or_else(|| {
            boxed(
                Error::new(Code::NotFound, "XLSX workbook has no active worksheet")
                    .with_format(FileFormat::Xlsx),
            )
        })?;
        Ok(Box::new(sheet))
    }

    fn worksheet_names(&self) -> &[String] {
        self.names()
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let sheet = self.0.owner.select(name).map_err(boxed)?.ok_or_else(|| {
            boxed(
                Error::new(
                    Code::NotFound,
                    format!("XLSX worksheet '{name}' was not found"),
                )
                .with_format(FileFormat::Xlsx),
            )
        })?;
        Ok(Box::new(sheet))
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        Ok(Box::new(self.sheet_at(index).map_err(boxed)?))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(Worksheets {
            workbook: self,
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.0.owner.len()
    }

    fn active_sheet_index(&self) -> usize {
        self.0.owner.active().map_or(0, |sheet| sheet.position())
    }

    fn is_1904_date_system(&self) -> bool {
        let date_system = match &self.0.owner {
            XlsxOwner::Source(workbook) => workbook.date_system(),
            XlsxOwner::Eager(workbook) => workbook.date_system(),
        };
        matches!(date_system, xlsx::DateSystem::Excel1904)
    }
}

struct Worksheets<'a> {
    workbook: &'a Xlsx,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for Worksheets<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreWorksheet + 'a>>> {
        if self.index >= self.workbook.0.owner.len() {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(WorkbookTrait::worksheet_by_index(self.workbook, index))
    }
}

#[cfg(test)]
mod tests;
