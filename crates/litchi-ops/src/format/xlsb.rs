//! Handle for Excel binary workbook (`.xlsb`) inputs.
//!
//! [`Xlsb`] is one `Arc` over its state: the two-state owner of record 0003
//! §2.3 (`Source` for the deferred, source-backed catalog with its
//! per-worksheet materialization cache; `Eager` for the owning workbook), the
//! core-properties metadata and its repairs read at open (§1.6), and the
//! retention policy from [`OpenOptions`]. Cloning bumps one reference count.
//!
//! # Routes
//!
//! The constructors take what detection already produced, so an open does
//! one detection pass and one container index:
//!
//! | Facade route today | Detection result | Constructor |
//! |---|---|---|
//! | `Workbook::open` on `unix`/`windows` | `WorkbookSourcePathDetection::Xlsb` | `Xlsb::from_source_workbook` |
//! | `Workbook::from_bytes` | `WorkbookSourceBytesDetection::Xlsb` | `Xlsb::from_source_package` |
//! | `Workbook::from_bytes` fallback (`detect_format_smart`) | `DetectedFormat::Xlsb` | `Xlsb::from_opc_package` |
//! | `Workbook::from_prepared` | `PreparedOoxmlSource` | `Xlsb::from_prepared` |
//!
//! The facade has no reader route for workbooks. The `sheet::open_xlsb_*`
//! functions are path and byte routes onto the same two owners.
//!
//! # Sheet names
//!
//! XLSB sheet names match exactly (record 0003 finding F15): unlike
//! [`Xlsx`](super::xlsx::Xlsx), `worksheet_by_name("simple")` does not find a
//! sheet named `Simple`. This handle keeps that leaf semantics unchanged.
//!
//! # Editing
//!
//! XLSB has no source-backed edit API. At the first edit the write kernel
//! (M2) asks for the eager [`xlsb::Workbook`], which a source-backed handle
//! materializes from the package indexed at open (L-XLSB-1).

use std::ops::Deref;
use std::sync::Arc;

use litchi_core::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as CoreWorksheet, WorksheetIterator,
};
use litchi_core::{FileFormat, Metadata};
use litchi_opc::{OpcPackage, PreparedOoxmlSource, SourceBackedPackage};
use litchi_xlsb as xlsb;

use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

#[allow(
    clippy::duplicate_mod,
    reason = "shared with the sibling workbook handle until S6 can declare it once in `format/mod.rs`"
)]
#[path = "cell_text.rs"]
mod cell_text;
mod source;

use source::{SourceWorkbook, SourceWorksheets, internal};

/// An opened Excel binary workbook.
///
/// Cheap to clone and `Send + Sync`. Read methods take `&self`; on a
/// source-backed workbook each read checks that the file has not changed
/// since it was opened and reports [`Code::IoError`] when it has.
///
/// ```no_run
/// use litchi_core::sheet::WorkbookTrait;
///
/// # fn demo(workbook: &litchi_ops::format::xlsb::Xlsb) -> litchi_ops::Result<()> {
/// println!("{:?}", workbook.worksheet_names()?);
/// let text = workbook.text()?;
/// // Generic sheet access, with exact name matching.
/// let first = WorkbookTrait::worksheet_by_index(workbook, 0);
/// # let _ = (text, first);
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct Xlsb(Arc<XlsbState>);

/// Shared state behind an [`Xlsb`] handle.
struct XlsbState {
    owner: XlsbOwner,
    /// Core properties, read at open as the facade does.
    metadata: Metadata,
    /// Repairs of that core-properties read, folded into
    /// [`Xlsb::tolerance_report`].
    metadata_repairs: MetadataRepairs,
    /// Whether write verbs may build an editor ([`OpenOptions::retain_source`]).
    /// Both owner states already hold everything an edit needs, so this
    /// releases nothing; it only refuses.
    retain_source: bool,
}

/// The two owner states of an XLSB handle (record 0003 §2.3).
enum XlsbOwner {
    /// Deferred catalog over the retained positional source.
    Source(SourceWorkbook),
    /// The owning workbook.
    Eager(Box<xlsb::Workbook>),
}

/// The eager workbook the write kernel (M2) edits through
/// [`xlsb::Workbook::edit_cell_values`] and the other eager edit entries.
pub(crate) enum XlsbEditor<'a> {
    /// The handle's own eager owner.
    Eager(&'a xlsb::Workbook),
    /// Materialized from the source-backed catalog (L-XLSB-1).
    Materialized(Box<xlsb::Workbook>),
}

impl Deref for XlsbEditor<'_> {
    type Target = xlsb::Workbook;

    fn deref(&self) -> &xlsb::Workbook {
        match self {
            Self::Eager(workbook) => workbook,
            Self::Materialized(workbook) => workbook,
        }
    }
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Xlsb>();
    assert!(size_of::<Xlsb>() == size_of::<usize>());
};

impl XlsbOwner {
    /// The facade's `ensure_source_current`: refuse a source-backed owner
    /// whose file changed since it was opened.
    fn ensure_current(&self) -> Result<()> {
        match self {
            Self::Source(workbook) => workbook.ensure_current(),
            Self::Eager(_) => Ok(()),
        }
    }
}

impl Xlsb {
    /// Adopt the source-backed catalog the filesystem detector built, with
    /// the metadata it read. Only the workbook catalog is read.
    ///
    /// Under [`SourceMode::Eager`] the catalog is materialized into the
    /// owning workbook instead.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_source_workbook(
        workbook: xlsb::SourceBackedWorkbook,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
    ) -> Result<Self> {
        Ok(Self::with_owner(
            Self::source_owner(workbook, options)?,
            metadata,
            metadata_repairs,
            options,
        ))
    }

    /// Adopt the source-backed package the byte detector validated: read its
    /// core properties, then build the deferred catalog over it.
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let (metadata, metadata_repairs) =
            read_source_backed_metadata(&package).map_err(metadata_error)?;
        let workbook = xlsb::SourceBackedWorkbook::from_source_backed_package(package)?;
        Ok(Self::with_owner(
            Self::source_owner(workbook, options)?,
            metadata,
            metadata_repairs,
            options,
        ))
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
        if detected != FileFormat::Xlsb {
            return Err(litchi_core::Error::UnexpectedFormat { detected }.into());
        }
        Self::from_source_package(prepared.into_package()?, options)
    }

    /// Build the owning workbook from the OPC package the byte detector
    /// already parsed (`DetectedFormat::Xlsb`).
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
        let workbook = xlsb::Package::from_opc(package)?.into_workbook()?;
        Ok(Self::with_owner(
            XlsbOwner::Eager(Box::new(workbook)),
            metadata,
            metadata_repairs,
            options,
        ))
    }

    fn source_owner(
        workbook: xlsb::SourceBackedWorkbook,
        options: &OpenOptions,
    ) -> Result<XlsbOwner> {
        Ok(if options.source_mode == SourceMode::Eager {
            XlsbOwner::Eager(Box::new(workbook.to_workbook()?))
        } else {
            XlsbOwner::Source(SourceWorkbook::new(workbook)?)
        })
    }

    fn with_owner(
        owner: XlsbOwner,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
    ) -> Self {
        Self(Arc::new(XlsbState {
            owner,
            metadata,
            metadata_repairs,
            retain_source: options.retain_source,
        }))
    }

    /// The format of this handle.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        FileFormat::Xlsb
    }

    /// Whether the handle reads through the retained source rather than an
    /// owning workbook.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the `File` tests from S6")
    )]
    pub(crate) fn is_source_backed(&self) -> bool {
        matches!(self.0.owner, XlsbOwner::Source(_))
    }

    /// Worksheet names in workbook order. Chart, dialog and macro sheets
    /// are not worksheets and are not listed.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        self.0.owner.ensure_current()?;
        Ok(WorkbookTrait::worksheet_names(self).to_vec())
    }

    /// The number of worksheets.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.0.owner.ensure_current()?;
        Ok(WorkbookTrait::worksheet_count(self))
    }

    /// All worksheet text: each row on its own line, cells separated by
    /// tabs. Chart sheets contribute nothing.
    ///
    /// # Errors
    ///
    /// Returns the XLSB error of the first worksheet that fails to read, or
    /// [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => Ok(workbook.workbook().text()?),
            XlsbOwner::Eager(workbook) => {
                let mut out = String::new();
                for index in 0..WorkbookTrait::worksheet_count(workbook.as_ref()) {
                    let worksheet = workbook.worksheet(index)?;
                    let mut rows = worksheet.rows();
                    while let Some(row) = rows.next() {
                        cell_text::append_row(&mut out, &row.map_err(sheet_error)?);
                    }
                }
                Ok(out)
            },
        }
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

    /// Everything the reader repaired, as the XLSB leaf records it, followed
    /// by the core-properties read at open.
    ///
    /// A source-backed workbook reports the workbook-scope repairs so far
    /// (the styles part joins once a worksheet first needs it), then every
    /// worksheet it has materialized so far. An eager workbook reports its
    /// open, then every worksheet, which this call materializes.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file
    /// changed, and the XLSB error when an eager worksheet cannot be
    /// materialized.
    pub fn tolerance_report(&self) -> Result<xlsb::ToleranceReport> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => {
                workbook.ensure_current()?;
                let mut report = workbook.tolerance_report();
                report.append_properties(&self.0.metadata_repairs);
                Ok(report)
            },
            XlsbOwner::Eager(workbook) => {
                let mut report = workbook.tolerance_report().clone();
                report.append_properties(&self.0.metadata_repairs);
                for index in 0..WorkbookTrait::worksheet_count(workbook.as_ref()) {
                    report.append(workbook.worksheet(index)?.tolerance_report());
                }
                Ok(report)
            },
        }
    }

    /// The eager workbook the write kernel (M2) edits, without detecting or
    /// indexing the input again.
    ///
    /// An eager handle lends its own workbook. A source-backed handle
    /// materializes one from the package indexed at open through the same
    /// positional source (L-XLSB-1); the handle stays usable.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedCapability`] when the handle was opened with
    /// [`OpenOptions::retain_source`] `false`, [`Code::IoError`] when the file
    /// changed, and the XLSB error of a failed materialization.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "consumed by the M2 write kernel")
    )]
    pub(crate) fn editor(&self) -> Result<XlsbEditor<'_>> {
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        match &self.0.owner {
            XlsbOwner::Source(workbook) => Ok(XlsbEditor::Materialized(Box::new(
                workbook.workbook().to_workbook()?,
            ))),
            XlsbOwner::Eager(workbook) => Ok(XlsbEditor::Eager(workbook)),
        }
    }
}

/// Classify a core-properties read failure. OPC failures keep their own
/// class (a changed source is [`Code::IoError`]); other failures are XLSB
/// package errors.
fn metadata_error(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Opc(error) => Error::from(error).with_format(FileFormat::Xlsb),
        error => Error::from(xlsb::package::PackageError::from(error)),
    }
}

/// Recover the typed error from a leaf `litchi_core::sheet` trait error.
fn sheet_error(error: Box<dyn std::error::Error + Send + Sync>) -> Error {
    match error.downcast::<xlsb::package::PackageError>() {
        Ok(error) => Error::from(*error),
        Err(error) => match error.downcast::<Error>() {
            Ok(error) => *error,
            Err(error) => internal(error.to_string()),
        },
    }
}

impl std::fmt::Debug for Xlsb {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut debug = formatter.debug_struct("Xlsb");
        match &self.0.owner {
            XlsbOwner::Source(workbook) => debug.field("source", workbook),
            XlsbOwner::Eager(workbook) => debug.field(
                "worksheet_count",
                &WorkbookTrait::worksheet_count(workbook.as_ref()),
            ),
        };
        debug.finish_non_exhaustive()
    }
}

impl WorkbookTrait for Xlsb {
    fn active_sheet_index(&self) -> usize {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.active_worksheet_ordinal().unwrap_or(0),
            XlsbOwner::Eager(workbook) => workbook.active_sheet_index(),
        }
    }

    fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.active_worksheet(),
            XlsbOwner::Eager(workbook) => workbook.active_worksheet(),
        }
    }

    fn is_1904_date_system(&self) -> bool {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.date1904(),
            XlsbOwner::Eager(workbook) => workbook.is_1904_date_system(),
        }
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.worksheet(index),
            XlsbOwner::Eager(workbook) => workbook.worksheet_by_index(index),
        }
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.worksheet_named(name),
            XlsbOwner::Eager(workbook) => workbook.worksheet_by_name(name),
        }
    }

    fn worksheet_count(&self) -> usize {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.names().len(),
            XlsbOwner::Eager(workbook) => workbook.worksheet_count(),
        }
    }

    fn worksheet_names(&self) -> &[String] {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.names(),
            XlsbOwner::Eager(workbook) => workbook.worksheet_names(),
        }
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => Box::new(SourceWorksheets { workbook, index: 0 }),
            XlsbOwner::Eager(workbook) => workbook.worksheets(),
        }
    }
}

#[cfg(test)]
mod tests;
