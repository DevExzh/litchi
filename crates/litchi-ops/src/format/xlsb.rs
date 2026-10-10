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
//! | Open route | Detection result | Constructor |
//! |---|---|---|
//! | `open` (path) on `unix`/`windows`, `open_bytes`, `open_prepared` | source-backed OPC package | `Xlsb::from_source_package` |
//! | [`SourceMode::Eager`], and paths elsewhere | eager `OpcPackage` | `Xlsb::from_opc_package` |
//!
//! XLSB has no decryption support, so `Locked::unlock` does not reach this
//! handle.
//!
//! # Sheet names
//!
//! Unified paths match sheet names with Unicode caseless normalization.
//! Ambiguous names are refused while positional paths remain available.
//! The legacy `worksheet_by_name` trait retains exact leaf-name matching.
//!
//! # Editing
//!
//! XLSB has no source-backed edit API. At the first edit the write kernel
//! (M2) asks for the eager [`xlsb::Workbook`], which a source-backed handle
//! materializes from the package indexed at open (L-XLSB-1).

use std::ops::Deref;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as CoreWorksheet, WorksheetIterator,
};
use litchi_core::{FileFormat, Metadata};
use litchi_opc::{OpcPackage, SourceBackedPackage, SourceLimits};
use litchi_xlsb as xlsb;

use super::cell_text;
use super::signature::{self, Opened, Signature};
use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

mod children;
mod header;
mod nav;
mod node;
mod owner;
#[cfg(test)]
mod read_tests;
mod source;
mod spec;
mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::SPEC;
pub(crate) use views::{Json as ViewJson, view_json, view_to};

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
pub struct Xlsb(Arc<Opened<XlsbState>>);

/// Shared state behind an [`Xlsb`] handle.
struct XlsbState {
    owner: XlsbOwner,
    /// Semantic catalog and Unicode name index; contains no worksheet payload.
    catalog: OnceLock<Arc<owner::Catalog>>,
    /// Shared semantic workbook names; child reads borrow one owner snapshot.
    defined_names: OnceLock<Arc<[String]>>,
    /// Core properties, read at open as the facade does.
    metadata: Metadata,
    /// Repairs of that core-properties read, folded into
    /// [`Xlsb::tolerance_report`].
    metadata_repairs: MetadataRepairs,
    /// Whether write verbs may build an editor ([`OpenOptions::retain_source`]).
    /// Both owner states already hold everything an edit needs, so this
    /// releases nothing; it only refuses.
    retain_source: bool,
    /// The OPC read limits of the open, which an unsigned copy is opened
    /// under (record 0013).
    ooxml_limits: SourceLimits,
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
    /// Materialized from the source-backed catalog (L-XLSB-1). After
    /// [`Signed::unsign`](crate::Signed::unsign), either owner yields an
    /// owning copy whose signature the leaf removed explicitly
    /// ([`xlsb::Workbook::unsign`], record 0013).
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
    pub(crate) fn ensure_current(&self) -> Result<()> {
        match self {
            Self::Source(workbook) => workbook.ensure_current(),
            Self::Eager(_) => Ok(()),
        }
    }
}

impl Xlsb {
    /// The stable path origin captured by a path open, if any.
    pub(crate) fn origin(&self) -> Option<Arc<Path>> {
        self.0.origin()
    }

    /// Attach path metadata to a handle after a path route has finished
    /// classification. Byte and create routes leave this unset.
    pub(crate) fn with_origin(&self, origin: Option<Arc<Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    /// Whether this handle is the same immutable open snapshot as `other`.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Adopt a validated native package without serializing and reopening it.
    /// The unified handle retains the source policy, metadata sidecar, origin,
    /// and signature transition of the input. A removal-authorized edit has
    /// already removed the native signature graph and therefore becomes an
    /// ordinary unsigned eager owner.
    pub(crate) fn adopt_package(&self, package: xlsb::Package) -> Result<Self> {
        let options = OpenOptions::new()
            .retain_source(self.0.retain_source)
            .ooxml_limits(self.0.ooxml_limits);
        let workbook = package.into_workbook()?;
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Ok(Self::with_owner(
            XlsbOwner::Eager(Box::new(workbook)),
            self.0.metadata.clone(),
            self.0.metadata_repairs.clone(),
            &options,
        )
        .with_origin(self.origin())
        .with_signature(signature))
    }

    /// Check signature disposition and retained-source freshness before a
    /// mutation. A signed handle refuses before target or property admission,
    /// including an otherwise exact no-op.
    pub(crate) fn check_source(&self) -> Result<()> {
        if self.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xlsb));
        }
        self.0.owner.ensure_current()
    }

    /// Admit the native structural editor without publishing a candidate.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        let editor = self.editor()?;
        let package = match editor {
            XlsbEditor::Eager(workbook) => workbook.to_package(),
            XlsbEditor::Materialized(workbook) => workbook.into_package(),
        };
        package
            .edit()
            .map(|_| ())
            .map_err(|error| Error::from(error).with_format(FileFormat::Xlsb))?;
        self.check_source()
    }

    /// The one-shot unified worksheet-property setter.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
    ) -> Result<write::Commit> {
        self.set_with_limits(
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Set unified worksheet properties under an explicit semantic budget.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        self.check_source()?;
        let result = write::set_with_limits(self, target, props, limits);
        self.check_source()?;
        result
    }

    /// Stream this immutable snapshot to a sequential sink.
    pub(crate) fn write_to<W: std::io::Write>(&self, sink: W) -> Result<()> {
        write::write_to(self, sink)
    }

    /// Atomically save this immutable snapshot to a filesystem path.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::save::SaveOptions) -> Result<u64> {
        write::save_to(self, path, options)
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        self.0.retain_source
    }

    /// Adopt the source-backed package detection validated: read its core
    /// properties, then build the deferred catalog over it. Only the
    /// workbook catalog is read.
    ///
    /// Under [`SourceMode::Eager`] the catalog is materialized into the
    /// owning workbook instead; the openers reach that only from
    /// `open_prepared`, because their other eager routes call
    /// [`Self::from_opc_package`].
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

    /// Build the owning workbook from the OPC package the byte detector
    /// already parsed (`DetectedFormat::Xlsb`).
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
        Self(Opened::new(XlsbState {
            owner,
            catalog: OnceLock::new(),
            defined_names: OnceLock::new(),
            metadata,
            metadata_repairs,
            retain_source: options.retain_source,
            ooxml_limits: options.ooxml_limits.unwrap_or_default(),
        }))
    }

    /// The same workbook under another signature disposition (record 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
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
        let names = WorkbookTrait::worksheet_names(self).to_vec();
        self.0.owner.ensure_current()?;
        Ok(names)
    }

    /// The number of worksheets.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.0.owner.ensure_current()?;
        let count = WorkbookTrait::worksheet_count(self);
        self.0.owner.ensure_current()?;
        Ok(count)
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

    /// The core document properties read at open. Always `Some` when it
    /// succeeds; a package without core properties reads as empty metadata,
    /// as the facade reports it.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        self.0.owner.ensure_current()?;
        Ok(Some(self.0.metadata.clone()))
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
                let mut report = workbook.tolerance_report()?;
                report.append_properties(&self.0.metadata_repairs);
                workbook.ensure_current()?;
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
    /// positional source (L-XLSB-1); the handle stays usable. After
    /// [`Signed::unsign`](crate::Signed::unsign) either owner yields an owning
    /// copy whose signature the leaf removed explicitly (record 0013): opened
    /// over the exact input bytes, or copied from the eager owner's package,
    /// then `unsign`ed.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`] while the workbook is still
    /// signed, with a suggestion naming `Signed::unsign`;
    /// [`Code::UnsupportedCapability`] when the handle was opened with
    /// [`OpenOptions::retain_source`] `false`; [`Code::IoError`] when the
    /// file changed; and the XLSB error of a failed materialization.
    pub(crate) fn editor(&self) -> Result<XlsbEditor<'_>> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xlsb));
        }
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        if self.0.signature() == Signature::RemovalAuthorized {
            return self
                .unsigned_workbook()
                .map(|workbook| XlsbEditor::Materialized(Box::new(workbook)));
        }
        match &self.0.owner {
            XlsbOwner::Source(workbook) => Ok(XlsbEditor::Materialized(Box::new(
                workbook.workbook().to_workbook()?,
            ))),
            XlsbOwner::Eager(workbook) => Ok(XlsbEditor::Eager(workbook)),
        }
    }

    /// An owning copy of the workbook whose signature the leaf removed
    /// explicitly ([`xlsb::Workbook::unsign`]).
    fn unsigned_workbook(&self) -> Result<xlsb::Workbook> {
        let mut workbook = match &self.0.owner {
            XlsbOwner::Source(workbook) => {
                let bytes = signature::exact_source(workbook.workbook().source_backed_package())?;
                xlsb::Package::from_bytes_with_limits(bytes, self.0.ooxml_limits)?
                    .into_workbook()?
            },
            XlsbOwner::Eager(workbook) => {
                xlsb::Workbook::from_opc_package(workbook.opc_package().clone())?
            },
        };
        workbook.unsign();
        Ok(workbook)
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
