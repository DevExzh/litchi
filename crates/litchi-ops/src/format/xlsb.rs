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
//! (M2) asks for the eager [`xlsb::Package`], which a source-backed handle
//! materializes from the package indexed at open (L-XLSB-1). Every edit
//! family of the package commits a new package, which the handle adopts
//! together with its already parsed read-only [`xlsb::Workbook`] projection.

use std::ops::Deref;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::sheet::WorkbookTrait;
use litchi_core::{FileFormat, Metadata};
use litchi_opc::{OpcPackage, SourceBackedPackage, SourceLimits};
use litchi_xlsb as xlsb;

use super::cell_text;
use super::signature::{self, Opened, Signature};
use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

#[cfg(feature = "eval")]
pub(crate) mod calc;
mod chart;
mod children;
pub(crate) mod create;
mod feature_write;
mod features;
mod header;
mod merge_write;
mod metadata;
mod name_write;
mod nav;
mod node;
mod owner;
#[cfg(test)]
mod read_tests;
pub(crate) mod source;
mod spec;
mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use features::FEATURE_ELEMENTS;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::{CELL_WRITE_TABLE, SPEC};
pub(crate) use views::{Json as ViewJson, view_json, view_to};

use crate::view::sheet::{Worksheet as Sheet, WorksheetsInner};
use source::{SourceWorkbook, SourceWorksheets};

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
    /// The owning package. Every constructor admits a package whose
    /// read-only projection is already parsed ([`eager_package`]), so
    /// [`projection`] borrows it without parsing.
    Eager(Box<xlsb::Package>),
}

/// The owning package the write kernel (M2) hands to an edit family.
pub(crate) enum XlsbEditor<'a> {
    /// The handle's own eager owner.
    Eager(&'a xlsb::Package),
    /// Materialized from the source-backed catalog (L-XLSB-1). After
    /// [`Signed::unsign`](crate::Signed::unsign), either owner yields an
    /// owning copy whose signature the leaf removed explicitly
    /// ([`xlsb::Package::unsign`], record 0013).
    Materialized(Box<xlsb::Package>),
}

impl Deref for XlsbEditor<'_> {
    type Target = xlsb::Package;

    fn deref(&self) -> &xlsb::Package {
        match self {
            Self::Eager(package) => package,
            Self::Materialized(package) => package,
        }
    }
}

/// Admit a package as an eager owner: parse its read-only projection now,
/// once, so every later read borrows it.
fn eager_package(package: xlsb::Package) -> Result<XlsbOwner> {
    package.workbook()?;
    Ok(XlsbOwner::Eager(Box::new(package)))
}

/// The read-only projection of an eager owner's package.
pub(crate) fn projection(package: &xlsb::Package) -> Result<&xlsb::Workbook> {
    package.workbook().map_err(Into::into)
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
        let owner = eager_package(package)?;
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Ok(Self::with_owner(
            owner,
            self.0.metadata.clone(),
            self.0.metadata_repairs.clone(),
            &options,
            self.0.ooxml_limits,
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
        self.check_source_for_replay()
    }

    /// Check retained source freshness while allowing an exact no-op replay
    /// to preserve the recipient's existing signature disposition.
    pub(crate) fn check_source_for_replay(&self) -> Result<()> {
        self.0.owner.ensure_current()
    }

    /// Admit an edit without materializing a second native workbook before
    /// the selected transaction owner does so.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        Ok(())
    }

    /// Admit a setter that may prove an exact signed no-op. Changed commits
    /// are rejected after native staging; source retention and freshness are
    /// still checked before target/property admission.
    pub(crate) fn admit_edit_for_write(&self) -> Result<()> {
        self.check_source_for_replay()?;
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        Ok(())
    }

    /// The one-shot unified worksheet and cell-property setter.
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
        self.check_source_for_replay()?;
        let result = write::set_with_limits(self, target, props, limits);
        self.check_source_for_replay()?;
        if self.signature() == Signature::Signed {
            result.and_then(|commit| {
                if commit.is_noop() {
                    Ok(commit)
                } else {
                    Err(Error::signed(FileFormat::Xlsb))
                }
            })
        } else {
            result
        }
    }

    /// Add one worksheet tab under the unified structural contract.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::new_node::NewNode<'n>,
    ) -> Result<(write::Commit, litchi_core::selector::Path<'static>)> {
        self.add_with_limits(parent, node, &crate::transaction::SemanticLimits::default())
    }

    /// Add one worksheet tab under a caller-selected semantic budget.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::new_node::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(write::Commit, litchi_core::selector::Path<'static>)> {
        self.check_source()?;
        let result = write::add_with_limits(self, parent, node, limits);
        self.check_source()?;
        result
    }

    /// Remove a complete-catalog sheet, or stored cells under the unified
    /// cell, range, and query target contract.  Sheet selectors resolve once
    /// against the immutable catalog and are staged together in one native
    /// structural edit.  Missing coordinates are virtual cells and therefore
    /// exact no-ops; the native cell owner never shifts neighboring cells.
    pub(crate) fn remove(&self, target: crate::target::TargetArg<'_>) -> Result<write::Commit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Remove cells under a caller-selected semantic budget.
    pub(crate) fn remove_with_limits(
        &self,
        target: crate::target::TargetArg<'_>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        self.check_source_for_replay()?;
        let result = write::remove_with_limits(self, target, limits);
        self.check_source_for_replay()?;
        self.enforce_signed_noop(result)
    }

    /// Move one complete-catalog tab using the unified placement contract.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::placement::Placement<'q>,
    ) -> Result<write::Commit> {
        self.move_to_with_limits(
            path,
            placement,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Move one complete-catalog tab under an explicit semantic budget.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::placement::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        self.check_source_for_replay()?;
        let result = write::move_to_with_limits(self, path, placement, limits);
        self.check_source_for_replay()?;
        self.enforce_signed_noop(result)
    }

    /// Swap two complete-catalog tabs in one native structural edit.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
    }

    /// Swap two complete-catalog tabs under an explicit semantic budget.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        self.check_source_for_replay()?;
        let result = write::swap_with_limits(self, left, right, limits);
        self.check_source_for_replay()?;
        self.enforce_signed_noop(result)
    }

    fn enforce_signed_noop(&self, result: Result<write::Commit>) -> Result<write::Commit> {
        result.and_then(|commit| {
            if self.signature() != Signature::Signed || commit.is_noop() {
                Ok(commit)
            } else {
                Err(Error::signed(FileFormat::Xlsb))
            }
        })
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

    /// The OPC ingress policy retained by this handle for raw re-admission.
    pub(crate) fn ooxml_read_limits(&self) -> SourceLimits {
        self.0.ooxml_limits
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
        let ooxml_limits = package.read_limits();
        let (metadata, metadata_repairs) =
            read_source_backed_metadata(&package).map_err(metadata_error)?;
        let workbook = xlsb::SourceBackedWorkbook::from_source_backed_package(package)?;
        Ok(Self::with_owner(
            Self::source_owner(workbook, options)?,
            metadata,
            metadata_repairs,
            options,
            ooxml_limits,
        ))
    }

    /// Build the owning workbook from the OPC package the byte detector
    /// already parsed (`DetectedFormat::Xlsb`).
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        let ooxml_limits = package.read_limits();
        let mut metadata_repairs = MetadataRepairs::default();
        let metadata =
            litchi_ooxml_common::properties::read_with_report(&package, &mut metadata_repairs)
                .map_err(metadata_error)?
                .unwrap_or_default();
        let metadata = super::metadata::with_app_props(metadata, &package);
        let owner = eager_package(xlsb::Package::from_opc(package)?)?;
        Ok(Self::with_owner(
            owner,
            metadata,
            metadata_repairs,
            options,
            ooxml_limits,
        ))
    }

    /// Adopt a newly created native package without serializing and reopening
    /// its binary workbook. The native creator has already validated the
    /// package and parsed its read-only projection.
    pub(crate) fn from_package(package: xlsb::Package, options: &OpenOptions) -> Result<Self> {
        let ooxml_limits = package.opc_package().read_limits();
        let mut metadata_repairs = MetadataRepairs::default();
        let metadata = litchi_ooxml_common::properties::read_with_report(
            package.opc_package(),
            &mut metadata_repairs,
        )
        .map_err(metadata_error)?
        .unwrap_or_default();
        let metadata = super::metadata::with_app_props(metadata, package.opc_package());
        let owner = eager_package(package)?;
        Ok(Self::with_owner(
            owner,
            metadata,
            metadata_repairs,
            options,
            ooxml_limits,
        ))
    }

    fn source_owner(
        workbook: xlsb::SourceBackedWorkbook,
        options: &OpenOptions,
    ) -> Result<XlsbOwner> {
        Ok(if options.source_mode == SourceMode::Eager {
            eager_package(workbook.to_workbook()?.into_package())?
        } else {
            XlsbOwner::Source(SourceWorkbook::new(workbook)?)
        })
    }

    fn with_owner(
        owner: XlsbOwner,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
        ooxml_limits: SourceLimits,
    ) -> Self {
        Self(Opened::new(XlsbState {
            owner,
            catalog: OnceLock::new(),
            defined_names: OnceLock::new(),
            metadata,
            metadata_repairs,
            retain_source: options.retain_source,
            ooxml_limits,
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
            XlsbOwner::Eager(package) => {
                let workbook = projection(package)?;
                let mut out = String::new();
                for index in 0..WorkbookTrait::worksheet_count(workbook) {
                    let worksheet = workbook.worksheet(index)?;
                    for row in litchi_core::sheet::Worksheet::rows(&worksheet) {
                        cell_text::append_row(&mut out, &row?);
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
            XlsbOwner::Eager(package) => {
                let workbook = projection(package)?;
                let mut report = workbook.tolerance_report().clone();
                report.append_properties(&self.0.metadata_repairs);
                for index in 0..WorkbookTrait::worksheet_count(workbook) {
                    report.append(workbook.worksheet(index)?.tolerance_report());
                }
                Ok(report)
            },
        }
    }

    /// The eager workbook the write kernel (M2) edits, without detecting or
    /// indexing the input again. Cell-value transactions start from its owning
    /// [`xlsb::Package`] so the package boundary remains the publication owner.
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
        self.editor_after_source_check(false)
    }

    /// Build a native editor for a setter that may prove an exact signed
    /// no-op. The caller rejects a changed native commit before publication.
    pub(crate) fn editor_for_write(&self) -> Result<XlsbEditor<'_>> {
        if self.0.signature() == Signature::Signed && !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        self.editor_after_source_check(self.0.signature() == Signature::Signed)
    }

    /// Materialize the original native package for a canonical signed no-op
    /// patch. The normal write editor explicitly removes signature parts;
    /// this identity patch must instead be rooted in the exact source bytes.
    pub(crate) fn original_package_for_noop(&self) -> Result<xlsb::Package> {
        self.admit_edit_for_write()?;
        match &self.0.owner {
            XlsbOwner::Source(workbook) => {
                let bytes = signature::exact_source(workbook.workbook().source_backed_package())?;
                Ok(xlsb::Package::from_bytes_with_options(
                    &bytes,
                    litchi_xlsb::OpenOptions::new().source_limits(self.0.ooxml_limits),
                )?)
            },
            XlsbOwner::Eager(package) => Ok((**package).clone()),
        }
    }

    fn editor_after_source_check(&self, allow_signed: bool) -> Result<XlsbEditor<'_>> {
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the workbook was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Xlsb)
            .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit"));
        }
        if allow_signed {
            self.check_source_for_replay()?;
        } else {
            self.check_source()?;
        }
        if matches!(
            self.0.signature(),
            Signature::Signed | Signature::RemovalAuthorized
        ) {
            return self
                .unsigned_package()
                .map(|package| XlsbEditor::Materialized(Box::new(package)));
        }
        match &self.0.owner {
            XlsbOwner::Source(workbook) => Ok(XlsbEditor::Materialized(Box::new(
                workbook.workbook().to_workbook()?.into_package(),
            ))),
            XlsbOwner::Eager(package) => Ok(XlsbEditor::Eager(package)),
        }
    }

    /// An owning copy of the package whose signature the leaf removed
    /// explicitly ([`xlsb::Package::unsign`]).
    fn unsigned_package(&self) -> Result<xlsb::Package> {
        Ok(match &self.0.owner {
            XlsbOwner::Source(workbook) => {
                let bytes = signature::exact_source(workbook.workbook().source_backed_package())?;
                xlsb::Package::from_bytes_with_options(
                    &bytes,
                    litchi_xlsb::OpenOptions::new().source_limits(self.0.ooxml_limits),
                )?
                .unsign()
            },
            XlsbOwner::Eager(package) => package.unsign(),
        })
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

impl std::fmt::Debug for Xlsb {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut debug = formatter.debug_struct("Xlsb");
        match &self.0.owner {
            XlsbOwner::Source(workbook) => debug.field("source", workbook),
            XlsbOwner::Eager(package) => debug.field(
                "worksheet_count",
                &projection(package).map_or(0, WorkbookTrait::worksheet_count),
            ),
        };
        debug.finish_non_exhaustive()
    }
}

impl WorkbookTrait for Xlsb {
    type Error = Error;
    type Worksheet<'a> = Sheet<'a>;
    type Worksheets<'a> = crate::view::sheet::Worksheets<'a>;

    fn active_sheet_index(&self) -> usize {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.active_worksheet_ordinal().unwrap_or(0),
            XlsbOwner::Eager(package) => {
                projection(package).map_or(0, WorkbookTrait::active_sheet_index)
            },
        }
    }

    fn active_worksheet(&self) -> Result<Sheet<'_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.active_worksheet().map(Sheet::xlsb_source),
            XlsbOwner::Eager(package) => projection(package)?
                .active_worksheet()
                .map(Sheet::xlsb_eager)
                .map_err(Into::into),
        }
    }

    fn is_1904_date_system(&self) -> bool {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.date1904(),
            XlsbOwner::Eager(package) => {
                projection(package).is_ok_and(WorkbookTrait::is_1904_date_system)
            },
        }
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Sheet<'_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.worksheet(index).map(Sheet::xlsb_source),
            XlsbOwner::Eager(package) => projection(package)?
                .worksheet_by_index(index)
                .map(Sheet::xlsb_eager)
                .map_err(Into::into),
        }
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Sheet<'_>> {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.worksheet_named(name).map(Sheet::xlsb_source),
            XlsbOwner::Eager(package) => projection(package)?
                .worksheet_by_name(name)
                .map(Sheet::xlsb_eager)
                .map_err(Into::into),
        }
    }

    fn worksheet_count(&self) -> usize {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.names().len(),
            XlsbOwner::Eager(package) => {
                projection(package).map_or(0, WorkbookTrait::worksheet_count)
            },
        }
    }

    fn worksheet_names(&self) -> &[String] {
        match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook.names(),
            XlsbOwner::Eager(package) => {
                projection(package).map_or(&[], WorkbookTrait::worksheet_names)
            },
        }
    }

    fn worksheets(&self) -> Self::Worksheets<'_> {
        crate::view::sheet::Worksheets::new(match &self.0.owner {
            XlsbOwner::Source(workbook) => {
                WorksheetsInner::XlsbSource(SourceWorksheets { workbook, index: 0 })
            },
            XlsbOwner::Eager(package) => match projection(package) {
                Ok(workbook) => WorksheetsInner::XlsbEager(workbook.worksheets()),
                Err(error) => WorksheetsInner::Refused(Some(error)),
            },
        })
    }
}

#[cfg(test)]
mod tests;
