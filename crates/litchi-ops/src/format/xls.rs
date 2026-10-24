//! Immutable legacy Excel handle backed by the native XLS Package.
//!
//! Detection hands its existing CFB parse or positional reader to the native
//! owner. Byte inputs retain eager BIFF coverage; filesystem inputs retain
//! lazy reads and source-change fences. The facade keeps signature disposition
//! and delegates semantic reads to the native Workbook projection.

use std::io::Cursor;
use std::mem::size_of;
use std::sync::{Arc, OnceLock};

use litchi_cfb::OleFile;
use litchi_core::sheet::WorkbookTrait;
use litchi_core::{FileFormat, Metadata};
use litchi_xls::{Package, SourceBackedWorkbook};

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::{OpenOptions, Result};

#[cfg(feature = "eval")]
pub(crate) mod calc;
mod children;
mod feature_write;
mod features;
mod layout;
mod metadata;
mod nav;
mod node;
pub(crate) mod owner;
mod spec;
mod structure;
mod style;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::{Nav, get};
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::{
    ADDABLE, CELL_WRITE_TABLE, REMOVABLE, SETTABLE, SPEC, WRITE_GAPS, capability_props,
};
pub(crate) use views::{Json as ViewJson, view_json, view_to};
pub(crate) use write::WriteCommit;

/// An immutable legacy Excel workbook. Cloning shares its native package.
#[derive(Clone)]
pub struct Xls(Arc<Opened<XlsState>>);

/// The state a handle shares with its clones: the native package and the
/// read caches of exactly that package (record 0315).
///
/// Every write publishes a new package into a new state
/// ([`Xls::with_package`]), so the caches never outlive the snapshot they
/// were filled from: a write invalidates them by construction.
pub(crate) struct XlsState {
    package: Package,
    /// The package's immutable tab catalog, retained only after the first
    /// checked read. Warm unified reads clone this `Arc` under their outer
    /// source fence instead of repeating the package metadata probe.
    sheet_catalog: OnceLock<Arc<[litchi_xls::SheetMetadata]>>,
    /// The sheet step of the last cell `get` ([`nav::SheetStep`]).
    sheet_step: nav::SheetStep,
}

impl XlsState {
    fn new(package: Package) -> Self {
        Self {
            package,
            sheet_catalog: OnceLock::new(),
            sheet_step: nav::SheetStep::default(),
        }
    }
}

type EagerWorkbook = litchi_xls::raw::Reader<Cursor<Vec<u8>>>;

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Xls>();
    assert!(size_of::<Xls>() == size_of::<usize>());
};

impl Xls {
    /// Whether two handles retain the same immutable owner snapshot.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Admit a primary package edit under signature and source guards.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        if self.is_encrypted() {
            return Err(litchi_xls::package::Error::EncryptionPolicy.into());
        }
        if self.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xls));
        }
        Ok(())
    }

    /// Admit a native edit that may still prove an exact signed no-op. A
    /// changed package is rejected by the publication fence.
    pub(crate) fn admit_edit_for_write(&self) -> Result<()> {
        self.check_source_for_write()?;
        if self.is_encrypted() {
            return Err(litchi_xls::package::Error::EncryptionPolicy.into());
        }
        Ok(())
    }

    /// Return the native package that may receive one immutable edit.
    ///
    /// An explicitly authorized signature removal starts from an unsigned
    /// copy of the exact source. Ordinary unsigned handles retain the native
    /// owner and therefore retain its source policy and lazy route.
    pub(crate) fn editable_package(&self) -> Result<Package> {
        self.admit_edit_for_write()?;
        if self.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                self.owner().source_bytes()?,
                litchi_sign::cfb::Format::Xls,
            )?;
            return Ok(litchi_xls::raw::package_from_replacement_bytes(
                self.owner(),
                bytes,
            )?);
        }
        Ok(self.owner().clone())
    }

    /// Admit an ordinary unsigned native edit against the remaining
    /// before/after snapshot budget before its package owner is cloned.
    pub(crate) fn editable_package_with_snapshot_limit(&self, maximum: usize) -> Result<Package> {
        self.preflight_snapshot_bytes(maximum, 0)?;
        self.editable_package()
    }

    /// Return the retained package for a setter that may prove an exact
    /// signed no-op. Removal-authorized handles use an explicit unsigned
    /// replacement; signed handles keep their original package until the
    /// publication fence sees whether the native patch changed.
    pub(crate) fn editable_package_for_write(&self) -> Result<Package> {
        self.admit_edit_for_write()?;
        if matches!(
            self.signature(),
            Signature::Signed | Signature::RemovalAuthorized
        ) {
            let bytes = signature::unsign_compound_file(
                self.owner().source_bytes()?,
                litchi_sign::cfb::Format::Xls,
            )?;
            return Ok(litchi_xls::raw::package_from_replacement_bytes(
                self.owner(),
                bytes,
            )?);
        }
        Ok(self.owner().clone())
    }

    /// Admit the native before/after package capture against the transaction's
    /// remaining snapshot budget before signature removal or package cloning
    /// can allocate its private edit owner.
    pub(crate) fn editable_package_for_write_with_snapshot_limit(
        &self,
        maximum: usize,
    ) -> Result<Package> {
        self.preflight_snapshot_bytes(maximum, 0)?;
        self.editable_package_for_write()
    }

    /// Check the native package, edit handoff and caller-authored staging
    /// against the remaining snapshot budget before a private XLS edit owner
    /// can be materialized. The patch's exact retained charge is still
    /// authoritative after commit.
    pub(crate) fn preflight_snapshot_bytes(&self, maximum: usize, additional: usize) -> Result<()> {
        let observed = self
            .retained_source_len()
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_xls::Patch>())
            .saturating_add(size_of::<litchi_xls::package::edit::Edit>())
            .saturating_add(additional);
        if observed > maximum {
            return Err(crate::transaction::Error::Limit {
                kind: crate::transaction::LimitKind::SnapshotBytes,
                observed,
                maximum,
            }
            .into());
        }
        Ok(())
    }

    /// Return the original native package for a canonical signed no-op patch.
    /// The ordinary write owner is unsigned before staging a changed edit;
    /// an empty cancellation patch must retain the signed CFB source instead.
    pub(crate) fn original_package_for_noop(&self) -> Result<Package> {
        self.admit_edit_for_write()?;
        Ok(self.owner().clone())
    }

    /// Carry one native package publication back into the immutable facade
    /// while retaining origin and the explicit signature disposition.
    pub(crate) fn with_package(&self, package: Package) -> Result<Self> {
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Ok(Self(
            Opened::new(XlsState::new(package))
                .with_origin(self.origin())
                .with_signature(signature),
        ))
    }

    /// The one-shot unified worksheet and scalar-cell setter.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
    ) -> Result<WriteCommit> {
        self.set_with_limits(
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Set scalar cell values, clear cells, or rename one worksheet under an
    /// explicit semantic budget.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let result = write::set_with_limits(self, target, props, limits);
        self.check_source_for_write()?;
        self.enforce_signed_noop(result)
    }

    fn enforce_signed_noop(&self, result: Result<WriteCommit>) -> Result<WriteCommit> {
        result.and_then(|commit| {
            if self.signature() != Signature::Signed || commit.after.same_snapshot(self) {
                Ok(commit)
            } else {
                Err(signature::refusal(FileFormat::Xls))
            }
        })
    }

    /// Remove stored cells selected by a path or scoped query.
    pub(crate) fn remove<'p>(&self, target: crate::target::TargetArg<'p>) -> Result<WriteCommit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Remove stored cells under an explicit semantic budget.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::target::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source()?;
        let result = write::remove_with_limits(self, target, limits);
        self.check_source()?;
        self.enforce_signed_noop(result)
    }

    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::create::{Conformance, Flavor};
        use crate::error::{Code, Error};
        let refusal = if options.flavor_value() != Flavor::Document {
            Some((
                "the XLS creator does not author templates",
                "use Flavor::Document",
            ))
        } else if options.conformance_value() == Conformance::Strict {
            Some((
                "legacy XLS does not support OOXML Strict conformance",
                "use Conformance::Transitional",
            ))
        } else {
            None
        };
        if let Some((message, suggestion)) = refusal {
            return Err(Error::new(Code::UnsupportedCapability, message)
                .with_format(FileFormat::Xls)
                .with_suggestion(suggestion));
        }
        match options.metadata_value() {
            Some(metadata) => Ok(Self(Opened::new(XlsState::new(metadata::create(
                metadata,
            )?)))),
            None => Ok(Self(Opened::new(XlsState::new(Package::new()?)))),
        }
    }

    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    pub(crate) fn origin(&self) -> Option<Arc<std::path::Path>> {
        self.0.origin()
    }

    pub(crate) fn is_encrypted(&self) -> bool {
        self.owner().is_encrypted()
    }

    pub(crate) fn check_source(&self) -> Result<()> {
        if let Some(source) = self.owner().source() {
            let expected = source.captured_source_version();
            let observed = source.source_arc().version()?;
            if observed != expected {
                return Err(
                    litchi_xls::SourceBackedError::SourceChanged { expected, observed }.into(),
                );
            }
        }
        Ok(())
    }

    /// Source freshness for a mutation that may still prove an exact signed
    /// no-op. Changed native results are rejected at publication.
    pub(crate) fn check_source_for_write(&self) -> Result<()> {
        self.check_source()
    }

    pub(crate) fn date_1904(&self) -> Result<bool> {
        let workbook = self.owner().workbook();
        if let Some(source) = workbook.source_backed_ref() {
            return Ok(matches!(
                source.date_system()?,
                litchi_xls::DateSystem::Excel1904
            ));
        }
        Ok(WorkbookTrait::is_1904_date_system(workbook.reader_ref()?))
    }

    pub(crate) fn write_to<W: std::io::Write>(&self, sink: &mut W) -> Result<()> {
        if self.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                self.owner().to_bytes()?,
                litchi_sign::cfb::Format::Xls,
            )?;
            sink.write_all(&bytes)?;
            self.check_source()?;
        } else {
            self.owner().write_to(sink)?;
        }
        Ok(())
    }

    pub(crate) fn retains_source(&self) -> bool {
        true
    }

    pub(crate) fn from_source(workbook: SourceBackedWorkbook) -> Result<Self> {
        Ok(Self(Opened::new(XlsState::new(Package::from_source(
            workbook,
        )?))))
    }

    pub(crate) fn from_ole_file(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: &OpenOptions,
        password: Option<&str>,
    ) -> Result<LegacyOpen<Self>> {
        let mut leaf_options = litchi_xls::OpenOptions::new().password(password);
        if let Some(maximum) = options.max_input_bytes {
            let maximum = maximum.min(litchi_xls::SourceLimits::MAX_INPUT_BYTES as u64);
            let maximum = usize::try_from(maximum).unwrap_or(usize::MAX);
            let limits = litchi_xls::SourceLimits::builder()
                .max_input_bytes(maximum)?
                .build()?;
            leaf_options = leaf_options.source_limits(limits);
        }
        match Package::from_ole_file_or_password_required(ole, leaf_options)? {
            Ok(package) => Ok(LegacyOpen::Opened(Self(Opened::new(XlsState::new(
                package,
            ))))),
            Err(ole) => Ok(LegacyOpen::PasswordRequired(Box::new(ole))),
        }
    }

    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    pub(super) fn owner(&self) -> &Package {
        &self.0.package
    }

    /// The retained compound-file length, when the native package exposes
    /// one. This O(1) lower bound is used before a private edit owner is
    /// materialized; native patch accounting remains the hard post-commit
    /// limit check.
    pub(crate) fn retained_source_len(&self) -> usize {
        self.owner()
            .source()
            .map(|source| usize::try_from(source.file_size()).unwrap_or(usize::MAX))
            .unwrap_or(0)
    }

    /// The handle's sheet-step cache (record 0315).
    pub(super) fn sheet_step(&self) -> &nav::SheetStep {
        &self.0.sheet_step
    }

    /// Return the adapter-local catalog after the package has captured it.
    /// This is used only inside an active read fence; callers outside that
    /// scope must use the checked package projection.
    pub(super) fn cached_sheet_catalog(&self) -> Option<Arc<[litchi_xls::SheetMetadata]>> {
        self.0.sheet_catalog.get().cloned()
    }

    /// Capture and retain the package catalog on the first checked read.
    pub(super) fn sheet_catalog(&self) -> Result<Arc<[litchi_xls::SheetMetadata]>> {
        let catalog = self.owner().sheet_catalog()?;
        let _ = self.0.sheet_catalog.set(Arc::clone(&catalog));
        Ok(self.cached_sheet_catalog().unwrap_or(catalog))
    }

    /// Run one read against the retained native owner, fencing both the
    /// source-backed and eager routes.  The native source owner itself keeps
    /// the payload scan lazy; this wrapper checks the captured source version
    /// once, after the operation.  Source versions only advance, so one
    /// probe after the read detects a change made before or during it, and a
    /// freshness failure takes precedence over the operation's own result.
    ///
    /// The operation runs in a fence scope for this handle, so the reads it
    /// nests (a `View::text` that formats through further fenced reads)
    /// share this one check instead of each adding their own; inside an
    /// enclosing scope (a resolver or query step, or an outer `finish`),
    /// whose check runs after the whole step, the operation runs alone
    /// (record 0291). Every public read still ends in at least one check.
    pub(crate) fn finish<T>(&self, operation: impl FnOnce() -> Result<T>) -> Result<T> {
        let key = self.fence_key();
        if crate::path::fence::is_active(key) {
            return operation();
        }
        let value = {
            let _scope = crate::path::fence::Scope::enter(key);
            operation()
        };
        match self.check_source() {
            Err(error) => Err(error),
            Ok(()) => value,
        }
    }

    /// This handle's key for [`crate::path::fence::Scope`]: its shared
    /// owner, which every clone of the handle reads.
    pub(crate) fn fence_key(&self) -> usize {
        Arc::as_ptr(&self.0).addr()
    }

    #[cfg(test)]
    pub(crate) fn is_source_backed(&self) -> bool {
        self.owner().is_source_backed()
    }

    /// Return ordinary worksheet names in tab order.
    ///
    /// # Errors
    /// Returns the native source or semantic read error.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        Ok(self.owner().workbook().worksheet_names()?)
    }

    /// Return the number of ordinary worksheets.
    ///
    /// # Errors
    /// Returns the native source or semantic read error.
    pub fn worksheet_count(&self) -> Result<usize> {
        Ok(self.owner().workbook().worksheet_count()?)
    }

    /// Extract worksheet text through the retained native read route.
    ///
    /// # Errors
    /// Returns the native source or semantic read error.
    pub fn text(&self) -> Result<String> {
        Ok(self.owner().workbook().text()?)
    }

    /// Read package metadata under the native source consistency fence.
    ///
    /// # Errors
    /// Returns a typed source or metadata read error.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(Some(self.owner().metadata()?))
    }

    /// Return the native report of accepted reader repairs.
    ///
    /// # Errors
    /// Returns a typed source-change error.
    pub fn tolerance_report(&self) -> Result<litchi_xls::ToleranceReport> {
        Ok(self.owner().workbook().tolerance_report()?)
    }

    /// The retained editable package after signature admission (unsigned
    /// first when removal is authorized): its `cell_values()` read view and
    /// `edit()` stage the native cell-value engine.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn cell_values_editor(&self) -> Result<Package> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xls));
        }
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = self.owner().source_bytes()?;
            let bytes = signature::unsign_compound_file(bytes, litchi_sign::cfb::Format::Xls)?;
            return Ok(Package::from_bytes(&bytes)?);
        }
        Ok(self.owner().clone())
    }

    fn trait_workbook(&self) -> Result<&EagerWorkbook> {
        Ok(self.owner().reader_ref()?)
    }
}

/// The worksheet the eager reader lends.
pub(crate) type XlsSheet<'a> = &'a litchi_xls::Worksheet;

use crate::view::sheet::{Worksheet as Sheet, WorksheetsInner};

/// The generic sheet traits over the eager reader.
///
/// On the filesystem route the first trait call reads the whole workbook
/// once. When that read fails, the fallible methods return its error and the
/// infallible ones report an empty workbook.
impl WorkbookTrait for Xls {
    type Error = crate::Error;
    type Worksheet<'a> = Sheet<'a>;
    type Worksheets<'a> = crate::view::sheet::Worksheets<'a>;

    fn active_sheet_index(&self) -> usize {
        self.trait_workbook()
            .map_or(0, WorkbookTrait::active_sheet_index)
    }

    fn active_worksheet(&self) -> Result<Sheet<'_>> {
        self.trait_workbook()?
            .active_worksheet()
            .map(Sheet::xls)
            .map_err(Into::into)
    }

    fn is_1904_date_system(&self) -> bool {
        self.trait_workbook()
            .is_ok_and(WorkbookTrait::is_1904_date_system)
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Sheet<'_>> {
        self.trait_workbook()?
            .worksheet_by_index(index)
            .map(Sheet::xls)
            .map_err(Into::into)
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Sheet<'_>> {
        self.trait_workbook()?
            .worksheet_by_name(name)
            .map(Sheet::xls)
            .map_err(Into::into)
    }

    fn worksheet_count(&self) -> usize {
        self.trait_workbook()
            .map_or(0, WorkbookTrait::worksheet_count)
    }

    fn worksheet_names(&self) -> &[String] {
        self.trait_workbook()
            .map_or(&[], WorkbookTrait::worksheet_names)
    }

    fn worksheets(&self) -> Self::Worksheets<'_> {
        crate::view::sheet::Worksheets::new(WorksheetsInner::Xls(match self.trait_workbook() {
            Ok(workbook) => Worksheets::Read(workbook.worksheets()),
            Err(error) => Worksheets::Failed(Some(error)),
        }))
    }
}

impl std::fmt::Debug for Xls {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Xls")
            .field("source_backed", &self.owner().is_source_backed())
            .finish_non_exhaustive()
    }
}

/// The worksheets of an XLS handle through the generic traits: the eager
/// reader's, or one error, the failed materialization.
pub(crate) enum Worksheets<'a> {
    Read(litchi_xls::workbook::Worksheets<'a>),
    Failed(Option<crate::Error>),
}

impl<'a> Iterator for Worksheets<'a> {
    type Item = Result<XlsSheet<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Read(worksheets) => worksheets.next().map(|sheet| sheet.map_err(Into::into)),
            Self::Failed(error) => error.take().map(Err),
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::Code;
    use crate::detect::{DetectedFormat, detect_format_smart};
    use litchi_cfb::SharedOleFile;
    use litchi_core::sheet::{CellValue, Worksheet as _};
    use litchi_ole_common::property_set::PropertySetReader;
    use litchi_xls::SourceBackedError;
    use std::io::Write;
    use std::path::{Path, PathBuf};

    #[test]
    fn explicit_eager_input_profile_can_exceed_default_within_hard_cap() {
        let maximum = 768 * 1024 * 1024;
        let options = OpenOptions {
            max_input_bytes: Some(maximum),
            ..OpenOptions::default()
        };
        let ole = OleFile::open(
            Cursor::new(fixture("Simple.xls")),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let LegacyOpen::Opened(xls) = Xls::from_ole_file(ole, &options, None).unwrap() else {
            panic!("unencrypted fixture");
        };
        assert_eq!(
            xls.owner().source_limits().unwrap().max_input_bytes() as u64,
            maximum
        );
    }

    #[test]
    #[cfg(any(unix, windows))]
    fn explicit_positional_input_profile_is_retained_by_cfb_and_workbook() {
        let maximum = 3 * 1024 * 1024 * 1024;
        let options = OpenOptions::new().max_input_bytes(maximum);
        let crate::File::Xls(xls) = crate::open_with(fixture_path("Simple.xls"), &options).unwrap()
        else {
            panic!("unencrypted XLS fixture");
        };
        assert_eq!(
            xls.owner().source().unwrap().limits().max_input_bytes(),
            maximum
        );
        assert_eq!(
            xls.owner().source_backed_limits().unwrap().max_input_bytes,
            maximum
        );
    }

    fn append_cell_text(out: &mut String, cell: &CellValue) {
        match cell {
            CellValue::Empty => {},
            CellValue::Bool(value) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
            CellValue::Int(value) => out.push_str(&value.to_string()),
            CellValue::Float(value) => out.push_str(&value.to_string()),
            CellValue::String(value) | CellValue::Error(value) => out.push_str(value),
            CellValue::DateTime(value) => out.push_str(&value.to_string()),
            CellValue::Formula {
                formula,
                cached_value,
                ..
            } => match cached_value.as_deref() {
                Some(CellValue::String(value)) => out.push_str(value),
                Some(CellValue::Int(value)) => out.push_str(&value.to_string()),
                Some(CellValue::Float(value)) => out.push_str(&value.to_string()),
                Some(CellValue::Bool(value)) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
                _ => {
                    out.push('=');
                    out.push_str(formula);
                },
            },
        }
    }

    fn fixture_path(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls")
            .join(name)
    }

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(fixture_path(name)).expect("XLS fixture")
    }

    /// The facade's byte route: detect, then hand the parse to the handle.
    /// A password-protected workbook reports the leaf's refusal, as the
    /// facade did; the unified opener turns it into `File::Locked`.
    fn open_bytes(bytes: Vec<u8>) -> Result<Xls> {
        match detect_format_smart(bytes, &crate::detect::SourceLimits::default()) {
            Some(DetectedFormat::Xls(ole)) => {
                match Xls::from_ole_file(ole, &OpenOptions::default(), None)? {
                    LegacyOpen::Opened(xls) => Ok(xls),
                    LegacyOpen::PasswordRequired(_) => {
                        Err(litchi_xls::Error::PasswordRequired.into())
                    },
                }
            },
            other => panic!("detected {other:?}, not XLS"),
        }
    }

    /// The filesystem route: [`crate::open`] opens the positional workbook.
    #[cfg(any(unix, windows))]
    fn open_path(path: &Path) -> Xls {
        match crate::open(path).unwrap() {
            crate::File::Xls(xls) => xls,
            other => panic!("{} was not opened as XLS: {other:?}", path.display()),
        }
    }

    /// The facade's eager text projection over the leaf `WorkbookTrait`, as
    /// an independent oracle for the handle's inherent-API port.
    fn trait_text(name: &str) -> String {
        let workbook = litchi_xls::raw::Reader::new(Cursor::new(fixture(name))).unwrap();
        let mut out = String::new();
        for index in 0..workbook.worksheet_count() {
            let worksheet = workbook.worksheet_by_index(index).unwrap();
            for row in worksheet.rows() {
                for (column, cell) in row.unwrap().iter().enumerate() {
                    if column > 0 {
                        out.push('\t');
                    }
                    append_cell_text(&mut out, cell);
                }
                out.push('\n');
            }
        }
        out
    }

    /// Every XLS in `test-data/ole/xls`.
    fn corpus() -> Vec<String> {
        let mut names = std::fs::read_dir(fixture_path(""))
            .unwrap()
            .map(|entry| entry.unwrap().file_name().into_string().unwrap())
            .filter(|name| name.ends_with(".xls"))
            .collect::<Vec<_>>();
        names.sort();
        assert!(names.len() > 40, "XLS corpus went missing");
        names
    }

    /// The same result, or the same typed failure.
    fn assert_same<T: PartialEq + std::fmt::Debug, E: std::fmt::Display>(
        handle: Result<T>,
        leaf: std::result::Result<T, E>,
        what: &str,
    ) {
        match (handle, leaf) {
            (Ok(handle), Ok(leaf)) => assert_eq!(handle, leaf, "{what}"),
            (Err(handle), Err(leaf)) => assert_eq!(handle.message(), leaf.to_string(), "{what}"),
            (handle, leaf) => panic!(
                "{what}: handle {:?} but leaf {:?}",
                handle.map_err(|error| error.message().to_owned()),
                leaf.map_err(|error| error.to_string())
            ),
        }
    }

    /// The facade's byte route against the eager leaf, over the corpus.
    #[test]
    fn byte_route_matches_the_eager_leaf_reader() {
        let mut opened = 0_usize;
        for name in corpus() {
            let Ok(mut leaf) = OleFile::open(
                Cursor::new(fixture(&name)),
                litchi_cfb::OleFileLimits::default(),
            ) else {
                continue;
            };
            let metadata = leaf.get_metadata().map(Metadata::from).unwrap_or_default();
            let workbook = litchi_xls::raw::Reader::from_ole_file(leaf);
            let xls = match (open_bytes(fixture(&name)), workbook) {
                (Ok(xls), Ok(workbook)) => {
                    assert_eq!(
                        xls.worksheet_names().unwrap(),
                        workbook.worksheet_names().to_vec(),
                        "{name}"
                    );
                    assert_eq!(xls.worksheet_count().unwrap(), workbook.worksheet_count());
                    assert_eq!(
                        &xls.tolerance_report().unwrap(),
                        workbook.tolerance_report(),
                        "{name}"
                    );
                    xls
                },
                (Err(error), Err(leaf)) => {
                    assert_eq!(error.message(), leaf.to_string(), "{name}");
                    continue;
                },
                (handle, leaf) => panic!(
                    "{name}: handle open {} but leaf open {}",
                    handle.is_ok(),
                    leaf.is_ok()
                ),
            };
            opened += 1;
            assert!(!xls.is_source_backed());
            assert_eq!(xls.text().unwrap(), trait_text(&name), "{name}");
            assert_eq!(
                format!("{:?}", xls.metadata().unwrap()),
                format!("{:?}", Some(metadata)),
                "{name}"
            );
        }
        assert!(opened > 30, "too few XLS fixtures opened: {opened}");
    }

    /// Fixtures whose names and text the facade's two XLS routes already read
    /// differently: the eager reader omits worksheets the source-backed owner
    /// reads. The handle keeps each route's reader, so it keeps these
    /// differences.
    const ROUTE_DIVERGENCES: &[&str] = &[
        "SimpleChart.xls",
        "SimpleScatterChart.xls",
        "SimpleWithColours copy.xls",
        "SimpleWithColours.xls",
        "SingleLetterRanges.xls",
        "WithChart.xls",
        "WithFormattedGraphTitle.xls",
        "WithThreeCharts.xls",
    ];

    /// The facade's filesystem route against the source-backed leaf, and
    /// against the byte route wherever both read, over the corpus.
    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_route_matches_the_source_leaf_and_the_byte_route() {
        let mut compared = 0_usize;
        let mut diverged = Vec::new();
        for name in corpus() {
            let path = fixture_path(&name);
            // The workbooks the positional owner reads; the path route reads
            // the others (BIFF5 and older, encrypted) eagerly or refuses them.
            let Ok(crate::File::Xls(xls)) = crate::open(&path) else {
                continue;
            };
            if !xls.is_source_backed() {
                continue;
            }
            let owner = SourceBackedWorkbook::open(&path).unwrap();
            assert_same(xls.text(), owner.text(), &name);
            assert_same(xls.worksheet_names(), owner.worksheet_names(), &name);
            assert_same(xls.worksheet_count(), owner.worksheet_count(), &name);
            let source = xls
                .metadata()
                .unwrap()
                .expect("XLS metadata is always reported");
            // Read twice: the second read comes from the kept projection.
            assert_eq!(
                format!("{:?}", xls.metadata().unwrap()),
                format!("{:?}", Some(&source))
            );
            assert!(xls.tolerance_report().is_ok());

            // The facade's own parity test: where both routes read, they
            // agree.
            let (Ok(text), Ok(bytes)) = (xls.text(), open_bytes(fixture(&name))) else {
                continue;
            };
            let eager = bytes
                .metadata()
                .unwrap()
                .expect("XLS metadata is always reported");
            assert_eq!(source.application, eager.application, "{name}");
            assert_eq!(source.author, eager.author, "{name}");
            assert_eq!(source.title, eager.title, "{name}");
            if ROUTE_DIVERGENCES.contains(&name.as_str()) {
                diverged.push(name);
                continue;
            }
            compared += 1;
            assert_eq!(
                xls.worksheet_names().unwrap(),
                bytes.worksheet_names().unwrap(),
                "{name}"
            );
            assert_eq!(text, bytes.text().unwrap(), "{name}");
        }
        assert_eq!(
            diverged, ROUTE_DIVERGENCES,
            "the known route divergences changed"
        );
        assert!(compared > 20, "too few XLS fixtures compared: {compared}");
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_route_refuses_a_changed_source() {
        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), fixture("Simple.xls")).unwrap();
        let xls = open_path(file.path());
        let mut appended = std::fs::OpenOptions::new()
            .append(true)
            .open(file.path())
            .unwrap();
        appended.write_all(&[0]).unwrap();
        appended.flush().unwrap();

        for error in [
            xls.text().unwrap_err(),
            xls.metadata().unwrap_err(),
            xls.tolerance_report().unwrap_err(),
            xls.worksheet_names().unwrap_err(),
        ] {
            assert_eq!(error.code(), Code::IoError);
            assert!(matches!(
                error.leaf(),
                crate::error::Source::XlsSource(SourceBackedError::SourceChanged { .. })
            ));
        }
    }

    /// An in-memory source that counts its version observations.
    struct Counting {
        inner: litchi_core::OwnedSource,
        versions: std::sync::atomic::AtomicUsize,
    }

    impl litchi_core::ReadAt for Counting {
        fn len(&self) -> std::io::Result<u64> {
            self.inner.len()
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.inner.read_at(offset, output)
        }

        fn version(&self) -> std::io::Result<litchi_core::SourceVersion> {
            self.versions
                .fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            self.inner.version()
        }
    }

    /// A source that counts its version observations and can be changed
    /// (record 0315).
    struct Changing {
        inner: litchi_core::OwnedSource,
        revision: std::sync::atomic::AtomicU64,
        versions: std::sync::atomic::AtomicUsize,
        bump_before_version: std::sync::atomic::AtomicUsize,
    }

    impl Changing {
        fn bump_before_next_version(&self) {
            let next = self
                .versions
                .load(std::sync::atomic::Ordering::SeqCst)
                .saturating_add(1);
            self.bump_before_version
                .store(next, std::sync::atomic::Ordering::SeqCst);
        }
    }

    impl litchi_core::ReadAt for Changing {
        fn len(&self) -> std::io::Result<u64> {
            self.inner.len()
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.inner.read_at(offset, output)
        }

        fn version(&self) -> std::io::Result<litchi_core::SourceVersion> {
            use std::sync::atomic::Ordering;
            let ordinal = self.versions.fetch_add(1, Ordering::SeqCst) + 1;
            if self.bump_before_version.load(Ordering::SeqCst) == ordinal {
                self.bump_before_version.store(0, Ordering::SeqCst);
                self.revision.fetch_add(1, Ordering::SeqCst);
            }
            Ok(litchi_core::SourceVersion::new(
                0x584c_5331_3533_3135,
                self.revision.load(Ordering::SeqCst),
            ))
        }
    }

    /// What a `get` of `path` gives: the node's path, kind and text, or
    /// the error's code and message.
    fn outcome(result: Result<crate::node::Node<'_>>) -> String {
        match result {
            Ok(node) => format!(
                "ok {} {} {:?}",
                node.path(),
                node.kind(),
                node.view()
                    .text()
                    .map(|text| text.map(|text| text.into_owned()))
            ),
            Err(error) => format!("err {:?} {error}", error.code()),
        }
    }

    /// The cell paths the sheet-step cache tests resolve on `xls`: every
    /// sheet step form for each worksheet, a few cells each, and misses.
    fn cached_get_paths(xls: &Xls) -> Vec<String> {
        let names = xls.worksheet_names().unwrap();
        let mut steps = vec!["sheet".to_owned(), "sheet[last()]".to_owned()];
        for (index, name) in names.iter().enumerate() {
            steps.push(format!("sheet[{}]", index + 1));
            steps.push(format!("\"{name}\""));
            steps.push(format!("\"{}\"", name.to_ascii_uppercase()));
        }
        steps.push(format!("sheet[{}]", names.len() + 1));
        steps.push("NoSuchSheet".to_owned());
        let mut paths = Vec::new();
        for step in &steps {
            for cell in ["A1", "B2", "a1", "C3", "IV65536", "IW1", "A65537"] {
                paths.push(format!("/{step}/{cell}"));
            }
        }
        paths
    }

    /// Record 0315 (record 0310 G8): a cell `get` whose sheet step the
    /// handle cached gives exactly what the general resolver gives on a
    /// handle that never cached a step: the same node, text and error, on
    /// both owners, whatever step the cache holds.
    #[test]
    fn cached_sheet_steps_resolve_as_the_resolver_does() {
        for name in ["Simple.xls", "HyperlinksOnManySheets.xls"] {
            let bytes = fixture(name);
            let eager = crate::File::Xls(open_bytes(bytes.clone()).unwrap());
            let source = Arc::new(litchi_core::OwnedSource::new(bytes.clone()));
            let positional = crate::File::Xls(
                Xls::from_source(SourceBackedWorkbook::from_read_at(source).unwrap()).unwrap(),
            );
            for file in [&eager, &positional] {
                let crate::File::Xls(xls) = file else {
                    unreachable!("an XLS handle")
                };
                let paths = cached_get_paths(xls);
                for path in &paths {
                    let arg = crate::path::IntoPath::into_path(path.as_str()).unwrap();
                    let general = outcome(
                        crate::path::resolve::get(&Nav::new(xls), &arg).map(crate::node::Node::Xls),
                    );
                    // Uncached, then cached on the same step, then after the
                    // cache moved to another step.
                    assert_eq!(outcome(file.get(path.as_str())), general, "{name} {path}");
                    assert_eq!(outcome(file.get(path.as_str())), general, "{name} {path}");
                }
                for path in paths.iter().rev() {
                    let arg = crate::path::IntoPath::into_path(path.as_str()).unwrap();
                    let general = outcome(
                        crate::path::resolve::get(&Nav::new(xls), &arg).map(crate::node::Node::Xls),
                    );
                    assert_eq!(outcome(file.get(path.as_str())), general, "{name} {path}");
                }
            }
        }
    }

    /// Record 0315: the cached route reads the source exactly as the
    /// resolver does (the same observations per `get`), a source changed
    /// between two cached `get`s fails the second, and a write publishes a
    /// handle whose cache starts empty, so a renamed worksheet is no longer
    /// found under its old name.
    #[test]
    fn cached_sheet_steps_keep_the_resolver_checks() {
        use std::sync::atomic::Ordering;
        let changing = Arc::new(Changing {
            inner: litchi_core::OwnedSource::new(fixture("Simple.xls")),
            revision: std::sync::atomic::AtomicU64::new(0),
            versions: std::sync::atomic::AtomicUsize::new(0),
            bump_before_version: std::sync::atomic::AtomicUsize::new(0),
        });
        let workbook = SourceBackedWorkbook::from_read_at(changing.clone()).unwrap();
        let file = crate::File::Xls(Xls::from_source(workbook).unwrap());
        let crate::File::Xls(xls) = &file else {
            unreachable!("an XLS handle")
        };
        let observe = |read: &dyn Fn()| {
            let before = changing.versions.load(Ordering::SeqCst);
            read();
            changing.versions.load(Ordering::SeqCst) - before
        };
        let resolver = || {
            let arg = crate::path::IntoPath::into_path("/sheet[1]/A1").unwrap();
            drop(crate::path::resolve::get(&Nav::new(xls), &arg).unwrap());
        };
        // The first reads warm the leaf's own caches (its cell index).
        for _ in 0..3 {
            resolver();
        }
        let general = observe(&resolver);
        assert_eq!(
            observe(&resolver),
            general,
            "a warm resolver read is stable"
        );
        // The first `get` fills the cache; the next ones read through it.
        let first = observe(&|| drop(file.get("/sheet[1]/A1").unwrap()));
        let cached = observe(&|| drop(file.get("/sheet[1]/A1").unwrap()));
        assert_eq!(first, general);
        assert_eq!(
            cached, general,
            "a cached get observes what the resolver does"
        );
        changing.revision.store(1, Ordering::SeqCst);
        let error = file.get("/sheet[1]/A1").unwrap_err();
        assert_eq!(error.code(), Code::IoError);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::XlsSource(SourceBackedError::SourceChanged { .. })
        ));

        let eager = crate::File::Xls(open_bytes(fixture("Simple.xls")).unwrap());
        let name = match &eager {
            crate::File::Xls(xls) => xls.worksheet_names().unwrap()[0].clone(),
            _ => unreachable!("an XLS handle"),
        };
        let by_name = format!("/\"{name}\"/A1");
        let before = outcome(eager.get(by_name.as_str()));
        assert!(before.starts_with("ok "), "{before}");
        let renamed = eager
            .set("/sheet[1]", crate::props! { "name" => "Renamed" })
            .unwrap();
        let after = outcome(renamed.get(by_name.as_str()));
        assert!(after.starts_with("err "), "{after}");
        assert!(outcome(renamed.get("/Renamed/A1")).starts_with("ok "));
        // The old handle keeps its snapshot and its cache.
        assert_eq!(outcome(eager.get(by_name.as_str())), before);
    }

    /// A stale source wins both when the first catalog read is fenced and
    /// when the source changes between resolution and the required trailing
    /// fence.  The latter also keeps a path miss from masking SourceChanged.
    #[test]
    fn source_change_precedes_cold_catalog_and_mid_resolve_results() {
        use std::sync::atomic::Ordering;

        let cold_source = Arc::new(Changing {
            inner: litchi_core::OwnedSource::new(fixture("Simple.xls")),
            revision: std::sync::atomic::AtomicU64::new(0),
            versions: std::sync::atomic::AtomicUsize::new(0),
            bump_before_version: std::sync::atomic::AtomicUsize::new(0),
        });
        let cold_workbook = SourceBackedWorkbook::from_read_at(cold_source.clone()).unwrap();
        let cold = crate::File::Xls(Xls::from_source(cold_workbook).unwrap());
        cold_source.revision.store(1, Ordering::SeqCst);
        let cold_error = cold.get("/sheet[999]").unwrap_err();
        assert_eq!(cold_error.code(), Code::IoError);
        assert!(matches!(
            cold_error.leaf(),
            crate::error::Source::XlsSource(SourceBackedError::SourceChanged { .. })
        ));

        let mid_source = Arc::new(Changing {
            inner: litchi_core::OwnedSource::new(fixture("Simple.xls")),
            revision: std::sync::atomic::AtomicU64::new(0),
            versions: std::sync::atomic::AtomicUsize::new(0),
            bump_before_version: std::sync::atomic::AtomicUsize::new(0),
        });
        let mid_workbook = SourceBackedWorkbook::from_read_at(mid_source.clone()).unwrap();
        let mid = crate::File::Xls(Xls::from_source(mid_workbook).unwrap());
        assert!(mid.get("/sheet[1]").is_ok(), "warm the adapter catalog");
        mid_source.bump_before_next_version();
        let mid_error = mid.get("/sheet[999]").unwrap_err();
        assert_eq!(mid_error.code(), Code::IoError);
        assert!(matches!(
            mid_error.leaf(),
            crate::error::Source::XlsSource(SourceBackedError::SourceChanged { .. })
        ));
    }

    /// Record 0291: a read nested in another fenced read of the same handle
    /// (a sheet's `View::text`, which reads the sheet name through a
    /// fenced reader) shares the outer read's one check; every public read
    /// still checks once, and every query match once.
    #[test]
    fn nested_fenced_reads_share_the_outer_check() {
        let counting = Arc::new(Counting {
            inner: litchi_core::OwnedSource::new(fixture("Simple.xls")),
            versions: std::sync::atomic::AtomicUsize::new(0),
        });
        let workbook = SourceBackedWorkbook::from_read_at(counting.clone()).unwrap();
        let unified = crate::File::Xls(Xls::from_source(workbook).unwrap());
        let observe = |read: &dyn Fn()| {
            let before = counting.versions.load(std::sync::atomic::Ordering::Relaxed);
            read();
            counting.versions.load(std::sync::atomic::Ordering::Relaxed) - before
        };
        let sheet = unified.get("/sheet[1]").unwrap();
        // `View::text` of a sheet reads its name through `sheet_name`, itself
        // a fenced read: one check for both.
        assert_eq!(observe(&|| drop(sheet.view().text().unwrap())), 1);
        assert_eq!(observe(&|| drop(sheet.view().text().unwrap())), 1);
        // A `get` checks once after resolving.
        assert_eq!(observe(&|| drop(unified.get("/sheet[1]").unwrap())), 1);
        // Each match of a query, and its end, is checked.
        // Each match of a query, and its end, is still checked (owner
        // decision 2026-10-14), besides what the leaf observes.
        let sheets = unified.query("sheet").unwrap().count();
        assert!(sheets >= 1);
        assert!(observe(&|| assert_eq!(unified.query("sheet").unwrap().count(), sheets)) > sheets);
    }

    /// Record 0291: nested reads under `View::text` share one check, and
    /// a node read before a change still refuses the changed source.
    #[cfg(any(unix, windows))]
    #[test]
    fn shared_nested_checks_still_refuse_a_changed_source() {
        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), fixture("Simple.xls")).unwrap();
        let unified = crate::open(file.path()).unwrap();
        let node = unified.get("/sheet[1]/A1").unwrap();
        let before = node.view().text().unwrap();
        assert!(before.is_some());
        assert_eq!(
            unified.get("/sheet[1]/A1").unwrap().view().text().unwrap(),
            before
        );
        let mut appended = std::fs::OpenOptions::new()
            .append(true)
            .open(file.path())
            .unwrap();
        appended.write_all(&[0]).unwrap();
        appended.flush().unwrap();
        for error in [
            node.view().text().unwrap_err(),
            unified.get("/sheet[1]/A1").unwrap_err(),
        ] {
            assert!(
                matches!(
                    error.leaf(),
                    crate::error::Source::XlsSource(SourceBackedError::SourceChanged { .. })
                ),
                "{error:?}"
            );
        }
        assert!(!crate::path::fence::is_active(match &unified {
            crate::File::Xls(xls) => xls.fence_key(),
            _ => unreachable!(),
        }));
    }

    #[test]
    fn byte_route_keeps_legacy_biff_the_source_owner_refuses() {
        // A BIFF5 workbook: the eager reader reads it, the source owner does
        // not. The byte route must keep reading it.
        let source = Arc::new(litchi_core::OwnedSource::new(fixture("Simple.xls")));
        let mut stream = SharedOleFile::open(source, litchi_cfb::SharedOleFileLimits::default())
            .unwrap()
            .open_stream(&["Workbook"])
            .unwrap();
        stream[4..6].copy_from_slice(&0x0500_u16.to_le_bytes());
        let mut writer = litchi_cfb::OleWriter::new();
        writer.create_stream(&["Workbook"], &stream).unwrap();
        let mut bytes = Vec::new();
        writer.write_to(&mut Cursor::new(&mut bytes)).unwrap();

        let xls = open_bytes(bytes.clone()).unwrap();
        assert!(!xls.worksheet_names().unwrap().is_empty());
        assert!(matches!(
            SourceBackedWorkbook::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes))),
            Err(SourceBackedError::UnsupportedBiffVersion(0x0500))
        ));
    }

    #[test]
    fn encrypted_input_keeps_todays_typed_refusal() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet/xor-encryption-abc.xls");
        let Ok(bytes) = std::fs::read(path) else {
            return;
        };
        let error = match open_bytes(bytes) {
            Ok(_) => panic!("an encrypted XLS opened without a password"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::DocumentProtected);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Xls(litchi_xls::Error::PasswordRequired)
        ));
    }

    #[test]
    fn editors_are_built_from_the_retained_source() {
        let bytes = fixture("Simple.xls");
        let oracle = Package::from_bytes(&bytes).unwrap().cell_values().unwrap();

        let eager = open_bytes(bytes).unwrap();
        let editor = eager.cell_values_editor().unwrap().cell_values().unwrap();
        assert_eq!(editor.bytes(), oracle.bytes());
        assert_eq!(editor.worksheet_count(), oracle.worksheet_count());

        #[cfg(any(unix, windows))]
        {
            let source = open_path(&fixture_path("Simple.xls"));
            let editor = source.cell_values_editor().unwrap().cell_values().unwrap();
            assert_eq!(editor.bytes(), oracle.bytes());
            assert_eq!(editor.worksheet_count(), oracle.worksheet_count());
        }
    }

    #[test]
    fn handles_are_cheap_shared_clones() {
        let xls = open_bytes(fixture("Simple.xls")).unwrap();
        let clone = xls.clone();
        assert!(Arc::ptr_eq(&xls.0, &clone.0));
        std::thread::scope(|scope| {
            scope.spawn(|| assert_eq!(clone.text().unwrap(), xls.text().unwrap()));
        });
    }
}
