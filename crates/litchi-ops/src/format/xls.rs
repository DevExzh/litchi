//! Immutable legacy Excel handle backed by the native XLS Package.
//!
//! Detection hands its existing CFB parse or positional reader to the native
//! owner. Byte inputs retain eager BIFF coverage; filesystem inputs retain
//! lazy reads and source-change fences. The facade keeps signature disposition
//! and delegates semantic reads to the native Workbook projection.

use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{OleFile, OleFileLimits};
use litchi_core::sheet::WorkbookTrait;
use litchi_core::{FileFormat, Metadata};
use litchi_xls::{Package, SourceBackedWorkbook, cell_values};

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::{OpenOptions, Result};

mod children;
mod nav;
mod node;
pub(crate) mod owner;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::{CELL_CAPABILITY_PROPS, SHEET_CAPABILITY_PROPS, SPEC};
pub(crate) use views::{Json as ViewJson, view_json, view_to};
pub(crate) use write::WriteCommit;

/// An immutable legacy Excel workbook. Cloning shares its native package.
#[derive(Clone)]
pub struct Xls(Arc<Opened<Package>>);

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

    /// Return the native package that may receive one immutable edit.
    ///
    /// An explicitly authorized signature removal starts from an unsigned
    /// copy of the exact source. Ordinary unsigned handles retain the native
    /// owner and therefore retain its source policy and lazy route.
    pub(crate) fn editable_package(&self) -> Result<Package> {
        self.admit_edit()?;
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

    /// Carry one native package publication back into the immutable facade
    /// while retaining origin and the explicit signature disposition.
    pub(crate) fn with_package(&self, package: Package) -> Result<Self> {
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Ok(Self(
            Opened::new(package)
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
        self.check_source()?;
        let result = write::set_with_limits(self, target, props, limits);
        self.check_source()?;
        result
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
        } else if options.metadata_value().is_some() {
            Some((
                "initial XLS metadata authoring is not integrated",
                "omit initial metadata",
            ))
        } else {
            None
        };
        if let Some((message, suggestion)) = refusal {
            return Err(Error::new(Code::UnsupportedCapability, message)
                .with_format(FileFormat::Xls)
                .with_suggestion(suggestion));
        }
        Ok(Self(Opened::new(Package::new()?)))
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
        Ok(Self(Opened::new(Package::from_source(workbook)?)))
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
            Ok(package) => Ok(LegacyOpen::Opened(Self(Opened::new(package)))),
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
        &self.0
    }

    /// Run one read against the retained native owner, fencing both the
    /// source-backed and eager routes.  The native source owner itself keeps
    /// the payload scan lazy; this wrapper only checks the captured source
    /// version before and after the operation.
    pub(crate) fn finish<T>(&self, operation: impl FnOnce() -> Result<T>) -> Result<T> {
        self.check_source()?;
        let value = operation();
        let freshness = self.check_source();
        match freshness {
            Err(error) => Err(error),
            Ok(()) => value,
        }
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

    /// Open the retained native cell-value engine after signature admission.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn cell_values_editor(&self) -> Result<cell_values::Snapshot> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xls));
        }
        let limits = litchi_ole_common::object::Limits::default();
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = self.owner().source_bytes()?;
            let bytes = signature::unsign_compound_file(bytes, litchi_sign::cfb::Format::Xls)?;
            return Ok(litchi_xls::raw::cell_values_snapshot_from_bytes(
                self.owner(),
                bytes,
            )?);
        }
        Ok(self
            .owner()
            .edit_snapshot_with_limits(limits, OleFileLimits::default())?)
    }

    fn trait_workbook(&self) -> Result<&EagerWorkbook> {
        Ok(self.owner().reader_ref()?)
    }
}

type SheetBox<'a> = Box<dyn litchi_core::sheet::Worksheet + 'a>;

fn boxed(error: crate::Error) -> Box<dyn std::error::Error + Send + Sync> {
    Box::new(error)
}

/// The generic sheet traits over the eager reader.
///
/// On the filesystem route the first trait call reads the whole workbook
/// once. When that read fails, the fallible methods return its error and the
/// infallible ones report an empty workbook.
impl WorkbookTrait for Xls {
    fn active_sheet_index(&self) -> usize {
        self.trait_workbook()
            .map_or(0, WorkbookTrait::active_sheet_index)
    }

    fn active_worksheet(&self) -> litchi_core::sheet::Result<SheetBox<'_>> {
        self.trait_workbook().map_err(boxed)?.active_worksheet()
    }

    fn is_1904_date_system(&self) -> bool {
        self.trait_workbook()
            .is_ok_and(WorkbookTrait::is_1904_date_system)
    }

    fn worksheet_by_index(&self, index: usize) -> litchi_core::sheet::Result<SheetBox<'_>> {
        self.trait_workbook()
            .map_err(boxed)?
            .worksheet_by_index(index)
    }

    fn worksheet_by_name(&self, name: &str) -> litchi_core::sheet::Result<SheetBox<'_>> {
        self.trait_workbook()
            .map_err(boxed)?
            .worksheet_by_name(name)
    }

    fn worksheet_count(&self) -> usize {
        self.trait_workbook()
            .map_or(0, WorkbookTrait::worksheet_count)
    }

    fn worksheet_names(&self) -> &[String] {
        self.trait_workbook()
            .map_or(&[], WorkbookTrait::worksheet_names)
    }

    fn worksheets(&self) -> Box<dyn litchi_core::sheet::WorksheetIterator<'_> + '_> {
        match self.trait_workbook() {
            Ok(workbook) => workbook.worksheets(),
            Err(error) => Box::new(Failed(Some(error))),
        }
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

/// A worksheet iterator that yields one error: the failed materialization.
struct Failed(Option<crate::Error>);

impl<'a> litchi_core::sheet::WorksheetIterator<'a> for Failed {
    fn next(&mut self) -> Option<litchi_core::sheet::Result<SheetBox<'a>>> {
        self.0.take().map(|error| Err(boxed(error)))
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
    use litchi_core::sheet::CellValue;
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
        let ole = OleFile::open(Cursor::new(fixture("Simple.xls"))).unwrap();
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
        match detect_format_smart(bytes) {
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
            let mut rows = worksheet.rows();
            while let Some(row) = rows.next() {
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
            let Ok(mut leaf) = OleFile::open(Cursor::new(fixture(&name))) else {
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
            let owner = SourceBackedWorkbook::from_path(&path).unwrap();
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

    #[test]
    fn byte_route_keeps_legacy_biff_the_source_owner_refuses() {
        // A BIFF5 workbook: the eager reader reads it, the source owner does
        // not. The byte route must keep reading it.
        let source = Arc::new(litchi_core::OwnedSource::new(fixture("Simple.xls")));
        let mut stream = SharedOleFile::open(source)
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
        let oracle = cell_values::Snapshot::from_bytes(bytes.clone()).unwrap();

        let eager = open_bytes(bytes).unwrap();
        let editor = eager.cell_values_editor().unwrap();
        assert_eq!(editor.bytes(), oracle.bytes());
        assert_eq!(editor.worksheet_count(), oracle.worksheet_count());

        #[cfg(any(unix, windows))]
        {
            let source = open_path(&fixture_path("Simple.xls"));
            let editor = source.cell_values_editor().unwrap();
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
