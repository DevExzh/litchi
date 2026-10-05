//! Handle for Excel binary (`.xls`) inputs (record 0003 §2.3, M0a step S5d).
//!
//! The facade's workbook enum had three XLS variants. `XlsFile` and `Other`
//! were never constructed (finding F12) and are gone. The two live routes
//! become the two states of one owner:
//!
//! - **Source**: the filesystem route, where detection opened a positional
//!   [`SourceBackedWorkbook`] over a [`SharedOleFile`]
//!   (`Xls::from_source`). Worksheets are read lazily through the
//!   retained source, and every read is fenced against a source change.
//! - **Eager**: the byte route, where detection parsed the input into an
//!   in-memory [`OleFile`] and the eager [`litchi_xls::Workbook`] reads it
//!   (`Xls::from_ole_file`).
//!
//! Record 0003 proposed moving the byte route onto the source-backed owner
//! as well. That would change what the byte route reads. The source owner
//! refuses BIFF5 and older workbooks (`UnsupportedBiffVersion`), which the
//! eager reader reads. Over the 54 workbooks in `test-data/ole/xls`, the
//! source owner's text read fails on 4 that the eager reader reads, and the
//! eager reader omits worksheets of 8 that the source owner reads in full;
//! the two routes of the facade already differ on those files today. The
//! tolerance report and the metadata error policy differ as well. The byte
//! route therefore keeps the eager reader, which is `Sync` and already owns
//! the input bytes, so it needs no extra retention.

use std::io::Cursor;
use std::sync::{Arc, OnceLock};

use litchi_cfb::{OleError, OleFile, OleFileLimits, SharedOleFile};
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet as _};
use litchi_core::{FileFormat, Metadata};
use litchi_ole_common::property_set::{PropertySetReader, SharedPropertySetReader};
use litchi_xls::{SourceBackedError, SourceBackedWorkbook, cell_values};

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::{OpenOptions, Result};

/// An opened legacy Excel (`.xls`) workbook.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`.
#[derive(Clone)]
pub struct Xls(Arc<Opened<XlsOwner>>);

/// The two states of [`Xls`] (record 0003 §2.3).
#[allow(
    clippy::large_enum_variant,
    reason = "one owner per handle, behind its `Arc`; boxing the source state would add an allocation to every filesystem open"
)]
enum XlsOwner {
    /// Filesystem route: the positional owner and its CFB view.
    Source(XlsSource),
    /// Byte route: the eager reader over the in-memory input.
    Eager(Box<EagerXls>),
}

struct XlsSource {
    workbook: SourceBackedWorkbook,
    /// The CFB view the owner was opened over. The owner holds it too; the
    /// handle keeps its own `Arc` for the metadata streams and for editors.
    cfb: Arc<SharedOleFile>,
    /// Metadata, read on first use and kept once it has been read.
    metadata: OnceLock<Metadata>,
    /// The eager reader the generic sheet traits read through, materialized
    /// from the retained source on their first use, never at open.
    eager: OnceLock<std::result::Result<Box<EagerWorkbook>, crate::Error>>,
}

type EagerWorkbook = litchi_xls::Workbook<Cursor<Vec<u8>>>;

struct EagerXls {
    workbook: EagerWorkbook,
    /// Metadata, read at open as the facade reads it.
    metadata: Metadata,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Xls>();
    assert!(size_of::<Xls>() == size_of::<usize>());
};

impl Xls {
    /// Adopts the positional owner filesystem detection opened (the facade's
    /// `WorkbookSourcePathDetection::Xls` arm). Nothing is read here.
    ///
    /// The owner already retains `cfb`, so [`OpenOptions::retain_source`]
    /// changes nothing on this route.
    pub(crate) fn from_source(
        workbook: SourceBackedWorkbook,
        cfb: Arc<SharedOleFile>,
        options: &OpenOptions,
    ) -> Self {
        let _ = options;
        Self(Opened::new(XlsOwner::Source(XlsSource {
            workbook,
            cfb,
            metadata: OnceLock::new(),
            eager: OnceLock::new(),
        })))
    }

    /// Opens the workbook from the OLE file byte detection already parsed
    /// (the facade's `DetectedFormat::Xls` arm): the metadata streams are
    /// projected, with an unreadable projection reported as empty metadata,
    /// and the eager reader adopts the parse (with `password` when
    /// [`Locked::unlock`] supplies one).
    ///
    /// The eager reader owns the input bytes for its lifetime, so
    /// [`OpenOptions::retain_source`] changes nothing on this route either.
    /// An encrypted workbook opened without a password comes back as
    /// [`LegacyOpen::PasswordRequired`] with the container unchanged.
    ///
    /// [`Locked::unlock`]: crate::Locked
    pub(crate) fn from_ole_file(
        mut ole: OleFile<Cursor<Vec<u8>>>,
        options: &OpenOptions,
        password: Option<&str>,
    ) -> Result<LegacyOpen<Self>> {
        let _ = options;
        let metadata = ole.get_metadata().map(Metadata::from).unwrap_or_default();
        let leaf_options = litchi_xls::OpenOptions::new().with_password(password);
        match litchi_xls::Workbook::from_ole_file_or_password_required(ole, leaf_options)? {
            Ok(workbook) => Ok(LegacyOpen::Opened(Self(Opened::new(XlsOwner::Eager(
                Box::new(EagerXls { workbook, metadata }),
            ))))),
            Err(ole) => Ok(LegacyOpen::PasswordRequired(Box::new(ole))),
        }
    }

    /// The same workbook under another signature disposition (record 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// The owner every read goes through.
    fn owner(&self) -> &XlsOwner {
        &self.0
    }

    /// Whether the owner is the positional source-backed workbook, for the
    /// tests that pin each route's owner.
    #[cfg(test)]
    pub(crate) fn is_source_backed(&self) -> bool {
        matches!(self.owner(), XlsOwner::Source(_))
    }

    /// The names of the ordinary worksheets, in tab order.
    ///
    /// # Errors
    ///
    /// On the filesystem route, returns a typed error when the source changed
    /// or the workbook globals cannot be read.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match self.owner() {
            XlsOwner::Source(source) => Ok(source.workbook.worksheet_names()?),
            XlsOwner::Eager(eager) => Ok(eager.workbook.worksheet_names().to_vec()),
        }
    }

    /// The number of ordinary worksheets.
    ///
    /// # Errors
    ///
    /// On the filesystem route, returns a typed error when the source
    /// changed.
    pub fn worksheet_count(&self) -> Result<usize> {
        match self.owner() {
            XlsOwner::Source(source) => Ok(source.workbook.worksheet_count()?),
            XlsOwner::Eager(eager) => Ok(eager.workbook.worksheet_count()),
        }
    }

    /// All worksheet text: one line per row, cells separated by tabs, as the
    /// facade's `Workbook::text` returns it.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a worksheet cannot be read or, on the
    /// filesystem route, when the source changed before or after the read.
    pub fn text(&self) -> Result<String> {
        match self.owner() {
            XlsOwner::Source(source) => source.text(),
            XlsOwner::Eager(eager) => eager.text(),
        }
    }

    /// The workbook's OLE property-set metadata.
    ///
    /// The byte route read it at open. The filesystem route reads it on
    /// first use, keeps it, and fences every call against a source change;
    /// absent or malformed metadata streams read as empty metadata.
    ///
    /// Always `Some` when it succeeds.
    ///
    /// # Errors
    ///
    /// On the filesystem route, returns a typed error when the source changed
    /// or a metadata stream cannot be read for a reason other than its
    /// content.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        match self.owner() {
            XlsOwner::Source(source) => source.metadata().map(Some),
            XlsOwner::Eager(eager) => Ok(Some(eager.metadata.clone())),
        }
    }

    /// Everything the reader repaired so far: the globals parse, the
    /// metadata streams, and the worksheet scans performed so far.
    ///
    /// # Errors
    ///
    /// On the filesystem route, returns a typed error when the source
    /// changed.
    pub fn tolerance_report(&self) -> Result<litchi_xls::ToleranceReport> {
        match self.owner() {
            XlsOwner::Source(source) => source.tolerance_report(),
            XlsOwner::Eager(eager) => Ok(eager.workbook.tolerance_report().clone()),
        }
    }

    /// Builds the cell-value editor (L-XLS-1).
    ///
    /// The filesystem route builds it over the retained CFB view through
    /// [`cell_values::Snapshot::from_shared_ole_file`]; the byte route from
    /// the eager reader's own bytes. Both copy the input once and let the
    /// editor index it, because the editor's package capture runs on the
    /// cursor-based reader (the fallback of record 0003, risk R2). Both costs
    /// fall here, at the first edit, not at open. After
    /// [`Signed::unsign`](crate::Signed::unsign) the copy is one whose
    /// signature storages the signature leaf removed
    /// (`litchi_sign::cfb::Editor::clear`, record 0013), so a commit from it
    /// carries no signature.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// while the workbook is still signed, with a suggestion naming
    /// `Signed::unsign`, and the leaf's typed refusal otherwise.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn cell_values_editor(&self) -> Result<cell_values::Snapshot> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xls));
        }
        let limits = litchi_ole_common::object::Limits::default();
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = match self.owner() {
                XlsOwner::Source(source) => signature::compound_file_bytes(&source.cfb)?,
                XlsOwner::Eager(eager) => eager.workbook.get_ref().get_ref().clone(),
            };
            let bytes = signature::unsign_compound_file(bytes, litchi_sign::cfb::Format::Xls)?;
            return Ok(cell_values::Snapshot::from_bytes_with_limits(
                bytes, limits,
            )?);
        }
        Ok(match self.owner() {
            XlsOwner::Source(source) => cell_values::Snapshot::from_shared_ole_file(
                &source.cfb,
                limits,
                OleFileLimits::default(),
            )?,
            XlsOwner::Eager(eager) => cell_values::Snapshot::from_bytes_with_limits(
                eager.workbook.get_ref().get_ref().clone(),
                limits,
            )?,
        })
    }
}

impl XlsSource {
    /// The TOCTOU fence of the facade's source adapter: the owner's source
    /// must still be the one it was opened over.
    fn ensure_current(&self) -> Result<()> {
        self.workbook.source_version()?;
        Ok(())
    }

    fn text(&self) -> Result<String> {
        self.ensure_current()?;
        let result = self.workbook.text();
        self.ensure_current()?;
        Ok(result?)
    }

    fn metadata(&self) -> Result<Metadata> {
        self.ensure_current()?;
        if let Some(metadata) = self.metadata.get() {
            let metadata = metadata.clone();
            self.ensure_current()?;
            return Ok(metadata);
        }
        let metadata = match self.cfb.get_metadata() {
            Ok(metadata) => Metadata::from(metadata),
            Err(error) if metadata_error_is_soft(&error) => Metadata::default(),
            Err(error) => return Err(SourceBackedError::from(error).into()),
        };
        self.ensure_current()?;
        let metadata = self.metadata.get_or_init(|| metadata).clone();
        self.ensure_current()?;
        Ok(metadata)
    }

    fn tolerance_report(&self) -> Result<litchi_xls::ToleranceReport> {
        self.ensure_current()?;
        // The source owner never reads the metadata streams; their repairs
        // are folded in here, as the eager owner records them at open.
        let mut report = self.workbook.tolerance_report();
        report.append_property_sets(|path| self.cfb.property_set_stream(path));
        self.ensure_current()?;
        Ok(report)
    }
}

impl EagerXls {
    fn text(&self) -> Result<String> {
        let mut out = String::new();
        for index in 0..self.workbook.worksheet_count() {
            let worksheet = self.workbook.xls_worksheet(index)?;
            for row in 0..worksheet.row_count() {
                for column in 0..worksheet.column_count() {
                    if column > 0 {
                        out.push('\t');
                    }
                    if let Some(cell) = worksheet.get_cell(cell_index(row), cell_index(column)) {
                        append_cell_text(&mut out, cell.value());
                    }
                }
                out.push('\n');
            }
        }
        Ok(out)
    }
}

/// Worksheet dimensions come from `u32` cell coordinates (`max + 1`), so
/// every index below them fits back into `u32`.
fn cell_index(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

/// One cell's text, exactly as the facade's workbook text projection writes
/// it.
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

impl Xls {
    /// The eager reader behind the generic sheet traits: the byte route's
    /// own, or one materialized once from the filesystem route's retained
    /// source. The facade's `open_xls_workbook_dyn` read paths through the
    /// same eager reader, so trait reads keep its semantics on both routes.
    fn trait_workbook(&self) -> Result<&EagerWorkbook> {
        match self.owner() {
            XlsOwner::Eager(eager) => Ok(&eager.workbook),
            XlsOwner::Source(source) => source
                .eager
                .get_or_init(|| {
                    source
                        .workbook
                        .materialize_eager()
                        .map(Box::new)
                        .map_err(crate::Error::from)
                })
                .as_deref()
                .map_err(Clone::clone),
        }
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
        let owner = match self.owner() {
            XlsOwner::Source(_) => "source",
            XlsOwner::Eager(_) => "eager",
        };
        formatter
            .debug_struct("Xls")
            .field("owner", &owner)
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

/// Metadata stream failures that mean "no usable metadata" rather than an
/// I/O, limit or source-change failure (the facade's
/// `xls_metadata_error_is_soft`).
fn metadata_error_is_soft(error: &OleError) -> bool {
    matches!(
        error,
        OleError::InvalidFormat(_)
            | OleError::InvalidData(_)
            | OleError::NotOleFile
            | OleError::CorruptedFile(_)
            | OleError::StreamNotFound
    )
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
    use std::io::Write;
    use std::path::{Path, PathBuf};

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
        let workbook = litchi_xls::Workbook::new(Cursor::new(fixture(name))).unwrap();
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
            let workbook = litchi_xls::Workbook::from_ole_file(leaf);
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
            assert!(matches!(xls.owner(), XlsOwner::Eager(_)));
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
            if !matches!(xls.owner(), XlsOwner::Source(_)) {
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
