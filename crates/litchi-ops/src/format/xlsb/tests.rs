//! `Xlsb` handle tests: every constructor route on real fixtures, parity
//! with the facade's `sheet::Workbook` output, the source-freshness guard,
//! the generic sheet traits with exact name matching, the semantics carried
//! over from `crates/litchi/tests/xlsb_facade.rs`, and editor construction.

use std::io::{self, Cursor};
use std::path::PathBuf;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Barrier};

use litchi_core::sheet::{CellValue, WorkbookTrait};
use litchi_core::{Metadata, ReadAt, SourceVersion};
use litchi_opc::{OpcPackage, PackURI, SourceBackedPackage, SourceLimits};
use litchi_xlsb as xlsb;

use super::{SourceWorkbook, Xlsb, XlsbEditor, XlsbOwner};
use crate::detect::workbook_metadata::MetadataRepairs;
use crate::error::{Code, Error, Source};
use crate::options::{OpenOptions, SourceMode};
use litchi_core::sheet::{Cell as _, Worksheet as _};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/"
    ))
    .join(name)
}

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("read fixture")
}

/// The path route, through the unified opener.
#[cfg(any(unix, windows))]
fn open_path(path: &std::path::Path, options: &OpenOptions) -> Xlsb {
    match crate::open_with(path, options).expect("path open") {
        crate::File::Xlsb(xlsb) => xlsb,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

/// The byte route, through the unified opener.
fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Xlsb {
    match crate::open_bytes_with(bytes, options).expect("bytes open") {
        crate::File::Xlsb(xlsb) => xlsb,
        other => panic!("bytes opened as {other:?}"),
    }
}

/// The eager fallback route: smart detection's parsed OPC package.
fn open_eager(bytes: Vec<u8>, options: &OpenOptions) -> Xlsb {
    let limits = options.detection_limits();
    let Some(crate::detect::DetectedFormat::Xlsb(package)) =
        crate::detect::detect_format_smart(bytes, &limits)
    else {
        panic!("smart detection did not find XLSB");
    };
    Xlsb::from_opc_package(package, options).expect("eager open")
}

/// The prepared route: `detect_and_prepare`, then the unified opener.
#[cfg(any(unix, windows))]
fn open_prepared(path: &std::path::Path, options: &OpenOptions) -> Xlsb {
    let limits = options.detection_limits();
    let detection =
        crate::detect::detect_and_prepare(path, &limits).expect("prepared OOXML source");
    match crate::open_prepared_with(detection, options).expect("prepared open") {
        crate::File::Xlsb(xlsb) => xlsb,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

fn fnv(text: &str) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in text.bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    hash
}

/// `(fixture, text length, FNV-1a of text, sheet names, title, author)`, as
/// the facade's `sheet::Workbook::{open, from_bytes}` report them at the
/// base of this step (captured with a throwaway facade example; both facade
/// routes agree).
type Expected = (
    &'static str,
    usize,
    u64,
    &'static [&'static str],
    Option<&'static str>,
    &'static str,
);
const FACADE: &[Expected] = &[
    (
        "xlsb/Simple.xlsb",
        75,
        0x2512_a350_ba83_4f34,
        &["Sheet1", "Sheet2", "Sheet3"],
        Some("Example document"),
        "Microsoft",
    ),
    (
        "xlsb/sample.xlsb",
        261,
        0x977c_7509_224d_2756,
        &["Sheet1", "rich test", "Sheet3"],
        None,
        "",
    ),
    (
        "xlsb/date.xlsb",
        6,
        0xb3af_fda1_8113_e51c,
        &["Sheet1"],
        None,
        "Allison, Timothy B.",
    ),
    (
        "xlsb/comments.xlsb",
        26,
        0xb4ab_dbd5_03da_9dae,
        &["Sheet1", "Sheet2", "Sheet3"],
        None,
        "",
    ),
];

fn assert_facade_parity(handle: &Xlsb, expected: &Expected) {
    let (name, length, hash, names, title, author) = *expected;
    let text = handle.text().expect("text");
    assert_eq!(text.len(), length, "{name}: text length");
    assert_eq!(fnv(&text), hash, "{name}: text digest");
    assert_eq!(handle.worksheet_names().unwrap(), names, "{name}: names");
    assert_eq!(handle.worksheet_count().unwrap(), names.len());
    let metadata = handle
        .metadata()
        .unwrap()
        .expect("XLSB metadata is always reported");
    assert_eq!(metadata.title.as_deref(), title, "{name}: title");
    assert_eq!(metadata.author.as_deref(), Some(author), "{name}: author");
    assert!(handle.tolerance_report().unwrap().is_clean(), "{name}");
}

#[cfg(any(unix, windows))]
#[test]
fn every_route_matches_the_facade_on_real_fixtures() {
    for expected in FACADE {
        let path = fixture_path(expected.0);
        let bytes = fixture(expected.0);

        let by_path = open_path(&path, &OpenOptions::new());
        assert!(by_path.is_source_backed());
        assert_facade_parity(&by_path, expected);

        let by_bytes = open_bytes(bytes.clone(), &OpenOptions::new());
        assert!(by_bytes.is_source_backed());
        assert_facade_parity(&by_bytes, expected);

        let by_prepared = open_prepared(&path, &OpenOptions::new());
        assert!(by_prepared.is_source_backed());
        assert_facade_parity(&by_prepared, expected);

        // The eager owner's text is the facade's eager projection; it equals
        // the source-backed text on these fixtures, as the facade test
        // `source_backed_facade_text_matches_eager_workbook` asserts.
        let eager = open_eager(bytes.clone(), &OpenOptions::new());
        assert!(!eager.is_source_backed());
        assert_facade_parity(&eager, expected);

        let forced = open_path(&path, &OpenOptions::new().source_mode(SourceMode::Eager));
        assert!(!forced.is_source_backed());
        assert_facade_parity(&forced, expected);
    }
}

#[test]
fn metadata_matches_the_leaf_core_properties_read() {
    let bytes = fixture("xlsb/Simple.xlsb");
    let package = SourceBackedPackage::from_vec(
        bytes.clone(),
        SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let (leaf, _) =
        crate::detect::workbook_metadata::read_source_backed_metadata(&package).unwrap();
    let handle = open_bytes(bytes, &OpenOptions::new());
    assert_eq!(
        format!("{:?}", handle.metadata().unwrap()),
        format!("{:?}", Some(leaf))
    );
}

#[test]
fn names_match_exactly_on_both_owner_states() {
    let bytes = fixture("xlsb/Simple.xlsb");
    let source = open_bytes(bytes.clone(), &OpenOptions::new());
    let eager = open_eager(bytes, &OpenOptions::new());
    for handle in [&source, &eager] {
        assert_eq!(
            WorkbookTrait::worksheet_by_name(handle, "Sheet1")
                .unwrap()
                .name(),
            "Sheet1"
        );
        assert!(WorkbookTrait::worksheet_by_name(handle, "sheet1").is_err());
        assert!(WorkbookTrait::worksheet_by_name(handle, "SHEET1").is_err());
    }
    let missing = WorkbookTrait::worksheet_by_name(&source, "sheet1")
        .err()
        .unwrap();
    assert_eq!(missing.code(), Code::NotFound);
    let out_of_range = WorkbookTrait::worksheet_by_index(&source, 3).err().unwrap();
    assert_eq!(out_of_range.code(), Code::OutOfRange);
}

#[test]
fn workbook_trait_surface_matches_across_owner_states() {
    let bytes = fixture("xlsb/Simple.xlsb");
    let source = open_bytes(bytes.clone(), &OpenOptions::new());
    let eager = open_eager(bytes, &OpenOptions::new());
    assert_eq!(
        WorkbookTrait::worksheet_names(&source),
        WorkbookTrait::worksheet_names(&eager)
    );
    assert_eq!(source.active_sheet_index(), eager.active_sheet_index());
    assert_eq!(source.is_1904_date_system(), eager.is_1904_date_system());
    assert_eq!(
        source.active_worksheet().unwrap().name(),
        eager.active_worksheet().unwrap().name()
    );
    let mut left = source.worksheets();
    let mut right = eager.worksheets();
    loop {
        match (left.next(), right.next()) {
            (None, None) => break,
            (Some(left), Some(right)) => {
                let (left, right) = (left.unwrap(), right.unwrap());
                assert_eq!(left.name(), right.name());
                // Carried over verbatim from the facade: the source-backed
                // adapter reports one-based bounds and takes one-based cell
                // coordinates, while the eager leaf owner is zero-based.
                assert_eq!(
                    left.dimensions(),
                    right
                        .dimensions()
                        .map(|(a, b, c, d)| (a + 1, b + 1, c + 1, d + 1))
                );
                assert_eq!(left.row_count(), right.row_count());
                for row in 0..right.row_count() {
                    assert_eq!(left.row(row).unwrap(), right.row(row).unwrap());
                }
                let mut left_cells = left.cells();
                for right in right.cells() {
                    let (left, right) = (left_cells.next().unwrap().unwrap(), right.unwrap());
                    assert_eq!(left.value(), right.value());
                    assert_eq!(left.row(), right.row() + 1);
                }
                assert_eq!(
                    left.cell_value(1, 1).unwrap(),
                    right.cell_value(0, 0).unwrap()
                );
            },
            _ => panic!("worksheet iterators differ in length"),
        }
    }
    assert!(format!("{source:?}").contains("worksheet_count: 3"));
    assert!(format!("{eager:?}").contains("worksheet_count: 3"));
}

// --- Builder-built workbooks (xlsb_facade.rs) -----------------------------

fn mixed_chart_tabs() -> Vec<u8> {
    use xlsb::builder::chart::{Anchor, Chart};
    use xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "DATA");
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");
    let mut writer = Builder::new();
    writer.add_worksheet(data);
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .unwrap();
    writer.add_worksheet(tail);
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .unwrap();
    output.into_inner()
}

fn rewrite_active_book_view(bytes: &[u8], active_catalog_position: u32) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let mut rewritten = Vec::new();
    {
        let mut writer = xlsb::raw::Encoder::new(&mut rewritten, xlsb::raw::Limits::DEFAULT);
        for record in xlsb::raw::Records::new(&original, xlsb::raw::Limits::DEFAULT) {
            let record = record.unwrap();
            let mut payload = record.payload().to_vec();
            if record.kind() == xlsb::raw::kind::BOOK_VIEW {
                payload[24..28].copy_from_slice(&active_catalog_position.to_le_bytes());
            }
            writer.write_record(record.kind(), &payload).unwrap();
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn malformed_second_sheet() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(
        Cursor::new(fixture("xlsb/Simple.xlsb")),
        SourceLimits::default(),
    )
    .unwrap();
    let sheet = PackURI::new("/xl/worksheets/sheet2.bin").unwrap();
    package
        .get_part_mut(&sheet)
        .unwrap()
        .set_blob(vec![0xff; 8]);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn chart_tabs_are_skipped_and_an_active_chart_is_a_typed_refusal() {
    let bytes = mixed_chart_tabs();
    let source = open_bytes(bytes.clone(), &OpenOptions::new());
    let eager = open_eager(bytes.clone(), &OpenOptions::new());
    for handle in [&source, &eager] {
        assert_eq!(handle.worksheet_names().unwrap(), ["Data", "Tail"]);
        assert_eq!(handle.text().unwrap(), "DATA\nTAIL\n");
        assert_eq!(
            WorkbookTrait::worksheet_by_index(handle, 1).unwrap().name(),
            "Tail"
        );
    }

    // Point the workbook view at the chart tab (catalog position 1).
    let active_chart = rewrite_active_book_view(&bytes, 1);
    let source = open_bytes(active_chart.clone(), &OpenOptions::new());
    let error = source.active_worksheet().err().unwrap();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(
        error.leaf(),
        Source::Xlsb(xlsb::package::PackageError::UnsupportedFeature(message))
            if message.contains("not a worksheet")
    ));
    assert_eq!(source.active_sheet_index(), 0);
    let eager = open_eager(active_chart, &OpenOptions::new());
    assert!(eager.active_worksheet().is_err());
    assert_eq!(eager.active_sheet_index(), 0);

    // Point it at the tail worksheet behind the chart (catalog position 2).
    let source = open_bytes(rewrite_active_book_view(&bytes, 2), &OpenOptions::new());
    assert_eq!(source.active_sheet_index(), 1);
    assert_eq!(source.active_worksheet().unwrap().name(), "Tail");
}

#[test]
fn source_owner_defers_a_malformed_unselected_worksheet() {
    let handle = open_bytes(malformed_second_sheet(), &OpenOptions::new());
    assert_eq!(
        handle.worksheet_names().unwrap(),
        ["Sheet1", "Sheet2", "Sheet3"]
    );
    let first = WorkbookTrait::worksheet_by_index(&handle, 0).unwrap();
    assert!(first.cell_value(1, 1).is_ok());
    let second = WorkbookTrait::worksheet_by_index(&handle, 1).unwrap();
    let error = second.cell_value(1, 1).unwrap_err();
    assert_eq!(error.code(), Code::CorruptFile);
    assert!(handle.text().is_err());
}

#[test]
fn one_based_trait_cell_reads_reject_zero_and_out_of_grid_before_materialization() {
    let (reads, handle) = counted();
    let worksheet = WorkbookTrait::worksheet_by_index(&handle, 0).unwrap();
    let before = reads.load(Ordering::SeqCst);

    for &(row, column) in &[
        (0, 1),
        (1, 0),
        (1_048_577, 1),
        (1, 16_385),
        (u32::MAX, u32::MAX),
    ] {
        let error = worksheet
            .cell(row, column)
            .expect_err("invalid cell must fail");
        assert_eq!(error.code(), Code::OutOfRange);

        let error = worksheet
            .cell_value(row, column)
            .expect_err("invalid cell value must fail");
        assert_eq!(error.code(), Code::OutOfRange);
    }

    let private_reference = "private-reference!A1";
    let error = worksheet
        .cell_by_coordinate(private_reference)
        .err()
        .unwrap();
    assert_eq!(error.code(), Code::InvalidArgument);
    assert!(!error.to_string().contains(private_reference));
    assert!(!format!("{error:?}").contains(private_reference));

    assert_eq!(
        reads.load(Ordering::SeqCst),
        before,
        "invalid coordinates must be rejected before grid materialization"
    );
}

// --- Source freshness ----------------------------------------------------

#[cfg(any(unix, windows))]
#[test]
fn a_changed_source_is_a_typed_io_error_on_every_read() {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), fixture("xlsb/Simple.xlsb")).unwrap();
    let handle = open_path(file.path(), &OpenOptions::new());
    let first = WorkbookTrait::worksheet_by_index(&handle, 0).unwrap();
    let mut worksheets = WorkbookTrait::worksheets(&handle);
    assert!(first.cell_value(1, 1).is_ok());

    std::fs::write(file.path(), mixed_chart_tabs()).unwrap();

    let changed = |error: Error| {
        assert_eq!(error.code(), Code::IoError, "{error}");
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Xlsb));
    };
    changed(handle.text().unwrap_err());
    changed(handle.metadata().unwrap_err());
    changed(handle.tolerance_report().unwrap_err());
    changed(handle.worksheet_names().unwrap_err());
    changed(handle.worksheet_count().unwrap_err());
    for result in [
        WorkbookTrait::worksheet_by_index(&handle, 0),
        WorkbookTrait::worksheet_by_index(&handle, usize::MAX),
        WorkbookTrait::worksheet_by_name(&handle, "Sheet1"),
        WorkbookTrait::worksheet_by_name(&handle, "missing"),
        WorkbookTrait::active_worksheet(&handle),
    ] {
        changed(result.expect_err("stale selector refuses before catalog lookup"));
    }
    changed(first.cell_value(1, 1).unwrap_err());
    let mut cells = first.cells();
    changed(
        cells
            .next()
            .unwrap()
            .expect_err("stale cell iterator reports one error"),
    );
    assert!(cells.next().is_none());
    let mut rows = first.rows();
    changed(
        rows.next()
            .unwrap()
            .expect_err("stale row iterator reports one error"),
    );
    assert!(rows.next().is_none());
    changed(
        worksheets
            .next()
            .unwrap()
            .expect_err("stale worksheet iterator reports one error"),
    );
    assert!(worksheets.next().is_none());
    changed(handle.editor().err().expect("materialization refuses"));
}

// --- Single materialization under concurrency ---------------------------

#[derive(Clone)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: Arc<AtomicUsize>,
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let offset = usize::try_from(offset).unwrap_or(usize::MAX);
        if output.is_empty() || offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x584c_5342_4f50_5300, 0))
    }
}

fn counted() -> (Arc<AtomicUsize>, Xlsb) {
    let reads = Arc::new(AtomicUsize::new(0));
    let source: Arc<dyn ReadAt> = Arc::new(CountingSource {
        bytes: Arc::new(fixture("xlsb/Simple.xlsb")),
        reads: Arc::clone(&reads),
    });
    let workbook = xlsb::SourceBackedWorkbook::from_read_at(source).unwrap();
    let owner = XlsbOwner::Source(SourceWorkbook::new(workbook).unwrap());
    let handle = Xlsb::with_owner(
        owner,
        Metadata::default(),
        MetadataRepairs::default(),
        &OpenOptions::new(),
        SourceLimits::default(),
    );
    reads.store(0, Ordering::SeqCst);
    (reads, handle)
}

#[test]
fn concurrent_first_reads_materialize_a_worksheet_once() {
    let (baseline_reads, baseline) = counted();
    let expected: CellValue = WorkbookTrait::worksheet_by_index(&baseline, 0)
        .unwrap()
        .cell_value(1, 1)
        .unwrap()
        .into_owned();
    let single = baseline_reads.load(Ordering::SeqCst);

    let (reads, handle) = counted();
    let barrier = Arc::new(Barrier::new(3));
    let workers = (0..2)
        .map(|_| {
            let handle = handle.clone();
            let barrier = Arc::clone(&barrier);
            std::thread::spawn(move || {
                barrier.wait();
                WorkbookTrait::worksheet_by_index(&handle, 0)
                    .unwrap()
                    .cell_value(1, 1)
                    .unwrap()
                    .into_owned()
            })
        })
        .collect::<Vec<_>>();
    barrier.wait();
    for worker in workers {
        assert_eq!(worker.join().unwrap(), expected);
    }
    assert_eq!(reads.load(Ordering::SeqCst), single);
}

// --- Editor construction (L-XLSB-1) --------------------------------------

#[test]
fn source_handles_materialize_the_eager_editor_from_the_same_package() {
    let bytes = fixture("xlsb/Simple.xlsb");
    let handle = open_bytes(bytes.clone(), &OpenOptions::new());
    let editor = handle.editor().expect("materialize the eager editor");
    assert!(matches!(editor, XlsbEditor::Materialized(_)));
    assert_eq!(
        WorkbookTrait::worksheet_names(editor.workbook().unwrap()),
        ["Sheet1", "Sheet2", "Sheet3"]
    );
    editor
        .edit_cell_values(0)
        .expect("eager cell-value edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet selector resolves");
    // The source-backed handle stays readable.
    assert_eq!(handle.text().unwrap().len(), 75);

    let eager = open_eager(bytes.clone(), &OpenOptions::new());
    let editor = eager.editor().unwrap();
    assert!(matches!(editor, XlsbEditor::Eager(_)));
    editor
        .edit_cell_values(0)
        .expect("eager cell-value edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet selector resolves");

    for handle in [
        open_bytes(bytes.clone(), &OpenOptions::new().retain_source(false)),
        open_eager(bytes.clone(), &OpenOptions::new().retain_source(false)),
    ] {
        assert_eq!(handle.text().unwrap().len(), 75, "reads still work");
        let error = handle.editor().err().expect("editing was not retained");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(error.suggestion().is_some());
    }
}

#[test]
fn clones_share_one_state() {
    let handle = open_bytes(fixture("xlsb/date.xlsb"), &OpenOptions::new());
    let clone = handle.clone();
    assert!(Arc::ptr_eq(&handle.0, &clone.0));
    assert_eq!(handle.format(), litchi_core::FileFormat::Xlsb);
    std::thread::spawn(move || clone.text().unwrap())
        .join()
        .unwrap();
}

// --- Fence scopes (record 0310) ------------------------------------------

/// A source that counts its version observations and can be changed.
struct Versioned {
    bytes: Vec<u8>,
    revision: std::sync::atomic::AtomicU64,
    versions: AtomicUsize,
}

impl ReadAt for Versioned {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset).unwrap_or(usize::MAX);
        if output.is_empty() || offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::SeqCst);
        Ok(SourceVersion::new(
            0x584c_5342_4f50_5301,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

/// Observations of `selector`'s second full run over a generated `Data`
/// sheet of `rows` numbers (1, 2, ...), and its match count.
fn observed_query(rows: u32, selector: &str) -> (usize, usize) {
    let mut sheet = xlsb::builder::Worksheet::new("Data");
    for row in 0..rows {
        sheet.set_cell(row, 0, f64::from(row + 1));
    }
    let mut builder = xlsb::Builder::new();
    builder.add_worksheet(sheet);
    let source = Arc::new(Versioned {
        bytes: builder.build().unwrap().to_bytes().unwrap(),
        revision: std::sync::atomic::AtomicU64::new(0),
        versions: AtomicUsize::new(0),
    });
    let workbook = xlsb::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let owner = XlsbOwner::Source(SourceWorkbook::new(workbook).unwrap());
    let file = crate::File::Xlsb(Xlsb::with_owner(
        owner,
        Metadata::default(),
        MetadataRepairs::default(),
        &OpenOptions::new(),
        SourceLimits::default(),
    ));
    let matches = file.query(selector).unwrap().count();
    let before = source.versions.load(Ordering::SeqCst);
    assert_eq!(file.query(selector).unwrap().count(), matches);
    (matches, source.versions.load(Ordering::SeqCst) - before)
}

/// Record 0310: inside a query step the node reads the step makes (a cell's
/// value for `cell[value>N]`, its text) share the step's closing check, as
/// the XLSX adapter's do. Every match is still checked (owner decision
/// 2026-10-14): ten more cells cost ten more observations when each
/// matches, and none when none does. A change between two pulls still
/// fails the next pull.
#[test]
fn query_steps_share_one_check_with_their_node_reads() {
    for (selector, per_cell) in [("cell", 1), ("cell[value>0]", 1), ("cell[value>1000]", 0)] {
        let (small_matches, small) = observed_query(10, selector);
        let (large_matches, large) = observed_query(20, selector);
        assert_eq!(large_matches - small_matches, 10 * per_cell, "{selector}");
        assert_eq!(
            large - small,
            10 * per_cell,
            "{selector}: {small} -> {large}"
        );
    }

    let source = Arc::new(Versioned {
        bytes: fixture("xlsb/Simple.xlsb"),
        revision: std::sync::atomic::AtomicU64::new(0),
        versions: AtomicUsize::new(0),
    });
    let workbook = xlsb::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let owner = XlsbOwner::Source(SourceWorkbook::new(workbook).unwrap());
    let file = crate::File::Xlsb(Xlsb::with_owner(
        owner,
        Metadata::default(),
        MetadataRepairs::default(),
        &OpenOptions::new(),
        SourceLimits::default(),
    ));
    let mut matches = file.query("sheet, cell").unwrap();
    assert!(matches.next().unwrap().is_ok());
    source.revision.store(1, Ordering::SeqCst);
    assert_eq!(matches.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(matches.next().is_none());
}
