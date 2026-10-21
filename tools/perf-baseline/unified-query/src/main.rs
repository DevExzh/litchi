//! Paired query cost (record 0262): full-match and early-terminating unified
//! queries and warm `get`s on the XLSX, DOCX and PPTX adapters, each against
//! a hand-written loop over the source-backed leaf it wraps.
//!
//! Only public APIs are used (`litchi_ops::open_bytes`, `File::get`,
//! `File::query`, `Matches::{first, limit}`, `Plan::compile` and the leaves'
//! own reads), so the same source builds against the code before and after a
//! change, and a driver can alternate the two binaries (ABBA across
//! processes, on top of the in-process ABBA below).
//!
//! ```text
//! unified-query-baseline list
//! unified-query-baseline time <route|prefix|all> [samples]   # JSON lines
//! unified-query-baseline profile <route> <calls>             # unified leg only
//! ```
//!
//! Corpora (generated in process, outside the timer):
//! - XLSX: one `Data` sheet with a bold header and N data rows under a
//!   `Sales` table (`crates/litchi-ops/tests/support/xlsx_workbook.rs`);
//!   N = 1,000 and 50,000 (the default stream-event limit refuses 100,000);
//! - DOCX: `test-data/ooxml/docx/drawing.docx` and
//!   `test-data/poi/test-data/document/saut_page.docx`;
//! - PPTX: generated decks of 10 slides × 8 boxes and 500 × 40
//!   (`crates/litchi-ops/tests/support/pptx_presentation.rs`);
//! - XLSB and XLS (record 0267): the XLSX corpus's shape (a header and N
//!   rows of region, sales, status) written by the leaves' own builders.
//!
//! Record 0267 adds `*_cold_*` routes, which open a fresh handle in every
//! call and stop at the first or tenth match (the cold first-match cost of
//! a worksheet no earlier read has decoded), and warm XLSB routes.
//!
//! Record 0315 adds `xlsb_*_all_cell_owned` and `xlsb_*_all_cell_fenced`,
//! the fair leaf legs record 0280 added for XLSX and PPTX.
//!
//! Each sample times `reps` calls of each leg, unified then leaf for samples
//! 0 and 3 (mod 4), leaf then unified otherwise. Every route checks before
//! timing that both legs return the same match count.

use std::hint::black_box;
use std::io::Cursor;
use std::time::Instant;

use litchi_core::{FileFormat, Position};
use litchi_docx::{BlockKind, Blocks, source_backed};
use litchi_ops::File;
use litchi_ops::query::{Plan, Query};
use litchi_pptx::SourceBackedPresentation;
use litchi_pptx::shape::Shape;
use litchi_xlsx::SourceBackedWorkbook;

#[allow(
    dead_code,
    reason = "shared test fixture; the harness uses one generator"
)]
#[path = "../../../../crates/litchi-ops/tests/support/pptx_presentation.rs"]
mod deck;
#[allow(dead_code, reason = "shared test fixture")]
#[path = "../../../../crates/litchi-ops/tests/support/xlsx_workbook.rs"]
mod workbook;

/// One measured route: a unified leg and a leaf leg doing the same work,
/// each returning its match count (or 1 for a get).
struct Route<'a> {
    name: String,
    reps: usize,
    unified: Box<dyn FnMut() -> usize + 'a>,
    leaf: Box<dyn FnMut() -> usize + 'a>,
}

impl<'a> Route<'a> {
    fn new(
        name: impl Into<String>,
        reps: usize,
        unified: impl FnMut() -> usize + 'a,
        leaf: impl FnMut() -> usize + 'a,
    ) -> Self {
        Self {
            name: name.into(),
            reps,
            unified: Box::new(unified),
            leaf: Box::new(leaf),
        }
    }
}

fn median(values: &mut [f64]) -> f64 {
    values.sort_by(f64::total_cmp);
    let middle = values.len() / 2;
    if values.len() % 2 == 1 {
        values[middle]
    } else {
        (values[middle - 1] + values[middle]) / 2.0
    }
}

fn time_leg(reps: usize, leg: &mut dyn FnMut() -> usize) -> f64 {
    let start = Instant::now();
    for _ in 0..reps {
        black_box(leg());
    }
    start.elapsed().as_nanos() as f64 / reps as f64
}

fn time(route: &mut Route<'_>, samples: usize) {
    let (unified_matches, leaf_matches) = ((route.unified)(), (route.leaf)());
    assert_eq!(
        unified_matches, leaf_matches,
        "{}: the unified and leaf legs disagree",
        route.name
    );
    for _ in 0..2 {
        time_leg(route.reps, &mut route.unified);
        time_leg(route.reps, &mut route.leaf);
    }
    let (mut unified, mut leaf, mut ratios) = (Vec::new(), Vec::new(), Vec::new());
    for sample in 0..samples {
        let (u, l) = if sample % 4 == 0 || sample % 4 == 3 {
            let u = time_leg(route.reps, &mut route.unified);
            (u, time_leg(route.reps, &mut route.leaf))
        } else {
            let l = time_leg(route.reps, &mut route.leaf);
            (time_leg(route.reps, &mut route.unified), l)
        };
        unified.push(u);
        leaf.push(l);
        ratios.push(u / l);
    }
    let line = serde_json::json!({
        "route": route.name,
        "matches": unified_matches,
        "reps": route.reps,
        "samples": samples,
        "unified_ns": unified.clone(),
        "leaf_ns": leaf.clone(),
        "unified_median_ns": median(&mut unified),
        "leaf_median_ns": median(&mut leaf),
        "ratio_median": median(&mut ratios),
    });
    println!("{line}");
}

fn plan(selector: &'static str, format: FileFormat) -> Plan<'static> {
    let query = Query::parse(selector).expect("parses");
    Plan::compile(&query, format).expect("compiles")
}

fn count_matches(file: &File, plan: &Plan<'_>) -> usize {
    let mut matches = 0;
    for node in file.query(plan).expect("runs") {
        black_box(node.expect("a match"));
        matches += 1;
    }
    matches
}

fn first_match(file: &File, plan: &Plan<'_>) -> usize {
    usize::from(black_box(file.query(plan).expect("runs").first().expect("reads")).is_some())
}

fn limited(file: &File, plan: &Plan<'_>, limit: usize) -> usize {
    let mut matches = 0;
    for node in file.query(plan).expect("runs").limit(limit) {
        black_box(node.expect("a match"));
        matches += 1;
    }
    matches
}

// --- Backing ---------------------------------------------------------------

/// How both legs open the corpus: from memory, or from a file on disk (the
/// default `litchi::open(path)` route, whose source-version fence reads the
/// file's metadata).
#[derive(Clone, Copy, PartialEq, Eq)]
enum Backing {
    Bytes,
    File,
}

impl Backing {
    const ALL: [Self; 2] = [Self::Bytes, Self::File];

    const fn name(self) -> &'static str {
        match self {
            Self::Bytes => "bytes",
            Self::File => "file",
        }
    }
}

/// Writes `bytes` to a scratch file named `name` (under `UNIFIED_QUERY_DIR`,
/// else the system temporary directory) and returns its path.
fn scratch_file(name: &str, bytes: &[u8]) -> std::path::PathBuf {
    let directory = std::env::var_os("UNIFIED_QUERY_DIR")
        .map_or_else(std::env::temp_dir, std::path::PathBuf::from);
    let path = directory.join(format!("unified-query-{}-{name}", std::process::id()));
    std::fs::write(&path, bytes).expect("writes the corpus file");
    path
}

/// The unified handle of `bytes` through `backing`.
fn open_unified(bytes: &[u8], name: &str, backing: Backing) -> (File, Option<std::path::PathBuf>) {
    match backing {
        Backing::Bytes => (litchi_ops::open_bytes(bytes.to_vec()).expect("opens"), None),
        Backing::File => {
            let path = scratch_file(name, bytes);
            (litchi_ops::open(&path).expect("opens"), Some(path))
        },
    }
}

// --- XLSX ------------------------------------------------------------------

fn xlsx_number(cell: &litchi_xlsx::Cell) -> Option<f64> {
    match cell {
        litchi_xlsx::Cell::Value(litchi_xlsx::Value::Number(number)) => number.as_f64(),
        _ => None,
    }
}

/// Every stored record of every sheet, stopping after `limit` that pass
/// `keep`.
fn xlsx_leaf_cells(
    leaf: &SourceBackedWorkbook,
    limit: usize,
    keep: fn(&litchi_xlsx::Cell) -> bool,
) -> usize {
    let mut matches = 0;
    for position in 0..leaf.len() {
        let sheet = leaf.sheet(position).expect("reads").expect("a sheet");
        let grid = sheet.grid().expect("a grid");
        let Some(extent) = grid.extents().stored() else {
            continue;
        };
        for (_, cell) in grid.cells(extent).expect("cells") {
            if keep(cell) {
                matches += 1;
                if matches == limit {
                    return matches;
                }
            }
        }
    }
    matches
}

/// Record 0280's fair leaf equivalent of a full `cell` query: every stored
/// record of every sheet, each kept as an owned, retainable handle (the
/// sheet position, a shared grid handle and the address, which is what a
/// unified node holds), and with `fenced`, one source-freshness check per
/// match, the guarantee a unified query gives on every pull.
fn xlsx_leaf_owned(leaf: &SourceBackedWorkbook, fenced: bool) -> usize {
    let mut matches = 0;
    for position in 0..leaf.len() {
        let sheet = leaf.sheet(position).expect("reads").expect("a sheet");
        let grid = sheet.grid().expect("a grid");
        let Some(extent) = grid.extents().stored() else {
            continue;
        };
        for (address, _) in grid.cells(extent).expect("cells") {
            if fenced {
                black_box(leaf.source_version().expect("current"));
            }
            black_box((position, grid.clone(), address));
            matches += 1;
        }
    }
    matches
}

fn over_1000(cell: &litchi_xlsx::Cell) -> bool {
    xlsx_number(cell).is_some_and(|value| value > 1000.0)
}

fn any_cell(_: &litchi_xlsx::Cell) -> bool {
    true
}

struct XlsxContext {
    rows: usize,
    tag: String,
    file: File,
    leaf: SourceBackedWorkbook,
    over: Plan<'static>,
    all: Plan<'static>,
}

fn xlsx_context(rows: usize, backing: Backing) -> XlsxContext {
    let bytes = workbook::workbook(rows);
    let (file, path) = open_unified(&bytes, &format!("{rows}.xlsx"), backing);
    assert_eq!(file.format(), Some(FileFormat::Xlsx));
    let leaf = match path {
        Some(path) => SourceBackedWorkbook::open(path).expect("opens"),
        None => SourceBackedWorkbook::from_reader(Cursor::new(bytes)).expect("opens"),
    };
    XlsxContext {
        rows,
        tag: format!("{rows}_{}", backing.name()),
        file,
        leaf,
        over: plan("cell[value>1000]", FileFormat::Xlsx),
        all: plan("cell", FileFormat::Xlsx),
    }
}

fn xlsx_routes(context: &XlsxContext) -> Vec<Route<'_>> {
    let (file, leaf) = (&context.file, &context.leaf);
    let (over, all) = (&context.over, &context.all);
    let rows = context.rows;
    let tag = &context.tag;
    let middle = rows / 2 + 1;
    let get_path = format!("/Data/B{middle}");
    let get_cell = format!("B{middle}");
    let full_reps = if rows > 10_000 { 3 } else { 300 };
    let mut routes = vec![
        Route::new(
            format!("xlsx_{tag}_get_cell_v1"),
            20_000,
            move || {
                black_box(file.get(black_box(get_path.as_str())).expect("resolves"));
                1
            },
            move || {
                let sheet = leaf
                    .sheet(black_box("Data"))
                    .expect("reads")
                    .expect("a sheet");
                black_box(sheet.cell(black_box(get_cell.as_str())).expect("reads"));
                1
            },
        ),
        Route::new(
            format!("xlsx_{tag}_first_cell_value_gt_v1"),
            20_000,
            move || first_match(file, over),
            move || xlsx_leaf_cells(leaf, 1, over_1000),
        ),
        Route::new(
            format!("xlsx_{tag}_limit10_cell_value_gt_v1"),
            5_000,
            move || limited(file, over, 10),
            move || xlsx_leaf_cells(leaf, 10, over_1000),
        ),
        Route::new(
            format!("xlsx_{tag}_all_cell_value_gt_v1"),
            full_reps,
            move || count_matches(file, over),
            move || xlsx_leaf_cells(leaf, usize::MAX, over_1000),
        ),
        Route::new(
            format!("xlsx_{tag}_all_cell_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsx_leaf_cells(leaf, usize::MAX, any_cell),
        ),
        // Record 0280: fair leaf legs for the same unified query.
        Route::new(
            format!("xlsx_{tag}_all_cell_owned_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsx_leaf_owned(leaf, false),
        ),
        Route::new(
            format!("xlsx_{tag}_all_cell_fenced_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsx_leaf_owned(leaf, true),
        ),
    ];
    if rows > 10_000 {
        routes.push(xlsx_rows_route(context));
    }
    routes
}

/// `row[Sales>1000]`: the rows under the table's header whose `Sales` cell
/// (column B) exceeds 1000.
fn xlsx_rows_route(context: &XlsxContext) -> Route<'_> {
    let (file, leaf) = (&context.file, &context.leaf);
    let rows = plan("row[Sales>1000]", FileFormat::Xlsx);
    Route::new(
        format!("xlsx_{}_all_row_header_gt_v1", context.tag),
        3,
        move || count_matches(file, &rows),
        move || {
            let sheet = leaf.sheet("Data").expect("reads").expect("a sheet");
            let grid = sheet.grid().expect("a grid");
            let mut matches = 0;
            for row in grid.rows().skip(1) {
                let at = litchi_xlsx::Address::at(row.index().get(), 1).expect("a coordinate");
                if let litchi_xlsx::cell::View::Stored(cell) = grid.cell(at).expect("a cell")
                    && over_1000(cell)
                {
                    matches += 1;
                }
            }
            matches
        },
    )
}

// --- Cold workbooks (record 0267) -------------------------------------------

/// The `Sales` values of the generated workbooks, repeating: half exceed
/// 1000 (as in `xlsx_workbook.rs`).
const SALES: [f64; 4] = [1200.0, 800.0, 450.5, 2000.0];
const REGIONS: [&str; 4] = ["West", "East", "North", "South"];
const STATUSES: [&str; 2] = ["open", "closed"];

/// An XLSB workbook shaped like the XLSX corpus: a `Data` sheet with a
/// header row and `rows` rows of (region, sales, status).
fn xlsb_workbook(rows: usize) -> Vec<u8> {
    let mut sheet = litchi_xlsb::builder::Worksheet::new("Data");
    for (column, header) in ["Region", "Sales", "Status"].into_iter().enumerate() {
        sheet.set_cell(0, column as u32, header.to_string());
    }
    for index in 0..rows {
        let row = index as u32 + 1;
        sheet.set_cell(row, 0, REGIONS[index % 4].to_string());
        sheet.set_cell(row, 1, SALES[index % 4]);
        sheet.set_cell(row, 2, STATUSES[index % 2].to_string());
    }
    let mut builder = litchi_xlsb::Builder::new();
    builder.add_worksheet(sheet);
    builder
        .build()
        .expect("builds")
        .to_bytes()
        .expect("serializes")
}

/// An XLS workbook of the same shape.
fn xls_workbook(rows: usize) -> Vec<u8> {
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Data").expect("a sheet");
    for (column, header) in ["Region", "Sales", "Status"].into_iter().enumerate() {
        builder
            .write_string(sheet, 0, column as u16, header)
            .expect("writes");
    }
    for index in 0..rows {
        let row = index as u32 + 1;
        builder
            .write_string(sheet, row, 0, REGIONS[index % 4])
            .expect("writes");
        builder
            .write_number(sheet, row, 1, SALES[index % 4])
            .expect("writes");
        builder
            .write_string(sheet, row, 2, STATUSES[index % 2])
            .expect("writes");
    }
    builder
        .build()
        .expect("builds")
        .to_bytes()
        .expect("serializes")
}

fn core_over_1000(value: &litchi_core::sheet::CellValue) -> bool {
    use litchi_core::sheet::CellValue;
    match value {
        CellValue::Float(value) => *value > 1000.0,
        CellValue::Int(value) => *value > 1000,
        _ => false,
    }
}

/// A freshly opened unified handle of `bytes` (or of the file at `path`),
/// then the first `limit` matches of `plan`: the cold early-stop route.
fn cold_unified(
    bytes: &[u8],
    path: Option<&std::path::Path>,
    plan: &Plan<'_>,
    limit: usize,
) -> usize {
    let file = match path {
        Some(path) => litchi_ops::open(path).expect("opens"),
        None => litchi_ops::open_bytes(bytes.to_vec()).expect("opens"),
    };
    if limit == 1 {
        first_match(&file, plan)
    } else {
        limited(&file, plan, limit)
    }
}

/// The same work through a freshly opened XLSX leaf: the first `limit`
/// stored cells over 1000, by the leaf's early-stopping range visitor.
fn cold_xlsx_leaf(bytes: &[u8], path: Option<&std::path::Path>, limit: usize) -> usize {
    let leaf = match path {
        Some(path) => SourceBackedWorkbook::open(path).expect("opens"),
        None => SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).expect("opens"),
    };
    let mut matches = 0;
    for position in 0..leaf.len() {
        let sheet = leaf.sheet(position).expect("reads").expect("a sheet");
        let flow = sheet
            .visit_cells_until(litchi_xlsx::Rect::ALL, |_, cell| {
                if over_1000(cell) {
                    matches += 1;
                    if matches == limit {
                        return std::ops::ControlFlow::Break(());
                    }
                }
                std::ops::ControlFlow::Continue(())
            })
            .expect("visits");
        if flow.is_break() {
            break;
        }
    }
    matches
}

/// The same work through a freshly opened XLSB leaf grid.
fn cold_xlsb_leaf(bytes: &[u8], path: Option<&std::path::Path>, limit: usize) -> usize {
    let leaf = xlsb_leaf(bytes, path);
    xlsb_leaf_cells(&leaf, limit, core_over_1000)
}

fn xlsb_leaf(bytes: &[u8], path: Option<&std::path::Path>) -> litchi_xlsb::SourceBackedWorkbook {
    match path {
        Some(path) => litchi_xlsb::SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
            litchi_core::FileSource::open(path).expect("opens"),
        ))
        .expect("opens"),
        None => litchi_xlsb::SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec()))
            .expect("opens"),
    }
}

/// Every stored cell of every XLSB worksheet that passes `keep`, stopping
/// after `limit`.
fn xlsb_leaf_cells(
    leaf: &litchi_xlsb::SourceBackedWorkbook,
    limit: usize,
    keep: fn(&litchi_core::sheet::CellValue) -> bool,
) -> usize {
    use litchi_core::sheet::Cell as _;
    let mut matches = 0;
    for sheet in leaf.worksheets().expect("reads") {
        let grid = sheet.grid().expect("a grid");
        for cell in grid.cells_in(litchi_xlsb::sheet::Rect::ALL).expect("cells") {
            if keep(cell.value()) {
                matches += 1;
                if matches == limit {
                    return matches;
                }
            }
        }
    }
    matches
}

/// Record 0315's fair leaf equivalent of a full XLSB `cell` query, as
/// record 0280's `xlsx_leaf_owned` is for XLSX: every stored cell of every
/// worksheet, each kept as an owned, retainable handle (the worksheet
/// position, a shared grid handle and the coordinate, which is what a
/// unified node holds), and with `fenced`, one source-freshness check per
/// match, the guarantee a unified query gives on every pull.
fn xlsb_leaf_owned(leaf: &litchi_xlsb::SourceBackedWorkbook, fenced: bool) -> usize {
    use litchi_core::sheet::Cell as _;
    let mut matches = 0;
    for (position, sheet) in leaf.worksheets().expect("reads").into_iter().enumerate() {
        let grid = sheet.grid().expect("a grid");
        for cell in grid.cells_in(litchi_xlsb::sheet::Rect::ALL).expect("cells") {
            if fenced {
                black_box(leaf.source_version().expect("current"));
            }
            black_box((position, grid.clone(), cell.row(), cell.column()));
            matches += 1;
        }
    }
    matches
}

/// The same work through a freshly opened XLS leaf: the eager reader's
/// cells from memory, the positional scanner from a file (stopped by its
/// visitor's error once `limit` cells matched).
fn cold_xls_leaf(bytes: &[u8], path: Option<&std::path::Path>, limit: usize) -> usize {
    let mut matches = 0;
    match path {
        Some(path) => {
            let leaf = litchi_xls::SourceBackedWorkbook::open(path).expect("opens");
            let worksheet = leaf
                .worksheet_by_index(0)
                .expect("reads")
                .expect("a worksheet");
            let result = worksheet.visit_cells(|cell| {
                if core_over_1000(cell.value()) {
                    matches += 1;
                    if matches == limit {
                        return Err(litchi_xls::SourceBackedError::InvalidData("stop".into()));
                    }
                }
                Ok(())
            });
            assert!(result.is_ok() || matches == limit, "visits");
        },
        None => {
            let leaf = litchi_xls::Package::from_bytes(bytes).expect("opens");
            let reader = leaf.reader_ref().expect("a reader");
            let worksheet =
                litchi_core::sheet::WorkbookTrait::worksheet_by_index(reader, 0).expect("a sheet");
            for cell in litchi_core::sheet::Worksheet::cells(worksheet) {
                if core_over_1000(litchi_core::sheet::Cell::value(&cell.expect("a cell"))) {
                    matches += 1;
                    if matches == limit {
                        break;
                    }
                }
            }
        },
    }
    matches
}

/// One cold corpus: its bytes, its scratch file for the file backing, and
/// the compiled plan.
struct ColdContext {
    format: &'static str,
    tag: String,
    bytes: Vec<u8>,
    path: Option<std::path::PathBuf>,
    over: Plan<'static>,
}

fn cold_context(format: &'static str, rows: usize, backing: Backing) -> ColdContext {
    let (bytes, file_format) = match format {
        "xlsx" => (workbook::workbook(rows), FileFormat::Xlsx),
        "xlsb" => (xlsb_workbook(rows), FileFormat::Xlsb),
        "xls" => (xls_workbook(rows), FileFormat::Xls),
        other => panic!("unknown cold format {other}"),
    };
    let path =
        (backing == Backing::File).then(|| scratch_file(&format!("cold-{rows}.{format}"), &bytes));
    ColdContext {
        format,
        tag: format!("{rows}_{}", backing.name()),
        bytes,
        path,
        over: plan("cell[value>1000]", file_format),
    }
}

fn cold_routes(context: &ColdContext) -> Vec<Route<'_>> {
    let (bytes, path, over) = (&context.bytes, context.path.as_deref(), &context.over);
    let leaf: fn(&[u8], Option<&std::path::Path>, usize) -> usize = match context.format {
        "xlsx" => cold_xlsx_leaf,
        "xlsb" => cold_xlsb_leaf,
        _ => cold_xls_leaf,
    };
    let reps = if context.tag.starts_with("50000") {
        5
    } else {
        100
    };
    let (format, tag) = (context.format, &context.tag);
    vec![
        Route::new(
            format!("{format}_{tag}_cold_first_cell_value_gt_v1"),
            reps,
            move || cold_unified(bytes, path, over, 1),
            move || leaf(bytes, path, 1),
        ),
        Route::new(
            format!("{format}_{tag}_cold_limit10_cell_value_gt_v1"),
            reps,
            move || cold_unified(bytes, path, over, 10),
            move || leaf(bytes, path, 10),
        ),
    ]
}

/// Warm XLSB routes: a handle opened once, as the XLSX routes do.
struct XlsbContext {
    tag: String,
    file: File,
    leaf: litchi_xlsb::SourceBackedWorkbook,
    rows: usize,
    over: Plan<'static>,
    all: Plan<'static>,
}

fn xlsb_context(rows: usize, backing: Backing) -> XlsbContext {
    let bytes = xlsb_workbook(rows);
    let (file, path) = open_unified(&bytes, &format!("{rows}.xlsb"), backing);
    assert_eq!(file.format(), Some(FileFormat::Xlsb));
    XlsbContext {
        tag: format!("{rows}_{}", backing.name()),
        leaf: xlsb_leaf(&bytes, path.as_deref()),
        file,
        rows,
        over: plan("cell[value>1000]", FileFormat::Xlsb),
        all: plan("cell", FileFormat::Xlsb),
    }
}

fn xlsb_routes(context: &XlsbContext) -> Vec<Route<'_>> {
    let (file, leaf) = (&context.file, &context.leaf);
    let (over, all) = (&context.over, &context.all);
    let tag = &context.tag;
    let middle = context.rows / 2 + 1;
    let get_path = format!("/Data/B{middle}");
    let full_reps = if context.rows > 10_000 { 3 } else { 300 };
    vec![
        Route::new(
            format!("xlsb_{tag}_get_cell_v1"),
            20_000,
            move || {
                black_box(file.get(black_box(get_path.as_str())).expect("resolves"));
                1
            },
            move || {
                let sheet = leaf
                    .worksheet_by_index(black_box(0))
                    .expect("reads")
                    .expect("a sheet");
                let grid = sheet.grid().expect("a grid");
                black_box(grid.cell_at((middle as u32 - 1, 1)).expect("reads"));
                1
            },
        ),
        Route::new(
            format!("xlsb_{tag}_first_cell_value_gt_v1"),
            20_000,
            move || first_match(file, over),
            move || xlsb_leaf_cells(leaf, 1, core_over_1000),
        ),
        Route::new(
            format!("xlsb_{tag}_all_cell_value_gt_v1"),
            full_reps,
            move || count_matches(file, over),
            move || xlsb_leaf_cells(leaf, usize::MAX, core_over_1000),
        ),
        Route::new(
            format!("xlsb_{tag}_all_cell_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsb_leaf_cells(leaf, usize::MAX, |_| true),
        ),
        // Record 0315: fair leaf legs for the same unified query, as record
        // 0280 added for XLSX and PPTX.
        Route::new(
            format!("xlsb_{tag}_all_cell_owned_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsb_leaf_owned(leaf, false),
        ),
        Route::new(
            format!("xlsb_{tag}_all_cell_fenced_v1"),
            full_reps,
            move || count_matches(file, all),
            move || xlsb_leaf_owned(leaf, true),
        ),
    ]
}

// --- DOCX ------------------------------------------------------------------

/// Runs with direct bold in the body, stopping after `limit`.
fn docx_leaf_bold_runs(body: &Blocks, limit: usize) -> usize {
    let mut count = 0;
    let mut stack = vec![body.clone()];
    while let Some(blocks) = stack.pop() {
        for block in blocks.iter() {
            match block.kind() {
                BlockKind::Paragraph => {
                    let paragraph = block.paragraph().expect("a paragraph");
                    for run in paragraph.addressable_runs().expect("runs") {
                        if run.direct_properties().expect("properties").bold == Some(true) {
                            count += 1;
                            if count == limit {
                                return count;
                            }
                        }
                    }
                },
                BlockKind::Table => {
                    let table = block.table().expect("a table");
                    for row in table.rows().expect("rows") {
                        for cell in row.cells().expect("cells") {
                            stack.push(cell.blocks().expect("cell blocks"));
                        }
                    }
                },
                BlockKind::ContentControl => {
                    let control = block.content_control().expect("a control");
                    stack.push(control.blocks().expect("control blocks"));
                },
                _ => {},
            }
        }
    }
    count
}

struct DocxContext {
    label: String,
    file: File,
    leaf: source_backed::Package,
    target: usize,
    bold: Plan<'static>,
}

fn docx_context(label: &'static str, path: &str, backing: Backing) -> DocxContext {
    let bytes = std::fs::read(path).expect("reads the fixture");
    let (file, path) = open_unified(&bytes, &format!("{label}.docx"), backing);
    assert_eq!(file.format(), Some(FileFormat::Docx));
    let leaf = match path {
        Some(path) => source_backed::Package::open(path).expect("opens"),
        None => source_backed::Package::from_bytes(&bytes).expect("opens"),
    };
    let label = format!("{label}_{}", backing.name());
    let paragraphs = leaf
        .document()
        .expect("document")
        .body_blocks()
        .expect("blocks")
        .count(BlockKind::Paragraph);
    DocxContext {
        label,
        file,
        leaf,
        target: paragraphs / 2,
        bold: plan("/body/r[bold=true]", FileFormat::Docx),
    }
}

fn docx_routes(context: &DocxContext) -> Vec<Route<'_>> {
    let (file, leaf, target) = (&context.file, &context.leaf, context.target);
    let bold = &context.bold;
    let label = &context.label;
    let get_path = format!("/body/p[{}]", target + 1);
    vec![
        Route::new(
            format!("docx_{label}_get_paragraph_v1"),
            20_000,
            move || {
                black_box(file.get(black_box(get_path.as_str())).expect("resolves"));
                1
            },
            move || {
                let document = leaf.document().expect("document");
                let paragraph = document
                    .body_paragraph(Position::new(black_box(target)))
                    .expect("reads");
                assert!(paragraph.is_some());
                black_box(paragraph);
                1
            },
        ),
        Route::new(
            format!("docx_{label}_first_run_bold_v1"),
            2_000,
            move || first_match(file, bold),
            move || {
                let document = leaf.document().expect("document");
                docx_leaf_bold_runs(&document.body_blocks().expect("blocks"), 1)
            },
        ),
        Route::new(
            format!("docx_{label}_all_run_bold_v1"),
            20,
            move || count_matches(file, bold),
            move || {
                let document = leaf.document().expect("document");
                docx_leaf_bold_runs(&document.body_blocks().expect("blocks"), usize::MAX)
            },
        ),
    ]
}

// --- PPTX ------------------------------------------------------------------

/// Root and nested `p:sp` shapes of every slide that pass `keep`, stopping
/// after `limit`.
fn pptx_leaf_shapes(
    leaf: &SourceBackedPresentation,
    limit: usize,
    keep: fn(Shape<'_>) -> bool,
) -> usize {
    let mut matches = 0;
    let mut position = 0;
    while let Some(slide) = leaf.slide(position) {
        let scene = slide.scene().expect("reads");
        for handle in scene.iter() {
            let view = handle.view();
            if matches!(view, Shape::Auto(_)) && keep(view) {
                matches += 1;
                if matches == limit {
                    return matches;
                }
            }
        }
        position += 1;
    }
    matches
}

/// Record 0280's fair leaf equivalent of a full `shape` query from the
/// root: the query visits every slide, master and layout scene, so the leaf
/// leg reads them all too (the masters and layouts through the leaf's own
/// catalog reads), keeps each matching shape as an owned handle, and with
/// `fenced` checks the source once per match.
fn pptx_leaf_owned(leaf: &SourceBackedPresentation, fenced: bool) -> usize {
    fn scene_shapes(
        scene: &litchi_pptx::shape::owned::Scene,
        leaf: &SourceBackedPresentation,
        fenced: bool,
    ) -> usize {
        let mut matches = 0;
        for handle in scene.iter() {
            if matches!(handle.view(), Shape::Auto(_)) {
                if fenced {
                    leaf.check_source().expect("current");
                }
                black_box(handle);
                matches += 1;
            }
        }
        matches
    }
    let mut matches = 0;
    let mut position = 0;
    while let Some(slide) = leaf.slide(position) {
        matches += scene_shapes(&slide.scene().expect("reads"), leaf, fenced);
        position += 1;
    }
    let masters = leaf.masters().expect("masters");
    for (index, master) in masters.iter().enumerate() {
        matches += scene_shapes(master.scene(), leaf, fenced);
        let layouts = leaf
            .layouts(Position::new(index))
            .expect("layouts")
            .unwrap_or_default();
        for layout in &layouts {
            matches += scene_shapes(layout.scene(), leaf, fenced);
        }
    }
    matches
}

fn wide(shape: Shape<'_>) -> bool {
    shape
        .bounds()
        .is_some_and(|bounds| bounds.width() >= 3_600_000)
}

fn contains_caseless(text: &str, needle: &str) -> bool {
    text.as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle.as_bytes()))
}

fn revenue(shape: Shape<'_>) -> bool {
    shape
        .text()
        .is_some_and(|text| contains_caseless(text, "Revenue"))
}

fn any_shape(_: Shape<'_>) -> bool {
    true
}

struct PptxContext {
    label: String,
    slides: usize,
    file: File,
    leaf: SourceBackedPresentation,
    plans: [Plan<'static>; 3],
}

fn pptx_context(label: &'static str, slides: usize, boxes: usize, backing: Backing) -> PptxContext {
    let bytes = deck::generated(slides, boxes);
    let (file, path) = open_unified(&bytes, &format!("{label}.pptx"), backing);
    assert_eq!(file.format(), Some(FileFormat::Pptx));
    let leaf = match path {
        Some(path) => SourceBackedPresentation::open(path).expect("opens"),
        None => SourceBackedPresentation::from_reader(Cursor::new(bytes)).expect("opens"),
    };
    PptxContext {
        label: format!("{label}_{}", backing.name()),
        slides,
        file,
        leaf,
        plans: [
            plan("shape[width>=10cm]", FileFormat::Pptx),
            plan("shape:contains(Revenue)", FileFormat::Pptx),
            plan("shape", FileFormat::Pptx),
        ],
    }
}

fn pptx_routes(context: &PptxContext) -> Vec<Route<'_>> {
    let (file, leaf) = (&context.file, &context.leaf);
    let label = &context.label;
    let middle = context.slides / 2;
    let get_path = format!("/slide[{}]/shape[3]", middle + 1);
    let full_reps = if context.slides > 100 { 3 } else { 500 };
    let first_reps = if context.slides > 100 { 2_000 } else { 20_000 };
    let [wide_plan, revenue_plan, all_plan] = &context.plans;
    vec![
        Route::new(
            format!("pptx_{label}_get_shape_v1"),
            20_000,
            move || {
                black_box(file.get(black_box(get_path.as_str())).expect("resolves"));
                1
            },
            move || {
                let slide = leaf.slide(black_box(middle)).expect("a slide");
                let scene = slide.scene().expect("reads");
                let shape = scene
                    .roots()
                    .filter(|handle| matches!(handle.view(), Shape::Auto(_)))
                    .nth(2)
                    .expect("a shape");
                black_box(shape);
                1
            },
        ),
        Route::new(
            format!("pptx_{label}_first_shape_contains_v1"),
            first_reps,
            move || first_match(file, revenue_plan),
            move || pptx_leaf_shapes(leaf, 1, revenue),
        ),
        Route::new(
            format!("pptx_{label}_limit10_shape_width_v1"),
            5_000,
            move || limited(file, wide_plan, 10),
            move || pptx_leaf_shapes(leaf, 10, wide),
        ),
        Route::new(
            format!("pptx_{label}_all_shape_width_v1"),
            full_reps,
            move || count_matches(file, wide_plan),
            move || pptx_leaf_shapes(leaf, usize::MAX, wide),
        ),
        Route::new(
            format!("pptx_{label}_all_shape_contains_v1"),
            full_reps,
            move || count_matches(file, revenue_plan),
            move || pptx_leaf_shapes(leaf, usize::MAX, revenue),
        ),
        Route::new(
            format!("pptx_{label}_all_shape_v1"),
            full_reps,
            move || count_matches(file, all_plan),
            move || pptx_leaf_shapes(leaf, usize::MAX, any_shape),
        ),
        // Record 0280: fair leaf legs for the same unified query.
        Route::new(
            format!("pptx_{label}_all_shape_owned_v1"),
            full_reps,
            move || count_matches(file, all_plan),
            move || pptx_leaf_owned(leaf, false),
        ),
        Route::new(
            format!("pptx_{label}_all_shape_fenced_v1"),
            full_reps,
            move || count_matches(file, all_plan),
            move || pptx_leaf_owned(leaf, true),
        ),
    ]
}

// --- Driver ----------------------------------------------------------------

/// Runs `leg` `calls` times; the profiled region of `profile` mode.
#[inline(never)]
fn profiled(leg: &mut dyn FnMut() -> usize, calls: usize) {
    for _ in 0..calls {
        black_box(leg());
    }
}

fn fixture(relative: &str) -> String {
    format!("{}/../../../{relative}", env!("CARGO_MANIFEST_DIR"))
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mode = args.first().map_or("list", String::as_str);
    let wanted = args.get(1).map_or("all", String::as_str);
    let count: usize = args.get(2).and_then(|text| text.parse().ok()).unwrap_or(0);
    let selected = |name: &str| wanted == "all" || name.starts_with(wanted);
    let family =
        |prefix: &str| wanted == "all" || wanted.starts_with(prefix) || prefix.starts_with(wanted);

    let xlsx: Vec<XlsxContext> = [1_000, 50_000]
        .into_iter()
        .flat_map(|rows| Backing::ALL.map(|backing| (rows, backing)))
        .filter(|(rows, backing)| family(&format!("xlsx_{rows}_{}_", backing.name())))
        .map(|(rows, backing)| xlsx_context(rows, backing))
        .collect();
    let docx: Vec<DocxContext> = [
        ("drawing", "test-data/ooxml/docx/drawing.docx"),
        ("saut", "test-data/poi/test-data/document/saut_page.docx"),
    ]
    .into_iter()
    .flat_map(|(label, path)| Backing::ALL.map(|backing| (label, path, backing)))
    .filter(|(label, _, backing)| family(&format!("docx_{label}_{}_", backing.name())))
    .map(|(label, path, backing)| docx_context(label, &fixture(path), backing))
    .collect();
    let pptx: Vec<PptxContext> = [("small", 10, 8), ("large", 500, 40)]
        .into_iter()
        .flat_map(|(label, slides, boxes)| {
            Backing::ALL.map(|backing| (label, slides, boxes, backing))
        })
        .filter(|(label, _, _, backing)| family(&format!("pptx_{label}_{}_", backing.name())))
        .map(|(label, slides, boxes, backing)| pptx_context(label, slides, boxes, backing))
        .collect();
    let cold: Vec<ColdContext> = ["xlsx", "xlsb", "xls"]
        .into_iter()
        .flat_map(|format| {
            [1_000, 50_000]
                .into_iter()
                .flat_map(move |rows| Backing::ALL.map(|backing| (format, rows, backing)))
        })
        .filter(|(format, rows, backing)| {
            family(&format!("{format}_{rows}_{}_cold", backing.name()))
        })
        .map(|(format, rows, backing)| cold_context(format, rows, backing))
        .collect();
    let xlsb: Vec<XlsbContext> = [1_000, 50_000]
        .into_iter()
        .flat_map(|rows| Backing::ALL.map(|backing| (rows, backing)))
        .filter(|(rows, backing)| family(&format!("xlsb_{rows}_{}_", backing.name())))
        .map(|(rows, backing)| xlsb_context(rows, backing))
        .collect();
    let mut routes: Vec<Route<'_>> = Vec::new();
    for context in &cold {
        routes.extend(cold_routes(context));
    }
    for context in &xlsb {
        routes.extend(xlsb_routes(context));
    }
    for context in &xlsx {
        routes.extend(xlsx_routes(context));
    }
    for context in &docx {
        routes.extend(docx_routes(context));
    }
    for context in &pptx {
        routes.extend(pptx_routes(context));
    }
    let routes = routes.iter_mut().filter(|route| selected(&route.name));
    match mode {
        "list" => {
            for route in routes {
                println!("{}", route.name);
            }
        },
        "time" => {
            let samples = if count == 0 { 15 } else { count };
            for route in routes {
                time(route, samples);
            }
        },
        "profile" => {
            let calls = count.max(1);
            for route in routes {
                black_box((route.unified)());
                profiled(route.unified.as_mut(), calls);
            }
        },
        other => panic!("unknown mode {other}"),
    }
}
