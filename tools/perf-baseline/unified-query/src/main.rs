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
//!   (`crates/litchi-ops/tests/support/pptx_presentation.rs`).
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
    let mut routes: Vec<Route<'_>> = Vec::new();
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
