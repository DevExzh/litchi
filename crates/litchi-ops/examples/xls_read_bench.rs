//! Paired native and unified XLS read timings.
//!
//! Each case authors a one-column numeric workbook and opens it once as an
//! eager byte owner and once as a filesystem positional owner, natively and
//! through the unified facade.  Two scenarios produce the same output on both
//! legs, the display text of every cell:
//!
//! * **get**: one lookup per cell — native `Worksheet::cell_value` plus
//!   number formatting, against `File::get` with a pre-parsed [`Path`] plus
//!   `view().text()`;
//! * **scan**: one traversal — native worksheet iteration plus formatting,
//!   against a compiled `cell` query plus `view().text()`.
//!
//! Samples alternate the leg order (ABBA) to spread drift and warm-cache
//! bias.  Generation, publication, opening, path parsing, plan compilation
//! and the correctness pass are outside the timings; every sample's checksum
//! is compared with the correctness pass.
//!
//! ```text
//! cargo build --release -p litchi-ops --features xls --example xls_read_bench
//! taskset -c 7 target/release/examples/xls_read_bench
//! ```
//!
//! Set `LITCHI_XLS_BENCH_SAMPLES` (default nine) and
//! `LITCHI_XLS_BENCH_REPEATS` (scans per sample, default chosen per case);
//! `LITCHI_XLS_BENCH_CELLS` and `LITCHI_XLS_BENCH_OWNER` (`eager` or
//! `positional`) select one case for profiling.

#[cfg(all(feature = "xls", any(unix, windows)))]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet as _};
    use litchi_ops::path::Path;
    use litchi_ops::query::{Plan, Query};
    use litchi_ops::{File, OpenOptions, SourceMode};
    use litchi_xls::Package;
    use litchi_xls::cell_values::Reference;
    use litchi_xls::writer::Builder;
    use serde::Serialize;

    const CASES: &[usize] = &[1, 100, 1_000, 10_000];
    const DEFAULT_SAMPLES: usize = 9;
    const FNV_OFFSET: u64 = 14_695_981_039_346_656_037;
    const FNV_PRIME: u64 = 1_099_511_628_211;

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        order: &'static str,
        generation_included: bool,
        opening_included: bool,
        path_parsing_included: bool,
        query_compilation_included: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        cells: usize,
        owner: &'static str,
        repeats: usize,
        checksum: u64,
        native_get_ns: Vec<u128>,
        unified_get_ns: Vec<u128>,
        native_scan_ns: Vec<u128>,
        unified_scan_ns: Vec<u128>,
        median_native_get_ns: u128,
        median_unified_get_ns: u128,
        median_native_scan_ns: u128,
        median_unified_scan_ns: u128,
        get_ratio: f64,
        scan_ratio: f64,
    }

    struct Fixture {
        native: Package,
        unified: File,
        references: Vec<Reference>,
        paths: Vec<Path<'static>>,
        plan: Plan<'static>,
        _guard: Option<tempfile::TempDir>,
    }

    pub(super) fn run() {
        let samples = env("LITCHI_XLS_BENCH_SAMPLES").unwrap_or(DEFAULT_SAMPLES);
        let mut cases = Vec::new();
        let only_cells = env("LITCHI_XLS_BENCH_CELLS");
        let only_owner = std::env::var("LITCHI_XLS_BENCH_OWNER").ok();
        for &cells in CASES {
            if only_cells.is_some_and(|only| only != cells) {
                continue;
            }
            let bytes = workbook(cells);
            for owner in ["eager", "positional"] {
                if only_owner.as_deref().is_some_and(|only| only != owner) {
                    continue;
                }
                let fixture = prepare(&bytes, cells, owner == "positional");
                let repeats = env("LITCHI_XLS_BENCH_REPEATS").unwrap_or((10_000 / cells).max(1));
                let checksum = native_get(&fixture);
                assert_eq!(checksum, unified_get(&fixture), "{owner} get checksum");
                assert_eq!(checksum, native_scan(&fixture), "{owner} scan checksum");
                assert_eq!(checksum, unified_scan(&fixture), "{owner} query checksum");
                let mut native_get_ns = Vec::new();
                let mut unified_get_ns = Vec::new();
                let mut native_scan_ns = Vec::new();
                let mut unified_scan_ns = Vec::new();
                for sample in 0..samples {
                    let forward = sample % 2 == 0;
                    let pair = |native: &mut Vec<u128>,
                                unified: &mut Vec<u128>,
                                native_leg: fn(&Fixture) -> u64,
                                unified_leg: fn(&Fixture) -> u64| {
                        if forward {
                            native.push(measure(repeats, checksum, || native_leg(&fixture)));
                            unified.push(measure(repeats, checksum, || unified_leg(&fixture)));
                        } else {
                            unified.push(measure(repeats, checksum, || unified_leg(&fixture)));
                            native.push(measure(repeats, checksum, || native_leg(&fixture)));
                        }
                    };
                    pair(
                        &mut native_get_ns,
                        &mut unified_get_ns,
                        native_get,
                        unified_get,
                    );
                    pair(
                        &mut native_scan_ns,
                        &mut unified_scan_ns,
                        native_scan,
                        unified_scan,
                    );
                }
                let median_native_get_ns = median(&native_get_ns);
                let median_unified_get_ns = median(&unified_get_ns);
                let median_native_scan_ns = median(&native_scan_ns);
                let median_unified_scan_ns = median(&unified_scan_ns);
                cases.push(CaseReport {
                    cells,
                    owner,
                    repeats,
                    checksum,
                    get_ratio: median_unified_get_ns as f64 / median_native_get_ns as f64,
                    scan_ratio: median_unified_scan_ns as f64 / median_native_scan_ns as f64,
                    native_get_ns,
                    unified_get_ns,
                    native_scan_ns,
                    unified_scan_ns,
                    median_native_get_ns,
                    median_unified_get_ns,
                    median_native_scan_ns,
                    median_unified_scan_ns,
                });
            }
        }
        let report = Report {
            schema: "litchi.xls_read_bench.v2",
            samples,
            order: "ABBA by sample; nanoseconds per complete scan",
            generation_included: false,
            opening_included: false,
            path_parsing_included: false,
            query_compilation_included: false,
            cases,
        };
        serde_json::to_writer_pretty(std::io::stdout().lock(), &report).expect("write JSON");
        println!();
    }

    fn env(name: &str) -> Option<usize> {
        std::env::var(name)
            .ok()
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 0)
    }

    fn workbook(cells: usize) -> Vec<u8> {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").expect("sheet");
        for row in 0..cells {
            builder
                .write_number(sheet, row as u32, 0, (row + 1) as f64 + 0.5)
                .expect("number");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn prepare(bytes: &[u8], cells: usize, positional: bool) -> Fixture {
        let (native, unified, guard) = if positional {
            let guard = tempfile::tempdir().expect("temporary directory");
            let path = guard.path().join("read.xls");
            std::fs::write(&path, bytes).expect("fixture");
            let native = Package::from_source(
                litchi_xls::SourceBackedWorkbook::open(&path).expect("native source"),
            )
            .expect("native positional open");
            assert!(native.is_source_backed());
            let unified =
                litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                    .expect("unified positional open");
            (native, unified, Some(guard))
        } else {
            let native = Package::from_bytes(bytes).expect("native eager open");
            let unified = litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("unified eager open");
            (native, unified, None)
        };
        let references = (0..cells)
            .map(|row| Reference::new(row as u32, 0).expect("reference"))
            .collect();
        let paths = (0..cells)
            .map(|row| {
                Path::parse(&format!("/sheet[1]/A{}", row + 1))
                    .expect("path")
                    .into_owned()
            })
            .collect();
        let query = Query::parse("cell").expect("query");
        let plan = Plan::compile(&query, FileFormat::Xls).expect("plan");
        Fixture {
            native,
            unified,
            references,
            paths,
            plan,
            _guard: guard,
        }
    }

    fn display(value: &CellValue) -> String {
        match value {
            CellValue::Float(value) | CellValue::DateTime(value) => value.to_string(),
            CellValue::Int(value) => value.to_string(),
            other => panic!("unexpected benchmark value {other:?}"),
        }
    }

    fn mix(checksum: u64, text: &str) -> u64 {
        text.bytes()
            .fold(checksum, |checksum, byte| {
                (checksum ^ u64::from(byte)).wrapping_mul(FNV_PRIME)
            })
            .wrapping_mul(FNV_PRIME)
    }

    fn native_get(fixture: &Fixture) -> u64 {
        let worksheet = fixture
            .native
            .workbook()
            .worksheet(0)
            .expect("worksheet")
            .expect("first worksheet");
        let mut checksum = FNV_OFFSET;
        for reference in &fixture.references {
            let value = worksheet
                .cell_value(*reference)
                .expect("native cell")
                .expect("stored cell");
            checksum = mix(checksum, &display(&value));
        }
        checksum
    }

    fn unified_get(fixture: &Fixture) -> u64 {
        let mut checksum = FNV_OFFSET;
        for path in &fixture.paths {
            let node = fixture.unified.get(path).expect("unified cell");
            let text = node.view().text().expect("text").expect("cell text");
            checksum = mix(checksum, &text);
        }
        checksum
    }

    fn native_scan(fixture: &Fixture) -> u64 {
        let mut checksum = FNV_OFFSET;
        if let Some(source) = fixture.native.source_workbook() {
            let worksheet = source
                .worksheet_by_index(0)
                .expect("worksheet")
                .expect("first worksheet");
            worksheet
                .visit_cells(|cell| {
                    checksum = mix(checksum, &display(cell.value()));
                    Ok(())
                })
                .expect("visit");
            return checksum;
        }
        let reader = fixture.native.reader_ref().expect("reader");
        let worksheet = WorkbookTrait::worksheet_by_index(reader, 0).expect("worksheet");
        for cell in worksheet.cells() {
            checksum = mix(checksum, &display(cell.expect("cell").value()));
        }
        checksum
    }

    fn unified_scan(fixture: &Fixture) -> u64 {
        let mut checksum = FNV_OFFSET;
        for node in fixture.unified.query(&fixture.plan).expect("query") {
            let node = node.expect("node");
            let text = node.view().text().expect("text").expect("cell text");
            checksum = mix(checksum, &text);
        }
        checksum
    }

    fn measure(repeats: usize, expected: u64, mut operation: impl FnMut() -> u64) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            assert_eq!(black_box(operation()), expected, "checksum changed");
        }
        start.elapsed().as_nanos() / repeats as u128
    }

    fn median(values: &[u128]) -> u128 {
        let mut sorted = values.to_vec();
        sorted.sort_unstable();
        sorted[sorted.len() / 2]
    }
}

fn main() {
    #[cfg(all(feature = "xls", any(unix, windows)))]
    enabled::run();
}
