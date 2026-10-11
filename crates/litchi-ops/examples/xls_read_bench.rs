//! Paired native and unified XLS read timings.
//!
//! The workbook is authored once for each case, then opened once before any
//! timed sample.  The timed regions only read the already-open native
//! worksheet or resolve `/sheet[1]/A1`-style unified paths.  A checksum and a
//! value-by-value verification run happen before timing so a benchmark result
//! cannot silently compare different projections.  JSON is written by
//! default; pass `--csv` for one row per case and sample.
//!
//! ```text
//! cargo run --release -p litchi-ops --features xls --example xls_read_bench
//! cargo run --release -p litchi-ops --features xls --example xls_read_bench -- --csv
//! ```
//!
//! Set `LITCHI_XLS_BENCH_SAMPLES` to change the number of repeat samples (the
//! default is five).  Workbook generation, byte publication, opening, path
//! construction, query-plan compilation and the initial correctness pass are
//! outside the timings.

#[cfg(feature = "xls")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_core::sheet::{CellValue, WorkbookTrait};
    use litchi_ops::node::Node;
    use litchi_ops::query::{Plan, Query};
    use litchi_ops::{File, OpenOptions, SourceMode};
    use litchi_xls::Package;
    use litchi_xls::cell_values::Reference;
    use litchi_xls::package::Worksheet as PackageWorksheet;
    use litchi_xls::writer::Builder;
    use serde::Serialize;

    const CASES: &[usize] = &[1, 100, 1_000, 10_000];
    const DEFAULT_SAMPLES: usize = 5;
    const FNV_OFFSET: u64 = 14_695_981_039_346_656_037;
    const FNV_PRIME: u64 = 1_099_511_628_211;

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        warmup_passes: usize,
        generation_included: bool,
        opening_included: bool,
        path_construction_included: bool,
        query_compilation_included: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        cells: usize,
        xls_bytes: usize,
        native_checksum: u64,
        unified_checksum: u64,
        native_query_checksum: u64,
        unified_query_checksum: u64,
        native_ns: Vec<u128>,
        unified_get_ns: Vec<u128>,
        native_query_ns: Vec<u128>,
        unified_query_ns: Vec<u128>,
    }

    /// The package and facade file stay alive for all repeat samples.  Their
    /// construction is deliberately kept in `prepare`, outside timing.
    struct Fixture {
        package: Package,
        native_sheet: PackageWorksheet,
        unified_file: File,
        references: Vec<Reference>,
        paths: Vec<String>,
        expected: Vec<f64>,
        query_plan: Plan<'static>,
        xls_bytes: usize,
    }

    pub(super) fn run() {
        let samples = std::env::var("LITCHI_XLS_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(DEFAULT_SAMPLES);
        let csv = std::env::args().any(|argument| argument == "--csv");

        let mut cases = Vec::with_capacity(CASES.len());
        for &cells in CASES {
            let fixture = prepare(cells);

            // This pass validates scalar values and checksums before output.
            // It also exercises the already-open lookup paths once; generation,
            // opening, caller path construction and this correctness pass stay
            // outside the timed samples.
            let native_checksum = native_scan(
                &fixture.native_sheet,
                &fixture.references,
                &fixture.expected,
            );
            let unified_checksum =
                unified_scan(&fixture.unified_file, &fixture.paths, &fixture.expected);
            assert_eq!(
                native_checksum, unified_checksum,
                "native and unified XLS checksums differ for {cells} cells"
            );
            let native_query_checksum = native_query_scan(&fixture.package, &fixture.expected);
            let unified_query_checksum = unified_query_scan(
                &fixture.unified_file,
                &fixture.query_plan,
                &fixture.expected,
            );
            assert_eq!(
                native_checksum, native_query_checksum,
                "native direct and traversal checksums differ for {cells} cells"
            );
            assert_eq!(
                native_checksum, unified_query_checksum,
                "native and unified traversal checksums differ for {cells} cells"
            );

            let (native_ns, native_checksums) = measure(samples, || {
                native_scan(
                    &fixture.native_sheet,
                    &fixture.references,
                    &fixture.expected,
                )
            });
            let (unified_ns, unified_checksums) = measure(samples, || {
                unified_scan(&fixture.unified_file, &fixture.paths, &fixture.expected)
            });
            let (native_query_ns, native_query_checksums) = measure(samples, || {
                native_query_scan(&fixture.package, &fixture.expected)
            });
            let (unified_query_ns, unified_query_checksums) = measure(samples, || {
                unified_query_scan(
                    &fixture.unified_file,
                    &fixture.query_plan,
                    &fixture.expected,
                )
            });
            assert!(
                native_checksums
                    .iter()
                    .all(|&checksum| checksum == native_checksum),
                "native checksum changed during repeat samples"
            );
            assert!(
                unified_checksums
                    .iter()
                    .all(|&checksum| checksum == unified_checksum),
                "unified checksum changed during repeat samples"
            );
            assert!(
                native_query_checksums
                    .iter()
                    .all(|&checksum| checksum == native_query_checksum),
                "native traversal checksum changed during repeat samples"
            );
            assert!(
                unified_query_checksums
                    .iter()
                    .all(|&checksum| checksum == unified_query_checksum),
                "unified traversal checksum changed during repeat samples"
            );

            cases.push(CaseReport {
                cells,
                xls_bytes: fixture.xls_bytes,
                native_checksum,
                unified_checksum,
                native_query_checksum,
                unified_query_checksum,
                native_ns,
                unified_get_ns: unified_ns,
                native_query_ns,
                unified_query_ns,
            });
        }

        let report = Report {
            schema: "litchi.xls_read_bench.v1",
            samples,
            warmup_passes: 1,
            generation_included: false,
            opening_included: false,
            path_construction_included: false,
            query_compilation_included: false,
            cases,
        };
        if csv {
            print_csv(&report);
        } else {
            serde_json::to_writer_pretty(std::io::stdout().lock(), &report)
                .expect("write XLS benchmark JSON");
            println!();
        }
    }

    fn prepare(cells: usize) -> Fixture {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").expect("add Sheet1");
        let mut references = Vec::with_capacity(cells);
        let mut paths = Vec::with_capacity(cells);
        let mut expected = Vec::with_capacity(cells);
        for row in 0..cells {
            let value = (row + 1) as f64;
            let row = u32::try_from(row).expect("benchmark row fits BIFF8");
            builder
                .write_number(sheet, row, 0, value)
                .expect("write benchmark number");
            references.push(Reference::new(row, 0).expect("benchmark reference"));
            paths.push(format!("/sheet[1]/A{}", row + 1));
            expected.push(value);
        }
        let package = builder.build().expect("build benchmark XLS");
        let xls_bytes = package.to_bytes().expect("publish benchmark XLS");
        let native_sheet = package
            .workbook()
            .worksheet(0)
            .expect("open native Sheet1")
            .expect("native Sheet1 exists");
        let unified_file = litchi_ops::open_bytes_with(
            xls_bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .expect("open unified XLS");
        let query = Query::parse("cell").expect("parse cell query");
        let query_plan = Plan::compile(&query, FileFormat::Xls).expect("compile cell query");
        Fixture {
            package,
            native_sheet,
            unified_file,
            references,
            paths,
            expected,
            query_plan,
            xls_bytes: xls_bytes.len(),
        }
    }

    fn measure(samples: usize, mut operation: impl FnMut() -> u64) -> (Vec<u128>, Vec<u64>) {
        let mut timings = Vec::with_capacity(samples);
        let mut checksums = Vec::with_capacity(samples);
        for _ in 0..samples {
            let start = Instant::now();
            let checksum = black_box(operation());
            timings.push(start.elapsed().as_nanos());
            checksums.push(checksum);
        }
        (timings, checksums)
    }

    fn native_scan(sheet: &PackageWorksheet, references: &[Reference], expected: &[f64]) -> u64 {
        assert_eq!(references.len(), expected.len());
        let mut checksum = FNV_OFFSET;
        for (reference, &expected) in references.iter().zip(expected) {
            let value = sheet
                .cell_value(*reference)
                .expect("native cell read")
                .expect("native benchmark cell exists");
            let observed = numeric_cell_value(&value);
            assert_eq!(
                observed.to_bits(),
                expected.to_bits(),
                "native value mismatch"
            );
            checksum = mix(checksum, observed);
        }
        checksum
    }

    fn native_query_scan(package: &Package, expected: &[f64]) -> u64 {
        let workbook = package.workbook();
        let reader = workbook.reader_ref().expect("native reader");
        let worksheet =
            WorkbookTrait::worksheet_by_index(reader, 0).expect("native worksheet traversal");
        let mut cells = worksheet.cells();
        let mut checksum = FNV_OFFSET;
        let mut index = 0;
        while let Some(cell) = cells.next() {
            let cell = cell.expect("native cell traversal");
            let expected = *expected
                .get(index)
                .expect("native traversal returned too many cells");
            let observed = numeric_cell_value(cell.value());
            assert_eq!(
                observed.to_bits(),
                expected.to_bits(),
                "native traversal value mismatch at cell {index}"
            );
            checksum = mix(checksum, observed);
            index += 1;
        }
        assert_eq!(
            index,
            expected.len(),
            "native traversal returned the wrong number of cells"
        );
        checksum
    }

    fn unified_scan(file: &File, paths: &[String], expected: &[f64]) -> u64 {
        assert_eq!(paths.len(), expected.len());
        let mut checksum = FNV_OFFSET;
        for (path, &expected) in paths.iter().zip(expected) {
            let node = file.get(path.as_str()).expect("unified cell lookup");
            let value = unified_numeric_value(&node);
            assert_eq!(
                value.to_bits(),
                expected.to_bits(),
                "unified value mismatch"
            );
            checksum = mix(checksum, value);
        }
        checksum
    }

    fn unified_query_scan(file: &File, plan: &Plan<'static>, expected: &[f64]) -> u64 {
        let mut matches = file.query(plan.clone()).expect("unified cell query");
        let mut checksum = FNV_OFFSET;
        for (index, &expected) in expected.iter().enumerate() {
            let node = matches
                .next()
                .expect("unified cell query result")
                .expect("unified cell query match");
            let observed = unified_numeric_value(&node);
            assert_eq!(
                observed.to_bits(),
                expected.to_bits(),
                "unified traversal value mismatch at cell {index}"
            );
            checksum = mix(checksum, observed);
        }
        assert!(
            matches.next().is_none(),
            "unified traversal returned too many cells"
        );
        checksum
    }

    fn numeric_cell_value(value: &CellValue) -> f64 {
        match value {
            CellValue::Int(value) => *value as f64,
            CellValue::Float(value) | CellValue::DateTime(value) => *value,
            CellValue::Formula { cached_value, .. } => cached_value
                .as_deref()
                .map(numeric_cell_value)
                .expect("benchmark formula has a numeric cache"),
            other => panic!("expected numeric native cell, got {other:?}"),
        }
    }

    fn unified_numeric_value(node: &Node<'_>) -> f64 {
        // Keep this conversion in one helper so the adapter's public raw-value
        // accessor can replace it without changing the benchmark protocol.
        // The current stable Node projection is sufficient for this numeric
        // corpus; the correctness pass still rejects a changed value.
        node.view()
            .text()
            .expect("unified cell text")
            .expect("unified benchmark cell text")
            .parse::<f64>()
            .expect("unified benchmark value is numeric")
    }

    fn mix(checksum: u64, value: f64) -> u64 {
        checksum.wrapping_mul(FNV_PRIME) ^ value.to_bits()
    }

    fn print_csv(report: &Report) {
        println!(
            "cells,xls_bytes,sample,native_ns,unified_get_ns,native_query_ns,unified_query_ns,native_checksum,unified_checksum,native_query_checksum,unified_query_checksum"
        );
        for case in &report.cases {
            for sample in 0..report.samples {
                println!(
                    "{},{},{},{},{},{},{},{},{},{},{}",
                    case.cells,
                    case.xls_bytes,
                    sample + 1,
                    case.native_ns[sample],
                    case.unified_get_ns[sample],
                    case.native_query_ns[sample],
                    case.unified_query_ns[sample],
                    case.native_checksum,
                    case.unified_checksum,
                    case.native_query_checksum,
                    case.unified_query_checksum,
                );
            }
        }
    }
}

#[cfg(feature = "xls")]
fn main() {
    enabled::run();
}

#[cfg(not(feature = "xls"))]
fn main() {
    eprintln!("xls_read_bench requires the litchi-ops `xls` feature");
}
