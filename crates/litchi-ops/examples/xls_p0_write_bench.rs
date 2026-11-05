//! Paired native and unified timings for the XLS P0 write paths.
//!
//! Each case authors a three-sheet workbook whose `Data` sheet holds `cells`
//! numbers in column A, opens it once natively and once through the unified
//! facade, and times one operation per scenario on both legs:
//!
//! * **sheet_move**: move the last tab first (`Edit::reorder_sheets` against
//!   `File::move_to(.., Placement::Start)`);
//! * **sheet_add**: append a named empty worksheet (`Edit::insert_sheet`
//!   against `File::add("/", sheet)`);
//! * **formula_set**: author `A1*2` into the absent `B1`;
//! * **row_height**: set row 1 to 24 points;
//! * **column_width**: set column C to 12.5 characters;
//! * **cell_format**: right-align `A1`.
//!
//! The native leg is `Package::edit`, the one native operation and `commit`
//! into a new `Package`; the unified leg is the single `File` call, including
//! its target-string parsing and property coercion.  Serialization is outside
//! the timing.  Before timing, both results are serialized and compared, so
//! the report states whether the unified write publishes byte-identical
//! output.  Samples alternate the leg order (ABBA).
//!
//! ```text
//! cargo build --release -p litchi-ops --features xls --example xls_p0_write_bench
//! taskset -c 7 target/release/examples/xls_p0_write_bench
//! ```
//!
//! `LITCHI_XLS_BENCH_SAMPLES` (default nine) and `LITCHI_XLS_BENCH_REPEATS`
//! (operations per sample, default chosen per case) tune the run.
//! `LITCHI_XLS_BENCH_ONLY` (a scenario name), `LITCHI_XLS_BENCH_OWNER`
//! (`eager` or `positional`) and `LITCHI_XLS_BENCH_CELLS` (record 0310) keep
//! only the matching cases, for profiling one of them.

#[cfg(all(feature = "xls", any(unix, windows)))]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::{File, NewNode, OpenOptions, Placement, SourceMode, props};
    use litchi_xls::HorizontalAlignment;
    use litchi_xls::Package;
    use litchi_xls::cell_values::{FormatChange, Reference, Selector};
    use litchi_xls::writer::Builder;
    use serde::Serialize;

    const CASES: &[usize] = &[100, 10_000];
    const DEFAULT_SAMPLES: usize = 9;
    const SCENARIOS: &[&str] = &[
        "sheet_move",
        "sheet_add",
        "formula_set",
        "row_height",
        "column_width",
        "cell_format",
    ];

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        order: &'static str,
        generation_included: bool,
        opening_included: bool,
        serialization_included: bool,
        target_parsing_included_in_unified: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        cells: usize,
        owner: &'static str,
        scenario: &'static str,
        repeats: usize,
        identical_output: bool,
        native_ns: Vec<u128>,
        unified_ns: Vec<u128>,
        median_native_ns: u128,
        median_unified_ns: u128,
        ratio: f64,
    }

    pub(super) fn run() {
        let samples = env("LITCHI_XLS_BENCH_SAMPLES").unwrap_or(DEFAULT_SAMPLES);
        let only = std::env::var("LITCHI_XLS_BENCH_ONLY").ok();
        let only_owner = std::env::var("LITCHI_XLS_BENCH_OWNER").ok();
        let only_cells = env("LITCHI_XLS_BENCH_CELLS");
        let mut cases = Vec::new();
        for &cells in CASES {
            if only_cells.is_some_and(|only| only != cells) {
                continue;
            }
            let bytes = workbook(cells);
            for owner in ["eager", "positional"] {
                if only_owner.as_deref().is_some_and(|only| only != owner) {
                    continue;
                }
                let (native, unified, _guard) = open(&bytes, owner == "positional");
                let repeats =
                    env("LITCHI_XLS_BENCH_REPEATS").unwrap_or(if cells > 1_000 { 20 } else { 200 });
                for &scenario in SCENARIOS {
                    if only.as_deref().is_some_and(|only| only != scenario) {
                        continue;
                    }
                    let identical = native_leg(&native, scenario)
                        .to_bytes()
                        .expect("native bytes")
                        == unified_leg(&unified, scenario)
                            .to_bytes()
                            .expect("unified bytes");
                    let mut native_ns = Vec::new();
                    let mut unified_ns = Vec::new();
                    for sample in 0..samples {
                        let native_time = || {
                            measure(repeats, || {
                                black_box(native_leg(&native, scenario));
                            })
                        };
                        let unified_time = || {
                            measure(repeats, || {
                                black_box(unified_leg(&unified, scenario));
                            })
                        };
                        if sample % 2 == 0 {
                            native_ns.push(native_time());
                            unified_ns.push(unified_time());
                        } else {
                            unified_ns.push(unified_time());
                            native_ns.push(native_time());
                        }
                    }
                    let median_native_ns = median(&native_ns);
                    let median_unified_ns = median(&unified_ns);
                    cases.push(CaseReport {
                        cells,
                        owner,
                        scenario,
                        repeats,
                        identical_output: identical,
                        ratio: median_unified_ns as f64 / median_native_ns as f64,
                        native_ns,
                        unified_ns,
                        median_native_ns,
                        median_unified_ns,
                    });
                }
            }
        }
        let report = Report {
            schema: "litchi.xls_p0_write_bench.v1",
            samples,
            order: "ABBA by sample; nanoseconds per operation",
            generation_included: false,
            opening_included: false,
            serialization_included: false,
            target_parsing_included_in_unified: true,
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
        let data = builder.add_worksheet("Data").expect("sheet");
        builder.add_worksheet("Archive").expect("sheet");
        builder.add_worksheet("Notes").expect("sheet");
        for row in 0..cells {
            builder
                .write_number(data, row as u32, 0, (row + 1) as f64)
                .expect("number");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn open(bytes: &[u8], positional: bool) -> (Package, File, Option<tempfile::TempDir>) {
        if !positional {
            let native = Package::from_bytes(bytes).expect("native eager open");
            let unified = litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("unified eager open");
            return (native, unified, None);
        }
        let guard = tempfile::tempdir().expect("temporary directory");
        let path = guard.path().join("write.xls");
        std::fs::write(&path, bytes).expect("fixture");
        let native = Package::from_source(
            litchi_xls::SourceBackedWorkbook::open(&path).expect("native source"),
        )
        .expect("native positional open");
        let unified =
            litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                .expect("unified positional open");
        (native, unified, Some(guard))
    }

    fn native_leg(package: &Package, scenario: &str) -> Package {
        let mut edit = package.edit().expect("edit");
        let a1 = Reference::new(0, 0).expect("A1");
        match scenario {
            "sheet_move" => edit
                .reorder_sheets(&[
                    Selector::Position(Position::new(2)),
                    Selector::Position(Position::new(0)),
                    Selector::Position(Position::new(1)),
                ])
                .expect("reorder"),
            "sheet_add" => edit
                .insert_sheet(Position::new(3), "Summary")
                .expect("insert"),
            "formula_set" => edit
                .set_formula("Data", Reference::new(0, 1).expect("B1"), "A1*2")
                .expect("formula"),
            "row_height" => edit.set_row_height("Data", 0, 480).expect("height"),
            "column_width" => edit.set_column_width("Data", 2, 3200).expect("width"),
            "cell_format" => {
                let mut change = FormatChange::default();
                change.horizontal = Some(HorizontalAlignment::Right);
                edit.set_cell_format("Data", a1, &change).expect("format");
            },
            other => unreachable!("scenario {other}"),
        }
        edit.commit().expect("commit").into_package()
    }

    fn unified_leg(file: &File, scenario: &str) -> File {
        match scenario {
            "sheet_move" => file.move_to("/Notes", Placement::Start).expect("move"),
            "sheet_add" => {
                file.add("/", NewNode::new("sheet").prop("name", "Summary"))
                    .expect("add")
                    .file
            },
            "formula_set" => file
                .set("/Data/B1", props! { "formula" => "A1*2" })
                .expect("formula"),
            "row_height" => file
                .set("/Data/row[1]", props! { "height" => "24pt" })
                .expect("height"),
            "column_width" => file
                .set("/Data/col[C]", props! { "width" => 12.5 })
                .expect("width"),
            "cell_format" => file
                .set("/Data/A1", props! { "alignment.horizontal" => "right" })
                .expect("format"),
            other => unreachable!("scenario {other}"),
        }
    }

    fn measure(repeats: usize, mut operation: impl FnMut()) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            operation();
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
