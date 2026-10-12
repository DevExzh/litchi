//! Paired native and unified XLS scalar-write timings.
//!
//! Each case authors one numeric workbook, publishes it, and opens that same
//! byte corpus through both the native [`litchi_xls::Package`] and the unified
//! [`litchi_ops::File`] API before timing starts.  A sample changes the
//! existing `A1` number to `-1.0`; the last cell and a complete numeric
//! checksum verify that the rest of the workbook survived the edit.  The
//! immutable source handles are checked after every sample as well.
//!
//! Generation, publication of the input, opening, path construction, and
//! correctness checks are outside the timed regions.  Serialization of the
//! result is also outside timing, so the measurements compare native
//! `Package::edit`/`set_value`/`commit` with `File::set`'s unified adapter
//! transaction.  The fixture is eager and in memory; this benchmark does not
//! measure source-backed I/O and does not require Python.
//!
//! ```text
//! cargo run --release -p litchi-ops --features xls --example xls_write_bench
//! ```
//!
//! Set `LITCHI_XLS_BENCH_SAMPLES` to change the repeat count (default five).

#[cfg(feature = "xls")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_core::sheet::CellValue;
    use litchi_ops::node::Node;
    use litchi_ops::query::{Plan, Query};
    use litchi_ops::{File, OpenOptions, SourceMode};
    use litchi_xls::Package;
    use litchi_xls::cell_values::{Reference, Value};
    use litchi_xls::writer::Builder;
    use serde::Serialize;

    const CASES: &[usize] = &[1, 100, 1_000, 10_000];
    const DEFAULT_SAMPLES: usize = 5;
    const FIRST_VALUE: f64 = 1.0;
    const TARGET_VALUE: f64 = -1.0;
    const FNV_OFFSET: u64 = 14_695_981_039_346_656_037;
    const FNV_PRIME: u64 = 1_099_511_628_211;
    const TARGET_PATH: &str = "/sheet[1]/A1";

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        warmup_passes: usize,
        generation_included: bool,
        opening_included: bool,
        path_construction_included: bool,
        serialization_included: bool,
        verification_included: bool,
        source_backed: bool,
        python_required: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        cells: usize,
        xls_bytes: usize,
        target_value: f64,
        expected_before_checksum: u64,
        expected_after_checksum: u64,
        native_checksum: u64,
        unified_checksum: u64,
        native_output_bytes: usize,
        unified_output_bytes: usize,
        output_bytes_equal: bool,
        original_native_unchanged: bool,
        original_unified_unchanged: bool,
        native_ns: Vec<u128>,
        unified_set_ns: Vec<u128>,
    }

    /// Both immutable owners and all caller-built inputs live outside timing.
    struct Fixture {
        cells: usize,
        original_bytes: Vec<u8>,
        native: Package,
        unified: File,
        references: Vec<Reference>,
        first_path: String,
        last_path: String,
        query_plan: Plan<'static>,
        expected_before_checksum: u64,
        expected_after_checksum: u64,
    }

    struct NativeSample {
        checksum: u64,
        bytes: Vec<u8>,
    }

    struct UnifiedSample {
        checksum: u64,
        bytes: Vec<u8>,
    }

    pub(super) fn run() {
        let samples = std::env::var("LITCHI_XLS_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(DEFAULT_SAMPLES);

        let mut cases = Vec::with_capacity(CASES.len());
        for &cells in CASES {
            let fixture = prepare(cells);

            // The warmup establishes the exact post-edit image shared by both
            // routes. It also proves the source package and source File were
            // left untouched before repeated timing starts.
            let warmup_native = native_edit(&fixture.native, fixture.references[0]);
            let warmup_unified = unified_edit(&fixture.unified);
            let native_warmup = verify_native(&fixture, &warmup_native);
            let unified_warmup = verify_unified(&fixture, &warmup_unified);
            assert_eq!(
                native_warmup.checksum, fixture.expected_after_checksum,
                "native warmup checksum differs for {cells} cells"
            );
            assert_eq!(
                unified_warmup.checksum, fixture.expected_after_checksum,
                "unified warmup checksum differs for {cells} cells"
            );
            assert_eq!(
                native_warmup.bytes, unified_warmup.bytes,
                "native and unified warmup output bytes differ for {cells} cells"
            );
            assert_source_unchanged(&fixture);

            let mut native_ns = Vec::with_capacity(samples);
            let mut unified_set_ns = Vec::with_capacity(samples);
            let mut native_checksum = 0;
            let mut unified_checksum = 0;
            let mut native_output_bytes = 0;
            let mut unified_output_bytes = 0;
            let mut output_bytes_equal = true;

            for _ in 0..samples {
                let start = Instant::now();
                let native = black_box(native_edit(&fixture.native, fixture.references[0]));
                native_ns.push(start.elapsed().as_nanos());
                let native_sample = verify_native(&fixture, &native);
                native_checksum = native_sample.checksum;
                native_output_bytes = native_sample.bytes.len();
                output_bytes_equal &= native_sample.bytes == native_warmup.bytes;

                let start = Instant::now();
                let unified = black_box(unified_edit(&fixture.unified));
                unified_set_ns.push(start.elapsed().as_nanos());
                let unified_sample = verify_unified(&fixture, &unified);
                unified_checksum = unified_sample.checksum;
                unified_output_bytes = unified_sample.bytes.len();
                output_bytes_equal &= unified_sample.bytes == unified_warmup.bytes;
                output_bytes_equal &= native_sample.bytes == unified_sample.bytes;

                assert_source_unchanged(&fixture);
            }

            assert_eq!(
                native_checksum, fixture.expected_after_checksum,
                "native checksum changed during repeat samples"
            );
            assert_eq!(
                unified_checksum, fixture.expected_after_checksum,
                "unified checksum changed during repeat samples"
            );
            assert!(
                output_bytes_equal,
                "edit output bytes changed between paired samples for {cells} cells"
            );

            cases.push(CaseReport {
                cells,
                xls_bytes: fixture.original_bytes.len(),
                target_value: TARGET_VALUE,
                expected_before_checksum: fixture.expected_before_checksum,
                expected_after_checksum: fixture.expected_after_checksum,
                native_checksum,
                unified_checksum,
                native_output_bytes,
                unified_output_bytes,
                output_bytes_equal,
                original_native_unchanged: true,
                original_unified_unchanged: true,
                native_ns,
                unified_set_ns,
            });
        }

        let report = Report {
            schema: "litchi.xls_write_bench.v1",
            samples,
            warmup_passes: 1,
            generation_included: false,
            opening_included: false,
            path_construction_included: false,
            serialization_included: false,
            verification_included: false,
            source_backed: false,
            python_required: false,
            cases,
        };
        serde_json::to_writer_pretty(std::io::stdout().lock(), &report)
            .expect("write XLS benchmark JSON");
        println!();
    }

    fn prepare(cells: usize) -> Fixture {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").expect("add Sheet1");
        let mut references = Vec::with_capacity(cells);
        for row in 0..cells {
            let row = u32::try_from(row).expect("benchmark row fits BIFF8");
            builder
                .write_number(sheet, row, 0, f64::from(row + 1))
                .expect("write benchmark number");
            references.push(Reference::new(row, 0).expect("benchmark reference"));
        }

        let authored = builder.build().expect("build benchmark XLS");
        let original_bytes = authored.to_bytes().expect("publish benchmark XLS");
        let native = Package::from_bytes(&original_bytes).expect("open native XLS");
        let unified = litchi_ops::open_bytes_with(
            original_bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .expect("open unified XLS");
        let query = Query::parse("cell").expect("parse cell query");
        let query_plan = Plan::compile(&query, FileFormat::Xls).expect("compile cell query");
        let first_path = TARGET_PATH.to_owned();
        let last_path = format!("/sheet[1]/A{}", cells);

        let fixture = Fixture {
            cells,
            original_bytes,
            native,
            unified,
            references,
            first_path,
            last_path,
            query_plan,
            expected_before_checksum: expected_checksum(cells, FIRST_VALUE),
            expected_after_checksum: expected_checksum(cells, TARGET_VALUE),
        };

        assert_source_unchanged(&fixture);
        fixture
    }

    fn native_edit(package: &Package, reference: Reference) -> Package {
        let mut edit = package.edit().expect("start native XLS edit");
        edit.set_value(0usize, reference, Value::Number(TARGET_VALUE))
            .expect("set native XLS number");
        edit.commit()
            .expect("commit native XLS edit")
            .into_package()
    }

    fn unified_edit(file: &File) -> File {
        file.set(TARGET_PATH, litchi_ops::props! { "value" => TARGET_VALUE })
            .expect("set unified XLS number")
    }

    fn verify_native(fixture: &Fixture, package: &Package) -> NativeSample {
        let checksum = native_checksum(package, &fixture.references, fixture.cells, TARGET_VALUE);
        let sheet = package
            .workbook()
            .worksheet(0)
            .expect("native output worksheet")
            .expect("native output Sheet1");
        let first = numeric_cell_value(
            &sheet
                .cell_value(fixture.references[0])
                .expect("native output first cell")
                .expect("native output first cell exists"),
        );
        let last = numeric_cell_value(
            &sheet
                .cell_value(fixture.references[fixture.cells - 1])
                .expect("native output last cell")
                .expect("native output last cell exists"),
        );
        assert_eq!(first.to_bits(), TARGET_VALUE.to_bits());
        let expected_last = if fixture.cells == 1 {
            TARGET_VALUE
        } else {
            fixture.cells as f64
        };
        assert_eq!(last.to_bits(), expected_last.to_bits());
        let bytes = package.to_bytes().expect("serialize native output");
        NativeSample { checksum, bytes }
    }

    fn verify_unified(fixture: &Fixture, file: &File) -> UnifiedSample {
        let checksum = unified_checksum(file, &fixture.query_plan, fixture.cells, TARGET_VALUE);
        let first = unified_cell_value(file, &fixture.first_path);
        let last = unified_cell_value(file, &fixture.last_path);
        assert_eq!(first.to_bits(), TARGET_VALUE.to_bits());
        let expected_last = if fixture.cells == 1 {
            TARGET_VALUE
        } else {
            fixture.cells as f64
        };
        assert_eq!(last.to_bits(), expected_last.to_bits());
        let bytes = file.to_bytes().expect("serialize unified output");
        UnifiedSample { checksum, bytes }
    }

    fn assert_source_unchanged(fixture: &Fixture) {
        let native_checksum = native_checksum(
            &fixture.native,
            &fixture.references,
            fixture.cells,
            FIRST_VALUE,
        );
        assert_eq!(native_checksum, fixture.expected_before_checksum);
        assert_eq!(
            fixture.native.to_bytes().expect("serialize native source"),
            fixture.original_bytes
        );

        let unified_checksum = unified_checksum(
            &fixture.unified,
            &fixture.query_plan,
            fixture.cells,
            FIRST_VALUE,
        );
        assert_eq!(unified_checksum, fixture.expected_before_checksum);
        assert_eq!(
            fixture
                .unified
                .to_bytes()
                .expect("serialize unified source"),
            fixture.original_bytes
        );
    }

    fn native_checksum(
        package: &Package,
        references: &[Reference],
        cells: usize,
        first_value: f64,
    ) -> u64 {
        let sheet = package
            .workbook()
            .worksheet(0)
            .expect("native checksum worksheet")
            .expect("native checksum Sheet1");
        let mut checksum = FNV_OFFSET;
        for (index, reference) in references.iter().enumerate() {
            let value = sheet
                .cell_value(*reference)
                .expect("native checksum cell")
                .expect("native checksum cell exists");
            let observed = numeric_cell_value(&value);
            let expected = if index == 0 {
                first_value
            } else {
                (index + 1) as f64
            };
            assert_eq!(observed.to_bits(), expected.to_bits());
            checksum = mix(checksum, observed);
        }
        assert_eq!(references.len(), cells);
        checksum
    }

    fn unified_checksum(file: &File, plan: &Plan<'static>, cells: usize, first_value: f64) -> u64 {
        let mut matches = file.query(plan.clone()).expect("unified checksum query");
        let mut checksum = FNV_OFFSET;
        for index in 0..cells {
            let node = matches
                .next()
                .expect("unified checksum match")
                .expect("unified checksum node");
            let observed = unified_numeric_value(&node);
            let expected = if index == 0 {
                first_value
            } else {
                (index + 1) as f64
            };
            assert_eq!(observed.to_bits(), expected.to_bits());
            checksum = mix(checksum, observed);
        }
        assert!(
            matches.next().is_none(),
            "unified checksum query returned extra cells"
        );
        checksum
    }

    fn unified_cell_value(file: &File, path: &str) -> f64 {
        let node = file.get(path).expect("unified cell lookup");
        unified_numeric_value(&node)
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
        node.view()
            .text()
            .expect("unified cell text")
            .expect("unified benchmark cell text")
            .parse::<f64>()
            .expect("unified benchmark value is numeric")
    }

    fn expected_checksum(cells: usize, first_value: f64) -> u64 {
        let mut checksum = FNV_OFFSET;
        for index in 0..cells {
            let value = if index == 0 {
                first_value
            } else {
                (index + 1) as f64
            };
            checksum = mix(checksum, value);
        }
        checksum
    }

    fn mix(checksum: u64, value: f64) -> u64 {
        checksum.wrapping_mul(FNV_PRIME) ^ value.to_bits()
    }
}

#[cfg(feature = "xls")]
fn main() {
    enabled::run();
}

#[cfg(not(feature = "xls"))]
fn main() {
    eprintln!("xls_write_bench requires the litchi-ops `xls` feature");
}
