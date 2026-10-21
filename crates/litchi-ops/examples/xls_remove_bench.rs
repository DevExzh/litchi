//! Paired native and unified XLS physical-cell-removal timings.
//!
//! Each case authors the same one-column workbook, removes its middle cell,
//! and checks every remaining coordinate and value.  Native
//! `Package::edit`/`remove_cell`/`commit` is paired with `File::remove` for
//! both eager owners and genuine filesystem positional owners.  Fixture
//! creation, opening, path construction, serialization, and verification are
//! outside the timed regions.
//!
//! ```text
//! cargo run --release -p litchi-ops --features xls --example xls_remove_bench
//! ```
//!
//! Set `LITCHI_XLS_BENCH_SAMPLES` to change the repeat count (default five).

#[cfg(all(feature = "xls", any(unix, windows)))]
mod enabled {
    use std::hint::black_box;
    use std::path::PathBuf;
    use std::sync::Arc;
    use std::time::Instant;

    use litchi_core::FileSource;
    use litchi_core::Position;
    use litchi_ops::{File, OpenOptions, SourceMode};
    use litchi_xls::cell_values::{Reference, Selector, Value};
    use litchi_xls::writer::Builder;
    use litchi_xls::{Package, SourceBackedWorkbook};
    use serde::Serialize;
    use tempfile::TempDir;

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
        serialization_included: bool,
        verification_included: bool,
        python_required: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        cells: usize,
        source_backed: bool,
        xls_bytes: usize,
        target_row: u16,
        expected_checksum: u64,
        native_checksum: u64,
        unified_checksum: u64,
        native_output_bytes: usize,
        unified_output_bytes: usize,
        output_bytes_equal: bool,
        original_native_unchanged: bool,
        original_unified_unchanged: bool,
        filesystem_source_unchanged: bool,
        native_remove_ns: Vec<u128>,
        unified_remove_ns: Vec<u128>,
    }

    /// The source owner and its temporary pathname stay alive for the whole
    /// case, keeping the pathname available for source freshness checks.
    struct Fixture {
        cells: usize,
        target_index: usize,
        original_bytes: Vec<u8>,
        native: Package,
        unified: File,
        references: Vec<Reference>,
        paths: Vec<String>,
        target_path: String,
        expected_checksum: u64,
        source_path: Option<PathBuf>,
        _tempdir: Option<TempDir>,
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

        let mut cases = Vec::with_capacity(CASES.len() * 2);
        for source_backed in [false, true] {
            for &cells in CASES {
                let fixture = prepare(cells, source_backed);
                let warmup_native =
                    native_remove(&fixture.native, fixture.references[fixture.target_index]);
                let warmup_unified = unified_remove(&fixture.unified, &fixture.target_path);
                let native_warmup = verify_native(&fixture, &warmup_native);
                let unified_warmup = verify_unified(&fixture, &warmup_unified);
                assert_eq!(
                    native_warmup.checksum, fixture.expected_checksum,
                    "native warmup checksum differs for {cells} cells, source={source_backed}"
                );
                assert_eq!(
                    unified_warmup.checksum, fixture.expected_checksum,
                    "unified warmup checksum differs for {cells} cells, source={source_backed}"
                );
                assert_eq!(
                    native_warmup.bytes, unified_warmup.bytes,
                    "native and unified warmup output bytes differ for {cells} cells, source={source_backed}"
                );
                assert_source_unchanged(&fixture);

                let mut native_remove_ns = Vec::with_capacity(samples);
                let mut unified_remove_ns = Vec::with_capacity(samples);
                let mut native_checksum = 0;
                let mut unified_checksum = 0;
                let mut native_output_bytes = 0;
                let mut unified_output_bytes = 0;
                let mut output_bytes_equal = true;

                for _ in 0..samples {
                    let start = Instant::now();
                    let native = black_box(native_remove(
                        &fixture.native,
                        fixture.references[fixture.target_index],
                    ));
                    native_remove_ns.push(start.elapsed().as_nanos());
                    let native_sample = verify_native(&fixture, &native);
                    native_checksum = native_sample.checksum;
                    native_output_bytes = native_sample.bytes.len();
                    output_bytes_equal &= native_sample.bytes == native_warmup.bytes;

                    let start = Instant::now();
                    let unified = black_box(unified_remove(&fixture.unified, &fixture.target_path));
                    unified_remove_ns.push(start.elapsed().as_nanos());
                    let unified_sample = verify_unified(&fixture, &unified);
                    unified_checksum = unified_sample.checksum;
                    unified_output_bytes = unified_sample.bytes.len();
                    output_bytes_equal &= unified_sample.bytes == unified_warmup.bytes;
                    output_bytes_equal &= native_sample.bytes == unified_sample.bytes;

                    assert_source_unchanged(&fixture);
                }

                assert_eq!(native_checksum, fixture.expected_checksum);
                assert_eq!(unified_checksum, fixture.expected_checksum);
                assert!(
                    output_bytes_equal,
                    "paired removal output changed for {cells} cells, source={source_backed}"
                );

                cases.push(CaseReport {
                    cells,
                    source_backed,
                    xls_bytes: fixture.original_bytes.len(),
                    target_row: fixture.references[fixture.target_index].row(),
                    expected_checksum: fixture.expected_checksum,
                    native_checksum,
                    unified_checksum,
                    native_output_bytes,
                    unified_output_bytes,
                    output_bytes_equal,
                    original_native_unchanged: true,
                    original_unified_unchanged: true,
                    filesystem_source_unchanged: true,
                    native_remove_ns,
                    unified_remove_ns,
                });
            }
        }

        let report = Report {
            schema: "litchi.xls_remove_bench.v1",
            samples,
            warmup_passes: 1,
            generation_included: false,
            opening_included: false,
            path_construction_included: false,
            serialization_included: false,
            verification_included: false,
            python_required: false,
            cases,
        };
        serde_json::to_writer_pretty(std::io::stdout().lock(), &report)
            .expect("write XLS removal benchmark JSON");
        println!();
    }

    fn prepare(cells: usize, source_backed: bool) -> Fixture {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").expect("add Sheet1");
        let mut references = Vec::with_capacity(cells);
        let mut paths = Vec::with_capacity(cells);
        for row in 0..cells {
            let row = u16::try_from(row).expect("benchmark row fits BIFF8");
            builder
                .write_number(sheet, u32::from(row), 0, f64::from(row) + 1.0)
                .expect("write benchmark number");
            references.push(Reference::new(u32::from(row), 0).expect("benchmark reference"));
            paths.push(format!("/sheet[1]/A{}", usize::from(row) + 1));
        }

        let original_bytes = builder
            .build()
            .expect("build benchmark XLS")
            .to_bytes()
            .expect("publish benchmark XLS");
        let target_index = cells / 2;
        let (native, unified, source_path, tempdir) = if source_backed {
            let tempdir = tempfile::tempdir().expect("create benchmark tempdir");
            let path = tempdir.path().join("fixture.xls");
            std::fs::write(&path, &original_bytes).expect("write benchmark source");
            let source = FileSource::open(&path).expect("open positional benchmark source");
            let workbook = SourceBackedWorkbook::from_read_at(Arc::new(source))
                .expect("open source-backed native workbook");
            let native = Package::from_source(workbook).expect("adopt source-backed package");
            let unified =
                litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                    .expect("open source-backed unified workbook");
            (native, unified, Some(path), Some(tempdir))
        } else {
            let native = Package::from_bytes(&original_bytes).expect("open eager native workbook");
            let unified = litchi_ops::open_bytes_with(
                original_bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("open eager unified workbook");
            (native, unified, None, None)
        };
        let target_path = paths[target_index].clone();
        let expected_checksum = expected_checksum(cells, target_index);
        let fixture = Fixture {
            cells,
            target_index,
            original_bytes,
            native,
            unified,
            references,
            paths,
            target_path,
            expected_checksum,
            source_path,
            _tempdir: tempdir,
        };
        assert_source_unchanged(&fixture);
        fixture
    }

    fn native_remove(package: &Package, reference: Reference) -> Package {
        let mut edit = package.edit().expect("start native XLS removal");
        edit.remove_cell(Selector::Position(Position::new(0)), reference)
            .expect("stage native XLS removal");
        edit.commit()
            .expect("commit native XLS removal")
            .into_package()
    }

    fn unified_remove(file: &File, path: &str) -> File {
        file.remove(path).expect("remove unified XLS cell")
    }

    fn verify_native(fixture: &Fixture, package: &Package) -> NativeSample {
        let bytes = package.to_bytes().expect("serialize native output");
        let snapshot = Package::from_bytes(&bytes.clone())
            .and_then(|package| package.cell_values())
            .expect("parse native output");
        let sheet = snapshot
            .worksheet(Selector::Position(Position::new(0)))
            .expect("resolve native output worksheet")
            .expect("native output Sheet1");
        assert_eq!(sheet.cells().count(), fixture.cells - 1);

        let mut checksum = FNV_OFFSET;
        for (index, reference) in fixture.references.iter().copied().enumerate() {
            let cell = sheet.cell(reference).expect("resolve native output cell");
            if index == fixture.target_index {
                assert!(cell.is_none(), "removed native cell remains stored");
                continue;
            }
            let cell = cell.expect("native remaining cell exists");
            let observed = numeric_value(cell.value());
            let expected = (index + 1) as f64;
            assert_eq!(observed.to_bits(), expected.to_bits());
            checksum = mix(checksum, reference, observed);
        }
        NativeSample { checksum, bytes }
    }

    fn verify_unified(fixture: &Fixture, file: &File) -> UnifiedSample {
        let mut checksum = FNV_OFFSET;
        let matches = file.query("cell").expect("query unified remaining cells");
        let mut count = 0;
        for node in matches {
            node.expect("read unified remaining cell");
            count += 1;
        }
        assert_eq!(count, fixture.cells - 1);

        for (index, path) in fixture.paths.iter().enumerate() {
            let node = file.get(path).expect("resolve unified output cell");
            let text = node
                .view()
                .text()
                .expect("read unified output cell")
                .expect("unified output cell text")
                .into_owned();
            if index == fixture.target_index {
                assert!(text.is_empty(), "removed unified cell is not blank");
                continue;
            }
            let observed = text.parse::<f64>().expect("unified value is numeric");
            let expected = (index + 1) as f64;
            assert_eq!(observed.to_bits(), expected.to_bits());
            checksum = mix(checksum, fixture.references[index], observed);
        }
        let bytes = file.to_bytes().expect("serialize unified output");
        UnifiedSample { checksum, bytes }
    }

    fn assert_source_unchanged(fixture: &Fixture) {
        assert_eq!(
            fixture.native.to_bytes().expect("serialize native source"),
            fixture.original_bytes,
            "native source owner changed"
        );
        assert_eq!(
            fixture
                .unified
                .to_bytes()
                .expect("serialize unified source"),
            fixture.original_bytes,
            "unified source owner changed"
        );
        if let Some(path) = &fixture.source_path {
            assert_eq!(
                std::fs::read(path).expect("read benchmark source"),
                fixture.original_bytes,
                "filesystem source changed"
            );
        }
    }

    fn numeric_value(value: &Value) -> f64 {
        match value {
            Value::Number(value) => *value,
            other => panic!("expected numeric native cell, got {other:?}"),
        }
    }

    fn expected_checksum(cells: usize, target_index: usize) -> u64 {
        let mut checksum = FNV_OFFSET;
        for index in 0..cells {
            if index != target_index {
                checksum = mix(
                    checksum,
                    Reference::new(index as u32, 0).expect("checksum reference"),
                    (index + 1) as f64,
                );
            }
        }
        checksum
    }

    fn mix(checksum: u64, reference: Reference, value: f64) -> u64 {
        let coordinate = (u64::from(reference.row()) << 8) | u64::from(reference.column());
        checksum
            .wrapping_mul(FNV_PRIME)
            .wrapping_add(coordinate)
            .wrapping_mul(FNV_PRIME)
            ^ value.to_bits()
    }
}

#[cfg(all(feature = "xls", any(unix, windows)))]
fn main() {
    enabled::run();
}

#[cfg(any(not(feature = "xls"), not(any(unix, windows))))]
fn main() {
    eprintln!("xls_remove_bench requires the litchi-ops `xls` feature and a filesystem target");
}
