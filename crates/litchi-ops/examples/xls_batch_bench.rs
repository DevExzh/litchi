//! Paired native and unified XLS batch timings for scalar sets and removals.
//!
//! Each case creates one numeric worksheet and prepares all paths, native
//! references, and batch payloads before timing.  Native and unified legs are
//! bounded at the XLS leaf's 4096 staged-operation limit: a larger case
//! flushes the current edit, starts a new one, and retries the item that hit
//! the limit.  Whole, staging, and commit portions of the native leg are
//! reported separately, including every bounded commit.  The unified legs use
//! one `File::edit` and one atomic `File::batch`, and their diagnostics must
//! report the same bounded commit count.  Eager and genuine filesystem
//! positional owners are measured with the same input bytes.
//!
//! Fixture creation, opening, path and payload construction, serialization,
//! and verification are outside timed regions.  Every output is checked by a
//! complete coordinate/value checksum and byte-for-byte equality between the
//! native and both unified routes.  The immutable native owner, unified owner,
//! and filesystem source are checked after every sample.
//!
//! ```text
//! cargo run --release -p litchi-ops --features xls --example xls_batch_bench
//! ```
//!
//! Set `LITCHI_XLS_BENCH_SAMPLES` to change the repeat count (default five).
//! Set `LITCHI_XLS_BENCH_REVERSE=1` to stage coordinates in descending order.

#[cfg(all(feature = "xls", any(unix, windows)))]
mod enabled {
    use std::hint::black_box;
    use std::path::PathBuf;
    use std::sync::Arc;
    use std::time::Instant;

    use litchi_core::FileSource;
    use litchi_core::Position;
    use litchi_ops::{BatchItem, BatchMode, File, OpenOptions, Props, SourceMode, props};
    use litchi_xls::cell_values::{Reference, Selector, Value};
    use litchi_xls::writer::Builder;
    use litchi_xls::{Package, SourceBackedWorkbook};
    use serde::Serialize;
    use tempfile::TempDir;

    const CASES: &[usize] = &[1, 10, 100, 1_000, 10_000];
    const DEFAULT_SAMPLES: usize = 5;
    const NATIVE_STAGE_LIMIT: usize = 4_096;
    const FNV_OFFSET: u64 = 14_695_981_039_346_656_037;
    const FNV_PRIME: u64 = 1_099_511_628_211;

    #[derive(Clone, Copy, Debug, Serialize)]
    #[serde(rename_all = "lowercase")]
    enum Operation {
        Set,
        Remove,
    }

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        staging_order: &'static str,
        warmup_passes: usize,
        native_stage_limit: usize,
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
        operation: Operation,
        source_backed: bool,
        xls_bytes: usize,
        expected_checksum: u64,
        native_checksum: u64,
        unified_edit_checksum: u64,
        unified_batch_checksum: u64,
        native_output_bytes: usize,
        unified_edit_output_bytes: usize,
        unified_batch_output_bytes: usize,
        native_edit_output_bytes_equal: bool,
        unified_edit_output_bytes_equal: bool,
        unified_batch_output_bytes_equal: bool,
        original_native_unchanged: bool,
        original_unified_unchanged: bool,
        filesystem_source_unchanged: bool,
        native_commit_count: usize,
        unified_edit_commit_count: usize,
        unified_batch_commit_count: usize,
        unified_edit_operations: usize,
        unified_batch_operations: usize,
        native_edit_ns: Vec<u128>,
        native_stage_ns: Vec<u128>,
        native_commit_ns: Vec<u128>,
        unified_edit_ns: Vec<u128>,
        unified_batch_ns: Vec<u128>,
    }

    struct Fixture {
        cells: usize,
        operation: Operation,
        original_bytes: Vec<u8>,
        native: Package,
        unified: File,
        references: Vec<Reference>,
        paths: Vec<String>,
        batch_items: Vec<BatchItem<'static>>,
        expected_checksum: u64,
        source_path: Option<PathBuf>,
        _tempdir: Option<TempDir>,
    }

    struct NativeRun {
        package: Package,
        whole_ns: u128,
        stage_ns: u128,
        commit_ns: u128,
        commits: usize,
    }

    struct UnifiedRun {
        file: File,
        commits: usize,
        operations: usize,
    }

    struct Verified {
        checksum: u64,
        bytes: Vec<u8>,
    }

    pub(super) fn run() {
        let samples = std::env::var("LITCHI_XLS_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(DEFAULT_SAMPLES);

        let reverse = std::env::var("LITCHI_XLS_BENCH_REVERSE").as_deref() == Ok("1");
        let mut cases = Vec::with_capacity(CASES.len() * 4);
        for source_backed in [false, true] {
            for operation in [Operation::Set, Operation::Remove] {
                for &cells in CASES {
                    let fixture = prepare(cells, operation, source_backed, reverse);
                    let warmup_native = native_edit(&fixture);
                    let warmup_edit = unified_edit(&fixture);
                    let warmup_batch = unified_batch(&fixture, fixture.batch_items.clone());
                    let native_warmup = verify_native(&fixture, &warmup_native);
                    let edit_warmup = verify_unified(&fixture, &warmup_edit);
                    let batch_warmup = verify_unified(&fixture, &warmup_batch);
                    assert_warmup(
                        &fixture,
                        &warmup_native,
                        &native_warmup,
                        &warmup_edit,
                        &edit_warmup,
                        &warmup_batch,
                        &batch_warmup,
                    );
                    assert_source_unchanged(&fixture);

                    let mut native_edit_ns = Vec::with_capacity(samples);
                    let mut native_stage_ns = Vec::with_capacity(samples);
                    let mut native_commit_ns = Vec::with_capacity(samples);
                    let mut unified_edit_ns = Vec::with_capacity(samples);
                    let mut unified_batch_ns = Vec::with_capacity(samples);
                    let mut native_checksum = 0;
                    let mut native_commit_count = 0;
                    let mut unified_edit_checksum = 0;
                    let mut unified_batch_checksum = 0;
                    let mut native_output_bytes = 0;
                    let mut unified_edit_output_bytes = 0;
                    let mut unified_batch_output_bytes = 0;
                    let mut native_edit_output_bytes_equal = true;
                    let mut unified_edit_output_bytes_equal = true;
                    let mut unified_batch_output_bytes_equal = true;
                    let mut unified_edit_commit_count = 0;
                    let mut unified_batch_commit_count = 0;
                    let mut unified_edit_operations = 0;
                    let mut unified_batch_operations = 0;

                    for _ in 0..samples {
                        let native = black_box(native_edit(&fixture));
                        native_edit_ns.push(native.whole_ns);
                        native_stage_ns.push(native.stage_ns);
                        native_commit_ns.push(native.commit_ns);
                        native_commit_count = native.commits;
                        let native_sample = verify_native(&fixture, &native);
                        native_checksum = native_sample.checksum;
                        native_output_bytes = native_sample.bytes.len();
                        native_edit_output_bytes_equal &=
                            native_sample.bytes == native_warmup.bytes;

                        let start = Instant::now();
                        let edit = black_box(unified_edit(&fixture));
                        unified_edit_ns.push(start.elapsed().as_nanos());
                        let edit_sample = verify_unified(&fixture, &edit);
                        unified_edit_checksum = edit_sample.checksum;
                        unified_edit_output_bytes = edit_sample.bytes.len();
                        unified_edit_output_bytes_equal &= edit_sample.bytes == edit_warmup.bytes;
                        unified_edit_output_bytes_equal &= edit_sample.bytes == native_sample.bytes;
                        unified_edit_commit_count = edit.commits;
                        unified_edit_operations = edit.operations;

                        let batch_items = fixture.batch_items.clone();
                        let start = Instant::now();
                        let batch = black_box(unified_batch(&fixture, batch_items));
                        unified_batch_ns.push(start.elapsed().as_nanos());
                        let batch_sample = verify_unified(&fixture, &batch);
                        unified_batch_checksum = batch_sample.checksum;
                        unified_batch_output_bytes = batch_sample.bytes.len();
                        unified_batch_output_bytes_equal &=
                            batch_sample.bytes == batch_warmup.bytes;
                        unified_batch_output_bytes_equal &=
                            batch_sample.bytes == native_sample.bytes;
                        unified_batch_commit_count = batch.commits;
                        unified_batch_operations = batch.operations;

                        assert_source_unchanged(&fixture);
                    }

                    assert_eq!(native_checksum, fixture.expected_checksum);
                    assert_eq!(unified_edit_checksum, fixture.expected_checksum);
                    assert_eq!(unified_batch_checksum, fixture.expected_checksum);
                    let expected_commits = expected_native_commit_count(cells);
                    assert_eq!(native_commit_count, expected_commits);
                    assert_eq!(unified_edit_commit_count, expected_commits);
                    assert_eq!(unified_batch_commit_count, expected_commits);
                    assert_eq!(unified_edit_operations, cells);
                    assert_eq!(unified_batch_operations, cells);
                    assert!(
                        native_edit_output_bytes_equal,
                        "native output changed across samples for {cells} {operation:?} source={source_backed}"
                    );
                    assert!(
                        unified_edit_output_bytes_equal,
                        "unified edit output differs for {cells} {operation:?} source={source_backed}"
                    );
                    assert!(
                        unified_batch_output_bytes_equal,
                        "unified batch output differs for {cells} {operation:?} source={source_backed}"
                    );

                    cases.push(CaseReport {
                        cells,
                        operation,
                        source_backed,
                        xls_bytes: fixture.original_bytes.len(),
                        expected_checksum: fixture.expected_checksum,
                        native_checksum,
                        unified_edit_checksum,
                        unified_batch_checksum,
                        native_output_bytes,
                        unified_edit_output_bytes,
                        unified_batch_output_bytes,
                        native_edit_output_bytes_equal,
                        unified_edit_output_bytes_equal,
                        unified_batch_output_bytes_equal,
                        original_native_unchanged: true,
                        original_unified_unchanged: true,
                        filesystem_source_unchanged: fixture.source_path.is_none()
                            || fixture.source_path.as_ref().is_some_and(|path| {
                                std::fs::read(path).expect("read source") == fixture.original_bytes
                            }),
                        native_commit_count,
                        unified_edit_commit_count,
                        unified_batch_commit_count,
                        unified_edit_operations,
                        unified_batch_operations,
                        native_edit_ns,
                        native_stage_ns,
                        native_commit_ns,
                        unified_edit_ns,
                        unified_batch_ns,
                    });
                }
            }
        }

        let report = Report {
            schema: "litchi.xls_batch_bench.v1",
            samples,
            staging_order: if reverse { "descending" } else { "ascending" },
            warmup_passes: 1,
            native_stage_limit: NATIVE_STAGE_LIMIT,
            generation_included: false,
            opening_included: false,
            path_construction_included: false,
            serialization_included: false,
            verification_included: false,
            python_required: false,
            cases,
        };
        serde_json::to_writer_pretty(std::io::stdout().lock(), &report)
            .expect("write XLS batch benchmark JSON");
        println!();
    }

    fn prepare(cells: usize, operation: Operation, source_backed: bool, reverse: bool) -> Fixture {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").expect("add Sheet1");
        let mut references = Vec::with_capacity(cells);
        let mut paths = Vec::with_capacity(cells);
        for row in 0..cells {
            let row = row as u32;
            builder
                .write_number(sheet, row, 0, f64::from(row + 1))
                .expect("write benchmark number");
            references.push(Reference::new(row, 0).expect("benchmark reference"));
            paths.push(format!("/sheet[1]/A{}", row + 1));
        }
        let original_bytes = builder
            .build()
            .expect("build benchmark XLS")
            .to_bytes()
            .expect("publish benchmark XLS");
        if reverse {
            references.reverse();
            paths.reverse();
        }
        let expected_checksum = expected_checksum(&references, operation);
        let batch_items = make_batch_items(&paths, operation);
        let (native, unified, source_path, tempdir) = if source_backed {
            let tempdir = tempfile::tempdir().expect("create benchmark tempdir");
            let path = tempdir.path().join("fixture.xls");
            std::fs::write(&path, &original_bytes).expect("write benchmark source");
            let source = FileSource::open(&path).expect("open positional native source");
            let workbook = SourceBackedWorkbook::from_read_at(Arc::new(source))
                .expect("open source-backed native workbook");
            let native = Package::from_source(workbook).expect("adopt native source workbook");
            let unified =
                litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
                    .expect("open source-backed unified workbook");
            (native, unified, Some(path), Some(tempdir))
        } else {
            let native = Package::from_bytes(&original_bytes).expect("open native workbook");
            let unified = litchi_ops::open_bytes_with(
                original_bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("open eager unified workbook");
            (native, unified, None, None)
        };
        let fixture = Fixture {
            cells,
            operation,
            original_bytes,
            native,
            unified,
            references,
            paths,
            batch_items,
            expected_checksum,
            source_path,
            _tempdir: tempdir,
        };
        assert_source_unchanged(&fixture);
        fixture
    }

    fn make_batch_items(paths: &[String], operation: Operation) -> Vec<BatchItem<'static>> {
        paths
            .iter()
            .map(|path| match operation {
                Operation::Set => {
                    let mut item = BatchItem::new("set");
                    item.path = Some(path.clone().into());
                    item.props = Some(Props::new().with("value", -1.0));
                    item
                },
                Operation::Remove => {
                    let mut item = BatchItem::new("remove");
                    item.path = Some(path.clone().into());
                    item
                },
            })
            .collect()
    }

    fn native_edit(fixture: &Fixture) -> NativeRun {
        let mut package = fixture.native.clone();
        let whole_start = Instant::now();
        let mut stage_ns = 0;
        let mut commit_ns = 0;
        let mut commits = 0;
        let mut index = 0;

        while index < fixture.references.len() {
            let stage_start = Instant::now();
            let mut edit = package.edit().expect("start native XLS edit");
            let mut staged = 0;
            let mut hit_limit = false;
            while index < fixture.references.len() {
                let reference = fixture.references[index];
                let result = match fixture.operation {
                    Operation::Set => edit.set_value(
                        Selector::Position(Position::new(0)),
                        reference,
                        Value::Number(-1.0),
                    ),
                    Operation::Remove => {
                        edit.remove_cell(Selector::Position(Position::new(0)), reference)
                    },
                };
                match result {
                    Ok(()) => {
                        staged += 1;
                        index += 1;
                    },
                    Err(error) if error.limit_kind() == Some(litchi_xls::LimitKind::Operations) => {
                        assert!(
                            staged > 0,
                            "native XLS operation limit was reached before any item was staged"
                        );
                        hit_limit = true;
                        break;
                    },
                    Err(error) => panic!("stage native XLS operation: {error}"),
                }
            }

            if hit_limit || index == fixture.references.len() {
                stage_ns += stage_start.elapsed().as_nanos();
                let commit_start = Instant::now();
                package = edit
                    .commit()
                    .expect(if hit_limit {
                        "flush native XLS edit at operation limit"
                    } else {
                        "commit native XLS edit"
                    })
                    .into_package();
                commit_ns += commit_start.elapsed().as_nanos();
                commits += 1;
            }
        }
        let whole_ns = whole_start.elapsed().as_nanos();
        NativeRun {
            package,
            whole_ns,
            stage_ns,
            commit_ns,
            commits,
        }
    }

    fn unified_edit(fixture: &Fixture) -> UnifiedRun {
        let commit = fixture
            .unified
            .edit(|transaction| {
                for path in &fixture.paths {
                    match fixture.operation {
                        Operation::Set => {
                            transaction.set(path.as_str(), props! { "value" => -1.0 })?
                        },
                        Operation::Remove => transaction.remove(path.as_str())?,
                    }
                }
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("unified File::edit");
        UnifiedRun {
            commits: commit.diagnostics().commits(),
            operations: commit.diagnostics().operations(),
            file: commit.into_file(),
        }
    }

    fn unified_batch(fixture: &Fixture, items: Vec<BatchItem<'static>>) -> UnifiedRun {
        let commit = fixture
            .unified
            .batch(items, BatchMode::Atomic)
            .expect("unified File::batch");
        UnifiedRun {
            commits: commit.diagnostics().commits(),
            operations: commit.diagnostics().operations(),
            file: commit.into_file(),
        }
    }

    fn verify_native(fixture: &Fixture, run: &NativeRun) -> Verified {
        let bytes = run.package.to_bytes().expect("serialize native output");
        let snapshot = Package::from_bytes(&bytes)
            .and_then(|package| package.cell_values())
            .expect("parse native output");
        let sheet = snapshot
            .worksheet(Selector::Position(Position::new(0)))
            .expect("resolve native output worksheet")
            .expect("native output Sheet1");
        let mut checksum = FNV_OFFSET;
        let mut count = 0;
        for reference in fixture.references.iter().copied() {
            let cell = sheet.cell(reference).expect("resolve native output cell");
            match fixture.operation {
                Operation::Set => {
                    let cell = cell.expect("native set output cell");
                    assert_eq!(numeric_value(cell.value()).to_bits(), (-1.0_f64).to_bits());
                    checksum = mix(checksum, reference, -1.0);
                    count += 1;
                },
                Operation::Remove => assert!(cell.is_none(), "removed native cell remains"),
            }
        }
        assert_eq!(count, expected_count_for_operations(fixture));
        Verified { checksum, bytes }
    }

    fn verify_unified(fixture: &Fixture, run: &UnifiedRun) -> Verified {
        let mut checksum = FNV_OFFSET;
        let mut count = 0;
        for (index, path) in fixture.paths.iter().enumerate() {
            let text = run
                .file
                .get(path)
                .expect("resolve unified output cell")
                .view()
                .text()
                .expect("read unified output cell")
                .expect("unified output cell text")
                .into_owned();
            match fixture.operation {
                Operation::Set => {
                    assert_eq!(text.parse::<f64>().expect("numeric unified output"), -1.0);
                    checksum = mix(checksum, fixture.references[index], -1.0);
                    count += 1;
                },
                Operation::Remove => assert!(text.is_empty(), "removed unified cell is not blank"),
            }
        }
        assert_eq!(count, expected_count_for_operations(fixture));
        Verified {
            checksum,
            bytes: run.file.to_bytes().expect("serialize unified output"),
        }
    }

    fn assert_warmup(
        fixture: &Fixture,
        native_run: &NativeRun,
        native: &Verified,
        edit_run: &UnifiedRun,
        edit: &Verified,
        batch_run: &UnifiedRun,
        batch: &Verified,
    ) {
        assert_eq!(native.checksum, fixture.expected_checksum);
        assert_eq!(edit.checksum, fixture.expected_checksum);
        assert_eq!(batch.checksum, fixture.expected_checksum);
        assert_eq!(native.bytes, edit.bytes, "native/edit warmup bytes differ");
        assert_eq!(
            native.bytes, batch.bytes,
            "native/batch warmup bytes differ"
        );
        let expected_commits = expected_native_commit_count(fixture.cells);
        assert_eq!(native_run.commits, expected_commits);
        assert_eq!(edit_run.commits, expected_commits);
        assert_eq!(batch_run.commits, expected_commits);
        assert_eq!(edit_run.operations, fixture.cells);
        assert_eq!(batch_run.operations, fixture.cells);
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
                std::fs::read(path).expect("read filesystem source"),
                fixture.original_bytes,
                "filesystem source changed"
            );
        }
    }

    fn expected_count_for_operations(fixture: &Fixture) -> usize {
        match fixture.operation {
            Operation::Set => fixture.cells,
            Operation::Remove => 0,
        }
    }

    fn expected_native_commit_count(cells: usize) -> usize {
        cells.div_ceil(NATIVE_STAGE_LIMIT)
    }

    fn expected_checksum(references: &[Reference], operation: Operation) -> u64 {
        match operation {
            Operation::Set => references
                .iter()
                .copied()
                .fold(FNV_OFFSET, |checksum, reference| {
                    mix(checksum, reference, -1.0)
                }),
            Operation::Remove => FNV_OFFSET,
        }
    }

    fn numeric_value(value: &Value) -> f64 {
        match value {
            Value::Number(value) => *value,
            other => panic!("expected numeric native cell, got {other:?}"),
        }
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
    eprintln!("xls_batch_bench requires the litchi-ops `xls` feature and a filesystem target");
}
