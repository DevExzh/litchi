//! Paired native and unified legacy-PPT direct shape-set timings.
//!
//! The native leg uses one leaf `Edit` per finite 4,096-operation profile and
//! rolls over when that profile is full. Native shape-text batches are split
//! into their 256-item wire limit. The unified closure and command-batch legs
//! retain distinct direct shape sets in one facade transaction whenever the
//! native profile permits it; their diagnostics expose the actual coalesced
//! commit count. Fixture creation, opening, path/payload construction,
//! serialization, and semantic verification are outside timed regions.
//!
//! ```text
//! cargo run --release -p litchi-ops --no-default-features --features ppt \
//!   --example ppt_batch_bench
//! ```
//!
//! Set `LITCHI_PPT_BATCH_BENCH_SAMPLES` to change the paired sample count
//! (default five). The 10,000-shape case opts into an explicit one-GiB
//! retained-snapshot profile because three native 4,096-operation rollovers
//! can exceed the facade's default 256-MiB retention budget. The 1,000-shape
//! case always uses the default profile and is the required default-budget
//! regression point. This is an exploratory probe; it does not replace a
//! frozen performance-harness route.

#[cfg(feature = "ppt")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::transaction::SemanticLimits;
    use litchi_ops::{BatchItem, BatchMode, File, Props, props};
    use litchi_ppt::Package;
    use litchi_ppt::slide_order::ShapeTextReplacement;
    use litchi_ppt::text_edit::Target;
    use litchi_ppt::writer::Builder;
    use serde::Serialize;
    use tempfile::NamedTempFile;

    const SHAPE_COUNTS: &[usize] = &[1, 10, 100, 1_000, 10_000];
    const DEFAULT_SAMPLES: usize = 5;
    const NATIVE_STAGE_LIMIT: usize = 4_096;
    const NATIVE_CHUNK_LIMIT: usize = 256;
    const LARGE_SNAPSHOT_LIMIT: usize = 1 << 30;

    #[derive(Debug, Serialize)]
    struct Report {
        schema: &'static str,
        samples: usize,
        native_stage_limit: usize,
        native_chunk_limit: usize,
        large_snapshot_limit_bytes: usize,
        warmup_passes: usize,
        generation_included: bool,
        opening_included: bool,
        path_construction_included: bool,
        serialization_included: bool,
        verification_included: bool,
        cases: Vec<CaseReport>,
    }

    #[derive(Debug, Serialize)]
    struct CaseReport {
        shapes: usize,
        source: &'static str,
        semantic_snapshot_limit_bytes: usize,
        native_commit_count: usize,
        unified_edit_commit_count: usize,
        unified_batch_commit_count: usize,
        unified_edit_operations: usize,
        unified_batch_operations: usize,
        output_bytes: usize,
        native_output_bytes_equal: bool,
        unified_edit_output_bytes_equal: bool,
        unified_batch_output_bytes_equal: bool,
        original_unified_unchanged: bool,
        filesystem_source_unchanged: bool,
        native_ns: Vec<u128>,
        unified_edit_ns: Vec<u128>,
        unified_batch_ns: Vec<u128>,
    }

    struct Fixture {
        shapes: usize,
        original_bytes: Vec<u8>,
        unified: File,
        paths: Vec<litchi_ops::path::Path<'static>>,
        texts: Vec<String>,
        batch_items: Vec<BatchItem<'static>>,
        limits: SemanticLimits,
        source_file: Option<NamedTempFile>,
    }

    struct NativeRun {
        package: Package,
        elapsed_ns: u128,
        commits: usize,
    }

    struct UnifiedRun {
        file: File,
        elapsed_ns: u128,
        commits: usize,
        operations: usize,
    }

    pub(super) fn run() {
        let samples = std::env::var("LITCHI_PPT_BATCH_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(DEFAULT_SAMPLES);
        let mut cases = Vec::with_capacity(SHAPE_COUNTS.len() * 2);

        for positional in [false, true] {
            for &shapes in SHAPE_COUNTS {
                let fixture = prepare(shapes, positional);
                let warmup_native = native_edit(&fixture);
                let warmup_edit = unified_edit(&fixture);
                let warmup_batch = unified_batch(&fixture);
                let native_bytes = verify_native(&fixture, &warmup_native.package);
                let edit_bytes = verify_unified(&fixture, &warmup_edit.file);
                let batch_bytes = verify_unified(&fixture, &warmup_batch.file);
                assert_eq!(native_bytes, edit_bytes, "native/edit output differs");
                assert_eq!(native_bytes, batch_bytes, "native/batch output differs");
                assert_eq!(
                    warmup_native.commits,
                    expected_native_commits(shapes),
                    "native finite profile count differs for {shapes} shapes"
                );
                assert_eq!(warmup_edit.commits, expected_facade_commits(shapes));
                assert_eq!(warmup_batch.commits, expected_facade_commits(shapes));
                assert_eq!(warmup_edit.operations, shapes);
                assert_eq!(warmup_batch.operations, shapes);
                assert_source_unchanged(&fixture);

                let mut native_ns = Vec::with_capacity(samples);
                let mut unified_edit_ns = Vec::with_capacity(samples);
                let mut unified_batch_ns = Vec::with_capacity(samples);
                let mut native_output_bytes_equal = true;
                let mut unified_edit_output_bytes_equal = true;
                let mut unified_batch_output_bytes_equal = true;
                let mut native_commits = 0;
                let mut edit_commits = 0;
                let mut batch_commits = 0;
                let mut edit_operations = 0;
                let mut batch_operations = 0;

                for sample in 0..samples {
                    if sample % 2 == 0 {
                        record_native(
                            &fixture,
                            &mut native_ns,
                            &mut native_commits,
                            &mut native_output_bytes_equal,
                            &native_bytes,
                        );
                        record_edit(
                            &fixture,
                            &mut unified_edit_ns,
                            &mut edit_commits,
                            &mut edit_operations,
                            &mut unified_edit_output_bytes_equal,
                            &edit_bytes,
                            &native_bytes,
                        );
                        record_batch(
                            &fixture,
                            &mut unified_batch_ns,
                            &mut batch_commits,
                            &mut batch_operations,
                            &mut unified_batch_output_bytes_equal,
                            &batch_bytes,
                            &native_bytes,
                        );
                    } else {
                        record_batch(
                            &fixture,
                            &mut unified_batch_ns,
                            &mut batch_commits,
                            &mut batch_operations,
                            &mut unified_batch_output_bytes_equal,
                            &batch_bytes,
                            &native_bytes,
                        );
                        record_edit(
                            &fixture,
                            &mut unified_edit_ns,
                            &mut edit_commits,
                            &mut edit_operations,
                            &mut unified_edit_output_bytes_equal,
                            &edit_bytes,
                            &native_bytes,
                        );
                        record_native(
                            &fixture,
                            &mut native_ns,
                            &mut native_commits,
                            &mut native_output_bytes_equal,
                            &native_bytes,
                        );
                    }
                    assert_source_unchanged(&fixture);
                }

                cases.push(CaseReport {
                    shapes,
                    source: if positional { "positional" } else { "memory" },
                    semantic_snapshot_limit_bytes: fixture.limits.max_snapshot_bytes(),
                    native_commit_count: native_commits,
                    unified_edit_commit_count: edit_commits,
                    unified_batch_commit_count: batch_commits,
                    unified_edit_operations: edit_operations,
                    unified_batch_operations: batch_operations,
                    output_bytes: native_bytes.len(),
                    native_output_bytes_equal,
                    unified_edit_output_bytes_equal,
                    unified_batch_output_bytes_equal,
                    original_unified_unchanged: true,
                    filesystem_source_unchanged: fixture.source_file.as_ref().is_none_or(|file| {
                        std::fs::read(file.path()).expect("read source") == fixture.original_bytes
                    }),
                    native_ns,
                    unified_edit_ns,
                    unified_batch_ns,
                });
            }
        }

        let report = Report {
            schema: "litchi.ppt_batch_bench.v1",
            samples,
            native_stage_limit: NATIVE_STAGE_LIMIT,
            native_chunk_limit: NATIVE_CHUNK_LIMIT,
            large_snapshot_limit_bytes: LARGE_SNAPSHOT_LIMIT,
            warmup_passes: 1,
            generation_included: false,
            opening_included: false,
            path_construction_included: false,
            serialization_included: false,
            verification_included: false,
            cases,
        };
        serde_json::to_writer_pretty(std::io::stdout().lock(), &report).expect("write JSON");
        println!();
    }

    fn prepare(shapes: usize, positional: bool) -> Fixture {
        let mut builder = Builder::new();
        let slide = builder.add_slide().expect("slide");
        builder
            .set_slide_timing(slide, litchi_ppt::writer::SlideTiming::default())
            .expect("timing");
        let mut paths = Vec::with_capacity(shapes);
        let mut texts = Vec::with_capacity(shapes);
        for index in 0..shapes {
            builder
                .add_textbox(slide, 10 + (index % 100) as i32, 10, 240, 40, "before")
                .expect("textbox");
            paths.push(format!("/slide[1]/shape[{}]", index + 1));
            texts.push(format!("after-{index}"));
        }
        let original_bytes = builder
            .build()
            .expect("build presentation")
            .to_bytes()
            .expect("serialize presentation");
        let source_file = if positional {
            let file = NamedTempFile::new().expect("create positional fixture");
            std::fs::write(file.path(), &original_bytes).expect("write positional fixture");
            Some(file)
        } else {
            None
        };
        let unified = if let Some(file) = source_file.as_ref() {
            litchi_ops::open(file.path()).expect("open positional unified file")
        } else {
            litchi_ops::open_bytes(original_bytes.clone()).expect("open unified file")
        };
        let batch_items = paths
            .iter()
            .zip(&texts)
            .map(|(path, text)| {
                let mut item = BatchItem::new("set");
                item.path = Some(path.clone().into());
                item.props = Some(Props::new().with("text", text.clone()));
                item
            })
            .collect();
        let limits = if shapes > NATIVE_STAGE_LIMIT {
            SemanticLimits::builder()
                .max_snapshot_bytes(LARGE_SNAPSHOT_LIMIT)
                .expect("large snapshot limit")
                .build()
                .expect("large snapshot profile")
        } else {
            SemanticLimits::default()
        };
        Fixture {
            shapes,
            original_bytes,
            unified,
            paths: paths
                .iter()
                .map(|path| {
                    litchi_ops::path::Path::parse(path)
                        .expect("typed path")
                        .into_owned()
                })
                .collect(),
            texts,
            batch_items,
            limits,
            source_file,
        }
    }

    fn native_edit(fixture: &Fixture) -> NativeRun {
        let mut package = if let Some(file) = fixture.source_file.as_ref() {
            Package::open(file.path()).expect("open positional native package")
        } else {
            Package::from_bytes(&fixture.original_bytes).expect("open native package")
        };
        let start = Instant::now();
        let mut commits = 0;
        for (profile_index, profile) in fixture.texts.chunks(NATIVE_STAGE_LIMIT).enumerate() {
            let profile_start = profile_index.saturating_mul(NATIVE_STAGE_LIMIT);
            let mut edit = package.edit().expect("start native edit");
            let mut offset = profile_start;
            for values in profile.chunks(NATIVE_CHUNK_LIMIT) {
                let replacements = values
                    .iter()
                    .enumerate()
                    .map(|(index, text)| {
                        ShapeTextReplacement::new(
                            Target::new(Position::new(0), Position::new(offset + index)),
                            text.as_str(),
                        )
                    })
                    .collect::<Vec<_>>();
                edit.set_shape_texts(&replacements)
                    .expect("stage native shape texts");
                offset += values.len();
            }
            package = edit.commit().expect("commit native edit").into_package();
            commits += 1;
        }
        NativeRun {
            package,
            elapsed_ns: start.elapsed().as_nanos(),
            commits,
        }
    }

    fn unified_edit(fixture: &Fixture) -> UnifiedRun {
        let start = Instant::now();
        let commit = fixture
            .unified
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(fixture.limits),
                |edit| {
                    for (path, text) in fixture.paths.iter().zip(&fixture.texts) {
                        edit.set(path, props! { "text" => text.as_str() })?;
                    }
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect("unified closure edit");
        let commits = commit.diagnostics().commits();
        let operations = commit.diagnostics().operations();
        let (file, patch, ()) = commit.into_parts();
        drop(patch);
        UnifiedRun {
            file,
            elapsed_ns: start.elapsed().as_nanos(),
            commits,
            operations,
        }
    }

    fn unified_batch(fixture: &Fixture) -> UnifiedRun {
        let start = Instant::now();
        let commit = fixture
            .unified
            .batch_with_options(
                fixture.batch_items.clone(),
                BatchMode::Atomic,
                litchi_ops::EditOptions::new().semantic_limits(fixture.limits),
            )
            .expect("unified command batch");
        let commits = commit.diagnostics().commits();
        let operations = commit.diagnostics().operations();
        let (file, patch, results) = commit.into_parts();
        drop(patch);
        drop(results);
        UnifiedRun {
            file,
            elapsed_ns: start.elapsed().as_nanos(),
            commits,
            operations,
        }
    }

    fn record_native(
        fixture: &Fixture,
        timings: &mut Vec<u128>,
        commits: &mut usize,
        equal: &mut bool,
        warmup: &[u8],
    ) {
        let run = black_box(native_edit(fixture));
        let bytes = verify_native(fixture, &run.package);
        timings.push(run.elapsed_ns);
        *commits = run.commits;
        *equal &= bytes == warmup;
    }

    fn record_edit(
        fixture: &Fixture,
        timings: &mut Vec<u128>,
        commits: &mut usize,
        operations: &mut usize,
        equal: &mut bool,
        warmup: &[u8],
        native_warmup: &[u8],
    ) {
        let run = black_box(unified_edit(fixture));
        let bytes = verify_unified(fixture, &run.file);
        timings.push(run.elapsed_ns);
        *commits = run.commits;
        *operations = run.operations;
        *equal &= bytes == warmup && bytes == native_warmup;
    }

    fn record_batch(
        fixture: &Fixture,
        timings: &mut Vec<u128>,
        commits: &mut usize,
        operations: &mut usize,
        equal: &mut bool,
        warmup: &[u8],
        native_warmup: &[u8],
    ) {
        let run = black_box(unified_batch(fixture));
        let bytes = verify_unified(fixture, &run.file);
        timings.push(run.elapsed_ns);
        *commits = run.commits;
        *operations = run.operations;
        *equal &= bytes == warmup && bytes == native_warmup;
    }

    fn verify_native(fixture: &Fixture, package: &Package) -> Vec<u8> {
        let bytes = package.to_bytes().expect("serialize native output");
        verify_output(fixture.shapes, &bytes);
        bytes
    }

    fn verify_unified(fixture: &Fixture, file: &File) -> Vec<u8> {
        let bytes = file.to_bytes().expect("serialize unified output");
        verify_output(fixture.shapes, &bytes);
        bytes
    }

    fn verify_output(shapes: usize, bytes: &[u8]) {
        let package = Package::from_bytes(bytes).expect("reopen output");
        let presentation = package.presentation().expect("read output presentation");
        let slide = presentation
            .slide_at(0)
            .expect("read output slide")
            .expect("output slide exists");
        let scene = slide.shapes().expect("read output scene");
        assert_eq!(scene.len(), shapes);
        for (index, shape) in scene.iter().enumerate() {
            assert_eq!(
                shape.text().expect("output shape text"),
                format!("after-{index}")
            );
        }
    }

    fn assert_source_unchanged(fixture: &Fixture) {
        assert_eq!(
            fixture
                .unified
                .to_bytes()
                .expect("serialize unified source"),
            fixture.original_bytes
        );
        if let Some(file) = fixture.source_file.as_ref() {
            assert_eq!(
                std::fs::read(file.path()).expect("read positional source"),
                fixture.original_bytes
            );
        }
    }

    const fn expected_native_commits(shapes: usize) -> usize {
        shapes.div_ceil(NATIVE_STAGE_LIMIT)
    }

    const fn expected_facade_commits(shapes: usize) -> usize {
        shapes.div_ceil(NATIVE_STAGE_LIMIT)
    }
}

fn main() {
    #[cfg(feature = "ppt")]
    enabled::run();
    #[cfg(not(feature = "ppt"))]
    eprintln!("ppt_batch_bench requires the litchi-ops `ppt` feature");
}
