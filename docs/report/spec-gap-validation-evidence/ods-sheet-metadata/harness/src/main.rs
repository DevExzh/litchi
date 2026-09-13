use std::{
    env,
    hint::black_box,
    time::{Duration, Instant},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Position,
    Profile, Resource,
};
use litchi_ods::sheet_metadata::{self, CellRange, CellSelector, Snapshot};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

#[derive(Clone, Copy, Debug)]
enum Workload {
    Sparse,
    Metadata,
    Parse,
    ParseMetadata,
    Lookup,
    StageOne,
    StageBatch,
    CommitOne,
    CommitBatch,
    Noop,
    EditOne,
    EditBatch,
}

impl Workload {
    fn parse(value: &str) -> Self {
        match value {
            "sparse" => Self::Sparse,
            "metadata" => Self::Metadata,
            "parse" => Self::Parse,
            "parse-metadata" => Self::ParseMetadata,
            "lookup" => Self::Lookup,
            "stage-one" => Self::StageOne,
            "stage-batch" => Self::StageBatch,
            "commit-one" => Self::CommitOne,
            "commit-batch" => Self::CommitBatch,
            "noop" => Self::Noop,
            "edit-one" => Self::EditOne,
            "edit-batch" => Self::EditBatch,
            other => panic!(
                "unknown workload {other:?}; expected sparse, metadata, parse, parse-metadata, lookup, stage-one, stage-batch, commit-one, commit-batch, noop, edit-one, or edit-batch"
            ),
        }
    }

    fn has_metadata_xml(self) -> bool {
        matches!(self, Self::Metadata | Self::ParseMetadata)
    }

    fn includes_parse(self) -> bool {
        matches!(
            self,
            Self::Parse
                | Self::ParseMetadata
                | Self::Sparse
                | Self::Metadata
                | Self::Noop
                | Self::EditOne
                | Self::EditBatch
        )
    }
}

#[derive(Clone, Copy, Debug)]
struct Config {
    workload: Workload,
    sheets: usize,
    rows: usize,
    columns: usize,
    iterations: usize,
    warmups: usize,
}

impl Config {
    fn from_args() -> Self {
        let mut workload = Workload::Sparse;
        let mut sheets = 1;
        let mut rows = 1;
        let mut columns = 1;
        let mut iterations = 10;
        let mut warmups = 2;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" => workload = Workload::parse(&value),
                "--sheets" => sheets = value.parse().expect("sheets must be an integer"),
                "--rows" => rows = value.parse().expect("rows must be an integer"),
                "--columns" => columns = value.parse().expect("columns must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer")
                },
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                other => panic!("unknown option {other}"),
            }
        }
        assert!(sheets > 0 && rows > 0 && columns > 0);
        assert!(iterations > 0);
        Self {
            workload,
            sheets,
            rows,
            columns,
            iterations,
            warmups,
        }
    }
}

fn xml(config: Config) -> String {
    let mut source = String::with_capacity(
        512 + config.sheets * (64 + config.rows * (64 + config.columns * 80)),
    );
    source.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:table=\"{TABLE}\" xmlns:text=\"{TEXT}\" xmlns:xlink=\"{XLINK}\" office:version=\"1.4\"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:spreadsheet>"
    ));
    for sheet in 0..config.sheets {
        source.push_str(&format!("<table:table table:name=\"Sheet{sheet}\">"));
        for row in 0..config.rows {
            source.push_str("<table:table-row>");
            for column in 0..config.columns {
                source.push_str("<table:table-cell office:value-type=\"string\">");
                if config.workload.has_metadata_xml() && column == 0 {
                    source.push_str(&format!(
                        "<table:cell-range-source table:name=\"Import{sheet}_{row}\" table:last-column-spanned=\"1\" table:last-row-spanned=\"1\" xlink:type=\"simple\" xlink:href=\"source.ods#Sheet{sheet}.A1\"/><table:detective><table:highlighted-range table:cell-range-address=\"Sheet{sheet}.A1:Sheet{sheet}.A1\" table:direction=\"from-same-table\" table:contains-error=\"false\"/><table:operation table:name=\"trace-errors\" table:index=\"{row}\"/></table:detective>"
                    ));
                }
                source.push_str("<text:p>x</text:p></table:table-cell>");
            }
            source.push_str("</table:table-row>");
        }
        source.push_str("</table:table>");
    }
    source.push_str("</office:spreadsheet></office:body></office:document-content>");
    source
}

fn context(iteration: usize) -> ExecutionContext {
    let budget = Budget::root(
        format!("ods-sheet-metadata-profile-{iteration}"),
        CoreLimits::for_profile(Profile::TrustedBatch),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("one worker"),
        std::num::NonZeroUsize::new(1).expect("one task"),
        std::num::NonZeroU64::new(256 * 1024 * 1024).expect("in-flight bytes"),
        0,
    )
    .expect("valid execution limits");
    ExecutionContext::new(budget, token, limits)
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    elapsed: Duration,
    memory_final: u64,
    memory_delta: u64,
    work_final: u64,
    work_delta: u64,
    probes: usize,
}

fn used(context: &ExecutionContext, resource: Resource) -> u64 {
    context.budget().used(resource)
}

fn lookup_all_cells(snapshot: &Snapshot, config: Config) -> usize {
    let mut probes = 0usize;
    for sheet in 0..config.sheets {
        for row in 0..config.rows {
            for column in 0..config.columns {
                let selector = CellSelector::by_position(Position::new(sheet), row, column);
                let view = snapshot
                    .cell_metadata(selector)
                    .expect("position selector should resolve")
                    .expect("synthetic cell should exist");
                black_box(view.location());
                probes += 1;
            }
        }
    }
    probes
}

fn stage_cells(edit: &mut litchi_ods::sheet_metadata::Edit, config: Config, count: usize) -> usize {
    let mut probes = 0usize;
    for row in 0..config.rows {
        for column in 0..config.columns {
            if probes >= count {
                return probes;
            }
            let selector = CellSelector::by_position(Position::new(0), row, column);
            let value =
                CellRange::new("Import", "source.ods#Sheet0.A1", 1, 1).expect("valid source");
            edit.set_cell_range_source(selector, value)
                .expect("source metadata should stage");
            probes += 1;
        }
    }
    probes
}

fn profile_once(source: &str, config: Config, iteration: usize) -> Sample {
    let workload = config.workload;
    let context = context(iteration);
    let (snapshot, mut started, mut memory_baseline, mut work_baseline) =
        if workload.includes_parse() {
            let started = Instant::now();
            let snapshot =
                Snapshot::parse_with_context(source, sheet_metadata::Limits::default(), &context)
                    .expect("synthetic XML should parse");
            (snapshot, started, 0, 0)
        } else {
            let snapshot =
                Snapshot::parse_with_context(source, sheet_metadata::Limits::default(), &context)
                    .expect("synthetic XML should parse");
            let memory_baseline = used(&context, Resource::Memory);
            let work_baseline = used(&context, Resource::Work);
            (snapshot, Instant::now(), memory_baseline, work_baseline)
        };
    let mut probes = 0usize;
    let mut retained_edit = None;
    let mut retained_commit = None;
    match workload {
        Workload::Sparse | Workload::Metadata => {
            for sheet in 0..source.matches("<table:table table:name=").count() {
                let selector = CellSelector::by_position(Position::new(sheet), 0, 0);
                let view = snapshot
                    .cell_metadata(selector)
                    .expect("position selector should resolve")
                    .expect("first cell should exist");
                black_box(view.location());
                probes += 1;
            }
            if matches!(workload, Workload::Metadata) {
                let selector = CellSelector::by_name("Sheet0", 0, 0);
                let view = snapshot
                    .cell_metadata(selector)
                    .expect("name selector should resolve")
                    .expect("first metadata cell should exist");
                black_box(view.range_source());
                black_box(view.detective());
                probes += 1;
            }
        },
        Workload::Parse | Workload::ParseMetadata => {
            black_box(snapshot.source_xml().len());
        },
        Workload::Lookup => {
            probes = lookup_all_cells(&snapshot, config);
        },
        Workload::StageOne | Workload::StageBatch | Workload::CommitOne | Workload::CommitBatch => {
            let count = match workload {
                Workload::StageOne | Workload::CommitOne => 1,
                Workload::StageBatch | Workload::CommitBatch => {
                    config.rows.saturating_mul(config.columns)
                },
                _ => unreachable!(),
            };
            let mut edit = snapshot.edit();
            probes = stage_cells(&mut edit, config, count);
            if matches!(workload, Workload::CommitOne | Workload::CommitBatch) {
                // Commit-only samples begin after staging so the timer and
                // deltas describe candidate rendering/reopen/readback.
                memory_baseline = used(&context, Resource::Memory);
                work_baseline = used(&context, Resource::Work);
                started = Instant::now();
                let commit = edit
                    .commit(&context)
                    .expect("metadata commit should succeed");
                black_box(commit.changed());
                retained_commit = Some(commit);
            }
            retained_edit = Some(edit);
        },
        Workload::Noop | Workload::EditOne | Workload::EditBatch => {
            let mut edit = snapshot.edit();
            if matches!(workload, Workload::EditOne | Workload::EditBatch) {
                let count = if matches!(workload, Workload::EditOne) {
                    1
                } else {
                    config.rows.saturating_mul(config.columns)
                };
                probes = stage_cells(&mut edit, config, count);
            }
            let commit = edit
                .commit(&context)
                .expect("metadata commit should succeed");
            black_box(commit.changed());
            retained_commit = Some(commit);
            retained_edit = Some(edit);
        },
    }
    let elapsed = started.elapsed();
    let memory_final = used(&context, Resource::Memory);
    let work_final = used(&context, Resource::Work);
    black_box(snapshot.source_xml().len());
    drop(retained_commit);
    drop(retained_edit);
    Sample {
        elapsed,
        memory_final,
        memory_delta: memory_final.saturating_sub(memory_baseline),
        work_final,
        work_delta: work_final.saturating_sub(work_baseline),
        probes,
    }
}

fn main() {
    let config = Config::from_args();
    let source = xml(config);
    println!(
        "config workload={:?} sheets={} rows={} columns={} source_bytes={} warmups={} iterations={}",
        config.workload,
        config.sheets,
        config.rows,
        config.columns,
        source.len(),
        config.warmups,
        config.iterations
    );
    for iteration in 0..config.warmups {
        let _ = profile_once(&source, config, iteration);
    }
    let mut samples = Vec::with_capacity(config.iterations);
    let mut probes = 0usize;
    for iteration in 0..config.iterations {
        let sample = profile_once(&source, config, iteration + config.warmups);
        probes = probes.saturating_add(sample.probes);
        samples.push(sample);
    }
    let mut elapsed: Vec<_> = samples.iter().map(|sample| sample.elapsed).collect();
    let mut memory_final: Vec<_> = samples.iter().map(|sample| sample.memory_final).collect();
    let mut memory_delta: Vec<_> = samples.iter().map(|sample| sample.memory_delta).collect();
    let mut work_final: Vec<_> = samples.iter().map(|sample| sample.work_final).collect();
    let mut work_delta: Vec<_> = samples.iter().map(|sample| sample.work_delta).collect();
    elapsed.sort_unstable();
    memory_final.sort_unstable();
    memory_delta.sort_unstable();
    work_final.sort_unstable();
    work_delta.sort_unstable();
    let sum_nanos: u128 = elapsed.iter().map(Duration::as_nanos).sum();
    let mean_nanos = sum_nanos / elapsed.len() as u128;
    let p50 = elapsed[elapsed.len() / 2].as_nanos();
    let p95 = elapsed[(elapsed.len() * 95).div_ceil(100).saturating_sub(1)].as_nanos();
    let p99 = elapsed[(elapsed.len() * 99).div_ceil(100).saturating_sub(1)].as_nanos();
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} memory_delta_p50={} memory_delta_max={} memory_final_p50={} memory_final_max={} work_delta_p50={} work_delta_max={} work_final_p50={} work_final_max={} probes={}",
        mean_nanos,
        p50,
        p95,
        p99,
        memory_delta[memory_delta.len() / 2],
        memory_delta[memory_delta.len() - 1],
        memory_final[memory_final.len() / 2],
        memory_final[memory_final.len() - 1],
        work_delta[work_delta.len() / 2],
        work_delta[work_delta.len() - 1],
        work_final[work_final.len() / 2],
        work_final[work_final.len() - 1],
        probes
    );
}
