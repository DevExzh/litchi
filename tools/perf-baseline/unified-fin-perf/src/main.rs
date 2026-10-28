//! Record 0377: performance qualification of the API-debt program (owner
//! decision 0350, contract 0351 §2.6 and §5).
//!
//! One source builds on the pre-program baseline (`--features base`, trunk
//! `839212fa9d`) and on the candidate. It pairs the leaf API with the
//! unified API on the same prepared input, for every claimed format:
//!
//! - `batch`: the record 0337 cases (one atomic `File::batch` of N items
//!   against one hand-written leaf edit), compiled from each tree's own
//!   `unified-batch-closure` modules, so the scenario and its leaf calls are
//!   that tree's;
//! - `set`: the same cases at one operation on the same fixture, the unified
//!   leg calling `File::set` (or `File::remove`) with the item's path and
//!   properties;
//! - `extra` (`extra.rs`): open, first and second save, metadata title set,
//!   DOCX paragraph add, DOC paragraph get, PPTX slide remove/move and shape
//!   remove.
//!
//! Every timed call runs inside [`fin_perf_region`], a non-inlined function,
//! so `valgrind --tool=callgrind --collect-atstart=no
//! --toggle-collect='*fin_perf_region*'` counts only the timed work.
//!
//! Usage:
//!   unified-fin-perf --list [--groups batch,set,extra] [--formats a,b]
//!   unified-fin-perf --abba [--blocks 5] [--rounds 5] [--routes r1,r2] ...
//!   unified-fin-perf --callgrind leaf|unified --routes r [--iterations 3]
//! `--fixture N` (default 1,000) is the number of targets in every fixture;
//! `--count N` (default 1,000) the batch size.

#![allow(clippy::too_many_arguments, dead_code)]

use std::collections::BTreeMap;
use std::hint::black_box;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::{Duration, Instant};

use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, File, ItemResult};

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#[path = "../../unified-batch-closure/src/iwork.rs"]
mod iwork;
#[cfg(any(
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "oth",
    feature = "odm",
    feature = "odf-formula"
))]
#[path = "../../unified-batch-closure/src/odf.rs"]
mod odf;
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
#[path = "../../unified-batch-closure/src/ole.rs"]
mod ole;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
#[path = "../../unified-batch-closure/src/ooxml.rs"]
mod ooxml;
#[cfg(feature = "rtf")]
#[path = "../../unified-batch-closure/src/rtf.rs"]
mod rtf;

mod extra;
#[cfg(feature = "peak-heap")]
mod heap;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

/// How the check compared the two outputs.
#[derive(Clone, Copy, Debug)]
pub enum Equal {
    /// Byte-identical outputs.
    Exact,
    /// Identical decoded ZIP members (entry order or compression differs).
    Members,
    /// The route returns no document (open, get); nothing to compare.
    NotApplicable,
    /// Both legs publish a document, but not the same bytes or members (the
    /// save routes of an eager unified owner may pass entries through where
    /// the leaf re-serializes); reported, not an error.
    Differs,
}

impl Equal {
    fn name(self) -> &'static str {
        match self {
            Self::Exact => "exact_bytes",
            Self::Members => "decoded_members",
            Self::NotApplicable => "not_applicable",
            Self::Differs => "outputs_differ",
        }
    }
}

/// The outcome of the untimed check.
pub struct Checked {
    pub equal: Equal,
    pub leaf_commits: usize,
    pub unified_commits: usize,
}

/// A timed leg: runs its own untimed setup and returns the timed interval.
pub type Leg<'a> = Box<dyn FnMut() -> Result<Duration> + 'a>;

/// One paired scenario at one size.
pub struct Case<'a> {
    pub format: &'static str,
    pub scenario: &'static str,
    pub owner: &'static str,
    pub count: usize,
    pub check: Box<dyn FnMut() -> Result<Checked> + 'a>,
    pub leaf: Leg<'a>,
    pub unified: Leg<'a>,
}

/// True while a callgrind run counts; the warm-up call runs outside the
/// counted function.
static COUNTING: AtomicBool = AtomicBool::new(false);

/// The counted region: the only function the callgrind runs collect.
#[inline(never)]
pub fn fin_perf_region<T>(call: impl FnOnce() -> Result<T>) -> Result<T> {
    black_box(call())
}

/// The uncounted twin. Its body must differ from [`fin_perf_region`]'s:
/// with identical bodies the two are folded into one symbol, and callgrind
/// then counts the warm-up call too (record 0377a).
#[inline(never)]
fn uncounted<T>(call: impl FnOnce() -> Result<T>) -> Result<T> {
    let output = call();
    black_box(COUNTING.load(Ordering::Relaxed));
    black_box(output)
}

pub fn run_timed<T>(call: impl FnOnce() -> Result<T>) -> Result<(T, Duration)> {
    let counting = COUNTING.load(Ordering::Relaxed);
    let start = Instant::now();
    let output = if counting {
        fin_perf_region(call)?
    } else {
        uncounted(call)?
    };
    let elapsed = start.elapsed();
    Ok((output, elapsed))
}

/// Times `call` and drops its output after the clock stops.
pub fn timed<T>(call: impl FnOnce() -> Result<T>) -> Result<Duration> {
    let (output, elapsed) = run_timed(call)?;
    drop(black_box(output));
    Ok(elapsed)
}

/// Like [`timed`], with an untimed input built first.
pub fn timed_with<I, T>(
    input: impl FnOnce() -> Result<I>,
    call: impl FnOnce(I) -> Result<T>,
) -> Result<Duration> {
    let input = input()?;
    let (output, elapsed) = run_timed(|| call(black_box(input)))?;
    drop(black_box(output));
    Ok(elapsed)
}

/// Limits large enough for every case (the leaf legs are unbounded there too).
pub fn limits() -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(100_000)
        .and_then(|b| b.max_expanded_targets(100_000))
        .and_then(|b| b.max_staged_bytes(256 << 20))
        .and_then(|b| b.max_snapshot_bytes(1 << 31))
        .and_then(|b| b.build())
        .expect("limits")
}

/// Whether `unified_leg` builds single-op (`File::set`/`File::remove`) legs.
static SINGLE: AtomicBool = AtomicBool::new(false);

/// The first unified handle each builder registers: the fixture of the
/// `extra` routes for that format.
static FIXTURE: Mutex<Option<&'static File>> = Mutex::new(None);

fn register(file: &'static File) {
    let mut slot = FIXTURE.lock().expect("fixture lock");
    if slot.is_none() {
        *slot = Some(file);
    }
}

/// The accounting of one unified batch: results, changed results, commits.
type Accounting = (File, usize, usize, usize);

fn changed(results: &[ItemResult]) -> usize {
    results
        .iter()
        .filter(|r| matches!(r, ItemResult::Applied(a) if a.changed))
        .count()
}

#[cfg(not(feature = "base"))]
fn batch_once(file: &File, items: Vec<BatchItem<'static>>) -> Result<Accounting> {
    let options = litchi_ops::EditOptions::new().semantic_limits(limits());
    let commit = file.batch_with_options(items, BatchMode::Atomic, options)?;
    let (results, changed) = (commit.value().len(), changed(commit.value()));
    let commits = commit.diagnostics().commits();
    Ok((commit.file().clone(), results, changed, commits))
}

#[cfg(feature = "base")]
fn batch_once(file: &File, items: Vec<BatchItem<'static>>) -> Result<Accounting> {
    let commit = file.batch_with_limits(items, BatchMode::Atomic, &limits())?;
    let (results, changed) = (commit.results.len(), changed(&commit.results));
    let commits = commit.diagnostics.commits();
    Ok((commit.file, results, changed, commits))
}

#[cfg(not(feature = "base"))]
fn batch_timed(file: &File, items: Vec<BatchItem<'static>>) -> Result<impl Sized> {
    let options = litchi_ops::EditOptions::new().semantic_limits(limits());
    Ok(file.batch_with_options(items, BatchMode::Atomic, options)?)
}

#[cfg(feature = "base")]
fn batch_timed(file: &File, items: Vec<BatchItem<'static>>) -> Result<impl Sized> {
    Ok(file.batch_with_limits(items, BatchMode::Atomic, &limits())?)
}

/// One item as a direct `File::set` / `File::remove` call.
fn single(file: &File, item: BatchItem<'static>) -> Result<File> {
    let path = item.path.ok_or("single item without a path")?;
    match item.command.as_ref() {
        "set" => Ok(file.set(path.as_ref(), item.props.ok_or("set without props")?)?),
        "remove" => Ok(file.remove(path.as_ref())?),
        other => Err(format!("single-op leg for `{other}`").into()),
    }
}

/// The unified leg: one atomic batch of a clone of `items`, or in the `set`
/// group the one item as a direct call.
pub fn unified_leg(file: &'static File, items: &'static [BatchItem<'static>]) -> Leg<'static> {
    register(file);
    if SINGLE.load(Ordering::Relaxed) && items.len() == 1 {
        return Box::new(move || timed_with(|| Ok(items[0].clone()), |item| single(file, item)));
    }
    Box::new(move || timed_with(|| Ok(items.to_vec()), |items| batch_timed(file, items)))
}

/// Runs the unified leg once and checks its accounting; returns the output
/// bytes and the number of leaf commits.
pub fn unified_output(file: &File, items: &[BatchItem<'static>]) -> Result<(Vec<u8>, usize)> {
    if SINGLE.load(Ordering::Relaxed) && items.len() == 1 {
        let out = single(file, items[0].clone())?;
        return Ok((out.to_bytes()?, 1));
    }
    let (out, results, changed, commits) = batch_once(file, items.to_vec())?;
    if results != items.len() || changed != items.len() {
        return Err(format!(
            "unified accounting: {} results, {} changed, {} items",
            results,
            changed,
            items.len()
        )
        .into());
    }
    Ok((out.to_bytes()?, commits))
}

/// Compares two outputs: exact bytes, else decoded ZIP members.
pub fn compare(leaf: &[u8], unified: &[u8]) -> Result<Equal> {
    if leaf == unified {
        return Ok(Equal::Exact);
    }
    if leaf.starts_with(b"PK") && unified.starts_with(b"PK") {
        let (a, b) = (members(leaf)?, members(unified)?);
        if a == b {
            return Ok(Equal::Members);
        }
        let differing: Vec<_> = a
            .keys()
            .chain(b.keys())
            .filter(|k| a.get(*k) != b.get(*k))
            .cloned()
            .collect();
        return Err(format!("decoded members differ: {differing:?}").into());
    }
    Err(format!("outputs differ ({} vs {} bytes)", leaf.len(), unified.len()).into())
}

fn members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
    let mut out = BTreeMap::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        let mut data = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut data)?;
        out.insert(entry.name().to_owned(), data);
    }
    Ok(out)
}

/// The standard check: the leaf output's bytes against one unified run.
pub fn check_pair<'a>(
    leaf_bytes: impl FnMut() -> Result<(Vec<u8>, usize)> + 'a,
    file: &'a File,
    items: &'a [BatchItem<'static>],
) -> Box<dyn FnMut() -> Result<Checked> + 'a> {
    let mut leaf_bytes = leaf_bytes;
    Box::new(move || {
        let (leaf, leaf_commits) = leaf_bytes()?;
        let (unified, unified_commits) = unified_output(file, items)?;
        Ok(Checked {
            equal: compare(&leaf, &unified)?,
            leaf_commits,
            unified_commits,
        })
    })
}

/// A `set` item.
pub fn set(
    path: String,
    key: &'static str,
    value: impl Into<litchi_ops::PropValue<'static>>,
) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(litchi_ops::Props::new().with(key, value));
    item
}

/// A `remove` item.
pub fn remove(path: String) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    let n = sorted.len();
    if n % 2 == 1 {
        sorted[n / 2]
    } else {
        (sorted[n / 2 - 1] + sorted[n / 2]) / 2
    }
}

fn arg(name: &str) -> Option<String> {
    let mut args = std::env::args();
    while let Some(a) = args.next() {
        if a == name {
            return args.next();
        }
        if let Some(v) = a.strip_prefix(&format!("{name}=")) {
            return Some(v.to_owned());
        }
    }
    None
}

fn flag(name: &str) -> bool {
    std::env::args().any(|a| a == name)
}

fn list(name: &str) -> Option<Vec<String>> {
    arg(name).map(|v| v.split(',').map(str::to_owned).collect())
}

/// The route name of a case in a group.
fn route(group: &str, case: &Case<'_>) -> String {
    match group {
        "batch" => format!(
            "{}_{}_{}_batch{}",
            case.format, case.owner, case.scenario, case.count
        ),
        "set" => format!("{}_{}_{}_single", case.format, case.owner, case.scenario),
        _ => format!("{}_{}_{}", case.format, case.owner, case.scenario),
    }
}

type Builder = (&'static str, fn(&[usize]) -> Result<Vec<Case<'static>>>);

fn builders() -> Vec<Builder> {
    let mut builders: Vec<Builder> = Vec::new();
    #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
    builders.extend(ooxml::BUILDERS.iter().copied());
    #[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
    builders.extend(ole::BUILDERS.iter().copied());
    #[cfg(feature = "rtf")]
    builders.extend(rtf::BUILDERS.iter().copied());
    #[cfg(any(
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "odg",
        feature = "odc",
        feature = "odi",
        feature = "odb",
        feature = "oth",
        feature = "odm",
        feature = "odf-formula"
    ))]
    builders.extend(odf::BUILDERS.iter().copied());
    #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
    builders.extend(iwork::BUILDERS.iter().copied());
    builders
}

/// Deterministic 64-bit digest (SipHash with zero keys) of a fixture, so the
/// baseline and candidate runs can show they measured the same input.
pub fn digest(bytes: &[u8]) -> String {
    use std::hash::Hasher;
    #[allow(deprecated)]
    let mut hasher = std::hash::SipHasher::new();
    hasher.write(bytes);
    format!("{:016x}", hasher.finish())
}

struct Named {
    group: &'static str,
    route: String,
    case: Case<'static>,
    fixture_digest: Option<String>,
}

fn collect(
    groups: &[String],
    formats: &Option<Vec<String>>,
    fixture: usize,
    count: usize,
) -> Result<Vec<Named>> {
    let wanted = |format: &str| {
        formats
            .as_ref()
            .is_none_or(|f| f.iter().any(|x| x == format))
    };
    let mut out = Vec::new();
    for (format, build) in builders() {
        // `odi` has one fixed frame and is measured only in the batch group.
        if !wanted(format) {
            continue;
        }
        if groups.iter().any(|g| g == "batch") {
            SINGLE.store(false, Ordering::Relaxed);
            let sizes: Vec<usize> = if count == fixture {
                vec![count]
            } else {
                vec![count, fixture]
            };
            for case in build(&sizes)? {
                if case.count == count || case.format == "odi" {
                    let route = route("batch", &case);
                    out.push(Named {
                        group: "batch",
                        route,
                        case,
                        fixture_digest: None,
                    });
                }
            }
        }
        let want_set = groups.iter().any(|g| g == "set");
        let want_extra = groups.iter().any(|g| g == "extra");
        if (want_set || want_extra) && format != "odi" {
            SINGLE.store(true, Ordering::Relaxed);
            *FIXTURE.lock().expect("fixture lock") = None;
            let cases = build(&[1, fixture])?;
            SINGLE.store(false, Ordering::Relaxed);
            let handle = FIXTURE.lock().expect("fixture lock").take();
            if want_set {
                for case in cases {
                    if case.count == 1 {
                        let route = route("set", &case);
                        out.push(Named {
                            group: "set",
                            route,
                            case,
                            fixture_digest: None,
                        });
                    }
                }
            }
            if want_extra {
                if let Some(handle) = handle {
                    let bytes: &'static [u8] = Box::leak(handle.to_bytes()?.into_boxed_slice());
                    let fixture_digest = digest(bytes);
                    for case in extra::cases(format, bytes, fixture)? {
                        let route = route("extra", &case);
                        out.push(Named {
                            group: "extra",
                            route,
                            case,
                            fixture_digest: Some(fixture_digest.clone()),
                        });
                    }
                }
            }
        }
    }
    Ok(out)
}

fn main() -> Result<()> {
    let groups =
        list("--groups").unwrap_or_else(|| vec!["batch".into(), "set".into(), "extra".into()]);
    let formats = list("--formats");
    let routes = list("--routes");
    let fixture: usize = arg("--fixture").map_or(Ok(1_000), |v| v.parse())?;
    let count: usize = arg("--count").map_or(Ok(1_000), |v| v.parse())?;
    let mut named = collect(&groups, &formats, fixture, count)?;
    if let Some(routes) = &routes {
        named.retain(|n| routes.iter().any(|r| r == &n.route));
    }
    if flag("--list") {
        for n in &named {
            println!("{}\t{}", n.group, n.route);
        }
        return Ok(());
    }
    if let Some(leg) = arg("--callgrind") {
        let iterations: usize = arg("--iterations").map_or(Ok(3), |v| v.parse())?;
        let skip_check = flag("--no-check");
        for n in &mut named {
            if !skip_check {
                (n.case.check)().map_err(|e| format!("{}: check failed: {e}", n.route))?;
            }
            let call = |n: &mut Named| -> Result<Duration> {
                if leg == "leaf" {
                    (n.case.leaf)()
                } else {
                    (n.case.unified)()
                }
            };
            call(n)?;
            COUNTING.store(true, Ordering::Relaxed);
            for _ in 0..iterations {
                call(n)?;
            }
            COUNTING.store(false, Ordering::Relaxed);
            println!(
                "{}",
                serde_json::json!({"route": n.route, "leg": leg, "iterations": iterations})
            );
        }
        return Ok(());
    }
    #[cfg(feature = "peak-heap")]
    if let Some(leg) = arg("--peak-heap") {
        for n in &mut named {
            (n.case.check)().map_err(|e| format!("{}: check failed: {e}", n.route))?;
            let call = |n: &mut Named| -> Result<Duration> {
                if leg == "leaf" {
                    (n.case.leaf)()
                } else {
                    (n.case.unified)()
                }
            };
            call(n)?;
            let peak = heap::peak_during(|| call(n))?;
            println!(
                "{}",
                serde_json::json!({"route": n.route, "leg": leg, "peak_heap_bytes": peak})
            );
        }
        return Ok(());
    }
    // Wall clock: `--blocks` blocks of `--rounds` ABBA rounds per route.
    let blocks: usize = arg("--blocks").map_or(Ok(5), |v| v.parse())?;
    let rounds: usize = arg("--rounds").map_or(Ok(5), |v| v.parse())?;
    let mut results = Vec::new();
    for n in &mut named {
        let checked = (n.case.check)().map_err(|e| format!("{}: check failed: {e}", n.route))?;
        (n.case.leaf)()?;
        (n.case.unified)()?;
        let (mut leaf_blocks, mut unified_blocks) = (Vec::new(), Vec::new());
        for _ in 0..blocks {
            let (mut leaf_ns, mut unified_ns) = (Vec::new(), Vec::new());
            for _ in 0..rounds {
                for is_leaf in [true, false, false, true] {
                    if is_leaf {
                        leaf_ns.push((n.case.leaf)()?.as_nanos());
                    } else {
                        unified_ns.push((n.case.unified)()?.as_nanos());
                    }
                }
            }
            leaf_blocks.push(median(&leaf_ns));
            unified_blocks.push(median(&unified_ns));
        }
        let ratios: Vec<f64> = leaf_blocks
            .iter()
            .zip(&unified_blocks)
            .map(|(l, u)| *u as f64 / *l as f64)
            .collect();
        let (l, u) = (median(&leaf_blocks), median(&unified_blocks));
        let lo = ratios.iter().copied().fold(f64::INFINITY, f64::min);
        let hi = ratios.iter().copied().fold(0.0, f64::max);
        eprintln!(
            "{:<52} leaf {:>12.1} µs unified {:>12.1} µs ratio {:.3} [{lo:.3}-{hi:.3}] ({})",
            n.route,
            l as f64 / 1e3,
            u as f64 / 1e3,
            u as f64 / l as f64,
            checked.equal.name(),
        );
        results.push(serde_json::json!({
            "group": n.group,
            "route": n.route,
            "format": n.case.format,
            "scenario": n.case.scenario,
            "owner": n.case.owner,
            "operations": n.case.count,
            "fixture_targets": fixture,
            "fixture_digest": n.fixture_digest,
            "equivalence": checked.equal.name(),
            "leaf_commits": checked.leaf_commits,
            "unified_commits": checked.unified_commits,
            "leaf_block_median_ns": leaf_blocks,
            "unified_block_median_ns": unified_blocks,
            "block_ratios": ratios,
            "leaf_median_ns": l,
            "unified_median_ns": u,
            "ratio": u as f64 / l as f64,
        }));
    }
    println!(
        "{}",
        serde_json::to_string(&serde_json::json!({
            "schema": "litchi.unified-fin-perf.v1",
            "record": "0377",
            "base_api": cfg!(feature = "base"),
            "blocks": blocks,
            "rounds_per_block": rounds,
            "order": "ABBA",
            "cases": results,
        }))?
    );
    Ok(())
}

/// The title the metadata routes set.
pub fn extra_title() -> &'static str {
    "Quarterly review"
}
