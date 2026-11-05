//! Record 0337: one N-item unified `File::batch` against one hand-written
//! leaf edit doing the same work, for every claimed format, at N = 100,
//! 1,000 and 10,000 (the goal's 1k-op batch budget and its linear-scaling
//! check).
//!
//! Each case prepares both owners outside the clock, checks once (outside
//! the clock) that the two legs give the same output, then times both legs
//! in ABBA order. Inputs that a leg consumes (the unified items, a leaf
//! package that an edit consumes) are built or cloned outside the clock;
//! outputs are dropped outside it. Route names are new (`*_batch_closure_v1`).
//!
//! Usage: `unified-batch-closure-baseline [--formats a,b] [--counts 100,1000,10000]
//! [--fixture N] [--rounds N] [--scenarios s1,s2]` prints one JSON document.
//! The fixture holds `max(counts, --fixture)` targets.

#![allow(clippy::too_many_arguments, dead_code)]

use std::collections::BTreeMap;
use std::hint::black_box;
use std::time::{Duration, Instant};

use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, File, ItemResult};

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
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
mod odf;
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
mod ole;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
mod ooxml;
#[cfg(feature = "rtf")]
mod rtf;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

/// How the check compared the two outputs.
#[derive(Clone, Copy, Debug)]
pub enum Equal {
    /// Byte-identical outputs.
    Exact,
    /// Identical decoded ZIP members (entry order or compression differs).
    Members,
}

impl Equal {
    fn name(self) -> &'static str {
        match self {
            Self::Exact => "exact_bytes",
            Self::Members => "decoded_members",
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

/// Times `call` and drops its output after the clock stops.
pub fn timed<T>(call: impl FnOnce() -> Result<T>) -> Result<Duration> {
    let start = Instant::now();
    let output = call()?;
    let elapsed = start.elapsed();
    drop(black_box(output));
    Ok(elapsed)
}

/// Like [`timed`], with an untimed input built first.
pub fn timed_with<I, T>(
    input: impl FnOnce() -> Result<I>,
    call: impl FnOnce(I) -> Result<T>,
) -> Result<Duration> {
    let input = input()?;
    let start = Instant::now();
    let output = call(black_box(input))?;
    let elapsed = start.elapsed();
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

/// The unified leg: one atomic batch of a clone of `items`.
pub fn unified_leg<'a>(file: &'a File, items: &'a [BatchItem<'static>]) -> Leg<'a> {
    let options = litchi_ops::EditOptions::new().semantic_limits(limits());
    Box::new(move || {
        timed_with(
            || Ok(items.to_vec()),
            |items| Ok(file.batch_with_options(items, BatchMode::Atomic, options)?),
        )
    })
}

/// Runs the unified batch once and checks its accounting; returns the output
/// bytes and the number of leaf commits.
pub fn unified_output(file: &File, items: &[BatchItem<'static>]) -> Result<(Vec<u8>, usize)> {
    let options = litchi_ops::EditOptions::new().semantic_limits(limits());
    let commit = file.batch_with_options(items.to_vec(), BatchMode::Atomic, options)?;
    let changed = commit
        .value()
        .iter()
        .filter(|r| matches!(r, ItemResult::Applied(a) if a.changed))
        .count();
    if commit.value().len() != items.len() || changed != items.len() {
        return Err(format!(
            "unified accounting: {} results, {} changed, {} items",
            commit.value().len(),
            changed,
            items.len()
        )
        .into());
    }
    Ok((commit.file().to_bytes()?, commit.diagnostics().commits()))
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

/// The standard check: the leaf output's bytes against one unified batch.
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

fn list(name: &str) -> Option<Vec<String>> {
    arg(name).map(|v| v.split(',').map(str::to_owned).collect())
}

fn main() -> Result<()> {
    let rounds: usize = arg("--rounds").map_or(Ok(5), |v| v.parse())?;
    if rounds < 3 {
        return Err("at least three ABBA rounds".into());
    }
    let counts: Vec<usize> = match list("--counts") {
        Some(v) => v
            .iter()
            .map(|s| s.parse())
            .collect::<std::result::Result<_, _>>()?,
        None => vec![100, 1_000, 10_000],
    };
    // The fixture size (targets in the document); every count runs on the
    // same fixture so the counts compare at a fixed document size.
    let fixture: Option<usize> = arg("--fixture").map(|v| v.parse()).transpose()?;
    let formats = list("--formats");
    let scenarios = list("--scenarios");
    let wanted = |format: &str| {
        formats
            .as_ref()
            .is_none_or(|f| f.iter().any(|x| x == format))
    };

    let mut results = Vec::new();
    let mut builders: Vec<(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)> = Vec::new();
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

    for (format, build) in builders {
        if !wanted(format) {
            continue;
        }
        let mut sizes = counts.clone();
        if let Some(fixture) = fixture {
            sizes.push(fixture);
        }
        let cases = build(&sizes)?;
        for mut case in cases {
            if !counts.contains(&case.count) && case.format != "odi" {
                continue;
            }
            if scenarios
                .as_ref()
                .is_some_and(|s| !s.iter().any(|x| x == case.scenario))
            {
                continue;
            }
            let checked = (case.check)().map_err(|e| {
                format!(
                    "{} {} {}: check failed: {e}",
                    case.format, case.scenario, case.count
                )
            })?;
            // Warm both legs.
            (case.leaf)()?;
            (case.unified)()?;
            let (mut leaf_ns, mut unified_ns) = (Vec::new(), Vec::new());
            for _ in 0..rounds {
                for is_leaf in [true, false, false, true] {
                    if is_leaf {
                        leaf_ns.push((case.leaf)()?.as_nanos());
                    } else {
                        unified_ns.push((case.unified)()?.as_nanos());
                    }
                }
            }
            let (l, u) = (median(&leaf_ns), median(&unified_ns));
            let route = format!("{}_{}_{}", case.format, case.owner, case.scenario);
            eprintln!(
                "{:<8} {:<22} {:>6} leaf {:>12.1} µs unified {:>12.1} µs ratio {:.3} ({}, commits {}/{})",
                case.format,
                case.scenario,
                case.count,
                l as f64 / 1e3,
                u as f64 / 1e3,
                u as f64 / l as f64,
                checked.equal.name(),
                checked.leaf_commits,
                checked.unified_commits
            );
            results.push(serde_json::json!({
                "format": case.format,
                "scenario": case.scenario,
                "owner": case.owner,
                "operations": case.count,
                "fixture_targets": counts.iter().copied().chain(fixture).max().unwrap_or(0),
                "leaf_route": format!("{route}_leaf_batch_closure_v1"),
                "unified_route": format!("{route}_unified_batch_closure_v1"),
                "equivalence": checked.equal.name(),
                "leaf_commits": checked.leaf_commits,
                "unified_commits": checked.unified_commits,
                "leaf_ns": leaf_ns,
                "unified_ns": unified_ns,
                "leaf_median_ns": l,
                "unified_median_ns": u,
                "ratio": u as f64 / l as f64,
            }));
        }
    }
    println!(
        "{}",
        serde_json::to_string(&serde_json::json!({
            "schema": "litchi.unified-batch-closure.v1",
            "record": "0337",
            "rounds": rounds,
            "order": "ABBA",
            "scope": "prepared owners; one unified File::batch (atomic) of N items against one hand-written leaf edit (split only where the leaf's own operation bound forces it); outputs checked equal outside the clock; open, item cloning, serialization and drops untimed",
            "cases": results,
        }))?
    );
    Ok(())
}
