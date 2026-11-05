//! Paired prepared-owner DOCX formatting batches; opening and saving are untimed.
use litchi_core::{FileFormat, Position};
use litchi_ops::{BatchItem, BatchMode, NewNode, OpenOptions, Props, SourceMode};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

fn leaf(base: &litchi_docx::Package, count: usize) -> Result<Vec<litchi_docx::document::Commit>> {
    let mut commits = Vec::new();
    let mut edit = base.edit()?;
    for index in 0..count {
        if edit.remaining_operations() == 0 {
            commits.push(edit.commit()?);
            edit = commits.last().unwrap().package().edit()?;
        }
        edit.format_run(Position::new(0), Position::new(0), |properties| {
            properties.bold = Some(index % 2 == 0);
        })?;
    }
    commits.push(edit.commit()?);
    Ok(commits)
}

fn source_leaf(
    base: &litchi_docx::source_backed::Package,
    count: usize,
) -> Result<Vec<litchi_docx::document::Commit<litchi_docx::source_backed::Package>>> {
    let mut commits = Vec::new();
    let mut edit = base.edit()?;
    for index in 0..count {
        if edit.remaining_operations() == 0 {
            commits.push(edit.commit()?);
            edit = commits.last().unwrap().package().edit()?;
        }
        edit.format_run(Position::new(0), Position::new(0), |properties| {
            properties.bold = Some(index % 2 == 0);
        })?;
    }
    commits.push(edit.commit()?);
    Ok(commits)
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn measure<T, H>(
    unified: &litchi_ops::File,
    native: &T,
    rounds: usize,
    leaf: fn(&T, usize) -> Result<H>,
    bytes: fn(&H, usize) -> Result<Vec<u8>>,
) -> Result<Vec<serde_json::Value>> {
    let input: Vec<_> = [true, false]
        .into_iter()
        .map(|bold| {
            let mut item = BatchItem::new("set");
            item.path = Some("/body/p[1]/r[1]".into());
            item.props = Some(Props::new().with("bold", bold));
            item
        })
        .collect();
    let mut results = Vec::new();
    for count in [1, 10, 100, 1_000, 10_000] {
        let oracle = leaf(native, count)?;
        let candidate = unified.batch(input.iter().cycle().take(count), BatchMode::Atomic)?;
        if candidate.value().len() != count
            || candidate.diagnostics().operations() != count
            || candidate.diagnostics().leaf_operations() != count
            || candidate
                .value()
                .iter()
                .any(|r| !matches!(r, litchi_ops::ItemResult::Applied(a) if a.changed))
        {
            return Err("unified operation/result accounting mismatch".into());
        }
        let expected = bytes(&oracle, count)?;
        if candidate.file().to_bytes()? != expected {
            return Err("leaf/unified byte mismatch".into());
        }
        let commits = candidate.diagnostics().commits();
        let snapshot_bytes = candidate.diagnostics().snapshot_bytes();
        drop((oracle, candidate));
        let mut leaf_ns = Vec::new();
        let mut unified_ns = Vec::new();
        // Untimed setup above also warms both routes. Destruction is outside each
        // interval, and both results retain all commits until after timing ends.
        for _ in 0..rounds {
            for is_leaf in [true, false, false, true] {
                if is_leaf {
                    let start = Instant::now();
                    let output = leaf(black_box(native), black_box(count))?;
                    let elapsed = start.elapsed().as_nanos();
                    black_box(&output);
                    leaf_ns.push(elapsed);
                    drop(output);
                } else {
                    let commands: Vec<_> = input.iter().cycle().take(count).cloned().collect();
                    let start = Instant::now();
                    let output =
                        black_box(&unified).batch(black_box(commands), BatchMode::Atomic)?;
                    let elapsed = start.elapsed().as_nanos();
                    black_box(&output);
                    unified_ns.push(elapsed);
                    drop(output);
                }
            }
        }
        results.push(serde_json::json!({
            "operations": count, "native_commits": commits,
            "unified_snapshot_bytes": snapshot_bytes,
            "leaf_ns": leaf_ns, "unified_ns": unified_ns,
            "leaf_median_ns": median(&leaf_ns), "unified_median_ns": median(&unified_ns),
            "unified_over_leaf": median(&unified_ns) as f64 / median(&leaf_ns) as f64,
        }));
    }
    Ok(results)
}

fn main() -> Result<()> {
    let rounds = std::env::args()
        .nth(1)
        .map(|s| s.parse::<usize>())
        .transpose()?
        .unwrap_or(9);
    if rounds < 3 {
        return Err("at least three ABBA rounds required".into());
    }
    let initial = litchi_ops::create(FileFormat::Docx)?
        .add(
            "/body",
            NewNode::new("paragraph").prop("text", "batch baseline"),
        )?
        .file;
    let corpus = initial.to_bytes()?;
    let corpus_sha256 = Sha256::digest(&corpus)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if corpus_sha256 != "f2f69d43d37d3103a1e82214e189b90ad627a8fa135427f9e82592bc6bb27a4a" {
        return Err("frozen corpus changed: add new route names before measuring".into());
    }
    let unified = litchi_ops::open_bytes_with(
        corpus.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let native = litchi_docx::Package::from_bytes(&corpus)?;
    let eager_results = measure(&unified, &native, rounds, leaf, |history, count| {
        let total: usize = history.iter().map(|c| c.diagnostics().operations()).sum();
        if total != count {
            return Err("native ledger count mismatch".into());
        }
        Ok(history.last().unwrap().package().clone().to_bytes()?)
    })?;
    let unified_source = litchi_ops::open_bytes(corpus.clone())?;
    let native_source = litchi_docx::source_backed::Package::from_bytes(&corpus)?;
    let source_results = measure(
        &unified_source,
        &native_source,
        rounds,
        source_leaf,
        |history, count| {
            if history
                .iter()
                .map(|c| c.diagnostics().operations())
                .sum::<usize>()
                != count
            {
                return Err("native ledger count mismatch".into());
            }
            let mut bytes = Vec::new();
            history.last().unwrap().package().write_to(&mut bytes)?;
            Ok(bytes)
        },
    )?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema": 1, "performance_claim": "scoped timing evidence only; full acceptance pending",
            "corpus_sha256": corpus_sha256,
            "corpus_bytes": corpus.len(), "rounds": rounds, "order": "ABBA",
            "scope": "prepared memory-backed owner, repeated single-run bold toggles, retained commits and unified item results; owned commands prepared before timing; excludes open/save/drop",
            "byte_equivalence": true, "routes": [
                {"leaf_route": "docx_eager_leaf_format_batch_v1", "unified_route": "docx_eager_unified_format_batch_v1", "measurements": eager_results},
                {"leaf_route": "docx_source_leaf_format_batch_v1", "unified_route": "docx_source_unified_format_batch_v1", "measurements": source_results},
            ],
        }))?
    );
    Ok(())
}
