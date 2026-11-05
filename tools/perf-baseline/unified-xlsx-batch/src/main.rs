//! Paired prepared-owner XLSX scalar-cell batches; opening and serialization are untimed.

use litchi_ops::{BatchItem, BatchMode, ItemResult, Props};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, io::Cursor, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];

struct Workload {
    count: usize,
    addresses: Vec<String>,
    commands: Vec<BatchItem<'static>>,
}

fn workloads() -> Vec<Workload> {
    COUNTS
        .into_iter()
        .map(|count| {
            let mut addresses = Vec::with_capacity(count);
            let mut commands = Vec::with_capacity(count);
            for index in 0..count {
                let row = index + 1;
                let address = format!("A{row}");
                let path = format!("Sheet1!{address}");
                addresses.push(address);

                let mut item = BatchItem::new("set");
                item.path = Some(path.into());
                item.props = Some(Props::new().with("value", row as i64));
                commands.push(item);
            }
            Workload {
                count,
                addresses,
                commands,
            }
        })
        .collect()
}

fn native_set(
    base: &litchi_xlsx::Package,
    addresses: &[String],
) -> Result<litchi_xlsx::workbook::Commit> {
    let mut edit = base.edit()?;
    {
        let mut sheet = edit
            .sheet("Sheet1")?
            .ok_or_else(|| std::io::Error::other("minimal workbook has no Sheet1"))?;
        for (index, address) in addresses.iter().enumerate() {
            sheet.set(address.as_str(), (index + 1) as i32)?;
        }
    }
    Ok(edit.commit()?)
}

fn promote_source(bytes: &[u8]) -> Result<litchi_xlsx::Package> {
    let source = litchi_xlsx::SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec()))?;
    let package = source
        .source_backed_package()?
        .to_opc_package_preserving_source()?;
    Ok(litchi_xlsx::Package::from_opc(package)?)
}

fn verify_unified(
    candidate: &litchi_ops::batch::BatchCommit,
    workload: &Workload,
    expected: &[u8],
) -> Result<()> {
    let changed = candidate
        .value()
        .iter()
        .filter(|result| matches!(result, ItemResult::Applied(applied) if applied.changed))
        .count();
    if candidate.value().len() != workload.count
        || changed != workload.count
        || candidate.diagnostics().operations() != workload.count
        || candidate.diagnostics().leaf_operations() != workload.count
        || candidate.diagnostics().commits() != 1
    {
        return Err(format!(
            "unified accounting mismatch for {} cells: results={}, changed={}, operations={}, leaf_operations={}, commits={}",
            workload.count,
            candidate.value().len(),
            changed,
            candidate.diagnostics().operations(),
            candidate.diagnostics().leaf_operations(),
            candidate.diagnostics().commits(),
        )
        .into());
    }
    if candidate.file().to_bytes()? != expected {
        return Err(format!("native/unified bytes differ for {} cells", workload.count).into());
    }
    Ok(())
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn measure(
    native: &litchi_xlsx::Package,
    unified: &litchi_ops::File,
    rounds: usize,
) -> Result<Vec<serde_json::Value>> {
    let mut results = Vec::new();
    for workload in workloads() {
        let native_warm = native_set(native, &workload.addresses)?;
        if native_warm.patch().operation_count() != workload.count {
            return Err(format!(
                "native patch accounting mismatch for {} cells: {} changes",
                workload.count,
                native_warm.patch().operation_count(),
            )
            .into());
        }
        let expected = native_warm.package().to_bytes()?;
        let unified_warm = unified.batch(workload.commands.iter().cloned(), BatchMode::Atomic)?;
        verify_unified(&unified_warm, &workload, &expected)?;
        drop((native_warm, unified_warm));

        let mut native_ns = Vec::with_capacity(rounds * 2);
        let mut unified_ns = Vec::with_capacity(rounds * 2);
        for _ in 0..rounds {
            for is_native in [true, false, false, true] {
                if is_native {
                    let start = Instant::now();
                    let output = native_set(black_box(native), &workload.addresses)?;
                    let elapsed = start.elapsed().as_nanos();
                    black_box(&output);
                    native_ns.push(elapsed);
                    drop(output);
                } else {
                    // Clone before starting the interval so wire command
                    // construction and ownership are not charged to execution.
                    let commands = workload.commands.clone();
                    let start = Instant::now();
                    let output =
                        black_box(unified).batch(black_box(commands), BatchMode::Atomic)?;
                    let elapsed = start.elapsed().as_nanos();
                    black_box(&output);
                    unified_ns.push(elapsed);
                    drop(output);
                }
            }
        }

        let native_median_ns = median(&native_ns);
        let unified_median_ns = median(&unified_ns);
        results.push(serde_json::json!({
            "cells": workload.count,
            "native_items_changed": workload.count,
            "native_patch_len": workload.count,
            "native_commits": 1,
            "unified_items_changed": workload.count,
            "unified_operations": workload.count,
            "unified_leaf_operations": workload.count,
            "unified_commits": 1,
            "native_ns": native_ns,
            "unified_ns": unified_ns,
            "native_median_ns": native_median_ns,
            "unified_median_ns": unified_median_ns,
            "unified_over_native": unified_median_ns as f64 / native_median_ns as f64,
        }));
    }
    Ok(results)
}

fn main() -> Result<()> {
    let rounds = std::env::args()
        .nth(1)
        .map(|value| value.parse::<usize>())
        .transpose()?
        .unwrap_or(9);
    if rounds < 3 {
        return Err("at least three ABBA rounds required".into());
    }

    let native_blank = litchi_xlsx::Package::new()?;
    let corpus = native_blank.to_bytes()?;
    let corpus_sha256 = Sha256::digest(&corpus)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();

    let eager_unified = litchi_ops::open_bytes_with(
        corpus.clone(),
        &litchi_ops::OpenOptions::new().source_mode(litchi_ops::SourceMode::Eager),
    )?;
    let eager_native = litchi_xlsx::Package::from_bytes(&corpus)?;
    let eager_results = measure(&eager_native, &eager_unified, rounds)?;

    let source_unified = litchi_ops::open_bytes(corpus.clone())?;
    let source_native = promote_source(&corpus)?;
    if source_native.to_bytes()? != corpus {
        return Err("source promotion changed the blank corpus".into());
    }
    let source_results = measure(&source_native, &source_unified, rounds)?;

    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema": 1,
            "performance_claim": "scoped timing evidence only; full acceptance pending",
            "corpus": {
                "generator": "litchi_xlsx::Package::create",
                "bytes": corpus.len(),
                "sha256": corpus_sha256,
            },
            "rounds": rounds,
            "order": "ABBA",
            "scope": "prepared blank XLSX owners, distinct scalar cell value assignments, native one-edit loop versus atomic unified BatchItem set batch; excludes open, source promotion, command construction, serialization and drop",
            "source_promotion": {
                "native_prepared_before_timing": true,
                "unified_warm_before_timing": true,
                "note": "the first source-owner unified batch warms exact-source promotion before ABBA samples"
            },
            "byte_equivalence": true,
            "routes": [
                {
                    "native_route": "xlsx_eager_leaf_scalar_batch_v1",
                    "unified_route": "xlsx_eager_unified_scalar_batch_v1",
                    "measurements": eager_results,
                },
                {
                    "native_route": "xlsx_source_leaf_scalar_batch_v1",
                    "unified_route": "xlsx_source_unified_scalar_batch_v1",
                    "measurements": source_results,
                },
            ],
        }))?
    );
    Ok(())
}
