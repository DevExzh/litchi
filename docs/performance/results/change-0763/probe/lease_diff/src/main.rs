//! Base-versus-candidate differential probe for change 0763, built once
//! against each tree (`Cargo.toml.in`, `@TREE@`).
//!
//! Drives the streaming DOCX and XLSX writers through fixed scripts under
//! budgets whose limit on one resource sweeps from 0 to one past the script's
//! total, with the limit on a single root, or on the parent of the writer's
//! budget; and cancels before each call. For every scenario it prints the
//! first refused call and its error (or the published package's SHA-256), and
//! every level's usage of every resource once the writer has finished or been
//! dropped. A sole holder's refusals and settled counters must be identical
//! on both trees; only counters read while a lease is held may differ, and
//! none are printed.

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_xlsx::{StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter};
use sha2::{Digest, Sha256};
use std::num::{NonZeroU64, NonZeroUsize};

const RESOURCES: [Resource; 5] = [
    Resource::Memory,
    Resource::InputBytes,
    Resource::OutputBytes,
    Resource::Objects,
    Resource::Work,
];
const LOOSE: u64 = 1 << 40;

fn hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes).iter().map(|b| format!("{b:02x}")).collect()
}

fn limits_with(resource: Option<Resource>, limit: u64) -> Limits {
    let value = |candidate: Resource| if Some(candidate) == resource { limit } else { LOOSE };
    Limits::new(
        value(Resource::Memory),
        value(Resource::InputBytes),
        value(Resource::OutputBytes),
        value(Resource::Objects),
        64,
        value(Resource::Work),
    )
}

/// The writer's budget and the levels to report: the leaf alone, or the
/// leaf under a parent that carries the limit.
fn budgets(resource: Option<Resource>, limit: u64, parent: bool) -> (Budget, Vec<Budget>) {
    if parent {
        let root = Budget::root("parent", limits_with(resource, limit));
        let leaf = root.child("leaf", limits_with(None, 0));
        (leaf.clone(), vec![root, leaf])
    } else {
        let root = Budget::root("root", limits_with(resource, limit));
        (root.clone(), vec![root])
    }
}

fn context(budget: Budget) -> (CancellationSource, ExecutionContext) {
    let (source, token) = CancellationSource::pair();
    let one = NonZeroUsize::new(1).unwrap();
    let limits = ExecutionLimits::new(one, one, NonZeroU64::new(1 << 20).unwrap(), 0).unwrap();
    (source, ExecutionContext::new(budget, token, limits))
}

fn counters(levels: &[Budget]) -> String {
    levels
        .iter()
        .map(|level| {
            RESOURCES
                .iter()
                .map(|&resource| level.used(resource).to_string())
                .collect::<Vec<_>>()
                .join(",")
        })
        .collect::<Vec<_>>()
        .join("/")
}

// ---- DOCX ----

#[derive(Clone, Copy)]
enum DocxCall<'a> {
    Paragraph,
    Run,
    Text(&'a str),
    EndRun,
    EndParagraph,
}

fn docx_texts() -> Vec<String> {
    (0..24)
        .map(|index| format!("probe {index}: café & <b> {}", "é".repeat(index * 11)))
        .collect()
}

fn docx_calls(texts: &[String]) -> Vec<DocxCall<'_>> {
    let mut calls = Vec::new();
    for (index, text) in texts.iter().enumerate() {
        calls.push(DocxCall::Paragraph);
        calls.push(DocxCall::Run);
        if index % 3 == 2 {
            let split = text.char_indices().nth(text.chars().count() / 2).map_or(0, |(at, _)| at);
            calls.push(DocxCall::Text(&text[..split]));
            calls.push(DocxCall::Text(&text[split..]));
        } else {
            calls.push(DocxCall::Text(text));
        }
        calls.push(DocxCall::EndRun);
        calls.push(DocxCall::EndParagraph);
        if index % 5 == 4 {
            // A second run in the paragraph.
            calls.pop();
            calls.push(DocxCall::Run);
            calls.push(DocxCall::Text("tail"));
            calls.push(DocxCall::EndRun);
            calls.push(DocxCall::EndParagraph);
        }
    }
    calls
}

fn docx_scenario(
    calls: &[DocxCall<'_>],
    resource: Option<Resource>,
    limit: u64,
    parent: bool,
    cancel_before: Option<usize>,
) -> String {
    let (budget, levels) = budgets(resource, limit, parent);
    let (source, execution) = context(budget);
    let outcome = match StreamingDocumentWriter::new(Vec::new(), execution, StreamingDocumentLimits::default()) {
        Err(error) => format!("new: {error:?}"),
        Ok(mut writer) => {
            let mut failed = None;
            for (index, &call) in calls.iter().enumerate() {
                if cancel_before == Some(index) {
                    source.cancel();
                }
                let result = match call {
                    DocxCall::Paragraph => writer.start_paragraph(),
                    DocxCall::Run => writer.start_run(),
                    DocxCall::Text(text) => writer.write_text(text),
                    DocxCall::EndRun => writer.finish_run(),
                    DocxCall::EndParagraph => writer.finish_paragraph(),
                };
                if let Err(error) = result {
                    failed = Some(format!("call {index}: {error:?}; poisoned {}", writer.is_poisoned()));
                    break;
                }
            }
            match failed {
                Some(failure) => {
                    drop(writer);
                    failure
                },
                None => match writer.finish() {
                    Ok(bytes) => format!("ok {} bytes sha256 {}", bytes.len(), hex(&bytes)),
                    Err(error) => format!("finish: {error:?}"),
                },
            }
        },
    };
    format!("{outcome} | settled {}", counters(&levels))
}

// ---- XLSX ----

fn xlsx_row(row: u32) -> Vec<StreamingCell<'static>> {
    let values = [
        StreamingCellValue::Number(f64::from(row)),
        StreamingCellValue::Text("probe & <row> café"),
        StreamingCellValue::Bool(row % 2 == 0),
        StreamingCellValue::Blank,
    ];
    let count = usize::try_from(row % 5).unwrap();
    values
        .into_iter()
        .take(count)
        .enumerate()
        .map(|(index, value)| StreamingCell::new(u32::try_from(index).unwrap() * 2 + 1, value))
        .collect()
}

fn xlsx_scenario(rows: u32, resource: Option<Resource>, limit: u64, parent: bool, cancel_before: Option<u32>) -> String {
    let (budget, levels) = budgets(resource, limit, parent);
    let (source, execution) = context(budget);
    let outcome = match StreamingWorkbookWriter::new(Vec::new(), execution, StreamingWorkbookLimits::default()) {
        Err(error) => format!("new: {error:?}"),
        Ok(mut writer) => {
            let mut failed = None;
            for row in 1..=rows {
                if cancel_before == Some(row) {
                    source.cancel();
                }
                if let Err(error) = writer.write_row(row, xlsx_row(row)) {
                    failed = Some(format!("row {row}: {error:?}; poisoned {}", writer.is_poisoned()));
                    if writer.is_poisoned() {
                        break;
                    }
                }
            }
            if cancel_before == Some(rows + 1) {
                source.cancel();
            }
            match failed {
                Some(failure) if writer.is_poisoned() => {
                    drop(writer);
                    failure
                },
                other => {
                    let prefix = other.map(|failure| format!("{failure}; ")).unwrap_or_default();
                    match writer.finish() {
                        Ok(bytes) => format!("{prefix}ok {} bytes sha256 {}", bytes.len(), hex(&bytes)),
                        Err(error) => format!("{prefix}finish: {error:?}"),
                    }
                },
            }
        },
    };
    format!("{outcome} | settled {}", counters(&levels))
}

fn main() {
    let texts = docx_texts();
    let calls = docx_calls(&texts);
    let mut lines = 0_u64;
    // Totals under loose limits bound each sweep.
    let (budget, _) = budgets(None, 0, false);
    let (_source, execution) = context(budget.clone());
    let mut writer = StreamingDocumentWriter::new(Vec::new(), execution, StreamingDocumentLimits::default()).unwrap();
    for &call in &calls {
        match call {
            DocxCall::Paragraph => writer.start_paragraph(),
            DocxCall::Run => writer.start_run(),
            DocxCall::Text(text) => writer.write_text(text),
            DocxCall::EndRun => writer.finish_run(),
            DocxCall::EndParagraph => writer.finish_paragraph(),
        }
        .unwrap();
    }
    writer.finish().unwrap();
    let docx_totals: Vec<u64> = RESOURCES.iter().map(|&r| budget.used(r)).collect();
    println!("docx calls {} totals {docx_totals:?}", calls.len());
    for (&resource, &total) in RESOURCES.iter().zip(&docx_totals) {
        // Memory is one fixed reservation: its boundary and zero.
        let sweep: Vec<u64> = if resource == Resource::Memory {
            vec![0, 1, 63, 64, 65]
        } else {
            (0..=total + 1).collect()
        };
        for parent in [false, true] {
            for &limit in &sweep {
                println!("docx {resource:?} parent={parent} limit={limit}: {}", docx_scenario(&calls, Some(resource), limit, parent, None));
                lines += 1;
            }
        }
    }
    for cancel in 0..=calls.len() {
        println!("docx cancel-before={cancel}: {}", docx_scenario(&calls, None, 0, false, Some(cancel)));
        lines += 1;
    }

    let rows = 40;
    let (budget, _) = budgets(None, 0, false);
    let (_source, execution) = context(budget.clone());
    let mut writer = StreamingWorkbookWriter::new(Vec::new(), execution, StreamingWorkbookLimits::default()).unwrap();
    for row in 1..=rows {
        writer.write_row(row, xlsx_row(row)).unwrap();
    }
    writer.finish().unwrap();
    let xlsx_totals: Vec<u64> = RESOURCES.iter().map(|&r| budget.used(r)).collect();
    println!("xlsx rows {rows} totals {xlsx_totals:?}");
    let row_bytes = StreamingWorkbookLimits::default().max_row_bytes;
    for (&resource, &total) in RESOURCES.iter().zip(&xlsx_totals) {
        let sweep: Vec<u64> = if resource == Resource::Memory {
            vec![0, 1, row_bytes - 1, row_bytes, row_bytes + 1]
        } else {
            (0..=total + 1).collect()
        };
        for parent in [false, true] {
            for &limit in &sweep {
                println!("xlsx {resource:?} parent={parent} limit={limit}: {}", xlsx_scenario(rows, Some(resource), limit, parent, None));
                lines += 1;
            }
        }
    }
    for cancel in 1..=rows + 1 {
        println!("xlsx cancel-before={cancel}: {}", xlsx_scenario(rows, None, 0, false, Some(cancel)));
        lines += 1;
    }
    eprintln!("scenarios {lines}");
}
