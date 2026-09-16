//! Corpus differential for change 0620.
//!
//! For every XLS file named on the command line, this drives the same five
//! `cell_values` publication paths the attribution probe drives, and prints one
//! deterministic line per (file, operation): either the exact typed refusal
//! text, or the SHA-256 of the complete published artifact together with its
//! byte length and the source-backed diagnostics.
//!
//! Diffing the output of the before-leg and after-leg builds is the oracle:
//! byte-identical published output and identical refusals on every fixture the
//! editors admit.

use std::fmt::Write as _;

use litchi_xls::cell_values::{Reference, Selector, Snapshot, Storage, Value};
use sha2::{Digest, Sha256};

fn digest(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let out = hasher.finalize();
    let mut text = String::new();
    for byte in out {
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn escape(text: &str) -> String {
    let mut out = String::new();
    for character in text.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

struct NumericTarget {
    sheet: String,
    row: u32,
    column: u32,
    before: f64,
}

struct TextTarget {
    sheet: String,
    row: u32,
    column: u32,
    after: String,
}

fn find_numeric(snapshot: &Snapshot) -> Option<NumericTarget> {
    for worksheet in snapshot.worksheets() {
        for cell in worksheet.cells() {
            if !matches!(
                cell.storage(),
                Storage::Number | Storage::Rk | Storage::MulRk
            ) {
                continue;
            }
            let Value::Number(before) = cell.value() else {
                continue;
            };
            if !before.is_finite() {
                continue;
            }
            return Some(NumericTarget {
                sheet: worksheet.name().to_string(),
                row: u32::from(cell.reference().row()),
                column: u32::from(cell.reference().column()),
                before: *before,
            });
        }
    }
    None
}

fn find_text(snapshot: &Snapshot) -> Option<TextTarget> {
    let mut first: Option<(String, u32, u32, String)> = None;
    for worksheet in snapshot.worksheets() {
        for cell in worksheet.cells() {
            if cell.storage() != Storage::LabelSst {
                continue;
            }
            let Value::Text(text) = cell.value() else {
                continue;
            };
            match &first {
                None => {
                    first = Some((
                        worksheet.name().to_string(),
                        u32::from(cell.reference().row()),
                        u32::from(cell.reference().column()),
                        text.clone(),
                    ));
                },
                Some((sheet, row, column, before)) if before != text => {
                    return Some(TextTarget {
                        sheet: sheet.clone(),
                        row: *row,
                        column: *column,
                        after: text.clone(),
                    });
                },
                Some(_) => {},
            }
        }
    }
    None
}

fn run(path: &str) {
    let bytes = match std::fs::read(path) {
        Ok(bytes) => bytes,
        Err(error) => {
            println!(
                "{{\"path\":\"{}\",\"operation\":\"read\",\"error\":\"{}\"}}",
                escape(path),
                escape(&error.to_string())
            );
            return;
        },
    };
    let snapshot = match Snapshot::from_bytes(bytes.clone()) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            println!(
                "{{\"path\":\"{}\",\"operation\":\"open\",\"refused\":\"{}\"}}",
                escape(path),
                escape(&error.to_string())
            );
            return;
        },
    };
    println!(
        "{{\"path\":\"{}\",\"operation\":\"open\",\"worksheets\":{},\"workbook_stream_bytes\":{}}}",
        escape(path),
        snapshot.worksheet_count(),
        snapshot.workbook_stream().len()
    );
    let numeric = find_numeric(&snapshot);
    let text = find_text(&snapshot);

    for operation in [
        "number-plan",
        "number-source-backed",
        "number-generic",
        "string-generic",
        "noop-generic",
    ] {
        let outcome = (|| -> Result<String, String> {
            let mut transaction = snapshot.edit();
            match operation {
                "string-generic" => {
                    let target = text.as_ref().ok_or("no two distinct SST cells")?;
                    transaction
                        .set_value(
                            Selector::Name(&target.sheet),
                            Reference::new(target.row, target.column)
                                .map_err(|error| error.to_string())?,
                            Value::Text(target.after.clone()),
                        )
                        .map_err(|error| error.to_string())?;
                },
                _ => {
                    let target = numeric.as_ref().ok_or("no editable numeric cell")?;
                    let replacement = if operation == "noop-generic" {
                        target.before
                    } else {
                        target.before + 1.0
                    };
                    transaction
                        .set_numeric(
                            Selector::Name(&target.sheet),
                            Reference::new(target.row, target.column)
                                .map_err(|error| error.to_string())?,
                            replacement,
                        )
                        .map_err(|error| error.to_string())?;
                },
            }
            let mut published = Vec::new();
            let extra = match operation {
                "number-plan" => {
                    let commit = transaction
                        .commit_source_backed_plan()
                        .map_err(|error| error.to_string())?;
                    commit
                        .write_to(&mut published)
                        .map_err(|error| error.to_string())?;
                    let diagnostics = commit.diagnostics();
                    format!(
                        ",\"splice_count\":{},\"replacement_bytes\":{},\"changed_spans\":{},\"target_workbook_bytes\":{}",
                        diagnostics.splice_count(),
                        diagnostics.replacement_bytes(),
                        diagnostics.changed_spans(),
                        diagnostics.target_workbook_bytes()
                    )
                },
                "number-source-backed" => {
                    let commit = transaction
                        .commit_source_backed()
                        .map_err(|error| error.to_string())?;
                    commit
                        .write_to(&mut published)
                        .map_err(|error| error.to_string())?;
                    let diagnostics = commit.diagnostics();
                    format!(
                        ",\"splice_count\":{},\"replacement_bytes\":{},\"changed_spans\":{},\"target_workbook_bytes\":{},\"is_noop\":{}",
                        diagnostics.splice_count(),
                        diagnostics.replacement_bytes(),
                        diagnostics.changed_spans(),
                        diagnostics.target_workbook_bytes(),
                        commit.is_noop()
                    )
                },
                _ => {
                    let commit = transaction.commit().map_err(|error| error.to_string())?;
                    let diagnostics = commit.diagnostics();
                    let summary = format!(
                        ",\"changed_cells\":{},\"touched_streams\":{}",
                        diagnostics.changed_cells(),
                        diagnostics.touched_streams()
                    );
                    let (target, patch, _) = commit.into_parts();
                    published.extend_from_slice(target.bytes());
                    let inverse = patch.inverse();
                    format!(
                        "{summary},\"patch_before\":{},\"patch_after\":{},\"patch_empty\":{},\"inverse_sha256\":\"{}\"",
                        patch.before().len(),
                        patch.after().len(),
                        patch.is_empty(),
                        digest(inverse.after())
                    )
                },
            };
            if let Ok(directory) = std::env::var("XLS_DUMP_DIR") {
                let stem = path.rsplit('/').next().unwrap_or("fixture");
                let _ = std::fs::write(format!("{directory}/{stem}-{operation}.bin"), &published);
            }
            Ok(format!(
                "\"published_bytes\":{},\"sha256\":\"{}\"{extra}",
                published.len(),
                digest(&published)
            ))
        })();
        match outcome {
            Ok(fields) => println!(
                "{{\"path\":\"{}\",\"operation\":\"{operation}\",{fields}}}",
                escape(path)
            ),
            Err(refusal) => println!(
                "{{\"path\":\"{}\",\"operation\":\"{operation}\",\"refused\":\"{}\"}}",
                escape(path),
                escape(&refusal)
            ),
        }
    }
}

fn main() {
    for path in std::env::args().skip(1) {
        run(&path);
    }
}
