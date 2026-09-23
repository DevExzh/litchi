//! Corpus differential for change 0620, extended by change 0746 and 0748.
//!
//! Change 0748 adds, to every source-backed line, the plan's source and target
//! fingerprints (from the commit diagnostics and from the publish report), the
//! digest of a second publication from the same plan, and, for the numeric
//! plan, the digest of the checked composed view read end to end and of an
//! atomic `save` into `CENSUS_SAVE_DIR`. Diffing these lines across legs proves
//! the fingerprint values and every published byte unchanged.
//!
//! Change 0746 adds three families per fixture: the comments owner (open, one
//! comment replacement through `commit` and `commit_source_backed`), the
//! sheet-visibility owner (open, one visibility change through both commits)
//! and a digest of the public `Workbook::new` reader's complete cell and
//! sheet state, which proves the public reader's output did not move.
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

fn hex(bytes: &[u8]) -> String {
    let mut text = String::new();
    for byte in bytes {
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn fingerprints(
    diagnostic_source: &litchi_cfb::ArtifactFingerprint,
    diagnostic_target: &litchi_cfb::ArtifactFingerprint,
    report: &litchi_cfb::PublishReport,
) -> String {
    format!(
        ",\"source_fp\":\"{}\",\"target_fp\":\"{}\",\"report_source_fp\":\"{}\",\"report_target_fp\":\"{}\"",
        hex(diagnostic_source.as_bytes()),
        hex(diagnostic_target.as_bytes()),
        hex(report.source_fingerprint().as_bytes()),
        hex(report.target_fingerprint().as_bytes())
    )
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
                    let report = commit
                        .write_to(&mut published)
                        .map_err(|error| error.to_string())?;
                    let diagnostics = commit.diagnostics();
                    let mut again = Vec::new();
                    commit.write_to(&mut again).map_err(|error| error.to_string())?;
                    let view = commit.composed_source().map_err(|error| error.to_string())?;
                    let mut viewed = vec![0; published.len()];
                    litchi_core::ReadAt::read_exact_at(&view, 0, &mut viewed)
                        .map_err(|error| error.to_string())?;
                    let saved = match std::env::var("CENSUS_SAVE_DIR") {
                        Ok(directory) => {
                            let target = format!("{directory}/census-save.xls");
                            commit.save(&target).map_err(|error| error.to_string())?;
                            let bytes = std::fs::read(&target).map_err(|error| error.to_string())?;
                            let _ = std::fs::remove_file(&target);
                            digest(&bytes)
                        },
                        Err(_) => String::from("not-run"),
                    };
                    format!(
                        ",\"splice_count\":{},\"replacement_bytes\":{},\"changed_spans\":{},\"target_workbook_bytes\":{},\"again_sha256\":\"{}\",\"composed_sha256\":\"{}\",\"save_sha256\":\"{}\"{}",
                        diagnostics.splice_count(),
                        diagnostics.replacement_bytes(),
                        diagnostics.changed_spans(),
                        diagnostics.target_workbook_bytes(),
                        digest(&again),
                        digest(&viewed),
                        saved,
                        fingerprints(
                            &diagnostics.source_fingerprint(),
                            &diagnostics.target_fingerprint(),
                            &report
                        )
                    )
                },
                "number-source-backed" => {
                    let commit = transaction
                        .commit_source_backed()
                        .map_err(|error| error.to_string())?;
                    let report = commit
                        .write_to(&mut published)
                        .map_err(|error| error.to_string())?;
                    let diagnostics = commit.diagnostics();
                    format!(
                        ",\"splice_count\":{},\"replacement_bytes\":{},\"changed_spans\":{},\"target_workbook_bytes\":{},\"is_noop\":{},\"snapshot_sha256\":\"{}\"{}",
                        diagnostics.splice_count(),
                        diagnostics.replacement_bytes(),
                        diagnostics.changed_spans(),
                        diagnostics.target_workbook_bytes(),
                        commit.is_noop(),
                        digest(commit.snapshot().bytes()),
                        fingerprints(
                            &diagnostics.source_fingerprint(),
                            &diagnostics.target_fingerprint(),
                            &report
                        )
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

fn line(path: &str, operation: &str, outcome: Result<String, String>) {
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

fn comments_family(path: &str, bytes: &[u8]) {
    use litchi_xls::cell_values::{Reference, Selector};
    use litchi_xls::comments::{Snapshot, Value};
    let snapshot = match Snapshot::from_bytes(bytes.to_vec()) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            line(path, "comments-open", Err(error.to_string()));
            return;
        },
    };
    line(
        path,
        "comments-open",
        Ok(format!("\"worksheets\":{}", snapshot.worksheet_count())),
    );
    let mut target = None;
    for position in 0..snapshot.worksheet_count() {
        let Ok(Some(worksheet)) = snapshot.worksheet(Selector::Position(position)) else {
            continue;
        };
        if let Some(comment) = worksheet.comments().next() {
            target = Some((
                position,
                u32::from(comment.row()),
                u32::from(comment.column()),
                comment.author().to_string(),
                comment.text().to_string(),
            ));
            break;
        }
    }
    for operation in ["comments-generic", "comments-source-backed"] {
        let outcome = (|| -> Result<String, String> {
            let (position, row, column, author, text) =
                target.clone().ok_or("no comment to edit")?;
            // The generic commit takes a different-length text; the
            // source-backed commit needs same-length NOTE/TXO ranges, so it
            // keeps the author and reverses the text's characters, which
            // keeps its UTF-16 length and its compressed/uncompressed width.
            let value = if operation == "comments-generic" {
                Value::new("Probe", "change 0746 probe comment")
            } else {
                let reversed = text.chars().rev().collect::<String>();
                if reversed == text {
                    return Err("comment text is its own reversal".into());
                }
                Value::new(author, reversed)
            }
            .map_err(|error| error.to_string())?;
            let mut edit = snapshot.edit();
            edit.replace(
                Selector::Position(position),
                Reference::new(row, column).map_err(|error| error.to_string())?,
                value,
            )
            .map_err(|error| error.to_string())?;
            let mut published = Vec::new();
            if operation == "comments-generic" {
                let commit = edit.commit().map_err(|error| error.to_string())?;
                let (target, patch, _) = commit.into_parts();
                published.extend_from_slice(target.bytes());
                Ok(format!(
                    "\"published_bytes\":{},\"sha256\":\"{}\",\"patch_empty\":{}",
                    published.len(),
                    digest(&published),
                    patch.is_empty()
                ))
            } else {
                let commit = edit.commit_source_backed().map_err(|error| error.to_string())?;
                let report = commit
                    .write_to(&mut published)
                    .map_err(|error| error.to_string())?;
                let diagnostics = commit.diagnostics();
                Ok(format!(
                    "\"published_bytes\":{},\"sha256\":\"{}\",\"is_noop\":{}{}",
                    published.len(),
                    digest(&published),
                    commit.is_noop(),
                    fingerprints(
                        &diagnostics.source_fingerprint(),
                        &diagnostics.target_fingerprint(),
                        &report
                    )
                ))
            }
        })();
        line(path, operation, outcome);
    }
}

fn visibility_family(path: &str, bytes: &[u8]) {
    use litchi_xls::SheetVisibility;
    use litchi_xls::sheet_visibility::{Selector, Snapshot};
    let snapshot = match Snapshot::from_bytes(bytes.to_vec()) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            line(path, "visibility-open", Err(error.to_string()));
            return;
        },
    };
    line(
        path,
        "visibility-open",
        Ok(format!("\"worksheets\":{}", snapshot.worksheet_count())),
    );
    let visible = snapshot
        .worksheets()
        .filter(|worksheet| worksheet.visibility() == SheetVisibility::Visible)
        .map(|worksheet| worksheet.position())
        .collect::<Vec<_>>();
    let hidden = snapshot
        .worksheets()
        .find(|worksheet| worksheet.visibility() != SheetVisibility::Visible)
        .map(|worksheet| worksheet.position());
    let change = if visible.len() >= 2 {
        visible.last().map(|position| (*position, SheetVisibility::Hidden))
    } else {
        hidden.map(|position| (position, SheetVisibility::Visible))
    };
    for operation in ["visibility-generic", "visibility-source-backed"] {
        let outcome = (|| -> Result<String, String> {
            let (position, visibility) = change.ok_or("no visibility change candidate")?;
            let mut edit = snapshot.edit();
            edit.set_visibility(Selector::Position(position), visibility)
                .map_err(|error| error.to_string())?;
            let mut published = Vec::new();
            if operation == "visibility-generic" {
                let commit = edit.commit().map_err(|error| error.to_string())?;
                let (target, patch, _) = commit.into_parts();
                published.extend_from_slice(target.bytes());
                Ok(format!(
                    "\"published_bytes\":{},\"sha256\":\"{}\",\"patch_empty\":{}",
                    published.len(),
                    digest(&published),
                    patch.is_empty()
                ))
            } else {
                let commit = edit.commit_source_backed().map_err(|error| error.to_string())?;
                let report = commit
                    .write_to(&mut published)
                    .map_err(|error| error.to_string())?;
                let diagnostics = commit.diagnostics();
                Ok(format!(
                    "\"published_bytes\":{},\"sha256\":\"{}\",\"is_noop\":{}{}",
                    published.len(),
                    digest(&published),
                    commit.is_noop(),
                    fingerprints(
                        &diagnostics.source_fingerprint(),
                        &diagnostics.target_fingerprint(),
                        &report
                    )
                ))
            }
        })();
        line(path, operation, outcome);
    }
}

/// Digest of the public reader's complete sheet directory and every decoded
/// cell (value, formula text and bytes, flags, XF, SST index, Array owner).
fn reader_family(path: &str, bytes: &[u8]) {
    use litchi_core::sheet::Worksheet as _;
    let outcome = (|| -> Result<String, String> {
        let workbook = litchi_xls::Workbook::new(std::io::Cursor::new(bytes))
            .map_err(|error| error.to_string())?;
        let mut hasher = Sha256::new();
        let mut cells = 0_u64;
        for metadata in workbook.sheets() {
            hasher.update(format!("{metadata:?}").as_bytes());
            let Some(index) = metadata.parsed_worksheet_index() else {
                continue;
            };
            let worksheet = workbook.xls_worksheet(index).map_err(|error| error.to_string())?;
            let mut iterator = worksheet.cells();
            while let Some(cell) = iterator.next() {
                let cell = cell.map_err(|error| error.to_string())?;
                let full = worksheet
                    .get_cell(cell.row(), cell.column())
                    .ok_or("cell iterator named an absent cell")?;
                hasher.update(format!("{full:?}").as_bytes());
                cells += 1;
            }
            hasher.update(format!("{:?}", worksheet.comments()).as_bytes());
            hasher.update(format!("{:?}", worksheet.protection()).as_bytes());
        }
        let out = hasher.finalize();
        let mut text = String::new();
        for byte in out {
            let _ = write!(text, "{byte:02x}");
        }
        Ok(format!("\"cells\":{cells},\"reader_sha256\":\"{text}\""))
    })();
    line(path, "reader", outcome);
}

fn main() {
    for path in std::env::args().skip(1) {
        run(&path);
        if let Ok(bytes) = std::fs::read(&path) {
            comments_family(&path, &bytes);
            visibility_family(&path, &bytes);
            reader_family(&path, &bytes);
        }
    }
}
