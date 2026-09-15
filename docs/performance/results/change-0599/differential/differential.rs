// Differential for change 0599: every `.xlsb` fixture, one leg per binary.
//
// Two binaries share this source. One links `litchi-xlsb` at the base commit
// 08d968f8e, the other at branch `perf/0599-xlsb-commit-single-parse`. Each
// drives the identical cell-value CRUD sequence over every fixture given on
// the command line and prints every published byte digest, every readback and
// every typed refusal, so `diff` of the two reports is the proof that the
// commit path's observable behaviour did not move.

use litchi_core::sheet::traits::WorkbookTrait as _;
use litchi_xlsb::Workbook;
use litchi_xlsb::cell_values::{Reference, StyleIndex, Value};
use sha2::{Digest, Sha256};
use std::io::Cursor;

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut out = String::with_capacity(64);
    for byte in digest {
        out.push_str(&format!("{byte:02x}"));
    }
    out
}

fn save(workbook: &Workbook) -> String {
    let mut out = Cursor::new(Vec::new());
    match workbook.save(&mut out) {
        Ok(()) => sha256(&out.into_inner()),
        Err(error) => format!("Err({error})"),
    }
}

fn leg(path: &str, bytes: &[u8]) -> Vec<String> {
    let mut lines = Vec::new();
    let mut say = |what: String, value: String| lines.push(format!("{path}\t{what}\t{value}"));

    let workbook = match Workbook::new(Cursor::new(bytes.to_vec())) {
        Ok(workbook) => workbook,
        Err(error) => {
            say("open".to_string(), format!("Err({error})"));
            return lines;
        },
    };
    say(
        "worksheet_names".to_string(),
        format!("{:?}", workbook.worksheet_names()),
    );
    let count = workbook.worksheet_count();
    say("worksheet_count".to_string(), count.to_string());
    say("save.baseline".to_string(), save(&workbook));

    for sheet in 0..count {
        match workbook.cell_values(sheet) {
            Ok(snapshot) => {
                let cells: Vec<String> = snapshot
                    .cells()
                    .map(|cell| {
                        format!(
                            "({},{},{},{:?})",
                            cell.reference().row(),
                            cell.reference().column(),
                            cell.style().get(),
                            cell.value()
                        )
                    })
                    .collect();
                say(
                    format!("sheet{sheet}.snapshot"),
                    format!(
                        "src={} n={} cells={}",
                        sha256(snapshot.source_bytes()),
                        cells.len(),
                        sha256(cells.join("|").as_bytes())
                    ),
                );
            },
            Err(error) => say(format!("sheet{sheet}.snapshot"), format!("Err({error})")),
        }
    }

    // 1. An exact no-op commit on every worksheet: nothing may change.
    for sheet in 0..count {
        let mut candidate = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
        let Ok(snapshot) = candidate.cell_values(sheet) else {
            continue;
        };
        let commit = match snapshot.edit().commit() {
            Ok(commit) => commit,
            Err(error) => {
                say(format!("sheet{sheet}.noop.commit"), format!("Err({error})"));
                continue;
            },
        };
        say(
            format!("sheet{sheet}.noop.patch_empty"),
            commit.patch().is_empty().to_string(),
        );
        match candidate.apply_cell_values(sheet, &commit) {
            Ok(published) => say(
                format!("sheet{sheet}.noop.published"),
                format!(
                    "src={} save={} names={:?} shared={} xfs={}",
                    sha256(published.source_bytes()),
                    save(&candidate),
                    candidate.worksheet_names(),
                    candidate.shared_strings().len(),
                    candidate.styles().cell_xfs.len()
                ),
            ),
            Err(error) => say(format!("sheet{sheet}.noop.published"), format!("Err({error})")),
        }
    }

    // 2. One real scalar edit per worksheet, wherever the fixture has one.
    for sheet in 0..count {
        let mut candidate = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
        let Ok(snapshot) = candidate.cell_values(sheet) else {
            continue;
        };
        let target = snapshot.cells().find_map(|cell| match cell.value() {
            Value::Number(value) => Some((
                cell.reference(),
                Value::Number(if value.to_bits() == 1.0f64.to_bits() {
                    2.0
                } else {
                    1.0
                }),
            )),
            Value::RkNumber(value) => Some((
                cell.reference(),
                Value::RkNumber(if value.to_bits() == 1.0f64.to_bits() {
                    2.0
                } else {
                    1.0
                }),
            )),
            Value::Boolean(value) => Some((cell.reference(), Value::Boolean(!value))),
            _ => None,
        });
        let Some((reference, after)) = target else {
            say(format!("sheet{sheet}.edit"), "no editable scalar".to_string());
            continue;
        };
        let mut edit = snapshot.edit();
        if let Err(error) = edit.set_value(reference, after) {
            say(format!("sheet{sheet}.edit"), format!("set Err({error})"));
            continue;
        }
        let commit = match edit.commit() {
            Ok(commit) => commit,
            Err(error) => {
                say(format!("sheet{sheet}.edit"), format!("commit Err({error})"));
                continue;
            },
        };
        match candidate.apply_cell_values(sheet, &commit) {
            Ok(published) => {
                let readback = match candidate.cell_values(sheet) {
                    Ok(snapshot) => match snapshot.cell(reference) {
                        Ok(Some(cell)) => format!("{:?}", cell.value()),
                        Ok(None) => "missing".to_string(),
                        Err(error) => format!("Err({error})"),
                    },
                    Err(error) => format!("Err({error})"),
                };
                say(
                    format!("sheet{sheet}.edit"),
                    format!(
                        "at=({},{}) src={} save={} readback={readback} names={:?} shared={} xfs={}",
                        reference.row(),
                        reference.column(),
                        sha256(published.source_bytes()),
                        save(&candidate),
                        candidate.worksheet_names(),
                        candidate.shared_strings().len(),
                        candidate.styles().cell_xfs.len()
                    ),
                );
            },
            Err(error) => say(format!("sheet{sheet}.edit"), format!("apply Err({error})")),
        }
    }

    // 3. A refused publication. ADR 0003: the published workbook must be
    //    byte-identical afterwards.
    let mut candidate = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
    let before_save = save(&candidate);
    match candidate.edit_cell_values(0) {
        Ok(mut edit) => {
            let reference = Reference::new(10_001, 100).unwrap();
            let style = StyleIndex::new(0x00FF_FFFF).unwrap();
            match edit.insert(reference, style, Value::Number(1.0)) {
                Ok(()) => match edit.commit() {
                    Ok(commit) => {
                        let outcome = match candidate.apply_cell_values(0, &commit) {
                            Ok(_) => "unexpectedly published".to_string(),
                            Err(error) => format!("Err({error})"),
                        };
                        say(
                            "refusal".to_string(),
                            format!("{outcome} unchanged={}", save(&candidate) == before_save),
                        );
                    },
                    Err(error) => say("refusal".to_string(), format!("commit Err({error})")),
                },
                Err(error) => say("refusal".to_string(), format!("insert Err({error})")),
            }
        },
        Err(error) => say("refusal".to_string(), format!("edit Err({error})")),
    }
    lines
}

fn main() {
    let mut fixtures: Vec<String> = std::env::args().skip(1).collect();
    fixtures.sort();
    for path in &fixtures {
        let Ok(bytes) = std::fs::read(path) else {
            println!("{path}\tread\tERROR");
            continue;
        };
        let display = path.rsplit('/').next().unwrap_or(path);
        for line in leg(display, &bytes) {
            println!("{line}");
        }
    }
}
