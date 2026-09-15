//! Sizing probe for change 0602 (XLSX real-producer admission).
//!
//! `census <file>...` reports, per package, whether the source-backed value
//! editor admits its first worksheet through the single-sheet door
//! (`SourceBackedEditor::edit`) and through the multi-sheet door
//! (`edit_many`), the typed refusal when it does not, and whether a complete
//! one-cell plan, commit and publication cycle succeeds.
//!
//! `edit <file> <sheet> <a1> <samples> [--publish]` opens one editor and runs
//! `samples` planning-plus-commit cycles against it, so a callgrind isolation
//! pair over two sample counts differences out the open.

use std::env;
use std::hint::black_box;
use std::io::{self, Write};
use std::path::Path;
use std::time::Instant;

use litchi_xlsx::cell::{Number, Value};
use litchi_xlsx::cell_values::{SheetCellValueEdit, SourceBackedEditor};
use litchi_xlsx::{Address, Cell as XCell, SourceBackedWorkbook};

/// A sink that counts bytes and drops them.
struct Sink(u64);

impl Write for Sink {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        self.0 += buf.len() as u64;
        Ok(buf.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Candidate replacement targets: stored non-formula value cells, in order.
fn candidates(path: &Path, limit: usize) -> Result<(String, Vec<Address>, usize), String> {
    let book = SourceBackedWorkbook::open(path).map_err(|e| format!("{e}"))?;
    let names: Vec<String> = book.sheets().map(|s| s.name().to_string()).collect();
    let first = names.first().cloned().ok_or_else(|| "no sheets".to_string())?;
    let sheet = book
        .sheet(first.as_str())
        .map_err(|e| format!("{e}"))?
        .ok_or_else(|| "sheet vanished".to_string())?;
    let Some(rect) = sheet.stored_extent().map_err(|e| format!("{e}"))? else {
        return Ok((first, Vec::new(), names.len()));
    };
    let cells = sheet.cells(rect).map_err(|e| format!("{e}"))?;
    let mut out = Vec::new();
    for c in cells {
        if matches!(c.cell, XCell::Value(_)) {
            out.push(c.address);
            if out.len() == limit {
                break;
            }
        }
    }
    Ok((first, out, names.len()))
}

fn one_line(s: &str) -> String {
    s.replace(['\n', '\t', '\r'], " ")
}

fn value() -> Value {
    Value::Number(Number::new("424242").expect("literal numeral is valid"))
}

/// One complete lifecycle through the multi-sheet door: plan, stage, commit,
/// publish. Returns the published byte count.
fn cycle(path: &Path, sheet: &str, address: Address) -> Result<u64, String> {
    let editor = SourceBackedEditor::open(path).map_err(|e| format!("{e}"))?;
    let edit = editor
        .edit_many([SheetCellValueEdit::set(sheet, address, value())])
        .map_err(|e| format!("{e}"))?;
    let commit = edit.commit().map_err(|e| format!("{e}"))?;
    let mut sink = Sink(0);
    editor
        .publish_multi_commit_to_stream(&mut sink, &commit)
        .map_err(|e| format!("{e}"))?;
    Ok(sink.0)
}

fn census(paths: &[String]) {
    println!("single\tmulti\tsheets\tsheet\ttarget\tcycle\tdetail\tpath");
    for path in paths {
        let p = Path::new(path);
        let (sheet, targets, nsheets) = match candidates(p, 64) {
            Ok(v) => v,
            Err(e) => {
                println!("read-err\t-\t-\t-\t-\t-\t{}\t{path}", one_line(&e));
                continue;
            },
        };
        let Ok(editor) = SourceBackedEditor::open(p) else {
            println!("open-err\t-\t{nsheets}\t{sheet}\t-\t-\t-\t{path}");
            continue;
        };
        let single = match editor.snapshot(sheet.as_str()) {
            Ok(_) => "yes".to_string(),
            Err(e) => format!("no:{}", one_line(&format!("{e}"))),
        };
        let first = targets.first().copied();
        let multi = match first {
            None => "no-target".to_string(),
            Some(a) => match editor.edit_many([SheetCellValueEdit::set(sheet.as_str(), a, value())])
            {
                Ok(_) => "yes".to_string(),
                Err(e) => format!("no:{}", one_line(&format!("{e}"))),
            },
        };
        let mut chosen = String::from("-");
        let mut outcome = String::from("no-target");
        let mut detail = String::from("-");
        for a in &targets {
            match cycle(p, &sheet, *a) {
                Ok(bytes) => {
                    chosen = a.a1();
                    outcome = "ok".into();
                    detail = bytes.to_string();
                    break;
                },
                Err(e) => {
                    chosen = a.a1();
                    outcome = "cycle-err".into();
                    detail = one_line(&e);
                },
            }
        }
        println!("{single}\t{multi}\t{nsheets}\t{sheet}\t{chosen}\t{outcome}\t{detail}\t{path}");
    }
}

/// Stage one edit of the requested kind: `set` replaces a stored scalar (the
/// reduced readback applies), `insert` adds a cell at an absent coordinate on a
/// row the layout does not contain, which leaves the rewrite's omission list
/// empty and takes the complete-candidate parse instead.
fn stage<'a>(sheet: &'a str, address: Address, insert: bool) -> SheetCellValueEdit<'a> {
    if insert {
        SheetCellValueEdit::insert(sheet, address, Number::new("424242").expect("numeral"))
    } else {
        SheetCellValueEdit::set(sheet, address, value())
    }
}

fn edit_mode_kind(
    path: &str,
    sheet: &str,
    a1: &str,
    samples: usize,
    publish: bool,
    insert: bool,
) {
    let p = Path::new(path);
    let address = Address::from_a1(a1).expect("valid A1 target");
    let started = Instant::now();
    let mut acc = 0u64;
    if publish {
        for _ in 0..samples {
            acc += cycle(p, sheet, address).expect("lifecycle");
        }
    } else {
        let editor = SourceBackedEditor::open(p).expect("open");
        for _ in 0..samples {
            let edit = editor
                .edit_many([stage(sheet, address, insert)])
                .expect("plan");
            let commit = edit.commit().expect("commit");
            acc += u64::from(black_box(&commit).changed());
        }
    }
    let elapsed = started.elapsed();
    println!(
        "samples={samples} publish={publish} elapsed_ns={} per_sample_ns={} acc={acc}",
        elapsed.as_nanos(),
        elapsed.as_nanos() / samples.max(1) as u128
    );
}

/// Run the OPC publication compactness audit over raw XML part bytes, the
/// same `verify_authored` call `litchi-opc` applies to BOTH the original and
/// the replacement bytes of every replaced part.
fn compact(paths: &[String]) {
    println!("verdict\tdetail\tpath");
    for path in paths {
        let Ok(bytes) = std::fs::read(path) else {
            println!("read-err\t-\t{path}");
            continue;
        };
        match xml_minifier::audit::verify_authored(&bytes, xml_minifier::audit::Limits::default()) {
            Ok(_) => println!("compact\t-\t{path}"),
            Err(e) => println!("noncompact\t{}\t{path}", one_line(&format!("{e}"))),
        }
    }
}

/// Emit one plan-plus-commit duration per line, after `warmup` untimed
/// cycles, so an external analyzer can compute per-leg quantiles and an A/A
/// floor from two interleaved legs of the identical binary.
fn bench_kind(path: &str, sheet: &str, a1: &str, warmup: usize, samples: usize, insert: bool) {
    let p = Path::new(path);
    let address = Address::from_a1(a1).expect("valid A1 target");
    let editor = SourceBackedEditor::open(p).expect("open");
    let once = |editor: &SourceBackedEditor| {
        let edit = editor
            .edit_many([stage(sheet, address, insert)])
            .expect("plan");
        let commit = edit.commit().expect("commit");
        u64::from(black_box(&commit).changed())
    };
    for _ in 0..warmup {
        black_box(once(&editor));
    }
    for _ in 0..samples {
        let started = Instant::now();
        black_box(once(&editor));
        println!("{}", started.elapsed().as_nanos());
    }
}

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("census") => census(&args[1..]),
        Some("compact") => compact(&args[1..]),
        Some("bench") => {
            let insert = args.iter().any(|a| a == "--insert");
            let pos: Vec<&String> = args[1..].iter().filter(|a| !a.starts_with("--")).collect();
            bench_kind(
                pos[0],
                pos[1],
                pos[2],
                pos[3].parse().expect("warmup"),
                pos[4].parse().expect("samples"),
                insert,
            );
        },
        Some("edit") => {
            let publish = args.iter().any(|a| a == "--publish");
            let insert = args.iter().any(|a| a == "--insert");
            let pos: Vec<&String> = args[1..].iter().filter(|a| !a.starts_with("--")).collect();
            edit_mode_kind(
                pos[0],
                pos[1],
                pos[2],
                pos[3].parse().expect("sample count"),
                publish,
                insert,
            );
        },
        _ => {
            eprintln!("usage: probe census <file>... | probe compact <xml>... | probe edit <file> <sheet> <a1> <samples> [--publish]");
            std::process::exit(2);
        },
    }
}
