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
//!
//! `roundtrip <file>...` runs one complete set-commit-publish cycle per
//! package against a real temporary file and then diffs the source archive
//! against the published archive member by member, so the report names every
//! part whose uncompressed bytes changed, were added or were dropped.
//!
//! `parts <file>...` reports the multi-sheet door's verdict for EVERY
//! worksheet part in each package, not only the first, by staging one set on
//! each sheet's first stored value cell.

use std::collections::BTreeMap;
use std::env;
use std::fs::File;
use std::hint::black_box;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
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

/// Every sheet in workbook order, paired with its first stored non-formula
/// value cell (`None` when the sheet stores no scalar).
fn sheet_targets(path: &Path) -> Result<Vec<(String, Option<Address>)>, String> {
    let book = SourceBackedWorkbook::open(path).map_err(|e| format!("{e}"))?;
    let names: Vec<String> = book.sheets().map(|s| s.name().to_string()).collect();
    let mut out = Vec::with_capacity(names.len());
    for name in names {
        let sheet = match book.sheet(name.as_str()) {
            Ok(Some(s)) => s,
            Ok(None) => {
                out.push((name, None));
                continue;
            },
            Err(_) => {
                out.push((name, None));
                continue;
            },
        };
        let target = match sheet.stored_extent() {
            Ok(Some(rect)) => match sheet.cells(rect) {
                Ok(cells) => cells
                    .into_iter()
                    .find(|c| matches!(c.cell, XCell::Value(_)))
                    .map(|c| c.address),
                Err(_) => None,
            },
            _ => None,
        };
        out.push((name, target));
    }
    Ok(out)
}

/// Read one zip archive into `member name -> uncompressed bytes`. Directory
/// entries are skipped; only file members participate in the diff.
fn members(path: &Path) -> Result<BTreeMap<String, Vec<u8>>, String> {
    let file = File::open(path).map_err(|e| format!("open {}: {e}", path.display()))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("zip {e}"))?;
    let mut out = BTreeMap::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).map_err(|e| format!("entry {i}: {e}"))?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        let mut bytes = Vec::with_capacity(entry.size() as usize);
        entry
            .read_to_end(&mut bytes)
            .map_err(|e| format!("inflate {name}: {e}"))?;
        out.insert(name, bytes);
    }
    Ok(out)
}

/// Names that differ between two archives: added, removed, or same-named with
/// different uncompressed bytes. Prefixed so the TSV column is unambiguous.
fn diff_members(src: &BTreeMap<String, Vec<u8>>, out: &BTreeMap<String, Vec<u8>>) -> Vec<String> {
    let mut names = Vec::new();
    for (name, bytes) in src {
        match out.get(name) {
            None => names.push(format!("-{name}")),
            Some(other) if other != bytes => names.push(name.clone()),
            Some(_) => {},
        }
    }
    for name in out.keys() {
        if !src.contains_key(name) {
            names.push(format!("+{name}"));
        }
    }
    names.sort();
    names
}

/// One set-commit-publish cycle into `dest`, reporting the typed refusal
/// separately from every other failure.
fn publish_to(path: &Path, dest: &Path) -> Result<(), (&'static str, String)> {
    let (sheet, targets, _) = candidates(path, 1).map_err(|e| ("error", e))?;
    let Some(address) = targets.first().copied() else {
        return Err(("error", "no stored value cell".to_string()));
    };
    let editor = SourceBackedEditor::open(path).map_err(|e| ("error", format!("{e}")))?;
    let edit = editor
        .edit_many([SheetCellValueEdit::set(sheet.as_str(), address, value())])
        .map_err(|e| ("refused", format!("{e}")))?;
    let commit = edit.commit().map_err(|e| ("commit-failed", format!("{e}")))?;
    let mut sink = File::create(dest).map_err(|e| ("error", format!("create: {e}")))?;
    editor
        .publish_multi_commit_to_stream(&mut sink, &commit)
        .map_err(|e| ("publish-failed", format!("{e}")))?;
    sink.flush().map_err(|e| ("error", format!("flush: {e}")))?;
    drop(sink);
    Ok(())
}

/// Members the value-only publication is allowed to touch. `xl/calcChain.xml`
/// may be dropped; the edited worksheet part, `xl/workbook.xml` and
/// `xl/_rels/workbook.xml.rels` may change. Anything else is a violation.
const CALC_CHAIN: &str = "xl/calcChain.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";

fn is_worksheet_part(name: &str) -> bool {
    name.starts_with("xl/worksheets/") && name.ends_with(".xml") && !name.contains("/_rels/")
}

/// `(prefix, suffix, before_middle, after_middle)` after trimming the longest
/// common prefix and suffix, so the report shows where the edit actually landed.
fn middle_diff<'a>(before: &'a [u8], after: &'a [u8]) -> (usize, usize, &'a [u8], &'a [u8]) {
    let pre = before
        .iter()
        .zip(after.iter())
        .take_while(|(a, b)| a == b)
        .count();
    let cap = before.len().min(after.len()) - pre;
    let suf = before
        .iter()
        .rev()
        .zip(after.iter().rev())
        .take_while(|(a, b)| a == b)
        .count()
        .min(cap);
    (pre, suf, &before[pre..before.len() - suf], &after[pre..after.len() - suf])
}

fn escape(bytes: &[u8], limit: usize) -> String {
    let shown = &bytes[..bytes.len().min(limit)];
    let mut out = String::from_utf8_lossy(shown).replace(['\n', '\t', '\r'], " ");
    if bytes.len() > limit {
        out.push_str("...[+");
        out.push_str(&(bytes.len() - limit).to_string());
        out.push_str("B]");
    }
    out
}

fn roundtrip(paths: &[String], scratch: &Path, keep: Option<&Path>, parts_out: Option<&Path>) {
    println!(
        "path\tverdict\tsrc_members\tout_members\tidentical\tdiffers\tadded\tremoved\tdiff\texpected_only"
    );
    let mut part_rows = vec![
        "path\tpart\tlen_before\tlen_after\tcommon_prefix\tcommon_suffix\tmiddle_ranges\tbefore_middle\tafter_middle"
            .to_string(),
    ];
    for path in paths {
        let p = Path::new(path);
        let stem = p.file_name().and_then(|s| s.to_str()).unwrap_or("out");
        let dest: PathBuf = match keep {
            Some(dir) => dir.join(format!("{stem}.published.xlsx")),
            None => scratch.join("roundtrip-out.xlsx"),
        };
        let _ = std::fs::remove_file(&dest);
        match publish_to(p, &dest) {
            Err((tag, msg)) => {
                println!(
                    "{path}\t{tag}:{}\t-\t-\t-\t-\t-\t-\t-\t-",
                    one_line(&msg)
                );
            },
            Ok(()) => match (members(p), members(&dest)) {
                (Err(e), _) | (_, Err(e)) => {
                    println!("{path}\terror:{}\t-\t-\t-\t-\t-\t-\t-\t-", one_line(&e));
                },
                (Ok(src), Ok(out)) => {
                    let (mut same, mut differ, mut added, mut removed) =
                        (0usize, Vec::new(), Vec::new(), Vec::new());
                    for (name, bytes) in &src {
                        match out.get(name) {
                            None => removed.push(name.clone()),
                            Some(o) if o != bytes => differ.push(name.clone()),
                            Some(_) => same += 1,
                        }
                    }
                    for name in out.keys() {
                        if !src.contains_key(name) {
                            added.push(name.clone());
                        }
                    }
                    let sheets: Vec<&String> =
                        differ.iter().filter(|n| is_worksheet_part(n)).collect();
                    // `xl/_rels/workbook.xml.rels` may change ONLY to drop the
                    // calcChain relationship, so it is allowed only when the
                    // source actually carried `xl/calcChain.xml`.
                    let had_chain = src.contains_key(CALC_CHAIN);
                    let expected = added.is_empty()
                        && sheets.len() <= 1
                        && removed
                            .iter()
                            .all(|n| had_chain && n.as_str() == CALC_CHAIN)
                        && differ.iter().all(|n| {
                            n == "xl/workbook.xml"
                                || is_worksheet_part(n)
                                || (had_chain && n.as_str() == WORKBOOK_RELS)
                        });
                    let mut listed: Vec<String> = differ.clone();
                    listed.extend(removed.iter().map(|n| format!("-{n}")));
                    listed.extend(added.iter().map(|n| format!("+{n}")));
                    listed.sort();
                    let text = if listed.is_empty() { "-".into() } else { listed.join(",") };
                    println!(
                        "{path}\tok\t{}\t{}\t{same}\t{}\t{}\t{}\t{text}\t{}",
                        src.len(),
                        out.len(),
                        differ.len(),
                        added.len(),
                        removed.len(),
                        if expected { "yes" } else { "no" }
                    );
                    for name in differ.iter().filter(|n| is_worksheet_part(n)) {
                        let (b, a) = (&src[name], &out[name]);
                        let (pre, suf, bm, am) = middle_diff(b, a);
                        let ranges = usize::from(!(bm.is_empty() && am.is_empty()));
                        part_rows.push(format!(
                            "{path}\t{name}\t{}\t{}\t{pre}\t{suf}\t{ranges}\t{}\t{}",
                            b.len(),
                            a.len(),
                            escape(bm, 200),
                            escape(am, 200)
                        ));
                    }
                },
            },
        }
        if keep.is_none() {
            let _ = std::fs::remove_file(&dest);
        }
    }
    if let Some(dest) = parts_out {
        let _ = std::fs::write(dest, part_rows.join("\n") + "\n");
    }
}

/// Read one cell back out of a published archive, plus a count of how many
/// cells of the sheet's stored extent still parse, so a readback that silently
/// truncates its neighbours is visible.
fn readback(path: &Path, sheet_name: &str, target: Address) -> Result<(String, usize), String> {
    let book = SourceBackedWorkbook::open(path).map_err(|e| format!("{e}"))?;
    let sheet = book
        .sheet(sheet_name)
        .map_err(|e| format!("{e}"))?
        .ok_or_else(|| "sheet missing after publication".to_string())?;
    let rect = sheet
        .stored_extent()
        .map_err(|e| format!("{e}"))?
        .ok_or_else(|| "published sheet stores no extent".to_string())?;
    let cells = sheet.cells(rect).map_err(|e| format!("{e}"))?;
    let mut found = None;
    let mut count = 0usize;
    for c in &cells {
        count += 1;
        if c.address == target {
            found = Some(match &c.cell {
                XCell::Value(Value::Number(n)) => format!("{n}"),
                XCell::Value(Value::Bool(b)) => format!("bool:{b}"),
                XCell::Value(Value::Text(_)) => "text".to_string(),
                XCell::Value(Value::Date(_)) => "date".to_string(),
                XCell::Value(Value::Error(_)) => "error".to_string(),
                XCell::Value(_) => "other-value".to_string(),
                XCell::Empty => "empty".to_string(),
                XCell::Formula(_) => "formula".to_string(),
                XCell::Unknown(_) => "unknown".to_string(),
                _ => "other".to_string(),
            });
        }
    }
    match found {
        Some(v) => Ok((v, count)),
        None => Err(format!("edited cell {} absent after publication", target.a1())),
    }
}

/// `reopen <file>...`: publish one edit, then reopen the published archive and
/// read the edited cell back.
fn reopen(paths: &[String], keep: &Path) {
    println!("package\treopened\tedited_cell\tvalue\tcells_src\tcells_out\tneighbours_ok\terror");
    for path in paths {
        let p = Path::new(path);
        let stem = p.file_name().and_then(|s| s.to_str()).unwrap_or("out");
        let dest = keep.join(format!("{stem}.reopen.xlsx"));
        let _ = std::fs::remove_file(&dest);
        let (sheet, targets, _) = match candidates(p, 1) {
            Ok(v) => v,
            Err(e) => {
                println!("{path}\tno\t-\t-\t-\t-\t-\t{}", one_line(&e));
                continue;
            },
        };
        let Some(address) = targets.first().copied() else {
            println!("{path}\tno\t-\t-\t-\t-\t-\tno stored value cell");
            continue;
        };
        if let Err((tag, msg)) = publish_to(p, &dest) {
            println!("{path}\tno\t{}\t-\t-\t-\t-\t{tag}: {}", address.a1(), one_line(&msg));
            let _ = std::fs::remove_file(&dest);
            continue;
        }
        let src_n = readback(p, &sheet, address).map(|(_, n)| n);
        match readback(&dest, &sheet, address) {
            Ok((v, n)) => {
                let (sn, ok) = match src_n {
                    Ok(sn) => (sn.to_string(), if sn == n { "yes" } else { "NO" }),
                    Err(_) => ("-".to_string(), "?"),
                };
                println!("{path}\tyes\t{}\t{v}\t{sn}\t{n}\t{ok}\t-", address.a1());
            },
            Err(e) => println!("{path}\tno\t{}\t-\t-\t-\t-\t{}", address.a1(), one_line(&e)),
        }
        let _ = std::fs::remove_file(&dest);
    }
}

fn parts(paths: &[String]) {
    println!("package\tindex\tsheet\tverdict");
    for path in paths {
        let p = Path::new(path);
        let sheets = match sheet_targets(p) {
            Ok(v) => v,
            Err(e) => {
                println!("{path}\t-\t-\tread-err:{}", one_line(&e));
                continue;
            },
        };
        if sheets.is_empty() {
            println!("{path}\t-\t-\tno-sheets");
            continue;
        }
        for (index, (name, target)) in sheets.iter().enumerate() {
            let verdict = match target {
                None => "no-target".to_string(),
                Some(address) => match SourceBackedEditor::open(p) {
                    Err(e) => format!("open-err:{}", one_line(&format!("{e}"))),
                    Ok(editor) => match editor
                        .edit_many([SheetCellValueEdit::set(name.as_str(), *address, value())])
                    {
                        Ok(_) => "ok".to_string(),
                        Err(e) => format!("refused:{}", one_line(&format!("{e}"))),
                    },
                },
            };
            println!("{path}\t{index}\t{}\t{verdict}", one_line(name));
        }
    }
}

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("census") => census(&args[1..]),
        Some("compact") => compact(&args[1..]),
        Some("parts") => parts(&args[1..]),
        Some("reopen") => {
            let keep = env::var("PROBE_SCRATCH").unwrap_or_else(|_| ".".to_string());
            reopen(&args[1..], Path::new(&keep));
        },
        Some("roundtrip") => {
            let scratch = env::var("PROBE_SCRATCH").unwrap_or_else(|_| ".".to_string());
            let keep = env::var("PROBE_KEEP").ok();
            let parts_out = env::var("PROBE_PART_DIFF").ok();
            let pos: Vec<String> = args[1..].to_vec();
            roundtrip(
                &pos,
                Path::new(&scratch),
                keep.as_deref().map(Path::new),
                parts_out.as_deref().map(Path::new),
            );
        },
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
            eprintln!(
                "usage: probe census <file>... | probe compact <xml>... | probe parts <file>... | probe roundtrip <file>... | probe edit <file> <sheet> <a1> <samples> [--publish]"
            );
            std::process::exit(2);
        },
    }
}
