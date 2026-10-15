//! Structural-edit census over the corpus (record 0293).
//!
//! An ignored diagnostic: for every corpus file of the formats whose
//! structural verbs record 0289 found refused on producer files, it tries
//! the structural edits and tallies the typed refusal messages:
//!
//! - XLS: insert a row and a column at the start of each worksheet (up to
//!   three) and remove them again; remove the first existing row and column.
//! - PPTX: remove the first and the last slide; move the first slide after
//!   the second and back.
//! - RTF: remove the first body paragraph; move the first paragraph after the
//!   second and back; swap the first two paragraphs and swap them back; add a
//!   paragraph at the start and remove it.
//!
//! Run with `cargo test -p litchi-ops --all-features --test structure_census
//! -- --ignored --nocapture`; it prints the tally and every refusal per file.

#![cfg(all(feature = "xls", feature = "pptx", feature = "rtf"))]

use std::collections::BTreeMap;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use litchi_core::FileFormat;
use litchi_ops::{File, NewNode, Placement, open};

fn test_data() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

fn guarded<T>(f: impl FnOnce() -> litchi_ops::Result<T>) -> Result<T, String> {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(Ok(value)) => Ok(value),
        Ok(Err(error)) => {
            let mut text = format!("{}: {}", error.code().as_str(), error.message());
            let mut source = std::error::Error::source(&error);
            while let Some(cause) = source {
                let cause_text = cause.to_string();
                if !text.contains(&cause_text) {
                    text.push_str(" <- ");
                    text.push_str(&cause_text);
                }
                source = cause.source();
            }
            Err(text)
        },
        Err(panic) => {
            let message = panic
                .downcast_ref::<&str>()
                .map(|text| (*text).to_owned())
                .or_else(|| panic.downcast_ref::<String>().cloned())
                .unwrap_or_default();
            Err(format!("panic: {message}"))
        },
    }
}

fn corpus() -> Vec<String> {
    let root = test_data();
    let mut found = Vec::new();
    let mut stack = vec![root.clone()];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else {
                found.push(
                    path.strip_prefix(&root)
                        .unwrap()
                        .to_string_lossy()
                        .into_owned(),
                );
            }
        }
    }
    found.sort();
    found
}

fn count(file: &File, parent: &str, kind: &str) -> usize {
    let Ok(node) = file.get(parent) else {
        return 0;
    };
    node.children()
        .filter_map(Result::ok)
        .filter(|child| child.kind() == kind)
        .count()
}

/// Publishes and reopens, the check every write needs.
fn republish(file: &File) -> Result<File, String> {
    let bytes = guarded(|| file.to_bytes())?;
    guarded(|| litchi_ops::open_bytes(bytes))
}

fn record(
    tally: &mut BTreeMap<String, Vec<String>>,
    case: &str,
    result: Result<(), String>,
    file: &str,
) {
    let key = match result {
        Ok(()) => format!("{case}: ok"),
        Err(error) => format!("{case}: {error}"),
    };
    tally.entry(key).or_default().push(file.to_owned());
}

fn xls(file: &File, input: &[u8], relative: &str, tally: &mut BTreeMap<String, Vec<String>>) {
    let sheets = count(file, "/", "sheet").min(3);
    for sheet in 1..=sheets {
        let parent = format!("/sheet[{sheet}]");
        for (kind, first) in [("row", "row[1]"), ("column", "col[A]")] {
            let result = guarded(|| {
                let added = file.add(parent.as_str(), NewNode::new(kind).at(Placement::Start))?;
                let removed = added.file.remove(&added.path)?;
                removed.to_bytes()
            })
            .and_then(|bytes| {
                if bytes == input {
                    Ok(())
                } else {
                    Err("round trip is not byte-exact".to_owned())
                }
            });
            record(tally, &format!("xls {kind} add+remove"), result, relative);
            let target = format!("{parent}/{first}");
            let result =
                guarded(|| file.remove(target.as_str())).and_then(|f| republish(&f).map(drop));
            record(
                tally,
                &format!("xls {kind} remove existing"),
                result,
                relative,
            );
        }
    }
}

fn pptx(file: &File, relative: &str, tally: &mut BTreeMap<String, Vec<String>>) {
    let slides = count(file, "/", "slide");
    if slides == 0 {
        return;
    }
    let result = guarded(|| file.remove("/slide[1]")).and_then(|f| republish(&f).map(drop));
    record(tally, "pptx remove first", result, relative);
    let last = format!("/slide[{slides}]");
    let result = guarded(|| file.remove(last.as_str())).and_then(|f| republish(&f).map(drop));
    record(tally, "pptx remove last", result, relative);
    if slides >= 2 {
        let result = guarded(|| {
            let moved = file.move_to("/slide[1]", Placement::after("/slide[2]")?)?;
            moved.move_to("/slide[2]", Placement::Start)
        })
        .and_then(|f| republish(&f).map(drop));
        record(tally, "pptx move+back", result, relative);
    }
}

fn exact(input: &[u8]) -> impl Fn(File) -> Result<(), String> + '_ {
    move |file| {
        let bytes = guarded(|| file.to_bytes())?;
        if bytes == input {
            Ok(())
        } else {
            republish(&file)?;
            Err("round trip is value-only (bytes differ)".to_owned())
        }
    }
}

fn rtf(file: &File, input: &[u8], relative: &str, tally: &mut BTreeMap<String, Vec<String>>) {
    let paragraphs = count(file, "/body", "paragraph");
    if paragraphs == 0 {
        return;
    }
    let result = guarded(|| file.remove("/body/p[1]")).and_then(|f| republish(&f).map(drop));
    record(tally, "rtf remove first", result, relative);
    let result = guarded(|| {
        let added = file.add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "Litchi")
                .at(Placement::Start),
        )?;
        added.file.remove(&added.path)
    })
    .and_then(|f| republish(&f).map(drop));
    record(tally, "rtf add+remove", result, relative);
    if paragraphs >= 2 {
        let result = guarded(|| {
            let moved = file.move_to("/body/p[1]", Placement::after("/body/p[2]")?)?;
            moved.move_to("/body/p[2]", Placement::before("/body/p[1]")?)
        })
        .and_then(exact(input));
        record(tally, "rtf move+back", result, relative);
        let result = guarded(|| {
            file.swap("/body/p[1]", "/body/p[2]")?
                .swap("/body/p[1]", "/body/p[2]")
        })
        .and_then(exact(input));
        record(tally, "rtf swap+back", result, relative);
    }
}

#[test]
#[ignore = "diagnostic census; run with --ignored --nocapture"]
fn structural_edit_census() {
    let mut tally: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let filter = std::env::var("CENSUS_FORMAT").ok();
    for relative in corpus() {
        let path = test_data().join(&relative);
        let Ok(file) = guarded(|| open(&path)) else {
            continue;
        };
        if matches!(file, File::Locked(_)) || file.is_signed() {
            continue;
        }
        let format = file.format();
        let wanted = |name: &str| filter.as_deref().is_none_or(|f| f == name);
        match format {
            Some(FileFormat::Xls) if wanted("xls") => {
                let input = std::fs::read(&path).unwrap();
                xls(&file, &input, &relative, &mut tally);
            },
            Some(FileFormat::Pptx) if wanted("pptx") => pptx(&file, &relative, &mut tally),
            Some(FileFormat::Rtf) if wanted("rtf") => {
                let input = std::fs::read(&path).unwrap();
                rtf(&file, &input, &relative, &mut tally);
            },
            _ => {},
        }
    }
    for (key, files) in &tally {
        println!("{:5}  {key}", files.len());
    }
    if std::env::var_os("CENSUS_FILES").is_some() {
        for (key, files) in &tally {
            println!("\n## {key}");
            for file in files {
                println!("  {file}");
            }
        }
    }
}

/// Writes the bytes of one structural round trip (`CENSUS_DUMP=<relative
/// path>`, `CENSUS_CASE=row|column|slide|paragraph`) to `CENSUS_OUT`.
#[test]
#[ignore = "diagnostic dump"]
fn structural_round_trip_dump() {
    let (Ok(relative), Ok(out)) = (std::env::var("CENSUS_DUMP"), std::env::var("CENSUS_OUT"))
    else {
        return;
    };
    let case = std::env::var("CENSUS_CASE").unwrap_or_else(|_| "row".to_owned());
    let file = open(test_data().join(&relative)).unwrap();
    let parent = std::env::var("CENSUS_PARENT").unwrap_or_else(|_| "/sheet[1]".to_owned());
    let added = file
        .add(
            parent.as_str(),
            NewNode::new(case.as_str()).at(Placement::Start),
        )
        .unwrap();
    std::fs::write(format!("{out}.added"), added.file.to_bytes().unwrap()).unwrap();
    let removed = added.file.remove(&added.path).unwrap();
    std::fs::write(out, removed.to_bytes().unwrap()).unwrap();
}

/// Prints the full error chain of an empty edit of `CENSUS_PROBE=<relative
/// path>`, for admission gaps between the reader and the editor.
#[test]
#[ignore = "diagnostic probe"]
fn empty_edit_probe() {
    let Ok(relative) = std::env::var("CENSUS_PROBE") else {
        return;
    };
    let file = open(test_data().join(&relative)).unwrap();
    match guarded(|| file.edit(|_| Ok(()))) {
        Ok(_) => println!("{relative}: empty edit admitted"),
        Err(error) => println!("{relative}: {error}"),
    }
}

/// Prints the number of `cell` query matches per sheet of
/// `CENSUS_CELLS=<relative path>`.
#[test]
#[ignore = "diagnostic probe"]
fn cell_count_probe() {
    let Ok(relative) = std::env::var("CENSUS_CELLS") else {
        return;
    };
    let file = open(test_data().join(&relative)).unwrap();
    let cells = file.query("cell").map(|matches| matches.count());
    println!("{relative}: {cells:?}");
}
