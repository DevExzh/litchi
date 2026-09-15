//! Change 0631 probe A: run each of the three repaired verdict paths against
//! every OOXML fixture in the repository corpus, `repeats` times per fixture,
//! and report the set of distinct outcomes.
//!
//! Each `Relationships::new()` builds a fresh `HashMap` whose `RandomState` is
//! seeded from a per-thread counter that advances on every construction, so
//! repeating a load in one process varies the relationship visit order exactly
//! as separate processes do. A fixture whose verdict depends on that order
//! therefore reports more than one outcome here.
//!
//! The three paths, one per crate:
//!
//! * **XLSX** — an exact no-op value-only patch captured from a source-backed
//!   editor and published against an independent `OpcPackage` load of the same
//!   bytes. `capture_auxiliary*` builds the styles-and-theme array this checks.
//! * **PPTX** — every `ActiveX` control snapshot's public `Revision`, which
//!   folds in the captured binary-part relationship array.
//! * **DOCX** — an exact no-op content-control package patch captured from one
//!   load and published against another, which compares the signature
//!   staleness token.
//!
//! Every fixture also gets an OPC open-and-republish SHA-256 so that the
//! publication path can be shown unmoved between the legs.
//!
//! Usage: corpus_verdicts <test-data-root> <repeats>

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{OwnedSource, Selector as CoreSelector};
use sha2::{Digest, Sha256};

const XLSX: &[&str] = &["xlsx", "xlsm", "xltx", "xltm"];
const PPTX: &[&str] = &["pptx", "pptm", "potx", "potm", "ppsx", "ppsm"];
const DOCX: &[&str] = &["docx", "docm", "dotx", "dotm"];

fn walk(root: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    let mut names: Vec<PathBuf> = entries.filter_map(|e| e.ok()).map(|e| e.path()).collect();
    names.sort();
    for path in names {
        if path.is_dir() {
            walk(&path, out);
        } else if path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase())
            .is_some_and(|e| {
                XLSX.contains(&e.as_str()) || PPTX.contains(&e.as_str()) || DOCX.contains(&e.as_str())
            })
        {
            out.push(path);
        }
    }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

/// One line of outcome text, with newlines flattened so the report stays
/// one record per line.
fn flatten(value: impl std::fmt::Display) -> String {
    value.to_string().replace('\n', " ").replace('\t', " ")
}

/// The typed error's variant name, taken from its `Debug` spelling. Two
/// refusals with the same variant are the same refusal even when the message
/// names a different first offender — `validate_package_relationships` and its
/// peers report whichever offending relationship they meet first, which is the
/// message-only ordering class change 0628 deliberately left alone.
fn variant(value: impl std::fmt::Debug) -> String {
    let text = format!("{value:?}");
    text.split(['(', '{', ' '])
        .next()
        .unwrap_or(&text)
        .to_owned()
}

/// A verdict is what happened and which typed refusal, without the message.
struct Verdict {
    class: String,
    message: String,
}

impl Verdict {
    fn ok(detail: String) -> Self {
        Self {
            class: format!("applied {detail}"),
            message: format!("applied {detail}"),
        }
    }

    fn refused(stage: &str, error: impl std::fmt::Debug + std::fmt::Display) -> Self {
        Self {
            class: format!("{stage}:{}", variant(&error)),
            message: format!("{stage}: {}", flatten(&error)),
        }
    }

    fn plain(text: &str) -> Self {
        Self {
            class: text.to_owned(),
            message: text.to_owned(),
        }
    }
}

/// XLSX: capture an exact no-op value-only patch against a source-backed
/// editor and publish it against an independent load of the same bytes.
fn xlsx_verdict(bytes: &[u8]) -> Verdict {
    use litchi_xlsx::cell_values::SourceBackedEditor;
    let editor = match SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
    {
        Ok(editor) => editor,
        Err(error) => return Verdict::refused("editor", error),
    };
    let selector: litchi_xlsx::Selector<'_> = CoreSelector::Position(0.into());
    let edit = match editor.edit_sheets([selector]) {
        Ok(edit) => edit,
        Err(error) => return Verdict::refused("edit", error),
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => return Verdict::refused("commit", error),
    };
    let mut replay = match litchi_opc::OpcPackage::from_bytes(bytes) {
        Ok(package) => package,
        Err(error) => return Verdict::refused("replay-open", error),
    };
    match commit.patch().apply(&mut replay) {
        Ok(()) => Verdict::ok(format!("changed={}", commit.changed())),
        Err(error) => Verdict::refused("apply", error),
    }
}

/// PPTX: every `ActiveX` control snapshot's public revision, in slide and
/// control order.
fn pptx_verdict(path: &Path) -> Verdict {
    use litchi_pptx::presentation::embedded::controls::{self, Limits, slide};
    let package = match litchi_pptx::Package::open(path) {
        Ok(package) => package,
        Err(error) => return Verdict::refused("open", error),
    };
    let Ok(opc) = package.opc() else {
        return Verdict::plain("opc-unavailable");
    };
    let presentation = match package.presentation() {
        Ok(presentation) => presentation,
        Err(error) => return Verdict::refused("presentation", error),
    };
    let slides = match presentation.slides() {
        Ok(slides) => slides,
        Err(error) => return Verdict::refused("slides", error),
    };
    let mut report = String::new();
    for (index, sheet) in slides.iter().enumerate() {
        let part = sheet.part().part();
        let mut limits = Limits::default();
        let found = match controls::load_slide(opc, index, part, &mut limits) {
            Ok(found) => found,
            Err(error) => {
                report.push_str(&format!("s{index}=load-refused:{} ", flatten(error)));
                continue;
            },
        };
        for control in 0..found.len() {
            match slide::load(opc, index, part, control, &mut Limits::default()) {
                Ok(snapshot) => {
                    report.push_str(&format!("s{index}c{control}={} ", snapshot.revision()));
                },
                Err(error) => {
                    report.push_str(&format!("s{index}c{control}=refused:{} ", flatten(error)));
                },
            }
        }
    }
    if report.is_empty() {
        Verdict::plain("no-controls")
    } else {
        Verdict::plain(report.trim_end())
    }
}

/// DOCX: capture an exact no-op content-control package patch from one load
/// and publish it against another load of the same file.
fn docx_verdict(path: &Path) -> Verdict {
    let source = match litchi_docx::Package::open(path) {
        Ok(package) => package,
        Err(error) => return Verdict::refused("open", error),
    };
    let snapshot = match source.content_control_snapshot() {
        Ok(snapshot) => snapshot,
        Err(error) => return Verdict::refused("snapshot", error),
    };
    let commit = match snapshot.edit().and_then(|edit| edit.commit()) {
        Ok(commit) => commit,
        Err(error) => return Verdict::refused("commit", error),
    };
    let mut target = match litchi_docx::Package::open(path) {
        Ok(package) => package,
        Err(error) => return Verdict::refused("target-open", error),
    };
    match target.apply_content_controls(&commit) {
        Ok(()) => Verdict::ok(format!(
            "changed={} noop={}",
            commit.changed(),
            commit.patch().is_noop()
        )),
        Err(error) => Verdict::refused("apply", error),
    }
}

fn main() {
    let mut args = std::env::args().skip(1);
    let root = PathBuf::from(args.next().expect("test-data root"));
    let repeats: usize = args
        .next()
        .map_or(8, |value| value.parse().expect("repeats"));

    let mut files = Vec::new();
    walk(&root, &mut files);

    let mut probed = [0usize; 3];
    let mut unstable = [0usize; 3];
    let mut message_unstable = [0usize; 3];
    let mut opened = 0usize;
    let mut open_refused = 0usize;
    let mut published = 0usize;
    let mut save_refused = 0usize;

    for path in &files {
        let relative = path
            .strip_prefix(&root)
            .unwrap_or(path)
            .display()
            .to_string();
        let Ok(bytes) = std::fs::read(path) else {
            println!("{relative}\tREAD-ERROR");
            continue;
        };

        // Publication digest: neither of the three sites is on this path, so it
        // must be identical on both legs.
        match litchi_opc::OpcPackage::from_bytes(&bytes) {
            Ok(package) => {
                opened += 1;
                match litchi_opc::PackageWriter::to_bytes(&package) {
                    Ok(output) => {
                        published += 1;
                        let mut hasher = Sha256::new();
                        hasher.update(&output);
                        println!(
                            "{relative}\tOPC\tout_bytes={}\tsha256={}",
                            output.len(),
                            hex(&hasher.finalize())
                        );
                    },
                    Err(error) => {
                        save_refused += 1;
                        println!("{relative}\tOPC\tSAVE-REFUSED\t{}", flatten(error));
                    },
                }
            },
            Err(error) => {
                open_refused += 1;
                println!("{relative}\tOPC\tOPEN-REFUSED\t{}", flatten(error));
            },
        }

        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .map(str::to_ascii_lowercase)
            .unwrap_or_default();
        let (family, index) = if XLSX.contains(&extension.as_str()) {
            ("XLSX", 0)
        } else if PPTX.contains(&extension.as_str()) {
            ("PPTX", 1)
        } else {
            ("DOCX", 2)
        };

        let mut classes = BTreeSet::new();
        let mut messages = BTreeSet::new();
        for _ in 0..repeats {
            let verdict = match index {
                0 => xlsx_verdict(&bytes),
                1 => pptx_verdict(path),
                _ => docx_verdict(path),
            };
            classes.insert(verdict.class);
            messages.insert(verdict.message);
        }
        probed[index] += 1;
        if classes.len() > 1 {
            unstable[index] += 1;
        }
        if messages.len() > 1 {
            message_unstable[index] += 1;
        }
        println!(
            "{relative}\t{family}\tverdicts={}\tmessages={}\t{}\t{}",
            classes.len(),
            messages.len(),
            classes.iter().map(String::as_str).collect::<Vec<_>>().join(" | "),
            messages.iter().map(String::as_str).collect::<Vec<_>>().join(" | ")
        );
    }

    println!();
    println!(
        "files={} repeats={repeats} opc_opened={opened} opc_published={published} \
         opc_open_refused={open_refused} opc_save_refused={save_refused}",
        files.len()
    );
    for (index, family) in ["XLSX", "PPTX", "DOCX"].iter().enumerate() {
        println!(
            "{family} probed={} verdict_unstable={} message_unstable={}",
            probed[index], unstable[index], message_unstable[index]
        );
    }
}
