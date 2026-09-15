//! Scratch sizing probe for change 0610 (design record for C2', lazy OPC part
//! decode behind the fallible package accessors).
//!
//! The question this probe answers is the one 0581 left open: an eager
//! `OpcPackage::open` inflates *every* admitted part before any caller has
//! asked for one, so how many parts does an ordinary open-then-one-edit-then-save
//! actually need decoded?
//!
//! Two measurements, both deterministic and build-independent:
//!
//! `census <root>`
//!     For every OOXML fixture under `<root>`: archive bytes on disk, the number
//!     of admitted parts, and the sum of their inflated payload bytes. This is
//!     exactly what `PackageReader::load_parts_eager` produces today, so it is
//!     the denominator: the work a lazy decode would be able to skip if nothing
//!     touched a part.
//!
//! `touch <fixture> <operation>`
//!     Ablation. The operation is run once on the unmodified fixture to fix a
//!     baseline outcome, then re-run once per part on a package whose payload
//!     for that one part has been replaced by a sentinel: a minimal compact XML
//!     document for an XML content type, an empty payload otherwise. A part
//!     whose ablation leaves the outcome unchanged *outside that part's own
//!     member* was never read by the operation; a part whose ablation changes
//!     the outcome (a typed refusal, a different digest for some other member,
//!     a different member set) was read. The count of the second group is the
//!     number of parts the operation touches.
//!
//!     The sentinel has to survive publication, not just open: a replaced XML
//!     payload is regenerated rather than copied, so it passes through
//!     `validate_authored_xml`, which refuses an empty payload and refuses a
//!     non-compact one (a `\r\n` after the declaration is the `NotCompact`
//!     refusal 0587 §4 reports). Hence the single-line declaration below.
//!
//!     Operations:
//!       `opc-reblob`  `OpcPackage::from_vec` -> replace the first XML part's
//!                     payload with an equal, freshly allocated one ->
//!                     `to_stream`. The pure physical open-edit-save.
//!       `opc-noop`    open and publish with no mutation (exact-source route).
//!       `xlsx-hide`   `Workbook::from_bytes` -> `edit()` -> hide the first
//!                     sheet that is not the only visible one -> `commit()` ->
//!                     `to_bytes()`. The documented semantic editor route.
//!
//! Usage:
//!   probe census <root>
//!   probe touch <fixture> <operation>
//!
//! Output is one whitespace-separated record per line, with a `#` header, so it
//! reduces with `awk` and diffs cleanly between runs.

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_opc::package::OpcPackage;
use litchi_opc::packuri::PackURI;
use sha2::{Digest, Sha256};

const EXTENSIONS: &[&str] = &[
    "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm", "potx", "ppsx",
    "xlsb",
];

/// Replacement payload for an XML part under ablation. It must be well formed
/// and compact, because a replaced XML part is regenerated and audited rather
/// than copied.
const SENTINEL_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ablated/>"#;

const USAGE: &str = "usage: probe census <root> | probe touch <fixture> <operation>";

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    match arguments.first().map(String::as_str) {
        Some("census") => {
            let root = arguments.get(1).expect(USAGE);
            census(Path::new(root));
        },
        Some("touch") => {
            let fixture = PathBuf::from(arguments.get(1).expect(USAGE));
            let operation = arguments.get(2).expect(USAGE);
            touch(&fixture, operation);
        },
        _ => {
            eprintln!("{USAGE}");
            std::process::exit(2);
        },
    }
}

// ---------------------------------------------------------------------------
// census
// ---------------------------------------------------------------------------

fn census(root: &Path) {
    let mut files = Vec::new();
    collect(root, &mut files);
    files.sort();
    println!("# fixture archive_bytes parts inflated_bytes xml_parts xml_bytes largest_part_bytes");
    for file in files {
        let Ok(data) = fs::read(&file) else { continue };
        let archive = data.len();
        match OpcPackage::from_vec(data) {
            Ok(package) => {
                let mut parts = 0_u64;
                let mut inflated = 0_u64;
                let mut xml_parts = 0_u64;
                let mut xml_bytes = 0_u64;
                let mut largest = 0_u64;
                for part in package.iter_parts() {
                    let length = part.blob().len() as u64;
                    parts += 1;
                    inflated += length;
                    largest = largest.max(length);
                    if part.content_type().contains("xml") {
                        xml_parts += 1;
                        xml_bytes += length;
                    }
                }
                println!(
                    "{} {archive} {parts} {inflated} {xml_parts} {xml_bytes} {largest}",
                    file.display()
                );
            },
            Err(error) => {
                println!("{} {archive} open-refused - - - - {error:?}", file.display());
            },
        }
    }
}

fn collect(path: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(path) else {
        return;
    };
    for entry in entries.flatten() {
        let child = entry.path();
        if child.is_dir() {
            collect(&child, files);
        } else if child
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| EXTENSIONS.contains(&extension))
        {
            files.push(child);
        }
    }
}

// ---------------------------------------------------------------------------
// touch (ablation)
// ---------------------------------------------------------------------------

/// The observable outcome of one run: either a typed error's `Debug` form, or
/// the published member set as (name, payload digest) pairs.
enum Outcome {
    Refused(String),
    Published(Vec<(String, String)>),
}

fn run(operation: &str, bytes: Vec<u8>) -> Outcome {
    match operation {
        "opc-noop" => match publish_opc(bytes, false) {
            Ok(output) => summarize(output),
            Err(error) => Outcome::Refused(error),
        },
        "opc-reblob" => match publish_opc(bytes, true) {
            Ok(output) => summarize(output),
            Err(error) => Outcome::Refused(error),
        },
        "xlsx-hide" => match publish_xlsx_hide(bytes) {
            Ok(output) => summarize(output),
            Err(error) => Outcome::Refused(error),
        },
        other => {
            eprintln!("unknown operation '{other}'");
            std::process::exit(2);
        },
    }
}

fn publish_opc(bytes: Vec<u8>, reblob: bool) -> Result<Vec<u8>, String> {
    let mut package = OpcPackage::from_vec(bytes).map_err(|error| format!("{error:?}"))?;
    if reblob {
        let target = package
            .iter_parts()
            .find(|part| part.content_type().contains("xml"))
            .map(|part| part.partname().clone());
        if let Some(name) = target {
            let replacement = package
                .get_part(&name)
                .map_err(|error| format!("{error:?}"))?
                .blob()
                .to_vec();
            package
                .get_part_mut(&name)
                .map_err(|error| format!("{error:?}"))?
                .set_blob(replacement);
        }
    }
    let mut output = Vec::new();
    package
        .to_stream(&mut output)
        .map_err(|error| format!("{error:?}"))?;
    Ok(output)
}

fn publish_xlsx_hide(bytes: Vec<u8>) -> Result<Vec<u8>, String> {
    let workbook = litchi_xlsx::Workbook::from_bytes(bytes).map_err(|error| format!("{error:?}"))?;
    let names: Vec<String> = workbook
        .sheets()
        .map(|sheet| sheet.name().to_string())
        .collect();
    if names.len() < 2 {
        return Err("skipped: fewer than two sheets".to_string());
    }
    let mut edit = workbook.edit().map_err(|error| format!("{error:?}"))?;
    {
        let mut tab = edit
            .tab(names[0].as_str())
            .map_err(|error| format!("{error:?}"))?
            .ok_or_else(|| "tab disappeared".to_string())?;
        tab.hide();
    }
    let committed = edit.commit().map_err(|error| format!("{error:?}"))?;
    committed
        .workbook()
        .to_bytes()
        .map_err(|error| format!("{error:?}"))
}

/// Reduce published bytes to a per-member view: every part's name with the
/// SHA-256 of its payload. Reopening is the comparison that matters here — two
/// runs that differ only in the ablated member must otherwise agree exactly.
fn summarize(output: Vec<u8>) -> Outcome {
    match OpcPackage::from_vec(output) {
        Ok(package) => {
            let mut members: Vec<(String, String)> = package
                .iter_parts()
                .map(|part| {
                    let mut hasher = Sha256::new();
                    hasher.update(part.blob());
                    (
                        part.partname().as_str().to_string(),
                        format!("{:x}", hasher.finalize()),
                    )
                })
                .collect();
            members.sort();
            Outcome::Published(members)
        },
        Err(error) => Outcome::Refused(format!("reopen: {error:?}")),
    }
}

/// True when two outcomes agree everywhere except at `ablated`.
fn agrees_outside(baseline: &Outcome, candidate: &Outcome, ablated: &str) -> bool {
    match (baseline, candidate) {
        (Outcome::Refused(left), Outcome::Refused(right)) => left == right,
        (Outcome::Published(left), Outcome::Published(right)) => {
            let filter = |members: &Vec<(String, String)>| -> Vec<(String, String)> {
                members
                    .iter()
                    .filter(|(name, _)| name != ablated)
                    .cloned()
                    .collect()
            };
            filter(left) == filter(right)
        },
        _ => false,
    }
}

fn touch(fixture: &Path, operation: &str) {
    let Ok(original) = fs::read(fixture) else {
        eprintln!("cannot read {}", fixture.display());
        std::process::exit(2);
    };
    let archive = original.len();

    let names: Vec<(PackURI, u64, String)> = match OpcPackage::from_vec(original.clone()) {
        Ok(package) => package
            .iter_parts()
            .map(|part| {
                (
                    part.partname().clone(),
                    part.blob().len() as u64,
                    part.content_type().to_string(),
                )
            })
            .collect(),
        Err(error) => {
            println!("# {} open-refused {error:?}", fixture.display());
            return;
        },
    };

    let baseline = run(operation, original.clone());
    if let Outcome::Refused(reason) = &baseline {
        println!(
            "# {} {operation} baseline-refused {reason}",
            fixture.display()
        );
        return;
    }

    println!("# fixture operation archive_bytes parts inflated_bytes touched touched_bytes");
    println!("# per-part: part <name> <bytes> <content_type> <touched|untouched|ablation-failed>");

    let mut parts = 0_u64;
    let mut inflated = 0_u64;
    let mut touched = 0_u64;
    let mut touched_bytes = 0_u64;

    for (name, length, content_type) in &names {
        parts += 1;
        inflated += length;

        // Build the ablated package: this one part's payload becomes the
        // sentinel.
        let sentinel = if content_type.contains("xml") {
            SENTINEL_XML.to_vec()
        } else {
            Vec::new()
        };
        let ablated_bytes = match OpcPackage::from_vec(original.clone()) {
            Ok(mut package) => match package.get_part_mut(name) {
                Ok(part) => {
                    part.set_blob(sentinel);
                    let mut output = Vec::new();
                    match package.to_stream(&mut output) {
                        Ok(()) => Ok(output),
                        Err(error) => Err(format!("{error:?}")),
                    }
                },
                Err(error) => Err(format!("{error:?}")),
            },
            Err(error) => Err(format!("{error:?}")),
        };

        let ablated_bytes = match ablated_bytes {
            Ok(bytes) => bytes,
            Err(reason) => {
                println!(
                    "part {} {length} {content_type} ablation-failed {reason}",
                    name.as_str()
                );
                touched += 1;
                touched_bytes += length;
                continue;
            },
        };

        let candidate = run(operation, ablated_bytes);
        if agrees_outside(&baseline, &candidate, name.as_str()) {
            println!("part {} {length} {content_type} untouched", name.as_str());
        } else {
            println!("part {} {length} {content_type} touched", name.as_str());
            touched += 1;
            touched_bytes += length;
        }
    }

    println!(
        "{} {operation} {archive} {parts} {inflated} {touched} {touched_bytes}",
        fixture.display()
    );
}
