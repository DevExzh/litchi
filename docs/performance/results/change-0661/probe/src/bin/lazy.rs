//! Within-binary eager/lazy differential and read-set counter for change 0661.
//!
//! This binary builds only against the **after** branch, because it uses
//! `try_iter_parts` and `deferred_decode_counters`. It compares the two owned
//! doors that the after branch exposes:
//!
//! * **eager-owned** — `OpcPackage::from_vec_reusing_payloads(data, limits,
//!   &OpcPackage::new())`. An empty donor donates nothing, so this is exactly
//!   the eager owned open the before checkout performs: `PhysPkgReader` +
//!   `PackageReader::from_phys_reader` + `authorize_owned_source`.
//! * **lazy-owned** — `OpcPackage::from_vec(data)`, which the after branch
//!   routes through `PackageReader::from_phys_reader_deferred`.
//!
//! Both doors retain the same owned source and the same preservation
//! provenance, so their publication routes are identical and any difference in
//! value, refusal or published bytes is attributable to the payload
//! representation alone. That makes this differential stronger than a
//! cross-checkout one: the two legs share every other line of code.
//!
//! Commands:
//!
//! ```text
//! lazy differ  <root>              # value, refusal and publication differential
//! lazy readset <root> <operation>  # parts and bytes inflated per operation
//! ```
//!
//! `readset` reports, per fixture: the admitted part count, the sum of every
//! part's inflated payload (the eager denominator), and the parts and bytes the
//! operation actually inflated, read from the package's own decode counters. A
//! clone of the package is retained across the operation; a clone shares the
//! deferred cells and the counter, so it observes exactly what the operation
//! decoded even when the operation consumed the package.

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_opc::limits::ReadLimits;
use litchi_opc::package::OpcPackage;
use litchi_opc::packuri::PackURI;
use sha2::{Digest, Sha256};

const EXTENSIONS: &[&str] = &[
    "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm", "potx", "ppsx",
    "xlsb",
];

const OPERATIONS: &[&str] = &["opc-noop", "opc-reblob", "xlsx-hide"];

const USAGE: &str = "usage: lazy differ <root> | lazy readset <root> <operation>";

/// Which owned door an open uses.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Door {
    Eager,
    Lazy,
}

fn open(door: Door, data: Vec<u8>) -> Result<OpcPackage, String> {
    match door {
        Door::Eager => {
            OpcPackage::from_vec_reusing_payloads(data, ReadLimits::default(), &OpcPackage::new())
        },
        Door::Lazy => OpcPackage::from_vec(data),
    }
    .map_err(|error| format!("{error:?}"))
}

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    match arguments.first().map(String::as_str) {
        Some("differ") => differ(Path::new(arguments.get(1).expect(USAGE))),
        Some("readset") => readset(
            Path::new(arguments.get(1).expect(USAGE)),
            arguments.get(2).expect(USAGE),
        ),
        _ => {
            eprintln!("{USAGE}");
            std::process::exit(2);
        },
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

fn fixtures(root: &Path) -> Vec<(String, PathBuf)> {
    let mut files = Vec::new();
    collect(root, &mut files);
    files.sort();
    files
        .into_iter()
        .map(|file| {
            let relative = file
                .strip_prefix(root)
                .unwrap_or(file.as_path())
                .display()
                .to_string();
            (relative, file)
        })
        .collect()
}

// ---------------------------------------------------------------------------
// differ
// ---------------------------------------------------------------------------

/// The read-side view of an open: the typed refusal, or every part's name,
/// content type and payload digest, taken through the fallible accessor.
fn read_view(door: Door, data: Vec<u8>) -> Result<Vec<String>, String> {
    let package = open(door, data)?;
    let mut rows = Vec::new();
    for part in package.try_iter_parts() {
        match part {
            Ok(part) => {
                let mut hasher = Sha256::new();
                hasher.update(part.blob());
                rows.push(format!(
                    "{}|{}|{:x}|{}|{}",
                    part.partname().as_str(),
                    part.content_type(),
                    hasher.finalize(),
                    part.blob().len(),
                    relationship_signature(part.rels())
                ));
            },
            Err(error) => rows.push(format!("decode-refused|{error:?}")),
        }
    }
    rows.sort();
    rows.push(format!(
        "package-rels|{}",
        relationship_signature(package.rels())
    ));
    rows.push(format!("non-part-members|{}", package.non_part_members().len()));
    Ok(rows)
}

fn relationship_signature(relationships: &litchi_opc::Relationships) -> String {
    let mut rows: Vec<String> = relationships
        .iter()
        .map(|relationship| {
            format!(
                "{}~{}~{}~{}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.is_external()
            )
        })
        .collect();
    rows.sort();
    rows.join(",")
}

fn publish(door: Door, operation: &str, data: Vec<u8>) -> Result<String, String> {
    let output = match operation {
        "opc-noop" => publish_opc(door, data, false)?,
        "opc-reblob" => publish_opc(door, data, true)?,
        "xlsx-hide" => {
            if door == Door::Eager {
                // The XLSX editor owns its own door, so there is no eager
                // variant of it; the cross-checkout `publish` binary carries
                // this comparison instead.
                return Err("skipped: no eager door".to_string());
            }
            publish_xlsx_hide(data)?
        },
        other => return Err(format!("unknown operation '{other}'")),
    };
    let mut hasher = Sha256::new();
    hasher.update(&output);
    Ok(format!("{:x}:{}", hasher.finalize(), output.len()))
}

fn publish_opc(door: Door, data: Vec<u8>, reblob: bool) -> Result<Vec<u8>, String> {
    let mut package = open(door, data)?;
    if reblob {
        if let Some(name) = first_xml_part(&package) {
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

fn first_xml_part(package: &OpcPackage) -> Option<PackURI> {
    package
        .iter_parts()
        .filter(|part| part.content_type().contains("xml"))
        .map(|part| part.partname().clone())
        .min_by(|left, right| left.as_str().cmp(right.as_str()))
}

fn publish_xlsx_hide(data: Vec<u8>) -> Result<Vec<u8>, String> {
    let workbook = litchi_xlsx::Workbook::from_bytes(data).map_err(|error| format!("{error:?}"))?;
    hide_first_tab(workbook)
}

fn hide_first_tab(workbook: litchi_xlsx::Workbook) -> Result<Vec<u8>, String> {
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

fn differ(root: &Path) {
    println!("# kind fixture detail eager lazy verdict");
    for (relative, file) in fixtures(root) {
        let Ok(data) = fs::read(&file) else { continue };

        let eager = read_view(Door::Eager, data.clone());
        let lazy = read_view(Door::Lazy, data.clone());
        match (&eager, &lazy) {
            (Err(left), Err(right)) => println!(
                "open {relative} - refused refused {}",
                verdict(left == right)
            ),
            (Ok(left), Ok(right)) => {
                println!("open {relative} - ok ok MATCH");
                println!(
                    "parts {relative} - {} {} {}",
                    left.len(),
                    right.len(),
                    verdict(left == right)
                );
                if left != right {
                    for (index, (l, r)) in left.iter().zip(right.iter()).enumerate() {
                        if l != r {
                            println!("part-diff {relative} {index} {l} {r} MISMATCH");
                        }
                    }
                }
            },
            (Ok(_), Err(right)) => println!("open {relative} - ok refused:{right} MISMATCH"),
            (Err(left), Ok(_)) => println!("open {relative} - refused:{left} ok MISMATCH"),
        }

        for operation in OPERATIONS {
            if *operation == "xlsx-hide" {
                continue;
            }
            let left = publish(Door::Eager, operation, data.clone());
            let right = publish(Door::Lazy, operation, data.clone());
            println!(
                "{operation} {relative} - {} {} {}",
                render(&left),
                render(&right),
                verdict(left == right)
            );
        }
    }
}

fn render(result: &Result<String, String>) -> String {
    match result {
        Ok(digest) => digest.clone(),
        Err(error) => format!("refused:{}", error.replace(' ', "_")),
    }
}

fn verdict(equal: bool) -> &'static str {
    if equal { "MATCH" } else { "MISMATCH" }
}

// ---------------------------------------------------------------------------
// readset
// ---------------------------------------------------------------------------

fn readset(root: &Path, operation: &str) {
    println!("# fixture operation parts inflated_bytes parts_inflated bytes_inflated outcome");
    for (relative, file) in fixtures(root) {
        let Ok(data) = fs::read(&file) else { continue };

        // Denominator: what the eager open inflates. Measured on its own
        // package so it cannot perturb the operation's counters.
        let (parts, inflated) = match OpcPackage::from_vec(data.clone()) {
            Ok(package) => {
                let mut parts = 0_u64;
                let mut inflated = 0_u64;
                for part in package.try_iter_parts() {
                    let Ok(part) = part else { continue };
                    parts += 1;
                    inflated += part.blob().len() as u64;
                }
                (parts, inflated)
            },
            Err(error) => {
                println!("{relative} {operation} - - - - open-refused:{error:?}");
                continue;
            },
        };

        let package = match OpcPackage::from_vec(data) {
            Ok(package) => package,
            Err(error) => {
                println!("{relative} {operation} {parts} {inflated} - - open-refused:{error:?}");
                continue;
            },
        };
        // A clone shares the deferred cells and the decode counters, so it
        // observes what the operation inflates even when the operation takes
        // the package by value.
        let observer = package.clone();
        let outcome = run_readset(operation, package);
        let (decoded_parts, decoded_bytes) = observer.deferred_decode_counters().unwrap_or((0, 0));
        println!(
            "{relative} {operation} {parts} {inflated} {decoded_parts} {decoded_bytes} {}",
            match outcome {
                Ok(length) => format!("published:{length}"),
                Err(error) => format!("refused:{}", error.replace(' ', "_")),
            }
        );
        if std::env::var_os("PROBE_NAME_DECODED").is_some() {
            let mut decoded: Vec<&str> = observer
                .iter_parts()
                .filter(|part| part.payload_is_decoded())
                .map(|part| part.partname().as_str())
                .collect();
            decoded.sort_unstable();
            for name in decoded {
                println!("decoded {relative} {name}");
            }
        }
    }
}

fn run_readset(operation: &str, mut package: OpcPackage) -> Result<usize, String> {
    match operation {
        "opc-noop" => {
            let mut output = Vec::new();
            package
                .to_stream(&mut output)
                .map_err(|error| format!("{error:?}"))?;
            Ok(output.len())
        },
        "opc-reblob" => {
            if let Some(name) = first_xml_part(&package) {
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
            let mut output = Vec::new();
            package
                .to_stream(&mut output)
                .map_err(|error| format!("{error:?}"))?;
            Ok(output.len())
        },
        "opc-open" => Ok(package.part_count()),
        "xlsx-hide" => {
            let workbook =
                litchi_xlsx::Workbook::from_package(package).map_err(|error| format!("{error:?}"))?;
            hide_first_tab(workbook).map(|output| output.len())
        },
        other => Err(format!("unknown operation '{other}'")),
    }
}
