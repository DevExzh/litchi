//! Cross-checkout publication differential for change 0661 (lazy OPC part
//! decode).
//!
//! This binary uses **only** API that exists on both the before checkout
//! (`ab07e2a47`, the eager owned open) and the after branch (the lazy owned
//! open), so the same source builds on both legs and the two outputs can be
//! diffed line for line. It deliberately avoids `iter_parts()`'s item type,
//! which the after leg narrows, by touching only `partname()` and
//! `content_type()` on the iterator's items — both present in both legs.
//!
//! For every OOXML fixture under `<root>` it runs each operation through the
//! ordinary owned door `OpcPackage::from_vec` and prints the SHA-256 of the
//! complete published stream, or the `Debug` form of the typed refusal.
//!
//! Usage: `publish <root> <operation>` with `operation` one of
//! `opc-noop`, `opc-reblob`, `opc-members`, `xlsx-hide`.
//!
//! * `opc-noop` — open, publish unchanged (the exact-source route).
//! * `opc-reblob` — open, replace the first XML part's payload with an equal,
//!   freshly allocated one, publish (the targeted-preservation route).
//! * `opc-members` — open and print each part's name, content type and payload
//!   digest, without publishing. This is the read-side value differential.
//! * `xlsx-hide` — the documented XLSX editor route from change 0610.

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

const USAGE: &str = "usage: publish <root> <opc-noop|opc-reblob|opc-members|xlsx-hide>";

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let root = arguments.first().expect(USAGE);
    let operation = arguments.get(1).expect(USAGE).as_str();

    let mut files = Vec::new();
    collect(Path::new(root), &mut files);
    files.sort();

    println!("# fixture operation result");
    for file in files {
        let relative = file
            .strip_prefix(root)
            .unwrap_or(file.as_path())
            .display()
            .to_string();
        let Ok(data) = fs::read(&file) else {
            println!("{relative} {operation} read-failed");
            continue;
        };
        if operation == "opc-members" {
            for line in members(data) {
                println!("{relative} member {line}");
            }
            continue;
        }
        let result = match run(operation, data) {
            Ok(output) => {
                let mut hasher = Sha256::new();
                hasher.update(&output);
                format!("{:x}:{}", hasher.finalize(), output.len())
            },
            Err(error) => format!("refused:{error}"),
        };
        println!("{relative} {operation} {result}");
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

/// Every part's name, content type and payload digest, read through the
/// fallible accessor so a deferred payload's refusal is visible.
fn members(data: Vec<u8>) -> Vec<String> {
    let package = match OpcPackage::from_vec(data) {
        Ok(package) => package,
        Err(error) => return vec![format!("open-refused:{error:?}")],
    };
    let mut names: Vec<PackURI> = package
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut lines = Vec::new();
    for name in names {
        match package.get_part(&name) {
            Ok(part) => {
                let mut hasher = Sha256::new();
                hasher.update(part.blob());
                lines.push(format!(
                    "{} {} {:x} {}",
                    name.as_str(),
                    part.content_type(),
                    hasher.finalize(),
                    part.blob().len()
                ));
            },
            Err(error) => lines.push(format!("{} get-refused:{error:?}", name.as_str())),
        }
    }
    lines
}

fn run(operation: &str, data: Vec<u8>) -> Result<Vec<u8>, String> {
    match operation {
        "opc-noop" => publish_opc(data, false),
        "opc-reblob" => publish_opc(data, true),
        "xlsx-hide" => publish_xlsx_hide(data),
        other => Err(format!("unknown operation '{other}'")),
    }
}

fn publish_opc(data: Vec<u8>, reblob: bool) -> Result<Vec<u8>, String> {
    let mut package = OpcPackage::from_vec(data).map_err(|error| format!("{error:?}"))?;
    if reblob {
        let target = package
            .iter_parts()
            .filter(|part| part.content_type().contains("xml"))
            .map(|part| part.partname().clone())
            .min_by(|left, right| left.as_str().cmp(right.as_str()));
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

fn publish_xlsx_hide(data: Vec<u8>) -> Result<Vec<u8>, String> {
    let workbook = litchi_xlsx::Workbook::from_bytes(data).map_err(|error| format!("{error:?}"))?;
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
