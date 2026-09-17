//! Change 0665's corpus probe.
//!
//! Every mode is a documented public entry point, run identically on both
//! legs. One line per fixture, tab separated, so a script can join the legs.
use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use std::io::Write;

fn sha256(bytes: &[u8]) -> String {
    use sha2::Digest as _;
    let mut hasher = sha2::Sha256::new();
    hasher.update(bytes);
    hasher.finalize().iter().map(|byte| format!("{byte:02x}")).collect()
}

fn escape(value: &str) -> String {
    value.replace('\\', "\\\\").replace('\t', "\\t").replace('\n', "\\n").replace('\r', "\\r")
}

/// Change 0610's `xlsx-hide` route, verbatim from change 0654's probe:
/// `Workbook::from_bytes` -> `edit()` -> hide the first sheet -> `commit()` ->
/// `to_bytes()`.
fn publish_xlsx_hide(bytes: Vec<u8>) -> Result<Vec<u8>, String> {
    let workbook = litchi_xlsx::Workbook::from_bytes(bytes).map_err(|error| format!("{error:?}"))?;
    let names: Vec<String> = workbook.sheets().map(|sheet| sheet.name().to_string()).collect();
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
    committed.workbook().to_bytes().map_err(|error| format!("{error:?}"))
}

/// The eager exact no-op: open with the eager reader, publish with the eager
/// writer, no edit at all. Every Part still holds the bytes it was decoded
/// from.
fn publish_eager_noop(bytes: Vec<u8>) -> Result<Vec<u8>, String> {
    let package = OpcPackage::from_vec(bytes).map_err(|error| format!("{error:?}"))?;
    PackageWriter::to_bytes(&package).map_err(|error| format!("{error:?}"))
}

const AUTHORED_PAYLOAD: &[u8] =
    b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><litchi0665/>";

/// The eager open-edit-save route at the OPC layer: replace the
/// lexicographically first XML Part with one fixed compact payload and
/// publish. Every other Part still holds the bytes it was decoded from.
fn publish_eager_edit(bytes: Vec<u8>) -> Result<(String, Vec<u8>), String> {
    let mut package = OpcPackage::from_vec(bytes).map_err(|error| format!("{error:?}"))?;
    let mut names: Vec<String> = package
        .iter_parts()
        .filter(|part| {
            xml_minifier::audit::package::is_xml_part(part.partname().as_str(), part.content_type())
        })
        .map(|part| part.partname().as_str().to_string())
        .collect();
    names.sort();
    let target = names.first().cloned().ok_or_else(|| "no XML part".to_string())?;
    let uri = PackURI::new(&target).map_err(|error| format!("{error:?}"))?;
    package
        .get_part_mut(&uri)
        .map_err(|error| format!("{error:?}"))?
        .set_blob(AUTHORED_PAYLOAD.to_vec());
    let published = PackageWriter::to_bytes(&package).map_err(|error| format!("{error:?}"))?;
    Ok((target, published))
}

/// The same route with a deliberately NON-compact authored payload, to show
/// that a caller's own spelling is published as written after change 0665.
fn publish_eager_edit_noncompact(bytes: Vec<u8>) -> Result<(String, Vec<u8>), String> {
    let mut package = OpcPackage::from_vec(bytes).map_err(|error| format!("{error:?}"))?;
    let mut names: Vec<String> = package
        .iter_parts()
        .filter(|part| {
            xml_minifier::audit::package::is_xml_part(part.partname().as_str(), part.content_type())
        })
        .map(|part| part.partname().as_str().to_string())
        .collect();
    names.sort();
    let target = names.first().cloned().ok_or_else(|| "no XML part".to_string())?;
    let uri = PackURI::new(&target).map_err(|error| format!("{error:?}"))?;
    package
        .get_part_mut(&uri)
        .map_err(|error| format!("{error:?}"))?
        .set_blob(b"<?xml version=\"1.0\"?>\r\n<litchi0665>\n  <child a=\"1\" />\n</litchi0665>".to_vec());
    let published = PackageWriter::to_bytes(&package).map_err(|error| format!("{error:?}"))?;
    Ok((target, published))
}

/// Every XML part of `bytes`, by partname, as the eager reader decodes it.
fn xml_parts(bytes: Vec<u8>) -> Result<Vec<(String, String, usize)>, String> {
    let package = OpcPackage::from_vec(bytes).map_err(|error| format!("{error:?}"))?;
    let mut parts: Vec<(String, String, usize)> = package
        .iter_parts()
        .map(|part| {
            (
                part.partname().as_str().to_string(),
                sha256(part.blob()),
                part.blob().len(),
            )
        })
        .collect();
    parts.sort();
    Ok(parts)
}

fn run(mode: &str, list: &str) {
    let mut output = std::io::BufWriter::new(std::io::stdout().lock());
    for path in std::fs::read_to_string(list).expect("fixture list").lines() {
        let bytes = match std::fs::read(path) {
            Ok(bytes) => bytes,
            Err(error) => {
                writeln!(output, "{path}\tread_error\t-\t{}\t-", escape(&error.to_string()))
                    .expect("line");
                continue;
            },
        };
        match mode {
            "members" => members(&mut output, path, bytes),
            _ => {
                let result = match mode {
                    "xlsx-hide" => publish_xlsx_hide(bytes).map(|published| (String::from("-"), published)),
                    "eager-noop" => publish_eager_noop(bytes).map(|published| (String::from("-"), published)),
                    "eager-edit" => publish_eager_edit(bytes),
                    "eager-edit-noncompact" => publish_eager_edit_noncompact(bytes),
                    other => panic!("unknown mode {other}"),
                };
                match result {
                    Ok((target, published)) => {
                        writeln!(output, "{path}\tok\t{target}\t-\t{}", sha256(&published))
                            .expect("line");
                    },
                    Err(error) => {
                        writeln!(output, "{path}\terror\t-\t{}\t-", escape(&error)).expect("line");
                    },
                }
            },
        }
    }
    output.flush().expect("flush");
}

/// For every fixture the `xlsx-hide` route publishes, compare each XML part of
/// the published artifact with the same part of the source, by digest. The
/// route edits `/xl/workbook.xml` and the worksheet it hides; every other part
/// must come back byte for byte.
fn members<W: Write>(output: &mut W, path: &str, bytes: Vec<u8>) {
    let source = match xml_parts(bytes.clone()) {
        Ok(parts) => parts,
        Err(error) => {
            writeln!(output, "{path}\tsource_error\t-\t{}\t-", escape(&error)).expect("line");
            return;
        },
    };
    let published = match publish_xlsx_hide(bytes) {
        Ok(published) => published,
        Err(error) => {
            writeln!(output, "{path}\troute_error\t-\t{}\t-", escape(&error)).expect("line");
            return;
        },
    };
    let after = match xml_parts(published) {
        Ok(parts) => parts,
        Err(error) => {
            writeln!(output, "{path}\treopen_error\t-\t{}\t-", escape(&error)).expect("line");
            return;
        },
    };
    let before: std::collections::BTreeMap<&str, (&str, usize)> = source
        .iter()
        .map(|(name, digest, len)| (name.as_str(), (digest.as_str(), *len)))
        .collect();
    let mut identical = 0usize;
    let mut changed = Vec::new();
    let mut added = Vec::new();
    for (name, digest, _len) in &after {
        match before.get(name.as_str()) {
            None => added.push(name.clone()),
            Some((source_digest, _)) if source_digest == digest => identical += 1,
            Some(_) => changed.push(name.clone()),
        }
    }
    let removed: Vec<&str> = before
        .keys()
        .filter(|name| !after.iter().any(|(after_name, _, _)| after_name == *name))
        .copied()
        .collect();
    writeln!(
        output,
        "{path}\tmembers\tidentical={identical}\tchanged={}\tadded={}\tremoved={}",
        changed.join(","),
        added.join(","),
        removed.join(","),
    )
    .expect("line");
}

fn main() {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    let mode = arguments.first().expect("mode").clone();
    let list = arguments.get(1).expect("fixture list").clone();
    run(&mode, &list);
}
