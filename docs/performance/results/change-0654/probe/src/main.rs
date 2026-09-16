//! Change 0654 corpus probe.
//!
//! Two subcommands, both run identically against the before and after legs so
//! that every difference in the output is a difference in the audited
//! contract.
//!
//! * `audit` reads length-prefixed XML documents from stdin and prints, for
//!   each, the verdict of the authored audit and of the original-bytes audit.
//!   The before leg has no separate original-bytes policy, so it runs the
//!   authored audit for both columns; that is exactly the contract being
//!   replaced.
//! * `publish` opens each named OOXML package with `SourceBackedPackage`,
//!   replaces the lexicographically first XML Part through
//!   `write_part_overlay_to_stream` with one fixed compact payload, and
//!   reports the outcome and the SHA-256 of the published archive.
#![allow(clippy::expect_used, clippy::print_stdout, clippy::unwrap_used)]

use std::io::{Read, Write};

use litchi_opc::{PackURI, SourceBackedPackage};

/// The one compact authored payload every publication writes.
const REPLACEMENT: &[u8] = b"<litchi0654/>";

fn escape(value: &str) -> String {
    value.replace('\\', "\\\\").replace('\t', "\\t").replace('\n', "\\n")
}

fn verdict(result: Result<xml_minifier::audit::Report, xml_minifier::audit::Error>) -> String {
    match result {
        Ok(_) => "ok".to_owned(),
        Err(error) => format!("ERR {error}"),
    }
}

fn audit_stdin() {
    let mut input = std::io::stdin().lock();
    let mut output = std::io::BufWriter::new(std::io::stdout().lock());
    let limits = xml_minifier::audit::Limits::default();
    loop {
        let mut header = [0_u8; 4];
        match input.read_exact(&mut header) {
            Ok(()) => {},
            Err(_) => break,
        }
        let length = u32::from_le_bytes(header) as usize;
        let mut bytes = vec![0_u8; length];
        input.read_exact(&mut bytes).expect("document body");
        let authored = verdict(xml_minifier::audit::verify_authored(&bytes, limits));
        let source = verdict(source_audit(&bytes, limits));
        writeln!(output, "{}\t{}", escape(&authored), escape(&source)).expect("verdict line");
    }
    output.flush().expect("flush verdicts");
}

#[cfg(feature = "source-policy")]
fn source_audit(
    bytes: &[u8],
    limits: xml_minifier::audit::Limits,
) -> Result<xml_minifier::audit::Report, xml_minifier::audit::Error> {
    xml_minifier::audit::verify_source(bytes, limits)
}

#[cfg(not(feature = "source-policy"))]
fn source_audit(
    bytes: &[u8],
    limits: xml_minifier::audit::Limits,
) -> Result<xml_minifier::audit::Report, xml_minifier::audit::Error> {
    xml_minifier::audit::verify_authored(bytes, limits)
}

fn sha256(bytes: &[u8]) -> String {
    use sha2::Digest as _;
    let mut hasher = sha2::Sha256::new();
    hasher.update(bytes);
    hasher
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn target_part(package: &SourceBackedPackage) -> Option<String> {
    let mut names: Vec<String> = package
        .iter_parts()
        .filter(|part| {
            xml_minifier::audit::package::is_xml_part(part.partname().as_str(), part.content_type())
        })
        .map(|part| part.partname().as_str().to_owned())
        .collect();
    names.sort();
    names.into_iter().next()
}

fn publish(paths: &[String], out_dir: &str) {
    let mut output = std::io::BufWriter::new(std::io::stdout().lock());
    for (index, path) in paths.iter().enumerate() {
        let package = match SourceBackedPackage::from_path(path) {
            Ok(package) => package,
            Err(error) => {
                writeln!(output, "{path}\t-\topen_error\t{}\t-", escape(&error.to_string()))
                    .expect("line");
                continue;
            },
        };
        let Some(name) = target_part(&package) else {
            writeln!(output, "{path}\t-\tno_xml_part\t-\t-").expect("line");
            continue;
        };
        let uri = PackURI::new(&name).expect("part name is a valid pack URI");
        let mut published = Vec::new();
        match package.write_part_overlay_to_stream(&mut published, &uri, REPLACEMENT.to_vec()) {
            Ok(()) => {
                let digest = sha256(&published);
                let target = format!("{out_dir}/{index:04}.zip");
                std::fs::write(&target, &published).expect("write published archive");
                writeln!(output, "{path}\t{name}\tok\t-\t{digest}").expect("line");
            },
            Err(error) => {
                writeln!(output, "{path}\t{name}\terror\t{}\t-", escape(&error.to_string()))
                    .expect("line");
            },
        }
    }
    output.flush().expect("flush publication lines");
}

/// Two synthetic packages whose *original* part bytes are compact, so the
/// compactness refusal never applies on either leg, and which therefore show
/// that the two refusals that surface after this change were already reachable
/// at the base.
fn witness() {
    use soapberry_zip::office::StreamingArchiveWriter;
    const TYPES: &str = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>";
    const RELS: &str = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>";
    const DOCUMENT: &[u8] = b"<document><body/></document>";

    let build = |trailing: bool| -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("[Content_Types].xml", TYPES.as_bytes()).expect("types");
        writer.write_stored("_rels/.rels", RELS.as_bytes()).expect("rels");
        writer.write_stored("word/document.xml", DOCUMENT).expect("document");
        let mut bytes = writer.finish_to_bytes().expect("finish");
        if trailing {
            bytes.extend_from_slice(b"TRAILING");
        }
        bytes
    };

    let uri = PackURI::new("/word/document.xml").expect("uri");
    for (label, bytes) in [
        ("compact-original+clean", build(false)),
        ("compact-original+trailing-bytes", build(true)),
    ] {
        let verdict = match SourceBackedPackage::from_vec(bytes) {
            Err(error) => format!("open_error {error}"),
            Ok(package) => {
                let mut published = Vec::new();
                match package.write_part_overlay_to_stream(&mut published, &uri, REPLACEMENT.to_vec()) {
                    Ok(()) => format!("ok {}", sha256(&published)),
                    Err(error) => format!("error {error}"),
                }
            },
        };
        println!("{label}\t{verdict}");
    }
}

/// Change 0610's `xlsx-hide` route: `Workbook::from_bytes` -> `edit()` ->
/// hide the first sheet -> `commit()` -> `to_bytes()`. This is the ordinary
/// documented semantic save, on which the XLSX writer compacts the parts it
/// regenerates, so the *original*-bytes audit is the only compactness gate
/// the route can fail.
fn xlsx_hide(list: &str) {
    let mut output = std::io::BufWriter::new(std::io::stdout().lock());
    for path in std::fs::read_to_string(list).expect("fixture list").lines() {
        let bytes = match std::fs::read(path) {
            Ok(bytes) => bytes,
            Err(error) => {
                writeln!(output, "{path}\tread_error\t{}\t-", escape(&error.to_string())).expect("line");
                continue;
            },
        };
        match publish_xlsx_hide(bytes) {
            Ok(published) => {
                writeln!(output, "{path}\tok\t-\t{}", sha256(&published)).expect("line");
            },
            Err(error) => {
                writeln!(output, "{path}\terror\t{}\t-", escape(&error)).expect("line");
            },
        }
    }
    output.flush().expect("flush");
}

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

/// For one fixture the eager `xlsx-hide` route refuses, report whether the
/// refused Part's blob is byte-identical to the source ZIP member. If it is,
/// the eager writer is applying the authored contract to bytes litchi did not
/// author, which is the same defect class this change fixes on the
/// source-backed route but at a site with no original/replacement pair.
fn blob_provenance(path: &str, member: &str) {
    let bytes = std::fs::read(path).expect("fixture");
    let package = litchi_opc::OpcPackage::from_vec(bytes.clone()).expect("package opens");
    let uri = format!("/{member}");
    for part in package.iter_parts() {
        if part.partname().as_str() != uri {
            continue;
        }
        let blob = part.blob();
        let source = zip_member(&bytes, member);
        println!(
            "part={uri} blob_len={} source_len={} identical={} blob_sha={} source_sha={}",
            blob.len(),
            source.len(),
            blob == source.as_slice(),
            sha256(blob),
            sha256(&source),
        );
        return;
    }
    println!("part={uri} not found");
}

/// Read one member out of a ZIP without a zip crate: the OPC reader already
/// decoded it, so ask a second `SourceBackedPackage` for the same Part.
fn zip_member(bytes: &[u8], member: &str) -> Vec<u8> {
    let package = SourceBackedPackage::from_vec(bytes.to_vec()).expect("source package opens");
    let uri = PackURI::new(&format!("/{member}")).expect("uri");
    let part = package.part(&uri).expect("part exists");
    part.data().expect("part payload").as_bytes().to_vec()
}

fn main() {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    match arguments.first().map(String::as_str) {
        Some("audit") => audit_stdin(),
        Some("witness") => witness(),
        Some("blob-provenance") => blob_provenance(
            arguments.get(1).expect("fixture path"),
            arguments.get(2).expect("member name"),
        ),
        Some("xlsx-hide") => {
            let list = arguments.get(1).expect("xlsx-hide needs a fixture list file");
            xlsx_hide(list);
        },
        Some("publish") => {
            let out_dir = arguments.get(1).expect("publish needs an output directory");
            let list = arguments.get(2).expect("publish needs a fixture list file");
            let text = std::fs::read_to_string(list).expect("fixture list");
            let paths: Vec<String> = text.lines().map(str::to_owned).collect();
            publish(&paths, out_dir);
        },
        other => panic!("unknown subcommand {other:?}"),
    }
}
