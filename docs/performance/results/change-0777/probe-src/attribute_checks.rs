//! Record 0770: fail-fast attribute checks, before and after.
//!
//! The same source builds against the base and the changed library, so it
//! uses only APIs both expose.
//!
//! * `differential --json PATH [--mutate-limit BYTES] ROOT...` opens every
//!   OOXML package under the roots with the OPC reader and, by its main part,
//!   the DOCX, PPTX or XLSX reader, and records a digest of each outcome (the
//!   text or cells read, or the error's message). It then mutates start tags of
//!   chosen members of every package no larger than the limit (a repeated
//!   name, a repeated name after 40 more names, malformed syntax) and records
//!   the same readers' outcomes on each mutated package, so two builds can be
//!   compared outcome by outcome.
//! * `probe --case NAME --n N --samples S --warmup W --json PATH` builds one
//!   package whose start tag at a fail-fast site carries `N` attributes, then
//!   times `W + S` opens and records each run's wall time and outcome.

#![forbid(unsafe_code)]

use std::collections::BTreeMap;
use std::error::Error;
use std::fmt::Write as _;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::time::Instant;

use quick_xml::events::BytesStart;
use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("differential") => differential(&args[1..]),
        Some("probe") => probe(&args[1..]),
        _ => Err("usage: attribute_checks differential|probe ...".into()),
    }
}

fn option<'a>(args: &'a [String], name: &str) -> Option<&'a str> {
    args.iter()
        .position(|arg| arg == name)
        .and_then(|index| args.get(index + 1))
        .map(String::as_str)
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

// --------------------------------------------------------------------------
// Readers
// --------------------------------------------------------------------------

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Format {
    Docx,
    Pptx,
    Xlsx,
    Other,
}

fn format_of(names: &[String]) -> Format {
    if names.iter().any(|name| name == "word/document.xml") {
        Format::Docx
    } else if names.iter().any(|name| name == "ppt/presentation.xml") {
        Format::Pptx
    } else if names.iter().any(|name| name == "xl/workbook.xml") {
        Format::Xlsx
    } else {
        Format::Other
    }
}

fn outcome<T, E: std::fmt::Display>(
    result: Result<T, E>,
    digest: impl FnOnce(T) -> String,
) -> String {
    match result {
        Ok(value) => format!("ok:{}", digest(value)),
        Err(error) => format!("err:{error}"),
    }
}

fn read_opc(bytes: &[u8]) -> String {
    outcome(litchi_opc::OpcPackage::from_bytes(bytes), |package| {
        // The package keeps its parts in a hash map, whose order differs from
        // one process to the next: digest them in name order.
        let mut parts: Vec<String> = package
            .iter_parts()
            .map(|part| format!("{} {}\n", part.partname(), part.content_type()))
            .collect();
        parts.sort_unstable();
        sha256_hex(parts.concat().as_bytes())
    })
}

fn read_docx(bytes: &[u8]) -> String {
    let text = litchi_docx::Package::from_reader(Cursor::new(bytes.to_vec()))
        .map_err(|error| error.to_string())
        .and_then(|package| {
            package
                .document()
                .and_then(|document| document.text())
                .map_err(|error| error.to_string())
        });
    outcome(text, |text| sha256_hex(text.as_bytes()))
}

fn read_pptx(bytes: &[u8]) -> String {
    let text = litchi_pptx::Package::from_bytes(bytes)
        .map_err(|error| error.to_string())
        .and_then(|package| {
            package
                .presentation()
                .and_then(|presentation| presentation.text())
                .map_err(|error| error.to_string())
        });
    outcome(text, |text| sha256_hex(text.as_bytes()))
}

fn read_xlsx(bytes: &[u8]) -> String {
    let cells = litchi_xlsx::Workbook::from_bytes(bytes.to_vec())
        .map_err(|error| error.to_string())
        .and_then(|workbook| {
            let mut summary = String::new();
            for sheet in workbook.sheets() {
                let cells = sheet
                    .cells("A1:XFD1048576")
                    .map_err(|error| error.to_string())?;
                for cell in cells {
                    let _ = writeln!(summary, "{cell:?}");
                }
                summary.push('\n');
            }
            Ok::<_, String>(summary)
        });
    outcome(cells, |summary| sha256_hex(summary.as_bytes()))
}

fn read_all(bytes: &[u8], format: Format, results: &mut BTreeMap<String, String>, key: &str) {
    results.insert(format!("{key}::opc"), read_opc(bytes));
    let (label, outcome) = match format {
        Format::Docx => ("docx", read_docx(bytes)),
        Format::Pptx => ("pptx", read_pptx(bytes)),
        Format::Xlsx => ("xlsx", read_xlsx(bytes)),
        Format::Other => return,
    };
    results.insert(format!("{key}::{label}"), outcome);
}

// --------------------------------------------------------------------------
// Differential
// --------------------------------------------------------------------------

fn differential(args: &[String]) -> Result<(), Box<dyn Error>> {
    let json = option(args, "--json").ok_or("--json is required")?;
    let mutate_limit: usize = option(args, "--mutate-limit")
        .unwrap_or("1048576")
        .parse()?;
    let mut roots = Vec::new();
    let mut skip = false;
    for arg in args {
        if skip {
            skip = false;
            continue;
        }
        if arg.starts_with("--") {
            skip = true;
            continue;
        }
        roots.push(PathBuf::from(arg));
    }
    let mut packages = Vec::new();
    for root in &roots {
        collect_packages(root, &mut packages)?;
    }
    packages.sort();
    let mut results = BTreeMap::new();
    let mut formats = BTreeMap::<String, usize>::new();
    let mut mutated = 0usize;
    let mut mutated_packages = 0usize;
    for path in &packages {
        let Ok(bytes) = std::fs::read(path) else {
            continue;
        };
        let Ok(archive) = ArchiveReader::new(&bytes) else {
            continue;
        };
        let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        let format = format_of(&names);
        *formats.entry(format!("{format:?}")).or_default() += 1;
        let key = path.display().to_string();
        read_all(&bytes, format, &mut results, &key);
        if bytes.len() > mutate_limit {
            continue;
        }
        mutated_packages += 1;
        for member in mutation_members(&names, format) {
            let Ok(xml) = archive.read(&member) else {
                continue;
            };
            for (tag_index, tag) in start_tags(&xml).into_iter().take(3).enumerate() {
                for (label, mutation) in mutations(&xml, &tag) {
                    let Some(package) = replace_member(&archive, &names, &member, &mutation) else {
                        continue;
                    };
                    mutated += 1;
                    read_all(
                        &package,
                        format,
                        &mut results,
                        &format!("{key}::{member}::tag{tag_index}::{label}"),
                    );
                }
            }
        }
    }
    let summary = serde_json::json!({
        "packages": packages.len(),
        "formats": formats,
        "mutated_packages": mutated_packages,
        "mutations": mutated,
        "outcomes": results.len(),
        "results": results,
    });
    std::fs::write(json, serde_json::to_vec(&summary)?)?;
    Ok(())
}

fn collect_packages(root: &Path, packages: &mut Vec<PathBuf>) -> Result<(), Box<dyn Error>> {
    let metadata = std::fs::symlink_metadata(root)?;
    if metadata.is_dir() {
        let mut entries: Vec<_> = std::fs::read_dir(root)?.filter_map(Result::ok).collect();
        entries.sort_by_key(std::fs::DirEntry::path);
        for entry in entries {
            collect_packages(&entry.path(), packages)?;
        }
    } else if metadata.is_file() {
        let Ok(bytes) = std::fs::read(root) else {
            return Ok(());
        };
        if bytes.starts_with(b"PK\x03\x04")
            && ArchiveReader::new(&bytes)
                .is_ok_and(|archive| archive.contains("[Content_Types].xml"))
        {
            packages.push(root.to_owned());
        }
    }
    Ok(())
}

/// The members whose start tags are mutated: the package manifests and the
/// parts every reader of the format reads.
fn mutation_members(names: &[String], format: Format) -> Vec<String> {
    let mut members = vec!["[Content_Types].xml".to_owned(), "_rels/.rels".to_owned()];
    let main: &[&str] = match format {
        Format::Docx => &[
            "word/document.xml",
            "word/styles.xml",
            "word/_rels/document.xml.rels",
            "word/numbering.xml",
        ],
        Format::Pptx => &[
            "ppt/presentation.xml",
            "ppt/slides/slide1.xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/slideLayouts/slideLayout1.xml",
        ],
        Format::Xlsx => &[
            "xl/workbook.xml",
            "xl/worksheets/sheet1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ],
        Format::Other => &[],
    };
    members.extend(main.iter().map(|name| (*name).to_owned()));
    members.retain(|member| names.contains(member));
    members
}

/// A start tag with attributes: the byte range of its content (after `<`,
/// before `>` or `/>`) and the names of its attributes.
struct Tag {
    content_end: usize,
    first: Vec<u8>,
    last: Vec<u8>,
}

/// The first start tags of `xml` that carry attributes, in document order.
fn start_tags(xml: &[u8]) -> Vec<Tag> {
    let mut tags = Vec::new();
    let mut index = 0;
    while index < xml.len() && tags.len() < 8 {
        let Some(open) = memchr(b'<', &xml[index..]).map(|found| index + found) else {
            break;
        };
        let start = open + 1;
        let Some(first) = xml.get(start) else {
            break;
        };
        if matches!(first, b'/' | b'?' | b'!') {
            index = start;
            continue;
        }
        // The end of the tag, outside quoted values.
        let mut quote = None;
        let mut end = None;
        for (offset, byte) in xml[start..].iter().enumerate() {
            match (quote, *byte) {
                (None, b'"' | b'\'') => quote = Some(*byte),
                (Some(open_quote), byte) if byte == open_quote => quote = None,
                (None, b'>') => {
                    end = Some(start + offset);
                    break;
                },
                _ => {},
            }
        }
        let Some(end) = end else {
            break;
        };
        let content_end = if xml[end - 1] == b'/' { end - 1 } else { end };
        let content = &xml[start..content_end];
        let name_len = content
            .iter()
            .position(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
            .unwrap_or(content.len());
        let Ok(content) = std::str::from_utf8(content) else {
            index = end + 1;
            continue;
        };
        let tag = BytesStart::from_content(content, name_len);
        let mut attributes = tag.attributes();
        attributes.with_checks(false);
        let keys: Vec<Vec<u8>> = attributes
            .map_while(Result::ok)
            .map(|attribute| attribute.key.as_ref().to_vec())
            .collect();
        if let (Some(first), Some(last)) = (keys.first(), keys.last()) {
            tags.push(Tag {
                content_end,
                first: first.clone(),
                last: last.clone(),
            });
        }
        index = end + 1;
    }
    tags
}

fn memchr(needle: u8, haystack: &[u8]) -> Option<usize> {
    haystack.iter().position(|byte| *byte == needle)
}

/// The mutations of one start tag: text appended to the tag's attributes.
fn mutations(xml: &[u8], tag: &Tag) -> Vec<(&'static str, Vec<u8>)> {
    let first = String::from_utf8_lossy(&tag.first).into_owned();
    let last = String::from_utf8_lossy(&tag.last).into_owned();
    let mut padding = String::new();
    for index in 0..40 {
        let _ = write!(padding, " zz{index}=\"{index}\"");
    }
    let appended: [(&'static str, String); 10] = [
        ("repeat_first", format!(" {first}=\"repeat\"")),
        ("repeat_last", format!(" {last}=\"repeat\"")),
        (
            "padded_repeat_first",
            format!("{padding} {first}=\"repeat\""),
        ),
        ("padded_repeat_padding", format!("{padding} zz3=\"repeat\"")),
        (
            "padded_repeat_unquoted",
            format!("{padding} {first}=repeat"),
        ),
        ("padded_unquoted", format!("{padding} zzq=value")),
        ("padded", padding.clone()),
        ("key_only", " zzflag".to_owned()),
        ("padded_key_only", format!("{padding} zzflag")),
        (
            "repeat_with_space_in_value",
            format!(" {first}=\"x {last}='y'\" zzt=\"t\""),
        ),
    ];
    appended
        .into_iter()
        .map(|(label, text)| {
            let mut mutated = Vec::with_capacity(xml.len() + text.len());
            mutated.extend_from_slice(&xml[..tag.content_end]);
            mutated.extend_from_slice(text.as_bytes());
            mutated.extend_from_slice(&xml[tag.content_end..]);
            (label, mutated)
        })
        .collect()
}

/// A copy of the package with one member replaced, stored uncompressed.
fn replace_member(
    archive: &ArchiveReader<'_>,
    names: &[String],
    member: &str,
    replacement: &[u8],
) -> Option<Vec<u8>> {
    let mut writer = StreamingArchiveWriter::new();
    for name in names {
        if name == member {
            writer.write_stored(name, replacement).ok()?;
        } else {
            writer.write_stored(name, &archive.read(name).ok()?).ok()?;
        }
    }
    writer.finish_to_bytes().ok()
}

// --------------------------------------------------------------------------
// Probes
// --------------------------------------------------------------------------

const CONTENT_TYPES: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    r#"</Types>"#
);

/// A package whose one package relationship carries `n` namespace
/// declarations, which the OPC relationship reader accepts and reads one by
/// one: the fail-fast site in `pkgreader`.
fn relationship_declarations(n: usize) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut rels = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="urn:probe" Target="probe.xml""#
    ));
    for index in 0..n {
        let _ = write!(rels, " xmlns:p{index}=\"u\"");
    }
    rels.push_str("/></Relationships>");
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", CONTENT_TYPES.as_bytes())?;
    writer.write_stored("_rels/.rels", rels.as_bytes())?;
    writer.write_stored("probe.xml", b"<probe/>")?;
    Ok(writer.finish_to_bytes()?)
}

fn probe(args: &[String]) -> Result<(), Box<dyn Error>> {
    let case = option(args, "--case").ok_or("--case is required")?;
    let n: usize = option(args, "--n").ok_or("--n is required")?.parse()?;
    let samples: usize = option(args, "--samples").unwrap_or("15").parse()?;
    let warmup: usize = option(args, "--warmup").unwrap_or("3").parse()?;
    let json = option(args, "--json").ok_or("--json is required")?;
    let input = match case {
        "opc_relationship_declarations" => relationship_declarations(n)?,
        _ => return Err(format!("unknown case {case}").into()),
    };
    let mut durations = Vec::with_capacity(samples);
    let mut outcomes = Vec::new();
    for run in 0..warmup + samples {
        let started = Instant::now();
        let result = read_opc(&input);
        let elapsed = started.elapsed();
        std::hint::black_box(&result);
        if run >= warmup {
            durations.push(elapsed.as_nanos());
        }
        if !outcomes.contains(&result) {
            outcomes.push(result);
        }
    }
    let summary = serde_json::json!({
        "case": case,
        "n": n,
        "input_bytes": input.len(),
        "samples": samples,
        "warmup": warmup,
        "durations_ns": durations,
        "outcomes": outcomes,
    });
    std::fs::write(json, serde_json::to_vec(&summary)?)?;
    Ok(())
}
