//! Scratch preservation census for change 0660.
//!
//! Change 0652's decision 10 gives the DOCX save route a public compaction
//! policy whose default leaves every unmodified paragraph and part untouched.
//! This probe drives the documented DOCX save, edit and publication routes over
//! a whole corpus and prints one self-describing line per fixture per aspect, so
//! the two legs can be differenced byte for byte.
//!
//! Aspects, all through documented entry points:
//!   open            `Package::open(path)`
//!   gate            whether `/word/document.xml` satisfies the OPC writer's
//!                   authored-XML compactness contract (`verify_authored`)
//!   source          length and SHA-256 of `/word/document.xml` as opened
//!   noop_save       open, `to_stream` with no edit; SHA-256 of the artifact
//!   exact_noop      open, `edit_document`, `commit` with no operation,
//!                   `publish_document_commit`, `to_stream`; SHA-256
//!   managed_edit    open, `edit_document`, `insert_paragraph(0, ..)`,
//!                   `publish_document_edit`, `to_stream`; SHA-256
//!   one_edit        open, `edit_document`, `replace_paragraph_text` on the
//!                   first paragraph that has text, publish, `to_stream`;
//!                   SHA-256 (the default policy on the change leg)
//!   one_edit_optin  the same edit under `CompactionPolicy::WholeDocument`
//!                   (change leg only; the base leg prints `unavailable`)
//!   one_edit_span   source vs published `/word/document.xml`: both lengths,
//!                   the common prefix and the common suffix, so the bytes the
//!                   edit actually moved are visible
//!   one_edit_reopen reopen the published artifact: paragraph count and the
//!                   edited paragraph's text
//!
//! And one isolation-pair region, which prices only the code the policy
//! changes — `Snapshot::edit()`, one `replace_paragraph_text` and `commit()` —
//! on a generated document of a chosen size, either compact or carrying a line
//! break after the XML declaration so the publication gate sends it down the
//! whole-document fallback:
//!
//!   probe commit <paragraphs> <compact|noncompact> <measured>
//!
//! It builds a constant four snapshots and runs the region on `measured` of
//! them, so differencing `measured = 3` against `measured = 1` and halving
//! leaves two regions and nothing else.
//!
//! Usage:
//!   probe census <root>
//!   probe commit <paragraphs> <compact|noncompact> <measured>

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_core::Position;
use litchi_docx::Package;
use litchi_opc::packuri::PackURI;
use sha2::{Digest, Sha256};

const EXTENSIONS: &[&str] = &["docx", "docm", "dotx", "dotm"];
const MAIN_PART: &str = "/word/document.xml";
const EDIT_SUFFIX: &str = " litchi0660";

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe census <root>";
    match arguments.first().map(String::as_str) {
        Some("census") => {
            let root = arguments.get(1).expect(usage);
            let mut files = Vec::new();
            collect(Path::new(root), &mut files);
            files.sort();
            for file in files {
                census(&file, root);
            }
        },
        Some("commit") => {
            let paragraphs: usize = arguments.get(1).expect(usage).parse().expect(usage);
            let compact = match arguments.get(2).map(String::as_str) {
                Some("compact") => true,
                Some("noncompact") => false,
                _ => panic!("{usage}"),
            };
            let measured: usize = arguments.get(3).expect(usage).parse().expect(usage);
            commit_region(paragraphs, compact, measured);
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}

/// Build one generated main document, optionally with a line break between the
/// XML declaration and the root element.
fn generated_xml(paragraphs: usize, compact: bool) -> Vec<u8> {
    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    if !compact {
        xml.push('\n');
    }
    xml.push_str(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
    );
    for index in 0..paragraphs {
        xml.push_str(&format!(
            "<w:p><w:r><w:t>litchi-0660-source-{index:05}</w:t></w:r></w:p>"
        ));
    }
    xml.push_str("<w:sectPr/></w:body></w:document>");
    xml.into_bytes()
}

fn commit_region(paragraphs: usize, compact: bool, measured: usize) {
    let xml = generated_xml(paragraphs, compact);
    let snapshots: Vec<litchi_docx::document::Snapshot> = (0..4)
        .map(|_index| litchi_docx::document::Snapshot::from_xml(xml.clone()).expect("snapshot"))
        .collect();
    let mut total = 0usize;
    for snapshot in snapshots.iter().take(measured) {
        let mut edit = snapshot.edit();
        edit.replace_paragraph_text(Position::new(paragraphs / 2), "litchi 0660 edited")
            .expect("rewrite");
        let commit = edit.commit().expect("commit");
        total += commit.snapshot().xml_bytes().len();
    }
    println!("commit paragraphs={paragraphs} compact={compact} measured={measured} bytes={total}");
}

fn collect(directory: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, files);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| EXTENSIONS.contains(&extension.to_ascii_lowercase().as_str()))
        {
            files.push(path);
        }
    }
}

fn census(file: &Path, root: &str) {
    let name = file
        .strip_prefix(root)
        .unwrap_or(file)
        .to_string_lossy()
        .into_owned();

    let package = match Package::open(file) {
        Ok(package) => {
            println!("{name} open ok");
            package
        },
        Err(error) => {
            println!("{name} open refused:{}", one_line(&error.to_string()));
            for aspect in [
                "gate",
                "gate_source",
                "source",
                "noop_save",
                "exact_noop",
                "managed_edit",
                "one_edit",
                "one_edit_optin",
                "one_edit_span",
                "one_edit_reopen",
            ] {
                println!("{name} {aspect} skipped");
            }
            return;
        },
    };

    let source_xml = main_part_bytes(&package);
    match source_xml.as_ref() {
        Some(bytes) => {
            println!("{name} source bytes={} {}", bytes.len(), hash(bytes));
            match xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default())
            {
                Ok(_report) => println!("{name} gate compact"),
                Err(error) => println!("{name} gate noncompact:{}", one_line(&error.to_string())),
            }
            // Change 0665's aspect: the verdict the publication gate reaches
            // after the eager writer stopped asserting compactness.
            match xml_minifier::audit::verify_source(bytes, xml_minifier::audit::Limits::default())
            {
                Ok(_report) => println!("{name} gate_source accepts"),
                Err(error) => {
                    println!("{name} gate_source refuses:{}", one_line(&error.to_string()));
                },
            }
        },
        None => {
            println!("{name} source unavailable");
            println!("{name} gate unavailable");
            println!("{name} gate_source unavailable");
        },
    }
    drop(package);

    report(&name, "noop_save", noop_save(file));
    report(&name, "exact_noop", exact_noop(file));
    report(&name, "managed_edit", managed_edit(file));

    let default_edit = one_edit(file, false);
    match default_edit.as_ref() {
        Ok(bytes) => {
            println!("{name} one_edit {}", hash(bytes));
            span(&name, source_xml.as_deref(), bytes);
            reopen(&name, bytes);
        },
        Err(message) => {
            println!("{name} one_edit refused:{}", one_line(message));
            println!("{name} one_edit_span skipped");
            println!("{name} one_edit_reopen skipped");
        },
    }

    if cfg!(feature = "policy") {
        report(&name, "one_edit_optin", one_edit(file, true));
    } else {
        println!("{name} one_edit_optin unavailable");
    }
}

fn report(name: &str, aspect: &str, outcome: Result<Vec<u8>, String>) {
    match outcome {
        Ok(bytes) => println!("{name} {aspect} {}", hash(&bytes)),
        Err(message) => println!("{name} {aspect} refused:{}", one_line(&message)),
    }
}

fn noop_save(file: &Path) -> Result<Vec<u8>, String> {
    let mut package = Package::open(file).map_err(|error| error.to_string())?;
    let mut published = Vec::new();
    package
        .to_stream(&mut published)
        .map_err(|error| error.to_string())?;
    Ok(published)
}

fn exact_noop(file: &Path) -> Result<Vec<u8>, String> {
    let mut package = Package::open(file).map_err(|error| error.to_string())?;
    let edit = package.edit_document().map_err(|error| error.to_string())?;
    package
        .publish_document_edit(edit)
        .map_err(|error| error.to_string())?;
    let mut published = Vec::new();
    package
        .to_stream(&mut published)
        .map_err(|error| error.to_string())?;
    Ok(published)
}

fn managed_edit(file: &Path) -> Result<Vec<u8>, String> {
    let mut package = Package::open(file).map_err(|error| error.to_string())?;
    let mut edit = package.edit_document().map_err(|error| error.to_string())?;
    edit.insert_paragraph(Position::new(0), "litchi 0660 probe")
        .map_err(|error| error.to_string())?;
    package
        .publish_document_edit(edit)
        .map_err(|error| error.to_string())?;
    let mut published = Vec::new();
    package
        .to_stream(&mut published)
        .map_err(|error| error.to_string())?;
    Ok(published)
}

fn one_edit(file: &Path, whole_document: bool) -> Result<Vec<u8>, String> {
    let mut package = Package::open(file).map_err(|error| error.to_string())?;
    let snapshot = package
        .document_snapshot()
        .map_err(|error| error.to_string())?;
    let mut selected = None;
    for index in 0..snapshot.paragraph_count() {
        let position = Position::new(index);
        let Some(paragraph) = snapshot.paragraph(position) else {
            continue;
        };
        let Ok(text) = paragraph.text() else {
            continue;
        };
        if !text.is_empty() {
            selected = Some((position, text.to_string()));
            break;
        }
    }
    let (position, text) = selected.ok_or_else(|| "no paragraph with text".to_string())?;
    let mut edit = package.edit_document().map_err(|error| error.to_string())?;
    if whole_document {
        edit = with_whole_document_policy(edit);
    }
    let mut replacement = text;
    replacement.push_str(EDIT_SUFFIX);
    edit.replace_paragraph_text(position, replacement)
        .map_err(|error| error.to_string())?;
    package
        .publish_document_edit(edit)
        .map_err(|error| error.to_string())?;
    let mut published = Vec::new();
    package
        .to_stream(&mut published)
        .map_err(|error| error.to_string())?;
    Ok(published)
}

#[cfg(feature = "policy")]
fn with_whole_document_policy(edit: litchi_docx::document::Edit) -> litchi_docx::document::Edit {
    edit.with_compaction_policy(litchi_docx::document::CompactionPolicy::WholeDocument)
}

#[cfg(not(feature = "policy"))]
fn with_whole_document_policy(edit: litchi_docx::document::Edit) -> litchi_docx::document::Edit {
    edit
}

fn span(name: &str, source: Option<&[u8]>, published: &[u8]) {
    let Some(source) = source else {
        println!("{name} one_edit_span unavailable");
        return;
    };
    let package = match Package::from_reader(std::io::Cursor::new(published.to_vec())) {
        Ok(package) => package,
        Err(error) => {
            println!(
                "{name} one_edit_span reopen-refused:{}",
                one_line(&error.to_string())
            );
            return;
        },
    };
    let Some(after) = main_part_bytes(&package) else {
        println!("{name} one_edit_span unavailable");
        return;
    };
    let prefix = source
        .iter()
        .zip(after.iter())
        .take_while(|(left, right)| left == right)
        .count();
    let suffix = source[prefix..]
        .iter()
        .rev()
        .zip(after[prefix..].iter().rev())
        .take_while(|(left, right)| left == right)
        .count();
    println!(
        "{name} one_edit_span source={} published={} prefix={prefix} suffix={suffix} moved_source={} moved_published={}",
        source.len(),
        after.len(),
        source.len() - prefix - suffix,
        after.len() - prefix - suffix,
    );
}

fn reopen(name: &str, published: &[u8]) {
    let package = match Package::from_reader(std::io::Cursor::new(published.to_vec())) {
        Ok(package) => package,
        Err(error) => {
            println!(
                "{name} one_edit_reopen refused:{}",
                one_line(&error.to_string())
            );
            return;
        },
    };
    match package.document_snapshot() {
        Ok(snapshot) => {
            let mut edited = 0usize;
            for index in 0..snapshot.paragraph_count() {
                let Some(paragraph) = snapshot.paragraph(Position::new(index)) else {
                    continue;
                };
                if paragraph
                    .text()
                    .is_ok_and(|text| text.ends_with(EDIT_SUFFIX))
                {
                    edited += 1;
                }
            }
            println!(
                "{name} one_edit_reopen ok paragraphs={} tables={} controls={} edited={edited}",
                snapshot.paragraph_count(),
                snapshot.table_count(),
                snapshot.block_content_control_count(),
            );
        },
        Err(error) => println!(
            "{name} one_edit_reopen snapshot-refused:{}",
            one_line(&error.to_string())
        ),
    }
}

fn main_part_bytes(package: &Package) -> Option<Vec<u8>> {
    let uri = PackURI::new(MAIN_PART).ok()?;
    let part = package.opc_package().get_part(&uri).ok()?;
    Some(part.blob().to_vec())
}

fn hash(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    format!("{:x}", hasher.finalize())
}

fn one_line(message: &str) -> String {
    message.replace(['\n', '\r'], " ")
}
