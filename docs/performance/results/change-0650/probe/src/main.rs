//! Scratch admission probe for change 0650.
//!
//! Change 0638 observed that `litchi_docx::Package::open` admits
//! `alt-chunk-header.docx` while `Package::document_mut()` refuses it. No
//! harness selector reports the two admissions side by side, so this probe
//! drives the documented DOCX read and edit routes over a whole corpus and
//! prints one self-describing line per fixture per aspect.
//!
//! Aspects, all through documented entry points:
//!   open          `Package::open(path)`
//!   reader        `Package::document()` then `paragraph_count()` and `text()`
//!   reader_bom    whether the part blob and the reader's normalized view
//!                 still carry a leading UTF-8 byte order mark
//!   sections      `Package::document()` then `sections()` (the reader's own
//!                 section parse, which `document_mut` runs eagerly)
//!   document_mut  `Package::document_mut()` (the editor's re-parse)
//!   noop_save     open, `save(tmp)` with no edit; SHA-256 of the artifact
//!   edit_save     open, `document_mut().add_paragraph_with_text(..)`, `save(tmp)`
//!   touch_save    open, `document_mut()` with no edit, `save(tmp)`
//!   managed_edit  open, `edit_document()`, `insert_paragraph(0, ..)`,
//!                 `publish_document_edit`, `save(tmp)` (the managed route)
//!   reopen        `Package::open` of the edited artifact, then `document()`
//!
//! Usage:
//!   probe census <root> <scratch-dir>   one line per fixture and aspect
//!   probe fromxml <file>                `MutableDocument::from_xml` on raw bytes
//!   probe partxml <docx> <out>          write `/word/document.xml` as opened

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_core::Position;
use litchi_docx::parts::DocumentPart;
use litchi_docx::{MutableDocument, Package};
use litchi_opc::packuri::PackURI;
use sha2::{Digest, Sha256};

const BYTE_ORDER_MARK: &[u8] = &[0xEF, 0xBB, 0xBF];
const EXTENSIONS: &[&str] = &["docx", "docm", "dotx", "dotm"];

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe census <root> <scratch> | probe fromxml <file> | probe partxml <docx> <out>";
    match arguments.first().map(String::as_str) {
        Some("census") => {
            let root = arguments.get(1).expect(usage);
            let scratch = PathBuf::from(arguments.get(2).expect(usage));
            fs::create_dir_all(&scratch).expect("scratch");
            let mut files = Vec::new();
            collect(Path::new(root), &mut files);
            files.sort();
            for file in files {
                census(&file, root, &scratch);
            }
        },
        Some("fromxml") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let bytes = fs::read(&file).expect("read");
            match std::str::from_utf8(&bytes) {
                Err(error) => println!("{} from_xml utf8-error:{error}", file.display()),
                Ok(text) => match MutableDocument::from_xml(text) {
                    Ok(_) => println!("{} from_xml ok", file.display()),
                    Err(error) => println!("{} from_xml refused:{error}", file.display()),
                },
            }
        },
        Some("partxml") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let out = PathBuf::from(arguments.get(2).expect(usage));
            let package = Package::open(&file).expect("open");
            let uri = PackURI::new("/word/document.xml").expect("uri");
            let part = package.opc_package().get_part(&uri).expect("part");
            fs::write(&out, part.blob()).expect("write");
            println!("{} partxml {} bytes", file.display(), part.blob().len());
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}

fn census(file: &Path, root: &str, scratch: &Path) {
    let name = file
        .strip_prefix(root)
        .unwrap_or(file)
        .to_string_lossy()
        .into_owned();

    let mut package = match Package::open(file) {
        Ok(package) => {
            println!("{name} open ok");
            package
        },
        Err(error) => {
            println!("{name} open refused:{}", one_line(&error.to_string()));
            println!("{name} reader skipped");
            println!("{name} reader_bom skipped");
            println!("{name} sections skipped");
            println!("{name} document_mut skipped");
            println!("{name} noop_save skipped");
            println!("{name} touch_save skipped");
            println!("{name} managed_edit skipped");
            println!("{name} edit_save skipped");
            println!("{name} reopen skipped");
            return;
        },
    };

    match package.document() {
        Ok(document) => println!(
            "{name} reader ok paragraphs={} text={}",
            count_outcome(&document),
            text_outcome(&document)
        ),
        Err(error) => println!("{name} reader refused:{}", one_line(&error.to_string())),
    }

    {
        let uri = PackURI::new("/word/document.xml").ok();
        let blob_bom = uri
            .as_ref()
            .and_then(|uri| package.opc_package().get_part(uri).ok())
            .map(|part| part.blob().starts_with(BYTE_ORDER_MARK));
        let view_bom = package
            .opc_package()
            .main_document_part()
            .ok()
            .and_then(|part| DocumentPart::from_part(part).ok())
            .map(|part| part.xml_bytes().starts_with(BYTE_ORDER_MARK));
        match (blob_bom, view_bom) {
            (Some(blob), Some(view)) => println!("{name} reader_bom part={blob} view={view}"),
            _ => println!("{name} reader_bom unavailable"),
        }
    }

    match package.document() {
        Ok(document) => match document.sections() {
            Ok(sections) => println!("{name} sections ok count={}", sections.len()),
            Err(error) => println!("{name} sections refused:{}", one_line(&error.to_string())),
        },
        Err(error) => println!("{name} sections open-refused:{}", one_line(&error.to_string())),
    }

    match package.document_mut() {
        Ok(_) => println!("{name} document_mut ok"),
        Err(error) => println!(
            "{name} document_mut refused:{}",
            one_line(&error.to_string())
        ),
    }

    // A fresh open per save aspect: `document_mut` above may have installed a
    // mutable document, and the no-op arm must publish what the open produced.
    let flat = name.replace(['/', '\\'], "_");
    let noop_path = scratch.join(format!("{flat}.noop"));
    match Package::open(file).and_then(|mut package| {
        package.save(&noop_path)?;
        Ok(())
    }) {
        Ok(()) => println!("{name} noop_save {}", digest(&noop_path)),
        Err(error) => println!("{name} noop_save refused:{}", one_line(&error.to_string())),
    }

    let touch_path = scratch.join(format!("{flat}.touch"));
    match Package::open(file).and_then(|mut package| {
        package.document_mut()?;
        package.save(&touch_path)?;
        Ok(())
    }) {
        Ok(()) => println!("{name} touch_save {}", digest(&touch_path)),
        Err(error) => println!("{name} touch_save refused:{}", one_line(&error.to_string())),
    }
    let _ = fs::remove_file(&touch_path);

    let managed_path = scratch.join(format!("{flat}.managed"));
    match managed_edit(file, &managed_path) {
        Ok(()) => println!("{name} managed_edit {}", digest(&managed_path)),
        Err(message) => println!("{name} managed_edit refused:{}", one_line(&message)),
    }
    let _ = fs::remove_file(&managed_path);

    let edit_path = scratch.join(format!("{flat}.edit"));
    let edited = Package::open(file).and_then(|mut package| {
        package
            .document_mut()?
            .add_paragraph_with_text("litchi 0650 probe");
        package.save(&edit_path)?;
        Ok(())
    });
    match edited {
        Ok(()) => {
            println!("{name} edit_save {}", digest(&edit_path));
            match Package::open(&edit_path) {
                Ok(package) => match package.document() {
                    Ok(document) => println!(
                        "{name} reopen ok paragraphs={} text={}",
                        count_outcome(&document),
                        text_outcome(&document)
                    ),
                    Err(error) => {
                        println!("{name} reopen reader-refused:{}", one_line(&error.to_string()));
                    },
                },
                Err(error) => {
                    println!("{name} reopen open-refused:{}", one_line(&error.to_string()));
                },
            }
        },
        Err(error) => {
            println!("{name} edit_save refused:{}", one_line(&error.to_string()));
            println!("{name} reopen skipped");
        },
    }
    let _ = fs::remove_file(&noop_path);
    let _ = fs::remove_file(&edit_path);
}

fn managed_edit(file: &Path, out: &Path) -> std::result::Result<(), String> {
    let mut package = Package::open(file).map_err(|error| error.to_string())?;
    let mut edit = package.edit_document().map_err(|error| error.to_string())?;
    edit.insert_paragraph(Position::new(0), "litchi 0650 probe")
        .map_err(|error| error.to_string())?;
    package
        .publish_document_edit(edit)
        .map_err(|error| error.to_string())?;
    package.save(out).map_err(|error| error.to_string())?;
    Ok(())
}

fn count_outcome(document: &litchi_docx::Document<'_>) -> String {
    match document.paragraph_count() {
        Ok(count) => format!("{count}"),
        Err(error) => format!("refused:{}", one_line(&error.to_string())),
    }
}

fn text_outcome(document: &litchi_docx::Document<'_>) -> String {
    match document.text() {
        Ok(text) => format!("{}bytes", text.len()),
        Err(error) => format!("refused:{}", one_line(&error.to_string())),
    }
}

fn digest(path: &Path) -> String {
    match fs::read(path) {
        Ok(bytes) => {
            let mut hasher = Sha256::new();
            hasher.update(&bytes);
            format!("{} bytes={}", hex(&hasher.finalize()), bytes.len())
        },
        Err(error) => format!("unreadable:{error}"),
    }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn one_line(value: &str) -> String {
    value.replace(['\n', '\r'], " ")
}

fn collect(root: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, out);
        } else if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| EXTENSIONS.contains(&value.to_ascii_lowercase().as_str()))
        {
            out.push(path);
        }
    }
}
