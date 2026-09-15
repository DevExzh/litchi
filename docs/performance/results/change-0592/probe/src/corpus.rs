//! Change 0592 differential corpus check.
//!
//! For every `.docx` under the given root this prints one line holding a
//! deterministic signature of every main-document read the change touches:
//! open outcome, `text()`, `paragraph_count()`, every `paragraph(i)` including
//! one past the end, `paragraphs()`, `tables()`, `elements()` and `blocks()`,
//! for both the eager and the source-backed facade. Errors are captured as
//! their `Display` text, so a moved refusal changes the signature.
//!
//! Usage: `corpus0592 <root>`

use std::fmt::Write as _;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::OwnedSource;

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

fn collect(root: &Path, found: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    let mut entries: Vec<_> = entries.filter_map(std::result::Result::ok).collect();
    entries.sort_by_key(std::fs::DirEntry::path);
    for entry in entries {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, found);
        } else if path.extension().is_some_and(|value| value == "docx") {
            found.push(path);
        }
    }
}

fn eager_signature(bytes: &[u8]) -> String {
    let mut out = String::new();
    let package = match litchi_docx::Package::from_reader(Cursor::new(bytes.to_vec())) {
        Ok(package) => package,
        Err(error) => return format!("open-error:{error}"),
    };
    let document = match package.document() {
        Ok(document) => document,
        Err(error) => return format!("document-error:{error}"),
    };
    let _ = write!(
        out,
        "text:{}|",
        match document.text() {
            Ok(text) => format!("{:016x}:{}", fnv1a(text.as_bytes()), text.len()),
            Err(error) => format!("err:{error}"),
        }
    );
    let count = document.paragraph_count();
    let _ = write!(
        out,
        "count:{}|",
        match &count {
            Ok(value) => value.to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "paragraphs:{}|",
        match document.paragraphs() {
            Ok(paragraphs) => {
                let mut joined = String::new();
                for paragraph in &paragraphs {
                    let _ = write!(
                        joined,
                        "{}\u{1}",
                        match paragraph.text() {
                            Ok(text) => text,
                            Err(error) => format!("err:{error}"),
                        }
                    );
                }
                format!("{:016x}:{}", fnv1a(joined.as_bytes()), paragraphs.len())
            },
            Err(error) => format!("err:{error}"),
        }
    );
    let limit = count.as_ref().copied().unwrap_or(0).min(4096);
    let mut selected = String::new();
    for index in 0..=limit {
        let _ = write!(
            selected,
            "{}\u{1}",
            match document.paragraph(index) {
                Ok(Some(paragraph)) => match paragraph.text() {
                    Ok(text) => text,
                    Err(error) => format!("err:{error}"),
                },
                Ok(None) => "<none>".to_owned(),
                Err(error) => format!("err:{error}"),
            }
        );
    }
    let _ = write!(out, "selected:{:016x}:{limit}|", fnv1a(selected.as_bytes()));
    let _ = write!(
        out,
        "tables:{}|",
        match document.tables() {
            Ok(tables) => tables.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "elements:{}|",
        match document.elements() {
            Ok(elements) => elements.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "blocks:{}",
        match document.blocks() {
            Ok(blocks) => blocks.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    out
}

fn source_signature(bytes: &[u8]) -> String {
    let mut out = String::new();
    let package = match litchi_docx::source_backed::Package::from_read_at(Arc::new(
        OwnedSource::new(bytes.to_vec()),
    )) {
        Ok(package) => package,
        Err(error) => return format!("open-error:{error}"),
    };
    let document = match package.document() {
        Ok(document) => document,
        Err(error) => return format!("document-error:{error}"),
    };
    let _ = write!(
        out,
        "text:{}|",
        match document.extract_text() {
            Ok(text) => format!("{:016x}:{}", fnv1a(text.as_bytes()), text.len()),
            Err(error) => format!("err:{error}"),
        }
    );
    let count = document.paragraph_count();
    let _ = write!(
        out,
        "count:{}|",
        match &count {
            Ok(value) => value.to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "paragraphs:{}|",
        match document.paragraphs() {
            Ok(paragraphs) => {
                let mut joined = String::new();
                for paragraph in &paragraphs {
                    let _ = write!(
                        joined,
                        "{}\u{1}",
                        match paragraph.text() {
                            Ok(text) => text,
                            Err(error) => format!("err:{error}"),
                        }
                    );
                }
                format!("{:016x}:{}", fnv1a(joined.as_bytes()), paragraphs.len())
            },
            Err(error) => format!("err:{error}"),
        }
    );
    let limit = count.as_ref().copied().unwrap_or(0).min(4096);
    let mut selected = String::new();
    let mut texts = String::new();
    for index in 0..=limit {
        let _ = write!(
            selected,
            "{}\u{1}",
            match document.paragraph(index) {
                Ok(Some(paragraph)) => match paragraph.text() {
                    Ok(text) => text,
                    Err(error) => format!("err:{error}"),
                },
                Ok(None) => "<none>".to_owned(),
                Err(error) => format!("err:{error}"),
            }
        );
        let _ = write!(
            texts,
            "{}\u{1}",
            match document.paragraph_text(index) {
                Ok(Some(text)) => text,
                Ok(None) => "<none>".to_owned(),
                Err(error) => format!("err:{error}"),
            }
        );
    }
    let _ = write!(out, "selected:{:016x}:{limit}|", fnv1a(selected.as_bytes()));
    let _ = write!(out, "paragraph_text:{:016x}|", fnv1a(texts.as_bytes()));
    let _ = write!(
        out,
        "tables:{}|",
        match document.tables() {
            Ok(tables) => tables.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "elements:{}|",
        match document.elements() {
            Ok(elements) => elements.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    let _ = write!(
        out,
        "blocks:{}",
        match document.blocks() {
            Ok(blocks) => blocks.len().to_string(),
            Err(error) => format!("err:{error}"),
        }
    );
    out
}

fn main() {
    let root = std::env::args().nth(1).expect("corpus root");
    let mut fixtures = Vec::new();
    collect(Path::new(&root), &mut fixtures);
    fixtures.sort();
    println!("fixtures={}", fixtures.len());
    for fixture in fixtures {
        let relative = fixture
            .strip_prefix(&root)
            .unwrap_or(&fixture)
            .to_string_lossy()
            .into_owned();
        let Ok(bytes) = std::fs::read(&fixture) else {
            println!("{relative}\tread-error");
            continue;
        };
        println!(
            "{relative}\teager[{}]\tsource[{}]",
            eager_signature(&bytes),
            source_signature(&bytes)
        );
    }
}
