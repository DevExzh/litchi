//! Change 0643 differential corpus check.
//!
//! Derived from the change-0592 corpus checker, which signed every
//! main-document read the lazy paragraph index could reach. This version adds
//! the reads change 0643 touches: `write_text_to` on both facades, under three
//! option sets, signing the exact sink bytes, the returned progress report and
//! — where the read fails — the full `TextOutputError` including its retained
//! partial progress. Errors are captured verbatim, so a moved refusal, a
//! changed message or a different amount of partial output changes the
//! signature.
//!
//! Usage: `corpus0643 <root>`

use std::fmt::Write as _;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{OwnedSource, TextOutputOptions};

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

/// A sink that keeps every byte handed to it, so the differential compares the
/// exact output and not only its length.
#[derive(Default)]
struct CapturingSink {
    bytes: Vec<u8>,
    writes: usize,
}

impl std::io::Write for CapturingSink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.writes += 1;
        self.bytes.extend_from_slice(buf);
        Ok(buf.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// A sink that refuses the third write, so the `Sink` failure branch and its
/// retained progress are signed too.
struct FailingSink {
    bytes: Vec<u8>,
    writes: usize,
}

impl std::io::Write for FailingSink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.writes += 1;
        if self.writes >= 3 {
            return Err(std::io::Error::other("corpus0643 sink refusal"));
        }
        self.bytes.extend_from_slice(buf);
        Ok(buf.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// The option sets the sink signature is taken under: the default policy, a
/// byte ceiling small enough to trip the limit on any non-trivial document,
/// and an object ceiling of one paragraph.
fn option_sets() -> [(&'static str, TextOutputOptions<'static>); 4] {
    [
        ("default", TextOutputOptions::default()),
        ("tight-bytes", TextOutputOptions::new("\n", "\n\n", 64, 1 << 20)),
        ("one-object", TextOutputOptions::new("\n", "\n\n", 1 << 20, 1)),
        (
            "no-empty",
            TextOutputOptions::new("|", "\n\n", 1 << 20, 1 << 20).with_empty_objects(false),
        ),
    ]
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
    for (name, options) in option_sets() {
        let mut sink = CapturingSink::default();
        let outcome = document.write_text_to(&mut sink, options);
        let _ = write!(
            out,
            "sink[{name}]:{:016x}:{}:{}:{}|",
            fnv1a(&sink.bytes),
            sink.bytes.len(),
            sink.writes,
            match outcome {
                Ok(report) => format!(
                    "ok({},{})",
                    report.bytes_written(),
                    report.objects_written()
                ),
                Err(error) => format!("err({error:?})"),
            }
        );
    }
    {
        let mut sink = FailingSink {
            bytes: Vec::new(),
            writes: 0,
        };
        let outcome = document.write_text_to(&mut sink, TextOutputOptions::default());
        let _ = write!(
            out,
            "sink[failing]:{:016x}:{}:{}:{}|",
            fnv1a(&sink.bytes),
            sink.bytes.len(),
            sink.writes,
            match outcome {
                Ok(report) => format!(
                    "ok({},{})",
                    report.bytes_written(),
                    report.objects_written()
                ),
                Err(error) => format!("err({error:?})"),
            }
        );
    }
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
    for (name, options) in option_sets() {
        let mut sink = CapturingSink::default();
        let outcome = package.write_text_to(&mut sink, options);
        let _ = write!(
            out,
            "sink[{name}]:{:016x}:{}:{}:{}|",
            fnv1a(&sink.bytes),
            sink.bytes.len(),
            sink.writes,
            match outcome {
                Ok(report) => format!(
                    "ok({},{})",
                    report.bytes_written(),
                    report.objects_written()
                ),
                Err(error) => format!("err({error:?})"),
            }
        );
    }
    {
        let mut sink = FailingSink {
            bytes: Vec::new(),
            writes: 0,
        };
        let outcome = package.write_text_to(&mut sink, TextOutputOptions::default());
        let _ = write!(
            out,
            "sink[failing]:{:016x}:{}:{}:{}|",
            fnv1a(&sink.bytes),
            sink.bytes.len(),
            sink.writes,
            match outcome {
                Ok(report) => format!(
                    "ok({},{})",
                    report.bytes_written(),
                    report.objects_written()
                ),
                Err(error) => format!("err({error:?})"),
            }
        );
    }
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
