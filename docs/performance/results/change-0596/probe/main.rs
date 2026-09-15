//! Scratch survey driver for the doc-ppt survey (not part of the tree).
//!
//! survey-doc <root>  : refusal census + size drivers for every *.doc
//! survey-ppt <root>  : refusal census for every *.ppt
//! profile <mode> <path> <warmups> <samples> : loop for callgrind
//!   modes: doc-facade-open doc-facade-text doc-snapshot-open
//!          ppt-source-open ppt-source-text ppt-textedit-open

use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{FileSource, ReadAt};

type BoxError = Box<dyn std::error::Error>;

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.eq_ignore_ascii_case(ext))
                .unwrap_or(false)
            {
                out.push(path);
            }
        }
    }
}

fn short(err: impl std::fmt::Debug) -> String {
    let s = format!("{err:?}");
    let s = s.replace('\n', " ");
    if s.len() > 200 { format!("{}...", &s[..200]) } else { s }
}

fn file_source(path: &Path) -> Result<Arc<dyn ReadAt>, BoxError> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn survey_doc(root: &Path) {
    let mut files = Vec::new();
    walk(root, "doc", &mut files);
    files.sort();
    println!("file\tbytes\tsnapshot\tfacade\tword_len\ttable_len\tdata_len\ttext_units\tparagraphs\tclx_len\tpieces\tchpx_pages\tpapx_pages");
    for path in files {
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let snapshot = match file_source(&path) {
            Ok(source) => match litchi_doc::body_text::source::SourceSnapshot::open(source) {
                Ok(_) => "OK".to_string(),
                Err(e) => format!("ERR {}", short(e)),
            },
            Err(e) => format!("IO {}", short(e)),
        };
        let mut facade = "OK".to_string();
        let mut word_len = 0usize;
        let mut table_len = 0usize;
        let mut data_len = 0usize;
        let mut text_units = 0usize;
        let mut paragraphs = 0usize;
        let mut clx_len = 0u32;
        let mut pieces = 0usize;
        let mut chpx_pages = 0usize;
        let mut papx_pages = 0usize;
        match litchi_doc::Package::open(&path) {
            Ok(mut package) => {
                // stream lengths via the OLE file
                {
                    let ole = package.ole_file();
                    word_len = ole.open_stream(&["WordDocument"]).map(|v| v.len()).unwrap_or(0);
                    table_len = ole
                        .open_stream(&["1Table"])
                        .or_else(|_| ole.open_stream(&["0Table"]))
                        .map(|v| v.len())
                        .unwrap_or(0);
                    data_len = ole.open_stream(&["Data"]).map(|v| v.len()).unwrap_or(0);
                }
                match package.document() {
                    Ok(document) => {
                        if let Ok(text) = document.text() {
                            text_units = text.encode_utf16().count();
                        }
                        paragraphs = document.paragraph_count().unwrap_or(0);
                        let fib = document.fib();
                        if let Some((_, len)) = fib.get_table_pointer(33) {
                            clx_len = len;
                        }
                        if let Some((_, len)) = fib.get_table_pointer(12) {
                            chpx_pages = (len as usize).saturating_sub(4) / 8;
                        }
                        if let Some((_, len)) = fib.get_table_pointer(13) {
                            papx_pages = (len as usize).saturating_sub(4) / 8;
                        }
                        // piece count: parse the CLX from the table stream
                        let ole = package.ole_file();
                        let table_name = if fib.which_table_stream() { "1Table" } else { "0Table" };
                        if let (Ok(table), Some((off, len))) =
                            (ole.open_stream(&[table_name]), fib.get_table_pointer(33))
                        {
                            let (off, len) = (off as usize, len as usize);
                            if let Some(clx) = table.get(off..off + len)
                                && let Some(pt) = litchi_doc::parts::piece_table::PieceTable::parse(clx)
                            {
                                pieces = pt.pieces().len();
                            }
                        }
                    }
                    Err(e) => facade = format!("ERR {}", short(e)),
                }
            }
            Err(e) => facade = format!("ERR {}", short(e)),
        }
        println!(
            "{}\t{bytes}\t{snapshot}\t{facade}\t{word_len}\t{table_len}\t{data_len}\t{text_units}\t{paragraphs}\t{clx_len}\t{pieces}\t{chpx_pages}\t{papx_pages}",
            path.display()
        );
    }
}

fn survey_ppt(root: &Path) {
    let mut files = Vec::new();
    walk(root, "ppt", &mut files);
    files.sort();
    println!("file\tbytes\tsource_pkg\tslides\ttext_edit");
    for path in files {
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let mut slides = 0usize;
        let source_pkg = match litchi_ppt::SourceBackedPackage::from_path(&path) {
            Ok(package) => match package.presentation() {
                Ok(p) => {
                    slides = p.slide_count();
                    "OK".to_string()
                }
                Err(e) => format!("ERR {}", short(e)),
            },
            Err(e) => format!("ERR {}", short(e)),
        };
        let text_edit = match file_source(&path) {
            Ok(source) => match litchi_ppt::text_edit::SourceSnapshot::open(source) {
                Ok(_) => "OK".to_string(),
                Err(e) => format!("ERR {}", short(e)),
            },
            Err(e) => format!("IO {}", short(e)),
        };
        println!("{}\t{bytes}\t{source_pkg}\t{slides}\t{text_edit}", path.display());
    }
}

/// Open one document, then loop only the named read so the isolation pair
/// attributes that read rather than the open.
fn profile_read(
    mode: &str,
    path: &Path,
    warmups: usize,
    samples: usize,
) -> Result<(), BoxError> {
    let mut package = litchi_doc::Package::open(path)?;
    let document = package.document()?;
    let mut checksum = 0u64;
    for _ in 0..(warmups + samples) {
        let value: usize = match mode {
            "count" => document.paragraph_count()?,
            "ranges" => {
                let ranges = document.fib().get_all_subdoc_ranges();
                let n = ranges.len();
                black_box(ranges);
                n
            },
            "text" => {
                let text = document.text()?;
                let n = text.len();
                black_box(text);
                n
            },
            other => return Err(format!("unknown read mode {other}").into()),
        };
        checksum = checksum.wrapping_add(black_box(value) as u64);
    }
    println!(
        "{{\"mode\":\"read-{mode}\",\"iterations\":{},\"checksum\":{checksum}}}",
        warmups + samples
    );
    Ok(())
}

fn profile(mode: &str, path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let mut checksum = 0u64;
    for _ in 0..(warmups + samples) {
        let value: usize = match mode {
            "doc-facade-open" => {
                let mut package = litchi_doc::Package::open(path)?;
                let document = package.document()?;
                black_box(&document);
                0
            }
            "doc-facade-text" => {
                let mut package = litchi_doc::Package::open(path)?;
                let document = package.document()?;
                let text = document.text()?;
                let n = text.len();
                black_box(text);
                n
            }
            "doc-snapshot-open" => {
                let source = file_source(path)?;
                let snapshot = litchi_doc::body_text::source::SourceSnapshot::open(source)?;
                black_box(&snapshot);
                0
            }
            "ppt-source-open" => {
                let package = litchi_ppt::SourceBackedPackage::from_path(path)?;
                let presentation = package.presentation()?;
                let n = presentation.slide_count();
                black_box(&presentation);
                n
            }
            "ppt-source-text" => {
                let package = litchi_ppt::SourceBackedPackage::from_path(path)?;
                let presentation = package.presentation()?;
                let text = presentation.text()?;
                let n = text.len();
                black_box(text);
                n
            }
            "ppt-textedit-open" => {
                let source = file_source(path)?;
                let snapshot = litchi_ppt::text_edit::SourceSnapshot::open(source)?;
                black_box(&snapshot);
                0
            }
            other => return Err(format!("unknown mode {other}").into()),
        };
        checksum = checksum.wrapping_add(black_box(value) as u64);
    }
    println!("{{\"mode\":\"{mode}\",\"iterations\":{},\"checksum\":{checksum}}}", warmups + samples);
    Ok(())
}

fn tree_stats(records: &[litchi_ppt::Record], depth: usize, out: &mut (usize, usize, usize)) {
    for record in records {
        out.0 += 1;
        out.1 += record.data.len();
        out.2 = out.2.max(depth);
        tree_stats(&record.children, depth + 1, out);
    }
}

/// Copy factor of the eager PPT record tree: sum of owned payload bytes over the stream length.
fn ppt_tree(path: &Path) -> Result<(), BoxError> {
    let file = std::fs::File::open(path)?;
    let mut ole = litchi_cfb::OleFile::open(file)?;
    let data = ole.open_stream(&["PowerPoint Document"])?;
    let mut offset = 0usize;
    let mut top = Vec::new();
    while offset + 8 <= data.len() {
        match litchi_ppt::Record::parse(&data, offset) {
            Ok((record, consumed)) => {
                top.push(record);
                if consumed == 0 { break; }
                offset += consumed;
            }
            Err(_) => offset += 1,
        }
    }
    let mut stats = (0usize, 0usize, 0usize);
    tree_stats(&top, 1, &mut stats);
    println!(
        "{}\tstream_len={}\ttop_records={}\ttotal_records={}\towned_payload_bytes={}\tcopy_factor={:.2}\tmax_depth={}",
        path.display(), data.len(), top.len(), stats.0, stats.1, stats.1 as f64 / data.len() as f64, stats.2
    );
    Ok(())
}


/// FNV-1a 64-bit over a byte stream, used as a stable cross-binary digest.
const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
fn fnv1a64(seed: u64, bytes: &[u8]) -> u64 {
    let mut hash = seed;
    for &byte in bytes {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

/// Differential oracle: one line per admitted `.doc`, covering the decoded text,
/// every resolved paragraph (which carries its runs' character properties), the
/// section table, the FIB bytes the parser retains, and the exact refusal text
/// for anything that does not open.
fn digest_doc(root: &Path) {
    let mut files = Vec::new();
    walk(root, "doc", &mut files);
    files.sort();
    println!(
        "file\topen\ttext_len\ttext_hash\tpara_count\tcount_api\tpara_hash\tsection_count\tsection_hash\tsubdoc\tfib_len\tfib_hash"
    );
    for path in files {
        let rel = path.display().to_string();
        let mut package = match litchi_doc::Package::open(&path) {
            Ok(package) => package,
            Err(error) => {
                println!("{rel}\tPKGERR {}\t\t\t\t\t\t\t\t\t\t", short(error));
                continue;
            },
        };
        let document = match package.document() {
            Ok(document) => document,
            Err(error) => {
                println!("{rel}\tDOCERR {}\t\t\t\t\t\t\t\t\t\t", short(error));
                continue;
            },
        };
        let (text_len, text_hash) = match document.text() {
            Ok(text) => (
                text.len() as i64,
                fnv1a64(FNV_OFFSET, text.as_bytes()).to_string(),
            ),
            Err(error) => (-1, format!("ERR {}", short(error))),
        };
        let (para_count, para_hash) = match document.paragraphs() {
            Ok(paragraphs) => {
                let mut hash = FNV_OFFSET;
                for paragraph in &paragraphs {
                    hash = fnv1a64(hash, format!("{paragraph:?}").as_bytes());
                }
                (paragraphs.len() as i64, hash.to_string())
            },
            Err(error) => (-1, format!("ERR {}", short(error))),
        };
        let count_api = match document.paragraph_count() {
            Ok(count) => count as i64,
            Err(_) => -1,
        };
        let sections = document.sections();
        let section_hash = fnv1a64(FNV_OFFSET, format!("{sections:?}").as_bytes());
        let fib = document.fib();
        let subdoc = fnv1a64(
            FNV_OFFSET,
            format!("{:?}", fib.get_all_subdoc_ranges()).as_bytes(),
        );
        let raw = fib.raw_data();
        println!(
            "{rel}\tOK\t{text_len}\t{text_hash}\t{para_count}\t{count_api}\t{para_hash}\t{}\t{section_hash}\t{subdoc}\t{}\t{}",
            sections.len(),
            raw.len(),
            fnv1a64(FNV_OFFSET, raw)
        );
    }
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("survey-doc") => survey_doc(Path::new(&args[1])),
        Some("digest-doc") => digest_doc(Path::new(&args[1])),
        Some("survey-ppt") => survey_ppt(Path::new(&args[1])),
        Some("ppt-tree") => { for p in &args[1..] { ppt_tree(Path::new(p))?; } },
        Some("profile") => profile(&args[1], Path::new(&args[2]), args[3].parse()?, args[4].parse()?)?,
        Some("profile-read") => profile_read(&args[1], Path::new(&args[2]), args[3].parse()?, args[4].parse()?)?,
        _ => return Err("usage: survey-doc <root> | digest-doc <root> | survey-ppt <root> | profile <mode> <path> <warmups> <samples>".into()),
    }
    Ok(())
}
