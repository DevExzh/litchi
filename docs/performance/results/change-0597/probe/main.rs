//! Scratch probe for change 0597 (selected-cell ineligibility gate).
//!
//! Modes:
//!   source <pkg.xlsx> <addr>   one source-backed cell read (callgrind subject);
//!                              same shape as the 0587 survey probe.
//!   eager  <pkg.xlsx> <addr>   one eager cell read (cross-check).
//!   oracle <dir|file> ...      deterministic transcript over every `.xlsx`
//!                              found: per (sheet, address) a FRESH
//!                              source-backed workbook over a counting
//!                              positional source, printing the exact
//!                              `Result<SourceCellView>` debug, the eager
//!                              result debug, and the logical read/byte
//!                              counts of the timed read.
//!   ir <pkg.xlsx> <addr> <n>   repeat the source-backed one-cell read n times
//!                              over fresh workbooks (callgrind isolation pair).

use std::env;
use std::fmt::Write as _;
use std::hint::black_box;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::source::{ReadAt, SourceVersion};

/// In-memory positional source counting logical reads and bytes.
struct CountingSource {
    bytes: Vec<u8>,
    reads: AtomicU64,
    read_bytes: AtomicU64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
        }
    }

    fn counts(&self) -> (u64, u64) {
        (
            self.reads.load(Ordering::Relaxed),
            self.read_bytes.load(Ordering::Relaxed),
        )
    }

    fn reset(&self) {
        self.reads.store(0, Ordering::Relaxed);
        self.read_bytes.store(0, Ordering::Relaxed);
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        let start = usize::try_from(offset).unwrap_or(usize::MAX).min(self.bytes.len());
        let available = &self.bytes[start..];
        let take = available.len().min(output.len());
        output[..take].copy_from_slice(&available[..take]);
        self.read_bytes.fetch_add(take as u64, Ordering::Relaxed);
        Ok(take)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(1, 1))
    }
}

const ADDRESSES: &[&str] = &[
    "A1", "A2", "B1", "B2", "C3", "D4", "E5", "H680", "M29", "Z100", "AA1", "BA50",
];
const RANGES: &[&str] = &["A1:E5", "B1:B50"];
const MAX_SHEETS: usize = 3;

fn collect_xlsx(root: &Path, out: &mut Vec<PathBuf>) {
    if root.is_file() {
        out.push(root.to_path_buf());
        return;
    }
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    let mut paths: Vec<PathBuf> = entries.filter_map(|e| e.ok()).map(|e| e.path()).collect();
    paths.sort();
    for path in paths {
        if path.is_dir() {
            collect_xlsx(&path, out);
        } else if path.extension().is_some_and(|ext| ext == "xlsx") {
            out.push(path);
        }
    }
}

/// Normalize a debug string so the transcript is comparable between legs.
fn norm(text: &str) -> String {
    text.replace('\n', " ")
}

fn oracle(paths: &[PathBuf]) {
    let mut transcript = String::new();
    for path in paths {
        let label = path.display().to_string();
        let Ok(bytes) = std::fs::read(path) else {
            let _ = writeln!(transcript, "{label}\tREADFAIL");
            continue;
        };
        // Sheet catalog and the eager oracle come from one open each.
        let catalog_source = Arc::new(CountingSource::new(bytes.clone()));
        let names: Vec<String> = match litchi_xlsx::SourceBackedWorkbook::from_read_at(
            catalog_source.clone() as Arc<dyn ReadAt>,
        ) {
            Ok(workbook) => workbook
                .sheets()
                .take(MAX_SHEETS)
                .map(|sheet| sheet.name().to_owned())
                .collect(),
            Err(error) => {
                let _ = writeln!(transcript, "{label}\tOPENERR\t{}", norm(&format!("{error:?}")));
                continue;
            },
        };
        let _ = writeln!(transcript, "{label}\tSHEETS\t{}", names.join("|"));

        for name in &names {
            for address in ADDRESSES {
                let source = Arc::new(CountingSource::new(bytes.clone()));
                let Ok(workbook) =
                    litchi_xlsx::SourceBackedWorkbook::from_read_at(source.clone() as Arc<dyn ReadAt>)
                else {
                    let _ = writeln!(transcript, "{label}\t{name}\t{address}\tREOPENERR");
                    continue;
                };
                let Some(sheet) = workbook.sheets().find(|sheet| sheet.name() == name) else {
                    let _ = writeln!(transcript, "{label}\t{name}\t{address}\tNOSHEET");
                    continue;
                };
                source.reset();
                let outcome = sheet.cell(*address);
                let (reads, read_bytes) = source.counts();
                let _ = writeln!(
                    transcript,
                    "{label}\t{name}\t{address}\tcell\treads={reads}\tbytes={read_bytes}\t{}",
                    norm(&format!("{outcome:?}"))
                );
            }
            for range in RANGES {
                let source = Arc::new(CountingSource::new(bytes.clone()));
                let Ok(workbook) =
                    litchi_xlsx::SourceBackedWorkbook::from_read_at(source.clone() as Arc<dyn ReadAt>)
                else {
                    let _ = writeln!(transcript, "{label}\t{name}\t{range}\tREOPENERR");
                    continue;
                };
                let Some(sheet) = workbook.sheets().find(|sheet| sheet.name() == name) else {
                    let _ = writeln!(transcript, "{label}\t{name}\t{range}\tNOSHEET");
                    continue;
                };
                source.reset();
                let outcome = sheet.cells(*range);
                let (reads, read_bytes) = source.counts();
                let _ = writeln!(
                    transcript,
                    "{label}\t{name}\t{range}\tcells\treads={reads}\tbytes={read_bytes}\t{}",
                    norm(&format!("{outcome:?}"))
                );
            }
        }
    }
    print!("{transcript}");
}

fn main() {
    let args: Vec<String> = env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("oracle");
    match mode {
        "source" => {
            let path = args.get(2).expect("path");
            let addr = args.get(3).map(String::as_str).unwrap_or("H680");
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path).expect("open");
            let sheet = workbook.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("source ok");
        },
        "eager" => {
            let path = args.get(2).expect("path");
            let addr = args.get(3).map(String::as_str).unwrap_or("H680");
            let workbook = litchi_xlsx::Workbook::open(path).expect("open");
            let sheet = workbook.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("eager ok");
        },
        "ir" => {
            let path = args.get(2).expect("path");
            let addr = args.get(3).map(String::as_str).unwrap_or("H680");
            let repeats: usize = args.get(4).map_or(1, |value| value.parse().expect("n"));
            let bytes = std::fs::read(path).expect("read");
            for _ in 0..repeats {
                let source = Arc::new(CountingSource::new(bytes.clone()));
                let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(
                    source.clone() as Arc<dyn ReadAt>
                )
                .expect("open");
                let sheet = workbook.sheets().next().expect("sheet");
                let view = sheet.cell(addr).expect("cell");
                black_box(view);
                let (reads, read_bytes) = source.counts();
                black_box((reads, read_bytes));
            }
            println!("ir ok {repeats}");
        },
        "eagoracle" => {
            let mut paths = Vec::new();
            for root in args.iter().skip(2) {
                collect_xlsx(Path::new(root), &mut paths);
            }
            paths.sort();
            let mut transcript = String::new();
            for path in &paths {
                let label = path.display().to_string();
                let workbook = match litchi_xlsx::Workbook::open(path) {
                    Ok(workbook) => workbook,
                    Err(error) => {
                        let _ = writeln!(
                            transcript,
                            "{label}\tOPENERR\t{}",
                            norm(&format!("{error:?}"))
                        );
                        continue;
                    },
                };
                let names: Vec<String> = workbook
                    .sheets()
                    .take(MAX_SHEETS)
                    .map(|sheet| sheet.name().to_owned())
                    .collect();
                for name in &names {
                    let Some(sheet) = workbook.sheets().find(|sheet| sheet.name() == name) else {
                        continue;
                    };
                    for address in ADDRESSES {
                        let outcome = sheet.cell(*address).map(|view| match view {
                            litchi_xlsx::cell::View::Missing => "Missing".to_owned(),
                            litchi_xlsx::cell::View::Covered(range) => format!("Covered({range:?})"),
                            litchi_xlsx::cell::View::Stored(cell) => format!("Stored({cell:?})"),
                            other => format!("{other:?}"),
                        });
                        let _ = writeln!(
                            transcript,
                            "{label}\t{name}\t{address}\tcell\t{}",
                            norm(&format!("{outcome:?}"))
                        );
                    }
                }
            }
            print!("{transcript}");
        },
        "oracle" => {
            let mut paths = Vec::new();
            for root in args.iter().skip(2) {
                collect_xlsx(Path::new(root), &mut paths);
            }
            paths.sort();
            oracle(&paths);
        },
        other => panic!("unknown mode {other}"),
    }
}
