//! Ceiling probe for change 0604: how much of an OLE2 open is the whole-stream
//! zero-fill?
//!
//! Every case runs `--samples` timed iterations after `--warmups` untimed ones
//! and prints one tiny JSON line, so the isolation-pair method (run N and N+M,
//! difference the `perf stat` totals, divide by M) removes process start-up,
//! the fixture read and the printing from the per-operation figure.
//!
//! Open cases build an OWNED in-memory source once, outside the loop, so the
//! measured operation is the open and nothing else.
//!
//! The two `slurp-*` cases model exactly what the design would change: today a
//! slurped stream is `alloc + memset + copy`; the appending shape is
//! `alloc + copy`. Their difference is the ceiling in cycles for the given
//! stream sizes. `memset` and `alloc-only` isolate the two halves of the
//! `alloc + memset` term on its own.

use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::source::OwnedSource;

type BoxError = Box<dyn std::error::Error>;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Case {
    DocOpen,
    PptOpen,
    PptSourceOpen,
    XlsSourceOpen,
    SlurpZero,
    SlurpAppend,
    Memset,
    AllocOnly,
}

struct Args {
    case: Case,
    input: Option<PathBuf>,
    bytes: Vec<usize>,
    warmups: usize,
    samples: usize,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut case: Option<Case> = None;
    let mut input: Option<PathBuf> = None;
    let mut bytes: Vec<usize> = Vec::new();
    let mut warmups: usize = 1;
    let mut samples: usize = 1;

    let mut argv = std::env::args().skip(1);
    while let Some(flag) = argv.next() {
        let mut value = || {
            argv.next()
                .ok_or_else(|| BoxError::from(format!("missing value for {flag}")))
        };
        match flag.as_str() {
            "--case" => {
                case = Some(match value()?.as_str() {
                    "doc-open" => Case::DocOpen,
                    "ppt-open" => Case::PptOpen,
                    "ppt-source-open" => Case::PptSourceOpen,
                    "xls-source-open" => Case::XlsSourceOpen,
                    "slurp-zero" => Case::SlurpZero,
                    "slurp-append" => Case::SlurpAppend,
                    "memset" => Case::Memset,
                    "alloc-only" => Case::AllocOnly,
                    other => return Err(format!("unknown --case {other}").into()),
                });
            },
            "--input" => input = Some(PathBuf::from(value()?)),
            "--bytes" => {
                for part in value()?.split(',') {
                    bytes.push(part.trim().parse()?);
                }
            },
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            other => return Err(format!("unknown flag {other}").into()),
        }
    }

    Ok(Args {
        case: case.ok_or("missing --case")?,
        input,
        bytes,
        warmups,
        samples,
    })
}

fn doc_iteration(bytes: &[u8]) -> Result<usize, BoxError> {
    let mut package = litchi_doc::Package::from_reader(Cursor::new(bytes))?;
    let document = package.document()?;
    let count = document.paragraph_count()?;
    black_box(&document);
    black_box(&package);
    Ok(count)
}

fn ppt_iteration(bytes: &[u8]) -> Result<usize, BoxError> {
    let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes))?;
    let presentation = package.presentation()?;
    let slides = presentation.slide_count();
    black_box(&presentation);
    black_box(&package);
    Ok(slides)
}

fn ppt_source_iteration(source: &Arc<Vec<u8>>) -> Result<usize, BoxError> {
    let owned: Arc<dyn litchi_core::source::ReadAt> =
        Arc::new(OwnedSource::from_arc(Arc::clone(source)));
    let package = litchi_ppt::SourceBackedPackage::from_read_at(owned)?;
    let presentation = package.presentation()?;
    let slides = presentation.slide_count();
    black_box(&presentation);
    black_box(&package);
    Ok(slides)
}

fn xls_source_iteration(source: &Arc<Vec<u8>>) -> Result<usize, BoxError> {
    let owned: Arc<dyn litchi_core::source::ReadAt> =
        Arc::new(OwnedSource::from_arc(Arc::clone(source)));
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(owned)?;
    let count = workbook.worksheet_count()?;
    black_box(&workbook);
    Ok(count)
}

/// Today's shape: reserve, zero-fill the whole stream, then overwrite it.
fn slurp_zero(sizes: &[usize], donor: &[u8]) -> Result<usize, BoxError> {
    let mut total = 0usize;
    for &size in sizes {
        let mut data: Vec<u8> = Vec::new();
        data.try_reserve_exact(size)?;
        data.resize(size, 0u8);
        data.copy_from_slice(&donor[..size]);
        total = total.wrapping_add(data[size - 1] as usize);
        black_box(&data);
    }
    Ok(total)
}

/// The designed shape: reserve, then append the physical runs.
fn slurp_append(sizes: &[usize], donor: &[u8]) -> Result<usize, BoxError> {
    let mut total = 0usize;
    for &size in sizes {
        let mut data: Vec<u8> = Vec::new();
        data.try_reserve_exact(size)?;
        data.extend_from_slice(&donor[..size]);
        total = total.wrapping_add(data[size - 1] as usize);
        black_box(&data);
    }
    Ok(total)
}

/// The zero-fill term alone: `try_reserve_exact` + `resize(len, 0)`, which is
/// exactly `try_zeroed_vec`/`try_filled_vec`.
fn memset_only(sizes: &[usize]) -> Result<usize, BoxError> {
    let mut total = 0usize;
    for &size in sizes {
        let mut data: Vec<u8> = Vec::new();
        data.try_reserve_exact(size)?;
        data.resize(size, 0u8);
        total = total.wrapping_add(data[size - 1] as usize);
        black_box(&data);
    }
    Ok(total)
}

/// The allocator term alone, so the zero-fill can be separated from it.
fn alloc_only(sizes: &[usize]) -> Result<usize, BoxError> {
    let mut total = 0usize;
    for &size in sizes {
        let mut data: Vec<u8> = Vec::new();
        data.try_reserve_exact(size)?;
        total = total.wrapping_add(data.capacity());
        black_box(&data);
    }
    Ok(total)
}

fn main() -> Result<(), BoxError> {
    let args = parse_args()?;
    let iterations = args.warmups + args.samples;

    let file_bytes: Arc<Vec<u8>> = match args.input.as_ref() {
        Some(path) => Arc::new(std::fs::read(path)?),
        None => Arc::new(Vec::new()),
    };
    let donor: Vec<u8> = if args.bytes.is_empty() {
        Vec::new()
    } else {
        let max = args.bytes.iter().copied().max().unwrap_or(0);
        (0..max).map(|index| (index % 251) as u8).collect()
    };

    let mut checksum: u64 = 0;
    let started = Instant::now();
    for _ in 0..iterations {
        let value = match args.case {
            Case::DocOpen => doc_iteration(file_bytes.as_slice())?,
            Case::PptOpen => ppt_iteration(file_bytes.as_slice())?,
            Case::PptSourceOpen => ppt_source_iteration(&file_bytes)?,
            Case::XlsSourceOpen => xls_source_iteration(&file_bytes)?,
            Case::SlurpZero => slurp_zero(&args.bytes, &donor)?,
            Case::SlurpAppend => slurp_append(&args.bytes, &donor)?,
            Case::Memset => memset_only(&args.bytes)?,
            Case::AllocOnly => alloc_only(&args.bytes)?,
        };
        checksum = checksum.wrapping_add(value as u64);
    }
    let elapsed = started.elapsed();

    println!(
        "{{\"iterations\":{iterations},\"checksum\":{checksum},\"elapsed_ns\":{}}}",
        elapsed.as_nanos()
    );
    Ok(())
}
