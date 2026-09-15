//! Instruction-level profiling driver for the OLE2 DOC and PPT read paths.
//!
//! Each iteration constructs a *fresh* reader so that the open cost is paid
//! every time, which is what the callgrind isolation pair differences away.
//! Output is a single tiny JSON line so the child's own printing never shows
//! up in the profile.

use std::hint::black_box;
use std::path::PathBuf;

type BoxError = Box<dyn std::error::Error>;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Format {
    Doc,
    Ppt,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Operation {
    Open,
    Text,
}

struct Args {
    input: PathBuf,
    format: Format,
    operation: Operation,
    warmups: usize,
    samples: usize,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut input: Option<PathBuf> = None;
    let mut format: Option<Format> = None;
    let mut operation: Option<Operation> = None;
    let mut warmups: usize = 1;
    let mut samples: usize = 1;

    let mut argv = std::env::args().skip(1);
    while let Some(flag) = argv.next() {
        let mut value = || {
            argv.next()
                .ok_or_else(|| BoxError::from(format!("missing value for {flag}")))
        };
        match flag.as_str() {
            "--input" => input = Some(PathBuf::from(value()?)),
            "--format" => {
                format = Some(match value()?.as_str() {
                    "doc" => Format::Doc,
                    "ppt" => Format::Ppt,
                    other => return Err(format!("unknown --format {other}").into()),
                });
            }
            "--operation" => {
                operation = Some(match value()?.as_str() {
                    "open" => Operation::Open,
                    "text" => Operation::Text,
                    other => return Err(format!("unknown --operation {other}").into()),
                });
            }
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            other => return Err(format!("unknown flag {other}").into()),
        }
    }

    Ok(Args {
        input: input.ok_or("missing --input")?,
        format: format.ok_or("missing --format")?,
        operation: operation.ok_or("missing --operation")?,
        warmups,
        samples,
    })
}

fn doc_iteration(path: &std::path::Path, operation: Operation) -> Result<usize, BoxError> {
    let mut package = litchi_doc::Package::open(path)?;
    let document = package.document()?;
    let checksum = match operation {
        Operation::Open => {
            black_box(&document);
            0
        }
        Operation::Text => {
            let text = document.text()?;
            let length = text.len();
            black_box(text);
            length
        }
    };
    black_box(&package);
    Ok(checksum)
}

fn ppt_iteration(path: &std::path::Path, operation: Operation) -> Result<usize, BoxError> {
    let package = litchi_ppt::SourceBackedPackage::from_path(path)?;
    let presentation = package.presentation()?;
    let slides = presentation.slide_count();
    let checksum = match operation {
        Operation::Open => {
            black_box(&presentation);
            slides
        }
        Operation::Text => {
            let text = presentation.text()?;
            let length = text.len();
            black_box(text);
            slides + length
        }
    };
    black_box(&package);
    Ok(checksum)
}

fn main() -> Result<(), BoxError> {
    let args = parse_args()?;
    let path = args.input.as_path();
    let iterations = args.warmups + args.samples;
    let mut checksum: u64 = 0;

    for _ in 0..iterations {
        let value = match args.format {
            Format::Doc => doc_iteration(path, args.operation)?,
            Format::Ppt => ppt_iteration(path, args.operation)?,
        };
        checksum = checksum.wrapping_add(black_box(value) as u64);
    }

    println!(
        "{{\"iterations\":{iterations},\"warmups\":{},\"samples\":{},\"checksum\":{checksum}}}",
        args.warmups, args.samples
    );
    Ok(())
}
