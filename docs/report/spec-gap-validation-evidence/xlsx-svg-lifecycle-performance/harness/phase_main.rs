//! Exploratory phase decomposition for same-drawing SVG attach batches.
//!
//! This binary is deliberately separate from the acceptance profile binary.
//! It reports phase-local allocator observations while retaining the public
//! API operation and complete semantic validation from the acceptance path.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    reason = "the exploratory profile emits bounded JSON"
)]

mod adapter;
mod support;

use std::env;
use std::process::ExitCode;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;

fn usage() {
    println!(
        "usage: xlsx-svg-lifecycle-phase-profile --pictures 16|64|256 [--warmup N] [--samples N]"
    );
}

fn parse_args() -> Result<(usize, usize, usize), String> {
    let mut arguments = env::args().skip(1);
    let mut pictures = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                usage();
                return Err(String::from("help"));
            },
            "--pictures" => {
                pictures = Some(
                    arguments
                        .next()
                        .ok_or("--pictures requires a value")?
                        .parse()
                        .map_err(|_| String::from("--pictures must be an integer"))?,
                );
            },
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a value")?
                    .parse()
                    .map_err(|_| String::from("--warmup must be an integer"))?;
            },
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a value")?
                    .parse()
                    .map_err(|_| String::from("--samples must be an integer"))?;
            },
            unknown => return Err(format!("unknown argument: {unknown}")),
        }
    }
    let pictures = pictures.ok_or("--pictures is required")?;
    if !matches!(pictures, 16 | 64 | 256) {
        return Err(String::from("--pictures must be one of 16, 64, or 256"));
    }
    if samples == 0 {
        return Err(String::from("--samples must be nonzero"));
    }
    Ok((pictures, warmup, samples))
}

fn main() -> ExitCode {
    let arguments = match parse_args() {
        Ok(arguments) => arguments,
        Err(error) if error == "help" => return ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            return ExitCode::from(2);
        },
    };
    let (pictures, warmup, samples) = arguments;
    match adapter::run_phase_decomposition(pictures, warmup, samples) {
        Ok(receipt) => {
            println!("{receipt}");
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("phase decomposition failed: {error}");
            ExitCode::from(1)
        },
    }
}
