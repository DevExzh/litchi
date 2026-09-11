//! Process-isolated XLSX ordinary-worksheet SVG lifecycle profile adapter.
//!
//! This binary is deliberately kept under the evidence tree. It links the
//! current public `litchi_xlsx` selector, source scanner, and transaction
//! methods, while the shell runner still requires an explicit freeze gate
//! before it may collect any receipt.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in profile owns synthetic fixtures and emits JSON"
)]

mod adapter;
mod support;

use std::env;
use std::process::ExitCode;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;

fn usage() {
    println!(
        "usage: xlsx-svg-lifecycle-profile --lane NAME [--warmup N] [--samples N]\n\n{} acceptance lanes plus {} exploratory lanes are wired to the current XLSX API; the shell runner remains freeze-gated",
        adapter::LANES.len(),
        adapter::EXPLORATORY_LANES.len(),
    );
}

fn parse_args() -> Result<(String, usize, usize), String> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                usage();
                return Err(String::from("help"));
            },
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
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
    let lane = lane.ok_or("--lane is required")?;
    if !adapter::is_known_lane(&lane) {
        return Err(format!("unknown lane: {lane}"));
    }
    if samples == 0 {
        return Err(String::from("--samples must be nonzero"));
    }
    Ok((lane, warmup, samples))
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
    let (lane, warmup, samples) = arguments;
    match adapter::run(&lane, warmup, samples) {
        Ok(receipt) => {
            println!("{receipt}");
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("profile fixture or adapter setup failed: {error}");
            ExitCode::from(1)
        },
    }
}
