//! Process-isolated XLSX SVG lifecycle profile scaffold.
//!
//! This binary intentionally has no `litchi-xlsx` dependency while the
//! source-backed owner and public API are provisional. The future adapter will
//! retain the lane names and receipt support, then wire the timed closures to
//! the frozen API. Running this binary today is an explicit refusal rather
//! than a synthetic or guessed measurement.

#![allow(clippy::print_stderr)]

use std::env;
use std::process::ExitCode;

#[path = "support.rs"]
mod support;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;

const LANES: &[&str] = &[
    "capture_raster_two_cell_small",
    "capture_raster_two_cell_large",
    "capture_raster_one_cell_small",
    "capture_raster_one_cell_large",
    "capture_raster_absolute_small",
    "capture_raster_absolute_large",
    "capture_attached_two_cell_small",
    "capture_attached_two_cell_large",
    "capture_attached_one_cell_small",
    "capture_attached_one_cell_large",
    "capture_attached_absolute_small",
    "capture_attached_absolute_large",
    "clone_raster_small",
    "clone_raster_large",
    "clone_attached_small",
    "clone_attached_large",
    "inventory_shared_256",
    "inventory_shared_1024",
    "inventory_distinct_256",
    "inventory_distinct_1024",
    "namespace_heavy",
    "namespace_limit_refusal",
    "attach_end_to_end_two_cell_small",
    "attach_end_to_end_two_cell_large",
    "attach_end_to_end_one_cell_small",
    "attach_end_to_end_one_cell_large",
    "attach_end_to_end_absolute_small",
    "attach_end_to_end_absolute_large",
    "detach_end_to_end_shared_first_two_cell",
    "detach_end_to_end_shared_first_one_cell",
    "detach_end_to_end_shared_first_absolute",
    "detach_end_to_end_shared_final_two_cell",
    "detach_end_to_end_shared_final_one_cell",
    "detach_end_to_end_shared_final_absolute",
    "detach_end_to_end_distinct_two_cell_small",
    "detach_end_to_end_distinct_two_cell_large",
    "detach_end_to_end_distinct_one_cell_small",
    "detach_end_to_end_distinct_one_cell_large",
    "detach_end_to_end_distinct_absolute_small",
    "detach_end_to_end_distinct_absolute_large",
    "noop_detach_two_cell",
    "noop_detach_one_cell",
    "noop_detach_absolute",
    "limit_small",
    "limit_large",
    "malformed_duplicate_owner",
    "malformed_mce_owner",
    "malformed_linked_owner",
    "malformed_unknown_uri",
];

fn usage() {
    println!(
        "usage: xlsx-svg-lifecycle-profile --lane NAME [--warmup N] [--samples N]\n\n{} lanes are reserved; API wiring is pending",
        LANES.len()
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
    if !LANES.contains(&lane.as_str()) {
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
    eprintln!(
        "api-wiring-pending: lane={lane} warmup={warmup} samples={samples}; no XLSX API is called"
    );
    eprintln!(
        "wire the frozen source-backed worksheet selector/attach/detach API before collecting receipts"
    );
    ExitCode::from(2)
}
