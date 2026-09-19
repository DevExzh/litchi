//! Standalone exact-output oracle for the shared OOXML markup-compatibility
//! processor.
//!
//! The binary intentionally calls only the public `mce` API.  It reads one
//! XML file, selects a named capability/limit profile, and emits one stable
//! record containing either the exact output SHA-256/length/report or the
//! typed error's `Debug` representation.  A separate Python driver invokes
//! the same binary built from the frozen baseline and the candidate source.

use litchi_ooxml_common::mce::{self, Capabilities, Limits, Name};
use sha2::{Digest, Sha256};
use std::{env, error::Error, fs, hint::black_box, path::Path, process, time::Instant};

const PROBE_ID: &str = "0698-mce-oracle";
const MAX_WARMUPS: usize = 100_000;
const MAX_SAMPLES: usize = 100_000;
const OPAQUE_NAMESPACE: &str = "urn:litchi:oracle:opaque";
const OPAQUE_LOCAL_NAME: &str = "payload";

type Fallible<T> = Result<T, Box<dyn Error + Send + Sync>>;

fn main() {
    if let Err(error) = run() {
        eprintln!("{PROBE_ID}: {error}");
        process::exit(2);
    }
}

fn run() -> Fallible<()> {
    let mut args = env::args().skip(1);
    let mode = args.next().unwrap_or_else(|| "run".to_owned());
    match mode.as_str() {
        "profiles" => {
            println!("baseline\tdefault capabilities\tdefault limits");
            println!("opaque\tdefault capabilities + {OPAQUE_NAMESPACE}#{OPAQUE_LOCAL_NAME}\tdefault limits");
            println!("opaque-small\tdefault capabilities + {OPAQUE_NAMESPACE}#{OPAQUE_LOCAL_NAME}\tsmall bounded limits");
            println!("opaque-large\tdefault capabilities + {OPAQUE_NAMESPACE}#{OPAQUE_LOCAL_NAME}\tlarge bounded limits");
            println!("opaque-many\tdefault capabilities + 4096 explicit opaque names\tdefault limits");
        },
        "run" => {
            let profile = args.next().ok_or("usage: run <profile> <xml-path>")?;
            let path = args.next().ok_or("usage: run <profile> <xml-path>")?;
            if args.next().is_some() {
                return Err("usage: run <profile> <xml-path>".into());
            }
            run_once(&profile, Path::new(&path))?;
        },
        "time" => {
            let profile = args.next().ok_or(
                "usage: time <profile> <xml-path> <warmups> <samples>",
            )?;
            let path = args.next().ok_or(
                "usage: time <profile> <xml-path> <warmups> <samples>",
            )?;
            let warmups = parse_count(
                args.next().ok_or("usage: time <profile> <xml-path> <warmups> <samples>")?,
                "warmups",
                MAX_WARMUPS,
            )?;
            let samples = parse_count(
                args.next().ok_or("usage: time <profile> <xml-path> <warmups> <samples>")?,
                "samples",
                MAX_SAMPLES,
            )?;
            if args.next().is_some() {
                return Err("usage: time <profile> <xml-path> <warmups> <samples>".into());
            }
            time_profile(&profile, Path::new(&path), warmups, samples)?;
        },
        _ => {
            return Err(format!(
                "unknown mode {mode:?}; use profiles, run, or time"
            )
            .into());
        },
    }
    Ok(())
}

fn parse_count(value: String, label: &str, maximum: usize) -> Fallible<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed > maximum {
        return Err(format!("{label} exceeds maximum {maximum}: {parsed}").into());
    }
    Ok(parsed)
}

fn read_xml(path: &Path) -> Fallible<Vec<u8>> {
    Ok(fs::read(path).map_err(|error| {
        format!("cannot read XML {}: {error}", path.display())
    })?)
}

fn profile(name: &str) -> Fallible<(Capabilities, Limits)> {
    let mut capabilities = Capabilities::default();
    let mut limits = Limits::default();
    match name {
        "baseline" => {},
        "opaque" => add_opaque_capability(&mut capabilities),
        "opaque-small" => {
            add_opaque_capability(&mut capabilities);
            limits.max_input_bytes = 4 * 1024;
            limits.max_output_bytes = 8 * 1024;
            limits.max_depth = 32;
            limits.max_namespace_bindings = 64;
            limits.max_directive_tokens = 16;
            limits.max_choices_per_alternate = 8;
        },
        "opaque-large" => {
            add_opaque_capability(&mut capabilities);
            limits.max_input_bytes = 8 * 1024 * 1024;
            limits.max_output_bytes = 16 * 1024 * 1024;
            limits.max_depth = 1024;
            limits.max_namespace_bindings = 16 * 1024;
            limits.max_directive_tokens = 16 * 1024;
            limits.max_choices_per_alternate = 4096;
        },
        "opaque-many" => {
            add_many_opaque_capabilities(&mut capabilities);
        },
        other => return Err(format!("unknown profile {other:?}; run profiles").into()),
    }
    Ok((capabilities, limits))
}

fn add_opaque_capability(capabilities: &mut Capabilities) {
    capabilities.preserve_extension_element(Name {
        namespace: OPAQUE_NAMESPACE.to_owned(),
        local_name: OPAQUE_LOCAL_NAME.to_owned(),
    });
}

fn add_many_opaque_capabilities(capabilities: &mut Capabilities) {
    add_opaque_capability(capabilities);
    for index in 1..4096 {
        capabilities.preserve_extension_element(Name {
            namespace: format!("{OPAQUE_NAMESPACE}:{index}"),
            local_name: OPAQUE_LOCAL_NAME.to_owned(),
        });
    }
}

fn digest(bytes: &[u8]) -> String {
    let mut result = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        result.push_str(&format!("{byte:02x}"));
    }
    result
}

fn run_once(profile_name: &str, path: &Path) -> Fallible<()> {
    let input = read_xml(path)?;
    let (capabilities, limits) = profile(profile_name)?;
    match mce::process_markup_compatibility(&input, &capabilities, &limits) {
        Ok(output) => {
            let borrowed = matches!(output.xml, std::borrow::Cow::Borrowed(_));
            println!(
                "OK\tprobe={PROBE_ID}\tprofile={profile_name}\tinput_len={}\toutput_len={}\toutput_sha256={}\tborrowed={borrowed}\treport={:?}",
                input.len(),
                output.xml.len(),
                digest(output.xml.as_ref()),
                output.report,
            );
        },
        Err(error) => {
            println!(
                "ERR\tprobe={PROBE_ID}\tprofile={profile_name}\tinput_len={}\tdebug={error:?}",
                input.len(),
            );
        },
    }
    Ok(())
}

fn time_profile(
    profile_name: &str,
    path: &Path,
    warmups: usize,
    samples: usize,
) -> Fallible<()> {
    let input = read_xml(path)?;
    let (capabilities, limits) = profile(profile_name)?;
    let process = || mce::process_markup_compatibility(black_box(&input), &capabilities, &limits);
    for _ in 0..warmups {
        black_box(process().map_err(|error| format!("warmup refused: {error:?}"))?);
    }
    println!(
        "TIMING\tprobe={PROBE_ID}\tprofile={profile_name}\tinput_len={}\twarmups={warmups}\tsamples={samples}",
        input.len(),
    );
    for index in 0..samples {
        let start = Instant::now();
        let output = process().map_err(|error| format!("sample refused: {error:?}"))?;
        black_box(output);
        println!("SAMPLE\tindex={index}\telapsed_ns={}", start.elapsed().as_nanos());
    }
    Ok(())
}
