//! Protocol-only probe for the paired DrawingML raw qualification lane.
//!
//! The five DrawingML owners are landing on separate source branches.  This
//! probe is therefore deliberately an adapter boundary: until the coordinator
//! supplies a lane adapter built against the frozen candidate, it emits
//! explicit `unavailable` rows and never manufactures timing or fidelity
//! evidence.  The command line and JSON shape are stable so the source pair,
//! allocator observer, and validator can be reviewed before those adapters
//! land.

#![forbid(unsafe_code)]

use std::env;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

const LANES: &[&str] = &[
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
];

const OPERATIONS: &[&str] = &[
    "entry_read",
    "read_shared_when_supported",
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
];

#[derive(Debug)]
struct Arguments {
    lane: String,
    role: String,
    mode: String,
    fixture: PathBuf,
    source_revision: String,
    iterations: usize,
    warmups: usize,
    blocks: usize,
    samples: usize,
    target_ms: u64,
    minimum_ms: u64,
    output: PathBuf,
}

fn usage() -> &'static str {
    "usage: drawingml-raw-{native,allocator} --lane LANE --role ROLE \
     --fixture PATH --source-revision SHA --iterations N --warmups N \
     --blocks N --samples N --target-ms N --minimum-ms N --json PATH"
}

fn argument_value(arguments: &mut impl Iterator<Item = String>, name: &str) -> io::Result<String> {
    arguments
        .next()
        .ok_or_else(|| invalid_argument(format!("missing value for {name}")))
}

fn invalid_argument(message: String) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message)
}

fn parse(arguments: impl Iterator<Item = String>, allocator: bool) -> io::Result<Arguments> {
    let mut lane = None;
    let mut role = None;
    let mut mode = None;
    let mut fixture = None;
    let mut source_revision = None;
    let mut iterations = None;
    let mut warmups = None;
    let mut blocks = None;
    let mut samples = None;
    let mut target_ms = None;
    let mut minimum_ms = None;
    let mut output = None;
    let mut arguments = arguments;
    while let Some(argument) = arguments.next() {
        let slot = match argument.as_str() {
            "--lane" => &mut lane,
            "--role" => &mut role,
            "--mode" => &mut mode,
            "--fixture" => &mut fixture,
            "--source-revision" => &mut source_revision,
            "--iterations" => &mut iterations,
            "--warmups" => &mut warmups,
            "--blocks" => &mut blocks,
            "--samples" => &mut samples,
            "--target-ms" => &mut target_ms,
            "--minimum-ms" => &mut minimum_ms,
            "--json" => &mut output,
            _ => return Err(invalid_argument(format!("unknown argument {argument:?}"))),
        };
        if slot.is_some() {
            return Err(invalid_argument(format!("duplicate argument {argument}")));
        }
        *slot = Some(argument_value(&mut arguments, &argument)?);
    }

    let lane = lane.ok_or_else(|| invalid_argument("--lane is required".into()))?;
    if !LANES.contains(&lane.as_str()) {
        return Err(invalid_argument(format!("unsupported DrawingML lane {lane:?}")));
    }
    let role = role.ok_or_else(|| invalid_argument("--role is required".into()))?;
    if !matches!(role.as_str(), "baseline" | "candidate") {
        return Err(invalid_argument(format!("unsupported role {role:?}")));
    }
    let expected_mode = if allocator { "allocator" } else { "normal" };
    let mode = mode.ok_or_else(|| invalid_argument("--mode is required".into()))?;
    if mode != expected_mode {
        return Err(invalid_argument(format!(
            "{expected_mode} binary requires --mode {expected_mode}"
        )));
    }
    let fixture = PathBuf::from(
        fixture.ok_or_else(|| invalid_argument("--fixture is required".into()))?,
    );
    if fixture.is_symlink() || !fixture.is_file() {
        return Err(invalid_argument(format!("fixture is not a regular file: {fixture:?}")));
    }
    let source_revision = source_revision
        .ok_or_else(|| invalid_argument("--source-revision is required".into()))?;
    if source_revision.len() != 40
        || !source_revision
            .bytes()
            .all(|byte| byte.is_ascii_hexdigit() && !byte.is_ascii_uppercase())
    {
        return Err(invalid_argument(
            "--source-revision must be a lowercase 40-character SHA".into(),
        ));
    }
    let parse_usize = |value: Option<String>, name: &str| -> io::Result<usize> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<usize>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let parse_u64 = |value: Option<String>, name: &str| -> io::Result<u64> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<u64>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let iterations = parse_usize(iterations, "--iterations")?;
    let warmups = parse_usize(warmups, "--warmups")?;
    let blocks = parse_usize(blocks, "--blocks")?;
    let samples = parse_usize(samples, "--samples")?;
    let target_ms = parse_u64(target_ms, "--target-ms")?;
    let minimum_ms = parse_u64(minimum_ms, "--minimum-ms")?;
    if iterations == 0 || warmups != 3 || blocks != 5 || samples != 20 {
        return Err(invalid_argument(
            "iterations must be positive, warmups=3, blocks=5, samples=20".into(),
        ));
    }
    if target_ms < minimum_ms || minimum_ms < 150 {
        return Err(invalid_argument(
            "target-ms must be at least minimum-ms and minimum-ms must be >= 150".into(),
        ));
    }
    let output = PathBuf::from(output.ok_or_else(|| invalid_argument("--json is required".into()))?);
    if output.as_os_str().is_empty() || output.is_symlink() || output.exists() {
        return Err(invalid_argument(format!(
            "JSON output must be a fresh regular-file path: {output:?}"
        )));
    }
    Ok(Arguments {
        lane,
        role,
        mode,
        fixture,
        source_revision,
        iterations,
        warmups,
        blocks,
        samples,
        target_ms,
        minimum_ms,
        output,
    })
}

fn json_string(value: &str) -> String {
    let mut result = String::with_capacity(value.len() + 2);
    result.push('"');
    for character in value.chars() {
        match character {
            '"' => result.push_str("\\\""),
            '\\' => result.push_str("\\\\"),
            '\n' => result.push_str("\\n"),
            '\r' => result.push_str("\\r"),
            '\t' => result.push_str("\\t"),
            character if character.is_control() => {
                use std::fmt::Write as _;
                let _ = write!(result, "\\u{:04x}", character as u32);
            },
            character => result.push(character),
        }
    }
    result.push('"');
    result
}

fn fingerprint(bytes: &[u8]) -> String {
    // This is intentionally only a workload receipt, not a source identity.
    // The Python runner records SHA-256 for every source, fixture, lock, and
    // binary path.  Keeping this tiny fingerprint in the probe avoids adding
    // a hashing dependency to the standalone performance package.
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    format!("fnv1a64-{hash:016x}")
}

fn report(arguments: &Arguments, fixture_bytes: &[u8]) -> String {
    let allocator_metrics = if arguments.mode == "allocator" {
        "{\"status\":\"unavailable\",\"elapsed_claimable\":false,\"direct_bytes\":null,\"realloc_bytes\":null,\"live_balance\":null,\"peak_live_delta\":null,\"process_rss_bytes\":null,\"peak_rss_bytes\":null}"
    } else {
        "null"
    };
    let mut rows = String::new();
    for (index, operation) in OPERATIONS.iter().enumerate() {
        if index != 0 {
            rows.push(',');
        }
        rows.push_str("{\"operation\":");
        rows.push_str(&json_string(operation));
        rows.push_str(",\"status\":\"unavailable\",\"reason\":\"lane adapter is not supplied until the final source freeze\",\"elapsed_ns\":null,\"fidelity_checked\":false,\"typed_refusal_checked\":false,\"sink_writes\":null}");
    }
    format!(
        "{{\"schema\":\"litchi.drawingml.raw.probe.v1\",\"lane\":{},\"role\":{},\"mode\":{},\"source_revision\":{},\"status\":\"unavailable\",\"qualification_claim\":false,\"availability_reason\":\"lane adapter is not supplied until the final source freeze\",\"fixture_bytes\":{},\"fixture_fingerprint\":{},\"protocol\":{{\"warmups\":{},\"samples_per_leg\":{},\"blocks\":{},\"target_ms\":{},\"minimum_ms\":{},\"fresh_process\":true,\"abba_order\":[\"baseline\",\"candidate\",\"candidate\",\"baseline\"]}},\"allocator_metrics\":{},\"resource_metrics\":{{\"process_rss_bytes\":null,\"peak_rss_bytes\":null}},\"rows\":[{}],\"iterations\":{}}}\n",
        json_string(&arguments.lane),
        json_string(&arguments.role),
        json_string(&arguments.mode),
        json_string(&arguments.source_revision),
        fixture_bytes.len(),
        json_string(&fingerprint(fixture_bytes)),
        arguments.warmups,
        arguments.samples,
        arguments.blocks,
        arguments.target_ms,
        arguments.minimum_ms,
        allocator_metrics,
        rows,
        arguments.iterations,
    )
}

fn write_report(path: &Path, contents: &str) -> io::Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| invalid_argument("JSON output has no parent directory".into()))?;
    if !parent.is_dir() || parent.is_symlink() {
        return Err(invalid_argument(format!(
            "JSON output parent is not a regular directory: {parent:?}"
        )));
    }
    let file_name = path
        .file_name()
        .ok_or_else(|| invalid_argument("JSON output has no file name".into()))?;
    let temporary = parent.join(format!(".{}.tmp", file_name.to_string_lossy()));
    if temporary.exists() || temporary.is_symlink() {
        return Err(invalid_argument(format!("temporary JSON output already exists: {temporary:?}")));
    }
    fs::write(&temporary, contents)?;
    fs::rename(temporary, path)
}

/// Run the unavailable-until-freeze protocol endpoint.
pub fn run(allocator: bool) -> io::Result<()> {
    let arguments = parse(env::args().skip(1), allocator).map_err(|error| {
        io::Error::new(
            error.kind(),
            format!("{error}; {}", usage()),
        )
    })?;
    let fixture_bytes = fs::read(&arguments.fixture)?;
    write_report(&arguments.output, &report(&arguments, &fixture_bytes))
}
