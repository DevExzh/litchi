//! Change 0745 measurement probe: PPT artifact-digest deferral.
//!
//! One process runs one `--mode` over one `--input` fixture. Every mode builds
//! its timed owner from the public `litchi_ppt::slide_order` (or, for the DOC
//! controls, `litchi_doc::body_text`) API only. Setup that a mode does not time
//! (fresh source snapshots, durable-patch decoding, oracle digests) happens
//! outside the timed interval. Each process prints one JSON report.
//!
//! Timed regions:
//! - `remove`: open + edit + remove slide 1 + commit + output copy (the
//!   0728/0732/0734 public slide-removal lifecycle).
//! - `remove-durable`: `remove` plus `Patch::to_durable` and deterministic
//!   JSON encoding of the forward patch.
//! - `chain-durable`: two successive remove/commit/serialize edits, the second
//!   on the first commit's snapshot.
//! - `apply-durable`: `Snapshot::apply_durable` of a decoded slide-removal
//!   patch on a fresh (untimed) source snapshot, plus output copy.
//! - `noop`: edit + commit of an empty transaction on a fresh source snapshot.
//! - `hide`: edit + toggle slide 0 hidden + commit + output copy.
//! - `doc-replace`: DOC control; open + edit + replace paragraph 0 + commit +
//!   output copy (the 0728/0734 public DOC route).
//! - `doc-replace-durable`: `doc-replace` plus DOC `to_durable` + JSON.
//!
//! `goldens` runs no timing: it prints the SHA-256 and length of deterministic
//! durable-patch JSON for fixed scenarios so two builds can be compared.

pub mod alloc_metrics;

use std::error::Error;
use std::hint::black_box;
use std::path::PathBuf;
use std::time::Instant;

use litchi_core::Position;
use litchi_core::patch::{BlobId, BlobLimits, Patch, PatchLimits, Reversible};

pub type BoxError = Box<dyn Error>;

fn fail<T>(message: impl Into<String>) -> Result<T, BoxError> {
    Err(message.into().into())
}

fn sha(bytes: &[u8]) -> String {
    BlobId::of(bytes).as_hex()
}

fn patch_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(8, 8 * 1024 * 1024, 16 * 1024 * 1024),
        20 * 1024 * 1024,
        64,
        8,
        4096,
        18 * 1024 * 1024,
    )
}

const DOC_TEXT: &str = "litchi copy-through baseline replacement text";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mode {
    Remove,
    RemoveDurable,
    ChainDurable,
    ApplyDurable,
    Noop,
    Hide,
    DocReplace,
    DocReplaceDurable,
    Goldens,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, BoxError> {
        Ok(match value {
            "remove" => Self::Remove,
            "remove-durable" => Self::RemoveDurable,
            "chain-durable" => Self::ChainDurable,
            "apply-durable" => Self::ApplyDurable,
            "noop" => Self::Noop,
            "hide" => Self::Hide,
            "doc-replace" => Self::DocReplace,
            "doc-replace-durable" => Self::DocReplaceDurable,
            "goldens" => Self::Goldens,
            other => return fail(format!("unknown --mode {other:?}")),
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::Remove => "remove",
            Self::RemoveDurable => "remove-durable",
            Self::ChainDurable => "chain-durable",
            Self::ApplyDurable => "apply-durable",
            Self::Noop => "noop",
            Self::Hide => "hide",
            Self::DocReplace => "doc-replace",
            Self::DocReplaceDurable => "doc-replace-durable",
            Self::Goldens => "goldens",
        }
    }
}

struct Args {
    mode: Mode,
    inputs: Vec<PathBuf>,
    warmups: usize,
    samples: usize,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut mode = None;
    let mut inputs = Vec::new();
    let mut warmups = 3usize;
    let mut samples = 15usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut value = || {
            arguments
                .next()
                .ok_or_else(|| BoxError::from(format!("missing value for {flag}")))
        };
        match flag.as_str() {
            "--mode" => mode = Some(Mode::parse(&value()?)?),
            "--input" => inputs.push(PathBuf::from(value()?)),
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            other => return fail(format!("unknown flag {other:?}")),
        }
    }
    let mode = mode.ok_or("missing --mode")?;
    if inputs.is_empty() {
        return fail("missing --input");
    }
    if samples == 0 || samples > 10_000 || warmups > 10_000 {
        return fail("--samples must be 1..=10000 and --warmups <= 10000");
    }
    Ok(Args {
        mode,
        inputs,
        warmups,
        samples,
    })
}

/// One timed owner's observable results. Only `output` and `durable` bytes
/// escape the timed interval; everything else is dropped inside it.
struct Outcome {
    output: Vec<u8>,
    durable: Option<Vec<u8>>,
}

fn ppt_remove(source: &[u8]) -> Result<Outcome, BoxError> {
    let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
    let mut edit = snapshot.edit()?;
    edit.remove_slide(Position::new(1))?;
    let commit = edit.commit()?;
    Ok(Outcome {
        output: commit.snapshot().bytes().to_vec(),
        durable: None,
    })
}

fn ppt_remove_durable(source: &[u8]) -> Result<Outcome, BoxError> {
    let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
    let mut edit = snapshot.edit()?;
    edit.remove_slide(Position::new(1))?;
    let commit = edit.commit()?;
    let durable = commit
        .patch()
        .to_durable(patch_limits())?
        .to_deterministic_json()?;
    Ok(Outcome {
        output: commit.snapshot().bytes().to_vec(),
        durable: Some(durable),
    })
}

fn ppt_chain_durable(source: &[u8]) -> Result<Outcome, BoxError> {
    let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
    let mut first = snapshot.edit()?;
    first.remove_slide(Position::new(1))?;
    let first = first.commit()?;
    let mut wire = first
        .patch()
        .to_durable(patch_limits())?
        .to_deterministic_json()?;
    let mut second = first.snapshot().edit()?;
    second.remove_slide(Position::new(1))?;
    let second = second.commit()?;
    wire.extend_from_slice(
        &second
            .patch()
            .to_durable(patch_limits())?
            .to_deterministic_json()?,
    );
    Ok(Outcome {
        output: second.snapshot().bytes().to_vec(),
        durable: Some(wire),
    })
}

fn ppt_apply_durable(
    snapshot: &litchi_ppt::slide_order::Snapshot,
    patch: &Patch<Reversible>,
) -> Result<Outcome, BoxError> {
    let applied = snapshot.apply_durable(patch)?;
    Ok(Outcome {
        output: applied.bytes().to_vec(),
        durable: None,
    })
}

fn ppt_noop(snapshot: &litchi_ppt::slide_order::Snapshot) -> Result<Outcome, BoxError> {
    let commit = snapshot.edit()?.commit()?;
    if !commit.patch().is_empty() {
        return fail("empty PPT transaction produced a changed patch");
    }
    Ok(Outcome {
        output: commit.snapshot().bytes().to_vec(),
        durable: None,
    })
}

fn ppt_hide(snapshot: &litchi_ppt::slide_order::Snapshot) -> Result<Outcome, BoxError> {
    let hidden = snapshot.slide_hidden(Position::new(0))?;
    let mut edit = snapshot.edit()?;
    edit.set_slide_hidden(Position::new(0), !hidden)?;
    let commit = edit.commit()?;
    Ok(Outcome {
        output: commit.snapshot().bytes().to_vec(),
        durable: None,
    })
}

fn doc_replace(source: &[u8], durable: bool) -> Result<Outcome, BoxError> {
    let snapshot = litchi_doc::body_text::Snapshot::open(
        source.to_vec(),
        litchi_doc::tracked_revision::Limits::default(),
    )?;
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph(Position::new(0), DOC_TEXT)?;
    let commit = edit.commit()?;
    let wire = if durable {
        Some(
            commit
                .patch()
                .to_durable(patch_limits())?
                .to_deterministic_json()?,
        )
    } else {
        None
    };
    Ok(Outcome {
        output: commit.snapshot().bytes().to_vec(),
        durable: wire,
    })
}

/// Per-mode untimed state built once per process.
enum Prepared {
    None,
    Durable(Patch<Reversible>),
}

fn prepare(mode: Mode, source: &[u8]) -> Result<Prepared, BoxError> {
    if mode != Mode::ApplyDurable {
        return Ok(Prepared::None);
    }
    let Outcome { durable, .. } = ppt_remove_durable(source)?;
    let wire = durable.ok_or("durable preparation produced no wire bytes")?;
    Ok(Prepared::Durable(Patch::<Reversible>::from_deterministic_json(
        &wire,
        patch_limits(),
    )?))
}

/// Untimed per-sample setup for modes that time an operation on an already
/// opened snapshot. A fresh snapshot is opened for every sample so no digest
/// memo or other state survives from one sample to the next.
fn fresh_snapshot(
    mode: Mode,
    source: &[u8],
) -> Result<Option<litchi_ppt::slide_order::Snapshot>, BoxError> {
    match mode {
        Mode::ApplyDurable | Mode::Noop | Mode::Hide => Ok(Some(
            litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?,
        )),
        _ => Ok(None),
    }
}

#[inline(never)]
fn timed_owner(
    mode: Mode,
    source: &[u8],
    prepared: &Prepared,
    snapshot: Option<&litchi_ppt::slide_order::Snapshot>,
) -> Result<Outcome, BoxError> {
    match (mode, prepared, snapshot) {
        (Mode::Remove, _, _) => ppt_remove(source),
        (Mode::RemoveDurable, _, _) => ppt_remove_durable(source),
        (Mode::ChainDurable, _, _) => ppt_chain_durable(source),
        (Mode::ApplyDurable, Prepared::Durable(patch), Some(snapshot)) => {
            ppt_apply_durable(snapshot, patch)
        },
        (Mode::Noop, _, Some(snapshot)) => ppt_noop(snapshot),
        (Mode::Hide, _, Some(snapshot)) => ppt_hide(snapshot),
        (Mode::DocReplace, _, _) => doc_replace(source, false),
        (Mode::DocReplaceDurable, _, _) => doc_replace(source, true),
        _ => fail("mode setup is inconsistent"),
    }
}

fn json_string(value: &str) -> String {
    let mut text = String::with_capacity(value.len() + 2);
    text.push('"');
    for character in value.chars() {
        match character {
            '"' => text.push_str("\\\""),
            '\\' => text.push_str("\\\\"),
            '\n' => text.push_str("\\n"),
            control if control.is_control() => {
                text.push_str(&format!("\\u{:04x}", control as u32));
            },
            other => text.push(other),
        }
    }
    text.push('"');
    text
}

fn run_timed(args: &Args, timing: bool) -> Result<(), BoxError> {
    let [input] = args.inputs.as_slice() else {
        return fail("timed modes take exactly one --input");
    };
    let source = std::fs::read(input)?;
    let prepared = prepare(args.mode, &source)?;
    let mut elapsed_ns = Vec::with_capacity(args.samples);
    let mut regions = Vec::new();
    let mut output_digest: Option<(String, usize)> = None;
    let mut durable_digest: Option<(String, usize)> = None;
    for iteration in 0..args.warmups + args.samples {
        let measured = iteration >= args.warmups;
        let snapshot = fresh_snapshot(args.mode, &source)?;
        let outcome = if timing {
            let started = Instant::now();
            let outcome = timed_owner(args.mode, &source, &prepared, snapshot.as_ref())?;
            let nanos = started.elapsed().as_nanos();
            black_box(outcome.output.len());
            if measured {
                elapsed_ns.push(nanos);
            }
            outcome
        } else {
            let mut captured = None;
            let region = alloc_metrics::region(|| {
                let outcome = timed_owner(args.mode, &source, &prepared, snapshot.as_ref())?;
                captured = Some(outcome);
                Ok::<(), BoxError>(())
            })?;
            if measured {
                regions.push(region);
            }
            captured.ok_or("allocation region produced no outcome")?
        };
        drop(snapshot);
        // Oracle: every owner of one process must publish identical bytes.
        let digest = (sha(&outcome.output), outcome.output.len());
        match &output_digest {
            Some(expected) if *expected != digest => {
                return fail("output bytes changed between owners of one process");
            },
            Some(_) => {},
            None => output_digest = Some(digest),
        }
        if let Some(wire) = &outcome.durable {
            let digest = (sha(wire), wire.len());
            match &durable_digest {
                Some(expected) if *expected != digest => {
                    return fail("durable bytes changed between owners of one process");
                },
                Some(_) => {},
                None => durable_digest = Some(digest),
            }
        }
    }
    let (output_sha, output_len) = output_digest.ok_or("no owner ran")?;
    let mut report = format!(
        "{{\"schema\":\"0745-probe-v1\",\"mode\":{},\"input\":{},\"input_sha256\":{},\"input_bytes\":{},\"warmups\":{},\"samples\":{},\"output_sha256\":{},\"output_bytes\":{}",
        json_string(args.mode.name()),
        json_string(&input.display().to_string()),
        json_string(&sha(&source)),
        source.len(),
        args.warmups,
        args.samples,
        json_string(&output_sha),
        output_len,
    );
    if let Some((durable_sha, durable_len)) = durable_digest {
        report.push_str(&format!(
            ",\"durable_sha256\":{},\"durable_bytes\":{}",
            json_string(&durable_sha),
            durable_len
        ));
    }
    if timing {
        report.push_str(",\"elapsed_ns\":[");
        report.push_str(
            &elapsed_ns
                .iter()
                .map(u128::to_string)
                .collect::<Vec<_>>()
                .join(","),
        );
        report.push(']');
    } else {
        report.push_str(",\"allocation\":[");
        report.push_str(
            &regions
                .iter()
                .map(alloc_metrics::Region::json)
                .collect::<Vec<_>>()
                .join(","),
        );
        report.push(']');
    }
    report.push('}');
    println!("{report}");
    Ok(())
}

/// Deterministic durable-patch wire digests for fixed scenarios. The same
/// scenario list runs against both builds; any refused stage is reported with
/// its error text rather than skipped, so refusals are compared as well.
fn run_goldens(args: &Args) -> Result<(), BoxError> {
    use litchi_ppt::slide_order::Snapshot;

    type Scenario = fn(&Snapshot) -> Result<litchi_ppt::slide_order::Commit, BoxError>;
    let generic: [(&str, Scenario); 6] = [
        ("remove-1", |source| {
            let mut edit = source.edit()?;
            edit.remove_slide(Position::new(1))?;
            Ok(edit.commit()?)
        }),
        ("move-0-to-last", |source| {
            let mut edit = source.edit()?;
            edit.move_slide(
                Position::new(0),
                Position::new(source.slide_count().saturating_sub(1)),
            )?;
            Ok(edit.commit()?)
        }),
        ("hide-0", |source| {
            let hidden = source.slide_hidden(Position::new(0))?;
            let mut edit = source.edit()?;
            edit.set_slide_hidden(Position::new(0), !hidden)?;
            Ok(edit.commit()?)
        }),
        ("hide-0-then-remove-1", |source| {
            let hidden = source.slide_hidden(Position::new(0))?;
            let mut edit = source.edit()?;
            edit.set_slide_hidden(Position::new(0), !hidden)?;
            edit.remove_slide(Position::new(1))?;
            Ok(edit.commit()?)
        }),
        ("move-0-1-and-back", |source| {
            let mut edit = source.edit()?;
            edit.move_slide(Position::new(0), Position::new(1))?;
            edit.move_slide(Position::new(1), Position::new(0))?;
            Ok(edit.commit()?)
        }),
        ("text-0-0-then-move-0-1", |source| {
            let mut edit = source.edit()?;
            edit.set_shape_text(
                litchi_ppt::text_edit::Target::new(Position::new(0), Position::new(0)),
                "golden replacement",
            )?;
            edit.move_slide(Position::new(0), Position::new(1))?;
            Ok(edit.commit()?)
        }),
    ];
    let authored: [(&str, Scenario); 2] = [
        ("transfer-0-into-remove-1", |donor| {
            let mut remove = donor.edit()?;
            remove.remove_slide(Position::new(1))?;
            let receiver = remove.commit()?.snapshot().clone();
            let plan = receiver.plan_transfer_from(donor, Position::new(0))?;
            let mut edit = receiver.edit()?;
            edit.insert_transfer(Position::new(1), &plan)?;
            Ok(edit.commit()?)
        }),
        ("anchor-0-0-then-remove-1", |source| {
            let mut edit = source.edit()?;
            edit.set_shape_anchor(
                litchi_ppt::text_edit::Target::new(Position::new(0), Position::new(0)),
                litchi_ppt::Anchor::small(25, 35, 325, 235)?,
            )?;
            edit.remove_slide(Position::new(1))?;
            Ok(edit.commit()?)
        }),
    ];

    let mut fixtures = Vec::new();
    for input in &args.inputs {
        let name = input
            .file_name()
            .and_then(|name| name.to_str())
            .ok_or("input has no UTF-8 file name")?
            .to_string();
        fixtures.push((name, std::fs::read(input)?, &generic[..]));
    }
    fixtures.push((
        "authored-two-slides".to_string(),
        authored_fixture()?,
        &generic[..],
    ));
    let mut entries = Vec::new();
    for (name, bytes, scenarios) in &fixtures {
        let extra: &[(&str, Scenario)] = if name == "authored-two-slides" {
            &authored[..]
        } else {
            &[]
        };
        for (scenario, run) in scenarios.iter().chain(extra) {
            entries.push(golden_entry(name, scenario, bytes, *run));
        }
    }
    println!(
        "{{\"schema\":\"0745-goldens-v2\",\"entries\":[{}]}}",
        entries.join(",")
    );
    Ok(())
}

fn authored_fixture() -> Result<Vec<u8>, BoxError> {
    let mut writer = litchi_ppt::writer::Writer::new();
    let first = writer.add_slide()?;
    writer.add_textbox(first, 10, 10, 240, 40, "first slide")?;
    let second = writer.add_slide()?;
    writer.add_textbox(second, 10, 10, 240, 40, "second slide")?;
    let mut output = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn stage<T>(result: Result<T, BoxError>, render: impl FnOnce(&T) -> String) -> (Option<T>, String) {
    match result {
        Ok(value) => {
            let text = render(&value);
            (Some(value), text)
        },
        Err(error) => (
            None,
            format!("{{\"error\":{}}}", json_string(&error.to_string())),
        ),
    }
}

fn golden_entry(
    name: &str,
    scenario: &str,
    bytes: &[u8],
    run: fn(
        &litchi_ppt::slide_order::Snapshot,
    ) -> Result<litchi_ppt::slide_order::Commit, BoxError>,
) -> String {
    use litchi_ppt::slide_order::Snapshot;

    let digest = |bytes: &[u8]| {
        format!(
            "{{\"sha256\":{},\"bytes\":{}}}",
            json_string(&sha(bytes)),
            bytes.len()
        )
    };
    let (commit, commit_text) = stage(
        Snapshot::from_bytes(bytes.to_vec())
            .map_err(BoxError::from)
            .and_then(|source| run(&source)),
        |commit| digest(commit.snapshot().bytes()),
    );
    let mut fields = vec![
        format!("\"fixture\":{}", json_string(name)),
        format!("\"scenario\":{}", json_string(scenario)),
        format!("\"commit\":{commit_text}"),
    ];
    if let Some(commit) = commit {
        let (forward, forward_text) = stage(
            commit
                .patch()
                .to_durable(patch_limits())
                .map_err(BoxError::from)
                .and_then(|patch| Ok(patch.to_deterministic_json()?)),
            |wire| digest(wire),
        );
        let (_, inverse_text) = stage(
            commit
                .patch()
                .inverse()
                .to_durable(patch_limits())
                .map_err(BoxError::from)
                .and_then(|patch| Ok(patch.to_deterministic_json()?)),
            |wire| digest(wire),
        );
        fields.push(format!("\"forward\":{forward_text}"));
        fields.push(format!("\"inverse\":{inverse_text}"));
        if let Some(wire) = forward {
            let decode = || Patch::<Reversible>::from_deterministic_json(&wire, patch_limits());
            let (_, replay_text) = stage(
                decode()
                    .map_err(BoxError::from)
                    .and_then(|patch| {
                        // Replay on a fresh snapshot of the patch's exact base.
                        let base = Snapshot::from_bytes(commit.patch().before().to_vec())?;
                        Ok(base.apply_durable(&patch)?)
                    }),
                |applied| digest(applied.bytes()),
            );
            let (_, restore_text) = stage(
                decode().map_err(BoxError::from).and_then(|patch| {
                    Ok(commit.snapshot().apply_durable(&patch.inverse())?)
                }),
                |restored| digest(restored.bytes()),
            );
            fields.push(format!("\"durable_replay\":{replay_text}"));
            fields.push(format!("\"durable_restore\":{restore_text}"));
        }
    }
    format!("{{{}}}", fields.join(","))
}

/// Entry point shared by the timing and allocation binaries.
///
/// # Errors
///
/// Returns any argument, I/O, library or oracle failure.
pub fn run(timing: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    if args.mode == Mode::Goldens {
        return run_goldens(&args);
    }
    run_timed(&args, timing)
}
