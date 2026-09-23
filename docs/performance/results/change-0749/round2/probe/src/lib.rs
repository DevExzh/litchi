//! Change 0749 measurement probe: CFB Reuse-plan validation.
//!
//! One process runs one `--mode` over one `--input` fixture and prints one
//! JSON report. Every timed owner uses public APIs only. Setup a mode does not
//! time (reading the fixture, the untimed public edit that derives stream
//! replacements, building a writer model, output digests) happens outside the
//! timed interval.
//!
//! Timed regions:
//! - `ppt-remove`: PPT open + edit + remove slide 1 + commit + output copy
//!   (the 0728/0734/0745 public slide-removal lifecycle).
//! - `ppt-noop`: edit + commit of an empty transaction on a fresh (untimed)
//!   snapshot, plus output copy. Control: no CFB write happens.
//! - `doc-replace`: DOC open + edit + replace paragraph 0 + commit + output
//!   copy (the 0728/0734/0745 public DOC route).
//! - `container-reuse` / `container-rewrite`: the 0728 common-container
//!   control. `litchi_ole_common::object::Editor::open` + policy +
//!   `put_streams_shared` of the exact changed streams an untimed public edit
//!   produced + `finish`.
//! - `cfb-write-reuse` / `cfb-write-rewrite`: CFB-only. An untimed
//!   `OleWriter` holds the edited model (every stream of the untimed public
//!   edit's output) and has adopted the original fixture; the timed region is
//!   one `write_to` into a fresh in-memory cursor (plan, validation, emission).
//! - `cfb-open-read`: control. `OleFile::open` of the fixture and
//!   `open_stream` of every stream. Reader path only.
//!
//! `describe` runs no timing: it prints the stream inventory and the sector
//! layout report of one Reuse and one Rewrite `write_to`.
//!
//! `--edit` selects the edited model the container and CFB lanes republish:
//! `public` (default) is the untimed public PPT slide removal or DOC
//! paragraph replace; `same` and `grow` are the two edits of the repository's
//! `sector_layout_corpus` test, applied to the largest stream: flip its last
//! byte (same length), or append `3 * sector_size + 7` bytes of `0xA5`.
//!
//! `generate --streams N --size BYTES --sector-size 512|4096 --output PATH`
//! writes a compound file of `N` streams `S00000`.. of `BYTES` deterministic
//! bytes each, from scratch, and prints its digest.

pub mod alloc_metrics;

use std::error::Error;
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter, SectorLayoutPolicy, SectorLayoutReport};
use litchi_core::Position;
use litchi_core::patch::BlobId;
use litchi_ole_common::object::{Editor, Limits, Targets};

pub type BoxError = Box<dyn Error>;

fn fail<T>(message: impl Into<String>) -> Result<T, BoxError> {
    Err(message.into().into())
}

fn sha(bytes: &[u8]) -> String {
    BlobId::of(bytes).as_hex()
}

/// The 0728/0745 DOC replacement: 45 UTF-16 code units.
const DOC_TEXT: &str = "litchi copy-through baseline replacement text";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mode {
    PptRemove,
    PptNoop,
    DocReplace,
    ContainerReuse,
    ContainerRewrite,
    CfbWriteReuse,
    CfbWriteRewrite,
    CfbOpenRead,
    Describe,
    Generate,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, BoxError> {
        Ok(match value {
            "ppt-remove" => Self::PptRemove,
            "ppt-noop" => Self::PptNoop,
            "doc-replace" => Self::DocReplace,
            "container-reuse" => Self::ContainerReuse,
            "container-rewrite" => Self::ContainerRewrite,
            "cfb-write-reuse" => Self::CfbWriteReuse,
            "cfb-write-rewrite" => Self::CfbWriteRewrite,
            "cfb-open-read" => Self::CfbOpenRead,
            "describe" => Self::Describe,
            "generate" => Self::Generate,
            other => return fail(format!("unknown --mode {other:?}")),
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::PptRemove => "ppt-remove",
            Self::PptNoop => "ppt-noop",
            Self::DocReplace => "doc-replace",
            Self::ContainerReuse => "container-reuse",
            Self::ContainerRewrite => "container-rewrite",
            Self::CfbWriteReuse => "cfb-write-reuse",
            Self::CfbWriteRewrite => "cfb-write-rewrite",
            Self::CfbOpenRead => "cfb-open-read",
            Self::Describe => "describe",
            Self::Generate => "generate",
        }
    }

    fn policy(self) -> SectorLayoutPolicy {
        match self {
            Self::ContainerRewrite | Self::CfbWriteRewrite => SectorLayoutPolicy::Rewrite,
            _ => SectorLayoutPolicy::Reuse,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Edit {
    Public,
    Same,
    Grow,
}

impl Edit {
    fn parse(value: &str) -> Result<Self, BoxError> {
        Ok(match value {
            "public" => Self::Public,
            "same" => Self::Same,
            "grow" => Self::Grow,
            other => return fail(format!("unknown --edit {other:?}")),
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::Public => "public",
            Self::Same => "same",
            Self::Grow => "grow",
        }
    }
}

struct Args {
    mode: Mode,
    input: PathBuf,
    warmups: usize,
    samples: usize,
    /// `all` (default) digests every owner's output; `first` digests only the
    /// first owner, for profiling runs where the digest would dominate.
    oracle_all: bool,
    edit: Edit,
    streams: usize,
    size: usize,
    sector_size: usize,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut mode = None;
    let mut input = None;
    let mut warmups = 3usize;
    let mut samples = 15usize;
    let mut oracle_all = true;
    let mut edit = Edit::Public;
    let mut streams = 0usize;
    let mut size = 0usize;
    let mut sector_size = 512usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut value = || {
            arguments
                .next()
                .ok_or_else(|| BoxError::from(format!("missing value for {flag}")))
        };
        match flag.as_str() {
            "--mode" => mode = Some(Mode::parse(&value()?)?),
            "--input" => input = Some(PathBuf::from(value()?)),
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            "--edit" => edit = Edit::parse(&value()?)?,
            "--streams" => streams = value()?.parse()?,
            "--size" => size = value()?.parse()?,
            "--sector-size" => sector_size = value()?.parse()?,
            "--output" => input = Some(PathBuf::from(value()?)),
            "--oracle" => {
                oracle_all = match value()?.as_str() {
                    "all" => true,
                    "first" => false,
                    other => return fail(format!("unknown --oracle {other:?}")),
                }
            },
            other => return fail(format!("unknown flag {other:?}")),
        }
    }
    let mode = mode.ok_or("missing --mode")?;
    let input = input.ok_or("missing --input (or --output for generate)")?;
    if samples == 0 || samples > 100_000 || warmups > 100_000 {
        return fail("--samples must be 1..=100000 and --warmups <= 100000");
    }
    Ok(Args {
        mode,
        input,
        warmups,
        samples,
        oracle_all,
        edit,
        streams,
        size,
        sector_size,
    })
}

fn is_ppt(source: &[u8]) -> Result<bool, BoxError> {
    let ole = OleFile::open(Cursor::new(source))?;
    Ok(ole.exists(&["PowerPoint Document"]))
}

/// The public edit whose output every container/CFB lane republishes.
fn public_edit(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    if is_ppt(source)? {
        ppt_remove(source)
    } else {
        doc_replace(source)
    }
}

fn ppt_remove(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
    let mut edit = snapshot.edit()?;
    edit.remove_slide(Position::new(1))?;
    let commit = edit.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

fn ppt_noop(snapshot: &litchi_ppt::slide_order::Snapshot) -> Result<Vec<u8>, BoxError> {
    let commit = snapshot.edit()?.commit()?;
    if !commit.patch().is_empty() {
        return fail("empty PPT transaction produced a changed patch");
    }
    Ok(commit.snapshot().bytes().to_vec())
}

fn doc_replace(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let snapshot = litchi_doc::body_text::Snapshot::open(
        source.to_vec(),
        litchi_doc::tracked_revision::Limits::default(),
    )?;
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph(Position::new(0), DOC_TEXT)?;
    let commit = edit.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

type Streams = Vec<(Vec<String>, Vec<u8>)>;

fn read_streams(bytes: &[u8]) -> Result<Streams, BoxError> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    let mut streams = Vec::new();
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = ole.open_stream(&refs)?;
        streams.push((path, data));
    }
    Ok(streams)
}

/// Every storage path of `bytes`, parents before children.
fn read_storages(bytes: &[u8]) -> Result<Vec<Vec<String>>, BoxError> {
    let ole = OleFile::open(Cursor::new(bytes))?;
    let mut storages = Vec::new();
    let mut pending: Vec<Vec<String>> = vec![Vec::new()];
    while let Some(prefix) = pending.pop() {
        let refs: Vec<&str> = prefix.iter().map(String::as_str).collect();
        for entry in ole.list_directory_entries(&refs)? {
            if entry.entry_type == litchi_cfb::consts::STGTY_STORAGE {
                let mut path = prefix.clone();
                path.push(entry.name.clone());
                storages.push(path.clone());
                pending.push(path);
            }
        }
    }
    storages.sort_by(|left, right| left.len().cmp(&right.len()).then_with(|| left.cmp(right)));
    Ok(storages)
}

struct Replacement {
    path: Vec<String>,
    data: Arc<[u8]>,
}

/// The exact changed streams of one untimed public edit (0728's derivation).
fn derive_replacements(source: &[u8], expected: &[u8]) -> Result<Vec<Replacement>, BoxError> {
    let before = read_streams(source)?;
    let after = read_streams(expected)?;
    let before_paths: Vec<&Vec<String>> = before.iter().map(|(path, _)| path).collect();
    let after_paths: Vec<&Vec<String>> = after.iter().map(|(path, _)| path).collect();
    if before_paths != after_paths {
        return fail("public edit added or deleted streams; no container mapping");
    }
    let mut replacements = Vec::new();
    for ((path, old), (_, new)) in before.iter().zip(after.iter()) {
        if old != new {
            replacements.push(Replacement {
                path: path.clone(),
                data: Arc::from(new.clone().into_boxed_slice()),
            });
        }
    }
    if replacements.is_empty() {
        return fail("public edit changed no stream");
    }
    Ok(replacements)
}

fn container(
    source: &[u8],
    replacements: &[Replacement],
    policy: SectorLayoutPolicy,
) -> Result<Vec<u8>, BoxError> {
    let mut editor = Editor::open(source.to_vec(), Targets::default(), Limits::default())?;
    editor.set_sector_layout_policy(policy);
    editor.put_streams_shared(
        replacements
            .iter()
            .map(|replacement| (replacement.path.as_slice(), Arc::clone(&replacement.data))),
    )?;
    Ok(editor.finish()?)
}

/// The edited logical model a container or CFB lane republishes.
struct Model {
    storages: Vec<Vec<String>>,
    streams: Streams,
}

fn edited_model(source: &[u8], edit: Edit) -> Result<Model, BoxError> {
    if edit == Edit::Public {
        let expected = public_edit(source)?;
        return Ok(Model {
            storages: read_storages(&expected)?,
            streams: read_streams(&expected)?,
        });
    }
    let sector_size = OleFile::open(Cursor::new(source))?.sector_size();
    let mut streams = read_streams(source)?;
    // The largest stream as `sector_layout_corpus` picks it: `max_by_key`,
    // which keeps the last of equally large streams.
    let largest = streams
        .iter()
        .enumerate()
        .max_by_key(|(_, (_, data))| data.len())
        .map(|(index, _)| index)
        .ok_or("fixture has no stream")?;
    let data = &mut streams[largest].1;
    match edit {
        Edit::Same => {
            let last = data.len().checked_sub(1).ok_or("largest stream is empty")?;
            data[last] ^= 0xFF;
        },
        Edit::Grow => data.extend(std::iter::repeat_n(0xA5u8, sector_size * 3 + 7)),
        Edit::Public => {},
    }
    Ok(Model {
        storages: read_storages(source)?,
        streams,
    })
}

/// An `OleWriter` holding the edited model and the adopted original source,
/// in the shape the PPT finish and the common editor give it.
fn prepared_writer(
    source: &[u8],
    model: &Model,
    policy: SectorLayoutPolicy,
) -> Result<OleWriter, BoxError> {
    let sector_size = OleFile::open(Cursor::new(source))?.sector_size();
    let mut writer = OleWriter::with_sector_size(sector_size)?;
    writer.set_sector_layout_policy(policy);
    if !writer.adopt_source_layout(source)? {
        return fail("fixture was not adopted as a source layout");
    }
    for storage in &model.storages {
        let refs: Vec<&str> = storage.iter().map(String::as_str).collect();
        writer.create_storage(&refs)?;
    }
    for (path, data) in &model.streams {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        writer.create_stream(&refs, data)?;
    }
    Ok(writer)
}

fn cfb_write(writer: &mut OleWriter) -> Result<Vec<u8>, BoxError> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn cfb_open_read(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let mut ole = OleFile::open(Cursor::new(source))?;
    let mut total = 0usize;
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        total += black_box(ole.open_stream(&refs)?).len();
    }
    Ok(total.to_le_bytes().to_vec())
}

fn report_json(report: Option<SectorLayoutReport>) -> String {
    match report {
        None => "null".to_string(),
        Some(report) => format!(
            "{{\"reused\":{},\"fallback\":{},\"output_sectors\":{},\"kept_sectors\":{},\"rewritten_sectors\":{},\"appended_sectors\":{},\"reclaimed_sectors\":{},\"free_sectors\":{}}}",
            report.reused_source_layout(),
            report
                .fallback()
                .map_or_else(|| "null".to_string(), |fallback| json_string(&format!("{fallback:?}"))),
            report.output_sectors(),
            report.kept_sectors(),
            report.rewritten_sectors(),
            report.appended_sectors(),
            report.reclaimed_sectors(),
            report.free_sectors(),
        ),
    }
}

/// Per-process untimed state.
enum Prepared {
    None,
    Replacements(Vec<Replacement>),
    Writer(Box<OleWriter>),
}

fn prepare(mode: Mode, edit: Edit, source: &[u8]) -> Result<Prepared, BoxError> {
    Ok(match mode {
        Mode::ContainerReuse | Mode::ContainerRewrite => {
            if edit != Edit::Public {
                return fail("container lanes take --edit public");
            }
            let expected = public_edit(source)?;
            Prepared::Replacements(derive_replacements(source, &expected)?)
        },
        Mode::CfbWriteReuse | Mode::CfbWriteRewrite => {
            let model = edited_model(source, edit)?;
            Prepared::Writer(Box::new(prepared_writer(source, &model, mode.policy())?))
        },
        _ => Prepared::None,
    })
}

fn fresh_snapshot(
    mode: Mode,
    source: &[u8],
) -> Result<Option<litchi_ppt::slide_order::Snapshot>, BoxError> {
    match mode {
        Mode::PptNoop => Ok(Some(litchi_ppt::slide_order::Snapshot::from_bytes(
            source.to_vec(),
        )?)),
        _ => Ok(None),
    }
}

#[inline(never)]
fn timed_owner(
    mode: Mode,
    source: &[u8],
    prepared: &mut Prepared,
    snapshot: Option<&litchi_ppt::slide_order::Snapshot>,
) -> Result<Vec<u8>, BoxError> {
    match (mode, prepared, snapshot) {
        (Mode::PptRemove, _, _) => ppt_remove(source),
        (Mode::PptNoop, _, Some(snapshot)) => ppt_noop(snapshot),
        (Mode::DocReplace, _, _) => doc_replace(source),
        (Mode::ContainerReuse | Mode::ContainerRewrite, Prepared::Replacements(list), _) => {
            container(source, list, mode.policy())
        },
        (Mode::CfbWriteReuse | Mode::CfbWriteRewrite, Prepared::Writer(writer), _) => {
            cfb_write(writer)
        },
        (Mode::CfbOpenRead, _, _) => cfb_open_read(source),
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
    let source = std::fs::read(&args.input)?;
    let mut prepared = prepare(args.mode, args.edit, &source)?;
    let mut elapsed_ns = Vec::with_capacity(args.samples);
    let mut regions = Vec::new();
    let mut output_digest: Option<(String, usize)> = None;
    for iteration in 0..args.warmups + args.samples {
        let measured = iteration >= args.warmups;
        let snapshot = fresh_snapshot(args.mode, &source)?;
        let output = if timing {
            let started = Instant::now();
            let output = timed_owner(args.mode, &source, &mut prepared, snapshot.as_ref())?;
            let nanos = started.elapsed().as_nanos();
            black_box(output.len());
            if measured {
                elapsed_ns.push(nanos);
            }
            output
        } else {
            let mut captured = None;
            let region = alloc_metrics::region(|| {
                let output = timed_owner(args.mode, &source, &mut prepared, snapshot.as_ref())?;
                captured = Some(output);
                Ok::<(), BoxError>(())
            })?;
            if measured {
                regions.push(region);
            }
            captured.ok_or("allocation region produced no outcome")?
        };
        drop(snapshot);
        if !args.oracle_all && output_digest.is_some() {
            continue;
        }
        // Oracle: every owner of one process must publish identical bytes.
        let digest = (sha(&output), output.len());
        match &output_digest {
            Some(expected) if *expected != digest => {
                return fail("output bytes changed between owners of one process");
            },
            Some(_) => {},
            None => output_digest = Some(digest),
        }
    }
    let layout = match &prepared {
        Prepared::Writer(writer) => report_json(writer.last_sector_layout()),
        _ => "null".to_string(),
    };
    let (output_sha, output_len) = output_digest.ok_or("no owner ran")?;
    let mut report = format!(
        "{{\"schema\":\"0749-probe-v2\",\"edit\":{},\"oracle\":{},\"mode\":{},\"input\":{},\"input_sha256\":{},\"input_bytes\":{},\"warmups\":{},\"samples\":{},\"output_sha256\":{},\"output_bytes\":{},\"sector_layout\":{}",
        json_string(args.edit.name()),
        json_string(if args.oracle_all { "all" } else { "first" }),
        json_string(args.mode.name()),
        json_string(&args.input.display().to_string()),
        json_string(&sha(&source)),
        source.len(),
        args.warmups,
        args.samples,
        json_string(&output_sha),
        output_len,
        layout,
    );
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

fn run_describe(args: &Args) -> Result<(), BoxError> {
    let source = std::fs::read(&args.input)?;
    let model = edited_model(&source, args.edit)?;
    let before = read_streams(&source)?;
    let mut rows = Vec::new();
    let mut mini_streams = 0usize;
    for ((path, old), (_, new)) in before.iter().zip(model.streams.iter()) {
        mini_streams += usize::from(old.len() < 4096);
        rows.push(format!(
            "{{\"path\":{},\"source_bytes\":{},\"edited_bytes\":{},\"changed\":{}}}",
            json_string(&path.join("/")),
            old.len(),
            new.len(),
            old != new
        ));
    }
    let mut layouts = Vec::new();
    for policy in [SectorLayoutPolicy::Reuse, SectorLayoutPolicy::Rewrite] {
        let mut writer = prepared_writer(&source, &model, policy)?;
        let output = cfb_write(&mut writer)?;
        let reread = read_streams(&output)?;
        if reread != model.streams {
            return fail("CFB write output streams differ from the edited model");
        }
        layouts.push(format!(
            "{{\"policy\":{},\"output_sha256\":{},\"output_bytes\":{},\"report\":{}}}",
            json_string(&format!("{policy:?}")),
            json_string(&sha(&output)),
            output.len(),
            report_json(writer.last_sector_layout())
        ));
    }
    println!(
        "{{\"schema\":\"0749-describe-v2\",\"input\":{},\"input_sha256\":{},\"edit\":{},\"storages\":{},\"stream_count\":{},\"source_mini_streams\":{},\"streams\":[{}],\"cfb_write\":[{}]}}",
        json_string(&args.input.display().to_string()),
        json_string(&sha(&source)),
        json_string(args.edit.name()),
        model.storages.len(),
        before.len(),
        mini_streams,
        if before.len() <= 40 { rows.join(",") } else { String::new() },
        layouts.join(",")
    );
    Ok(())
}

/// Writes `streams` deterministic streams of `size` bytes each, from scratch.
fn run_generate(args: &Args) -> Result<(), BoxError> {
    if args.streams == 0 || !matches!(args.sector_size, 512 | 4096) {
        return fail("generate needs --streams > 0 and --sector-size 512 or 4096");
    }
    let mut writer = OleWriter::with_sector_size(args.sector_size)?;
    for index in 0..args.streams {
        let payload: Vec<u8> = (0..args.size)
            .map(|byte| u8::try_from((byte * 7 + index * 13) % 251).unwrap_or(0))
            .collect();
        let name = format!("S{index:05}");
        writer.create_stream_owned(&[name.as_str()], payload)?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let bytes = output.into_inner();
    std::fs::write(&args.input, &bytes)?;
    println!(
        "{{\"schema\":\"0749-generate-v1\",\"output\":{},\"streams\":{},\"size\":{},\"sector_size\":{},\"sha256\":{},\"bytes\":{}}}",
        json_string(&args.input.display().to_string()),
        args.streams,
        args.size,
        args.sector_size,
        json_string(&sha(&bytes)),
        bytes.len()
    );
    Ok(())
}

/// Entry point shared by the timing and allocation binaries.
///
/// # Errors
///
/// Returns any argument, I/O, library or oracle failure.
pub fn run(timing: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    if args.mode == Mode::Describe {
        return run_describe(&args);
    }
    if args.mode == Mode::Generate {
        return run_generate(&args);
    }
    run_timed(&args, timing)
}
