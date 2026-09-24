//! Change 0767 measurement probe: CFB structural parsing and reads.
//!
//! One process runs one `--mode` over one `--input` file and prints one JSON
//! report. Every timed owner uses public APIs only. Setup a mode does not
//! time (reading the file, building a writer model, a fresh untimed reader for
//! the read-only lane, output digests) happens outside the timed interval.
//!
//! Timed regions:
//! - `cfb-open`: `OleFile::open` of the input (header, FAT, directory,
//!   MiniFAT, stream-allocation and partition validation).
//! - `shared-open`: `SharedOleFile::open_owned` of the input, which runs the
//!   same validating parse and keeps its index.
//! - `cfb-open-read`: `OleFile::open`, `list_streams` and `open_stream` of
//!   every stream (0749's reader control).
//! - `cfb-read-all`: `list_streams` and `open_stream` of every stream on a
//!   reader opened, untimed, for this owner.
//! - `editor-open`: `litchi_ole_common::object::Editor::open` with no
//!   targets, which parses the file and captures every stream.
//! - `cfb-write-reuse` / `cfb-write-rewrite`: one `OleWriter::write_to` of
//!   the edited model over the adopted original (0749's CFB-only lane). The
//!   Rewrite policy never reparses, so it is a control here.
//!
//! `--edit` selects the edited model the write lanes republish: `public`
//! (default) is the untimed public PPT slide removal or DOC paragraph
//! replace; `same` flips the largest stream's last byte; `grow` appends
//! `3 * sector_size + 7` bytes of `0xA5` to it.
//!
//! `verdicts --input PATH --cases N --seed S` applies `N` seeded sets of
//! byte-level faults to the input and prints every open and read verdict
//! (see `verdicts.rs`); the before and after binaries' outputs are compared
//! byte for byte.
//!
//! `generate --streams N --size BYTES --sector-size 512|4096 --output PATH`
//! writes a compound file of `N` root streams `S00000`.. of `BYTES`
//! deterministic bytes each, from scratch, and prints its digest. This is
//! 0749's generator; `BYTES` below 4096 makes every stream a mini stream.

pub mod alloc_metrics;
mod verdicts;

use std::error::Error;
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter, SectorLayoutPolicy, SectorLayoutReport, SharedOleFile};
use litchi_core::patch::BlobId;
use litchi_core::{Position, SourceVersion};
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
    CfbOpen,
    SharedOpen,
    CfbOpenRead,
    CfbReadAll,
    EditorOpen,
    CfbWriteReuse,
    CfbWriteRewrite,
    Generate,
    Verdicts,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, BoxError> {
        Ok(match value {
            "cfb-open" => Self::CfbOpen,
            "shared-open" => Self::SharedOpen,
            "cfb-open-read" => Self::CfbOpenRead,
            "cfb-read-all" => Self::CfbReadAll,
            "editor-open" => Self::EditorOpen,
            "cfb-write-reuse" => Self::CfbWriteReuse,
            "cfb-write-rewrite" => Self::CfbWriteRewrite,
            "generate" => Self::Generate,
            "verdicts" => Self::Verdicts,
            other => return fail(format!("unknown --mode {other:?}")),
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::CfbOpen => "cfb-open",
            Self::SharedOpen => "shared-open",
            Self::CfbOpenRead => "cfb-open-read",
            Self::CfbReadAll => "cfb-read-all",
            Self::EditorOpen => "editor-open",
            Self::CfbWriteReuse => "cfb-write-reuse",
            Self::CfbWriteRewrite => "cfb-write-rewrite",
            Self::Generate => "generate",
            Self::Verdicts => "verdicts",
        }
    }

    fn policy(self) -> SectorLayoutPolicy {
        match self {
            Self::CfbWriteRewrite => SectorLayoutPolicy::Rewrite,
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
    /// first owner, for counter runs where the digest would dominate.
    oracle_all: bool,
    edit: Edit,
    streams: usize,
    size: usize,
    sector_size: usize,
    cases: usize,
    seed: u64,
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
    let mut cases = 100usize;
    let mut seed = 767u64;
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut value = || {
            arguments
                .next()
                .ok_or_else(|| BoxError::from(format!("missing value for {flag}")))
        };
        match flag.as_str() {
            "--mode" => mode = Some(Mode::parse(&value()?)?),
            "--input" | "--output" => input = Some(PathBuf::from(value()?)),
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            "--edit" => edit = Edit::parse(&value()?)?,
            "--streams" => streams = value()?.parse()?,
            "--size" => size = value()?.parse()?,
            "--sector-size" => sector_size = value()?.parse()?,
            "--cases" => cases = value()?.parse()?,
            "--seed" => seed = value()?.parse()?,
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
    if samples == 0 || samples > 1_000_000 || warmups > 1_000_000 {
        return fail("--samples must be 1..=1000000 and --warmups <= 1000000");
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
        cases,
        seed,
    })
}

fn is_ppt(source: &[u8]) -> Result<bool, BoxError> {
    let ole = OleFile::open(Cursor::new(source))?;
    Ok(ole.exists(&["PowerPoint Document"]))
}

/// The public edit whose output the `public` write lanes republish.
fn public_edit(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    if is_ppt(source)? {
        let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
        let mut edit = snapshot.edit()?;
        edit.remove_slide(Position::new(1))?;
        let commit = edit.commit()?;
        Ok(commit.snapshot().bytes().to_vec())
    } else {
        let snapshot = litchi_doc::body_text::Snapshot::open(
            source.to_vec(),
            litchi_doc::tracked_revision::Limits::default(),
        )?;
        let mut edit = snapshot.edit()?;
        edit.replace_paragraph(Position::new(0), DOC_TEXT)?;
        let commit = edit.commit()?;
        Ok(commit.snapshot().bytes().to_vec())
    }
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

/// The edited logical model a write lane republishes.
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
    let largest = streams
        .iter()
        .enumerate()
        .max_by_key(|(_, (_, data))| data.len())
        .map(|(index, _)| index)
        .ok_or("file has no stream")?;
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

/// An `OleWriter` holding the edited model and the adopted original source.
fn prepared_writer(
    source: &[u8],
    model: &Model,
    policy: SectorLayoutPolicy,
) -> Result<OleWriter, BoxError> {
    let sector_size = OleFile::open(Cursor::new(source))?.sector_size();
    let mut writer = OleWriter::with_sector_size(sector_size)?;
    writer.set_sector_layout_policy(policy);
    if !writer.adopt_source_layout(source)? {
        return fail("input was not adopted as a source layout");
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

/// A digest-sized summary of an open: the file size, sector size, stream
/// count and the root entry's size.
fn open_summary<R: std::io::Read + std::io::Seek>(ole: &OleFile<R>) -> Vec<u8> {
    let mut out = Vec::with_capacity(24);
    out.extend_from_slice(&ole.file_size().to_le_bytes());
    out.extend_from_slice(&(ole.sector_size() as u64).to_le_bytes());
    out.extend_from_slice(&(ole.directory_entry_count() as u64).to_le_bytes());
    out
}

fn cfb_open(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let ole = OleFile::open(Cursor::new(source))?;
    Ok(open_summary(black_box(&ole)))
}

fn shared_open(source: &Arc<[u8]>) -> Result<Vec<u8>, BoxError> {
    let shared = SharedOleFile::open_owned(Arc::clone(source), SourceVersion::new(767, 0))?;
    let shared = black_box(shared);
    let mut out = Vec::with_capacity(16);
    out.extend_from_slice(&shared.file_size().to_le_bytes());
    out.extend_from_slice(&(shared.directory_entries().count() as u64).to_le_bytes());
    Ok(out)
}

/// Reads every stream; the result is the total length followed by the
/// concatenated per-stream lengths (a digest of it catches any difference).
fn read_all<R: std::io::Read + std::io::Seek>(ole: &mut OleFile<R>) -> Result<Vec<u8>, BoxError> {
    let mut total = 0usize;
    let mut checksum = 0u64;
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = black_box(ole.open_stream(&refs)?);
        total += data.len();
        // A cheap order-dependent fold of the first and last byte, so the
        // oracle sees the bytes, not only the lengths.
        let first = u64::from(data.first().copied().unwrap_or(0));
        let last = u64::from(data.last().copied().unwrap_or(0));
        checksum = checksum.rotate_left(5) ^ (first << 8 | last) ^ data.len() as u64;
    }
    let mut out = total.to_le_bytes().to_vec();
    out.extend_from_slice(&checksum.to_le_bytes());
    Ok(out)
}

fn cfb_open_read(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let mut ole = OleFile::open(Cursor::new(source))?;
    read_all(&mut ole)
}

fn editor_open(source: &[u8]) -> Result<Vec<u8>, BoxError> {
    let editor = Editor::open(source.to_vec(), Targets::default(), Limits::default())?;
    let editor = black_box(editor);
    drop(editor);
    Ok(source.len().to_le_bytes().to_vec())
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
    Shared(Arc<[u8]>),
    Writer(Box<OleWriter>),
}

fn prepare(mode: Mode, edit: Edit, source: &[u8]) -> Result<Prepared, BoxError> {
    Ok(match mode {
        Mode::CfbWriteReuse | Mode::CfbWriteRewrite => {
            let model = edited_model(source, edit)?;
            Prepared::Writer(Box::new(prepared_writer(source, &model, mode.policy())?))
        },
        Mode::SharedOpen => Prepared::Shared(Arc::from(source.to_vec().into_boxed_slice())),
        _ => Prepared::None,
    })
}

/// Per-owner untimed state: a fresh reader for the read-only lane.
fn fresh_reader(mode: Mode, source: &[u8]) -> Result<Option<OleFile<Cursor<&[u8]>>>, BoxError> {
    match mode {
        Mode::CfbReadAll => Ok(Some(OleFile::open(Cursor::new(source))?)),
        _ => Ok(None),
    }
}

#[inline(never)]
fn timed_owner(
    mode: Mode,
    source: &[u8],
    prepared: &mut Prepared,
    reader: Option<&mut OleFile<Cursor<&[u8]>>>,
) -> Result<Vec<u8>, BoxError> {
    match (mode, prepared, reader) {
        (Mode::CfbOpen, _, _) => cfb_open(source),
        (Mode::SharedOpen, Prepared::Shared(bytes), _) => shared_open(bytes),
        (Mode::CfbOpenRead, _, _) => cfb_open_read(source),
        (Mode::CfbReadAll, _, Some(reader)) => read_all(reader),
        (Mode::EditorOpen, _, _) => editor_open(source),
        (Mode::CfbWriteReuse | Mode::CfbWriteRewrite, Prepared::Writer(writer), _) => {
            cfb_write(writer)
        },
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
        let mut reader = fresh_reader(args.mode, &source)?;
        let output = if timing {
            let started = Instant::now();
            let output = timed_owner(args.mode, &source, &mut prepared, reader.as_mut())?;
            let nanos = started.elapsed().as_nanos();
            black_box(output.len());
            if measured {
                elapsed_ns.push(nanos);
            }
            output
        } else {
            let mut captured = None;
            let region = alloc_metrics::region(|| {
                let output = timed_owner(args.mode, &source, &mut prepared, reader.as_mut())?;
                captured = Some(output);
                Ok::<(), BoxError>(())
            })?;
            if measured {
                regions.push(region);
            }
            captured.ok_or("allocation region produced no outcome")?
        };
        drop(reader);
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
        "{{\"schema\":\"0767-probe-v1\",\"edit\":{},\"oracle\":{},\"mode\":{},\"input\":{},\"input_sha256\":{},\"input_bytes\":{},\"warmups\":{},\"samples\":{},\"output_sha256\":{},\"output_bytes\":{},\"sector_layout\":{}",
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
        "{{\"schema\":\"0767-generate-v1\",\"output\":{},\"streams\":{},\"size\":{},\"sector_size\":{},\"sha256\":{},\"bytes\":{}}}",
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
    if args.mode == Mode::Generate {
        return run_generate(&args);
    }
    if args.mode == Mode::Verdicts {
        return verdicts::run(&args.input, args.cases, args.seed);
    }
    run_timed(&args, timing)
}
