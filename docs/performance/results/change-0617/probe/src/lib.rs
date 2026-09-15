//! Phase attribution for a length-changing OLE2 edit-and-save.
//!
//! Change 0587 ranked CFB-2 ("a copy-through writer for length-changing OLE2
//! saves") with size `unknown` and wrote its own falsification criterion: the
//! item is falsified if a phase attribution of one XLS/DOC/PPT save shows that
//! the crates' own record re-encoding, not the container copy, dominates. No
//! registered selector performs such a save, so this probe is the baseline.
//!
//! Three operations per format, so that callgrind isolation pairs can difference
//! them:
//!
//! * `open`     — construct the editable snapshot and nothing else.
//! * `commit`   — `open` plus one length-changing edit plus the save. This is
//!                the whole operation CFB-2 is about.
//! * `container`— the container rebuild alone: open the generic OLE2 object
//!                editor over the same source bytes and replace exactly the
//!                streams the format crate replaced, with the exact bytes the
//!                format crate produced (captured once, outside the interval).
//!                No format record is decoded or re-encoded in this leg.
//!
//! `commit - open` is the edit plus the container rebuild; `container` is the
//! container rebuild alone, on the same source and the same replacement bytes.
//!
//! `--report` prints one JSON line of exact, deterministic counts instead of
//! running an iteration loop.

use std::hint::black_box;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;

pub mod alloc_metrics;

pub type BoxError = Box<dyn std::error::Error>;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Format {
    Xls,
    Doc,
    Ppt,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Operation {
    Open,
    Commit,
    Container,
    ContainerChangedOnly,
}

struct Args {
    input: PathBuf,
    format: Format,
    operation: Operation,
    warmups: usize,
    samples: usize,
    report: bool,
    emit_ppt_fixture: Option<PathBuf>,
    slides: usize,
    shapes: usize,
    text: String,
    sheet: usize,
    row: u32,
    column: u32,
    slide: usize,
    shape: usize,
    paragraph: usize,
    ppt_op: PptOp,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum PptOp {
    Text,
    RemoveSlide,
    MoveSlide,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut input: Option<PathBuf> = None;
    let mut format: Option<Format> = None;
    let mut operation = Operation::Commit;
    let mut warmups = 1usize;
    let mut samples = 1usize;
    let mut report = false;
    let mut emit_ppt_fixture: Option<PathBuf> = None;
    let mut slides = 3usize;
    let mut shapes = 2usize;
    let mut text = String::from("litchi copy-through baseline replacement text");
    let mut sheet = 0usize;
    let mut row = 0u32;
    let mut column = 0u32;
    let mut slide = 0usize;
    let mut shape = 0usize;
    let mut paragraph = 0usize;
    let mut ppt_op = PptOp::Text;

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
                    "xls" => Format::Xls,
                    "doc" => Format::Doc,
                    "ppt" => Format::Ppt,
                    other => return Err(format!("unknown --format {other}").into()),
                });
            },
            "--operation" => {
                operation = match value()?.as_str() {
                    "open" => Operation::Open,
                    "commit" => Operation::Commit,
                    "container" => Operation::Container,
                    "container-changed-only" => Operation::ContainerChangedOnly,
                    other => return Err(format!("unknown --operation {other}").into()),
                };
            },
            "--warmups" => warmups = value()?.parse()?,
            "--samples" => samples = value()?.parse()?,
            "--report" => report = true,
            "--emit-ppt-fixture" => emit_ppt_fixture = Some(PathBuf::from(value()?)),
            "--slides" => slides = value()?.parse()?,
            "--shapes" => shapes = value()?.parse()?,
            "--text" => text = value()?,
            "--sheet" => sheet = value()?.parse()?,
            "--row" => row = value()?.parse()?,
            "--column" => column = value()?.parse()?,
            "--slide" => slide = value()?.parse()?,
            "--shape" => shape = value()?.parse()?,
            "--paragraph" => paragraph = value()?.parse()?,
            "--ppt-op" => {
                ppt_op = match value()?.as_str() {
                    "text" => PptOp::Text,
                    "remove-slide" => PptOp::RemoveSlide,
                    "move-slide" => PptOp::MoveSlide,
                    other => return Err(format!("unknown --ppt-op {other}").into()),
                };
            },
            other => return Err(format!("unknown flag {other}").into()),
        }
    }

    Ok(Args {
        input: input.ok_or("missing --input")?,
        format: format.ok_or("missing --format")?,
        operation,
        warmups,
        samples,
        report,
        emit_ppt_fixture,
        slides,
        shapes,
        text,
        sheet,
        row,
        column,
        slide,
        shape,
        paragraph,
        ppt_op,
    })
}

// ---------------------------------------------------------------------------
// Container inventory, read through the ordinary public CFB parser.
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
struct Inventory {
    file_bytes: u64,
    sector_size: usize,
    streams: Vec<(Vec<String>, usize)>,
}

impl Inventory {
    fn of(bytes: &[u8]) -> Result<Self, BoxError> {
        let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes.to_vec()))?;
        let sector_size = ole.sector_size();
        let paths = ole.list_streams();
        let mut streams = Vec::new();
        for path in paths {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&refs)?;
            streams.push((path, data.len()));
        }
        streams.sort();
        Ok(Self {
            file_bytes: bytes.len() as u64,
            sector_size,
            streams,
        })
    }

    fn stream_bytes(&self) -> u64 {
        self.streams.iter().map(|(_, len)| *len as u64).sum()
    }
}

/// Minimal RFC 8259 string escaping: CFB stream names carry control characters
/// such as U+0001 and U+0005, which `escape_default` renders as `\u{1}`.
fn json_escape(value: &str) -> String {
    let mut out = String::new();
    for ch in value.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

fn json_streams(inventory: &Inventory) -> String {
    let mut parts = Vec::new();
    for (path, len) in &inventory.streams {
        parts.push(format!(
            "{{\"path\":\"{}\",\"bytes\":{len}}}",
            json_escape(&path.join("/"))
        ));
    }
    format!("[{}]", parts.join(","))
}

/// Streams present in both inventories whose bytes differ, plus added/removed.
fn diff(_before: &Inventory, after: &Inventory, before_bytes: &[u8], after_bytes: &[u8])
-> Result<(Vec<(String, usize, usize)>, u64), BoxError> {
    let mut before_ole = litchi_cfb::OleFile::open(Cursor::new(before_bytes.to_vec()))?;
    let mut after_ole = litchi_cfb::OleFile::open(Cursor::new(after_bytes.to_vec()))?;
    let mut changed = Vec::new();
    let mut unchanged_bytes = 0u64;
    for (path, len) in &after.streams {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let new = after_ole.open_stream(&refs)?;
        match before_ole.open_stream(&refs) {
            Ok(old) => {
                if old == new {
                    unchanged_bytes += old.len() as u64;
                } else {
                    changed.push((path.join("/"), old.len(), *len));
                }
            },
            Err(_) => changed.push((path.join("/"), 0, *len)),
        }
    }
    Ok((changed, unchanged_bytes))
}

// ---------------------------------------------------------------------------
// XLS: one string cell value set to a string that is not yet in the SST, which
// appends to the SST tail and therefore lengthens the Workbook stream.
// ---------------------------------------------------------------------------

fn xls_commit(bytes: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    use litchi_xls::cell_values::{Reference, Selector, Snapshot, Value};
    let snapshot = Snapshot::from_bytes(bytes.to_vec())?;
    let mut transaction = snapshot.edit();
    transaction.set_value(
        Selector::Position(args.sheet),
        Reference::new(args.row, args.column)?,
        Value::Text(args.text.clone()),
    )?;
    let commit = transaction.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

fn xls_open(bytes: &[u8]) -> Result<usize, BoxError> {
    let snapshot = litchi_xls::cell_values::Snapshot::from_bytes(bytes.to_vec())?;
    Ok(snapshot.bytes().len())
}

// ---------------------------------------------------------------------------
// DOC: one main-story paragraph replaced by text of a different length.
// ---------------------------------------------------------------------------

fn doc_limits() -> litchi_doc::tracked_revision::Limits {
    litchi_doc::tracked_revision::Limits::default()
}

fn doc_commit(bytes: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    use litchi_core::Position;
    use litchi_doc::body_text::Snapshot;
    let snapshot = Snapshot::open(bytes.to_vec(), doc_limits())?;
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph(Position::new(args.paragraph), &args.text)?;
    let commit = edit.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

fn doc_open(bytes: &[u8]) -> Result<usize, BoxError> {
    let snapshot = litchi_doc::body_text::Snapshot::open(bytes.to_vec(), doc_limits())?;
    Ok(snapshot.bytes().len())
}

// ---------------------------------------------------------------------------
// PPT: one shape's text replaced by text of a different length.
// ---------------------------------------------------------------------------

fn ppt_commit(bytes: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    match args.ppt_op {
        PptOp::Text => ppt_text_commit(bytes, args),
        PptOp::RemoveSlide => ppt_structural_commit(bytes, args, false),
        PptOp::MoveSlide => ppt_structural_commit(bytes, args, true),
    }
}

/// A structural PPT edit: remove or move one slide. Both rewrite the
/// `PowerPoint Document` stream through an append-only user edit, so both are
/// length-changing OLE2 saves.
fn ppt_structural_commit(bytes: &[u8], args: &Args, move_slide: bool) -> Result<Vec<u8>, BoxError> {
    use litchi_core::Position;
    use litchi_ppt::slide_order::Snapshot;
    let snapshot = Snapshot::from_bytes(bytes.to_vec())?;
    let mut transaction = snapshot.edit()?;
    if move_slide {
        transaction.move_slide(Position::new(args.slide), Position::new(args.shape))?;
    } else {
        transaction.remove_slide(Position::new(args.slide))?;
    }
    let commit = transaction.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

fn ppt_text_commit(bytes: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    use litchi_core::Position;
    use litchi_ppt::text_edit::{Snapshot, Target};
    let snapshot = Snapshot::from_bytes(bytes.to_vec())?;
    let mut transaction =
        snapshot.edit_text(Target::new(Position::new(args.slide), Position::new(args.shape)))?;
    transaction.set_text(args.text.clone())?;
    let commit = transaction.commit()?;
    Ok(commit.snapshot().bytes().to_vec())
}

fn ppt_open(bytes: &[u8], op: PptOp) -> Result<usize, BoxError> {
    if op == PptOp::Text {
        let snapshot = litchi_ppt::text_edit::Snapshot::from_bytes(bytes.to_vec())?;
        Ok(snapshot.bytes().len())
    } else {
        let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(bytes.to_vec())?;
        Ok(snapshot.bytes().len())
    }
}

// ---------------------------------------------------------------------------
// The container rebuild alone.
// ---------------------------------------------------------------------------

/// A CFB holding only the streams a commit changed, built with the ordinary
/// `OleWriter`. The container rebuild over this source performs the same
/// header/FAT/DIFAT/MiniFAT/directory work and the same changed-stream copies
/// as the rebuild over the real source, and none of the work proportional to
/// the untouched streams. The difference between the two is therefore the
/// ceiling of what a copy-through writer could remove from the container leg.
fn changed_only_source(
    bytes: &[u8],
    replacements: &[(Vec<String>, Arc<[u8]>)],
) -> Result<Vec<u8>, BoxError> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes.to_vec()))?;
    let mut writer = litchi_cfb::OleWriter::new();
    if let Some(root) = ole.root_entry() {
        writer.set_root_clsid(clsid_bytes(root));
    }
    for (path, data) in replacements {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        for depth in 1..refs.len() {
            let parent = &refs[..depth];
            let _ = writer.create_storage(parent);
        }
        writer.create_stream(&refs, data)?;
    }
    let mut out = Vec::new();
    writer.write_to(&mut Cursor::new(&mut out))?;
    Ok(out)
}

fn clsid_bytes(entry: &litchi_cfb::DirectoryEntry) -> [u8; 16] {
    let mut bytes = [0u8; 16];
    let text: String = entry.clsid.chars().filter(|c| c.is_ascii_hexdigit()).collect();
    if text.len() == 32 {
        for (index, slot) in bytes.iter_mut().enumerate() {
            *slot = u8::from_str_radix(&text[index * 2..index * 2 + 2], 16).unwrap_or(0);
        }
    }
    bytes
}

/// Replaces `replacements` in the source through the generic OLE2 object
/// editor: `Editor::open` (parse + capture every stream) then one
/// failure-atomic `put_streams_shared` (check + `OleWriter` render + reopen +
/// recapture). This is exactly the container work a format commit performs and
/// nothing else.
fn container_rebuild(
    bytes: &[u8],
    replacements: &[(Vec<String>, Arc<[u8]>)],
) -> Result<usize, BoxError> {
    use litchi_ole_common::object::{Editor, Limits, Targets};
    let mut editor = Editor::open(bytes.to_vec(), Targets::default(), Limits::default())?;
    editor.put_streams_shared(
        replacements
            .iter()
            .map(|(path, data)| (path.as_slice(), Arc::clone(data))),
    )?;
    let out = editor.finish()?;
    Ok(out.len())
}

fn changed_replacements(
    before_bytes: &[u8],
    after_bytes: &[u8],
) -> Result<Vec<(Vec<String>, Arc<[u8]>)>, BoxError> {
    let after = Inventory::of(after_bytes)?;
    let mut before_ole = litchi_cfb::OleFile::open(Cursor::new(before_bytes.to_vec()))?;
    let mut after_ole = litchi_cfb::OleFile::open(Cursor::new(after_bytes.to_vec()))?;
    let mut replacements = Vec::new();
    for (path, _) in &after.streams {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let new = after_ole.open_stream(&refs)?;
        let changed = match before_ole.open_stream(&refs) {
            Ok(old) => old != new,
            Err(_) => true,
        };
        if changed {
            replacements.push((path.clone(), Arc::from(new.into_boxed_slice())));
        }
    }
    Ok(replacements)
}

// ---------------------------------------------------------------------------

fn commit_once(bytes: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    match args.format {
        Format::Xls => xls_commit(bytes, args),
        Format::Doc => doc_commit(bytes, args),
        Format::Ppt => ppt_commit(bytes, args),
    }
}

fn open_once(bytes: &[u8], args: &Args) -> Result<usize, BoxError> {
    match args.format {
        Format::Xls => xls_open(bytes),
        Format::Doc => doc_open(bytes),
        Format::Ppt => ppt_open(bytes, args.ppt_op),
    }
}

fn report(bytes: &[u8], args: &Args) -> Result<(), BoxError> {
    let before = Inventory::of(bytes)?;
    let output = commit_once(bytes, args)?;
    let after = Inventory::of(&output)?;
    let (changed, unchanged_bytes) = diff(&before, &after, bytes, &output)?;
    let replacements = changed_replacements(bytes, &output)?;
    let container_output = container_rebuild(bytes, &replacements)?;

    // Allocation regions. Each is measured on a fresh operation so the peak is
    // that operation's own high-water mark. The counters are no-ops unless the
    // `peak` binary installed the counting allocator.
    let commit_region = alloc_metrics::region(|| commit_once(bytes, args).map(|value| value.len()))?;
    let open_region = alloc_metrics::region(|| open_once(bytes, args))?;
    let container_region =
        alloc_metrics::region(|| container_rebuild(bytes, &replacements))?;
    let changed_only = changed_only_source(bytes, &replacements)?;
    let changed_only_region =
        alloc_metrics::region(|| container_rebuild(&changed_only, &replacements))?;

    let changed_json = changed
        .iter()
        .map(|(path, old, new)| {
            format!(
                "{{\"path\":\"{}\",\"before_bytes\":{old},\"after_bytes\":{new}}}",
                json_escape(path)
            )
        })
        .collect::<Vec<_>>()
        .join(",");

    println!(
        "{{\"format\":\"{:?}\",\"fixture\":\"{}\",\
\"source_bytes\":{},\"output_bytes\":{},\
\"source_sector_size\":{},\"output_sector_size\":{},\
\"source_stream_count\":{},\"output_stream_count\":{},\
\"source_stream_bytes\":{},\"output_stream_bytes\":{},\
\"unchanged_stream_bytes\":{},\"changed_streams\":[{}],\
\"container_only_output_bytes\":{},\
\"changed_only_source_bytes\":{},\
\"alloc\":{{\"instrumented\":{},\"open\":{},\"commit\":{},\"container\":{},\"container_changed_only\":{}}},\
\"source_streams\":{},\"output_streams\":{}}}",
        args.format,
        args.input.display(),
        before.file_bytes,
        after.file_bytes,
        before.sector_size,
        after.sector_size,
        before.streams.len(),
        after.streams.len(),
        before.stream_bytes(),
        after.stream_bytes(),
        unchanged_bytes,
        changed_json,
        container_output,
        changed_only.len(),
        alloc_metrics::instrumented(),
        open_region.json(),
        commit_region.json(),
        container_region.json(),
        changed_only_region.json(),
        json_streams(&before),
        json_streams(&after),
    );
    Ok(())
}

/// Runs the probe. Both binaries call this; only the `peak` binary installs
/// the counting global allocator.
/// Writes a PPT deck authored by `litchi_ppt::writer::Writer`.
///
/// No `.ppt` fixture in this repository accepts a length-changing shape-text
/// replacement: real PowerPoint textboxes carry extra `ClientTextbox` children
/// that set `can_resize = false`, so every target is refused with
/// `Refusal::DependencyClosure`. An authored deck is the only way to reach the
/// shape-text edit, and it is labelled as generated wherever it is reported.
fn emit_ppt_fixture(path: &Path, slides: usize, shapes: usize) -> Result<(), BoxError> {
    let mut writer = litchi_ppt::writer::Writer::new();
    for slide_index in 0..slides {
        let slide = writer.add_slide()?;
        for shape_index in 0..shapes {
            let x = 10 + i32::try_from(shape_index)? * 260;
            let text = format!("slide-{slide_index:02}-shape-{shape_index:02}");
            writer.add_textbox(slide, x, 10, 240, 40, &text)?;
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    std::fs::write(path, output.into_inner())?;
    Ok(())
}

pub fn run() -> Result<(), BoxError> {
    let args = parse_args()?;
    if let Some(path) = &args.emit_ppt_fixture {
        emit_ppt_fixture(path, args.slides, args.shapes)?;
        println!("{{\"emitted\":\"{}\"}}", json_escape(&path.display().to_string()));
        return Ok(());
    }
    let bytes = std::fs::read(Path::new(&args.input))?;

    if args.report {
        return report(&bytes, &args);
    }

    // Prepare the container leg's replacement bytes once, outside the loop, so
    // the timed interval contains only the container rebuild.
    let replacements = if matches!(
        args.operation,
        Operation::Container | Operation::ContainerChangedOnly
    ) {
        let output = commit_once(&bytes, &args)?;
        changed_replacements(&bytes, &output)?
    } else {
        Vec::new()
    };
    let changed_only = if args.operation == Operation::ContainerChangedOnly {
        changed_only_source(&bytes, &replacements)?
    } else {
        Vec::new()
    };

    let iterations = args.warmups + args.samples;
    let mut checksum: u64 = 0;
    for _ in 0..iterations {
        let value = match args.operation {
            Operation::Open => open_once(&bytes, &args)?,
            Operation::Commit => {
                let output = commit_once(&bytes, &args)?;
                let len = output.len();
                black_box(output);
                len
            },
            Operation::Container => container_rebuild(&bytes, &replacements)?,
            Operation::ContainerChangedOnly => container_rebuild(&changed_only, &replacements)?,
        };
        checksum = checksum.wrapping_add(black_box(value) as u64);
    }

    println!(
        "{{\"format\":\"{:?}\",\"operation\":\"{:?}\",\"iterations\":{iterations},\"warmups\":{},\"samples\":{},\"checksum\":{checksum}}}",
        args.format, args.operation, args.warmups, args.samples
    );
    Ok(())
}
