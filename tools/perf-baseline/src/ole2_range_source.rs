//! Opt-in OLE2 range-source selectors (change 0627).
//!
//! Change [0587](../../../docs/performance/0587-remaining-opportunity-survey.md)
//! recorded, as evidence gap 4, that the program has three independent OOXML
//! range-source mechanisms — `SimulatedRangeSource` over OPC and XLSX, a PPTX
//! module, and provider pacing — and **no** OLE2 one: no CFB, XLS, DOC or PPT
//! case ever ran over a caller-supplied remote source. The survey also
//! recorded why that gap is cheap to close: `litchi_xls::SourceBackedWorkbook`
//! and `litchi_ppt::text_edit::SourceSnapshot` already take
//! `Arc<dyn litchi_core::ReadAt>`, and the harness's `SimulatedRangeSource`
//! already implements that trait. Closing it is harness plumbing, not a
//! production-crate change.
//!
//! This module adds fourteen opt-in selectors, none of them in
//! `Case::DEFAULT`:
//!
//! * seven over `SimulatedRangeSource`, which splits every logical read into
//!   physical requests no larger than `--range-max-physical-bytes` and pays
//!   0447/0448's deterministic fixed latency, per-request overhead and
//!   bandwidth for each one, exactly as change
//!   [0572](../../../docs/performance/0572-ooxml-range-source-attribution.md)
//!   does for OPC and XLSX;
//! * seven owned-source controls that run the *same* phases over the same
//!   bytes through a plain in-process `ReadAt`, so the range-source request
//!   sequence can be read as a ratio against the same reader's ordinary
//!   logical reads rather than against a differently shaped corpus.
//!
//! Both families read a caller-named real fixture supplied with
//! `--ole2-file PATH` (repeatable, at most one file per format). Like 0601's
//! `--real-file` opt-in, this is an input whose bytes do not come from the
//! process, so the corpus identity records the path, the size and the SHA-256,
//! every scenario target is derived from the file rather than assumed, and the
//! selectors stay out of the default matrix.
//!
//! Every case is a complete fresh-open lifecycle — a fresh source, a fresh
//! owner, then the operation — because that is the shape the existing XLS
//! source-backed family (`xls_source_backed_open_list_worksheets` and
//! siblings) and the standalone `xls_source_attribution` profiler already use,
//! and because a caller paying per request cares about the cost of *reaching*
//! a cell, not about a marginal query on an already-open workbook.
//!
//! These selectors take no timing, allocation, physical-I/O, cold-cache or
//! speedup claim. They are a descriptive baseline: the first measurement of
//! any OLE2 reader over a range source.

use std::{
    error::Error,
    fs,
    io::Cursor,
    path::Path,
    sync::Arc,
    time::{Duration, Instant},
};

use litchi_core::{OwnedSource, Position, ReadAt, sheet::CellValue};
use serde::Serialize;

use crate::{
    Case, CaseResult, Corpus, CorpusManifest, InstrumentedSource, RangeSimulationConfig,
    RangeSimulationSnapshot, SourceSummary, boxed_source, iteration_count,
    producer_shape::RealFileProvenance, record_elapsed, sha256_hex, simulated_source, statistics,
    verify_simulation_snapshot,
};

/// Generator identity of the XLS range-source family. Nothing in it is
/// generated, so the identity says so.
pub(crate) const XLS_REAL_FILE_GENERATOR: &str = "litchi-xls-real-file-v1";
/// Generator identity of the PPT range-source family.
pub(crate) const PPT_REAL_FILE_GENERATOR: &str = "litchi-ppt-real-file-v1";

/// Largest `--ole2-file` input this harness will read. A caller-named file is
/// still a bounded resource; the limit matches 0601's `--real-file` bound.
const MAX_OLE2_FILE_BYTES: u64 = 32 * 1024 * 1024;

/// Ranges retained verbatim from the first sample's physical request
/// sequence. The complete sequence identity travels as a digest; this prefix
/// exists so a record can quote the first offsets a reader asks for without
/// the report carrying tens of thousands of triples per sample.
const REQUEST_SEQUENCE_PREVIEW: usize = 32;

/// Which OLE2 reader a corpus and a scenario belong to.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Format {
    Xls,
    Ppt,
}

impl Format {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Xls => "XLS",
            Self::Ppt => "PPT",
        }
    }

    const fn package_format(self) -> &'static str {
        match self {
            Self::Xls => "XLS/CFB/OLE2",
            Self::Ppt => "PPT/CFB/OLE2",
        }
    }

    const fn generator(self) -> &'static str {
        match self {
            Self::Xls => XLS_REAL_FILE_GENERATOR,
            Self::Ppt => PPT_REAL_FILE_GENERATOR,
        }
    }

    const fn corpus_name(self) -> &'static str {
        match self {
            Self::Xls => "xls-real-file",
            Self::Ppt => "ppt-real-file",
        }
    }
}

/// One measured phase. Every variant includes a fresh open.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Scenario {
    /// `SourceBackedWorkbook::from_read_at`.
    XlsOpen,
    /// Open, then `worksheet_names()`.
    XlsListWorksheets,
    /// Open, then one `cell_value_by_index` at the derived target.
    XlsOneCell,
    /// Open, then `SourceBackedWorksheet::visit_cells` over one worksheet:
    /// change 0605's whole-sheet walk.
    XlsAllCells,
    /// Open, then `SourceBackedWorkbook::text()`.
    XlsFullText,
    /// `litchi_ppt::text_edit::SourceSnapshot::open`.
    PptOpen,
    /// Open, then `read_text` of the derived slide/shape target.
    PptOneShapeText,
}

impl Scenario {
    const fn as_str(self) -> &'static str {
        match self {
            Self::XlsOpen | Self::PptOpen => "open",
            Self::XlsListWorksheets => "open+list-worksheets",
            Self::XlsOneCell => "open+one-cell",
            Self::XlsAllCells => "open+all-cells",
            Self::XlsFullText => "open+full-text",
            Self::PptOneShapeText => "open+one-shape-text",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::XlsOpen => "fresh-source-and-source-backed-workbook-open",
            Self::XlsListWorksheets => "fresh-source-open-plus-worksheet-names",
            Self::XlsOneCell => "fresh-source-open-plus-one-selected-cell",
            Self::XlsAllCells => "fresh-source-open-plus-one-worksheet-walk",
            Self::XlsFullText => "fresh-source-open-plus-complete-text",
            Self::PptOpen => "fresh-source-and-source-snapshot-open",
            Self::PptOneShapeText => "fresh-source-open-plus-selected-shape-text",
        }
    }

    pub(crate) const fn format(self) -> Format {
        match self {
            Self::XlsOpen
            | Self::XlsListWorksheets
            | Self::XlsOneCell
            | Self::XlsAllCells
            | Self::XlsFullText => Format::Xls,
            Self::PptOpen | Self::PptOneShapeText => Format::Ppt,
        }
    }
}

/// How the reader reaches the bytes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Transport {
    /// `SimulatedRangeSource`: bounded physical requests, each paying the
    /// configured latency, overhead and bandwidth.
    RangeSource,
    /// A plain instrumented in-process `ReadAt`: the control leg.
    OwnedSource,
}

impl Transport {
    const fn as_str(self) -> &'static str {
        match self {
            Self::RangeSource => "simulated-range-source",
            Self::OwnedSource => "owned-source-control",
        }
    }

    const fn counter_scope(self) -> &'static str {
        match self {
            Self::RangeSource => {
                "logical reads are the reader's own calls; physical requests are the simulator's \
                 bounded splits of them, each paying fixed latency, per-request overhead and \
                 bandwidth"
            },
            Self::OwnedSource => {
                "logical reads are the reader's own calls against an in-process source; there is \
                 no physical request model and no simulated delay"
            },
        }
    }
}

/// Everything a range-source corpus states about itself beyond the ordinary
/// [`CorpusManifest`]: where the bytes came from, what the CFB holds, and the
/// scenario targets derived from it.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct Ole2Evidence {
    pub(crate) generator: &'static str,
    pub(crate) format: &'static str,
    pub(crate) real_file: RealFileProvenance,
    pub(crate) cfb_stream_count: usize,
    pub(crate) cfb_sector_size: usize,
    pub(crate) target_stream: String,
    pub(crate) target_stream_bytes: usize,
    pub(crate) target_stream_sha256: String,
    /// XLS: the ordinary worksheet names the reader reports.
    pub(crate) worksheet_names: Vec<String>,
    /// XLS: the worksheet the one-cell and all-cells scenarios select.
    pub(crate) selected_worksheet_index: Option<usize>,
    pub(crate) selected_row: Option<u32>,
    pub(crate) selected_column: Option<u32>,
    /// PPT: the slide and shape the one-shape-text scenario selects.
    pub(crate) selected_slide: Option<usize>,
    pub(crate) selected_shape: Option<usize>,
    /// Frozen oracle projection per scenario, derived once from the file.
    pub(crate) scenario_oracles: Vec<ScenarioOracle>,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct ScenarioOracle {
    pub(crate) scenario: &'static str,
    pub(crate) observation: String,
}

/// A range-source corpus: an ordinary [`Corpus`] plus its evidence block.
#[derive(Debug)]
pub(crate) struct Ole2Corpus {
    pub(crate) corpus: Corpus,
    pub(crate) evidence: Ole2Evidence,
    format: Format,
    selected_worksheet_index: usize,
    selected_row: u32,
    selected_column: u32,
    selected_slide: usize,
    selected_shape: usize,
}

/// Per-case evidence written into `source.ole2_range_source`.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct Ole2RangeSourceSummary {
    pub(crate) format: &'static str,
    pub(crate) transport: &'static str,
    pub(crate) scenario: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) source_counter_scope: &'static str,
    /// Present only on the range-source leg.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) transport_parameters: Option<RangeSimulationConfig>,
    /// The corpus identity and the derived scenario targets, restated per
    /// case so one result is self-describing.
    pub(crate) corpus: Ole2Evidence,
    /// The reader's own read calls and bytes, comparable across both legs.
    pub(crate) logical_read_calls: Vec<u64>,
    pub(crate) logical_read_bytes: Vec<u64>,
    /// Number of bounded physical requests per retained sample; empty on the
    /// owned-source control.
    pub(crate) physical_request_count: Vec<u64>,
    pub(crate) physical_request_bytes: Vec<u64>,
    /// SHA-256 of the complete `(offset, requested, returned)` request
    /// sequence in call order, one per retained sample.
    pub(crate) request_sequence_sha256: Vec<String>,
    pub(crate) request_sequence_identical: bool,
    /// First [`REQUEST_SEQUENCE_PREVIEW`] triples of the first retained
    /// sample's sequence.
    pub(crate) request_sequence_preview: Vec<[u64; 3]>,
    /// The frozen oracle, derived once from the file before any sample runs.
    pub(crate) observation: String,
    /// SHA-256 of each retained sample's own projection, and whether all of
    /// them reproduced the oracle.
    pub(crate) observation_sha256: Vec<String>,
    pub(crate) observations_identical: bool,
    /// Sum of the configured per-request service times over the first
    /// retained sample's request sequence: the part of elapsed time the
    /// transport model, not the reader, accounts for.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) simulated_service_floor_ns: Option<u64>,
}

// ---------------------------------------------------------------------------
// Corpus construction
// ---------------------------------------------------------------------------

fn read_bounded(path: &Path) -> Result<Vec<u8>, Box<dyn Error>> {
    let metadata = fs::metadata(path)
        .map_err(|source| format!("--ole2-file {} is unreadable: {source}", path.display()))?;
    if !metadata.is_file() {
        return Err(format!("--ole2-file {} is not a regular file", path.display()).into());
    }
    if metadata.len() > MAX_OLE2_FILE_BYTES {
        return Err(format!(
            "--ole2-file {} is {} bytes, above the {MAX_OLE2_FILE_BYTES}-byte bound",
            path.display(),
            metadata.len()
        )
        .into());
    }
    fs::read(path).map_err(|source| {
        format!("--ole2-file {} could not be read: {source}", path.display()).into()
    })
}

/// Reads the CFB inventory: the stream paths, the sector size, and the bytes
/// of the named target stream. Everything here is untimed and happens once.
fn cfb_inventory(
    archive: &[u8],
    candidates: &[&[&str]],
    label: &str,
) -> Result<(Vec<Vec<String>>, usize, Vec<String>, Vec<u8>), Box<dyn Error>> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(archive))?;
    let streams = ole.list_streams();
    let sector_size = ole.sector_size();
    let path = candidates
        .iter()
        .find(|candidate| ole.exists(candidate))
        .ok_or_else(|| format!("{label} fixture has no {label} stream"))?;
    let bytes = ole.open_stream(path)?;
    let owned = path.iter().map(|part| (*part).to_owned()).collect();
    Ok((streams, sector_size, owned, bytes))
}

fn manifest_of(
    format: Format,
    archive: &[u8],
    streams: &[Vec<String>],
    sector_size: usize,
    target_entry: &str,
    target_payload: &[u8],
    entry_count: usize,
) -> CorpusManifest {
    CorpusManifest {
        name: format.corpus_name().to_owned(),
        generator: format.generator(),
        package_format: format.package_format(),
        shape: "real-file",
        payload_kind: "real-producer-ole2",
        compression: "none",
        entry_count,
        archive_member_count: streams.len(),
        entry_bytes: sector_size,
        uncompressed_payload_bytes: target_payload.len(),
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(archive),
        target_entry: target_entry.to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: sha256_hex(target_payload),
        rtf_variant: None,
        xlsx: None,
    }
}

fn provenance_of(path: &Path, archive: &[u8]) -> Result<RealFileProvenance, Box<dyn Error>> {
    Ok(RealFileProvenance {
        path: path.display().to_string(),
        bytes: u64::try_from(archive.len())
            .map_err(|_error| "OLE2 fixture length does not fit u64")?,
        sha256: sha256_hex(archive),
    })
}

fn cell_projection(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "empty".to_owned(),
        CellValue::Bool(value) => format!("bool:{value}"),
        CellValue::Int(value) => format!("int:{value}"),
        CellValue::Float(value) => format!("float:{:016x}", value.to_bits()),
        CellValue::String(value) => format!("string:{}:{value}", value.len()),
        CellValue::DateTime(value) => format!("datetime:{value}"),
        CellValue::Error(value) => format!("error:{}:{value}", value.len()),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => {
            let cached = cached_value
                .as_deref()
                .map_or_else(|| "none".to_owned(), cell_projection);
            format!("formula:{}:{formula}:cached:{cached}", formula.len())
        },
    }
}

/// Position-keyed projection of a walked worksheet. The walk reports records
/// in stream order and may report one position twice; the last record wins,
/// exactly as a selected-cell query reports the last record it saw.
fn cells_digest(cells: &[(u32, u32, String)]) -> String {
    let mut by_position = std::collections::BTreeMap::new();
    for (row, column, projection) in cells {
        let _ = by_position.insert((*row, *column), projection.clone());
    }
    let mut buffer = Vec::new();
    for ((row, column), projection) in &by_position {
        buffer.extend_from_slice(format!("{row},{column}={projection}\n").as_bytes());
    }
    format!("cells:{}:{}", by_position.len(), sha256_hex(&buffer))
}

fn names_digest(names: &[String]) -> String {
    let mut buffer = Vec::new();
    for name in names {
        buffer.extend_from_slice(&(name.len() as u64).to_le_bytes());
        buffer.extend_from_slice(name.as_bytes());
    }
    format!("names:{}:{}", names.len(), sha256_hex(&buffer))
}

fn text_outcome(text: &Result<String, impl std::fmt::Display>) -> String {
    match text {
        Ok(text) => format!("text:{}:{}", text.len(), sha256_hex(text.as_bytes())),
        Err(error) => format!("refused:{error}"),
    }
}

fn walk_worksheet(
    workbook: &litchi_xls::SourceBackedWorkbook,
    worksheet_index: usize,
) -> Result<Vec<(u32, u32, String)>, Box<dyn Error>> {
    let worksheet = workbook
        .worksheet_by_index(worksheet_index)?
        .ok_or("XLS range-source worksheet index is out of range")?;
    let mut cells = Vec::new();
    worksheet.visit_cells(|cell| {
        cells.push((cell.row(), cell.column(), cell_projection(cell.value())));
        Ok(())
    })?;
    Ok(cells)
}

/// Builds the XLS corpus from a caller-named fixture.
///
/// Nothing about the file is assumed. The worksheet the cell scenarios select
/// is the first ordinary worksheet whose walk reports at least one cell, and
/// the selected cell is the median position of that walk, so the target is
/// derived from the bytes and is stable for a given fixture.
pub(crate) fn build_xls_corpus(path: &Path) -> Result<Ole2Corpus, Box<dyn Error>> {
    let archive = read_bounded(path)?;
    let provenance = provenance_of(path, &archive)?;
    let (streams, sector_size, target_path, target_payload) =
        cfb_inventory(&archive, &[&["Workbook"], &["Book"]], "Workbook")?;
    let target_entry = target_path.join("/");

    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
        archive.clone(),
    )))?;
    let worksheet_names = workbook.worksheet_names()?;
    if worksheet_names.is_empty() {
        return Err("--ole2-file XLS fixture declares no ordinary worksheet".into());
    }
    // The first ordinary worksheet whose walk reports a cell with a value,
    // falling back to the first that reports any cell at all.
    let mut selected = None;
    let mut fallback = None;
    for index in 0..worksheet_names.len() {
        let cells = walk_worksheet(&workbook, index)?;
        if cells.is_empty() {
            continue;
        }
        if cells.iter().any(|(_, _, projection)| projection != "empty") {
            selected = Some((index, cells));
            break;
        }
        if fallback.is_none() {
            fallback = Some((index, cells));
        }
    }
    let selected = selected.or(fallback);
    let (selected_worksheet_index, walked) =
        selected.ok_or("--ole2-file XLS fixture has no worksheet with a stored cell")?;
    // Prefer a position the reader reports with a value: a worksheet may
    // store blank records, and a target whose projection is `empty` would
    // measure the scan without proving a value came back.
    let mut positions = walked
        .iter()
        .filter(|(_, _, projection)| projection != "empty")
        .map(|(row, column, _)| (*row, *column))
        .collect::<Vec<_>>();
    if positions.is_empty() {
        positions = walked
            .iter()
            .map(|(row, column, _)| (*row, *column))
            .collect();
    }
    positions.sort_unstable();
    positions.dedup();
    let (selected_row, selected_column) = positions[positions.len() / 2];
    let selected_value =
        workbook.cell_value_by_index(selected_worksheet_index, selected_row, selected_column)?;
    let one_cell = selected_value
        .as_ref()
        .map_or_else(|| "absent".to_owned(), cell_projection);
    let full_text = text_outcome(&workbook.text());
    let worksheet_count = workbook.worksheet_count()?;
    drop(workbook);

    let scenario_oracles = vec![
        ScenarioOracle {
            scenario: Scenario::XlsOpen.as_str(),
            observation: format!("worksheets:{worksheet_count}"),
        },
        ScenarioOracle {
            scenario: Scenario::XlsListWorksheets.as_str(),
            observation: names_digest(&worksheet_names),
        },
        ScenarioOracle {
            scenario: Scenario::XlsOneCell.as_str(),
            observation: one_cell,
        },
        ScenarioOracle {
            scenario: Scenario::XlsAllCells.as_str(),
            observation: cells_digest(&walked),
        },
        ScenarioOracle {
            scenario: Scenario::XlsFullText.as_str(),
            observation: full_text,
        },
    ];

    let evidence = Ole2Evidence {
        generator: XLS_REAL_FILE_GENERATOR,
        format: Format::Xls.as_str(),
        real_file: provenance,
        cfb_stream_count: streams.len(),
        cfb_sector_size: sector_size,
        target_stream: target_entry.clone(),
        target_stream_bytes: target_payload.len(),
        target_stream_sha256: sha256_hex(&target_payload),
        worksheet_names,
        selected_worksheet_index: Some(selected_worksheet_index),
        selected_row: Some(selected_row),
        selected_column: Some(selected_column),
        selected_slide: None,
        selected_shape: None,
        scenario_oracles,
    };
    let manifest = manifest_of(
        Format::Xls,
        &archive,
        &streams,
        sector_size,
        &target_entry,
        &target_payload,
        worksheet_count,
    );
    Ok(Ole2Corpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: target_entry,
            target_payload,
            xlsx: None,
        },
        evidence,
        format: Format::Xls,
        selected_worksheet_index,
        selected_row,
        selected_column,
        selected_slide: 0,
        selected_shape: 0,
    })
}

/// Builds the PPT corpus from a caller-named fixture.
///
/// The selected shape is the first slide/shape position, in the eager
/// reader's source order, whose text that reader reports as non-empty — a
/// structural fact about the deck, derived from the bytes rather than
/// assumed. Whether the *source-backed* reader admits that position is a
/// separate question and is not required: a typed refusal is frozen as the
/// scenario's outcome, exactly as change 0605's XLS full-text scenario
/// freezes one, because a caller paying per request still pays for the bytes
/// the reader reads before refusing.
pub(crate) fn build_ppt_corpus(path: &Path) -> Result<Ole2Corpus, Box<dyn Error>> {
    use litchi_ppt::text_edit::{SourceSnapshot, Target};

    let archive = read_bounded(path)?;
    let provenance = provenance_of(path, &archive)?;
    let (streams, sector_size, target_path, target_payload) = cfb_inventory(
        &archive,
        &[
            &["PowerPoint Document"],
            &["PP97_DUALSTORAGE", "PowerPoint Document"],
        ],
        "PowerPoint Document",
    )?;
    let target_entry = target_path.join("/");

    let mut package = litchi_ppt::Package::from_reader(Cursor::new(archive.as_slice()))?;
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let slide_count = slides.len();
    let mut chosen = None;
    'outer: for (slide_index, slide) in slides.iter().enumerate() {
        for (shape_index, shape) in slide.shapes()?.iter().enumerate() {
            if shape.text().is_ok_and(|text| !text.is_empty()) {
                chosen = Some((slide_index, shape_index));
                break 'outer;
            }
        }
    }
    drop(slides);
    drop(presentation);
    drop(package);
    let (selected_slide, selected_shape) =
        chosen.ok_or("--ole2-file PPT fixture has no slide shape carrying text")?;

    let snapshot = SourceSnapshot::open(Arc::new(OwnedSource::new(archive.clone())))?;
    let target = Target::new(Position::new(selected_slide), Position::new(selected_shape));
    let one_shape_text = text_outcome(&snapshot.read_text(target));
    drop(snapshot);

    let scenario_oracles = vec![
        ScenarioOracle {
            scenario: Scenario::PptOpen.as_str(),
            observation: format!("slides:{slide_count}"),
        },
        ScenarioOracle {
            scenario: Scenario::PptOneShapeText.as_str(),
            observation: one_shape_text,
        },
    ];

    let evidence = Ole2Evidence {
        generator: PPT_REAL_FILE_GENERATOR,
        format: Format::Ppt.as_str(),
        real_file: provenance,
        cfb_stream_count: streams.len(),
        cfb_sector_size: sector_size,
        target_stream: target_entry.clone(),
        target_stream_bytes: target_payload.len(),
        target_stream_sha256: sha256_hex(&target_payload),
        worksheet_names: Vec::new(),
        selected_worksheet_index: None,
        selected_row: None,
        selected_column: None,
        selected_slide: Some(selected_slide),
        selected_shape: Some(selected_shape),
        scenario_oracles,
    };
    let manifest = manifest_of(
        Format::Ppt,
        &archive,
        &streams,
        sector_size,
        &target_entry,
        &target_payload,
        slide_count,
    );
    Ok(Ole2Corpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: target_entry,
            target_payload,
            xlsx: None,
        },
        evidence,
        format: Format::Ppt,
        selected_worksheet_index: 0,
        selected_row: 0,
        selected_column: 0,
        selected_slide,
        selected_shape,
    })
}

/// Decides which OLE2 reader a caller-named file belongs to by its CFB
/// stream inventory, so `--ole2-file` needs no format flag and no extension
/// heuristic. A file that offers both, or neither, is refused.
pub(crate) fn classify(path: &Path) -> Result<Format, Box<dyn Error>> {
    let archive = read_bounded(path)?;
    let ole = litchi_cfb::OleFile::open(Cursor::new(archive.as_slice()))?;
    let workbook = ole.exists(&["Workbook"]) || ole.exists(&["Book"]);
    let document = ole.exists(&["PowerPoint Document"])
        || ole.exists(&["PP97_DUALSTORAGE", "PowerPoint Document"]);
    match (workbook, document) {
        (true, false) => Ok(Format::Xls),
        (false, true) => Ok(Format::Ppt),
        (true, true) => Err(format!(
            "--ole2-file {} holds both a Workbook and a PowerPoint Document stream",
            path.display()
        )
        .into()),
        (false, false) => Err(format!(
            "--ole2-file {} holds neither a Workbook nor a PowerPoint Document stream",
            path.display()
        )
        .into()),
    }
}

// ---------------------------------------------------------------------------
// Measurement
// ---------------------------------------------------------------------------

/// Everything one retained iteration contributes.
struct Observation {
    duration: Duration,
    projection: String,
    logical_read_calls: u64,
    logical_read_bytes: u64,
    simulation: Option<RangeSimulationSnapshot>,
}

fn request_sequence_sha256(ranges: &[[u64; 3]]) -> String {
    let mut buffer = Vec::with_capacity(ranges.len() * 24);
    for range in ranges {
        for field in range {
            buffer.extend_from_slice(&field.to_le_bytes());
        }
    }
    format!("ranges:{}:{}", ranges.len(), sha256_hex(&buffer))
}

fn oracle_for(corpus: &Ole2Corpus, scenario: Scenario) -> Result<&str, Box<dyn Error>> {
    corpus
        .evidence
        .scenario_oracles
        .iter()
        .find(|oracle| oracle.scenario == scenario.as_str())
        .map(|oracle| oracle.observation.as_str())
        .ok_or_else(|| format!("{} scenario has no frozen oracle", scenario.as_str()).into())
}

pub(crate) fn run_case(
    case: Case,
    scenario: Scenario,
    transport: Transport,
    corpus: &Ole2Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    if scenario.format() != corpus.format {
        return Err(format!(
            "{} scenario cannot run on a {} corpus",
            scenario.as_str(),
            corpus.format.as_str()
        )
        .into());
    }
    let oracle = oracle_for(corpus, scenario)?.to_owned();

    let mut elapsed = Vec::with_capacity(samples);
    let mut summary = SourceSummary::default();
    let mut logical_read_calls = Vec::with_capacity(samples);
    let mut logical_read_bytes = Vec::with_capacity(samples);
    let mut physical_request_count = Vec::with_capacity(samples);
    let mut physical_request_bytes = Vec::with_capacity(samples);
    let mut request_sequence = Vec::with_capacity(samples);
    let mut preview = Vec::new();
    let mut observation_sha256 = Vec::with_capacity(samples);
    let mut simulated_service_floor_ns = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = Arc::new(InstrumentedSource::new(
            corpus.corpus.archive.clone(),
            Vec::new(),
        ));
        let observation = match transport {
            Transport::RangeSource => {
                let source = simulated_source(Arc::clone(&backing), config);
                let started = Instant::now();
                let projection = measure(corpus, scenario, Arc::clone(&source) as Arc<dyn ReadAt>)?;
                let duration = started.elapsed();
                let simulation = source.snapshot()?;
                verify_simulation_snapshot(
                    &simulation,
                    config,
                    &format!("simulated {} {}", corpus.format.as_str(), scenario.as_str()),
                )?;
                Observation {
                    duration,
                    projection,
                    logical_read_calls: simulation.logical_read_calls,
                    logical_read_bytes: simulation.logical_read_bytes,
                    simulation: Some(simulation),
                }
            },
            Transport::OwnedSource => {
                let started = Instant::now();
                let projection =
                    measure(corpus, scenario, Arc::clone(&backing) as Arc<dyn ReadAt>)?;
                let duration = started.elapsed();
                let metrics = backing.snapshot();
                if metrics.read_calls == 0 || metrics.read_bytes == 0 {
                    return Err(format!(
                        "owned-source control {} {} performed no source I/O",
                        corpus.format.as_str(),
                        scenario.as_str()
                    )
                    .into());
                }
                Observation {
                    duration,
                    projection,
                    logical_read_calls: metrics.read_calls,
                    logical_read_bytes: metrics.read_bytes,
                    simulation: None,
                }
            },
        };
        if iteration >= warmup_iterations {
            observation_sha256.push(sha256_hex(observation.projection.as_bytes()));
            summary.record(backing.snapshot());
            logical_read_calls.push(observation.logical_read_calls);
            logical_read_bytes.push(observation.logical_read_bytes);
            if let Some(simulation) = observation.simulation {
                physical_request_count.push(simulation.physical_request_count);
                physical_request_bytes.push(simulation.physical_request_bytes);
                request_sequence.push(request_sequence_sha256(&simulation.physical_ranges));
                if preview.is_empty() {
                    preview = simulation
                        .physical_ranges
                        .iter()
                        .take(REQUEST_SEQUENCE_PREVIEW)
                        .copied()
                        .collect();
                    simulated_service_floor_ns =
                        Some(crate::simulated_service_floor_ns(&simulation, config)?);
                }
                summary.record_simulation(simulation);
            }
        }
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            observation.duration,
        )?;
    }

    let oracle_sha256 = sha256_hex(oracle.as_bytes());
    let observations_identical = observation_sha256
        .iter()
        .all(|entry| *entry == oracle_sha256);
    if !observations_identical {
        return Err(format!(
            "{} {} produced a sample that differs from the frozen oracle",
            corpus.format.as_str(),
            scenario.as_str()
        )
        .into());
    }

    let request_sequence_identical = request_sequence
        .first()
        .is_none_or(|first| request_sequence.iter().all(|entry| entry == first));
    if !request_sequence_identical {
        return Err(format!(
            "simulated {} {} issued a different physical request sequence across repeats",
            corpus.format.as_str(),
            scenario.as_str()
        )
        .into());
    }

    summary.ole2_range_source = Some(Box::new(Ole2RangeSourceSummary {
        format: corpus.format.as_str(),
        transport: transport.as_str(),
        scenario: scenario.as_str(),
        timing_scope: scenario.timing_scope(),
        source_counter_scope: transport.counter_scope(),
        transport_parameters: matches!(transport, Transport::RangeSource).then_some(config),
        corpus: corpus.evidence.clone(),
        logical_read_calls,
        logical_read_bytes,
        physical_request_count,
        physical_request_bytes,
        request_sequence_sha256: request_sequence,
        request_sequence_identical,
        request_sequence_preview: preview,
        observation: oracle,
        observation_sha256,
        observations_identical,
        simulated_service_floor_ns,
    }));

    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: boxed_source(summary),
        execution: None,
        output_sha256: None,
        operation_metrics: None,
    })
}

/// Runs one measured phase against the supplied source and returns the
/// scenario's projection. Every branch opens a fresh owner.
fn measure(
    corpus: &Ole2Corpus,
    scenario: Scenario,
    source: Arc<dyn ReadAt>,
) -> Result<String, Box<dyn Error>> {
    match scenario {
        Scenario::XlsOpen => {
            let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
            let count = std::hint::black_box(workbook.worksheet_count()?);
            Ok(format!("worksheets:{count}"))
        },
        Scenario::XlsListWorksheets => {
            let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
            let names = std::hint::black_box(workbook.worksheet_names()?);
            Ok(names_digest(&names))
        },
        Scenario::XlsOneCell => {
            let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
            let value = std::hint::black_box(workbook.cell_value_by_index(
                corpus.selected_worksheet_index,
                corpus.selected_row,
                corpus.selected_column,
            )?);
            Ok(value
                .as_ref()
                .map_or_else(|| "absent".to_owned(), cell_projection))
        },
        Scenario::XlsAllCells => {
            let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
            let cells =
                std::hint::black_box(walk_worksheet(&workbook, corpus.selected_worksheet_index)?);
            Ok(cells_digest(&cells))
        },
        Scenario::XlsFullText => {
            let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
            Ok(text_outcome(&std::hint::black_box(workbook.text())))
        },
        Scenario::PptOpen => {
            let snapshot = litchi_ppt::text_edit::SourceSnapshot::open(source)?;
            std::hint::black_box(&snapshot);
            Ok(corpus
                .evidence
                .scenario_oracles
                .iter()
                .find(|oracle| oracle.scenario == Scenario::PptOpen.as_str())
                .map(|oracle| oracle.observation.clone())
                .ok_or("PPT open scenario has no frozen oracle")?)
        },
        Scenario::PptOneShapeText => {
            use litchi_ppt::text_edit::{SourceSnapshot, Target};
            let snapshot = SourceSnapshot::open(source)?;
            let target = Target::new(
                Position::new(corpus.selected_slide),
                Position::new(corpus.selected_shape),
            );
            Ok(text_outcome(&std::hint::black_box(
                snapshot.read_text(target),
            )))
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(relative: &str) -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative)
    }

    #[test]
    fn scenarios_and_transports_describe_themselves() {
        assert_eq!(Scenario::XlsOpen.format(), Format::Xls);
        assert_eq!(Scenario::PptOneShapeText.format(), Format::Ppt);
        assert_eq!(Scenario::XlsAllCells.as_str(), "open+all-cells");
        assert_eq!(Transport::RangeSource.as_str(), "simulated-range-source");
        assert_eq!(Transport::OwnedSource.as_str(), "owned-source-control");
    }

    #[test]
    fn every_ole2_range_source_case_is_opt_in_and_round_trips() {
        let mut seen = 0;
        for case in [
            Case::XlsRangeSourceOpen,
            Case::XlsRangeSourceOpenListWorksheets,
            Case::XlsRangeSourceOpenOneCell,
            Case::XlsRangeSourceOpenAllCells,
            Case::XlsRangeSourceOpenFullText,
            Case::XlsOwnedSourceControlOpen,
            Case::XlsOwnedSourceControlOpenListWorksheets,
            Case::XlsOwnedSourceControlOpenOneCell,
            Case::XlsOwnedSourceControlOpenAllCells,
            Case::XlsOwnedSourceControlOpenFullText,
            Case::PptRangeSourceOpen,
            Case::PptRangeSourceOpenOneShapeText,
            Case::PptOwnedSourceControlOpen,
            Case::PptOwnedSourceControlOpenOneShapeText,
        ] {
            assert!(case.is_ole2_range_source());
            assert!(!Case::DEFAULT.contains(&case));
            assert_eq!(crate::parse_case(case.name()), Some(case));
            seen += 1;
        }
        assert_eq!(seen, 14);
    }

    #[test]
    fn xls_corpus_derives_its_targets_from_the_file() {
        let path = fixture("test-data/ole/xls/WithCustomViews.xls");
        let corpus = build_xls_corpus(&path).unwrap();
        assert_eq!(corpus.format, Format::Xls);
        assert_eq!(corpus.corpus.manifest.generator, XLS_REAL_FILE_GENERATOR);
        assert!(!corpus.evidence.worksheet_names.is_empty());
        assert_eq!(corpus.evidence.scenario_oracles.len(), 5);
        assert_eq!(corpus.evidence.target_stream, "Workbook");
        assert!(corpus.evidence.target_stream_bytes > 0);
        assert_eq!(classify(&path).unwrap(), Format::Xls);
        assert_eq!(
            corpus.evidence.real_file.sha256,
            corpus.corpus.manifest.archive_sha256
        );
    }

    #[test]
    fn xls_range_source_and_owned_legs_agree_on_logical_reads() {
        let path = fixture("test-data/ole/xls/WithCustomViews.xls");
        let corpus = build_xls_corpus(&path).unwrap();
        let config = RangeSimulationConfig {
            fixed_latency_us: 0,
            request_overhead_us: 0,
            bandwidth_bytes_per_second: 1 << 40,
            max_physical_range_bytes: 4096,
        };
        let simulated = run_case(
            Case::XlsRangeSourceOpen,
            Scenario::XlsOpen,
            Transport::RangeSource,
            &corpus,
            0,
            2,
            config,
        )
        .unwrap();
        let owned = run_case(
            Case::XlsOwnedSourceControlOpen,
            Scenario::XlsOpen,
            Transport::OwnedSource,
            &corpus,
            0,
            2,
            config,
        )
        .unwrap();
        let simulated_source = simulated.source.unwrap().ole2_range_source.unwrap();
        let owned_source = owned.source.unwrap().ole2_range_source.unwrap();
        assert_eq!(
            simulated_source.logical_read_calls,
            owned_source.logical_read_calls
        );
        assert_eq!(
            simulated_source.logical_read_bytes,
            owned_source.logical_read_bytes
        );
        assert!(simulated_source.request_sequence_identical);
        assert!(
            simulated_source.physical_request_count[0] >= simulated_source.logical_read_calls[0]
        );
        assert!(owned_source.physical_request_count.is_empty());
        assert_eq!(simulated_source.observation, owned_source.observation);
    }

    #[test]
    fn ppt_corpus_selects_a_text_shape_and_freezes_its_outcome() {
        let path = fixture("test-data/ole/ppt/SampleShow.ppt");
        let corpus = build_ppt_corpus(&path).unwrap();
        assert_eq!(corpus.format, Format::Ppt);
        assert_eq!(corpus.corpus.manifest.generator, PPT_REAL_FILE_GENERATOR);
        assert_eq!(corpus.evidence.target_stream, "PowerPoint Document");
        assert_eq!(classify(&path).unwrap(), Format::Ppt);
        assert_eq!(corpus.evidence.scenario_oracles.len(), 2);
        assert!(corpus.evidence.selected_slide.is_some());
        let config = RangeSimulationConfig {
            fixed_latency_us: 0,
            request_overhead_us: 0,
            bandwidth_bytes_per_second: 1 << 40,
            max_physical_range_bytes: 4096,
        };
        let result = run_case(
            Case::PptRangeSourceOpenOneShapeText,
            Scenario::PptOneShapeText,
            Transport::RangeSource,
            &corpus,
            0,
            2,
            config,
        )
        .unwrap();
        let source = result.source.unwrap().ole2_range_source.unwrap();
        assert!(source.request_sequence_identical);
        assert!(
            source.observation.starts_with("text:") || source.observation.starts_with("refused:"),
            "unexpected PPT one-shape-text outcome {}",
            source.observation
        );
    }
}
