//! Supplemental valid-source event-cap guard for the 0544 XLSX campaign.
//!
//! The root harness copies this file temporarily to
//! `crates/litchi-xlsx/examples/perf_cap_boundary.rs` and builds it against
//! the frozen baseline and candidate.  It deliberately uses only public
//! `OwnedSource`, `SourceBackedEditor`, and the existing `soapberry-zip` and
//! `serde_json` dependencies; no Cargo metadata is changed.
//!
//! `--size 160` produces `5*N*N + 2*N + 6 == 128_326` XML events, below the
//! shared provisional cap.  `--size 164` produces 134_814 events and
//! `--size 256` produces 328_198 events, both above that cap and below the
//! ordinary one-million-event parser bound.  Every fixture is a deterministic
//! single-worksheet, stored-ZIP package with an ASCII/UTF-8 marker-free source
//! worksheet and a dense numeric `N x N` grid.
//!
//! Only `edit_sheets` is timed.  The source/editor and selector are prepared
//! before the clock, and the returned transaction remains alive until the
//! clock stops.  Cell inspection, empty commit, source-byte comparison, and
//! exact no-op publication are outside the measured interval.  The optional
//! `--fixture-out PATH` writes the exact archive before timing; the parent
//! harness binds that file's SHA-256 because this standalone example has no
//! hash dependency.  The JSON report therefore records byte lengths and the
//! deterministic XML/package identity rather than a locally calculated hash.

use std::env;
use std::error::Error;
use std::ffi::OsString;
use std::fmt::Write as _;
use std::fs;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::OwnedSource;
use litchi_xlsx::cell_values::SourceBackedEditor;
use litchi_xlsx::{Address, Number, Selector, Value};
use quick_xml::Reader;
use quick_xml::events::Event;
use serde_json::{Value as JsonValue, json};
use soapberry_zip::office::StreamingArchiveWriter;

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WORKSHEET_PART: &str = "xl/worksheets/sheet1.xml";
const SHARED_SOURCE_BYTES: usize = 8 * 1024 * 1024;
const SHARED_EVENT_CAP: usize = 131_072;
const RAW_EVENT_CAP: usize = 1_000_000;
const MAX_WARMUP: usize = 1_000;
const MAX_SAMPLES: usize = 1_000;

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const ALTERNATE_CONTENT: &[u8] = b"AlternateContent";
const DY_DESCENT: &[u8] = b"dyDescent";

#[derive(Debug)]
struct Arguments {
    size: usize,
    warmup: usize,
    samples: usize,
    json_path: Option<PathBuf>,
    fixture_out: Option<PathBuf>,
}

fn usage() -> &'static str {
    "usage: perf_cap_boundary --size <160|164|256> [--warmup N] [--samples N] [--json PATH] [--fixture-out PATH]"
}

fn parse_size(value: &str) -> AnyResult<usize> {
    match value {
        "160" | "164" | "256" => Ok(value.parse()?),
        other => Err(format!("unknown cap-boundary size '{other}'").into()),
    }
}

fn parse_count(value: &str, label: &str, maximum: usize) -> AnyResult<usize> {
    let count = value
        .parse::<usize>()
        .map_err(|source| format!("invalid {label} '{value}': {source}"))?;
    if count > maximum {
        return Err(format!("{label} {count} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn parse_args<I>(arguments: I) -> AnyResult<Arguments>
where
    I: IntoIterator<Item = OsString>,
{
    let mut size = None;
    let mut warmup = 10;
    let mut samples = 100;
    let mut json_path = None;
    let mut fixture_out = None;
    let mut args = arguments
        .into_iter()
        .map(|argument| argument.to_string_lossy().into_owned());

    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--size" => {
                let value = args.next().ok_or("missing value for --size")?;
                size = Some(parse_size(&value)?);
            },
            "--warmup" => {
                let value = args.next().ok_or("missing value for --warmup")?;
                warmup = parse_count(&value, "warmup", MAX_WARMUP)?;
            },
            "--samples" => {
                let value = args.next().ok_or("missing value for --samples")?;
                samples = parse_count(&value, "samples", MAX_SAMPLES)?;
            },
            "--json" => {
                let value = args.next().ok_or("missing value for --json")?;
                json_path = Some(PathBuf::from(value));
            },
            "--fixture-out" => {
                let value = args.next().ok_or("missing value for --fixture-out")?;
                fixture_out = Some(PathBuf::from(value));
            },
            "--help" | "-h" => return Err(usage().into()),
            other => return Err(format!("unknown argument '{other}'\n{}", usage()).into()),
        }
    }

    let size = size.ok_or_else(|| format!("missing --size\n{}", usage()))?;
    if samples == 0 {
        return Err("samples must be positive".into());
    }
    Ok(Arguments {
        size,
        warmup,
        samples,
        json_path,
        fixture_out,
    })
}

#[derive(Debug)]
struct Fixture {
    size: usize,
    cells: usize,
    last_cell: String,
    worksheet: Vec<u8>,
    archive: Vec<u8>,
    event_count: usize,
    expected_event_count: usize,
}

fn column_name(mut column: usize) -> String {
    let mut output = String::new();
    column += 1;
    while column != 0 {
        let remainder = ((column - 1) % 26) as u8;
        output.push(char::from(b'A' + remainder));
        column = (column - 1) / 26;
    }
    output.chars().rev().collect()
}

fn expected_event_count(size: usize) -> AnyResult<usize> {
    let cells = size
        .checked_mul(size)
        .ok_or("cap-boundary cell count overflow")?;
    cells
        .checked_mul(5)
        .and_then(|count| count.checked_add(size.checked_mul(2)?))
        .and_then(|count| count.checked_add(6))
        .ok_or_else(|| "cap-boundary event count overflow".into())
}

fn worksheet_xml(size: usize) -> AnyResult<Vec<u8>> {
    let cells = size
        .checked_mul(size)
        .ok_or("cap-boundary cell count overflow")?;
    let last_cell = format!("{}{}", column_name(size - 1), size);
    let capacity = 1_024usize
        .checked_add(
            cells
                .checked_mul(40)
                .ok_or("cap-boundary XML capacity overflow")?,
        )
        .ok_or("cap-boundary XML capacity overflow")?;
    let mut xml = String::new();
    xml.try_reserve(capacity)
        .map_err(|source| format!("cap-boundary XML reservation failed: {source}"))?;
    write!(
        &mut xml,
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:{last_cell}"/><sheetData>"#
    )?;
    for row in 0..size {
        let row_number = row + 1;
        write!(&mut xml, r#"<row r="{row_number}">"#)?;
        for column in 0..size {
            let address = format!("{}{}", column_name(column), row_number);
            let ordinal = row
                .checked_mul(size)
                .and_then(|value| value.checked_add(column))
                .and_then(|value| value.checked_add(1))
                .ok_or("cap-boundary cell ordinal overflow")?;
            write!(&mut xml, r#"<c r="{address}"><v>{ordinal}</v></c>"#)?;
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    let bytes = xml.into_bytes();
    if bytes.len() >= SHARED_SOURCE_BYTES {
        return Err(format!(
            "cap-boundary worksheet is {} bytes; expected < {}",
            bytes.len(),
            SHARED_SOURCE_BYTES
        )
        .into());
    }
    Ok(bytes)
}

fn archive_for(worksheet: &[u8]) -> AnyResult<Vec<u8>> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{OPC_REL}"><Relationship Id="rIdRoot" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{OPC_REL}"><Relationship Id="rIdSheet" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", content_types.as_bytes())?;
    writer.write_stored("_rels/.rels", root_relationships.as_bytes())?;
    writer.write_stored("xl/workbook.xml", workbook.as_bytes())?;
    writer.write_stored(
        "xl/_rels/workbook.xml.rels",
        workbook_relationships.as_bytes(),
    )?;
    writer.write_stored(WORKSHEET_PART, worksheet)?;
    let archive = writer.finish_to_bytes()?;
    if archive.len() >= SHARED_SOURCE_BYTES {
        return Err(format!(
            "cap-boundary archive is {} bytes; expected < {}",
            archive.len(),
            SHARED_SOURCE_BYTES
        )
        .into());
    }
    Ok(archive)
}

fn marker_free(worksheet: &[u8]) -> bool {
    [
        MCE_NAMESPACE,
        X14AC_NAMESPACE,
        ALTERNATE_CONTENT,
        DY_DESCENT,
    ]
    .into_iter()
    .all(|marker| {
        !worksheet
            .windows(marker.len())
            .any(|window| window == marker)
    })
}

fn count_xml_events(worksheet: &[u8]) -> AnyResult<usize> {
    let mut reader = Reader::from_reader(worksheet);
    let mut buffer = Vec::new();
    let mut count = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer)?;
        let eof = matches!(event, Event::Eof);
        count = count
            .checked_add(1)
            .ok_or("cap-boundary event count overflow while reading XML")?;
        buffer.clear();
        if eof {
            return Ok(count);
        }
    }
}

fn fixture(size: usize) -> AnyResult<Fixture> {
    let cells = size
        .checked_mul(size)
        .ok_or("cap-boundary cell count overflow")?;
    let worksheet = worksheet_xml(size)?;
    if std::str::from_utf8(&worksheet).is_err() {
        return Err("cap-boundary worksheet is not UTF-8".into());
    }
    if !marker_free(&worksheet) {
        return Err("cap-boundary worksheet unexpectedly contains MCE/extension markers".into());
    }
    let expected_event_count = expected_event_count(size)?;
    let event_count = count_xml_events(&worksheet)?;
    if event_count != expected_event_count {
        return Err(format!(
            "cap-boundary event count mismatch: expected {expected_event_count}, observed {event_count}"
        )
        .into());
    }
    if event_count > RAW_EVENT_CAP {
        return Err(format!(
            "cap-boundary fixture exceeds ordinary parser event cap {RAW_EVENT_CAP}"
        )
        .into());
    }
    let archive = archive_for(&worksheet)?;
    Ok(Fixture {
        size,
        cells,
        last_cell: format!("{}{}", column_name(size - 1), size),
        worksheet,
        archive,
        event_count,
        expected_event_count,
    })
}

fn verify_noop_publication(fixture: &Fixture) -> AnyResult<()> {
    let source = Arc::new(OwnedSource::new(fixture.archive.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone())?;
    let commit = editor.edit_sheets([Selector::from("Sheet1")])?.commit()?;
    if commit.changed() || !commit.patch().is_empty() {
        return Err("cap-boundary empty commit changed the source".into());
    }
    let mut published = Vec::new();
    editor.publish_multi_commit_to_stream(&mut published, &commit)?;
    if published != fixture.archive {
        return Err("cap-boundary no-op publication changed archive bytes".into());
    }
    if source.as_slice() != fixture.archive.as_slice() {
        return Err("cap-boundary no-op publication changed source bytes".into());
    }
    Ok(())
}

fn run_iteration(fixture: &Fixture) -> AnyResult<u128> {
    let source = Arc::new(OwnedSource::new(fixture.archive.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone())?;
    let selectors = [Selector::from("Sheet1")];

    let started = Instant::now();
    let planned = editor.edit_sheets(selectors.into_iter());
    // `planned` remains alive beyond this statement and is inspected only
    // after the measured interval has closed.
    let duration_ns = started.elapsed().as_nanos();

    let transaction = planned?;
    if transaction.worksheet_count() != 1 {
        return Err("cap-boundary selected an unexpected worksheet count".into());
    }
    let a1 = Address::from_a1("A1")?;
    let last = Address::from_a1(&fixture.last_cell)?;
    let first_expected = Value::Number(Number::new("1")?);
    let last_expected = Value::Number(Number::new(fixture.cells.to_string())?);
    if transaction.before().value(0, a1) != Some(&first_expected)
        || transaction.before().value(0, last) != Some(&last_expected)
    {
        return Err("cap-boundary source snapshot lost A1 or the last cell".into());
    }

    let commit = transaction.commit()?;
    if commit.changed() || !commit.patch().is_empty() {
        return Err("cap-boundary empty commit was not an exact no-op".into());
    }
    if commit.snapshot().value(0, a1) != Some(&first_expected)
        || commit.snapshot().value(0, last) != Some(&last_expected)
    {
        return Err("cap-boundary committed snapshot lost A1 or the last cell".into());
    }
    if source.as_slice() != fixture.archive.as_slice() {
        return Err("cap-boundary iteration changed source bytes".into());
    }
    Ok(duration_ns)
}

fn report(
    arguments: &Arguments,
    fixture: &Fixture,
    durations: &[u128],
    fixture_dump: Option<&PathBuf>,
) -> JsonValue {
    let relation = if fixture.event_count <= SHARED_EVENT_CAP {
        "below"
    } else {
        "above"
    };
    let sample_order: Vec<usize> = (0..durations.len()).collect();
    let fixture_dump = fixture_dump.map(|path| path.to_string_lossy().into_owned());
    let binary_path = env::current_exe()
        .ok()
        .map(|path| path.to_string_lossy().into_owned());
    let binary_bytes = env::current_exe()
        .ok()
        .and_then(|path| fs::metadata(path).ok())
        .map(|metadata| metadata.len());
    json!({
        "schema": "litchi.xlsx.cap-boundary-guard.v1",
        "tool": "perf_cap_boundary",
        "case": "valid",
        "size": fixture.size,
        "warmup_iterations": arguments.warmup,
        "samples": arguments.samples,
        "rows": fixture.size,
        "columns": fixture.size,
        "cells": fixture.cells,
        "last_cell": fixture.last_cell,
        "event_count": fixture.event_count,
        "expected_event_count": fixture.expected_event_count,
        "event_count_formula": "5*N*N+2*N+6 (including EOF)",
        "shared_provisional_event_cap": SHARED_EVENT_CAP,
        "ordinary_parser_event_cap": RAW_EVENT_CAP,
        "event_cap_relation": relation,
        "source_stream_byte_limit": SHARED_SOURCE_BYTES,
        "source_stream_eligible": true,
        "archive_bytes": fixture.archive.len(),
        "source_xml_bytes": fixture.worksheet.len(),
        "fixture_out": fixture_dump.clone(),
        "source": {
            "shape": format!("{}x{}", fixture.size, fixture.size),
            "rows": fixture.size,
            "columns": fixture.size,
            "format": "OOXML/XLSX",
            "fixture_kind": "single_worksheet_dense_numeric_grid",
            "generator": "litchi-xlsx-cap-boundary-stored-grid-v1",
            "worksheet_member": WORKSHEET_PART,
            "compression": "stored",
            "encoding": "UTF-8",
            "marker_free": true,
            "source_xml_bytes": fixture.worksheet.len(),
            "worksheet_bytes": fixture.worksheet.len(),
            "worksheet_xml_bytes": fixture.worksheet.len(),
            "archive_bytes": fixture.archive.len(),
            "fixture_dump": fixture_dump,
            "identity_method": "parent binds SHA-256 of fixture_dump bytes",
        },
        "binary": {
            "path": binary_path,
            "bytes": binary_bytes,
            "profile": "release",
            "identity_method": "parent binds SHA-256 of the captured binary",
        },
        "phase": {
            "name": "edit_sheets",
            "timing_scope": "source/editor/selector setup excluded; transaction retained through clock stop; inspection/commit/drop excluded",
            "warmup_iterations": arguments.warmup,
            "samples": arguments.samples,
            "sample_order": sample_order,
            "duration_ns": durations,
        },
        "correctness": {
            "source_unchanged": true,
            "source_bytes_unchanged": true,
            "source_xml_unchanged": true,
            "valid_snapshot_values": true,
            "snapshot_a1": 1,
            "snapshot_last_cell": fixture.cells,
            "empty_commit_is_noop": true,
            "commit_outside_timing": true,
            "no_op_publication_exact": true,
        },
        "performance_claim": "none: valid event-cap boundary diagnostic only",
    })
}

fn run_from_args<I>(arguments: I) -> AnyResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let arguments = parse_args(arguments)?;
    let fixture = fixture(arguments.size)?;
    if let Some(path) = arguments.fixture_out.as_ref() {
        fs::write(path, &fixture.archive)?;
    }
    // Keep this expensive publication oracle outside every timing sample. It
    // still proves the exact archive identity for each fresh executable.
    verify_noop_publication(&fixture)?;

    for _ in 0..arguments.warmup {
        let _ = run_iteration(&fixture)?;
    }
    let mut durations = Vec::new();
    durations.try_reserve_exact(arguments.samples)?;
    for _ in 0..arguments.samples {
        durations.push(run_iteration(&fixture)?);
    }
    if durations.len() != arguments.samples {
        return Err("cap-boundary produced an incomplete sample vector".into());
    }

    let encoded = serde_json::to_vec_pretty(&report(
        &arguments,
        &fixture,
        &durations,
        arguments.fixture_out.as_ref(),
    ))?;
    if let Some(path) = arguments.json_path {
        fs::write(path, &encoded)?;
    }
    println!("{}", String::from_utf8(encoded)?);
    Ok(())
}

fn main() -> AnyResult<()> {
    run_from_args(env::args_os().skip(1))
}
