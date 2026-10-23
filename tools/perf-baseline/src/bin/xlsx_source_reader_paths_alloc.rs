//! Count-only allocation evidence for the ordinary source reader paths.
//!
//! Each route opens a fresh owner over the same fixture, then exercises one
//! cold operation: a cell read, a cell record plus format lookup, or a cell
//! cursor. The fixture and route observations are checked in-process; the
//! JSON is paired by the coordinator against the pre-W4 owner and the W4
//! owner. This is allocator evidence, not a latency or RSS measurement.

#[path = "support/counting_allocator.rs"]
mod allocator;

use std::error::Error;
use std::hint::black_box;
use std::io::Write;
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use serde::Serialize;
use sha2::{Digest, Sha256};

use litchi_perf_baseline::allocation_metrics;

type AnyError = Box<dyn Error>;
type AnyResult<T> = Result<T, AnyError>;

const SCHEMA: &str = "litchi.xlsx.source-reader-paths-alloc.v1";
const MAIN_PART: &str = "/xl/workbook.xml";
const SHEET_PART: &str = "/xl/worksheets/sheet1.xml";
const STYLES_PART: &str = "/xl/styles.xml";

const API_VARIANT: &str = "w4-source-reader-lazy-cache";

#[derive(Debug, Clone, Copy)]
struct Arguments {
    iterations: usize,
    warmups: usize,
    print_fixture_identity: bool,
    open_only: bool,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api_variant: &'static str,
    iterations: usize,
    warmups: usize,
    fixture_sha256: String,
    fixture_bytes: usize,
    cell_observations: usize,
    style_observations: usize,
    cursor_observations: usize,
    routes: Routes,
}

#[derive(Debug, Serialize)]
struct Routes {
    first_cell_read: AllocationDelta,
    first_style_read: AllocationDelta,
    first_cursor_read: AllocationDelta,
}

#[derive(Debug, Serialize)]
struct AllocationDelta {
    status: &'static str,
    scope: &'static str,
    operation_count: usize,
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes_before: u64,
    live_bytes_after: u64,
    peak_live_bytes_before: u64,
    peak_live_bytes_after: u64,
}

fn main() -> AnyResult<()> {
    allocation_metrics::enable();
    let arguments = parse_arguments(std::env::args().skip(1))?;
    let fixture = fixture()?;
    let fixture_sha256 = sha256_hex(&fixture);
    if arguments.print_fixture_identity {
        eprintln!(
            "fixture_sha256={} fixture_bytes={}",
            fixture_sha256,
            fixture.len()
        );
    }

    if arguments.open_only {
        let route = run_open_route(arguments.iterations, arguments.warmups, &fixture)?;
        let stdout = std::io::stdout();
        let mut output = stdout.lock();
        serde_json::to_writer_pretty(&mut output, &route)?;
        output.write_all(b"\n")?;
        return Ok(());
    }

    let (first_cell_read, cell_observations) = run_route(
        arguments.iterations,
        arguments.warmups,
        &fixture,
        "first_cell_read",
        |workbook| {
            let sheet = workbook
                .sheet("Sheet1")?
                .ok_or("reader fixture did not expose Sheet1")?;
            let value = sheet.cell("A1")?;
            let observed = usize::from(!matches!(value, litchi_xlsx::SourceCellView::Missing));
            black_box(&value);
            Ok(observed)
        },
    )?;
    let (first_style_read, style_observations) = run_route(
        arguments.iterations,
        arguments.warmups,
        &fixture,
        "first_style_read",
        |workbook| {
            let sheet = workbook
                .sheet("Sheet1")?
                .ok_or("reader fixture did not expose Sheet1")?;
            let record = sheet
                .cell_record("A1")?
                .ok_or("reader fixture did not expose A1")?;
            let format = sheet.format(record.style.as_ref())?;
            let observed = usize::from(format.is_some() && record.style.is_some());
            black_box((&record, &format));
            Ok(observed)
        },
    )?;
    let (first_cursor_read, cursor_observations) = run_route(
        arguments.iterations,
        arguments.warmups,
        &fixture,
        "first_cursor_read",
        |workbook| {
            let sheet = workbook
                .sheet("Sheet1")?
                .ok_or("reader fixture did not expose Sheet1")?;
            let mut cursor = sheet.cell_cursor()?;
            let mut count = 0usize;
            while let Some(record) = cursor.next() {
                let (_address, _index) = record?;
                count = count.checked_add(1).ok_or("cursor count overflow")?;
            }
            if count != 3 {
                return Err(format!("reader fixture cursor returned {count} records").into());
            }
            black_box(cursor.is_complete());
            Ok(count)
        },
    )?;

    let report = Report {
        schema: SCHEMA,
        api_variant: API_VARIANT,
        iterations: arguments.iterations,
        warmups: arguments.warmups,
        fixture_sha256,
        fixture_bytes: fixture.len(),
        cell_observations,
        style_observations,
        cursor_observations,
        routes: Routes {
            first_cell_read,
            first_style_read,
            first_cursor_read,
        },
    };
    let stdout = std::io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer_pretty(&mut output, &report)?;
    output.write_all(b"\n")?;
    Ok(())
}

fn run_route<F>(
    iterations: usize,
    warmups: usize,
    fixture: &[u8],
    scope: &'static str,
    mut operation: F,
) -> AnyResult<(AllocationDelta, usize)>
where
    F: FnMut(&litchi_xlsx::SourceBackedWorkbook) -> AnyResult<usize>,
{
    for _ in 0..warmups {
        let owner = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            fixture.to_vec(),
        )))?;
        black_box(operation(&owner)?);
    }
    let before = allocation_metrics::snapshot();
    let mut observations = 0usize;
    for _ in 0..iterations {
        let owner = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            fixture.to_vec(),
        )))?;
        observations = observations
            .checked_add(operation(&owner)?)
            .ok_or("reader observation count overflow")?;
        black_box(&owner);
    }
    let after = allocation_metrics::snapshot();
    Ok((
        allocation_delta(before, after, scope, iterations)?,
        observations,
    ))
}

fn run_open_route(iterations: usize, warmups: usize, fixture: &[u8]) -> AnyResult<AllocationDelta> {
    for _ in 0..warmups {
        let owner = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            fixture.to_vec(),
        )))?;
        black_box(&owner);
    }
    let before = allocation_metrics::snapshot();
    for _ in 0..iterations {
        let owner = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            fixture.to_vec(),
        )))?;
        black_box(&owner);
    }
    let after = allocation_metrics::snapshot();
    allocation_delta(before, after, "open_only", iterations)
}

fn fixture() -> AnyResult<Vec<u8>> {
    let workbook = br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><definedNames><definedName name="Rate">Sheet1!$A$1</definedName></definedNames><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#;
    let sheet = br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>hello</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="C2"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>"#;
    let styles = br#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/><xf numFmtId="2"/></cellXfs></styleSheet>"#;
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (
            MAIN_PART,
            litchi_opc::constants::content_type::SML_SHEET_MAIN,
            workbook.to_vec(),
        ),
        (
            SHEET_PART,
            litchi_opc::constants::content_type::SML_WORKSHEET,
            sheet.to_vec(),
        ),
        (
            STYLES_PART,
            litchi_opc::constants::content_type::SML_STYLES,
            styles.to_vec(),
        ),
    ] {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(name)?,
            content_type.to_owned(),
            bytes,
        )))?;
    }
    let main = PackURI::new(MAIN_PART)?;
    let workbook_part = package.get_part_mut(&main)?;
    workbook_part.rels_mut().try_add_relationship(
        litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
        "worksheets/sheet1.xml".to_owned(),
        "rIdSheet".to_owned(),
        TargetMode::Internal,
    )?;
    workbook_part.rels_mut().try_add_relationship(
        litchi_opc::constants::relationship_type::STYLES.to_owned(),
        "styles.xml".to_owned(),
        "rIdStyles".to_owned(),
        TargetMode::Internal,
    )?;
    package.relate_to(
        "xl/workbook.xml",
        litchi_opc::constants::relationship_type::OFFICE_DOCUMENT,
    );
    Ok(PackageWriter::to_bytes(&package)?)
}

fn allocation_delta(
    before: allocation_metrics::Snapshot,
    after: allocation_metrics::Snapshot,
    scope: &'static str,
    operation_count: usize,
) -> AnyResult<AllocationDelta> {
    if after.observer_invalid || after.overflowed {
        return Err("allocator counters became invalid during source reader paths".into());
    }
    let subtract = |after: u64, before: u64, label: &str| -> AnyResult<u64> {
        after
            .checked_sub(before)
            .ok_or_else(|| format!("allocator {label} counter regressed").into())
    };
    Ok(AllocationDelta {
        status: "measured-count-only",
        scope,
        operation_count,
        allocation_calls: subtract(
            after.allocation_calls,
            before.allocation_calls,
            "allocation",
        )?,
        deallocation_calls: subtract(
            after.deallocation_calls,
            before.deallocation_calls,
            "deallocation",
        )?,
        reallocation_calls: subtract(
            after.reallocation_calls,
            before.reallocation_calls,
            "reallocation",
        )?,
        allocated_bytes: subtract(after.allocated_bytes, before.allocated_bytes, "allocated")?,
        deallocated_bytes: subtract(
            after.deallocated_bytes,
            before.deallocated_bytes,
            "deallocated",
        )?,
        live_bytes_before: before.live_bytes,
        live_bytes_after: after.live_bytes,
        peak_live_bytes_before: before.peak_live_bytes,
        peak_live_bytes_after: after.peak_live_bytes,
    })
}

fn parse_arguments<I>(arguments: I) -> AnyResult<Arguments>
where
    I: IntoIterator<Item = String>,
{
    let mut parsed = Arguments {
        iterations: 64,
        warmups: 2,
        print_fixture_identity: false,
        open_only: false,
    };
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let (name, value) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(name, value)| {
                (name, Some(value))
            });
        match name {
            "--iterations" => {
                let value = value
                    .map(str::to_owned)
                    .or_else(|| arguments.next())
                    .ok_or("--iterations requires a value")?;
                parsed.iterations = value.parse()?;
            },
            "--warmups" => {
                let value = value
                    .map(str::to_owned)
                    .or_else(|| arguments.next())
                    .ok_or("--warmups requires a value")?;
                parsed.warmups = value.parse()?;
            },
            "--print-fixture-identity" => parsed.print_fixture_identity = true,
            "--open-only" => parsed.open_only = true,
            unknown => return Err(format!("unknown argument `{unknown}`").into()),
        }
    }
    if parsed.iterations == 0 {
        return Err("--iterations must be greater than zero".into());
    }
    Ok(parsed)
}

fn sha256_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push(HEX[usize::from(byte >> 4)] as char);
        output.push(HEX[usize::from(byte & 0x0f)] as char);
    }
    output
}
