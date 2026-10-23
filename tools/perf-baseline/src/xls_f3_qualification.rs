//! Native XLS source-backed cell-reopen qualification (memo 0359a-F3).
//!
//! This probe is deliberately separate from the normal selector registry. It
//! keeps the existing numeric corpus builders and correctness gates, while
//! exposing the three lifecycle boundaries needed to qualify the deferred
//! composed-workbook reopen:
//!
//! ```text
//! write-only       commit -> write_to
//! first-read       commit -> first semantic read
//! read-after-write commit -> write_to -> first semantic read
//! ```
//!
//! Source opening, corpus setup, expected output, and all replay/refusal
//! checks are outside the route clock. Each measured operation opens a fresh
//! source owner and drops its committed owner after the checks. The example is
//! copied unchanged into the pre-reopen source snapshot for the paired run.

use std::error::Error;
use std::hint::black_box;
use std::io::{self, Cursor, Write};
use std::sync::Arc;
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::Position;
use litchi_core::sheet::CellValue;
use litchi_xls::SourceBackedWorkbook;
use litchi_xls::cell_values::{Reference, Selector, Storage, Value};
use litchi_xls::source_backed::cells;
use sha2::{Digest, Sha256};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const MAX_WRITE: usize = 64 * 1024;
const REAL_PRODUCER: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/spreadsheet/54016.xls");

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Route {
    WriteOnly,
    FirstRead,
    ReadAfterWrite,
}

impl Route {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "write-only" => Ok(Self::WriteOnly),
            "first-read" => Ok(Self::FirstRead),
            "read-after-write" => Ok(Self::ReadAfterWrite),
            _ => Err(format!(
                "unknown XLS F3 route {value:?}; expected write-only, first-read, or read-after-write"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::WriteOnly => "write-only",
            Self::FirstRead => "first-read",
            Self::ReadAfterWrite => "read-after-write",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CorpusChoice {
    Number,
    RkMulrk,
    All,
}

impl CorpusChoice {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "number" => Ok(Self::Number),
            "rk-mulrk" => Ok(Self::RkMulrk),
            "all" => Ok(Self::All),
            _ => Err(
                format!("unknown XLS F3 corpus {value:?}; expected number, rk-mulrk, or all")
                    .into(),
            ),
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct EditSpec {
    selector: Selector<'static>,
    sheet_name: &'static str,
    sheet_index: usize,
    reference: Reference,
    replacement: f64,
    storage: Storage,
}

#[derive(Debug)]
struct Fixture {
    name: &'static str,
    archive: Vec<u8>,
    edits: Vec<EditSpec>,
}

#[derive(Debug)]
struct Expected {
    output: Vec<u8>,
    digest: String,
}

#[derive(Debug)]
struct Measurement {
    elapsed_ns: u64,
    edit_ns: u64,
    set_ns: u64,
    commit_ns: u64,
    write_ns: u64,
    read_ns: u64,
    output: Vec<u8>,
    source: SourceBackedWorkbook,
    commit: cells::Commit,
}

#[derive(Debug)]
struct CountingSink {
    bytes: Vec<u8>,
    write_calls: u64,
    max_write: usize,
}

impl CountingSink {
    fn new(capacity: usize) -> Result<Self> {
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(capacity)?;
        Ok(Self {
            bytes,
            write_calls: 0,
            max_write: MAX_WRITE,
        })
    }
}

impl Write for CountingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        let accepted = bytes.len().min(self.max_write);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.write_calls = self.write_calls.saturating_add(1);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct PrefixSink {
    accepted: usize,
    limit: usize,
}

impl Write for PrefixSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("intentional XLS F3 sink failure"));
        }
        let accepted = bytes.len().min(self.limit - self.accepted);
        self.accepted = self.accepted.saturating_add(accepted);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64> {
    u64::try_from(duration.as_nanos()).map_err(|_| "XLS F3 route duration overflows u64".into())
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn source_workbook(bytes: &[u8]) -> Result<SourceBackedWorkbook> {
    Ok(SourceBackedWorkbook::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(bytes.to_vec()),
    ))?)
}

fn cell_snapshot(bytes: &[u8]) -> Result<litchi_xls::cell_values::Snapshot> {
    Ok(litchi_xls::Package::from_bytes(bytes)?.cell_values()?)
}

fn write_package(package: &SourceBackedWorkbook) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn write_commit(commit: &cells::Commit, capacity: usize) -> Result<CountingSink> {
    let mut sink = CountingSink::new(capacity)?;
    commit.package().write_to(&mut sink)?;
    if sink.bytes.len() != capacity {
        return Err(format!(
            "XLS F3 publication wrote {} bytes, expected {capacity}",
            sink.bytes.len()
        )
        .into());
    }
    let max_calls = capacity
        .checked_add(MAX_WRITE - 1)
        .ok_or("XLS F3 publication write-call bound overflows")?
        / MAX_WRITE;
    if sink.write_calls == 0 || sink.write_calls > u64::try_from(max_calls)? {
        return Err(format!(
            "XLS F3 publication used {} writes, expected at most {max_calls}",
            sink.write_calls
        )
        .into());
    }
    Ok(sink)
}

fn stage(edit: &mut cells::Edit, fixture: &Fixture) -> Result<()> {
    for specification in &fixture.edits {
        match specification.storage {
            Storage::Number => edit.set_number(
                specification.selector,
                specification.reference,
                specification.replacement,
            )?,
            Storage::Rk | Storage::MulRk => edit.set_numeric(
                specification.selector,
                specification.reference,
                specification.replacement,
            )?,
            _ => return Err("XLS F3 fixture selected an unsupported storage family".into()),
        }
    }
    Ok(())
}

fn read_first_target(package: &SourceBackedWorkbook, fixture: &Fixture) -> Result<()> {
    let specification = fixture
        .edits
        .first()
        .ok_or("XLS F3 fixture has no selected cell")?;
    let value = package
        .cell_value_by_index(
            specification.sheet_index,
            u32::from(specification.reference.row()),
            u32::from(specification.reference.column()),
        )?
        .ok_or("XLS F3 first read lost the selected cell")?;
    let (CellValue::Float(value) | CellValue::DateTime(value)) = value else {
        return Err("XLS F3 first read returned a non-numeric value".into());
    };
    if value.to_bits() != specification.replacement.to_bits() {
        return Err("XLS F3 first read returned the wrong numeric value".into());
    }
    Ok(())
}

fn verify_committed_values(package: &SourceBackedWorkbook, fixture: &Fixture) -> Result<()> {
    for specification in &fixture.edits {
        let value = package
            .cell_value_by_index(
                specification.sheet_index,
                u32::from(specification.reference.row()),
                u32::from(specification.reference.column()),
            )?
            .ok_or("XLS F3 committed workbook lost an edited cell")?;
        let (CellValue::Float(value) | CellValue::DateTime(value)) = value else {
            return Err("XLS F3 committed workbook returned a non-numeric value".into());
        };
        if value.to_bits() != specification.replacement.to_bits() {
            return Err("XLS F3 committed workbook value differs from expectation".into());
        }
    }
    // The second lookup of the first target is deliberately outside the clock:
    // candidate materialization must be cached after the route's first read.
    read_first_target(package, fixture)
}

fn commit_fixture(source: &SourceBackedWorkbook, fixture: &Fixture) -> Result<cells::Commit> {
    let mut edit = source.edit_cells()?;
    stage(&mut edit, fixture)?;
    Ok(edit.commit()?)
}

fn run_once(route: Route, fixture: &Fixture, output_capacity: usize) -> Result<Measurement> {
    // Source opening is intentionally outside every route clock, matching the
    // existing xls_numeric selectors. Each call receives fresh source bytes.
    let source = source_workbook(&fixture.archive)?;
    let route_started = Instant::now();

    let edit_started = Instant::now();
    let mut edit = source.edit_cells()?;
    let edit_ns = elapsed_ns(edit_started.elapsed())?;

    let set_started = Instant::now();
    stage(&mut edit, fixture)?;
    let set_ns = elapsed_ns(set_started.elapsed())?;

    let commit_started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed_ns(commit_started.elapsed())?;

    let mut output = Vec::new();
    let mut write_ns = 0;
    let mut read_ns = 0;
    match route {
        Route::WriteOnly => {
            let write_started = Instant::now();
            let sink = write_commit(&commit, output_capacity)?;
            write_ns = elapsed_ns(write_started.elapsed())?;
            output = sink.bytes;
        },
        Route::FirstRead => {
            let read_started = Instant::now();
            read_first_target(commit.package(), fixture)?;
            read_ns = elapsed_ns(read_started.elapsed())?;
            // Publication is intentionally after the first-read clock. This
            // route isolates the moved target reopen from output streaming.
            output = write_commit(&commit, output_capacity)?.bytes;
        },
        Route::ReadAfterWrite => {
            let write_started = Instant::now();
            let sink = write_commit(&commit, output_capacity)?;
            write_ns = elapsed_ns(write_started.elapsed())?;
            output = sink.bytes;
            let read_started = Instant::now();
            read_first_target(commit.package(), fixture)?;
            read_ns = elapsed_ns(read_started.elapsed())?;
        },
    }
    let elapsed_ns = elapsed_ns(route_started.elapsed())?;

    Ok(Measurement {
        elapsed_ns,
        edit_ns,
        set_ns,
        commit_ns,
        write_ns,
        read_ns,
        output,
        source,
        commit,
    })
}

fn prepare_expected(fixture: &Fixture) -> Result<Expected> {
    let source = source_workbook(&fixture.archive)?;
    let commit = commit_fixture(&source, fixture)?;
    let output = write_package(commit.package())?;
    if output == fixture.archive {
        return Err(format!("XLS F3 {} edit did not change its source", fixture.name).into());
    }
    verify_output(fixture, &output)?;
    verify_patch(&source, &commit)?;
    verify_partial_publication(&commit, output.len())?;
    Ok(Expected {
        digest: sha256_hex(&output),
        output,
    })
}

fn verify_sample(
    route: Route,
    fixture: &Fixture,
    expected: &Expected,
    measurement: Measurement,
) -> Result<Measurement> {
    if measurement.output != expected.output || sha256_hex(&measurement.output) != expected.digest {
        return Err(format!(
            "XLS F3 {} {} output differs from preflight",
            fixture.name,
            route.name()
        )
        .into());
    }
    verify_output(fixture, &measurement.output)?;
    verify_patch(&measurement.source, &measurement.commit)?;
    if matches!(route, Route::FirstRead | Route::ReadAfterWrite) {
        verify_committed_values(measurement.commit.package(), fixture)?;
    }
    black_box((&measurement.output, &measurement.commit));
    Ok(measurement)
}

fn verify_partial_publication(commit: &cells::Commit, output_bytes: usize) -> Result<()> {
    let mut sink = PrefixSink {
        accepted: 0,
        limit: output_bytes / 2,
    };
    if commit.package().write_to(&mut sink).is_ok()
        || sink.accepted == 0
        || sink.accepted >= output_bytes
    {
        return Err("XLS F3 partial sink did not produce a typed prefix failure".into());
    }
    Ok(())
}

fn verify_patch(source: &SourceBackedWorkbook, commit: &cells::Commit) -> Result<()> {
    let patch = commit.patch();
    let applied = source.apply_cells(patch)?;
    let applied_bytes = write_package(applied.package())?;
    let committed_bytes = write_package(commit.package())?;
    if applied_bytes != committed_bytes || commit.package().apply_cells(patch).is_ok() {
        return Err("XLS F3 patch replay or stale-recipient gate failed".into());
    }
    let restored = commit.package().apply_cells(&patch.inverse())?;
    let restored_bytes = write_package(restored.package())?;
    let source_bytes = write_package(source)?;
    if restored_bytes != source_bytes {
        return Err("XLS F3 inverse patch did not restore source bytes".into());
    }
    Ok(())
}

fn verify_output(fixture: &Fixture, output: &[u8]) -> Result<()> {
    verify_untouched_streams(&fixture.archive, output)?;
    let snapshot = cell_snapshot(output)?;
    for specification in &fixture.edits {
        let sheet = snapshot
            .worksheet(specification.selector)?
            .ok_or("XLS F3 output lost the selected worksheet")?;
        let cell = sheet
            .cell(specification.reference)?
            .ok_or("XLS F3 output lost an edited cell")?;
        let Value::Number(value) = cell.value() else {
            return Err("XLS F3 output edited cell is no longer numeric".into());
        };
        if value.to_bits() != specification.replacement.to_bits()
            || cell.storage() != specification.storage
        {
            return Err("XLS F3 output value or storage family differs".into());
        }
    }

    // The independent raw reader must agree with the semantic snapshot.
    let reader = litchi_xls::raw::Reader::new(Cursor::new(output))?;
    for specification in &fixture.edits {
        let sheet = reader
            .sheets()
            .iter()
            .find(|sheet| sheet.name() == specification.sheet_name)
            .ok_or("XLS F3 raw reader lost the selected worksheet")?;
        let worksheet = reader.xls_worksheet(
            sheet
                .parsed_worksheet_index()
                .ok_or("XLS F3 raw reader selected a non-worksheet tab")?,
        )?;
        let value = worksheet
            .get_cell(
                u32::from(specification.reference.row()),
                u32::from(specification.reference.column()),
            )
            .ok_or("XLS F3 raw reader lost an edited cell")?
            .value()
            .as_float()
            .ok_or("XLS F3 raw reader returned a non-numeric cell")?;
        if value.to_bits() != specification.replacement.to_bits() {
            return Err("XLS F3 raw reader value differs from expectation".into());
        }
    }
    Ok(())
}

fn verify_untouched_streams(source: &[u8], candidate: &[u8]) -> Result<()> {
    let mut source_ole = OleFile::open(Cursor::new(source))?;
    let mut candidate_ole = OleFile::open(Cursor::new(candidate))?;
    let mut source_paths = source_ole.list_streams();
    let mut candidate_paths = candidate_ole.list_streams();
    source_paths.sort();
    candidate_paths.sort();
    if source_paths != candidate_paths
        || !source_ole.directory_exists(&["OpaquePayloads"])
        || !candidate_ole.directory_exists(&["OpaquePayloads"])
    {
        return Err("XLS F3 publication changed the CFB stream inventory".into());
    }
    for path in source_paths {
        let references = path.iter().map(String::as_str).collect::<Vec<_>>();
        let source_stream = source_ole.open_stream(&references)?;
        let candidate_stream = candidate_ole.open_stream(&references)?;
        let workbook = path.len() == 1
            && path
                .first()
                .is_some_and(|name| name == "Workbook" || name == "Book");
        if workbook {
            if source_stream == candidate_stream || source_stream.len() != candidate_stream.len() {
                return Err("XLS F3 Workbook stream was not changed with equal length".into());
            }
        } else if source_stream != candidate_stream {
            return Err("XLS F3 publication changed an untouched CFB stream".into());
        }
    }
    Ok(())
}

fn build_fixtures(choice: CorpusChoice) -> Result<Vec<Fixture>> {
    let mut fixtures = Vec::new();
    if matches!(choice, CorpusChoice::Number | CorpusChoice::All) {
        let archive = crate::xls_f3_number_fixture()?;
        let snapshot = cell_snapshot(&archive)?;
        let sheet = snapshot
            .worksheet(Selector::Name("Untouched"))?
            .ok_or("XLS F3 Number corpus lost Untouched worksheet")?;
        let reference = Reference::new(20, 4)?;
        let cell = sheet
            .cell(reference)?
            .ok_or("XLS F3 Number corpus lost Untouched!E21")?;
        if cell.storage() != Storage::Number
            || !matches!(cell.value(), Value::Number(value) if value.to_bits() == 42.0_f64.to_bits())
        {
            return Err("XLS F3 Number corpus target is not Number/42".into());
        }
        fixtures.push(Fixture {
            name: "number",
            archive,
            edits: vec![EditSpec {
                selector: Selector::Name("Untouched"),
                sheet_name: "Untouched",
                sheet_index: sheet.position(),
                reference,
                replacement: 43.0,
                storage: Storage::Number,
            }],
        });
    }
    if matches!(choice, CorpusChoice::RkMulrk | CorpusChoice::All) {
        let archive = crate::xls_numeric::build_rk_mulrk_fixture()?;
        let snapshot = cell_snapshot(&archive)?;
        let sheet = snapshot
            .worksheet(Selector::Name("Packed"))?
            .ok_or("XLS F3 RK/MulRK corpus lost Packed worksheet")?;
        let positions = [(0_u32, 0_u32, 2.0), (1, 0, 5.0), (1, 1, 6.0)];
        let mut edits = Vec::new();
        for (row, column, replacement) in positions {
            let reference = Reference::new(row, column)?;
            let cell = sheet
                .cell(reference)?
                .ok_or("XLS F3 RK/MulRK corpus lost an expected cell")?;
            if !matches!(cell.storage(), Storage::Rk | Storage::MulRk)
                || !matches!(cell.value(), Value::Number(_))
            {
                return Err("XLS F3 RK/MulRK corpus storage family changed".into());
            }
            edits.push(EditSpec {
                selector: Selector::Name("Packed"),
                sheet_name: "Packed",
                sheet_index: sheet.position(),
                reference,
                replacement,
                storage: cell.storage(),
            });
        }
        fixtures.push(Fixture {
            name: "rk-mulrk",
            archive,
            edits,
        });
    }
    if fixtures.is_empty() {
        return Err("XLS F3 selected no corpus".into());
    }
    Ok(fixtures)
}

fn verify_noop_and_unsupported(fixture: &Fixture) -> Result<()> {
    let source = source_workbook(&fixture.archive)?;
    let noop = source.edit_cells()?.commit()?;
    if noop.diagnostics().changed() || !noop.patch().is_noop() {
        return Err(format!("XLS F3 {} no-op changed its source", fixture.name).into());
    }
    if write_package(noop.package())? != fixture.archive {
        return Err(format!("XLS F3 {} no-op output differs", fixture.name).into());
    }

    let first = fixture
        .edits
        .first()
        .ok_or("XLS F3 fixture has no edit for refusal checks")?;
    let mut unsupported = source.edit_cells()?;
    if unsupported
        .set_value(
            first.selector,
            first.reference,
            Value::Text("unsupported numeric family".to_owned()),
        )
        .is_ok()
    {
        if unsupported.commit().is_ok() {
            return Err(format!("XLS F3 {} accepted a text cell replacement", fixture.name).into());
        }
    } else if unsupported.commit()?.diagnostics().changed() {
        return Err(format!(
            "XLS F3 {} unsupported edit was not failure-atomic",
            fixture.name
        )
        .into());
    }

    let mut wrong_family = source.edit_cells()?;
    if first.storage == Storage::Number {
        wrong_family.set_numeric(first.selector, first.reference, first.replacement)?;
    } else if wrong_family
        .set_number(first.selector, first.reference, first.replacement)
        .is_ok()
    {
        return Err(format!(
            "XLS F3 {} accepted set_number for compressed storage",
            fixture.name
        )
        .into());
    }
    Ok(())
}

fn copy_streams(source: &[u8]) -> Result<OleWriter> {
    let mut ole = OleFile::open(Cursor::new(source))?;
    let mut writer = OleWriter::new();
    let mut streams = ole.list_streams();
    streams.sort();
    for path in streams {
        let references = path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream_owned(&references, ole.open_stream(&references)?)?;
    }
    Ok(writer)
}

fn add_stream(source: &[u8], path: &[&str], payload: &[u8]) -> Result<Vec<u8>> {
    let mut writer = copy_streams(source)?;
    writer.create_stream(path, payload)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn macro_fixture(source: &[u8]) -> Result<Vec<u8>> {
    let mut writer = copy_streams(source)?;
    writer.create_storage(&["_VBA_PROJECT_CUR"])?;
    writer.create_storage(&["_VBA_PROJECT_CUR", "VBA"])?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "_VBA_PROJECT"], b"project")?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "dir"], b"dir")?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "Module1"], b"module")?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn protected_fixture() -> Result<Vec<u8>> {
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Protected")?;
    builder.write_number(sheet, 0, 0, 1.0)?;
    builder.protect_sheet(sheet, Some("password"), true, false)?;
    let mut output = Cursor::new(Vec::new());
    builder.build()?.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn encrypted_fixture() -> Result<Vec<u8>> {
    use litchi_xls::writer::{Builder, EncryptionProfile};
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Encrypted")?;
    builder.write_number(sheet, 0, 0, 1.0)?;
    builder.set_xor_obfuscation_password(
        "legacy",
        litchi_xls::WeakEncryptionPolicy::allow_xor_obfuscation(),
    )?;
    if builder.encryption_profile() != Some(EncryptionProfile::XorObfuscation) {
        return Err("XLS F3 encrypted fixture did not retain its profile".into());
    }
    Ok(builder.build()?.to_encrypted_bytes()?)
}

fn expect_refused(bytes: &[u8], label: &str) -> Result<()> {
    let workbook = match source_workbook(bytes) {
        Ok(workbook) => workbook,
        Err(_) => return Ok(()),
    };
    let mut edit = match workbook.edit_cells() {
        Ok(edit) => edit,
        Err(_) => return Ok(()),
    };
    let snapshot = cell_snapshot(bytes)?;
    let Some((sheet, reference, storage, value)) = snapshot.worksheets().find_map(|sheet| {
        sheet.cells().find_map(|cell| {
            let Value::Number(value) = cell.value() else {
                return None;
            };
            matches!(
                cell.storage(),
                Storage::Number | Storage::Rk | Storage::MulRk
            )
            .then_some((sheet.position(), cell.reference(), cell.storage(), *value))
        })
    }) else {
        return Err(format!("XLS F3 {label} refusal fixture has no numeric cell").into());
    };
    let staged = match storage {
        Storage::Number => edit.set_number(
            Selector::Position(Position::new(sheet)),
            reference,
            value + 1.0,
        ),
        Storage::Rk | Storage::MulRk => edit.set_numeric(
            Selector::Position(Position::new(sheet)),
            reference,
            value + 1.0,
        ),
        _ => return Err(format!("XLS F3 {label} selected an invalid storage family").into()),
    };
    if staged.is_err() {
        return Ok(());
    }
    if edit.commit().is_ok() {
        return Err(format!("XLS F3 {label} edit was accepted").into());
    }
    Ok(())
}

fn verify_security_fixtures(fixture: &Fixture) -> Result<()> {
    expect_refused(&protected_fixture()?, "protected")?;
    expect_refused(
        &add_stream(&fixture.archive, &["DigitalSignature"], b"signature")?,
        "signed",
    )?;
    expect_refused(&macro_fixture(&fixture.archive)?, "macro")?;
    expect_refused(&encrypted_fixture()?, "encrypted")?;
    Ok(())
}

fn verify_real_producer() -> Result<()> {
    let source = source_workbook(REAL_PRODUCER)?;
    let snapshot = cell_snapshot(REAL_PRODUCER)?;
    let Some((sheet, reference, value)) = snapshot.worksheets().find_map(|sheet| {
        sheet.cells().find_map(|cell| {
            let Value::Number(value) = cell.value() else {
                return None;
            };
            matches!(cell.storage(), Storage::Rk | Storage::MulRk).then_some((
                sheet.position(),
                cell.reference(),
                *value,
            ))
        })
    }) else {
        return Err("XLS F3 real producer has no RK/MulRK numeric cell".into());
    };
    let replacement = value + 1.0;
    let mut edit = source.edit_cells()?;
    edit.set_numeric(
        Selector::Position(Position::new(sheet)),
        reference,
        replacement,
    )?;
    let commit = edit.commit()?;
    let output = write_package(commit.package())?;
    let output_snapshot = cell_snapshot(&output)?;
    let cell = output_snapshot
        .worksheet(Selector::Position(Position::new(sheet)))?
        .ok_or("XLS F3 real producer output lost worksheet")?
        .cell(reference)?
        .ok_or("XLS F3 real producer output lost cell")?;
    if cell.value() != &Value::Number(replacement) {
        return Err("XLS F3 real producer output value differs".into());
    }
    let restored = commit.package().apply_cells(&commit.patch().inverse())?;
    if write_package(restored.package())? != REAL_PRODUCER {
        return Err("XLS F3 real producer inverse did not restore bytes".into());
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct Options {
    route: Route,
    corpus: CorpusChoice,
    warmup: usize,
    samples: usize,
    batch: usize,
}

fn parse_args<I>(arguments: I) -> Result<Options>
where
    I: IntoIterator<Item = String>,
{
    let mut route = None;
    let mut corpus = CorpusChoice::All;
    let mut warmup = 20;
    let mut samples = 40;
    let mut batch = 1;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--route" => {
                route = Some(Route::parse(
                    &arguments.next().ok_or("--route needs a value")?,
                )?)
            },
            "--corpus" => {
                corpus = CorpusChoice::parse(&arguments.next().ok_or("--corpus needs a value")?)?
            },
            "--warmup" => warmup = arguments.next().ok_or("--warmup needs a value")?.parse()?,
            "--samples" => samples = arguments.next().ok_or("--samples needs a value")?.parse()?,
            "--batch" => batch = arguments.next().ok_or("--batch needs a value")?.parse()?,
            "--help" | "-h" => {
                println!(
                    "usage: xls_f3_qualification --route <write-only|first-read|read-after-write> [--corpus <number|rk-mulrk|all>] [--warmup N] [--samples N] [--batch N]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown XLS F3 argument {other:?}").into()),
        }
    }
    let route = route.ok_or("--route is required")?;
    if samples == 0 || batch == 0 {
        return Err("--samples and --batch must be nonzero".into());
    }
    Ok(Options {
        route,
        corpus,
        warmup,
        samples,
        batch,
    })
}

#[derive(Debug, Default)]
struct BatchSummary {
    elapsed_ns: u64,
    edit_ns: u64,
    set_ns: u64,
    commit_ns: u64,
    write_ns: u64,
    read_ns: u64,
    output: Vec<u8>,
}

fn add_duration(total: &mut u64, value: u64) -> Result<()> {
    *total = total
        .checked_add(value)
        .ok_or("XLS F3 batched duration overflows u64")?;
    Ok(())
}

fn run_batch(
    route: Route,
    fixture: &Fixture,
    expected: &Expected,
    options: Options,
) -> Result<BatchSummary> {
    let mut summary = BatchSummary::default();
    for _ in 0..options.batch {
        let measurement = verify_sample(
            route,
            fixture,
            expected,
            run_once(route, fixture, expected.output.len())?,
        )?;
        add_duration(&mut summary.elapsed_ns, measurement.elapsed_ns)?;
        add_duration(&mut summary.edit_ns, measurement.edit_ns)?;
        add_duration(&mut summary.set_ns, measurement.set_ns)?;
        add_duration(&mut summary.commit_ns, measurement.commit_ns)?;
        add_duration(&mut summary.write_ns, measurement.write_ns)?;
        add_duration(&mut summary.read_ns, measurement.read_ns)?;
        summary.output = measurement.output.clone();
        drop(measurement);
    }
    Ok(summary)
}

fn run_fixture(route: Route, fixture: &Fixture, options: Options) -> Result<()> {
    let expected = prepare_expected(fixture)?;
    let total_iterations = options
        .warmup
        .checked_add(options.samples)
        .ok_or("XLS F3 warmup/sample count overflows")?;
    for iteration in 0..total_iterations {
        let summary = run_batch(route, fixture, &expected, options)?;
        if iteration >= options.warmup {
            println!(
                "{},{},{},{},{},{},{},{},{},{},{},{}",
                route.name(),
                fixture.name,
                iteration - options.warmup,
                summary.elapsed_ns,
                summary.edit_ns,
                summary.set_ns,
                summary.commit_ns,
                summary.write_ns,
                summary.read_ns,
                summary.output.len(),
                sha256_hex(&summary.output),
                options.batch,
            );
        }
    }
    Ok(())
}

/// Run the qualification probe from an example or an external copied source.
pub fn run_from_args<I>(arguments: I) -> Result<()>
where
    I: IntoIterator<Item = String>,
{
    let options = parse_args(arguments)?;
    let fixtures = build_fixtures(options.corpus)?;
    // These are invariant gates. They run once before any warmup or measured
    // sample and therefore never enter route timing.
    for fixture in &fixtures {
        verify_noop_and_unsupported(fixture)?;
        verify_security_fixtures(fixture)?;
    }
    verify_real_producer()?;
    println!(
        "route,corpus,sample,nanoseconds,edit_ns,set_ns,commit_ns,write_ns,read_ns,output_bytes,output_sha256,batch"
    );
    for fixture in &fixtures {
        run_fixture(options.route, fixture, options)?;
    }
    Ok(())
}
