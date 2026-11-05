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
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicU64, Ordering},
};
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::edit::RefusalKind;
use litchi_core::sheet::{Cell, CellValue};
use litchi_core::{Position, ReadAt, SourceVersion};
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Implementation {
    Baseline,
    Candidate,
}

impl Implementation {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "baseline" => Ok(Self::Baseline),
            "candidate" => Ok(Self::Candidate),
            _ => Err(format!(
                "unknown XLS F3 implementation {value:?}; expected baseline or candidate"
            )
            .into()),
        }
    }
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
    original: f64,
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
    bytes: Vec<u8>,
    accepted: usize,
    limit: usize,
}

impl Write for PrefixSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("intentional XLS F3 sink failure"));
        }
        let accepted = bytes.len().min(self.limit - self.accepted);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.accepted = self.accepted.saturating_add(accepted);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// A source whose bytes and revision can be changed after an owner opens it.
/// The F3 preflight uses this to prove that the public owner reports the typed
/// `SourceChanged` refusal instead of silently reading a stale source.
#[derive(Debug)]
struct VersionedSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: &[u8]) -> Self {
        Self {
            bytes: RwLock::new(bytes.to_vec()),
            revision: AtomicU64::new(0),
        }
    }

    fn mutate(&self) {
        let mut bytes = self
            .bytes
            .write()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if let Some(first) = bytes.first_mut() {
            *first ^= 1;
        }
        self.revision.fetch_add(1, Ordering::Relaxed);
    }

    fn bytes(&self) -> Vec<u8> {
        self.bytes
            .read()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        let bytes = self
            .bytes
            .read()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        u64::try_from(bytes.len())
            .map_err(|_| io::Error::other("XLS F3 source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).map_err(|_| {
            io::Error::new(io::ErrorKind::InvalidInput, "XLS F3 source offset overflow")
        })?;
        let bytes = self
            .bytes
            .read()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if output.is_empty() || start >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_535f_4633_0001,
            self.revision.load(Ordering::Relaxed),
        ))
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

fn source_workbook(
    bytes: &[u8],
) -> std::result::Result<SourceBackedWorkbook, litchi_xls::SourceBackedError> {
    Ok(SourceBackedWorkbook::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(bytes.to_vec()),
    ))?)
}

fn verify_source_freshness(fixture: &Fixture, implementation: Implementation) -> Result<()> {
    let specification = fixture
        .edits
        .first()
        .ok_or("XLS F3 fixture has no selected cell for freshness")?;

    // An unchanged source is the control: it must still return the original
    // value, and the source bytes must remain exactly the fixture bytes.
    let control_source = Arc::new(VersionedSource::new(&fixture.archive));
    let control_workbook = SourceBackedWorkbook::from_read_at(control_source.clone())?;
    let control = control_workbook
        .cell_value_by_index(
            specification.sheet_index,
            u32::from(specification.reference.row()),
            u32::from(specification.reference.column()),
        )?
        .ok_or("XLS F3 unchanged source lost the selected cell")?;
    let (CellValue::Float(value) | CellValue::DateTime(value)) = control else {
        return Err("XLS F3 unchanged source returned a non-numeric value".into());
    };
    if value.to_bits() != specification.original.to_bits()
        || control_source.bytes() != fixture.archive
    {
        return Err("XLS F3 unchanged source control changed its bytes or value".into());
    }

    // A stale read must report the source-backed typed freshness error.
    let source = Arc::new(VersionedSource::new(&fixture.archive));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
    source.mutate();
    let error = match workbook.cell_value_by_index(
        specification.sheet_index,
        u32::from(specification.reference.row()),
        u32::from(specification.reference.column()),
    ) {
        Ok(_) => return Err("XLS F3 changed source was unexpectedly accepted".into()),
        Err(error) => error,
    };
    if !matches!(error, litchi_xls::SourceBackedError::SourceChanged { .. }) {
        return Err(format!("XLS F3 freshness oracle returned the wrong error: {error:?}").into());
    }

    // A mutation after staging must be caught at commit, before a deferred
    // owner can retain a plan rooted in stale source bytes.
    let source = Arc::new(VersionedSource::new(&fixture.archive));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
    let mut edit = workbook.edit_cells()?;
    stage(&mut edit, fixture)?;
    source.mutate();
    match edit.commit() {
        Ok(commit) if implementation == Implementation::Baseline => {
            require_baseline_nonempty_output(
                write_package(commit.package()),
                "stage/mutate/commit",
            )?;
        },
        Ok(_) => return Err("XLS F3 stale source was accepted at commit".into()),
        Err(error) if implementation == Implementation::Candidate => {
            require_commit_source_changed(error, "stage/mutate/commit")?;
        },
        Err(error) => {
            return Err(format!(
                "XLS F3 baseline stage/mutate/commit disposition changed: {error:?}"
            )
            .into());
        },
    }

    // Replay must not let Arc identity bypass the original-source fence.  The
    // forward patch's before-owner is this same workbook, so mutating its
    // source exercises the fast path directly.
    let source = Arc::new(VersionedSource::new(&fixture.archive));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
    let commit = commit_fixture(&workbook, fixture)?;
    source.mutate();
    match workbook.apply_cells(commit.patch()) {
        Ok(replayed) if implementation == Implementation::Baseline => {
            require_baseline_nonempty_output(
                write_package(replayed.package()),
                "same-owner forward replay",
            )?;
        },
        Ok(_) => return Err("XLS F3 stale same-owner forward replay was accepted".into()),
        Err(error) if implementation == Implementation::Candidate => {
            require_commit_source_changed(error, "same-owner forward replay")?;
        },
        Err(error) => {
            return Err(format!(
                "XLS F3 baseline same-owner forward replay disposition changed: {error:?}"
            )
            .into());
        },
    }

    // The inverse patch's before-owner is the committed package.  Keep the
    // recipient identical to that owner so this covers the inverse fast path,
    // not only the independent-content replay route.
    let inverse = commit.patch().inverse();
    match commit.package().apply_cells(&inverse) {
        Ok(restored) if implementation == Implementation::Baseline => {
            require_baseline_publication_source_changed(
                write_package(restored.package()),
                "same-owner inverse replay",
            )?;
        },
        Ok(_) => return Err("XLS F3 stale same-owner inverse replay was accepted".into()),
        Err(error) if implementation == Implementation::Candidate => {
            require_commit_source_changed(error, "same-owner inverse replay")?;
        },
        Err(error) => {
            return Err(format!(
                "XLS F3 baseline same-owner inverse replay disposition changed: {error:?}"
            )
            .into());
        },
    }

    // A mutation after a deferred commit must also be caught when the first
    // semantic read asks effective_inner to materialize the composed target.
    let source = Arc::new(VersionedSource::new(&fixture.archive));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
    let mut edit = workbook.edit_cells()?;
    stage(&mut edit, fixture)?;
    let commit = edit.commit()?;
    source.mutate();
    match commit.package().cell_value_by_index(
        specification.sheet_index,
        u32::from(specification.reference.row()),
        u32::from(specification.reference.column()),
    ) {
        Ok(Some(CellValue::Float(value) | CellValue::DateTime(value)))
            if implementation == Implementation::Baseline
                && value.to_bits() == specification.replacement.to_bits() => {},
        Ok(_) => return Err("XLS F3 stale deferred source was read".into()),
        Err(error) if implementation == Implementation::Candidate => {
            if !matches!(error, litchi_xls::SourceBackedError::SourceChanged { .. }) {
                return Err(
                    format!("XLS F3 deferred read returned the wrong error: {error:?}").into(),
                );
            }
        },
        Err(error) => {
            return Err(
                format!("XLS F3 baseline deferred-read disposition changed: {error:?}").into(),
            );
        },
    }

    // A committed deferred owner must retain the original source freshness
    // fence through publication as well.
    let source = Arc::new(VersionedSource::new(&fixture.archive));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
    let mut edit = workbook.edit_cells()?;
    stage(&mut edit, fixture)?;
    let commit = edit.commit()?;
    source.mutate();
    let mut output = Vec::new();
    match commit.package().write_to(&mut output) {
        Ok(_) if implementation == Implementation::Baseline => {
            if output.is_empty() {
                return Err(
                    "XLS F3 baseline mutate/committed-write published empty stale output".into(),
                );
            }
        },
        Ok(_) => return Err("XLS F3 stale committed source was published".into()),
        Err(error) if implementation == Implementation::Candidate => {
            require_output_source_changed(error, "mutate/committed-write")?;
        },
        Err(error) if implementation == Implementation::Baseline => {
            require_output_source_changed(error, "mutate/committed-write")?;
        },
        Err(error) => {
            return Err(format!(
                "XLS F3 baseline mutate/committed-write disposition changed: {error:?}"
            )
            .into());
        },
    }
    Ok(())
}

fn require_baseline_nonempty_output(output: Result<Vec<u8>>, label: &str) -> Result<()> {
    match output {
        Ok(output) if !output.is_empty() => Ok(()),
        Ok(_) => Err(format!("XLS F3 baseline {label} published empty stale output").into()),
        Err(error) => Err(format!(
            "XLS F3 baseline {label} returned an unrecorded publication error: {error:?}"
        )
        .into()),
    }
}

fn require_baseline_publication_source_changed(output: Result<Vec<u8>>, label: &str) -> Result<()> {
    match output {
        Ok(_) => Err(format!("XLS F3 baseline {label} unexpectedly published stale output").into()),
        Err(error)
            if error
                .downcast_ref::<litchi_xls::source_backed::OutputError>()
                .is_some_and(|error| {
                    matches!(
                        error,
                        litchi_xls::source_backed::OutputError::Package(
                            litchi_xls::package::Error::Source(
                                litchi_xls::SourceBackedError::SourceChanged { .. }
                            )
                        )
                    )
                }) =>
        {
            Ok(())
        },
        Err(error) => Err(format!(
            "XLS F3 baseline {label} returned an unrecorded publication error: {error:?}"
        )
        .into()),
    }
}

fn require_commit_source_changed(error: litchi_xls::Error, label: &str) -> Result<()> {
    match error {
        litchi_xls::Error::Io(error)
            if error.kind() == io::ErrorKind::Other
                && error.to_string() == "source changed during source-backed replay" =>
        {
            Ok(())
        },
        other => Err(format!(
            "XLS F3 {label} returned the wrong source freshness error: {other:?}"
        )
        .into()),
    }
}

fn require_output_source_changed(
    error: litchi_xls::source_backed::OutputError,
    label: &str,
) -> Result<()> {
    match error {
        litchi_xls::source_backed::OutputError::Package(litchi_xls::package::Error::Source(
            litchi_xls::SourceBackedError::SourceChanged { .. },
        )) => Ok(()),
        other => Err(format!(
            "XLS F3 {label} returned the wrong output freshness error: {other:?}"
        )
        .into()),
    }
}

fn cell_snapshot(bytes: &[u8]) -> Result<litchi_xls::cell_values::Snapshot> {
    Ok(litchi_xls::Package::from_bytes(bytes)?.cell_values()?)
}

fn write_package(package: &SourceBackedWorkbook) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn write_commit(
    commit: &cells::Commit,
    mut sink: CountingSink,
    capacity: usize,
) -> Result<CountingSink> {
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
    // The publication sink's bounded allocation is setup, not route work.
    // Allocate it before the route timer so first-read and write-only timings
    // compare only the owner operation and the bounded stream writes.
    let publication_sink = CountingSink::new(output_capacity)?;
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

    let output;
    let mut write_ns = 0;
    let mut read_ns = 0;
    let route_elapsed_ns;
    match route {
        Route::WriteOnly => {
            let write_started = Instant::now();
            let sink = write_commit(&commit, publication_sink, output_capacity)?;
            write_ns = elapsed_ns(write_started.elapsed())?;
            output = sink.bytes;
            route_elapsed_ns = elapsed_ns(route_started.elapsed())?;
        },
        Route::FirstRead => {
            let read_started = Instant::now();
            read_first_target(commit.package(), fixture)?;
            read_ns = elapsed_ns(read_started.elapsed())?;
            // Stop the first-read route clock before publication. The output
            // write remains an untimed correctness check for this route.
            route_elapsed_ns = elapsed_ns(route_started.elapsed())?;
            // Publication is intentionally after the first-read clock. This
            // route isolates the moved target reopen from output streaming.
            output = write_commit(&commit, publication_sink, output_capacity)?.bytes;
        },
        Route::ReadAfterWrite => {
            let write_started = Instant::now();
            let sink = write_commit(&commit, publication_sink, output_capacity)?;
            write_ns = elapsed_ns(write_started.elapsed())?;
            output = sink.bytes;
            let read_started = Instant::now();
            read_first_target(commit.package(), fixture)?;
            read_ns = elapsed_ns(read_started.elapsed())?;
            route_elapsed_ns = elapsed_ns(route_started.elapsed())?;
        },
    }

    Ok(Measurement {
        elapsed_ns: route_elapsed_ns,
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
    if write_package(&source)? != fixture.archive {
        return Err(format!(
            "XLS F3 {} source bytes changed during commit preflight",
            fixture.name
        )
        .into());
    }
    let output = write_package(commit.package())?;
    if output == fixture.archive {
        return Err(format!("XLS F3 {} edit did not change its source", fixture.name).into());
    }
    verify_output(fixture, &output)?;
    verify_patch(&source, &commit)?;
    verify_partial_publication(&commit, &output)?;
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
    if write_package(&measurement.source)? != fixture.archive {
        return Err(format!(
            "XLS F3 {} {} source bytes changed during measurement",
            fixture.name,
            route.name()
        )
        .into());
    }
    verify_patch(&measurement.source, &measurement.commit)?;
    if matches!(route, Route::FirstRead | Route::ReadAfterWrite) {
        verify_committed_values(measurement.commit.package(), fixture)?;
    }
    black_box((&measurement.output, &measurement.commit));
    Ok(measurement)
}

fn verify_partial_publication(commit: &cells::Commit, expected: &[u8]) -> Result<()> {
    let mut sink = PrefixSink {
        bytes: Vec::new(),
        accepted: 0,
        limit: expected.len() / 2,
    };
    let error = match commit.package().write_to(&mut sink) {
        Ok(_) => return Err("XLS F3 partial sink unexpectedly accepted complete output".into()),
        Err(error) => error,
    };
    let litchi_xls::source_backed::OutputError::Overlay(
        litchi_cfb::OverlayError::IncompleteOutput {
            progress:
                litchi_cfb::OutputProgress::Prefix {
                    accepted,
                    expected: progress_expected,
                },
            source,
        },
    ) = error
    else {
        return Err(
            format!("XLS F3 partial sink returned an untyped output error: {error:?}").into(),
        );
    };
    let authored = u64::try_from(sink.bytes.len())?;
    let expected_bytes = u64::try_from(expected.len())?;
    if accepted != authored
        || progress_expected != expected_bytes
        || sink.bytes.is_empty()
        || sink.bytes.len() >= expected.len()
        || !matches!(source.as_ref(), litchi_cfb::OverlayError::Io(_))
        || sink.bytes.as_slice() != &expected[..sink.bytes.len()]
    {
        return Err(format!(
            "XLS F3 partial sink progress/prefix mismatch: accepted={accepted}, authored={authored}, expected={progress_expected}, sink_bytes={}",
            sink.bytes.len()
        )
        .into());
    }
    Ok(())
}

fn verify_patch(source: &SourceBackedWorkbook, commit: &cells::Commit) -> Result<()> {
    let patch = commit.patch();
    // Reopen both equal-content recipients independently.  Reusing the
    // original owner or committed owner would only prove Arc identity and
    // would not exercise the patch's content lineage contract.
    let source_bytes = write_package(source)?;
    let committed_bytes = write_package(commit.package())?;
    let independent_source = source_workbook(&source_bytes)?;
    let applied = independent_source.apply_cells(patch)?;
    let applied_bytes = write_package(applied.package())?;
    if applied_bytes != committed_bytes {
        return Err("XLS F3 patch replay changed committed bytes".into());
    }
    let stale_error = match commit.package().apply_cells(patch) {
        Ok(_) => return Err("XLS F3 stale recipient unexpectedly accepted the patch".into()),
        Err(error) => error,
    };
    if !matches!(stale_error, litchi_xls::Error::PatchConflict { .. }) {
        return Err(
            format!("XLS F3 stale recipient returned the wrong error: {stale_error:?}").into(),
        );
    }
    let independent_target = source_workbook(&committed_bytes)?;
    let restored = independent_target.apply_cells(&patch.inverse())?;
    let restored_bytes = write_package(restored.package())?;
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
    let mut source_ole = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default())?;
    let mut candidate_ole =
        OleFile::open(Cursor::new(candidate), litchi_cfb::OleFileLimits::default())?;
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
                original: 42.0,
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
                original: match cell.value() {
                    Value::Number(value) => *value,
                    _ => unreachable!("the storage/value guard above admits only numbers"),
                },
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

fn verify_noop_and_unsupported(fixture: &Fixture, implementation: Implementation) -> Result<()> {
    let source = source_workbook(&fixture.archive)?;
    let noop = source.edit_cells()?.commit()?;
    let diagnostics = noop.diagnostics();
    if diagnostics.changed()
        || !noop.patch().is_noop()
        || diagnostics.source_fingerprint() != diagnostics.target_fingerprint()
    {
        return Err(format!("XLS F3 {} no-op changed its source", fixture.name).into());
    }
    if write_package(noop.package())? != fixture.archive {
        return Err(format!("XLS F3 {} no-op output differs", fixture.name).into());
    }

    let first = fixture
        .edits
        .first()
        .ok_or("XLS F3 fixture has no edit for refusal checks")?;
    // Boolean values have no fixed-width representation in any of the
    // numeric Number/RK/MulRK slots. This is deliberately a typed refusal,
    // not a broad "some error" assertion.
    let mut unsupported = source.edit_cells()?;
    let staged = unsupported.set_value(first.selector, first.reference, Value::Boolean(true));
    match staged {
        Err(error) => require_refusal(
            error,
            &format!("{} unsupported target", fixture.name),
            RefusalKind::UnsupportedTarget,
            "value is not representable in the fixed-width slot",
        )?,
        Ok(()) => {
            let error = match unsupported.commit() {
                Ok(_) => {
                    return Err(format!(
                        "XLS F3 {} accepted an unsupported numeric target",
                        fixture.name
                    )
                    .into());
                },
                Err(error) => error,
            };
            require_unsupported_target_commit(
                error,
                &format!("{} unsupported target commit", fixture.name),
                implementation,
            )?;
        },
    }

    let mut wrong_family = source.edit_cells()?;
    if matches!(first.storage, Storage::Rk | Storage::MulRk) {
        let error =
            match wrong_family.set_number(first.selector, first.reference, first.replacement) {
                Ok(()) => {
                    return Err(format!(
                        "XLS F3 {} accepted set_number for compressed storage",
                        fixture.name
                    )
                    .into());
                },
                Err(error) => error,
            };
        let legacy_expected = format!(
            "cell ({}, {}) is not encoded as a standalone BIFF8 Number record",
            first.reference.row(),
            first.reference.column()
        );
        match implementation {
            Implementation::Baseline => match error {
                litchi_xls::Error::UnsupportedFeature(message) if message == legacy_expected => {},
                other => {
                    return Err(format!(
                        "XLS F3 {} baseline wrong-family disposition changed: {other:?}",
                        fixture.name
                    )
                    .into());
                },
            },
            Implementation::Candidate => require_refusal(
                error,
                &format!("{} wrong-family", fixture.name),
                RefusalKind::UnsupportedTarget,
                "the selected cell is not an existing standalone BIFF8 Number",
            )?,
        }
    }
    Ok(())
}

fn require_unsupported_target_commit(
    error: litchi_xls::Error,
    label: &str,
    implementation: Implementation,
) -> Result<()> {
    match implementation {
        Implementation::Baseline => match error {
            litchi_xls::Error::UnsupportedFeature(message)
                if message
                    == "source-backed numeric plan accepts no structural or resource edits" =>
            {
                Ok(())
            },
            other => Err(format!("XLS F3 {label} baseline disposition changed: {other:?}").into()),
        },
        Implementation::Candidate => require_refusal(
            error,
            label,
            RefusalKind::UnsupportedOperation,
            "source-backed numeric plan accepts no structural or resource edits",
        ),
    }
}

fn require_refusal(
    error: litchi_xls::Error,
    label: &str,
    kind: RefusalKind,
    reason: &'static str,
) -> Result<()> {
    match error {
        litchi_xls::Error::Refused(refusal)
            if refusal.kind() == kind
                && refusal.operation() == "edit cells"
                && refusal.reason() == reason =>
        {
            Ok(())
        },
        other => Err(format!("XLS F3 {label} returned an unexpected refusal: {other:?}").into()),
    }
}

fn copy_streams(source: &[u8]) -> Result<OleWriter> {
    let mut ole = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default())?;
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

#[derive(Clone, Copy)]
enum SecurityExpectation {
    RefuseEdit {
        kind: RefusalKind,
        reason: &'static str,
    },
    RefuseSignedOpen,
    RefuseEncryptedOpen,
}

#[derive(Debug)]
enum SecurityOpenError {
    Source(litchi_xls::SourceBackedError),
    Edit(litchi_xls::Error),
}

const EXPECTED_TYPED_PROTECTED_REFUSAL: &str =
    "signed, encrypted, or DRM containers are not eligible for source-backed editing";
const EXPECTED_LEGACY_PROTECTED_REFUSAL: &str =
    "signed, encrypted, or DRM containers are not eligible for object editing";

fn require_signed_marker(bytes: &[u8]) -> Result<()> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let mut pending = vec![Vec::<String>::new()];
    let mut signed_path = None;
    while let Some(path) = pending.pop() {
        let references = path.iter().map(String::as_str).collect::<Vec<_>>();
        for entry in ole.list_directory_entries(&references)? {
            let mut full_path = path.clone();
            full_path.push(entry.name.clone());
            if entry.entry_type == 2
                && entry.name.eq_ignore_ascii_case("DigitalSignature")
                && litchi_ole_common::protection::is_protected_component(&entry.name)
            {
                signed_path = Some(full_path);
            } else if entry.entry_type == 1 {
                pending.push(full_path);
            }
        }
    }
    let Some(signed_path) = signed_path else {
        return Err("XLS F3 signed fixture lacks its DigitalSignature marker".into());
    };
    let references = signed_path.iter().map(String::as_str).collect::<Vec<_>>();
    if ole.open_stream(&references)?.is_empty() {
        return Err("XLS F3 signed fixture has an empty DigitalSignature marker".into());
    }
    Ok(())
}

fn require_encrypted_marker(bytes: &[u8]) -> Result<()> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let workbook = ole
        .open_stream(&["Workbook"])
        .or_else(|_| ole.open_stream(&["Book"]))?;
    let mut offset = 0usize;
    while offset < workbook.len() {
        let header_end = offset
            .checked_add(4)
            .ok_or("XLS F3 encrypted marker header offset overflow")?;
        let header = workbook
            .get(offset..header_end)
            .ok_or("XLS F3 encrypted marker has a truncated BIFF header")?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = header_end
            .checked_add(payload_len)
            .ok_or("XLS F3 encrypted marker record length overflow")?;
        let payload = workbook
            .get(header_end..end)
            .ok_or("XLS F3 encrypted marker has a truncated BIFF payload")?;
        if kind == 0x002f && payload.len() >= 2 {
            return Ok(());
        }
        offset = end;
    }
    Err("XLS F3 encrypted fixture lacks a BIFF FILEPASS marker".into())
}

fn require_signed_refusal(
    error: SecurityOpenError,
    label: &str,
    implementation: Implementation,
) -> Result<()> {
    match implementation {
        Implementation::Baseline => match error {
            SecurityOpenError::Edit(litchi_xls::Error::Object(litchi_ole_common::Error::Cfb(
                litchi_cfb::OleError::InvalidFormat(message),
            ))) if message == EXPECTED_LEGACY_PROTECTED_REFUSAL => Ok(()),
            other => {
                Err(format!("XLS F3 {label} baseline signed disposition changed: {other:?}").into())
            },
        },
        Implementation::Candidate => match error {
            SecurityOpenError::Edit(litchi_xls::Error::Refused(refusal))
                if refusal.kind() == RefusalKind::Protected
                    && refusal.operation() == "edit cells"
                    && refusal.reason() == EXPECTED_TYPED_PROTECTED_REFUSAL =>
            {
                Ok(())
            },
            other => Err(
                format!("XLS F3 {label} candidate signed disposition changed: {other:?}").into(),
            ),
        },
    }
}

fn require_encrypted_refusal(error: SecurityOpenError, label: &str) -> Result<()> {
    match error {
        SecurityOpenError::Source(litchi_xls::SourceBackedError::EncryptedUnsupported)
        | SecurityOpenError::Edit(litchi_xls::Error::PasswordRequired) => Ok(()),
        other => Err(format!(
            "XLS F3 {label} fixture returned an unexpected encryption refusal: {other:?}"
        )
        .into()),
    }
}

fn verify_security_refusal(
    bytes: &[u8],
    label: &str,
    expectation: SecurityExpectation,
    implementation: Implementation,
) -> Result<()> {
    match expectation {
        SecurityExpectation::RefuseEdit { kind, reason } => {
            verify_policy_refusal(bytes, label, kind, reason)
        },
        SecurityExpectation::RefuseSignedOpen => {
            require_signed_marker(bytes)?;
            let opened = source_workbook(bytes)
                .map_err(SecurityOpenError::Source)
                .and_then(|workbook| {
                    workbook
                        .edit_cells()
                        .map(|_| ())
                        .map_err(SecurityOpenError::Edit)
                });
            match opened {
                Ok(()) => Err(format!("XLS F3 {label} fixture unexpectedly opened").into()),
                Err(error) => require_signed_refusal(error, label, implementation),
            }
        },
        SecurityExpectation::RefuseEncryptedOpen => {
            require_encrypted_marker(bytes)?;
            let opened = source_workbook(bytes)
                .map_err(SecurityOpenError::Source)
                .and_then(|workbook| {
                    workbook
                        .edit_cells()
                        .map(|_| ())
                        .map_err(SecurityOpenError::Edit)
                });
            match opened {
                Ok(()) => Err(format!("XLS F3 {label} fixture unexpectedly opened").into()),
                Err(error) => require_encrypted_refusal(error, label),
            }
        },
    }
}

fn verify_policy_refusal(
    bytes: &[u8],
    label: &str,
    kind: RefusalKind,
    reason: &'static str,
) -> Result<()> {
    let workbook = source_workbook(bytes)?;
    let snapshot = litchi_xls::Package::from_source(workbook.clone())?.cell_values()?;
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
        return Err(format!("XLS F3 {label} fixture has no numeric guard cell").into());
    };
    let mut edit = match workbook.edit_cells() {
        Ok(edit) => edit,
        Err(error) => return require_refusal(error, label, kind, reason),
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
    match staged {
        Err(error) => require_refusal(error, label, kind, reason),
        Ok(()) => {
            let error = match edit.commit() {
                Ok(_) => return Err(format!("XLS F3 {label} edit was accepted").into()),
                Err(error) => error,
            };
            require_refusal(error, label, kind, reason)
        },
    }
}

fn verify_security_fixtures(fixture: &Fixture, implementation: Implementation) -> Result<()> {
    verify_security_refusal(
        &protected_fixture()?,
        "protected",
        SecurityExpectation::RefuseEdit {
            kind: RefusalKind::Protected,
            reason: "protected worksheets are not eligible for source-backed numeric edits",
        },
        implementation,
    )?;
    verify_security_refusal(
        &add_stream(&fixture.archive, &["DigitalSignature"], b"signature")?,
        "signed",
        SecurityExpectation::RefuseSignedOpen,
        implementation,
    )?;
    verify_security_refusal(
        &macro_fixture(&fixture.archive)?,
        "macro",
        SecurityExpectation::RefuseEdit {
            kind: RefusalKind::UnsupportedOperation,
            reason: "macro-bearing XLS sources are not eligible for source-backed numeric edits",
        },
        implementation,
    )?;
    verify_security_refusal(
        &encrypted_fixture()?,
        "encrypted",
        SecurityExpectation::RefuseEncryptedOpen,
        implementation,
    )?;
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
    let replayed = source_workbook(REAL_PRODUCER)?.apply_cells(commit.patch())?;
    if write_package(replayed.package())? != output {
        return Err("XLS F3 real producer forward replay changed bytes".into());
    }
    let independent_target = source_workbook(&output)?;
    let restored = independent_target.apply_cells(&commit.patch().inverse())?;
    if write_package(restored.package())? != REAL_PRODUCER {
        return Err("XLS F3 real producer inverse did not restore bytes".into());
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct Options {
    implementation: Implementation,
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
    let mut implementation = None;
    let mut corpus = CorpusChoice::All;
    let mut warmup = 20;
    let mut samples = 40;
    let mut batch = 1;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--implementation" => {
                implementation = Some(Implementation::parse(
                    &arguments.next().ok_or("--implementation needs a value")?,
                )?)
            },
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
                    "usage: xls_f3_qualification --implementation <baseline|candidate> --route <write-only|first-read|read-after-write> [--corpus <number|rk-mulrk|all>] [--warmup N] [--samples N] [--batch N]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown XLS F3 argument {other:?}").into()),
        }
    }
    let implementation = implementation.ok_or("--implementation is required")?;
    let route = route.ok_or("--route is required")?;
    if samples == 0 || batch == 0 {
        return Err("--samples and --batch must be nonzero".into());
    }
    Ok(Options {
        implementation,
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
        verify_source_freshness(fixture, options.implementation)?;
        verify_noop_and_unsupported(fixture, options.implementation)?;
        verify_security_fixtures(fixture, options.implementation)?;
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
