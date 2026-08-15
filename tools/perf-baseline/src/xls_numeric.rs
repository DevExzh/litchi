//! Fixed-width native XLS numeric CRUD evidence.
//!
//! The four selectors in this module are intentionally opt-in.  They measure
//! the semantic edit/commit boundary separately from publication while all
//! corpus preparation, source ingress, complete reopen, and safety checks stay
//! outside the timed interval.

use super::{
    Case, CaseResult, Corpus, CorpusManifest, CountingSink, InstrumentedSource, PayloadKind,
    SourceSnapshot, SourceSummary, sha256_hex,
};
use litchi_cfb::{OleFile, OleWriter, PublishReport};
use litchi_core::ReadAt;
use serde::Serialize;
use std::io::{self, Cursor, Write};
use std::time::Instant;

const CORPUS_GENERATOR: &str = "litchi-xls-rk-mulrk-publication-v1";
const OPAQUE_STREAM_COUNT: usize = 2;
const OPAQUE_STREAM_BYTES: usize = 96 * 1024;
const MAX_WRITE: u64 = 64 * 1024;
const REAL_PRODUCER: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/spreadsheet/54016.xls");

fn reserve_exact<T>(
    values: &mut Vec<T>,
    amount: usize,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    values
        .try_reserve_exact(amount)
        .map_err(|error| format!("{label} allocation failed: {error}").into())
}

fn push_bounded<T>(
    values: &mut Vec<T>,
    value: T,
    limit: usize,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    if values.len() >= limit {
        return Err(format!("{label} exceeded its bounded capacity").into());
    }
    values
        .try_reserve(1)
        .map_err(|error| format!("{label} allocation failed: {error}"))?;
    values.push(value);
    Ok(())
}

fn clone_bytes(bytes: &[u8], label: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut cloned = Vec::new();
    reserve_exact(&mut cloned, bytes.len(), label)?;
    cloned.extend_from_slice(bytes);
    Ok(cloned)
}

fn clone_record(record: &RawRecord, label: &str) -> Result<RawRecord, Box<dyn std::error::Error>> {
    Ok(RawRecord {
        kind: record.kind,
        payload: clone_bytes(&record.payload, label)?,
    })
}

/// Per-case native XLS numeric publication evidence.  The complete target
/// bytes are deliberately reported for both implementations: source-backed
/// publication avoids Workbook reserialization but retains a complete target
/// snapshot and therefore makes no bounded-artifact-memory claim.
#[derive(Clone, Debug, Default, Serialize)]
pub(crate) struct XlsNumericSourceSummary {
    pub(crate) source_counter_scope: &'static str,
    pub(crate) implementation: &'static str,
    pub(crate) family: &'static str,
    pub(crate) source_backed: bool,
    pub(crate) update_count: usize,
    pub(crate) sample_count: usize,
    pub(crate) input_cfb_bytes: u64,
    pub(crate) output_cfb_bytes: u64,
    pub(crate) source_workbook_bytes: u64,
    pub(crate) target_workbook_bytes: u64,
    pub(crate) sink_capacity_bytes: u64,
    pub(crate) expected_output_sha256: String,
    pub(crate) owned_input_scope: &'static str,
    pub(crate) edit_ns: Vec<u64>,
    pub(crate) set_ns: Vec<u64>,
    pub(crate) commit_ns: Vec<u64>,
    pub(crate) publication_ns: Vec<u64>,
    pub(crate) total_ns: Vec<u64>,
    pub(crate) complete_target_materialized_bytes: Vec<u64>,
    pub(crate) sink_bytes: Vec<u64>,
    pub(crate) sink_write_calls: Vec<u64>,
    pub(crate) sink_digests: Vec<String>,
    pub(crate) source_bytes: Vec<u64>,
    pub(crate) source_workbook_bytes_per_sample: Vec<u64>,
    pub(crate) target_workbook_bytes_per_sample: Vec<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) splice_count: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) replacement_bytes: Option<Vec<u64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) changed_spans: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) source_fingerprints: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) target_fingerprints: Option<Vec<String>>,
}

fn reserve_summary(
    summary: &mut XlsNumericSourceSummary,
    samples: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    reserve_exact(
        &mut summary.edit_ns,
        samples,
        "XLS numeric edit timing evidence",
    )?;
    reserve_exact(
        &mut summary.set_ns,
        samples,
        "XLS numeric set timing evidence",
    )?;
    reserve_exact(
        &mut summary.commit_ns,
        samples,
        "XLS numeric commit timing evidence",
    )?;
    reserve_exact(
        &mut summary.publication_ns,
        samples,
        "XLS numeric publication timing evidence",
    )?;
    reserve_exact(
        &mut summary.total_ns,
        samples,
        "XLS numeric total timing evidence",
    )?;
    reserve_exact(
        &mut summary.complete_target_materialized_bytes,
        samples,
        "XLS numeric materialized-target evidence",
    )?;
    reserve_exact(
        &mut summary.sink_bytes,
        samples,
        "XLS numeric sink-byte evidence",
    )?;
    reserve_exact(
        &mut summary.sink_write_calls,
        samples,
        "XLS numeric sink-write evidence",
    )?;
    reserve_exact(
        &mut summary.sink_digests,
        samples,
        "XLS numeric sink-digest evidence",
    )?;
    reserve_exact(
        &mut summary.source_bytes,
        samples,
        "XLS numeric source-size evidence",
    )?;
    reserve_exact(
        &mut summary.source_workbook_bytes_per_sample,
        samples,
        "XLS numeric source-workbook evidence",
    )?;
    reserve_exact(
        &mut summary.target_workbook_bytes_per_sample,
        samples,
        "XLS numeric target-workbook evidence",
    )?;
    if let Some(values) = summary.splice_count.as_mut() {
        reserve_exact(values, samples, "XLS numeric splice evidence")?;
    }
    if let Some(values) = summary.replacement_bytes.as_mut() {
        reserve_exact(values, samples, "XLS numeric replacement-byte evidence")?;
    }
    if let Some(values) = summary.changed_spans.as_mut() {
        reserve_exact(values, samples, "XLS numeric changed-span evidence")?;
    }
    if let Some(values) = summary.source_fingerprints.as_mut() {
        reserve_exact(values, samples, "XLS numeric source-fingerprint evidence")?;
    }
    if let Some(values) = summary.target_fingerprints.as_mut() {
        reserve_exact(values, samples, "XLS numeric target-fingerprint evidence")?;
    }
    Ok(())
}

/// Builds the deterministic native workbook used by the packed-RK selectors.
///
/// The public writer emits the source Number records.  The narrow corpus
/// conversion below changes one standalone Number to RK and one adjacent pair
/// to MulRK using the same BIFF fields and RK encoding as the writer.  The
/// complete CFB is then rebuilt with two opaque sibling streams so publication
/// checks can prove that untouched members and topology survive.
pub(crate) fn build_rk_mulrk_corpus() -> Result<Corpus, Box<dyn std::error::Error>> {
    let mut workbook_writer = litchi_xls::writer::Writer::new();
    let packed = workbook_writer.add_worksheet("Packed")?;
    workbook_writer.write_number(packed, 0, 0, 1.0)?;
    workbook_writer.write_number(packed, 1, 0, 2.0)?;
    workbook_writer.write_number(packed, 1, 1, 3.0)?;
    workbook_writer.write_number(packed, 3, 2, 17.0)?;
    let mut package = Cursor::new(Vec::new());
    workbook_writer.write_to(&mut package)?;

    let mut source_ole = OleFile::open(Cursor::new(package.into_inner()))?;
    let source_workbook = source_ole.open_stream(&["Workbook"])?;
    let rebuilt_workbook = convert_numbers_to_rk_mulrk(&source_workbook)?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], rebuilt_workbook)?;
    writer.create_storage(&["OpaquePayloads"])?;
    for index in 0..OPAQUE_STREAM_COUNT {
        let name = format!("Payload{index:03}");
        writer.create_stream_owned(
            &["OpaquePayloads", name.as_str()],
            super::payload_bytes(
                PayloadKind::Incompressible,
                70_000 + index,
                OPAQUE_STREAM_BYTES,
            ),
        )?;
    }
    writer.create_stream(
        &["OpaqueMetadata"],
        b"litchi-xls-rk-mulrk-opaque-metadata-v1",
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();

    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    let archive_member_count = parsed.list_streams().len();
    let target_payload = parsed.open_stream(&["Workbook"])?;
    if archive_member_count != OPAQUE_STREAM_COUNT + 2 {
        return Err("XLS RK/MulRK corpus stream inventory differs from specification".into());
    }
    let uncompressed_payload_bytes = target_payload
        .len()
        .checked_add(OPAQUE_STREAM_COUNT * OPAQUE_STREAM_BYTES)
        .and_then(|bytes| bytes.checked_add(b"litchi-xls-rk-mulrk-opaque-metadata-v1".len()))
        .ok_or("XLS RK/MulRK corpus payload size overflow")?;

    let snapshot = litchi_xls::cell_values::Snapshot::from_bytes(clone_bytes(
        &archive,
        "XLS RK/MulRK corpus snapshot",
    )?)?;
    let sheet = snapshot
        .worksheet(litchi_xls::cell_values::Selector::Name("Packed"))?
        .ok_or("XLS RK/MulRK corpus lost Packed worksheet")?;
    let cell_count = sheet.cells().count();
    let rk_count = sheet
        .cells()
        .filter(|cell| cell.storage() == litchi_xls::cell_values::Storage::Rk)
        .count();
    let mul_rk_count = sheet
        .cells()
        .filter(|cell| cell.storage() == litchi_xls::cell_values::Storage::MulRk)
        .count();
    if rk_count != 1 || mul_rk_count != 2 {
        return Err("XLS RK/MulRK corpus does not contain one RK and one MulRK record".into());
    }

    Ok(Corpus {
        manifest: CorpusManifest {
            name: "xls-rk-mulrk-deterministic".to_owned(),
            generator: CORPUS_GENERATOR,
            package_format: "XLS/CFB",
            shape: "one-rk-one-mulrk",
            payload_kind: PayloadKind::Incompressible.name(),
            compression: "none",
            entry_count: cell_count,
            archive_member_count,
            entry_bytes: OPAQUE_STREAM_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "Workbook".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "Workbook".to_owned(),
        target_payload,
        xlsx: None,
    })
}

#[derive(Clone, Debug)]
struct RawRecord {
    kind: u16,
    payload: Vec<u8>,
}

fn convert_numbers_to_rk_mulrk(workbook: &[u8]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let records = scan_records(workbook)?;
    let mut number_indices = Vec::new();
    reserve_exact(
        &mut number_indices,
        records.len(),
        "XLS RK/MulRK Number-record index",
    )?;
    for (index, record) in records.iter().enumerate() {
        if record.kind == 0x0203 {
            push_bounded(
                &mut number_indices,
                index,
                records.len(),
                "XLS RK/MulRK Number-record index",
            )?;
        }
    }
    let standalone = *number_indices
        .first()
        .ok_or("XLS RK/MulRK source writer emitted no Number records")?;
    let pair = number_indices
        .windows(2)
        .find(|pair| {
            let left = &records[pair[0]].payload;
            let right = &records[pair[1]].payload;
            pair[1] == pair[0] + 1
                && left.get(0..2) == right.get(0..2)
                && left
                    .get(2..4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .map(u16::from_le_bytes)
                    .zip(
                        right
                            .get(2..4)
                            .and_then(|bytes| bytes.try_into().ok())
                            .map(u16::from_le_bytes),
                    )
                    .is_some_and(|(left, right)| right == left + 1)
                && pair[0] != standalone
        })
        .map(|pair| [pair[0], pair[1]])
        .ok_or("XLS RK/MulRK source writer emitted no contiguous Number pair")?;

    let mut transformed = Vec::new();
    reserve_exact(
        &mut transformed,
        records.len(),
        "XLS transformed BIFF records",
    )?;
    for (index, record) in records.iter().enumerate() {
        if index == pair[1] {
            continue;
        }
        if index == standalone {
            let mut payload = clone_bytes(&record.payload, "XLS RK standalone payload")?;
            payload.truncate(10);
            let value = f64::from_le_bytes(
                record
                    .payload
                    .get(6..14)
                    .ok_or("XLS Number record is truncated")?
                    .try_into()?,
            );
            payload[6..10].copy_from_slice(&encode_integer_rk(value)?.to_le_bytes());
            push_bounded(
                &mut transformed,
                RawRecord {
                    kind: 0x027e,
                    payload,
                },
                records.len(),
                "XLS transformed BIFF records",
            )?;
            continue;
        }
        if index == pair[0] {
            let first = records
                .get(pair[0])
                .ok_or("XLS MulRK first Number disappeared")?;
            let second = records
                .get(pair[1])
                .ok_or("XLS MulRK second Number disappeared")?;
            let first_value = f64::from_le_bytes(first.payload[6..14].try_into()?);
            let second_value = f64::from_le_bytes(second.payload[6..14].try_into()?);
            let mut payload = Vec::new();
            reserve_exact(&mut payload, 18, "XLS MulRK payload")?;
            payload.extend_from_slice(&first.payload[0..4]);
            payload.extend_from_slice(&first.payload[4..6]);
            payload.extend_from_slice(&encode_integer_rk(first_value)?.to_le_bytes());
            payload.extend_from_slice(&second.payload[4..6]);
            payload.extend_from_slice(&encode_integer_rk(second_value)?.to_le_bytes());
            payload.extend_from_slice(
                &u16::from_le_bytes(second.payload[2..4].try_into()?).to_le_bytes(),
            );
            push_bounded(
                &mut transformed,
                RawRecord {
                    kind: 0x00bd,
                    payload,
                },
                records.len(),
                "XLS transformed BIFF records",
            )?;
            continue;
        }
        push_bounded(
            &mut transformed,
            clone_record(record, "XLS transformed BIFF payload")?,
            records.len(),
            "XLS transformed BIFF records",
        )?;
    }

    let mut rebuilt = Vec::new();
    reserve_exact(
        &mut rebuilt,
        workbook.len(),
        "XLS transformed BIFF workbook",
    )?;
    for record in transformed {
        rebuilt.extend_from_slice(&record.kind.to_le_bytes());
        rebuilt.extend_from_slice(
            &u16::try_from(record.payload.len())
                .map_err(|_error| "XLS transformed BIFF payload exceeds u16")?
                .to_le_bytes(),
        );
        rebuilt.extend_from_slice(&record.payload);
    }
    Ok(rebuilt)
}

fn scan_records(workbook: &[u8]) -> Result<Vec<RawRecord>, Box<dyn std::error::Error>> {
    let mut records = Vec::new();
    reserve_exact(
        &mut records,
        workbook.len().saturating_div(4).saturating_add(1),
        "XLS BIFF record inventory",
    )?;
    let mut offset = 0usize;
    while offset < workbook.len() {
        let header_end = offset
            .checked_add(4)
            .ok_or("XLS BIFF header offset overflow")?;
        let header = workbook
            .get(offset..header_end)
            .ok_or("XLS BIFF record has a truncated header")?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = header_end
            .checked_add(payload_len)
            .ok_or("XLS BIFF record length overflow")?;
        let payload = clone_bytes(
            workbook
                .get(header_end..end)
                .ok_or("XLS BIFF record has a truncated payload")?,
            "XLS BIFF record payload",
        )?;
        push_bounded(
            &mut records,
            RawRecord { kind, payload },
            workbook.len().saturating_div(4).saturating_add(1),
            "XLS BIFF record inventory",
        )?;
        offset = end;
    }
    Ok(records)
}

fn encode_integer_rk(value: f64) -> Result<u32, Box<dyn std::error::Error>> {
    let integer = i32::try_from(value as i64).map_err(|_error| "XLS RK integer overflow")?;
    if f64::from(integer) != value || !(-(1_i32 << 29)..(1_i32 << 29)).contains(&integer) {
        return Err("XLS deterministic RK corpus value is not an integer RK".into());
    }
    Ok((integer as u32).wrapping_shl(2) | 0x02)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Family {
    Number,
    RkMulrk,
}

#[derive(Clone, Copy, Debug)]
struct CellEdit {
    selector: litchi_xls::cell_values::Selector<'static>,
    reference: litchi_xls::cell_values::Reference,
    replacement: f64,
    storage: litchi_xls::cell_values::Storage,
}

fn family_for(case: Case) -> Result<(bool, Family), Box<dyn std::error::Error>> {
    match case {
        Case::XlsNumericEagerNumberEditSave => Ok((false, Family::Number)),
        Case::XlsNumericSourceBackedNumberEditSave => Ok((true, Family::Number)),
        Case::XlsNumericEagerRkMulrkEditSave => Ok((false, Family::RkMulrk)),
        Case::XlsNumericSourceBackedRkMulrkEditSave => Ok((true, Family::RkMulrk)),
        _ => Err("non-XLS-numeric case passed to numeric runner".into()),
    }
}

fn edits_for(
    snapshot: &litchi_xls::cell_values::Snapshot,
    family: Family,
) -> Result<Vec<CellEdit>, Box<dyn std::error::Error>> {
    use litchi_xls::cell_values::{Reference, Selector, Storage, Value};

    match family {
        Family::Number => {
            let reference = Reference::new(20, 4)?;
            let sheet = snapshot
                .worksheet(Selector::Name("Untouched"))?
                .ok_or("XLS Number corpus lost Untouched worksheet")?;
            let cell = sheet
                .cell(reference)?
                .ok_or("XLS Number corpus lost Untouched!E21")?;
            let Value::Number(source) = cell.value() else {
                return Err("XLS Number corpus target is not numeric".into());
            };
            if cell.storage() != Storage::Number || source.to_bits() != 42.0_f64.to_bits() {
                return Err("XLS Number corpus target does not retain Number/42".into());
            }
            let mut edits = Vec::new();
            reserve_exact(&mut edits, 1, "XLS Number edit inventory")?;
            push_bounded(
                &mut edits,
                CellEdit {
                    selector: Selector::Name("Untouched"),
                    reference,
                    replacement: 43.0,
                    storage: Storage::Number,
                },
                1,
                "XLS Number edit inventory",
            )?;
            Ok(edits)
        },
        Family::RkMulrk => {
            let sheet = snapshot
                .worksheet(Selector::Name("Packed"))?
                .ok_or("XLS RK/MulRK corpus lost Packed worksheet")?;
            let values = [(0, 0, 2.0), (1, 0, 5.0), (1, 1, 6.0)];
            let mut edits = Vec::new();
            reserve_exact(&mut edits, values.len(), "XLS RK/MulRK edit inventory")?;
            for (row, column, replacement) in values {
                let reference = Reference::new(row, column)?;
                let cell = sheet
                    .cell(reference)?
                    .ok_or("XLS RK/MulRK corpus lost an expected cell")?;
                let Value::Number(_source) = cell.value() else {
                    return Err("XLS RK/MulRK corpus cell is not numeric".into());
                };
                let storage = cell.storage();
                if !matches!(storage, Storage::Rk | Storage::MulRk) {
                    return Err("XLS RK/MulRK corpus cell family changed during preparation".into());
                }
                push_bounded(
                    &mut edits,
                    CellEdit {
                        selector: Selector::Name("Packed"),
                        reference,
                        replacement,
                        storage,
                    },
                    values.len(),
                    "XLS RK/MulRK edit inventory",
                )?;
            }
            Ok(edits)
        },
    }
}

fn stage(
    transaction: &mut litchi_xls::cell_values::Transaction,
    edits: &[CellEdit],
    family: Family,
) -> Result<(), Box<dyn std::error::Error>> {
    for edit in edits {
        match family {
            Family::Number => {
                transaction.set_number(edit.selector, edit.reference, edit.replacement)?
            },
            Family::RkMulrk => {
                transaction.set_numeric(edit.selector, edit.reference, edit.replacement)?
            },
        }
    }
    Ok(())
}

fn read_owned(source: &[u8]) -> Result<(Vec<u8>, SourceSnapshot), Box<dyn std::error::Error>> {
    let instrumented =
        InstrumentedSource::new(clone_bytes(source, "XLS numeric owned source")?, Vec::new());
    let mut bytes = Vec::new();
    reserve_exact(&mut bytes, source.len(), "XLS numeric source read buffer")?;
    bytes.resize(source.len(), 0_u8);
    let mut offset = 0_u64;
    for chunk in bytes.chunks_mut(64 * 1024) {
        instrumented.read_exact_at(offset, chunk)?;
        offset = offset
            .checked_add(u64::try_from(chunk.len())?)
            .ok_or("XLS numeric source offset overflow")?;
    }
    Ok((bytes, instrumented.snapshot()))
}

enum Publication {
    Eager(Box<litchi_xls::cell_values::Commit>),
    SourceBacked(Box<litchi_xls::cell_values::SourceBackedCommit>),
}

impl Publication {
    fn snapshot(&self) -> &litchi_xls::cell_values::Snapshot {
        match self {
            Self::Eager(commit) => commit.snapshot(),
            Self::SourceBacked(commit) => commit.snapshot(),
        }
    }

    fn patch(&self) -> &litchi_xls::cell_values::Patch {
        match self {
            Self::Eager(commit) => commit.patch(),
            Self::SourceBacked(commit) => commit.patch(),
        }
    }

    fn diagnostics(&self, source_workbook_bytes: u64) -> NumericDiagnostics {
        match self {
            Self::Eager(commit) => NumericDiagnostics {
                source_bytes: u64::try_from(commit.patch().before().len()).unwrap_or(0),
                source_workbook_bytes,
                target_workbook_bytes: u64::try_from(commit.snapshot().workbook_stream().len())
                    .unwrap_or(0),
                splice_count: None,
                replacement_bytes: None,
                changed_spans: None,
                source_fingerprint: None,
                target_fingerprint: None,
            },
            Self::SourceBacked(commit) => {
                let diagnostics = commit.diagnostics();
                NumericDiagnostics {
                    source_bytes: diagnostics.source_bytes(),
                    source_workbook_bytes: diagnostics.source_workbook_bytes(),
                    target_workbook_bytes: diagnostics.target_workbook_bytes(),
                    splice_count: Some(diagnostics.splice_count()),
                    replacement_bytes: Some(diagnostics.replacement_bytes()),
                    changed_spans: Some(diagnostics.changed_spans()),
                    source_fingerprint: Some(fingerprint_hex(
                        diagnostics.source_fingerprint().as_bytes(),
                    )),
                    target_fingerprint: Some(fingerprint_hex(
                        diagnostics.target_fingerprint().as_bytes(),
                    )),
                }
            },
        }
    }

    fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> Result<Option<PublishReport>, Box<dyn std::error::Error>> {
        match self {
            Self::Eager(commit) => {
                for chunk in commit.snapshot().bytes().chunks(64 * 1024) {
                    writer.write_all(chunk)?;
                }
                writer.flush()?;
                Ok(None)
            },
            Self::SourceBacked(commit) => Ok(Some(commit.write_to(writer)?)),
        }
    }
}

#[derive(Clone, Debug, Default, Serialize)]
struct NumericDiagnostics {
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    splice_count: Option<usize>,
    replacement_bytes: Option<u64>,
    changed_spans: Option<usize>,
    source_fingerprint: Option<String>,
    target_fingerprint: Option<String>,
}

fn prepare(
    snapshot: litchi_xls::cell_values::Snapshot,
    source_backed: bool,
    family: Family,
    edits: &[CellEdit],
) -> Result<Publication, Box<dyn std::error::Error>> {
    let mut transaction = snapshot.edit();
    stage(&mut transaction, edits, family)?;
    if source_backed {
        Ok(Publication::SourceBacked(Box::new(
            transaction.commit_source_backed()?,
        )))
    } else {
        Ok(Publication::Eager(Box::new(transaction.commit()?)))
    }
}

fn fingerprint_hex(bytes: &[u8; 32]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn verify_noop_and_security(
    corpus: &Corpus,
    family: Family,
    edits: &[CellEdit],
) -> Result<(), Box<dyn std::error::Error>> {
    use litchi_xls::cell_values::{Snapshot, Value};

    let source = Snapshot::from_bytes(clone_bytes(&corpus.archive, "XLS numeric no-op source")?)?;
    let noop = source.transaction().commit_source_backed()?;
    if !noop.is_noop() || noop.snapshot().bytes() != source.bytes() {
        return Err("XLS numeric source-backed no-op changed source identity".into());
    }
    let mut noop_bytes = Vec::new();
    let report = noop.write_to(&mut noop_bytes)?;
    if noop_bytes != corpus.archive
        || report.bytes() != u64::try_from(corpus.archive.len())?
        || report.changed_spans() != 0
        || report.source_fingerprint() != report.target_fingerprint()
    {
        return Err("XLS numeric no-op fingerprint/publication gate failed".into());
    }

    let mut unsupported = source.edit();
    let staged = unsupported
        .set_value(
            edits[0].selector,
            edits[0].reference,
            Value::Text("unsupported numeric family".to_owned()),
        )
        .is_ok();
    if staged {
        if unsupported.commit_source_backed().is_ok() {
            return Err("XLS numeric unsupported edit was accepted".into());
        }
    } else {
        let noop_after_refusal = unsupported.commit_source_backed()?;
        if !noop_after_refusal.is_noop() {
            return Err("XLS numeric unsupported edit was not failure-atomic".into());
        }
    }
    let mut bad_family = source.edit();
    if family == Family::Number
        && bad_family
            .set_numeric(edits[0].selector, edits[0].reference, edits[0].replacement)
            .is_err()
    {
        return Err("XLS Number set_numeric unexpectedly refused Number".into());
    }
    if family == Family::RkMulrk {
        let mut wrong = source.edit();
        if wrong
            .set_number(edits[0].selector, edits[0].reference, edits[0].replacement)
            .is_ok()
        {
            return Err("XLS RK/MulRK set_number changed a compressed family".into());
        }
    }

    let protected = protected_fixture()?;
    verify_security_refusal(&protected, "protected", SecurityFixtureExpectation::Open)?;
    let signed = source_with_stream(corpus, &["DigitalSignature"], b"signature")?;
    verify_security_refusal(
        &signed,
        "signed",
        SecurityFixtureExpectation::RefuseSignedOpen,
    )?;
    let macro_source = macro_fixture(corpus)?;
    verify_security_refusal(&macro_source, "macro", SecurityFixtureExpectation::Open)?;
    let encrypted = encrypted_fixture()?;
    verify_security_refusal(
        &encrypted,
        "encrypted",
        SecurityFixtureExpectation::RefuseEncryptedOpen,
    )?;
    Ok(())
}

#[derive(Clone, Copy)]
enum SecurityFixtureExpectation {
    Open,
    // Signed and encrypted CFBs are rejected before a Snapshot is exposed by
    // the public cell-values contract; each variant independently validates
    // its marker before requiring the typed refusal below.
    RefuseSignedOpen,
    RefuseEncryptedOpen,
}

fn verify_security_refusal(
    bytes: &[u8],
    label: &str,
    expectation: SecurityFixtureExpectation,
) -> Result<(), Box<dyn std::error::Error>> {
    use litchi_xls::cell_values::{Selector, Snapshot, Storage, Value};

    match expectation {
        SecurityFixtureExpectation::Open => {},
        SecurityFixtureExpectation::RefuseSignedOpen => {
            require_signed_marker(bytes)?;
        },
        SecurityFixtureExpectation::RefuseEncryptedOpen => {
            require_encrypted_marker(bytes)?;
        },
    }
    let snapshot = match Snapshot::from_bytes(clone_bytes(bytes, "XLS security fixture")?) {
        Ok(snapshot) => {
            if !matches!(expectation, SecurityFixtureExpectation::Open) {
                return Err(format!("XLS {label} fixture unexpectedly opened").into());
            }
            snapshot
        },
        Err(error) => {
            if matches!(expectation, SecurityFixtureExpectation::Open) {
                return Err(format!("XLS {label} fixture failed to open: {error}").into());
            }
            return match expectation {
                SecurityFixtureExpectation::RefuseSignedOpen => {
                    require_signed_refusal(error, label)
                },
                SecurityFixtureExpectation::RefuseEncryptedOpen => {
                    require_encrypted_refusal(error, label)
                },
                SecurityFixtureExpectation::Open => {
                    Err(format!("XLS {label} fixture failed to open: {error}").into())
                },
            };
        },
    };
    let Some((sheet, reference, storage, value)) = snapshot.worksheets().find_map(|sheet| {
        sheet.cells().find_map(|cell| {
            let Value::Number(value) = cell.value() else {
                return None;
            };
            matches!(
                cell.storage(),
                Storage::Number | Storage::Rk | Storage::MulRk
            )
            .then_some((sheet.position(), cell.reference(), cell.storage(), value))
        })
    }) else {
        return Err(format!("XLS {label} fixture has no numeric guard cell").into());
    };
    let mut edit = snapshot.edit();
    let staged = match storage {
        Storage::Number => edit.set_number(Selector::Position(sheet), reference, value + 1.0),
        Storage::Rk | Storage::MulRk => {
            edit.set_numeric(Selector::Position(sheet), reference, value + 1.0)
        },
        _ => {
            return Err(format!("XLS {label} numeric guard selected an unsupported family").into());
        },
    };
    if staged.is_ok() {
        if edit.commit_source_backed().is_ok() {
            return Err(format!("XLS {label} source-backed edit was accepted").into());
        }
    } else if !edit.commit_source_backed()?.is_noop() {
        return Err(format!("XLS {label} refusal was not failure-atomic").into());
    }
    Ok(())
}

const EXPECTED_PROTECTED_REFUSAL: &str =
    "signed, encrypted, or DRM containers are not eligible for object editing";

fn require_signed_refusal(
    error: litchi_xls::Error,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    match error {
        litchi_xls::Error::Cfb(litchi_cfb::OleError::InvalidFormat(message))
            if message == EXPECTED_PROTECTED_REFUSAL =>
        {
            Ok(())
        },
        other => Err(format!(
            "XLS {label} fixture returned an unexpected security refusal: {other:?}"
        )
        .into()),
    }
}

fn require_encrypted_refusal(
    error: litchi_xls::Error,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    match error {
        litchi_xls::Error::PasswordRequired => Ok(()),
        other => Err(format!(
            "XLS {label} fixture returned an unexpected encryption refusal: {other:?}"
        )
        .into()),
    }
}

fn require_signed_marker(bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
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
        return Err("XLS signed fixture lacks its DigitalSignature stream marker".into());
    };
    let references = signed_path.iter().map(String::as_str).collect::<Vec<_>>();
    if ole.open_stream(&references)?.is_empty() {
        return Err("XLS signed fixture has an empty DigitalSignature marker".into());
    }
    Ok(())
}

fn require_encrypted_marker(bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    let workbook = ole
        .open_stream(&["Workbook"])
        .or_else(|_| ole.open_stream(&["Book"]))?;
    let mut offset = 0usize;
    while offset < workbook.len() {
        let header_end = offset
            .checked_add(4)
            .ok_or("XLS encrypted marker header offset overflow")?;
        let header = workbook
            .get(offset..header_end)
            .ok_or("XLS encrypted marker has a truncated BIFF header")?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = header_end
            .checked_add(payload_len)
            .ok_or("XLS encrypted marker record length overflow")?;
        let payload = workbook
            .get(header_end..end)
            .ok_or("XLS encrypted marker has a truncated BIFF payload")?;
        if kind == 0x002f && payload.len() >= 2 {
            return Ok(());
        }
        offset = end;
    }
    Err("XLS encrypted fixture lacks a BIFF FILEPASS marker".into())
}

fn protected_fixture() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = litchi_xls::writer::Writer::new();
    let sheet = writer.add_worksheet("Protected")?;
    writer.write_number(sheet, 0, 0, 1.0)?;
    writer.protect_sheet(sheet, Some("password"), true, false)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn encrypted_fixture() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    use litchi_xls::WeakEncryptionPolicy;
    use litchi_xls::writer::{EncryptionProfile, Writer};

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Encrypted")?;
    writer.write_number(sheet, 0, 0, 1.0)?;
    writer.set_xor_obfuscation_password("legacy", WeakEncryptionPolicy::allow_xor_obfuscation())?;
    if writer.encryption_profile() != Some(EncryptionProfile::XorObfuscation) {
        return Err("XLS encrypted guard fixture did not retain its encryption profile".into());
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn source_with_stream(
    corpus: &Corpus,
    path: &[&str],
    payload: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut source = OleFile::open(Cursor::new(corpus.archive.as_slice()))?;
    let mut writer = OleWriter::new();
    let mut streams = source.list_streams();
    streams.sort();
    for stream_path in streams {
        let references = stream_references(&stream_path)?;
        writer.create_stream_owned(&references, source.open_stream(&references)?)?;
    }
    writer.create_stream(path, payload)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn macro_fixture(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut source = OleFile::open(Cursor::new(corpus.archive.as_slice()))?;
    let mut writer = OleWriter::new();
    let mut streams = source.list_streams();
    streams.sort();
    for stream_path in streams {
        let references = stream_references(&stream_path)?;
        writer.create_stream_owned(&references, source.open_stream(&references)?)?;
    }
    writer.create_storage(&["_VBA_PROJECT_CUR"])?;
    writer.create_storage(&["_VBA_PROJECT_CUR", "VBA"])?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "_VBA_PROJECT"], b"project")?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "dir"], b"dir")?;
    writer.create_stream(&["_VBA_PROJECT_CUR", "VBA", "Module1"], b"module")?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn stream_references(path: &[String]) -> Result<Vec<&str>, Box<dyn std::error::Error>> {
    let mut references = Vec::new();
    reserve_exact(&mut references, path.len(), "XLS CFB stream references")?;
    for component in path {
        push_bounded(
            &mut references,
            component.as_str(),
            path.len(),
            "XLS CFB stream references",
        )?;
    }
    Ok(references)
}

fn verify_output(
    corpus: &Corpus,
    output: &[u8],
    family: Family,
    edits: &[CellEdit],
    source_backed: bool,
) -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::sheet::Cell as _;
    use litchi_xls::cell_values::{Snapshot, Storage, Value};

    super::verify_xls_untouched_streams(&corpus.archive, output, source_backed)?;
    let snapshot = Snapshot::from_bytes(clone_bytes(output, "XLS numeric output reopen")?)?;
    for edit in edits {
        let sheet = snapshot
            .worksheet(edit.selector)?
            .ok_or("XLS numeric output lost selected worksheet")?;
        let cell = sheet
            .cell(edit.reference)?
            .ok_or("XLS numeric output lost edited cell")?;
        let Value::Number(value) = cell.value() else {
            return Err("XLS numeric output edited cell is no longer numeric".into());
        };
        if value.to_bits() != edit.replacement.to_bits() || cell.storage() != edit.storage {
            return Err("XLS numeric output semantic family/value differs from expectation".into());
        }
    }
    if family == Family::RkMulrk
        && snapshot
            .worksheet(litchi_xls::cell_values::Selector::Name("Packed"))?
            .ok_or("XLS RK/MulRK output lost Packed worksheet")?
            .cells()
            .filter(|cell| matches!(cell.storage(), Storage::Rk | Storage::MulRk))
            .count()
            != 3
    {
        return Err("XLS RK/MulRK output family inventory changed".into());
    }

    let workbook = litchi_xls::Workbook::new(Cursor::new(output))?;
    let sheet = workbook
        .sheets()
        .iter()
        .find(|sheet| sheet.name() == "Untouched" || sheet.name() == "Packed")
        .ok_or("XLS numeric Workbook reopen lost selected worksheet")?;
    let worksheet = workbook.xls_worksheet(
        sheet
            .parsed_worksheet_index()
            .ok_or("XLS numeric Workbook selected tab is not a worksheet")?,
    )?;
    for edit in edits {
        let cell = worksheet
            .get_cell(
                u32::from(edit.reference.row()),
                u32::from(edit.reference.column()),
            )
            .ok_or("XLS numeric Workbook reopen lost edited cell")?;
        let value = cell
            .value()
            .as_float()
            .ok_or("XLS numeric Workbook reopen edited cell is not numeric")?;
        if value.to_bits() != edit.replacement.to_bits() {
            return Err("XLS numeric Workbook readback differs from expectation".into());
        }
    }
    Ok(())
}

fn verify_patch(
    source: &litchi_xls::cell_values::Snapshot,
    publication: &Publication,
) -> Result<(), Box<dyn std::error::Error>> {
    let applied = publication.patch().apply(source)?;
    if applied.bytes() != publication.snapshot().bytes()
        || publication.patch().apply(publication.snapshot()).is_ok()
    {
        return Err("XLS numeric patch source/stale precondition gate failed".into());
    }
    let restored = publication.patch().inverse().apply(&applied)?;
    if restored.bytes() != source.bytes() {
        return Err("XLS numeric patch inverse did not restore source".into());
    }
    Ok(())
}

#[derive(Debug)]
struct PrefixSink {
    accepted: usize,
    limit: usize,
}

impl Write for PrefixSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("intentional XLS numeric sink failure"));
        }
        let accepted = bytes.len().min(self.limit - self.accepted);
        self.accepted = self.accepted.saturating_add(accepted);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn verify_partial_publication(
    publication: &Publication,
    output_bytes: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut sink = PrefixSink {
        accepted: 0,
        limit: output_bytes / 2,
    };
    if publication.write_to(&mut sink).is_ok()
        || sink.accepted == 0
        || sink.accepted >= output_bytes
    {
        return Err("XLS numeric partial sink gate did not retain a typed prefix".into());
    }
    Ok(())
}

fn verify_real_producer() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_xls::cell_values::{Selector, Snapshot, Storage, Value};

    let source = Snapshot::from_bytes(clone_bytes(REAL_PRODUCER, "XLS real-producer source")?)?;
    let mut selected = None;
    for sheet in source.worksheets() {
        for cell in sheet.cells() {
            let Value::Number(value) = cell.value() else {
                continue;
            };
            if !matches!(cell.storage(), Storage::Rk | Storage::MulRk) {
                continue;
            }
            let replacement = [*value + 1.0, *value - 1.0].into_iter().find(|candidate| {
                let mut edit = source.edit();
                edit.set_numeric(
                    Selector::Position(sheet.position()),
                    cell.reference(),
                    *candidate,
                )
                .is_ok()
            });
            if let Some(replacement) = replacement {
                selected = Some((sheet.position(), cell.reference(), replacement));
                break;
            }
        }
        if selected.is_some() {
            break;
        }
    }
    let (sheet, reference, replacement) =
        selected.ok_or("XLS real-producer fixture has no representable RK/MulRK replacement")?;
    let mut edit = source.edit();
    edit.set_numeric(Selector::Position(sheet), reference, replacement)?;
    let commit = edit.commit_source_backed()?;
    let mut output = Vec::new();
    commit.write_to(&mut output)?;
    let reopened = Snapshot::from_bytes(output)?;
    let cell = reopened
        .worksheet(Selector::Position(sheet))?
        .ok_or("XLS real-producer output lost worksheet")?
        .cell(reference)?
        .ok_or("XLS real-producer output lost cell")?;
    if cell.value() != &Value::Number(replacement)
        || commit.patch().inverse().apply(commit.snapshot())?.bytes() != source.bytes()
    {
        return Err("XLS real-producer RK/MulRK reopen/inverse gate failed".into());
    }
    Ok(())
}

fn prepare_expected(
    corpus: &Corpus,
    source_backed: bool,
    family: Family,
) -> Result<Expected, Box<dyn std::error::Error>> {
    let source = litchi_xls::cell_values::Snapshot::from_bytes(clone_bytes(
        &corpus.archive,
        "XLS numeric expected source",
    )?)?;
    let edits = edits_for(&source, family)?;
    verify_noop_and_security(corpus, family, &edits)?;
    let publication = prepare(source.clone(), source_backed, family, &edits)?;
    let mut output = Vec::new();
    let report = publication.write_to(&mut output)?;
    if output == corpus.archive {
        return Err("XLS numeric expected publication did not change source".into());
    }
    verify_output(corpus, &output, family, &edits, source_backed)?;
    verify_patch(&source, &publication)?;
    verify_partial_publication(&publication, output.len())?;
    if let Some(report) = report
        && (report.bytes() != u64::try_from(output.len())? || report.changed_spans() == 0)
    {
        return Err("XLS numeric source-backed publication report is incomplete".into());
    }
    let target_workbook_bytes = u64::try_from(publication.snapshot().workbook_stream().len())?;
    Ok(Expected {
        source,
        edits,
        output_digest: sha256_hex(&output),
        output_bytes: output,
        target_workbook_bytes,
    })
}

struct Expected {
    source: litchi_xls::cell_values::Snapshot,
    edits: Vec<CellEdit>,
    output_digest: String,
    output_bytes: Vec<u8>,
    target_workbook_bytes: u64,
}

/// Runs one of the four opt-in selectors.
pub(crate) fn run(
    case: Case,
    number_corpus: &Corpus,
    rk_mulrk_corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn std::error::Error>> {
    let (source_backed, family) = family_for(case)?;
    let corpus = match family {
        Family::Number => number_corpus,
        Family::RkMulrk => rk_mulrk_corpus,
    };
    if family == Family::Number
        && corpus.manifest.generator != super::XLS_COMMENTS_EDIT_CORPUS_GENERATOR
    {
        return Err("XLS Number selector requires the deterministic comments corpus".into());
    }
    if family == Family::RkMulrk && corpus.manifest.generator != CORPUS_GENERATOR {
        return Err("XLS RK/MulRK selector requires its fixed native corpus".into());
    }

    // All guards, source ingress, expected outputs, and the untimed
    // real-producer reopen/inverse gate happen before any measured sample.
    verify_real_producer()?;
    let expected = prepare_expected(corpus, source_backed, family)?;
    let maximum = u64::try_from(expected.output_bytes.len())?;
    let source_workbook_bytes = u64::try_from(expected.source.workbook_stream().len())?;
    let mut elapsed = Vec::new();
    reserve_exact(&mut elapsed, samples, "XLS numeric total timing")?;
    let mut source_read_calls = Vec::new();
    reserve_exact(
        &mut source_read_calls,
        samples,
        "XLS numeric source read-call evidence",
    )?;
    let mut source_read_bytes = Vec::new();
    reserve_exact(
        &mut source_read_bytes,
        samples,
        "XLS numeric source read-byte evidence",
    )?;
    let mut source_summary = XlsNumericSourceSummary {
        source_counter_scope: "owned-source-ingress-only",
        implementation: if source_backed {
            "source_backed"
        } else {
            "eager"
        },
        family: match family {
            Family::Number => "Number",
            Family::RkMulrk => "RK+MulRK",
        },
        source_backed,
        update_count: expected.edits.len(),
        input_cfb_bytes: u64::try_from(corpus.archive.len())?,
        output_cfb_bytes: u64::try_from(expected.output_bytes.len())?,
        source_workbook_bytes,
        target_workbook_bytes: expected.target_workbook_bytes,
        sink_capacity_bytes: maximum,
        expected_output_sha256: expected.output_digest.clone(),
        owned_input_scope: "complete in-memory CFB bytes; no positional/physical I/O",
        splice_count: source_backed.then(Vec::new),
        replacement_bytes: source_backed.then(Vec::new),
        changed_spans: source_backed.then(Vec::new),
        source_fingerprints: source_backed.then(Vec::new),
        target_fingerprints: source_backed.then(Vec::new),
        ..XlsNumericSourceSummary::default()
    };
    reserve_summary(&mut source_summary, samples)?;
    let mut sink_summaries = Vec::new();
    reserve_exact(&mut sink_summaries, samples, "XLS numeric sink summaries")?;
    let mut measured_digests = Vec::new();
    reserve_exact(
        &mut measured_digests,
        samples,
        "XLS numeric measured digests",
    )?;

    for iteration in 0..super::iteration_count(warmup_iterations, samples)? {
        let (owned, source_metrics) = read_owned(&corpus.archive)?;
        let snapshot = litchi_xls::cell_values::Snapshot::from_bytes(owned)?;
        let mut sink = CountingSink::bounded(maximum, MAX_WRITE);
        sink.reserve_budget()?;

        let edit_started = Instant::now();
        let mut transaction = snapshot.edit();
        let edit_duration = edit_started.elapsed();
        let set_started = Instant::now();
        stage(&mut transaction, &expected.edits, family)?;
        let set_duration = set_started.elapsed();
        let commit_started = Instant::now();
        let publication = if source_backed {
            Publication::SourceBacked(Box::new(transaction.commit_source_backed()?))
        } else {
            Publication::Eager(Box::new(transaction.commit()?))
        };
        let commit_duration = commit_started.elapsed();
        let publication_started = Instant::now();
        let report = publication.write_to(&mut sink)?;
        let publication_duration = publication_started.elapsed();
        let total = edit_duration
            .checked_add(set_duration)
            .and_then(|duration| duration.checked_add(commit_duration))
            .and_then(|duration| duration.checked_add(publication_duration))
            .ok_or("XLS numeric sample duration overflow")?;
        let metrics = source_metrics;
        let diagnostics = publication.diagnostics(source_workbook_bytes);

        if metrics.read_calls == 0
            || metrics.read_bytes != u64::try_from(corpus.archive.len())?
            || sink.bytes != expected.output_bytes
            || sink.summary().accepted_bytes != u64::try_from(expected.output_bytes.len())?
            || diagnostics.target_workbook_bytes
                != u64::try_from(publication.snapshot().workbook_stream().len())?
        {
            return Err(
                "XLS numeric sample source/publication evidence differs from expected".into(),
            );
        }
        if source_backed {
            let report = report.ok_or("XLS numeric source-backed publication has no report")?;
            let expected_replacement_bytes = match family {
                Family::Number => 8,
                Family::RkMulrk => 4 * expected.edits.len(),
            };
            if diagnostics.splice_count != Some(expected.edits.len())
                || diagnostics.replacement_bytes != Some(u64::try_from(expected_replacement_bytes)?)
                || diagnostics.changed_spans != Some(report.changed_spans())
                || report.bytes() != sink.summary().accepted_bytes
                || diagnostics.source_fingerprint
                    != Some(fingerprint_hex(report.source_fingerprint().as_bytes()))
                || diagnostics.target_fingerprint
                    != Some(fingerprint_hex(report.target_fingerprint().as_bytes()))
            {
                return Err(
                    "XLS numeric source-backed diagnostics disagree with publication".into(),
                );
            }
        } else if report.is_some() {
            return Err("XLS numeric eager publication reported overlay-only evidence".into());
        }
        verify_output(corpus, &sink.bytes, family, &expected.edits, source_backed)?;
        verify_patch(&expected.source, &publication)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected.output_digest {
            return Err("XLS numeric measured output digest is unstable".into());
        }

        if iteration >= warmup_iterations {
            let commit_elapsed = super::elapsed_ns(commit_duration)?;
            let publication_elapsed = super::elapsed_ns(publication_duration)?;
            let edit_elapsed = super::elapsed_ns(edit_duration)?;
            let set_elapsed = super::elapsed_ns(set_duration)?;
            let total_elapsed = super::elapsed_ns(total)?;
            push_bounded(
                &mut elapsed,
                total_elapsed,
                samples,
                "XLS numeric total timing",
            )?;
            push_bounded(
                &mut source_read_calls,
                metrics.read_calls,
                samples,
                "XLS numeric source read-call evidence",
            )?;
            push_bounded(
                &mut source_read_bytes,
                metrics.read_bytes,
                samples,
                "XLS numeric source read-byte evidence",
            )?;
            push_bounded(
                &mut source_summary.commit_ns,
                commit_elapsed,
                samples,
                "XLS numeric commit timing evidence",
            )?;
            push_bounded(
                &mut source_summary.publication_ns,
                publication_elapsed,
                samples,
                "XLS numeric publication timing evidence",
            )?;
            push_bounded(
                &mut source_summary.edit_ns,
                edit_elapsed,
                samples,
                "XLS numeric edit timing evidence",
            )?;
            push_bounded(
                &mut source_summary.set_ns,
                set_elapsed,
                samples,
                "XLS numeric set timing evidence",
            )?;
            push_bounded(
                &mut source_summary.total_ns,
                total_elapsed,
                samples,
                "XLS numeric total timing evidence",
            )?;
            push_bounded(
                &mut source_summary.complete_target_materialized_bytes,
                u64::try_from(publication.snapshot().bytes().len())?,
                samples,
                "XLS numeric materialized-target evidence",
            )?;
            push_bounded(
                &mut source_summary.sink_bytes,
                sink.summary().accepted_bytes,
                samples,
                "XLS numeric sink-byte evidence",
            )?;
            push_bounded(
                &mut source_summary.sink_write_calls,
                sink.summary().write_calls,
                samples,
                "XLS numeric sink-write evidence",
            )?;
            push_bounded(
                &mut source_summary.sink_digests,
                digest.clone(),
                samples,
                "XLS numeric sink-digest evidence",
            )?;
            push_bounded(
                &mut source_summary.source_bytes,
                diagnostics.source_bytes,
                samples,
                "XLS numeric source-size evidence",
            )?;
            push_bounded(
                &mut source_summary.source_workbook_bytes_per_sample,
                diagnostics.source_workbook_bytes,
                samples,
                "XLS numeric source-workbook evidence",
            )?;
            push_bounded(
                &mut source_summary.target_workbook_bytes_per_sample,
                diagnostics.target_workbook_bytes,
                samples,
                "XLS numeric target-workbook evidence",
            )?;
            if let Some(value) = diagnostics.splice_count {
                let values = source_summary
                    .splice_count
                    .as_mut()
                    .ok_or("XLS numeric eager path produced splice evidence")?;
                push_bounded(values, value, samples, "XLS numeric splice evidence")?;
            }
            if let Some(value) = diagnostics.replacement_bytes {
                let values = source_summary
                    .replacement_bytes
                    .as_mut()
                    .ok_or("XLS numeric eager path produced replacement evidence")?;
                push_bounded(
                    values,
                    value,
                    samples,
                    "XLS numeric replacement-byte evidence",
                )?;
            }
            if let Some(value) = diagnostics.changed_spans {
                let values = source_summary
                    .changed_spans
                    .as_mut()
                    .ok_or("XLS numeric eager path produced changed-span evidence")?;
                push_bounded(values, value, samples, "XLS numeric changed-span evidence")?;
            }
            if let Some(value) = diagnostics.source_fingerprint {
                let values = source_summary
                    .source_fingerprints
                    .as_mut()
                    .ok_or("XLS numeric eager path produced source-fingerprint evidence")?;
                push_bounded(
                    values,
                    value,
                    samples,
                    "XLS numeric source-fingerprint evidence",
                )?;
            }
            if let Some(value) = diagnostics.target_fingerprint {
                let values = source_summary
                    .target_fingerprints
                    .as_mut()
                    .ok_or("XLS numeric eager path produced target-fingerprint evidence")?;
                push_bounded(
                    values,
                    value,
                    samples,
                    "XLS numeric target-fingerprint evidence",
                )?;
            }
            push_bounded(
                &mut sink_summaries,
                sink.summary(),
                samples,
                "XLS numeric sink summaries",
            )?;
            push_bounded(
                &mut measured_digests,
                digest,
                samples,
                "XLS numeric measured digests",
            )?;
        }
        std::hint::black_box(&sink.bytes);
    }

    if measured_digests
        .iter()
        .any(|digest| digest != &expected.output_digest)
        || elapsed.len() != samples
        || source_read_calls.len() != samples
        || source_read_bytes.len() != samples
        || source_summary.edit_ns.len() != samples
        || source_summary.set_ns.len() != samples
        || source_summary.commit_ns.len() != samples
        || source_summary.publication_ns.len() != samples
        || source_summary.total_ns.len() != samples
        || source_summary.complete_target_materialized_bytes.len() != samples
        || source_summary.sink_bytes.len() != samples
        || source_summary.sink_write_calls.len() != samples
        || source_summary.sink_digests.len() != samples
        || source_summary.source_bytes.len() != samples
        || source_summary.source_workbook_bytes_per_sample.len() != samples
        || source_summary.target_workbook_bytes_per_sample.len() != samples
    {
        return Err("XLS numeric measured vectors are incomplete or nondeterministic".into());
    }
    let optional_lengths = [
        source_summary.splice_count.as_ref().map(Vec::len),
        source_summary.replacement_bytes.as_ref().map(Vec::len),
        source_summary.changed_spans.as_ref().map(Vec::len),
        source_summary.source_fingerprints.as_ref().map(Vec::len),
        source_summary.target_fingerprints.as_ref().map(Vec::len),
    ];
    if optional_lengths.iter().any(|length| {
        length.is_some() != source_backed || length.is_some_and(|length| length != samples)
    }) {
        return Err("XLS numeric optional evidence vectors are incomplete".into());
    }
    source_summary.sample_count = samples;
    let sink = super::deterministic_sink_summary(&sink_summaries, "XLS numeric publication")?;
    let source = SourceSummary {
        read_calls: source_read_calls,
        read_bytes: source_read_bytes,
        xls_numeric: Some(source_summary),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: super::statistics(elapsed),
        sink: Some(sink),
        source: Some(source),
        execution: None,
        output_sha256: Some(expected.output_digest),
        operation_metrics: None,
    })
}

#[cfg(test)]
mod tests {
    use super::{Family, build_rk_mulrk_corpus, family_for, run};
    use crate::{Case, build_xls_comments_edit_corpus, parse_case};

    #[test]
    fn selectors_are_opt_in_and_have_stable_families() {
        let cases = [
            (
                "xls_numeric_eager_number_edit_save",
                Case::XlsNumericEagerNumberEditSave,
                (false, Family::Number),
            ),
            (
                "xls_numeric_source_backed_number_edit_save",
                Case::XlsNumericSourceBackedNumberEditSave,
                (true, Family::Number),
            ),
            (
                "xls_numeric_eager_rk_mulrk_edit_save",
                Case::XlsNumericEagerRkMulrkEditSave,
                (false, Family::RkMulrk),
            ),
            (
                "xls_numeric_source_backed_rk_mulrk_edit_save",
                Case::XlsNumericSourceBackedRkMulrkEditSave,
                (true, Family::RkMulrk),
            ),
        ];
        for (name, case, expected) in cases {
            assert_eq!(parse_case(name), Some(case));
            assert_eq!(family_for(case).unwrap(), expected);
            assert!(!Case::DEFAULT.contains(&case));
        }
    }

    #[test]
    fn packed_corpus_is_deterministic_and_runs_one_sample() {
        let first = build_rk_mulrk_corpus().unwrap();
        let second = build_rk_mulrk_corpus().unwrap();
        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.archive_sha256,
            second.manifest.archive_sha256
        );
        assert_eq!(first.manifest.generator, super::CORPUS_GENERATOR);
        let number = build_xls_comments_edit_corpus().unwrap();
        let mut number_digest = None;
        let mut rk_mulrk_digest = None;
        for case in [
            Case::XlsNumericEagerNumberEditSave,
            Case::XlsNumericSourceBackedNumberEditSave,
            Case::XlsNumericEagerRkMulrkEditSave,
            Case::XlsNumericSourceBackedRkMulrkEditSave,
        ] {
            let result = run(case, &number, &first, 0, 1).unwrap();
            assert_eq!(result.elapsed_ns.samples.len(), 1);
            let digest = result.output_sha256.clone().unwrap();
            let evidence = result.source.unwrap().xls_numeric.unwrap();
            assert_eq!(evidence.sample_count, 1);
            assert_eq!(evidence.complete_target_materialized_bytes.len(), 1);
            assert_eq!(evidence.sink_digests.len(), 1);
            let previous = if matches!(
                case,
                Case::XlsNumericEagerNumberEditSave | Case::XlsNumericSourceBackedNumberEditSave
            ) {
                &mut number_digest
            } else {
                &mut rk_mulrk_digest
            };
            if let Some(previous) = previous {
                assert_eq!(previous, &digest);
            } else {
                *previous = Some(digest);
            }
        }
    }
}
