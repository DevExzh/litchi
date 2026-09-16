//! Deterministic differential harness for the ZIP read grammar (change 0611).
//!
//! This is change 0582's `strict_scope_differential.rs`, verbatim, plus two
//! extra per-member verdict APIs:
//!
//!   * `I.read_entry` -- `IndexedArchive::read_entry`, the ordinary indexed
//!     read path.  Change 0611 modifies exactly this path, and change 0582's
//!     harness never called it: its four member APIs all enter
//!     `strict_layout_for`, which the ordinary path does not.
//!   * `R.read` -- `ArchiveReader::read`, the ordinary slice-backed read path.
//!     A control: change 0611 does not touch the slice path, so every `R.read`
//!     verdict must be identical on both sides.
//!
//! Everything else -- the limit profiles, the four strict-layout APIs, the
//! `read_stored_borrowed` control, the ported fuzz-target body, the slice parse
//! record and the verdict line format -- is unchanged.
//!
//! This is **not** a fuzzer.  It is a fixed-corpus, single-threaded, fully
//! deterministic driver that records one canonical verdict line per
//! (input, profile, API, member) and writes them to a file.  Running it against
//! two builds of `soapberry-zip` and diffing the two reports is the whole
//! experiment.
//!
//! Build it as an example of `soapberry-zip` so it links the crate under test
//! without adding a dependency to any crate:
//!
//! ```text
//! cp read_grammar_differential.rs <tree>/crates/soapberry-zip/examples/
//! CARGO_TARGET_DIR=<dir> cargo run --release -p soapberry-zip \
//!     --example read_grammar_differential -- <corpus-dir> <report-path>
//! ```
//!
//! The body of `exercise_fuzz_target_body` is a port of
//! `crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs`.  What was kept, what
//! was changed and what was dropped is stated in
//! `docs/performance/0582-zip-strict-scope-differential-harness.md`.

use std::collections::BTreeSet;
use std::fmt::Write as _;
use std::hint::black_box;
use std::io::{self, Write};
use std::panic::{self, AssertUnwindSafe};
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use soapberry_zip::office::{ArchiveLimits, ArchiveReader, IndexedArchive};
use soapberry_zip::{
    CompressionMethod, ErrorKind, PreservationPlan, RECOMMENDED_BUFFER_SIZE, ReaderAt,
    RegeneratedEntry, ReplayLimits, ZipArchive, ZipOperationAccounting,
};

// ---------------------------------------------------------------------------
// Limits.  `fuzz` reproduces parse_zip.rs exactly.  `wide` exists because the
// fuzz limits refuse almost every real Office fixture at open, which would make
// the test-data half of the corpus prove nothing about the changed code.
// ---------------------------------------------------------------------------

const MAX_INPUT_BYTES: usize = 1 << 20;
const MAX_FILES: usize = 256;
const MAX_MEMBER_NAME_BYTES: u64 = 4 << 10;
const MAX_METADATA_BYTES: u64 = 64 << 10;
const MAX_ENTRY_BYTES: u64 = 1 << 20;
const MAX_PRECOMPRESSED_PROGRESS_EVENTS: usize = 8;
const MAX_REPUBLISHED_OUTPUT_BYTES: usize = 3 * MAX_INPUT_BYTES;
const MAX_REPUBLISHED_WRITE_CHUNK: usize = 4096;
const REPUBLISHED_MEMBER_NAME: &str = "__fuzz_precompressed_member__";
const MAX_REPLAY_PAYLOAD_BYTES: usize = 4096;

fn fuzz_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: MAX_FILES,
        max_member_name_bytes: MAX_MEMBER_NAME_BYTES,
        max_metadata_bytes: MAX_METADATA_BYTES,
        max_compressed_size: MAX_ENTRY_BYTES,
        max_entry_size: MAX_ENTRY_BYTES,
        max_total_size: MAX_ENTRY_BYTES,
    }
}

fn wide_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: 65_535,
        max_member_name_bytes: 64 << 10,
        max_metadata_bytes: 16 << 20,
        max_compressed_size: 256 << 20,
        max_entry_size: 256 << 20,
        max_total_size: 1 << 30,
    }
}

// ---------------------------------------------------------------------------
// Report sink
// ---------------------------------------------------------------------------

struct Report {
    out: String,
}

impl Report {
    fn line(&mut self, text: &str) {
        self.out.push_str(text);
        self.out.push('\n');
    }
}

/// A stable identity for a typed refusal.
///
/// `ErrorKind` derives `Debug` and carries its message and numeric fields, so
/// the debug rendering is the error identity for differential purposes.
fn error_id(error: &soapberry_zip::Error) -> String {
    escape(&format!("{:?}", error.kind()))
}

fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '\\' => out.push_str("\\\\"),
            c if (c as u32) < 0x20 || (c as u32) == 0x7f => {
                let _ = write!(out, "\\x{:02x}", c as u32);
            },
            c => out.push(c),
        }
    }
    out
}

/// A member name is printed as one whitespace-free token, so the comparator
/// can never merge two members into one key.
fn escape_name(text: &str) -> String {
    escape(text).replace(' ', "\\x20")
}

fn verdict<T>(result: Result<T, soapberry_zip::Error>, describe: impl Fn(&T) -> String) -> String {
    match result {
        Ok(value) => format!("ok({})", describe(&value)),
        Err(error) => format!("err({})", error_id(&error)),
    }
}

// ---------------------------------------------------------------------------
// Sinks
// ---------------------------------------------------------------------------

/// Counts and fingerprints without retaining, so a large member costs no heap.
struct DigestSink {
    len: u64,
    hash: u64,
    chunks: Vec<u8>,
}

impl Default for DigestSink {
    fn default() -> Self {
        Self {
            len: 0,
            hash: 0xcbf2_9ce4_8422_2325,
            chunks: Vec::new(),
        }
    }
}

impl Write for DigestSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.len += input.len() as u64;
        // FNV-1a over the concatenation.  `crc32` of a growing prefix is not
        // composable through the public API, so the whole member is hashed here
        // and a bounded prefix is kept for a second, independent fingerprint.
        if self.chunks.len() < 4096 {
            let take = (4096 - self.chunks.len()).min(input.len());
            self.chunks.extend_from_slice(&input[..take]);
        }
        for byte in input {
            self.hash ^= u64::from(*byte);
            self.hash = self.hash.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl DigestSink {
    fn describe(&self) -> String {
        format!(
            "len={} fnv={:016x} head={:08x}",
            self.len,
            self.hash,
            soapberry_zip::crc32(&self.chunks)
        )
    }
}

/// The fuzz target's bounded sink, ported verbatim.
#[derive(Debug)]
struct BoundedSink {
    bytes: Vec<u8>,
    max_bytes: usize,
    max_write: usize,
    fail_after: Option<usize>,
}

impl BoundedSink {
    fn new(max_bytes: usize, max_write: usize, fail_after: Option<usize>) -> Self {
        assert!(max_bytes > 0);
        assert!(max_write > 0);
        Self {
            bytes: Vec::new(),
            max_bytes,
            max_write,
            fail_after,
        }
    }
}

impl Write for BoundedSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        if self.bytes.len() >= self.max_bytes {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "fuzz output limit reached",
            ));
        }
        if self
            .fail_after
            .is_some_and(|fail_after| self.bytes.len() >= fail_after)
        {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "fuzz partial sink failure",
            ));
        }

        let mut accepted = input.len().min(self.max_write);
        accepted = accepted.min(self.max_bytes - self.bytes.len());
        if let Some(fail_after) = self.fail_after {
            accepted = accepted.min(fail_after.saturating_sub(self.bytes.len()));
        }
        if accepted == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "fuzz partial sink failure",
            ));
        }
        self.bytes.extend_from_slice(&input[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// The strict-layout sweep.  This is the part the fuzz target does not do and
// the part change 0580 actually changes.
// ---------------------------------------------------------------------------

fn sorted_names(reader: &ArchiveReader<'_>) -> Vec<String> {
    let mut names: Vec<String> = reader.file_names().map(|name| name.to_string()).collect();
    names.sort();
    names.dedup();
    names
}

fn sorted_names_indexed<R: ReaderAt>(archive: &IndexedArchive<R>) -> Vec<String> {
    let mut names: Vec<String> = archive.file_names().map(|name| name.to_string()).collect();
    names.sort();
    names.dedup();
    names
}

/// `ArchiveReader::read_to`, the slice-backed strict-layout entry point.
fn sweep_slice_read_to(reader: &ArchiveReader<'_>, names: &[String]) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            let mut sink = DigestSink::default();
            match reader.read_to(name, &mut sink) {
                Ok(count) => format!("ok(count={count} {})", sink.describe()),
                Err(error) => format!("err({})", error_id(&error)),
            }
        })
        .collect()
}

/// `ArchiveReader::read`, the ordinary slice-backed read path.  Not a
/// strict-layout entry point: change 0611 leaves the slice path alone, so this
/// is a control.
fn sweep_slice_read(reader: &ArchiveReader<'_>, names: &[String]) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            verdict(reader.read(name), |payload| {
                format!(
                    "len={} crc={:08x}",
                    payload.len(),
                    soapberry_zip::crc32(payload)
                )
            })
        })
        .collect()
}

/// `IndexedArchive::read_entry_to`, the source-backed strict-layout entry point.
fn sweep_indexed_read_entry_to<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    names: &[String],
) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            let Some(entry_id) = archive.entry_id(name) else {
                return "no-entry-id".to_string();
            };
            let mut sink = DigestSink::default();
            match archive.read_entry_to(entry_id, &mut sink) {
                Ok(count) => format!("ok(count={count} {})", sink.describe()),
                Err(error) => format!("err({})", error_id(&error)),
            }
        })
        .collect()
}

/// `IndexedArchive::read_entry`, the ordinary source-backed read path.  This is
/// the single body change 0611 modifies, and it enters no strict-layout proof,
/// so none of the four APIs above reach it.
fn sweep_indexed_read_entry<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    names: &[String],
) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            let Some(entry_id) = archive.entry_id(name) else {
                return "no-entry-id".to_string();
            };
            verdict(archive.read_entry(entry_id), |payload| {
                format!(
                    "len={} crc={:08x}",
                    payload.len(),
                    soapberry_zip::crc32(payload)
                )
            })
        })
        .collect()
}

/// `IndexedArchive::with_verified_entry_reader`, the third strict-layout entry
/// point.
fn sweep_indexed_verified_reader<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    names: &[String],
) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            let Some(entry_id) = archive.entry_id(name) else {
                return "no-entry-id".to_string();
            };
            let outcome = archive.with_verified_entry_reader(entry_id, |reader| {
                let mut sink = DigestSink::default();
                io::copy(reader, &mut sink).map(|_| sink)
            });
            match outcome {
                Ok(sink) => format!("ok({})", sink.describe()),
                Err(error) => {
                    use soapberry_zip::office::VerifiedEntryReaderError as Failure;
                    match error {
                        Failure::Archive { error, .. } => {
                            format!("err-archive({})", error_id(&error))
                        },
                        Failure::Transport { error, .. } => {
                            format!("err-transport({})", escape(&format!("{error:?}")))
                        },
                        Failure::Callback(source) => {
                            format!("err-callback({})", escape(&format!("{source:?}")))
                        },
                        other => format!("err-other({})", escape(&format!("{other:?}"))),
                    }
                },
            }
        })
        .collect()
}

/// `IndexedArchive::read_entry_precompressed_and_decoded_with_progress` for
/// every member, not just the first.  This is the one strict-layout entry point
/// the real fuzz target reaches, and it reaches it once per input.
fn sweep_indexed_precompressed<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    names: &[String],
) -> Vec<String> {
    names
        .iter()
        .map(|name| {
            let Some(entry_id) = archive.entry_id(name) else {
                return "no-entry-id".to_string();
            };
            let mut events = 0usize;
            let outcome = archive.read_entry_precompressed_and_decoded_with_progress(
                entry_id,
                |progress| {
                    events = events.saturating_add(1);
                    let _ = black_box(progress);
                    if events > MAX_PRECOMPRESSED_PROGRESS_EVENTS {
                        Err(())
                    } else {
                        Ok(())
                    }
                },
            );
            match outcome {
                Ok((token, decoded)) => format!(
                    "ok(csize={} usize={} crc={:08x} decoded={})",
                    token.compressed_size(),
                    token.uncompressed_size(),
                    token.crc32(),
                    decoded.len()
                ),
                Err(error) => {
                    use soapberry_zip::office::VerifiedPrecompressedError as Failure;
                    match error {
                        Failure::Archive(error) => format!("err-archive({})", error_id(&error)),
                        Failure::Transport(error) => {
                            format!("err-transport({})", escape(&format!("{error:?}")))
                        },
                        Failure::Callback(()) => "err-callback".to_string(),
                        other => format!("err-other({})", escape(&format!("{other:?}"))),
                    }
                },
            }
        })
        .collect()
}

/// The strict-layout sweep, on a fresh reader per order, so the per-reader memo
/// cannot carry a fact between the two traversals.
fn strict_sweep(report: &mut Report, data: &[u8], limits: ArchiveLimits, tag: &str) -> SweepStats {
    let mut stats = SweepStats::default();

    // --- slice-backed ---------------------------------------------------
    match ArchiveReader::new_with_limits(data, limits) {
        Ok(reader) => {
            let names = sorted_names(&reader);
            report.line(&format!("{tag} slice.open ok(len={})", names.len()));
            let forward = sweep_slice_read_to(&reader, &names);
            for (name, outcome) in names.iter().zip(forward.iter()) {
                report.line(&format!("{tag} R.read_to {} {}", escape_name(name), outcome));
                stats.record(outcome);
            }
            // Reverse order on a *fresh* reader: an order-dependent verdict is a
            // defect on either side.
            let reversed: Vec<String> = names.iter().rev().cloned().collect();
            let fresh = ArchiveReader::new_with_limits(data, limits)
                .expect("a reader that opened once opens again");
            let mut back = sweep_slice_read_to(&fresh, &reversed);
            back.reverse();
            report.line(&format!(
                "{tag} R.read_to.order_independent {}",
                back == forward
            ));
            // Same reader, second pass: the memo must not change a verdict.
            let again = sweep_slice_read_to(&reader, &names);
            report.line(&format!("{tag} R.read_to.memo_stable {}", again == forward));

            // The ordinary slice-backed read path, driven exactly as
            // `R.read_to` is: forward, reverse on a fresh reader, and a second
            // pass on the same reader.
            let read_forward = sweep_slice_read(&reader, &names);
            for (name, outcome) in names.iter().zip(read_forward.iter()) {
                report.line(&format!("{tag} R.read {} {}", escape_name(name), outcome));
                stats.record(outcome);
            }
            let read_reversed: Vec<String> = names.iter().rev().cloned().collect();
            let read_fresh = ArchiveReader::new_with_limits(data, limits)
                .expect("a reader that opened once opens again");
            let mut read_back = sweep_slice_read(&read_fresh, &read_reversed);
            read_back.reverse();
            report.line(&format!(
                "{tag} R.read.order_independent {}",
                read_back == read_forward
            ));
            let read_again = sweep_slice_read(&reader, &names);
            report.line(&format!(
                "{tag} R.read.memo_stable {}",
                read_again == read_forward
            ));
        },
        Err(error) => report.line(&format!("{tag} slice.open err({})", error_id(&error))),
    }

    // --- source-backed --------------------------------------------------
    match IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits) {
        Ok(archive) => {
            let names = sorted_names_indexed(&archive);
            report.line(&format!("{tag} indexed.open ok(len={})", names.len()));

            let forward = sweep_indexed_read_entry_to(&archive, &names);
            for (name, outcome) in names.iter().zip(forward.iter()) {
                report.line(&format!("{tag} I.read_entry_to {} {}", escape_name(name), outcome));
                stats.record(outcome);
            }
            let reversed: Vec<String> = names.iter().rev().cloned().collect();
            let fresh = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
                .expect("an archive that opened once opens again");
            let mut back = sweep_indexed_read_entry_to(&fresh, &reversed);
            back.reverse();
            report.line(&format!(
                "{tag} I.read_entry_to.order_independent {}",
                back == forward
            ));
            let again = sweep_indexed_read_entry_to(&archive, &names);
            report.line(&format!(
                "{tag} I.read_entry_to.memo_stable {}",
                again == forward
            ));

            // The ordinary indexed read path -- the path change 0611 modifies
            // -- driven exactly as `I.read_entry_to` is, on its own archive so
            // no strict-layout memo built above can carry a fact into it.
            let ordinary = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
                .expect("an archive that opened once opens again");
            let entry_forward = sweep_indexed_read_entry(&ordinary, &names);
            for (name, outcome) in names.iter().zip(entry_forward.iter()) {
                report.line(&format!("{tag} I.read_entry {} {}", escape_name(name), outcome));
                stats.record(outcome);
            }
            let entry_reversed: Vec<String> = names.iter().rev().cloned().collect();
            let entry_fresh =
                IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
                    .expect("an archive that opened once opens again");
            let mut entry_back = sweep_indexed_read_entry(&entry_fresh, &entry_reversed);
            entry_back.reverse();
            report.line(&format!(
                "{tag} I.read_entry.order_independent {}",
                entry_back == entry_forward
            ));
            let entry_again = sweep_indexed_read_entry(&ordinary, &names);
            report.line(&format!(
                "{tag} I.read_entry.memo_stable {}",
                entry_again == entry_forward
            ));

            let verified = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
                .expect("an archive that opened once opens again");
            for (name, outcome) in names
                .iter()
                .zip(sweep_indexed_verified_reader(&verified, &names))
            {
                report.line(&format!(
                    "{tag} I.verified_reader {} {outcome}",
                    escape_name(name)
                ));
                stats.record(&outcome);
            }

            let precompressed =
                IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
                    .expect("an archive that opened once opens again");
            for (name, outcome) in names
                .iter()
                .zip(sweep_indexed_precompressed(&precompressed, &names))
            {
                report.line(&format!("{tag} I.precompressed {} {outcome}", escape_name(name)));
                stats.record(&outcome);
            }

            // Archive-wide surfaces that 0580 leaves alone, recorded so a
            // regression in them is visible in the same diff.
            let borrowed = ArchiveReader::new_with_limits(data, limits);
            if let Ok(reader) = &borrowed {
                for name in &names {
                    let outcome = verdict(reader.read_stored_borrowed(name), |payload| {
                        payload.map_or_else(
                            || "none".to_string(),
                            |bytes| {
                                format!("len={} crc={:08x}", bytes.len(), soapberry_zip::crc32(bytes))
                            },
                        )
                    });
                    report.line(&format!(
                        "{tag} R.read_stored_borrowed {} {outcome}",
                        escape_name(name)
                    ));
                }
            }
            let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
            let outcome = verdict(
                archive.preservation_index_with_limits(&mut scratch, limits),
                |index| format!("entries={}", index.entries().len()),
            );
            report.line(&format!("{tag} I.preservation_index {outcome}"));
        },
        Err(error) => report.line(&format!("{tag} indexed.open err({})", error_id(&error))),
    }

    stats
}

#[derive(Default, Debug, Clone, Copy)]
struct SweepStats {
    accepted: u64,
    refused_overlap: u64,
    refused_other: u64,
}

impl SweepStats {
    fn record(&mut self, outcome: &str) {
        if outcome.starts_with("ok(") {
            self.accepted += 1;
        } else if outcome.contains("overlapping ZIP local spans") {
            self.refused_overlap += 1;
        } else if outcome.starts_with("err") {
            self.refused_other += 1;
        }
    }
}

// ---------------------------------------------------------------------------
// The port of parse_zip.rs.  Kept as close to the original as a non-libFuzzer
// build allows.
// ---------------------------------------------------------------------------

fn exercise_borrowed_reader(reader: &ArchiveReader<'_>) {
    let mut borrowed_probe_done = false;
    let mut file_count = 0usize;

    for name in reader.file_names() {
        file_count += 1;
        let _ = black_box(reader.metadata(name));
        let stored = reader.is_stored(name);

        if !borrowed_probe_done && matches!(stored, Ok(true)) {
            if let Ok(Some(payload)) = reader.read_stored_borrowed(name) {
                if let Ok(metadata) = reader.metadata(name) {
                    assert_eq!(
                        u64::try_from(payload.len()).ok(),
                        Some(metadata.uncompressed_size())
                    );
                    assert!(!metadata.is_directory());
                }
                black_box(payload);
            }
            borrowed_probe_done = true;
        }
    }

    assert_eq!(file_count, reader.len());
}

fn exercise_reader_at_metadata<R: ReaderAt>(archive: &IndexedArchive<R>) {
    black_box(archive.archive_end_offset());
    black_box(archive.preservation_entry_count());
    black_box(archive.preservation_metadata_bytes());
    black_box(archive.has_encrypted_entries());
    black_box(archive.has_data_descriptor_entries());
    black_box(archive.archive_is_zip64());
    black_box(archive.has_zip64_metadata());
    black_box(archive.all_local_spans_bounded());

    let mut file_count = 0usize;
    for name in archive.file_names() {
        file_count += 1;
        assert!(archive.contains(name));
        let Some(entry_id) = archive.entry_id(name) else {
            continue;
        };
        if let Ok(metadata) = archive.metadata_for(entry_id) {
            assert!(!metadata.is_directory());
            black_box((metadata.compressed_size(), metadata.uncompressed_size()));
        }
        let _ = black_box(archive.metadata(name));
        let _ = black_box(archive.is_stored(name));
    }
    assert_eq!(file_count, archive.len());
}

fn exercise_preservation_index<R: ReaderAt>(archive: &IndexedArchive<R>, limits: ArchiveLimits) {
    let mut scratch = [0u8; RECOMMENDED_BUFFER_SIZE];
    let Ok(index) = archive.preservation_index_with_limits(&mut scratch, limits) else {
        return;
    };

    assert_eq!(index.entries().len(), archive.preservation_entry_count());
    assert_eq!(index.archive_end_offset(), archive.archive_end_offset());
    for entry in index.entries() {
        let local_span = entry.local_span();
        let central_record = entry.central_record();
        assert!(local_span.start < local_span.end);
        assert!(central_record.start < central_record.end);
        let _ = black_box(entry.id());
        let _ = black_box(entry.compression_method());
        let _ = black_box(entry.compressed_size());
        let _ = black_box(entry.uncompressed_size());
        let _ = black_box(entry.raw_name_bytes());
    }
}

fn exercise_precompressed<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    borrowed: &ArchiveReader<'_>,
    limits: ArchiveLimits,
) {
    for name in archive.file_names() {
        let Some(entry_id) = archive.entry_id(name) else {
            continue;
        };
        let Ok(metadata) = archive.metadata_for(entry_id) else {
            continue;
        };
        if metadata.uncompressed_size() > MAX_ENTRY_BYTES {
            continue;
        }

        let Ok(decoded) = borrowed.read(name) else {
            continue;
        };
        let Ok(decoded_size) = u64::try_from(decoded.len()) else {
            continue;
        };
        assert_eq!(decoded_size, metadata.uncompressed_size());

        let mut progress_events = 0usize;
        let result =
            archive.read_entry_precompressed_with_progress(entry_id, &decoded, |progress| {
                progress_events = progress_events.saturating_add(1);
                let _ = black_box(progress);
                if progress_events > MAX_PRECOMPRESSED_PROGRESS_EVENTS {
                    Err(())
                } else {
                    Ok(())
                }
            });
        if let Ok(token) = result {
            assert_eq!(token.compressed_size(), metadata.compressed_size());
            assert_eq!(token.uncompressed_size(), decoded_size);
            assert_eq!(token.crc32(), soapberry_zip::crc32(&decoded));
            let _ = black_box(token.compression_method());
            exercise_precompressed_republication(archive, limits, &decoded, token);
        }

        break;
    }
}

fn republished_payload_range(output: &[u8]) -> Option<(usize, usize)> {
    let archive = ZipArchive::from_slice(output).ok()?;
    let record = archive.entries().find_map(|entry| {
        let entry = entry.ok()?;
        (entry.file_path().as_ref() == REPUBLISHED_MEMBER_NAME.as_bytes()).then_some(entry)
    })?;
    let entry = archive.get_entry(record.wayfinder()).ok()?;
    let (start, end) = entry.compressed_data_range();
    Some((usize::try_from(start).ok()?, usize::try_from(end).ok()?))
}

fn exercise_precompressed_republication<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    limits: ArchiveLimits,
    decoded: &[u8],
    token: soapberry_zip::office::VerifiedPrecompressedEntry,
) {
    if archive.contains(REPUBLISHED_MEMBER_NAME) {
        return;
    }

    let mut scratch = [0u8; RECOMMENDED_BUFFER_SIZE];
    let Ok(index) = archive.preservation_index_with_limits(&mut scratch, limits) else {
        return;
    };
    let mut plan = PreservationPlan::copy_all(&index);
    if plan
        .try_append(RegeneratedEntry::new_precompressed_shared(
            REPUBLISHED_MEMBER_NAME,
            token.clone(),
        ))
        .is_err()
    {
        return;
    }

    let mut baseline_sink = BoundedSink::new(MAX_REPUBLISHED_OUTPUT_BYTES, usize::MAX, None);
    let mut baseline_accounting = ZipOperationAccounting::default();
    if index
        .write_to_with_accounting(&plan, &mut baseline_sink, &mut baseline_accounting)
        .is_err()
    {
        return;
    }
    let output = baseline_sink.bytes;
    let (payload_start, payload_end) = republished_payload_range(&output)
        .expect("successful publication must contain the generated member");
    assert!(payload_start <= payload_end && payload_end <= output.len());

    let mut output_limits = limits;
    output_limits.max_files = limits.max_files.saturating_add(1);
    output_limits.max_metadata_bytes = limits.max_metadata_bytes.saturating_add(128 * 1024);
    output_limits.max_total_size = MAX_REPUBLISHED_OUTPUT_BYTES as u64;
    let output_reader = ArchiveReader::new_with_limits(&output, output_limits)
        .expect("successful bounded publication must reopen within expanded limits");
    let readback = output_reader
        .read(REPUBLISHED_MEMBER_NAME)
        .expect("verified precompressed republish must decode");
    assert_eq!(readback, decoded);

    let payload_bytes = (payload_end - payload_start) as u64;
    match token.compression_method() {
        CompressionMethod::Store => {
            assert_eq!(
                baseline_accounting.stored_payload_bytes_emitted(),
                payload_bytes
            );
            assert_eq!(baseline_accounting.precompressed_payload_bytes_emitted(), 0);
        },
        CompressionMethod::Deflate => {
            assert_eq!(
                baseline_accounting.precompressed_payload_bytes_emitted(),
                payload_bytes
            );
            assert_eq!(baseline_accounting.stored_payload_bytes_emitted(), 0);
        },
        _ => return,
    }
    assert_eq!(
        baseline_accounting.generated_deflate_payload_bytes_emitted(),
        0
    );

    let mut short_sink = BoundedSink::new(
        MAX_REPUBLISHED_OUTPUT_BYTES,
        MAX_REPUBLISHED_WRITE_CHUNK,
        None,
    );
    let mut short_accounting = ZipOperationAccounting::default();
    index
        .write_to_with_accounting(&plan, &mut short_sink, &mut short_accounting)
        .expect("bounded short-write sink must complete");
    assert_eq!(short_sink.bytes, output);
    assert_eq!(short_accounting, baseline_accounting);

    let accepted_prefix = (payload_end - payload_start).min(3);
    let fail_after = payload_start.saturating_add(accepted_prefix);
    if fail_after < output.len() {
        let mut partial_sink = BoundedSink::new(
            MAX_REPUBLISHED_OUTPUT_BYTES,
            MAX_REPUBLISHED_WRITE_CHUNK,
            Some(fail_after),
        );
        let mut partial_accounting = ZipOperationAccounting::default();
        let error = index
            .write_to_with_accounting(&plan, &mut partial_sink, &mut partial_accounting)
            .expect_err("partial sink must fail after its accepted prefix");
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(partial_sink.bytes, output[..fail_after]);
        let accepted_payload = accepted_prefix as u64;
        match token.compression_method() {
            CompressionMethod::Store => {
                assert_eq!(
                    partial_accounting.stored_payload_bytes_emitted(),
                    accepted_payload
                );
                assert_eq!(partial_accounting.precompressed_payload_bytes_emitted(), 0);
            },
            CompressionMethod::Deflate => {
                assert_eq!(
                    partial_accounting.precompressed_payload_bytes_emitted(),
                    accepted_payload
                );
                assert_eq!(partial_accounting.stored_payload_bytes_emitted(), 0);
            },
            _ => return,
        }
        assert_eq!(
            partial_accounting.generated_deflate_payload_bytes_emitted(),
            0
        );
    }
}

fn exercise_fused_precompressed<R: ReaderAt>(
    archive: &IndexedArchive<R>,
    borrowed: &ArchiveReader<'_>,
) {
    for name in archive.file_names() {
        let Some(entry_id) = archive.entry_id(name) else {
            continue;
        };
        let mut progress_events = 0usize;
        let result =
            archive.read_entry_precompressed_and_decoded_with_progress(entry_id, |progress| {
                progress_events = progress_events.saturating_add(1);
                let _ = black_box(progress);
                if progress_events > MAX_PRECOMPRESSED_PROGRESS_EVENTS {
                    Err(())
                } else {
                    Ok(())
                }
            });
        if let Ok((token, decoded)) = result {
            let metadata = archive
                .metadata_for(entry_id)
                .expect("verified entry remains indexed");
            assert_eq!(token.compressed_size(), metadata.compressed_size());
            assert_eq!(token.uncompressed_size(), metadata.uncompressed_size());
            assert_eq!(token.uncompressed_size(), decoded.len() as u64);
            assert_eq!(token.crc32(), soapberry_zip::crc32(&decoded));
            assert_eq!(
                borrowed
                    .read(name)
                    .expect("verified capture remains readable"),
                decoded
            );
            black_box(token);
        }
        break;
    }
}

fn exercise_replay<R: ReaderAt>(archive: &IndexedArchive<R>, data: &[u8]) {
    let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
    let Ok(index) = archive.preservation_index(&mut scratch) else {
        return;
    };
    let Some(target) = index.entries().iter().find(|entry| {
        std::str::from_utf8(entry.raw_name_bytes()).is_ok_and(|name| !name.ends_with('/'))
    }) else {
        return;
    };
    let name = std::str::from_utf8(target.raw_name_bytes()).expect("selected UTF-8 member");
    let payload = &data[..data.len().min(MAX_REPLAY_PAYLOAD_BYTES)];
    let method = if data.len() & 1 != 0 {
        CompressionMethod::Store
    } else {
        CompressionMethod::Deflate
    };
    let limits = ReplayLimits::new(
        MAX_REPLAY_PAYLOAD_BYTES as u64,
        (MAX_REPLAY_PAYLOAD_BYTES * 2) as u64,
        MAX_REPUBLISHED_OUTPUT_BYTES as u64,
    )
    .expect("finite fuzz replay limits");
    let mut sink = BoundedSink::new(MAX_REPUBLISHED_OUTPUT_BYTES, 37, None);
    let mut calls = 0;
    let result = index.write_replacing_with_replay(target.id(), method, limits, &mut sink, |out| {
        calls += 1;
        let chunk = if calls == 1 { 127 } else { 31 };
        for bytes in payload.chunks(chunk) {
            out.write_all(bytes)?;
        }
        Ok::<_, io::Error>(())
    });
    if result.is_err() {
        assert!(
            sink.bytes.is_empty(),
            "stable replay failed after writing output"
        );
        return;
    }
    assert_eq!(calls, 2);
    let output = sink.bytes;
    let mut output_limits = fuzz_limits();
    output_limits.max_total_size += MAX_REPLAY_PAYLOAD_BYTES as u64;
    let reopened = ArchiveReader::new_with_limits(&output, output_limits)
        .expect("successful bounded replay must have readable metadata");
    assert_eq!(
        reopened.read(name).expect("replayed member verifies"),
        payload
    );

    let fail_after = output.len() / 2;
    let mut partial = BoundedSink::new(MAX_REPUBLISHED_OUTPUT_BYTES, 19, Some(fail_after));
    let error = index
        .write_replacing_with_replay(target.id(), method, limits, &mut partial, |out| {
            out.write_all(payload)
        })
        .expect_err("replay sink must fail at its accepted prefix");
    assert_eq!(partial.bytes, output[..fail_after]);
    assert_eq!(error.progress().accepted(), fail_after as u64);

    if !payload.is_empty() {
        let mut drift = BoundedSink::new(MAX_REPUBLISHED_OUTPUT_BYTES, 53, None);
        let mut pass = 0;
        let result =
            index.write_replacing_with_replay(target.id(), method, limits, &mut drift, |out| {
                pass += 1;
                let first = if pass == 1 {
                    payload[0]
                } else {
                    payload[0] ^ 1
                };
                out.write_all(&[first])?;
                out.write_all(&payload[1..])
            });
        assert!(
            result.is_err(),
            "changed replay must never publish successfully"
        );
    }
}

fn exercise_bounded_paths(data: &[u8]) {
    let limits = fuzz_limits();
    let borrowed = ArchiveReader::new_with_limits(data, limits);
    let reader_at = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits);

    if let Ok(reader) = &borrowed {
        exercise_borrowed_reader(reader);
    }
    if let Ok(archive) = &reader_at {
        exercise_reader_at_metadata(archive);
        exercise_preservation_index(archive, limits);
    }

    if let (Ok(reader), Ok(archive)) = (&borrowed, &reader_at) {
        exercise_replay(archive, data);
        exercise_fused_precompressed(archive, reader);
        exercise_precompressed(archive, reader, limits);
        for name in reader.file_names() {
            assert!(archive.contains(name));
            if let (Ok(borrowed_metadata), Ok(reader_at_metadata)) =
                (reader.metadata(name), archive.metadata(name))
            {
                assert_eq!(borrowed_metadata, reader_at_metadata);
            }
        }
    }
}

/// The `fuzz_target!` body of `parse_zip.rs`.
fn exercise_fuzz_target_body(data: &[u8]) {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }

    if let Ok(archive) = ZipArchive::from_slice(data) {
        let _ = archive.entries_hint();
        let _ = archive.eocd_offset();
        let _ = archive.directory_offset();
        let _ = archive.end_offset();

        for entry_result in archive.entries() {
            let Ok(entry) = entry_result else { break };
            let path = entry.file_path();
            let _ = path.as_ref();
            let _ = path.try_normalize();
            let _ = entry.is_dir();
            let _ = entry.compression_method();
            let _ = entry.compressed_size_hint();
            let _ = entry.uncompressed_size_hint();
            let _ = entry.crc32();
        }
    }

    exercise_bounded_paths(data);
}

// ---------------------------------------------------------------------------
// Slice-level metadata record, so a difference in the pure parse shows up too.
// ---------------------------------------------------------------------------

fn record_slice_parse(report: &mut Report, data: &[u8]) {
    match ZipArchive::from_slice(data) {
        Ok(archive) => {
            report.line(&format!(
                "slice.from_slice ok(hint={} eocd={} dir={} end={})",
                archive.entries_hint(),
                archive.eocd_offset(),
                archive.directory_offset(),
                archive.end_offset()
            ));
            let mut count = 0usize;
            for entry_result in archive.entries() {
                let Ok(_entry) = entry_result else {
                    report.line("slice.entries stopped-early");
                    break;
                };
                count += 1;
                if count > 4096 {
                    report.line("slice.entries truncated-record");
                    break;
                }
            }
            report.line(&format!("slice.entries count={count}"));
        },
        Err(error) => report.line(&format!("slice.from_slice err({})", error_id(&error))),
    }
}

// ---------------------------------------------------------------------------
// Driver
// ---------------------------------------------------------------------------

static PANIC_MESSAGE: Mutex<Option<String>> = Mutex::new(None);

fn run_guarded<F: FnOnce()>(body: F) -> Option<String> {
    {
        let mut slot = PANIC_MESSAGE.lock().expect("panic slot");
        *slot = None;
    }
    let outcome = panic::catch_unwind(AssertUnwindSafe(body));
    if outcome.is_ok() {
        return None;
    }
    let captured = PANIC_MESSAGE
        .lock()
        .expect("panic slot")
        .clone()
        .unwrap_or_else(|| "<no message captured>".to_string());
    Some(captured)
}

fn collect_inputs(root: &Path) -> Vec<PathBuf> {
    let mut found = BTreeSet::new();
    let mut stack = vec![root.to_path_buf()];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path.is_file() {
                found.insert(path);
            }
        }
    }
    found.into_iter().collect()
}

fn main() {
    let mut args = std::env::args_os().skip(1);
    let corpus = PathBuf::from(args.next().expect("usage: <corpus-dir> <report-path>"));
    let report_path = PathBuf::from(args.next().expect("usage: <corpus-dir> <report-path>"));

    panic::set_hook(Box::new(|info| {
        let message = if let Some(text) = info.payload().downcast_ref::<&str>() {
            (*text).to_string()
        } else if let Some(text) = info.payload().downcast_ref::<String>() {
            text.clone()
        } else {
            "<non-string panic payload>".to_string()
        };
        let location = info
            .location()
            .map(|loc| format!("{}:{}:{}", loc.file(), loc.line(), loc.column()))
            .unwrap_or_else(|| "<unknown>".to_string());
        let mut slot = PANIC_MESSAGE.lock().expect("panic slot");
        *slot = Some(format!("{message} @ {location}"));
    }));

    let inputs = collect_inputs(&corpus);
    let mut report = Report {
        out: String::with_capacity(1 << 22),
    };
    let mut panics = 0usize;
    let mut totals = (SweepStats::default(), SweepStats::default());

    report.line(&format!("harness 1 inputs {}", inputs.len()));

    for path in &inputs {
        let relative = path
            .strip_prefix(&corpus)
            .unwrap_or(path)
            .to_string_lossy()
            .replace('\\', "/");
        let Ok(data) = std::fs::read(path) else {
            report.line(&format!("=== {relative} unreadable"));
            continue;
        };
        report.line(&format!(
            "=== {relative} len={} sha={:08x}",
            data.len(),
            soapberry_zip::crc32(&data)
        ));

        // Slice-level parse record.
        let mut section = Report { out: String::new() };
        if let Some(message) = run_guarded(|| record_slice_parse(&mut section, &data)) {
            panics += 1;
            report.line(&format!("PANIC slice-parse {}", escape(&message)));
        }
        report.out.push_str(&section.out);

        // Profile A: the fuzz target's own limits and body.
        if let Some(message) = run_guarded(|| exercise_fuzz_target_body(&data)) {
            panics += 1;
            report.line(&format!("PANIC fuzz-body {}", escape(&message)));
        }

        // Profile A sweep.
        let mut section = Report { out: String::new() };
        let mut stats = SweepStats::default();
        if let Some(message) =
            run_guarded(|| stats = strict_sweep(&mut section, &data, fuzz_limits(), "fuzz"))
        {
            panics += 1;
            report.line(&format!("PANIC sweep-fuzz {}", escape(&message)));
        }
        report.out.push_str(&section.out);
        totals.0.accepted += stats.accepted;
        totals.0.refused_overlap += stats.refused_overlap;
        totals.0.refused_other += stats.refused_other;

        // Profile B: wide limits, so real Office fixtures reach the changed code.
        let mut section = Report { out: String::new() };
        let mut stats = SweepStats::default();
        if let Some(message) =
            run_guarded(|| stats = strict_sweep(&mut section, &data, wide_limits(), "wide"))
        {
            panics += 1;
            report.line(&format!("PANIC sweep-wide {}", escape(&message)));
        }
        report.out.push_str(&section.out);
        totals.1.accepted += stats.accepted;
        totals.1.refused_overlap += stats.refused_overlap;
        totals.1.refused_other += stats.refused_other;
    }

    report.line(&format!(
        "TOTAL inputs={} panics={} fuzz[accept={} overlap={} other={}] wide[accept={} overlap={} other={}]",
        inputs.len(),
        panics,
        totals.0.accepted,
        totals.0.refused_overlap,
        totals.0.refused_other,
        totals.1.accepted,
        totals.1.refused_overlap,
        totals.1.refused_other,
    ));

    std::fs::write(&report_path, report.out.as_bytes()).expect("report write");
    eprintln!(
        "inputs={} panics={} report={}",
        inputs.len(),
        panics,
        report_path.display()
    );
}
