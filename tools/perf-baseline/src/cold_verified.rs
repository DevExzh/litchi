//! Strict Linux page-cache verification for the filesystem evidence harness.
//!
//! This module is deliberately harness-only.  It does not claim that a block
//! filesystem read reached physical media: the proof is limited to an
//! external `fincore` observation immediately before the operation and a
//! positive process `read_bytes` delta during the operation.

use std::{
    fs::{self, File},
    path::Path,
    process::Command,
};

use serde::{Deserialize, Deserializer, Serialize, de};

use crate::process_metrics;

/// The cache states and failure reasons are part of the report's evidence
/// vocabulary.  Ineligible states never produce a timed `CaseResult`.
#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub(crate) enum Status {
    Eligible,
    IneligibleNonLinux,
    IneligibleFilesystemUnknown,
    IneligibleFilesystemUnsupported,
    IneligibleFincoreUnavailable,
    IneligibleFincoreFailed,
    IneligibleFincoreInvalidJson,
    IneligibleFincoreMultipleRecords,
    IneligibleFincorePathMismatch,
    IneligibleFincoreSizeMismatch,
    IneligibleSourceNotRegular,
    IneligibleSourceEmpty,
    IneligibleSourcePageSizeUnavailable,
    IneligibleSourceNotPageAligned,
    IneligibleSourceFsyncFailed,
    IneligibleSourceAdviceFailed,
    IneligibleSourceResident,
    IneligibleSourceDirty,
    IneligibleSourceWriteback,
    IneligibleProcIoUnavailable,
    IneligibleReadBytesZero,
    IneligiblePreparedQueryControl,
    IneligibleSourceAlignmentUnavailable,
    IneligibleSourceWriteFailed,
}

impl Status {
    pub(crate) const fn is_eligible(self) -> bool {
        matches!(self, Self::Eligible)
    }
}

/// One verifier result.  No source path is retained in reports.  Optional
/// values are omitted when a precondition failed before the corresponding
/// observation could be made.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct Sample {
    pub status: Status,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filesystem_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_size_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_pages: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fsync_completed: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub advice: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fincore_size_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub resident_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dirty_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub writeback_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_bytes_before: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_bytes_after: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_bytes_delta: Option<u64>,
}

impl Sample {
    pub(crate) fn ineligible(status: Status) -> Self {
        Self {
            status,
            filesystem_type: None,
            page_size_bytes: None,
            source_bytes: None,
            source_pages: None,
            fsync_completed: None,
            advice: None,
            fincore_size_bytes: None,
            resident_bytes: None,
            dirty_bytes: None,
            writeback_bytes: None,
            read_bytes_before: None,
            read_bytes_after: None,
            read_bytes_delta: None,
        }
    }

    fn with_source(
        status: Status,
        filesystem_type: String,
        page_size_bytes: u64,
        source_bytes: u64,
    ) -> Self {
        Self {
            status,
            filesystem_type: Some(filesystem_type),
            page_size_bytes: Some(page_size_bytes),
            source_bytes: Some(source_bytes),
            source_pages: Some(source_bytes / page_size_bytes),
            fsync_completed: Some(false),
            advice: None,
            fincore_size_bytes: None,
            resident_bytes: None,
            dirty_bytes: None,
            writeback_bytes: None,
            read_bytes_before: None,
            read_bytes_after: None,
            read_bytes_delta: None,
        }
    }
}

/// A successful verifier is scoped to these facts only.  It is intentionally
/// not a physical-media or device-temperature assertion.
pub(crate) const CLAIM_SCOPE: &str = "external fincore page-cache residency/dirty/writeback proof plus positive process read_bytes; no physical-media claim";
pub(crate) const FINCORE_COMMAND: &str =
    "fincore --json --bytes --output FILE,SIZE,RES,DIRTY,WRITEBACK";
pub(crate) const ADVICE: &str = "posix_fadvise_dontneed_accepted";

pub(crate) fn page_size_for_harness() -> Result<u64, Status> {
    #[cfg(not(target_os = "linux"))]
    {
        Err(Status::IneligibleNonLinux)
    }
    #[cfg(target_os = "linux")]
    {
        page_size_bytes().ok_or(Status::IneligibleSourcePageSizeUnavailable)
    }
}

/// A deliberately conservative allowlist of Linux filesystems for which the
/// verifier's page-residency proof is meaningful.  tmpfs, overlayfs, procfs,
/// network filesystems, and unknown filesystems are ineligible.
pub(crate) const SUPPORTED_BLOCK_FILESYSTEMS: &[&str] =
    &["ext2", "ext3", "ext4", "xfs", "btrfs", "f2fs", "zfs"];

pub(crate) const fn supported_block_filesystem(filesystem_type: &str) -> bool {
    let mut index = 0;
    while index < SUPPORTED_BLOCK_FILESYSTEMS.len() {
        if string_eq(filesystem_type, SUPPORTED_BLOCK_FILESYSTEMS[index]) {
            return true;
        }
        index += 1;
    }
    false
}

const fn string_eq(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

/// Fincore's JSON output changes its default columns across util-linux
/// versions.  The harness requests an exact column list and accepts exactly
/// one record with exactly the requested fields.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct FincoreObservation {
    pub size_bytes: u64,
    pub resident_bytes: u64,
    pub dirty_bytes: u64,
    pub writeback_bytes: u64,
}

#[derive(Debug)]
struct FincoreDocument {
    records: Vec<FincoreRecord>,
}

#[derive(Debug)]
struct FincoreRecord {
    file: String,
    size_bytes: u64,
    resident_bytes: u64,
    dirty_bytes: u64,
    writeback_bytes: u64,
}

impl<'de> Deserialize<'de> for FincoreDocument {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        struct DocumentVisitor;

        impl<'de> de::Visitor<'de> for DocumentVisitor {
            type Value = FincoreDocument;

            fn expecting(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.write_str("an object with one fincore array")
            }

            fn visit_map<A>(self, mut map: A) -> Result<Self::Value, A::Error>
            where
                A: de::MapAccess<'de>,
            {
                let mut records = None;
                while let Some(key) = map.next_key::<String>()? {
                    if key != "fincore" || records.is_some() {
                        return Err(de::Error::custom(
                            "fincore JSON has an unknown or duplicate field",
                        ));
                    }
                    records = Some(map.next_value::<Vec<FincoreRecord>>()?);
                }
                Ok(FincoreDocument {
                    records: records.ok_or_else(|| de::Error::missing_field("fincore"))?,
                })
            }
        }

        deserializer.deserialize_map(DocumentVisitor)
    }
}

impl<'de> Deserialize<'de> for FincoreRecord {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        struct RecordVisitor;

        impl<'de> de::Visitor<'de> for RecordVisitor {
            type Value = FincoreRecord;

            fn expecting(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.write_str("one strict fincore record")
            }

            fn visit_map<A>(self, mut map: A) -> Result<Self::Value, A::Error>
            where
                A: de::MapAccess<'de>,
            {
                let mut file = None;
                let mut size_bytes = None;
                let mut resident_bytes = None;
                let mut dirty_bytes = None;
                let mut writeback_bytes = None;
                while let Some(key) = map.next_key::<String>()? {
                    match key.as_str() {
                        "file" if file.is_none() => file = Some(map.next_value::<String>()?),
                        "size" if size_bytes.is_none() => {
                            size_bytes = Some(map.next_value::<u64>()?)
                        },
                        "res" if resident_bytes.is_none() => {
                            resident_bytes = Some(map.next_value::<u64>()?)
                        },
                        "dirty" if dirty_bytes.is_none() => {
                            dirty_bytes = Some(map.next_value::<u64>()?)
                        },
                        "writeback" if writeback_bytes.is_none() => {
                            writeback_bytes = Some(map.next_value::<u64>()?)
                        },
                        _ => {
                            return Err(de::Error::custom(
                                "fincore record has an unknown, duplicate, or missing-column field",
                            ));
                        },
                    }
                }
                Ok(FincoreRecord {
                    file: file.ok_or_else(|| de::Error::missing_field("file"))?,
                    size_bytes: size_bytes.ok_or_else(|| de::Error::missing_field("size"))?,
                    resident_bytes: resident_bytes
                        .ok_or_else(|| de::Error::missing_field("res"))?,
                    dirty_bytes: dirty_bytes.ok_or_else(|| de::Error::missing_field("dirty"))?,
                    writeback_bytes: writeback_bytes
                        .ok_or_else(|| de::Error::missing_field("writeback"))?,
                })
            }
        }

        deserializer.deserialize_map(RecordVisitor)
    }
}

pub(crate) fn parse_fincore_json(output: &[u8]) -> Result<FincoreObservation, Status> {
    let document = serde_json::from_slice::<FincoreDocument>(output)
        .map_err(|_| Status::IneligibleFincoreInvalidJson)?;
    if document.records.len() != 1 {
        return Err(Status::IneligibleFincoreMultipleRecords);
    }
    let record = document
        .records
        .into_iter()
        .next()
        .ok_or(Status::IneligibleFincoreMultipleRecords)?;
    Ok(FincoreObservation {
        size_bytes: record.size_bytes,
        resident_bytes: record.resident_bytes,
        dirty_bytes: record.dirty_bytes,
        writeback_bytes: record.writeback_bytes,
    })
}

fn observation_status(observation: FincoreObservation, source_bytes: u64) -> Status {
    if observation.size_bytes != source_bytes {
        Status::IneligibleFincoreSizeMismatch
    } else if observation.resident_bytes != 0 {
        Status::IneligibleSourceResident
    } else if observation.dirty_bytes != 0 {
        Status::IneligibleSourceDirty
    } else if observation.writeback_bytes != 0 {
        Status::IneligibleSourceWriteback
    } else {
        Status::Eligible
    }
}

/// Checks source metadata, filesystem identity, durable source state, and
/// page residency.  The operation itself is not run here; the child repeats
/// this immediately before its timed source-touching operation.
pub(crate) fn prepare(path: &Path) -> Sample {
    #[cfg(not(target_os = "linux"))]
    {
        let _ = path;
        return Sample::ineligible(Status::IneligibleNonLinux);
    }

    #[cfg(target_os = "linux")]
    {
        let symlink_metadata = match fs::symlink_metadata(path) {
            Ok(metadata) => metadata,
            Err(_) => return Sample::ineligible(Status::IneligibleSourceNotRegular),
        };
        if !symlink_metadata.file_type().is_file() {
            return Sample::ineligible(Status::IneligibleSourceNotRegular);
        }
        let metadata = match fs::metadata(path) {
            Ok(metadata) if metadata.is_file() => metadata,
            _ => return Sample::ineligible(Status::IneligibleSourceNotRegular),
        };
        let source_bytes = metadata.len();
        if source_bytes == 0 {
            return Sample::ineligible(Status::IneligibleSourceEmpty);
        }
        let filesystem_type = match filesystem_type(path) {
            Some(filesystem_type) => filesystem_type,
            None => return Sample::ineligible(Status::IneligibleFilesystemUnknown),
        };
        if !supported_block_filesystem(&filesystem_type) {
            return Sample {
                status: Status::IneligibleFilesystemUnsupported,
                filesystem_type: Some(filesystem_type),
                source_bytes: Some(source_bytes),
                ..Sample::ineligible(Status::IneligibleFilesystemUnsupported)
            };
        }
        let page_size_bytes = match page_size_bytes() {
            Some(page_size_bytes) if page_size_bytes > 0 => page_size_bytes,
            _ => return Sample::ineligible(Status::IneligibleSourcePageSizeUnavailable),
        };
        if source_bytes % page_size_bytes != 0 {
            let mut sample = Sample::with_source(
                Status::IneligibleSourceNotPageAligned,
                filesystem_type,
                page_size_bytes,
                source_bytes,
            );
            sample.fsync_completed = Some(false);
            return sample;
        }
        let mut sample = Sample::with_source(
            Status::Eligible,
            filesystem_type,
            page_size_bytes,
            source_bytes,
        );
        let file = match File::open(path) {
            Ok(file) => file,
            Err(_) => {
                sample.status = Status::IneligibleSourceNotRegular;
                return sample;
            },
        };
        if file.sync_all().is_err() {
            sample.status = Status::IneligibleSourceFsyncFailed;
            return sample;
        }
        sample.fsync_completed = Some(true);
        if rustix::fs::fadvise(&file, 0, None, rustix::fs::Advice::DontNeed).is_err() {
            sample.status = Status::IneligibleSourceAdviceFailed;
            return sample;
        }
        sample.advice = Some(ADVICE.to_owned());
        let observation = match run_fincore(path) {
            Ok(observation) => observation,
            Err(status) => {
                sample.status = status;
                return sample;
            },
        };
        sample.fincore_size_bytes = Some(observation.size_bytes);
        sample.resident_bytes = Some(observation.resident_bytes);
        sample.dirty_bytes = Some(observation.dirty_bytes);
        sample.writeback_bytes = Some(observation.writeback_bytes);
        sample.status = observation_status(observation, source_bytes);
        sample
    }
}

/// Completes the proof after the timed operation.  `read_bytes` is the
/// process-local Linux storage-read counter; a zero delta is ineligible.
pub(crate) fn complete(
    mut sample: Sample,
    before: Option<process_metrics::Snapshot>,
    after: Option<process_metrics::Snapshot>,
) -> Sample {
    if !sample.status.is_eligible() {
        return sample;
    }
    let (Some(before), Some(after)) = (before, after) else {
        sample.status = Status::IneligibleProcIoUnavailable;
        return sample;
    };
    let read_bytes_delta = after.read_bytes.saturating_sub(before.read_bytes);
    sample.read_bytes_before = Some(before.read_bytes);
    sample.read_bytes_after = Some(after.read_bytes);
    sample.read_bytes_delta = Some(read_bytes_delta);
    if read_bytes_delta == 0 {
        sample.status = Status::IneligibleReadBytesZero;
    }
    sample
}

/// Pads a private verifier copy without altering the logical package bytes.
/// ZIP archives use the EOCD comment field so the EOCD remains the final
/// record; CFB archives use trailing zeroes, which the CFB reader ignores
/// after the declared sector chain.
pub(crate) fn page_aligned_archive(
    bytes: &[u8],
    page_size_bytes: u64,
    zip_archive: bool,
) -> Result<Vec<u8>, Status> {
    if bytes.is_empty() || page_size_bytes == 0 {
        return Err(Status::IneligibleSourceAlignmentUnavailable);
    }
    let page_size = usize::try_from(page_size_bytes)
        .map_err(|_| Status::IneligibleSourceAlignmentUnavailable)?;
    if page_size > usize::MAX / 2 {
        return Err(Status::IneligibleSourceAlignmentUnavailable);
    }
    let padding = (page_size - bytes.len() % page_size) % page_size;
    if !zip_archive {
        if padding == 0 {
            return Ok(bytes.to_vec());
        }
        let mut aligned = bytes.to_vec();
        aligned.resize(
            aligned
                .len()
                .checked_add(padding)
                .ok_or(Status::IneligibleSourceAlignmentUnavailable)?,
            0,
        );
        return Ok(aligned);
    }
    if padding > usize::from(u16::MAX) {
        return Err(Status::IneligibleSourceAlignmentUnavailable);
    }
    let search_start = bytes.len().saturating_sub(65_557);
    let eocd = bytes
        .get(search_start..)
        .and_then(|suffix| {
            suffix
                .windows(4)
                .enumerate()
                .rev()
                .find_map(|(offset, window)| {
                    if window != b"PK\x05\x06" {
                        return None;
                    }
                    let candidate = search_start.checked_add(offset)?;
                    let comment_len_offset = candidate.checked_add(20)?;
                    let comment_len_end = comment_len_offset.checked_add(2)?;
                    let comment_len_bytes = bytes.get(comment_len_offset..comment_len_end)?;
                    let comment_len = usize::from(u16::from_le_bytes([
                        comment_len_bytes[0],
                        comment_len_bytes[1],
                    ]));
                    let comment_end = comment_len_end.checked_add(comment_len)?;
                    (comment_end == bytes.len()).then_some(candidate)
                })
        })
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    let comment_len_offset = eocd
        .checked_add(20)
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    let comment_len_end = comment_len_offset
        .checked_add(2)
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    let comment_len_bytes = bytes
        .get(comment_len_offset..comment_len_end)
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    let comment_len = usize::from(u16::from_le_bytes([
        comment_len_bytes[0],
        comment_len_bytes[1],
    ]));
    let comment_end = comment_len_end
        .checked_add(comment_len)
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    if comment_end != bytes.len() {
        return Err(Status::IneligibleSourceAlignmentUnavailable);
    }
    if padding == 0 {
        return Ok(bytes.to_vec());
    }
    let new_comment_len = comment_len
        .checked_add(padding)
        .filter(|length| *length <= usize::from(u16::MAX))
        .ok_or(Status::IneligibleSourceAlignmentUnavailable)?;
    let mut aligned = bytes.to_vec();
    let new_comment_len = u16::try_from(new_comment_len)
        .map_err(|_| Status::IneligibleSourceAlignmentUnavailable)?
        .to_le_bytes();
    aligned[comment_len_offset..comment_len_end].copy_from_slice(&new_comment_len);
    aligned.resize(
        aligned
            .len()
            .checked_add(padding)
            .ok_or(Status::IneligibleSourceAlignmentUnavailable)?,
        0,
    );
    Ok(aligned)
}

#[cfg(target_os = "linux")]
fn filesystem_type(path: &Path) -> Option<String> {
    let output = Command::new("stat")
        .args(["-f", "-c", "%T"])
        .arg("--")
        .arg(path)
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }
    String::from_utf8(output.stdout)
        .ok()
        .map(|value| value.trim().to_owned())
        .filter(|value| !value.is_empty())
}

#[cfg(not(target_os = "linux"))]
fn filesystem_type(_path: &Path) -> Option<String> {
    None
}

#[cfg(target_os = "linux")]
fn page_size_bytes() -> Option<u64> {
    command_output("getconf", &["PAGESIZE"])
        .and_then(|value| value.parse::<u64>().ok())
        .filter(|value| *value > 0)
}

#[cfg(not(target_os = "linux"))]
fn page_size_bytes() -> Option<u64> {
    None
}

#[cfg(target_os = "linux")]
fn run_fincore(path: &Path) -> Result<FincoreObservation, Status> {
    let output = Command::new("fincore")
        .args([
            "--json",
            "--bytes",
            "--output",
            "FILE,SIZE,RES,DIRTY,WRITEBACK",
        ])
        .arg("--")
        .arg(path)
        .output()
        .map_err(|_| Status::IneligibleFincoreUnavailable)?;
    if !output.status.success() {
        return Err(Status::IneligibleFincoreFailed);
    }
    let observation = parse_fincore_json(&output.stdout)?;
    let expected_path = path.to_str().ok_or(Status::IneligibleFincorePathMismatch)?;
    // Parse the path separately only after strict structural parsing.  This
    // keeps the parser useful in adversarial unit tests without retaining a
    // caller path in the report.
    let document = serde_json::from_slice::<FincoreDocument>(&output.stdout)
        .map_err(|_| Status::IneligibleFincoreInvalidJson)?;
    let record = document
        .records
        .first()
        .ok_or(Status::IneligibleFincoreMultipleRecords)?;
    if record.file != expected_path {
        return Err(Status::IneligibleFincorePathMismatch);
    }
    Ok(observation)
}

#[cfg(not(target_os = "linux"))]
fn run_fincore(_path: &Path) -> Result<FincoreObservation, Status> {
    Err(Status::IneligibleNonLinux)
}

#[cfg(target_os = "linux")]
fn command_output(program: &str, arguments: &[&str]) -> Option<String> {
    let output = Command::new(program).args(arguments).output().ok()?;
    if !output.status.success() {
        return None;
    }
    String::from_utf8(output.stdout)
        .ok()
        .map(|value| value.trim().to_owned())
}

#[cfg(test)]
mod tests {
    use super::{
        FincoreObservation, SUPPORTED_BLOCK_FILESYSTEMS, Sample, Status, complete,
        observation_status, page_aligned_archive, parse_fincore_json, supported_block_filesystem,
    };

    fn valid_json() -> &'static [u8] {
        br#"{"fincore":[{"file":"/tmp/source","size":8192,"res":0,"dirty":0,"writeback":0}]}"#
    }

    #[test]
    fn strict_fincore_json_accepts_requested_columns() {
        let observation = parse_fincore_json(valid_json()).unwrap();
        assert_eq!(observation.size_bytes, 8192);
        assert_eq!(observation.resident_bytes, 0);
        assert_eq!(observation.dirty_bytes, 0);
        assert_eq!(observation.writeback_bytes, 0);
    }

    #[test]
    fn strict_fincore_json_rejects_unknown_and_duplicate_fields() {
        for json in [
            br#"{"fincore":[{"file":"/tmp/source","size":8192,"res":0,"dirty":0,"writeback":0,"extra":1}]}"#,
            br#"{"fincore":[{"file":"/tmp/source","size":8192,"res":0,"dirty":0,"writeback":0,"res":0}]}"#,
            br#"{"fincore":[{"file":"/tmp/source","size":"8192","res":0,"dirty":0,"writeback":0}]}"#,
            br#"{"fincore":[{"file":"/tmp/source","size":8192,"res":0,"dirty":0}]}"#,
            br#"{"fincore":[{"file":"/tmp/source","size":8192,"res":0,"dirty":0,"writeback":0}],"extra":[]}"#,
        ] {
            assert_eq!(parse_fincore_json(json), Err(Status::IneligibleFincoreInvalidJson));
        }
    }

    #[test]
    fn strict_fincore_json_rejects_multiple_or_empty_records() {
        assert_eq!(
            parse_fincore_json(br#"{"fincore":[]}"#),
            Err(Status::IneligibleFincoreMultipleRecords)
        );
        assert_eq!(
            parse_fincore_json(
                br#"{"fincore":[{"file":"/tmp/a","size":1,"res":0,"dirty":0,"writeback":0},{"file":"/tmp/b","size":1,"res":0,"dirty":0,"writeback":0}]}"#
            ),
            Err(Status::IneligibleFincoreMultipleRecords)
        );
    }

    #[test]
    fn residency_dirty_and_writeback_states_fail_closed_independently() {
        let source_bytes = 8192;
        for (observation, status) in [
            (
                FincoreObservation {
                    size_bytes: 4096,
                    resident_bytes: 0,
                    dirty_bytes: 0,
                    writeback_bytes: 0,
                },
                Status::IneligibleFincoreSizeMismatch,
            ),
            (
                FincoreObservation {
                    size_bytes: source_bytes,
                    resident_bytes: 4096,
                    dirty_bytes: 0,
                    writeback_bytes: 0,
                },
                Status::IneligibleSourceResident,
            ),
            (
                FincoreObservation {
                    size_bytes: source_bytes,
                    resident_bytes: 0,
                    dirty_bytes: 4096,
                    writeback_bytes: 0,
                },
                Status::IneligibleSourceDirty,
            ),
            (
                FincoreObservation {
                    size_bytes: source_bytes,
                    resident_bytes: 0,
                    dirty_bytes: 0,
                    writeback_bytes: 4096,
                },
                Status::IneligibleSourceWriteback,
            ),
            (
                FincoreObservation {
                    size_bytes: source_bytes,
                    resident_bytes: 0,
                    dirty_bytes: 0,
                    writeback_bytes: 0,
                },
                Status::Eligible,
            ),
        ] {
            assert_eq!(observation_status(observation, source_bytes), status);
        }
    }

    #[test]
    fn process_read_bytes_gate_requires_a_positive_delta() {
        let sample = Sample::with_source(Status::Eligible, "ext4".to_owned(), 4096, 8192);
        let before = crate::process_metrics::Snapshot {
            read_bytes: 10,
            ..Default::default()
        };
        let after = crate::process_metrics::Snapshot {
            read_bytes: 18,
            ..Default::default()
        };
        let completed = complete(sample.clone(), Some(before), Some(after));
        assert_eq!(completed.status, Status::Eligible);
        assert_eq!(completed.read_bytes_delta, Some(8));

        let zero = complete(sample, Some(before), Some(before));
        assert_eq!(zero.status, Status::IneligibleReadBytesZero);
        assert_eq!(zero.read_bytes_delta, Some(0));
    }

    #[test]
    fn process_snapshot_failure_is_explicitly_ineligible() {
        let sample = Sample::with_source(Status::Eligible, "ext4".to_owned(), 4096, 8192);
        let completed = complete(sample, None, None);
        assert_eq!(completed.status, Status::IneligibleProcIoUnavailable);
        assert_eq!(completed.read_bytes_delta, None);
    }

    #[test]
    fn block_filesystem_allowlist_is_conservative() {
        assert_eq!(SUPPORTED_BLOCK_FILESYSTEMS.len(), 7);
        assert!(supported_block_filesystem("ext4"));
        assert!(supported_block_filesystem("xfs"));
        assert!(!supported_block_filesystem("tmpfs"));
        assert!(!supported_block_filesystem("overlayfs"));
        assert!(!supported_block_filesystem("nfs"));
        assert!(!supported_block_filesystem(""));
    }

    #[test]
    fn page_alignment_preserves_zip_eocd_and_cfb_prefix() {
        let zip = [
            b'P', b'K', 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        ];
        let aligned_zip = page_aligned_archive(&zip, 4096, true).unwrap();
        assert_eq!(aligned_zip.len() % 4096, 0);
        assert_eq!(&aligned_zip[..22], &zip);
        assert_eq!(
            u16::from_le_bytes([aligned_zip[20], aligned_zip[21]]) as usize,
            aligned_zip.len() - 22
        );

        let cfb = vec![0x11_u8; 4095];
        let aligned_cfb = page_aligned_archive(&cfb, 4096, false).unwrap();
        assert_eq!(aligned_cfb.len(), 4096);
        assert_eq!(&aligned_cfb[..cfb.len()], cfb.as_slice());
        assert!(aligned_cfb[cfb.len()..].iter().all(|byte| *byte == 0));
    }

    #[test]
    fn zip_alignment_rejects_missing_eocd_and_oversized_page() {
        assert_eq!(
            page_aligned_archive(b"not-a-zip", 4096, true),
            Err(Status::IneligibleSourceAlignmentUnavailable)
        );
        assert_eq!(
            page_aligned_archive(b"payload", 65_536, true),
            Err(Status::IneligibleSourceAlignmentUnavailable)
        );
        assert_eq!(
            page_aligned_archive(&vec![0_u8; 4096], 4096, true),
            Err(Status::IneligibleSourceAlignmentUnavailable)
        );
    }
}
