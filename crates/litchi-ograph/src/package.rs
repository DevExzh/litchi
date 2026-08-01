use std::io::Cursor;

use litchi_cfb::consts::STGTY_STREAM;
use litchi_cfb::{DirectoryEntry, OleFile};

use crate::limits::as_u64;
use crate::raw::{Kind, RecordRef, Records};
use crate::{Error, Limits, Result};

const WORKBOOK: &str = "Workbook";
const COMP_OBJ: &str = "\u{1}CompObj";
const OLE: &str = "\u{1}Ole";
const BOF: Kind = Kind::new(0x0809);
const EOF: Kind = Kind::new(0x000A);
const BOF_BYTES: usize = 16;
const OGRAPH_VERSION: u16 = 0x0680;
const GLOBALS: u16 = 0x0005;
const CHART_SHEET: u16 = 0x8000;

/// Validated root-stream topology of a standalone OGraph compound file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Topology {
    workbook_bytes: u64,
    comp_obj_bytes: Option<u64>,
    ole_bytes: Option<u64>,
}

impl Topology {
    /// Declared byte size of the required `Workbook` stream.
    pub const fn workbook_bytes(self) -> u64 {
        self.workbook_bytes
    }

    /// Declared size of the optional `\u{1}CompObj` stream.
    pub const fn comp_obj_bytes(self) -> Option<u64> {
        self.comp_obj_bytes
    }

    /// Declared size of the optional `\u{1}Ole` stream.
    pub const fn ole_bytes(self) -> Option<u64> {
        self.ole_bytes
    }

    /// Number of allowed root streams present in the package.
    pub const fn stream_count(self) -> usize {
        1 + self.comp_obj_bytes.is_some() as usize + self.ole_bytes.is_some() as usize
    }
}

/// Borrowed, validated standalone OGraph package.
///
/// The compound-file bytes remain caller-owned. CFB stream extraction returns
/// an owned buffer because a logical stream can span non-contiguous sectors;
/// record traversal over that buffer is then zero-copy.
#[derive(Debug, Clone, Copy)]
pub struct PackageRef<'a> {
    bytes: &'a [u8],
    topology: Topology,
    limits: Limits,
}

impl<'a> PackageRef<'a> {
    /// Validates borrowed bytes with conservative default limits.
    pub fn open(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Validates borrowed bytes with explicit resource limits.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_limit("package bytes", bytes.len(), limits.max_package_bytes)?;
        let topology = validate(bytes, limits)?;
        Ok(Self {
            bytes,
            topology,
            limits,
        })
    }

    /// Original compound-file bytes.
    pub const fn as_bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// Validated stream topology.
    pub const fn topology(self) -> Topology {
        self.topology
    }

    /// Resource limits used to validate the package.
    pub const fn limits(self) -> Limits {
        self.limits
    }

    /// Reads the fragmented `Workbook` stream into one contiguous buffer.
    pub fn workbook(self) -> Result<Vec<u8>> {
        let mut cfb = OleFile::open(Cursor::new(self.bytes))?;
        let workbook = cfb.open_stream(&[WORKBOOK])?;
        check_limit(
            "Workbook bytes",
            workbook.len(),
            self.limits.max_workbook_bytes,
        )?;
        Ok(workbook)
    }
}

/// Move-owned, validated standalone OGraph package.
#[derive(Debug)]
pub struct Package {
    bytes: Vec<u8>,
    topology: Topology,
    limits: Limits,
}

impl Package {
    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates with explicit resource limits.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let validated = PackageRef::with_limits(&bytes, limits)?;
        let topology = validated.topology;
        let limits = validated.limits;
        Ok(Self {
            bytes,
            topology,
            limits,
        })
    }

    /// Borrows this package without revalidation or copying.
    pub fn as_ref(&self) -> PackageRef<'_> {
        PackageRef {
            bytes: &self.bytes,
            topology: self.topology,
            limits: self.limits,
        }
    }

    /// Validated stream topology.
    pub const fn topology(&self) -> Topology {
        self.topology
    }

    /// Consumes the structurally validated package as an opaque attachment payload.
    ///
    /// This is a move: the compound-file allocation is neither cloned nor
    /// rebuilt.
    pub fn finish(self) -> Payload {
        Payload {
            bytes: self.bytes,
            topology: self.topology,
            limits: self.limits,
        }
    }
}

/// Opaque standalone OGraph bytes with validated CFB topology and BIFF framing.
///
/// This capability does not claim that the still-opaque chart records satisfy
/// the complete `[MS-OGRAPH]` chart grammar. Hosts that require that guarantee
/// must validate the typed chart model before attaching the payload.
#[derive(Debug)]
pub struct Payload {
    bytes: Vec<u8>,
    topology: Topology,
    limits: Limits,
}

impl Payload {
    /// Takes ownership and validates bytes for attachment.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Package::open(bytes).map(Package::finish)
    }

    /// Takes ownership and validates bytes with explicit resource limits.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        Package::with_limits(bytes, limits).map(Package::finish)
    }

    /// Borrows the validated attachment bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrows the validated package view without copying.
    pub fn as_package(&self) -> PackageRef<'_> {
        PackageRef {
            bytes: &self.bytes,
            topology: self.topology,
            limits: self.limits,
        }
    }

    /// Recovers the original allocation without copying.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

fn validate(bytes: &[u8], limits: Limits) -> Result<Topology> {
    let mut cfb = OleFile::open(Cursor::new(bytes))?;
    let entries = cfb.list_directory_entries(&[])?;
    check_limit("root entries", entries.len(), limits.max_streams)?;

    let mut workbook = None;
    let mut comp_obj = None;
    let mut ole = None;
    for entry in entries {
        validate_entry(entry, limits)?;
        let slot = match entry.name.as_str() {
            WORKBOOK => &mut workbook,
            COMP_OBJ => &mut comp_obj,
            OLE => &mut ole,
            _ => {
                return Err(Error::UnexpectedEntry {
                    name: entry.name.clone(),
                    entry_type: entry.entry_type,
                });
            },
        };
        if slot.replace(entry.size).is_some() {
            return Err(Error::DuplicateStream {
                name: entry.name.clone(),
            });
        }
    }

    let workbook_bytes = workbook.ok_or(Error::MissingStream { name: WORKBOOK })?;
    check_limit_u64("Workbook bytes", workbook_bytes, limits.max_workbook_bytes)?;
    let workbook = cfb.open_stream(&[WORKBOOK])?;
    check_limit("Workbook bytes", workbook.len(), limits.max_workbook_bytes)?;
    validate_workbook(&workbook, limits)?;

    Ok(Topology {
        workbook_bytes,
        comp_obj_bytes: comp_obj,
        ole_bytes: ole,
    })
}

fn validate_entry(entry: &DirectoryEntry, limits: Limits) -> Result<()> {
    if entry.entry_type != STGTY_STREAM {
        return Err(Error::UnexpectedEntry {
            name: entry.name.clone(),
            entry_type: entry.entry_type,
        });
    }
    check_limit_u64("stream bytes", entry.size, limits.max_stream_bytes)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StreamState {
    GlobalsBof,
    Globals,
    ChartBof,
    Chart,
    Done,
}

fn validate_workbook(bytes: &[u8], limits: Limits) -> Result<()> {
    let mut state = StreamState::GlobalsBof;
    for item in Records::with_limits(bytes, limits)? {
        let record = item?;
        state = match state {
            StreamState::GlobalsBof => {
                validate_bof(record, GLOBALS)?;
                StreamState::Globals
            },
            StreamState::Globals => {
                if record.kind() == BOF {
                    return workbook_error(record.offset(), "nested BOF in globals substream");
                }
                if record.kind() == EOF {
                    validate_eof(record)?;
                    StreamState::ChartBof
                } else {
                    StreamState::Globals
                }
            },
            StreamState::ChartBof => {
                validate_bof(record, CHART_SHEET)?;
                StreamState::Chart
            },
            StreamState::Chart => {
                if record.kind() == BOF {
                    return workbook_error(record.offset(), "nested BOF in chart substream");
                }
                if record.kind() == EOF {
                    validate_eof(record)?;
                    StreamState::Done
                } else {
                    StreamState::Chart
                }
            },
            StreamState::Done => {
                return workbook_error(record.offset(), "records follow the chart substream EOF");
            },
        };
    }

    match state {
        StreamState::Done => Ok(()),
        StreamState::GlobalsBof => workbook_error(0, "missing globals substream BOF"),
        StreamState::Globals => workbook_error(bytes.len(), "missing globals substream EOF"),
        StreamState::ChartBof => workbook_error(bytes.len(), "missing chart substream BOF"),
        StreamState::Chart => workbook_error(bytes.len(), "missing chart substream EOF"),
    }
}

fn validate_bof(record: RecordRef<'_>, expected_doc_type: u16) -> Result<()> {
    if record.kind() != BOF {
        return workbook_error(record.offset(), "substream does not begin with BOF");
    }
    if record.payload().len() != BOF_BYTES {
        return workbook_error(record.offset(), "BOF payload is not 16 bytes");
    }
    let payload = record.payload();
    let version = le_u16(payload, 0).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF version is truncated",
    })?;
    let doc_type = le_u16(payload, 2).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF docType is truncated",
    })?;
    if version != OGRAPH_VERSION {
        return workbook_error(record.offset(), "BOF version is not 0x0680");
    }
    if doc_type != expected_doc_type {
        return if expected_doc_type == GLOBALS {
            workbook_error(record.offset(), "first BOF docType is not workbook globals")
        } else {
            workbook_error(record.offset(), "second BOF docType is not chart sheet")
        };
    }
    Ok(())
}

fn validate_eof(record: RecordRef<'_>) -> Result<()> {
    if !record.payload().is_empty() {
        return workbook_error(record.offset(), "EOF record has a non-empty payload");
    }
    Ok(())
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let pair = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([pair[0], pair[1]]))
}

fn workbook_error<T>(offset: usize, reason: &'static str) -> Result<T> {
    Err(Error::InvalidWorkbook { offset, reason })
}

fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed: as_u64(observed),
            maximum: as_u64(maximum),
        });
    }
    Ok(())
}

fn check_limit_u64(resource: &'static str, observed: u64, maximum: usize) -> Result<()> {
    let maximum = as_u64(maximum);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed,
            maximum,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use litchi_cfb::OleWriter;

    use super::*;
    use crate::raw::Encoder;

    fn bof(doc_type: u16) -> [u8; BOF_BYTES] {
        let mut payload = [0; BOF_BYTES];
        payload[0..2].copy_from_slice(&OGRAPH_VERSION.to_le_bytes());
        payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
        payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
        payload
    }

    fn workbook() -> Vec<u8> {
        let mut out = Encoder::new();
        out.push(BOF, &bof(GLOBALS)).expect("globals BOF");
        out.push(Kind::new(0x7777), &[1, 2, 3])
            .expect("unknown record");
        out.push(EOF, &[]).expect("globals EOF");
        out.push(BOF, &bof(CHART_SHEET)).expect("chart BOF");
        out.push(Kind::new(0x7778), &[4, 5])
            .expect("unknown record");
        out.push(EOF, &[]).expect("chart EOF");
        out.finish()
    }

    fn package(workbook: Option<&[u8]>, extras: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = OleWriter::new();
        if let Some(workbook) = workbook {
            writer
                .create_stream(&[WORKBOOK], workbook)
                .expect("Workbook stream");
        }
        for (name, bytes) in extras {
            writer.create_stream(&[*name], bytes).expect("extra stream");
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("write CFB");
        output.into_inner()
    }

    #[test]
    fn accepts_exact_standalone_topology_and_reads_workbook() {
        let workbook = workbook();
        let bytes = package(Some(&workbook), &[(COMP_OBJ, &[1, 2]), (OLE, &[3, 4, 5])]);
        let parsed = PackageRef::open(&bytes).expect("valid package");
        assert_eq!(parsed.topology().stream_count(), 3);
        assert_eq!(parsed.topology().workbook_bytes(), workbook.len() as u64);
        assert_eq!(parsed.workbook().expect("read Workbook"), workbook);
    }

    #[test]
    fn owned_finish_reuses_the_input_allocation() {
        let bytes = package(Some(&workbook()), &[]);
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();
        let payload = Package::open(bytes).expect("valid").finish();
        let bytes = payload.into_bytes();
        assert_eq!(bytes.as_ptr(), pointer);
        assert_eq!(bytes.capacity(), capacity);
    }

    #[test]
    fn rejects_missing_unknown_and_nested_root_entries() {
        let missing = package(None, &[(COMP_OBJ, &[])]);
        assert!(matches!(
            PackageRef::open(&missing),
            Err(Error::MissingStream { name: WORKBOOK })
        ));

        let unknown = package(Some(&workbook()), &[("Other", &[])]);
        assert!(matches!(
            PackageRef::open(&unknown),
            Err(Error::UnexpectedEntry { name, .. }) if name == "Other"
        ));

        let mut writer = OleWriter::new();
        writer.create_storage(&["Nested"]).expect("storage");
        writer
            .create_stream(&[WORKBOOK], &workbook())
            .expect("Workbook");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("write CFB");
        assert!(matches!(
            PackageRef::open(output.get_ref()),
            Err(Error::UnexpectedEntry { name, .. }) if name == "Nested"
        ));
    }

    #[test]
    fn rejects_wrong_substream_shape_and_trailing_records() {
        let mut wrong = Encoder::new();
        wrong.push(BOF, &bof(GLOBALS)).expect("BOF");
        wrong.push(EOF, &[]).expect("EOF");
        wrong.push(BOF, &bof(GLOBALS)).expect("wrong BOF");
        wrong.push(EOF, &[]).expect("EOF");
        let bytes = package(Some(&wrong.finish()), &[]);
        assert!(matches!(
            PackageRef::open(&bytes),
            Err(Error::InvalidWorkbook { .. })
        ));

        let mut trailing = workbook();
        trailing.extend_from_slice(&[0x0A, 0, 0, 0]);
        let bytes = package(Some(&trailing), &[]);
        assert!(matches!(
            PackageRef::open(&bytes),
            Err(Error::InvalidWorkbook { .. })
        ));
    }

    #[test]
    fn checks_package_and_workbook_limits_before_exposing_data() {
        let bytes = package(Some(&workbook()), &[]);
        let package_limits = Limits {
            max_package_bytes: bytes.len() - 1,
            ..Limits::default()
        };
        assert!(matches!(
            PackageRef::with_limits(&bytes, package_limits),
            Err(Error::LimitExceeded {
                resource: "package bytes",
                ..
            })
        ));

        let workbook_limits = Limits {
            max_workbook_bytes: 1,
            ..Limits::default()
        };
        assert!(matches!(
            PackageRef::with_limits(&bytes, workbook_limits),
            Err(Error::LimitExceeded {
                resource: "Workbook bytes",
                ..
            })
        ));
    }
}
