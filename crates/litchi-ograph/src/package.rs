use std::io::Cursor;

use litchi_cfb::consts::STGTY_STREAM;
use litchi_cfb::{DirectoryEntry, OleFile};

use crate::chart;
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
const OGRAPH_YEAR_1996: u16 = 0x07CC;
const OGRAPH_YEAR_1997: u16 = 0x07CD;
const REQUIRED_PLATFORM_FLAGS: u32 = 0x0000_0009;
const FORBIDDEN_PLATFORM_FLAGS: u32 = 0x0000_0136;
const RESERVED1: u32 = 0xFFF8_0000;
const RESERVED2: u32 = 0xFFFF_F000;

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
    workbook: WorkbookLayout,
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
        let validated = validate(bytes, limits)?;
        Ok(Self {
            bytes,
            topology: validated.topology,
            workbook: validated.workbook,
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

    /// Reads the fragmented `Workbook` stream into one validated owned view.
    pub fn workbook(self) -> Result<Workbook> {
        let mut cfb = OleFile::open(Cursor::new(self.bytes))?;
        let bytes = cfb.open_stream(&[WORKBOOK])?;
        check_limit(
            "Workbook bytes",
            bytes.len(),
            self.limits.max_workbook_bytes,
        )?;
        self.workbook.check(bytes.len())?;
        Ok(Workbook {
            bytes,
            layout: self.workbook,
            limits: self.limits,
        })
    }
}

/// Move-owned, validated standalone OGraph package.
#[derive(Debug)]
pub struct Package {
    bytes: Vec<u8>,
    topology: Topology,
    workbook: WorkbookLayout,
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
        let workbook = validated.workbook;
        let limits = validated.limits;
        Ok(Self {
            bytes,
            topology,
            workbook,
            limits,
        })
    }

    /// Borrows this package without revalidation or copying.
    pub fn as_ref(&self) -> PackageRef<'_> {
        PackageRef {
            bytes: &self.bytes,
            topology: self.topology,
            workbook: self.workbook,
            limits: self.limits,
        }
    }

    /// Validated stream topology.
    pub const fn topology(&self) -> Topology {
        self.topology
    }

    /// Reads the bounded standalone Workbook stream.
    pub fn workbook(&self) -> Result<Workbook> {
        self.as_ref().workbook()
    }

    /// Consumes the structurally validated package as an opaque attachment payload.
    ///
    /// This is a move: the compound-file allocation is neither cloned nor
    /// rebuilt.
    pub fn finish(self) -> Payload {
        Payload {
            bytes: self.bytes,
            topology: self.topology,
            workbook: self.workbook,
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
    workbook: WorkbookLayout,
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
            workbook: self.workbook,
            limits: self.limits,
        }
    }

    /// Recovers the original allocation without copying.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Borrowed, validated standalone OGraph `Workbook` stream.
#[derive(Debug, Clone, Copy)]
pub struct WorkbookRef<'a> {
    bytes: &'a [u8],
    layout: WorkbookLayout,
    limits: Limits,
}

impl<'a> WorkbookRef<'a> {
    /// Validates the exact standalone globals-plus-chart topology.
    pub fn open(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Validates the standalone topology with explicit resource bounds.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_limit("Workbook bytes", bytes.len(), limits.max_workbook_bytes)?;
        let layout = validate_workbook(bytes, limits)?;
        Ok(Self {
            bytes,
            layout,
            limits,
        })
    }

    /// Exact Workbook stream bytes.
    pub const fn as_bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// The one standalone Microsoft Graph chart substream.
    pub fn chart(self) -> chart::Ref<'a> {
        let bytes = match self
            .bytes
            .get(self.layout.chart_start..self.layout.chart_end)
        {
            Some(bytes) => bytes,
            None => &[],
        };
        chart::Ref::from_validated(
            bytes,
            chart::Kind::Graph,
            self.layout.chart_start,
            self.limits,
        )
    }

    /// Resource limits under which the Workbook was validated.
    pub const fn limits(self) -> Limits {
        self.limits
    }
}

/// Move-owned, validated standalone OGraph `Workbook` stream.
#[derive(Debug)]
pub struct Workbook {
    bytes: Vec<u8>,
    layout: WorkbookLayout,
    limits: Limits,
}

impl Workbook {
    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates under explicit resource bounds.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let workbook = WorkbookRef::with_limits(&bytes, limits)?;
        let layout = workbook.layout;
        let limits = workbook.limits;
        Ok(Self {
            bytes,
            layout,
            limits,
        })
    }

    /// Borrow without copying or revalidation.
    pub fn as_ref(&self) -> WorkbookRef<'_> {
        WorkbookRef {
            bytes: &self.bytes,
            layout: self.layout,
            limits: self.limits,
        }
    }

    /// Exact Workbook stream bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The one standalone Microsoft Graph chart substream.
    pub fn chart(&self) -> chart::Ref<'_> {
        self.as_ref().chart()
    }

    /// Recover the original Workbook allocation without copying.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

#[derive(Debug, Clone, Copy)]
struct ValidatedPackage {
    topology: Topology,
    workbook: WorkbookLayout,
}

#[derive(Debug, Clone, Copy)]
struct WorkbookLayout {
    chart_start: usize,
    chart_end: usize,
    stream_end: usize,
}

impl WorkbookLayout {
    fn check(self, len: usize) -> Result<()> {
        if self.stream_end != len
            || self.chart_start > self.chart_end
            || self.chart_end > self.stream_end
        {
            return workbook_error(0, "validated Workbook layout no longer matches its bytes");
        }
        Ok(())
    }
}

fn validate(bytes: &[u8], limits: Limits) -> Result<ValidatedPackage> {
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
    let workbook_layout = validate_workbook(&workbook, limits)?;

    Ok(ValidatedPackage {
        topology: Topology {
            workbook_bytes,
            comp_obj_bytes: comp_obj,
            ole_bytes: ole,
        },
        workbook: workbook_layout,
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

fn validate_workbook(bytes: &[u8], limits: Limits) -> Result<WorkbookLayout> {
    let mut state = StreamState::GlobalsBof;
    let mut chart_start = None;
    let mut chart_end = None;
    let mut chart_records = 0usize;
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
                chart_start = Some(record.offset());
                chart_records = 1;
                StreamState::Chart
            },
            StreamState::Chart => {
                chart_records = chart_records.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "chart record count",
                })?;
                check_limit(
                    "chart record count",
                    chart_records,
                    limits.max_chart_records,
                )?;
                if record.kind() == BOF {
                    return workbook_error(record.offset(), "nested BOF in chart substream");
                }
                if record.kind() == EOF {
                    validate_eof(record)?;
                    chart_end = Some(record.offset().checked_add(record.encoded().len()).ok_or(
                        Error::SizeOverflow {
                            resource: "chart substream",
                        },
                    )?);
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
        StreamState::Done => {
            let chart_start = chart_start.ok_or(Error::InvalidWorkbook {
                offset: 0,
                reason: "missing chart substream start",
            })?;
            let chart_end = chart_end.ok_or(Error::InvalidWorkbook {
                offset: bytes.len(),
                reason: "missing chart substream end",
            })?;
            Ok(WorkbookLayout {
                chart_start,
                chart_end,
                stream_end: bytes.len(),
            })
        },
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
    let year = le_u16(payload, 6).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF application year is truncated",
    })?;
    if !matches!(year, OGRAPH_YEAR_1996 | OGRAPH_YEAR_1997) {
        return workbook_error(record.offset(), "BOF application year is invalid");
    }
    let flags = le_u32(payload, 8).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF platform flags are truncated",
    })?;
    if flags & REQUIRED_PLATFORM_FLAGS != REQUIRED_PLATFORM_FLAGS
        || flags & FORBIDDEN_PLATFORM_FLAGS != 0
    {
        return workbook_error(record.offset(), "BOF platform flags are invalid");
    }
    if flags & RESERVED1 != 0 {
        return workbook_error(record.offset(), "BOF reserved1 bits are nonzero");
    }
    let highest = (flags >> 14) & 0xF;
    if !valid_version(highest) {
        return workbook_error(
            record.offset(),
            "BOF highest application version is invalid",
        );
    }
    let versions = le_u32(payload, 12).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF version flags are truncated",
    })?;
    if versions & 0xFF != 0x06 {
        return workbook_error(record.offset(), "BOF lowest BIFF version is not 0x06");
    }
    if versions & RESERVED2 != 0 {
        return workbook_error(record.offset(), "BOF reserved2 bits are nonzero");
    }
    let last = (versions >> 8) & 0xF;
    if !valid_version(last) || last > highest {
        return workbook_error(
            record.offset(),
            "BOF last-saved application version is invalid",
        );
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
    Some(u16::from_le_bytes([*pair.first()?, *pair.get(1)?]))
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    let value = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([
        *value.first()?,
        *value.get(1)?,
        *value.get(2)?,
        *value.get(3)?,
    ]))
}

const fn valid_version(value: u32) -> bool {
    matches!(value, 0 | 1 | 2 | 3 | 4 | 6)
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
        payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
        payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
        payload[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14)).to_le_bytes());
        payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
        payload
    }

    fn workbook() -> Vec<u8> {
        workbook_with_bofs(bof(GLOBALS), bof(CHART_SHEET))
    }

    fn workbook_with_bofs(globals: [u8; BOF_BYTES], chart: [u8; BOF_BYTES]) -> Vec<u8> {
        let mut out = Encoder::new();
        out.push(BOF, &globals).expect("globals BOF");
        out.push(Kind::new(0x7777), &[1, 2, 3])
            .expect("unknown record");
        out.push(EOF, &[]).expect("globals EOF");
        out.push(BOF, &chart).expect("chart BOF");
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
        let opened = parsed.workbook().expect("read Workbook");
        assert_eq!(opened.as_bytes(), workbook);
        assert_eq!(opened.chart().kind(), chart::Kind::Graph);
        assert_eq!(opened.chart().records().count(), 3);
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

    #[test]
    fn strict_bof_checks_must_fields_but_preserves_undefined_bits() {
        let mut globals = bof(GLOBALS);
        let ignored = (0b11 << 6) | (1 << 9) | (1 << 10) | (0b11 << 11) | (1 << 13) | (1 << 18);
        globals[8..12]
            .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | ignored | (6 << 14)).to_le_bytes());
        let bytes = workbook_with_bofs(globals, bof(CHART_SHEET));
        let opened = Workbook::open(bytes).expect("undefined BOF bits are preserved");
        assert_eq!(opened.as_bytes(), opened.as_ref().as_bytes());

        let mut invalid = Vec::new();

        let mut year = bof(GLOBALS);
        year[6..8].copy_from_slice(&0x07CB_u16.to_le_bytes());
        invalid.push(year);

        let mut platform = bof(GLOBALS);
        platform[8..12].copy_from_slice(&(6_u32 << 14).to_le_bytes());
        invalid.push(platform);

        let mut forbidden = bof(GLOBALS);
        forbidden[8..12]
            .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (1 << 1) | (6 << 14)).to_le_bytes());
        invalid.push(forbidden);

        let mut reserved1 = bof(GLOBALS);
        reserved1[8..12]
            .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14) | (1 << 19)).to_le_bytes());
        invalid.push(reserved1);

        let mut highest = bof(GLOBALS);
        highest[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (5 << 14)).to_le_bytes());
        invalid.push(highest);

        let mut lowest = bof(GLOBALS);
        lowest[12..16].copy_from_slice(&(0x05_u32 | (4 << 8)).to_le_bytes());
        invalid.push(lowest);

        let mut last = bof(GLOBALS);
        last[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (4 << 14)).to_le_bytes());
        last[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
        invalid.push(last);

        let mut reserved2 = bof(GLOBALS);
        reserved2[12..16].copy_from_slice(&(0x06_u32 | (6 << 8) | (1 << 12)).to_le_bytes());
        invalid.push(reserved2);

        for globals in invalid {
            assert!(matches!(
                WorkbookRef::open(&workbook_with_bofs(globals, bof(CHART_SHEET))),
                Err(Error::InvalidWorkbook { .. })
            ));
        }
    }
}
