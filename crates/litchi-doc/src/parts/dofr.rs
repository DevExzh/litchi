//! Bounded, passive readers for the DOC `RgDofr` frame/list records.
//!
//! The `RgDofr` table is a sequence of `Dofrh` records.  Each record carries
//! its own byte count, so the record boundaries can be checked without
//! interpreting frame layout or opening a referenced file.  This module keeps
//! the complete source bytes and exposes the fixed record fields and bounded
//! variable payloads as typed views.  It deliberately does not build a frame
//! renderer tree, resolve frame file names, or apply list styles.
//!
//! The only supported write operation is a source-checked replacement of a
//! complete record with another record of the same serialized length.  The
//! candidate array is reparsed before the transaction is committed; therefore
//! unknown payload bytes and all record boundaries remain intact.
//!
//! Unknown `Dofrh.dofrt` values are accepted as inert, source-preserved
//! records. Known record types still enforce their typed field domains; this
//! distinction lets an existing producer extension round-trip without making
//! malformed modeled fields look valid.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use std::sync::Arc;

/// FIB `FibRgFcLcb2000` index for `fcRgDofr`/`lcbRgDofr`.
pub const FIB_INDEX_RG_DOFR: usize = 99;
/// Maximum `RgDofr` payload materialized by this reader.
pub const MAX_DOFR_BYTES: usize = 16 * 1024 * 1024;
/// Maximum number of `Dofrh` records accepted in one array.
pub const MAX_DOFR_RECORDS: usize = 65_536;
const DOFR_HEADER_SIZE: usize = 8;
const DOFR_FSN_SIZE: usize = 36;
const DOFR_FSNP_SIZE: usize = 4;
const DOFR_FSN_SPBD_SIZE: usize = 12;
const MAX_FSN_NAME_CHARS: u16 = 255;
const MAX_FSN_FILE_NAME_CHARS: u16 = 258;
const MAX_LIST_STYLES: usize = 65_536;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| corrupted(format!("Dofr {field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| corrupted(format!("Dofr {field} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| corrupted(format!("Dofr {field} is truncated")))?;
    Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// The `Dofrh.dofrt` value (MS-DOC 2.9.63).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DofrType {
    /// Frame-set root record (`dofrtFs`).
    FrameSet,
    /// Frame attributes (`dofrtFsn`).
    Frame,
    /// Child-frame push/pop marker (`dofrtFsnp`).
    ChildMarker,
    /// Frame name (`dofrtFsnName`).
    FrameName,
    /// Frame file path (`dofrtFsnFnm`).
    FrameFileName,
    /// Frame-set border and splitter attributes (`dofrtFsnSpbd`).
    FrameSplitter,
    /// List-style array (`dofrtRglstsf`).
    ListStyles,
    /// An unrecognized producer value retained as an inert record.
    Unknown(u32),
}

impl DofrType {
    fn from_raw(value: u32) -> Self {
        match value {
            0 => Self::FrameSet,
            1 => Self::Frame,
            2 => Self::ChildMarker,
            3 => Self::FrameName,
            4 => Self::FrameFileName,
            5 => Self::FrameSplitter,
            6 => Self::ListStyles,
            other => Self::Unknown(other),
        }
    }

    /// The serialized `Dofrt` value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::FrameSet => 0,
            Self::Frame => 1,
            Self::ChildMarker => 2,
            Self::FrameName => 3,
            Self::FrameFileName => 4,
            Self::FrameSplitter => 5,
            Self::ListStyles => 6,
            Self::Unknown(value) => value,
        }
    }
}

/// The `DofrFsn.fsnk` value (MS-DOC 2.9.97).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DofrFrameKind {
    /// No specified frame kind (`fsnkNil`).
    Nil,
    /// Attributes for the frame set (`fsnkFrameset`).
    FrameSet,
    /// Basic specifications for a frame (`fsnkFrame`).
    Frame,
    /// An unrecognized raw value. Strict `DofrArray` parsing rejects it, while
    /// the variant keeps the public raw-value projection forward compatible.
    Unknown(u32),
}

impl DofrFrameKind {
    fn from_raw(value: u32) -> Self {
        match value {
            0 => Self::Nil,
            1 => Self::FrameSet,
            2 => Self::Frame,
            other => Self::Unknown(other),
        }
    }

    /// The serialized `Fsnk` value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Nil => 0,
            Self::FrameSet => 1,
            Self::Frame => 2,
            Self::Unknown(value) => value,
        }
    }
}

/// The units in a frame-set divider (`Fssd.Units`, MS-DOC 2.9.99).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DofrDividerUnits {
    /// No units are specified (`iFssUnitsNil`).
    Nil,
    /// Pixels (`iFssUnitsPxl`).
    Pixels,
    /// Percentage of the parent (`iFssUnitsPct`).
    Percentage,
    /// Relative position (`iFssUnitsRel`).
    Relative,
    /// An unrecognized raw value. Strict `DofrArray` parsing rejects it, while
    /// the variant keeps the public raw-value projection forward compatible.
    Unknown(u32),
}

impl DofrDividerUnits {
    fn from_raw(value: u32) -> Self {
        match value {
            0 => Self::Nil,
            1 => Self::Pixels,
            2 => Self::Percentage,
            3 => Self::Relative,
            other => Self::Unknown(other),
        }
    }

    /// The serialized `FssUnits` value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Nil => 0,
            Self::Pixels => 1,
            Self::Percentage => 2,
            Self::Relative => 3,
            Self::Unknown(value) => value,
        }
    }
}

/// The scrollbar behavior in `DofrFsn.iidsScroll` (MS-DOC 2.9.122).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DofrScrollType {
    /// A scrollbar appears only when needed (`iScrollAuto`).
    Auto,
    /// Always show a scrollbar (`iScrollYes`).
    Yes,
    /// Never show a scrollbar (`iScrollNo`).
    No,
    /// An unrecognized raw value. Strict `DofrArray` parsing rejects it, while
    /// the variant keeps the public raw-value projection forward compatible.
    Unknown(u32),
}

impl DofrScrollType {
    fn from_raw(value: u32) -> Self {
        match value {
            0 => Self::Auto,
            1 => Self::Yes,
            2 => Self::No,
            other => Self::Unknown(other),
        }
    }

    /// The serialized `IScrollType` value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Auto => 0,
            Self::Yes => 1,
            Self::No => 2,
            Self::Unknown(value) => value,
        }
    }
}

/// The fixed-size divider-position structure from `DofrFsn.fssd`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DofrDivider {
    units: DofrDividerUnits,
    value: i32,
}

impl DofrDivider {
    /// Units used to interpret [`Self::value`].
    #[must_use]
    pub const fn units(&self) -> DofrDividerUnits {
        self.units
    }

    /// Serialized divider position. It is inert and is never applied to a UI.
    #[must_use]
    pub const fn value(&self) -> i32 {
        self.value
    }
}

/// Typed fields of a `DofrFsn` frame record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DofrFrame {
    divider: DofrDivider,
    t_cols: i32,
    frame_kind: DofrFrameKind,
    dx_margin: i32,
    dy_margin: i32,
    scroll: DofrScrollType,
    linked: bool,
    no_resize: bool,
    unused_flags: u32,
    unused2: u32,
}

impl DofrFrame {
    /// Divider position and units (`fssd`).
    #[must_use]
    pub const fn divider(&self) -> DofrDivider {
        self.divider
    }

    /// Child-frame arrangement: -1 means no children, 0 rows, 1 columns.
    #[must_use]
    pub const fn t_cols(&self) -> i32 {
        self.t_cols
    }

    /// Kind of frame described by this record (`fsnk`).
    #[must_use]
    pub const fn frame_kind(&self) -> DofrFrameKind {
        self.frame_kind
    }

    /// Left/right margin in pixels.
    #[must_use]
    pub const fn dx_margin(&self) -> i32 {
        self.dx_margin
    }

    /// Top/bottom margin in pixels.
    #[must_use]
    pub const fn dy_margin(&self) -> i32 {
        self.dy_margin
    }

    /// Scrollbar behavior.
    #[must_use]
    pub const fn scroll(&self) -> DofrScrollType {
        self.scroll
    }

    /// Whether the frame is linked to an external file.
    #[must_use]
    pub const fn linked(&self) -> bool {
        self.linked
    }

    /// Whether the frame size is locked.
    #[must_use]
    pub const fn no_resize(&self) -> bool {
        self.no_resize
    }

    /// Undefined `fUnused1` bits, retained as a raw value for inspection.
    #[must_use]
    pub const fn unused_flags(&self) -> u32 {
        self.unused_flags
    }

    /// Undefined `fUnused2` value.
    #[must_use]
    pub const fn unused2(&self) -> u32 {
        self.unused2
    }
}

/// Typed fields of a `DofrFsnp` child-frame marker.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DofrChildMarker {
    push: bool,
    unused: u32,
}

impl DofrChildMarker {
    /// Whether this marker begins a nested child-frame group.
    #[must_use]
    pub const fn push(&self) -> bool {
        self.push
    }

    /// Undefined marker bits retained as a raw value.
    #[must_use]
    pub const fn unused(&self) -> u32 {
        self.unused
    }
}

/// Typed fields of the fixed `DofrFsnSpbd` splitter/border record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DofrSplitter {
    width_twips: i32,
    color: u32,
    no_border: bool,
    three_d_border: bool,
    unused: u32,
}

impl DofrSplitter {
    /// Border and divider width in twips.
    #[must_use]
    pub const fn width_twips(&self) -> i32 {
        self.width_twips
    }

    /// Raw `COLORREF` value.
    #[must_use]
    pub const fn color(&self) -> u32 {
        self.color
    }

    /// Whether frame-set borders are hidden.
    #[must_use]
    pub const fn no_border(&self) -> bool {
        self.no_border
    }

    /// Whether the border uses a raised three-dimensional style.
    #[must_use]
    pub const fn three_d_border(&self) -> bool {
        self.three_d_border
    }

    /// Required-zero splitter flags (always zero in a valid record).
    #[must_use]
    pub const fn unused(&self) -> u32 {
        self.unused
    }
}

/// A bounded `Xstz` view used by a frame name or file-path record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DofrXstz<'a> {
    source: &'a [u8],
    cch: u16,
}

impl<'a> DofrXstz<'a> {
    /// Number of UTF-16 code units, excluding the required null terminator.
    #[must_use]
    pub const fn cch(&self) -> u16 {
        self.cch
    }

    /// Exact serialized `Xstz` bytes, including its length and terminator.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.source
    }

    /// Borrow the UTF-16 code units without applying Unicode normalization.
    pub fn utf16_units(&self) -> impl Iterator<Item = u16> + 'a {
        self.source[2..2 + usize::from(self.cch) * 2]
            .as_chunks::<2>()
            .0
            .iter()
            .copied()
            .map(u16::from_le_bytes)
    }

    /// Decode the name as Rust text when its UTF-16 units are valid Unicode.
    pub fn text(&self) -> Result<String> {
        String::from_utf16(&self.utf16_units().collect::<Vec<_>>())
            .map_err(|_| corrupted("Dofr Xstz contains invalid UTF-16"))
    }
}

/// One fixed-size list-style entry from a `DofrRglstsf` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DofrListStyle {
    ilst: u16,
    istd_list: u16,
    style_defined: bool,
    unused: u8,
}

impl DofrListStyle {
    /// Zero-based list-definition index (`ilst`).
    #[must_use]
    pub const fn ilst(&self) -> u16 {
        self.ilst
    }

    /// Standard style index (`istdList`, 12 bits).
    #[must_use]
    pub const fn istd_list(&self) -> u16 {
        self.istd_list
    }

    /// Whether this entry is a custom list-style definition.
    #[must_use]
    pub const fn style_defined(&self) -> bool {
        self.style_defined
    }

    /// Required-zero high bits (always zero in a valid record).
    #[must_use]
    pub const fn unused(&self) -> u8 {
        self.unused
    }
}

/// A borrowed view of the `DofrRglstsf` list-style array.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DofrListStyles<'a> {
    source: &'a [u8],
    count: usize,
}

impl<'a> DofrListStyles<'a> {
    /// Number of list-style records in this payload.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.count
    }

    /// Whether this list-style payload contains no entries.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.count == 0
    }

    /// Return one fixed-size list-style entry.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<DofrListStyle> {
        if index >= self.count {
            return None;
        }
        let raw = u32::from_le_bytes(
            self.source[4 + index * 4..8 + index * 4]
                .try_into()
                .expect("validated list-style range"),
        );
        Some(DofrListStyle {
            ilst: raw as u16,
            istd_list: ((raw >> 16) & 0x0FFF) as u16,
            style_defined: raw & (1 << 28) != 0,
            unused: (raw >> 29) as u8,
        })
    }

    /// Exact serialized `DofrRglstsf` payload, including `clstsf`.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.source
    }
}

/// A typed view of the payload selected by a [`DofrType`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DofrPayload<'a> {
    /// The frame-set root has no payload by definition.
    FrameSet,
    /// Fixed frame attributes.
    Frame(DofrFrame),
    /// Child-frame push/pop marker.
    ChildMarker(DofrChildMarker),
    /// Frame name.
    FrameName(DofrXstz<'a>),
    /// Frame file path.
    FrameFileName(DofrXstz<'a>),
    /// Frame-set border/splitter attributes.
    FrameSplitter(DofrSplitter),
    /// List-style array.
    ListStyles(DofrListStyles<'a>),
    /// An unrecognized `Dofrt` value and its inert payload.
    Unknown { kind: u32, bytes: &'a [u8] },
}

/// One borrowed `Dofrh` record from a [`DofrArray`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DofrRecord<'a> {
    bytes: &'a [u8],
    kind: DofrType,
}

impl<'a> DofrRecord<'a> {
    /// Exact serialized record bytes, including `Dofrh`.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.bytes
    }

    /// Serialized `Dofrh.cb`, equal to [`Self::bytes`].len().
    #[must_use]
    pub fn cb(&self) -> u32 {
        u32::from_le_bytes(self.bytes[0..4].try_into().expect("validated Dofrh header"))
    }

    /// Typed record kind.
    #[must_use]
    pub const fn kind(&self) -> DofrType {
        self.kind
    }

    /// Payload bytes following `cb` and `dofrt`.
    #[must_use]
    pub fn payload_bytes(&self) -> &'a [u8] {
        &self.bytes[DOFR_HEADER_SIZE..]
    }

    /// Decode the bounded payload into its typed view.
    pub fn payload(&self) -> Result<DofrPayload<'a>> {
        decode_payload(self.kind, self.payload_bytes())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DofrRecordMeta {
    start: usize,
    end: usize,
    kind: DofrType,
}

/// A bounded `RgDofr` array from the main table stream.
#[derive(Debug, Clone)]
pub struct DofrArray {
    source: Vec<u8>,
    records: Vec<DofrRecordMeta>,
    /// Optional owning DOC source used by the high-level facade to keep a
    /// component patch bound to its complete immutable artifact. Detached
    /// component readers leave this unset intentionally.
    owner: Option<Arc<[u8]>>,
}

impl PartialEq for DofrArray {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.records == other.records
    }
}

impl Eq for DofrArray {}

impl DofrArray {
    /// Parse the optional `RgDofr` table selected by `fcRgDofr`.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_RG_DOFR) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let length = usize::try_from(length)
            .map_err(|_| corrupted("RgDofr length does not fit in memory"))?;
        if length > MAX_DOFR_BYTES {
            return Err(corrupted(format!(
                "RgDofr exceeds the {MAX_DOFR_BYTES}-byte limit"
            )));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("RgDofr offset does not fit in memory"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("RgDofr range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("RgDofr extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete, bounded `RgDofr` byte array.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.is_empty() {
            return Err(corrupted("RgDofr array is empty"));
        }
        if data.len() > MAX_DOFR_BYTES {
            return Err(corrupted(format!(
                "RgDofr exceeds the {MAX_DOFR_BYTES}-byte limit"
            )));
        }

        let mut records = Vec::new();
        let estimated_records = data.len().div_ceil(DOFR_HEADER_SIZE).min(MAX_DOFR_RECORDS);
        records
            .try_reserve_exact(estimated_records)
            .map_err(|error| corrupted(format!("could not reserve RgDofr records: {error}")))?;
        let mut offset = 0usize;
        while offset < data.len() {
            if records.len() >= MAX_DOFR_RECORDS {
                return Err(corrupted(format!(
                    "RgDofr exceeds the {MAX_DOFR_RECORDS}-record limit"
                )));
            }
            let (kind, end) = parse_record(data, offset)?;
            records.push(DofrRecordMeta {
                start: offset,
                end,
                kind,
            });
            offset = end;
        }
        validate_sequence(data, &records)?;
        let mut source = Vec::new();
        source
            .try_reserve_exact(data.len())
            .map_err(|error| corrupted(format!("could not retain RgDofr source: {error}")))?;
        source.extend_from_slice(data);
        Ok(Self {
            source,
            records,
            owner: None,
        })
    }

    /// Attach the immutable owning DOC source used by facade-level patches.
    ///
    /// The component bytes remain independently retained for borrowed record
    /// access. The owner is an existing shared allocation, so this operation
    /// does not copy the complete DOC artifact.
    pub(crate) fn with_owner(mut self, owner: Arc<[u8]>) -> Self {
        self.owner = Some(owner);
        self
    }

    fn with_optional_owner(mut self, owner: Option<Arc<[u8]>>) -> Self {
        self.owner = owner;
        self
    }

    /// Exact source bytes, including every unknown payload byte.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Number of `Dofrh` records.
    #[must_use]
    pub fn len(&self) -> usize {
        self.records.len()
    }

    /// Whether the array contains no records.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.records.is_empty()
    }

    /// Borrow one record by logical array index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<DofrRecord<'_>> {
        let meta = *self.records.get(index)?;
        Some(DofrRecord {
            bytes: &self.source[meta.start..meta.end],
            kind: meta.kind,
        })
    }

    /// Iterate over all records without copying their source payloads.
    pub fn records(&self) -> impl ExactSizeIterator<Item = DofrRecord<'_>> + '_ {
        self.records.iter().map(|meta| DofrRecord {
            bytes: &self.source[meta.start..meta.end],
            kind: meta.kind,
        })
    }

    /// Start a source-bound, same-length record replacement transaction.
    #[must_use]
    pub fn transaction(&self) -> DofrTransaction {
        DofrTransaction {
            source: self.clone(),
            draft: self.source.clone(),
        }
    }

    /// Validate a complete record replacement without creating a transaction.
    ///
    /// The returned value is `false` when the replacement is an exact byte
    /// no-op. Index and fixed-record-length failures are reported before a
    /// transaction can retain a second copy of the array.
    pub fn preflight_record_replacement(&self, index: usize, replacement: &[u8]) -> Result<bool> {
        let meta = *self
            .records
            .get(index)
            .ok_or_else(|| corrupted("RgDofr record index is out of range"))?;
        let expected_len = meta.end - meta.start;
        if replacement.len() != expected_len {
            return Err(corrupted(format!(
                "RgDofr record replacement must remain {expected_len} bytes"
            )));
        }
        Ok(&self.source[meta.start..meta.end] != replacement)
    }
}

/// A checked transaction over complete `Dofrh` records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DofrTransaction {
    source: DofrArray,
    draft: Vec<u8>,
}

impl DofrTransaction {
    /// Original parsed array that this transaction is based on.
    #[must_use]
    pub fn source(&self) -> &DofrArray {
        &self.source
    }

    /// Current validated draft.
    pub fn snapshot(&self) -> Result<DofrArray> {
        DofrArray::parse_bytes(&self.draft)
            .map(|array| array.with_optional_owner(self.source.owner.clone()))
    }

    /// Replace one complete record with a same-length serialized `Dofrh`.
    ///
    /// The replacement is copied only after the candidate array has been
    /// reparsed successfully. A failed replacement leaves the draft unchanged.
    pub fn replace_record_bytes(&mut self, index: usize, replacement: &[u8]) -> Result<&mut Self> {
        self.source
            .preflight_record_replacement(index, replacement)?;
        let meta = *self
            .source
            .records
            .get(index)
            .ok_or_else(|| corrupted("RgDofr record index is out of range"))?;
        let mut candidate = self.draft.clone();
        candidate[meta.start..meta.end].copy_from_slice(replacement);
        DofrArray::parse_bytes(&candidate)?;
        self.draft = candidate;
        Ok(self)
    }

    /// Commit the validated draft and return its reversible source patch.
    pub fn commit(self) -> Result<DofrCommit> {
        let owner = self.source.owner.clone();
        let snapshot = DofrArray::parse_bytes(&self.draft)?.with_optional_owner(owner.clone());
        let patch = DofrPatch {
            source: self.source.source,
            replacement: snapshot.source.clone(),
            owner,
        };
        Ok(DofrCommit { snapshot, patch })
    }
}

/// The result of committing a [`DofrTransaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DofrCommit {
    snapshot: DofrArray,
    patch: DofrPatch,
}

impl DofrCommit {
    /// Validated post-edit array.
    #[must_use]
    pub fn snapshot(&self) -> &DofrArray {
        &self.snapshot
    }

    /// Source-checked reversible patch.
    #[must_use]
    pub fn patch(&self) -> &DofrPatch {
        &self.patch
    }

    /// Whether any serialized byte changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.changed()
    }

    /// Consume the commit and return its post-edit array.
    #[must_use]
    pub fn into_snapshot(self) -> DofrArray {
        self.snapshot
    }
}

/// A reversible, source-checked same-length byte patch for `RgDofr`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DofrPatch {
    source: Vec<u8>,
    replacement: Vec<u8>,
    owner: Option<Arc<[u8]>>,
}

impl DofrPatch {
    /// Exact source bytes required before applying this patch.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Exact replacement bytes written by this patch.
    #[must_use]
    pub fn replacement_bytes(&self) -> &[u8] {
        &self.replacement
    }

    /// Whether applying this patch changes any serialized `RgDofr` byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.source != self.replacement
    }

    /// Validate the FIB-selected range and its exact source bytes without
    /// taking a mutable stream copy.
    ///
    /// The returned range is suitable for a subsequent in-place apply after
    /// the caller has completed any policy checks. No bytes are modified.
    pub fn preflight_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<(usize, usize)> {
        let (start, end) = table_range(fib, table_stream.len())?;
        if table_stream[start..end] != self.source {
            return Err(corrupted(
                "RgDofr patch source does not match the table stream",
            ));
        }
        Ok((start, end))
    }

    /// Apply this patch to a matching parsed array.
    pub fn apply(&self, source: &DofrArray) -> Result<DofrArray> {
        if source.source != self.source {
            return Err(corrupted("RgDofr patch source does not match the array"));
        }
        DofrArray::parse_bytes(&self.replacement)
            .map(|array| array.with_optional_owner(source.owner.clone()))
    }

    /// Revert this patch from a matching post-edit array.
    pub fn revert(&self, replacement: &DofrArray) -> Result<DofrArray> {
        if replacement.source != self.replacement {
            return Err(corrupted(
                "RgDofr patch replacement does not match the array",
            ));
        }
        DofrArray::parse_bytes(&self.source)
            .map(|array| array.with_optional_owner(replacement.owner.clone()))
    }

    /// Apply this patch in place to the FIB-selected table-stream range.
    pub fn apply_to_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &mut [u8],
    ) -> Result<()> {
        let (start, end) = self.preflight_table_stream(fib, table_stream)?;
        table_stream[start..end].copy_from_slice(&self.replacement);
        Ok(())
    }

    /// Revert this patch in place from the matching table-stream range.
    pub fn revert_to_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &mut [u8],
    ) -> Result<()> {
        let (start, end) = table_range(fib, table_stream.len())?;
        if table_stream[start..end] != self.replacement {
            return Err(corrupted(
                "RgDofr patch replacement does not match the table stream",
            ));
        }
        table_stream[start..end].copy_from_slice(&self.source);
        Ok(())
    }

    /// Whether this changed patch was produced from the supplied immutable
    /// DOC source. Pointer identity is a fast path, while exact bytes permit
    /// independently reopened snapshots of the same source to authorize the
    /// patch without admitting a different document.
    pub(crate) fn owner_matches(&self, owner: &Arc<[u8]>) -> bool {
        self.owner
            .as_ref()
            .is_some_and(|bound| Arc::ptr_eq(bound, owner) || bound.as_ref() == owner.as_ref())
    }
}

fn parse_record(data: &[u8], start: usize) -> Result<(DofrType, usize)> {
    let cb = usize::try_from(read_u32(data, start, "Dofrh.cb")?)
        .map_err(|_| corrupted("Dofrh.cb does not fit in memory"))?;
    if cb < DOFR_HEADER_SIZE {
        return Err(corrupted("Dofrh.cb is smaller than its header"));
    }
    let end = start
        .checked_add(cb)
        .ok_or_else(|| corrupted("Dofrh range overflows"))?;
    if end > data.len() {
        return Err(corrupted("Dofrh extends beyond the RgDofr array"));
    }
    let kind = DofrType::from_raw(read_u32(data, start + 4, "Dofrh.dofrt")?);
    let payload = &data[start + DOFR_HEADER_SIZE..end];
    decode_payload(kind, payload)?;
    Ok((kind, end))
}

fn decode_payload<'a>(kind: DofrType, payload: &'a [u8]) -> Result<DofrPayload<'a>> {
    match kind {
        DofrType::FrameSet => {
            if !payload.is_empty() {
                return Err(corrupted("dofrtFs record must not carry a Dofr payload"));
            }
            Ok(DofrPayload::FrameSet)
        },
        DofrType::Frame => {
            if payload.len() != DOFR_FSN_SIZE {
                return Err(corrupted(format!(
                    "DofrFsn must be {DOFR_FSN_SIZE} bytes, got {}",
                    payload.len()
                )));
            }
            let divider_units_raw = read_u32(payload, 0, "DofrFsn.fssd.Units")?;
            if divider_units_raw > 3 {
                return Err(corrupted(format!(
                    "DofrFsn.fssd.Units has unsupported value {divider_units_raw}"
                )));
            }
            let t_cols = read_i32(payload, 8, "DofrFsn.tCols")?;
            if !(-1..=1).contains(&t_cols) {
                return Err(corrupted(format!(
                    "DofrFsn.tCols has unsupported value {t_cols}"
                )));
            }
            let frame_kind_raw = read_u32(payload, 12, "DofrFsn.fsnk")?;
            if frame_kind_raw > 2 {
                return Err(corrupted(format!(
                    "DofrFsn.fsnk has unsupported value {frame_kind_raw}"
                )));
            }
            let scroll_raw = read_u32(payload, 24, "DofrFsn.iidsScroll")?;
            if scroll_raw > 2 {
                return Err(corrupted(format!(
                    "DofrFsn.iidsScroll has unsupported value {scroll_raw}"
                )));
            }
            let flags = read_u32(payload, 28, "DofrFsn.flags")?;
            Ok(DofrPayload::Frame(DofrFrame {
                divider: DofrDivider {
                    units: DofrDividerUnits::from_raw(divider_units_raw),
                    value: read_i32(payload, 4, "DofrFsn.fssd.Val")?,
                },
                t_cols,
                frame_kind: DofrFrameKind::from_raw(frame_kind_raw),
                dx_margin: read_i32(payload, 16, "DofrFsn.dxMargin")?,
                dy_margin: read_i32(payload, 20, "DofrFsn.dyMargin")?,
                scroll: DofrScrollType::from_raw(scroll_raw),
                linked: flags & 1 != 0,
                no_resize: flags & 2 != 0,
                unused_flags: flags >> 2,
                unused2: read_u32(payload, 32, "DofrFsn.fUnused2")?,
            }))
        },
        DofrType::ChildMarker => {
            if payload.len() != DOFR_FSNP_SIZE {
                return Err(corrupted(format!(
                    "DofrFsnp must be {DOFR_FSNP_SIZE} bytes, got {}",
                    payload.len()
                )));
            }
            let flags = read_u32(payload, 0, "DofrFsnp.flags")?;
            Ok(DofrPayload::ChildMarker(DofrChildMarker {
                push: flags & 1 != 0,
                unused: flags >> 1,
            }))
        },
        DofrType::FrameName => Ok(DofrPayload::FrameName(parse_xstz(
            payload,
            MAX_FSN_NAME_CHARS,
            "DofrFsnName",
        )?)),
        DofrType::FrameFileName => Ok(DofrPayload::FrameFileName(parse_xstz(
            payload,
            MAX_FSN_FILE_NAME_CHARS,
            "DofrFsnFnm",
        )?)),
        DofrType::FrameSplitter => {
            if payload.len() != DOFR_FSN_SPBD_SIZE {
                return Err(corrupted(format!(
                    "DofrFsnSpbd must be {DOFR_FSN_SPBD_SIZE} bytes, got {}",
                    payload.len()
                )));
            }
            let width_twips = read_i32(payload, 0, "DofrFsnSpbd.dzaSpb")?;
            if !(0..=31_680).contains(&width_twips) {
                return Err(corrupted(format!(
                    "DofrFsnSpbd.dzaSpb has unsupported value {width_twips}"
                )));
            }
            let flags = read_u32(payload, 8, "DofrFsnSpbd.flags")?;
            if flags >> 2 != 0 {
                return Err(corrupted("DofrFsnSpbd.fUnused must be zero"));
            }
            Ok(DofrPayload::FrameSplitter(DofrSplitter {
                width_twips,
                color: read_u32(payload, 4, "DofrFsnSpbd.cvSpb")?,
                no_border: flags & 1 != 0,
                three_d_border: flags & 2 != 0,
                unused: 0,
            }))
        },
        DofrType::ListStyles => Ok(DofrPayload::ListStyles(parse_list_styles(payload)?)),
        DofrType::Unknown(kind) => Ok(DofrPayload::Unknown {
            kind,
            bytes: payload,
        }),
    }
}

fn parse_xstz<'a>(data: &'a [u8], max_chars: u16, field: &str) -> Result<DofrXstz<'a>> {
    let cch = read_u16(data, 0, &format!("{field}.cch"))?;
    if cch > max_chars {
        return Err(corrupted(format!("{field} exceeds its character limit")));
    }
    let char_bytes = usize::from(cch)
        .checked_mul(2)
        .ok_or_else(|| corrupted(format!("{field} character range overflows")))?;
    let end_without_term = 2usize
        .checked_add(char_bytes)
        .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
    let end = end_without_term
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{field} terminator range overflows")))?;
    if end != data.len() {
        return Err(corrupted(format!(
            "{field} Xstz does not consume its record payload"
        )));
    }
    if read_u16(data, end_without_term, &format!("{field}.chTerm"))? != 0 {
        return Err(corrupted(format!("{field}.chTerm is not zero")));
    }
    Ok(DofrXstz { source: data, cch })
}

fn parse_list_styles<'a>(data: &'a [u8]) -> Result<DofrListStyles<'a>> {
    let count = read_i32(data, 0, "DofrRglstsf.clstsf")?;
    if count < 0 {
        return Err(corrupted("DofrRglstsf.clstsf is negative"));
    }
    let count = usize::try_from(count).map_err(|_| corrupted("DofrRglstsf count overflows"))?;
    if count > MAX_LIST_STYLES {
        return Err(corrupted(format!(
            "DofrRglstsf exceeds the {MAX_LIST_STYLES}-entry limit"
        )));
    }
    let expected = count
        .checked_mul(4)
        .and_then(|bytes| bytes.checked_add(4))
        .ok_or_else(|| corrupted("DofrRglstsf range overflows"))?;
    if expected != data.len() {
        return Err(corrupted(
            "DofrRglstsf list-style array is not record-bounded",
        ));
    }
    for index in 0..count {
        let raw = read_u32(data, 4 + index * 4, "DofrRglstsf.rglstsf")?;
        if raw >> 29 != 0 {
            return Err(corrupted("DofrRglstsf.Lstsf.fUnused must be zero"));
        }
    }
    Ok(DofrListStyles {
        source: data,
        count,
    })
}

fn validate_sequence(data: &[u8], records: &[DofrRecordMeta]) -> Result<()> {
    let first = records
        .first()
        .ok_or_else(|| corrupted("RgDofr array has no Dofrh records"))?;
    match first.kind {
        DofrType::FrameSet => validate_frame_sequence(data, records),
        DofrType::ListStyles => {
            for record in records.iter().skip(1) {
                if !matches!(record.kind, DofrType::ListStyles | DofrType::Unknown(_)) {
                    return Err(corrupted(
                        "RgDofr list-style array contains a frame-set record",
                    ));
                }
            }
            Ok(())
        },
        _ => Err(corrupted("RgDofr must begin with dofrtFs or dofrtRglstsf")),
    }
}

fn validate_frame_sequence(data: &[u8], records: &[DofrRecordMeta]) -> Result<()> {
    let mut depth = 0usize;
    let mut most_recent_frame = None;
    let mut previous_had_frame_association = false;
    for (index, record) in records.iter().enumerate() {
        match record.kind {
            DofrType::FrameSet if index == 0 => {},
            DofrType::FrameSet => {
                return Err(corrupted("RgDofr has more than one frame-set root"));
            },
            DofrType::Frame => {
                most_recent_frame = Some(index);
                previous_had_frame_association = true;
            },
            DofrType::ChildMarker => {
                let parsed = parse_record_payload(data, *record)?;
                let DofrPayload::ChildMarker(marker) = parsed else {
                    return Err(corrupted("DofrFsnp payload has the wrong type"));
                };
                if marker.push() {
                    if !previous_had_frame_association {
                        return Err(corrupted(
                            "DofrFsnp push marker is not attached to the preceding frame-associated record",
                        ));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| corrupted("DofrFsnp nesting depth overflows"))?;
                } else if depth == 0 {
                    return Err(corrupted("DofrFsnp pop marker has no matching push"));
                } else {
                    depth -= 1;
                }
                previous_had_frame_association = false;
            },
            DofrType::FrameName | DofrType::FrameFileName => {
                if most_recent_frame.is_none() {
                    return Err(corrupted(
                        "Dofr frame name/path has no most-recent DofrFsn record",
                    ));
                }
                previous_had_frame_association = true;
            },
            DofrType::FrameSplitter => {
                previous_had_frame_association = false;
            },
            DofrType::ListStyles => {
                return Err(corrupted(
                    "RgDofr frame-set array contains a list-style root record",
                ));
            },
            DofrType::Unknown(_) => {
                // Unknown record kinds remain inert and source-preserved. They
                // do not replace the most recently read DofrFsn for a later
                // DofrFsnName or DofrFsnFnm record.
                previous_had_frame_association = false;
            },
        }
    }
    if depth != 0 {
        return Err(corrupted("DofrFsnp push markers are not balanced"));
    }
    Ok(())
}

fn parse_record_payload<'a>(data: &'a [u8], record: DofrRecordMeta) -> Result<DofrPayload<'a>> {
    let bytes = data
        .get(record.start..record.end)
        .ok_or_else(|| corrupted("Dofr record range is invalid"))?;
    decode_payload(record.kind, &bytes[DOFR_HEADER_SIZE..])
}

fn table_range(fib: &FileInformationBlock, table_stream_len: usize) -> Result<(usize, usize)> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_RG_DOFR) else {
        return Err(corrupted("RgDofr table pointer is unavailable"));
    };
    let length =
        usize::try_from(length).map_err(|_| corrupted("RgDofr length does not fit in memory"))?;
    if length == 0 {
        return Err(corrupted("RgDofr table pointer has zero length"));
    }
    if length > MAX_DOFR_BYTES {
        return Err(corrupted(format!(
            "RgDofr exceeds the {MAX_DOFR_BYTES}-byte limit"
        )));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted("RgDofr offset does not fit in memory"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("RgDofr range overflows"))?;
    if end > table_stream_len {
        return Err(corrupted("RgDofr extends beyond the table stream"));
    }
    Ok((start, end))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u32, payload: &[u8]) -> Vec<u8> {
        let cb = u32::try_from(DOFR_HEADER_SIZE + payload.len()).unwrap();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&cb.to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn frame_payload(frame_kind: u32) -> [u8; DOFR_FSN_SIZE] {
        let mut payload = [0; DOFR_FSN_SIZE];
        payload[12..16].copy_from_slice(&frame_kind.to_le_bytes());
        payload
    }

    fn frame_set() -> Vec<u8> {
        let mut bytes = record(0, &[]);
        bytes.extend_from_slice(&record(1, &frame_payload(2)));
        bytes
    }

    fn empty_xstz() -> Vec<u8> {
        vec![0, 0, 0, 0]
    }

    #[test]
    fn parses_typed_records_and_retains_unknown_payload() {
        let mut bytes = frame_set();
        let marker_payload = 1u32.to_le_bytes();
        bytes.extend_from_slice(&record(2, &marker_payload));
        bytes.extend_from_slice(&record(2, &[0, 0, 0, 0]));
        bytes.extend_from_slice(&record(0xCAFE, &[0xA5, 0x5A]));
        let array = DofrArray::parse_bytes(&bytes).unwrap();
        assert_eq!(array.len(), 5);
        assert_eq!(array.get(1).unwrap().kind(), DofrType::Frame);
        let DofrPayload::Frame(frame) = array.get(1).unwrap().payload().unwrap() else {
            panic!("expected frame payload");
        };
        assert_eq!(frame.frame_kind(), DofrFrameKind::Frame);
        let unknown = array.get(4).unwrap();
        assert_eq!(unknown.kind(), DofrType::Unknown(0xCAFE));
        assert_eq!(unknown.payload_bytes(), &[0xA5, 0x5A]);
        assert_eq!(unknown.bytes(), &bytes[frame_set().len() + 24..]);
        assert_eq!(array.bytes(), bytes.as_slice());
    }

    #[test]
    fn validates_record_shapes_and_marker_balance() {
        let mut bad_root = record(0, &[0]);
        assert!(DofrArray::parse_bytes(&bad_root).is_err());

        bad_root = record(2, &[0, 0, 0, 0]);
        assert!(DofrArray::parse_bytes(&bad_root).is_err());

        let mut unbalanced = record(0, &[]);
        unbalanced.extend_from_slice(&record(2, &1u32.to_le_bytes()));
        assert!(DofrArray::parse_bytes(&unbalanced).is_err());

        let mut malformed_frame = record(0, &[]);
        malformed_frame.extend_from_slice(&record(1, &[0; DOFR_FSN_SIZE - 1]));
        assert!(DofrArray::parse_bytes(&malformed_frame).is_err());
    }

    #[test]
    fn rejects_out_of_domain_typed_fields_and_keeps_ignored_bits() {
        let invalid_frame_fields = [
            (0, 4u32, "divider units"),
            (12, 3u32, "frame kind"),
            (24, 3u32, "scroll type"),
        ];
        for (offset, value, field) in invalid_frame_fields {
            let mut payload = frame_payload(2);
            payload[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            let mut bytes = record(0, &[]);
            bytes.extend_from_slice(&record(1, &payload));
            assert!(
                DofrArray::parse_bytes(&bytes).is_err(),
                "invalid {field} should be rejected"
            );
        }

        for value in [-2i32, 2i32] {
            let mut payload = frame_payload(2);
            payload[8..12].copy_from_slice(&value.to_le_bytes());
            let mut bytes = record(0, &[]);
            bytes.extend_from_slice(&record(1, &payload));
            assert!(DofrArray::parse_bytes(&bytes).is_err());
        }

        for width in [-1i32, 31_681i32] {
            let mut splitter = [0u8; DOFR_FSN_SPBD_SIZE];
            splitter[..4].copy_from_slice(&width.to_le_bytes());
            let mut bytes = record(0, &[]);
            bytes.extend_from_slice(&record(5, &splitter));
            assert!(DofrArray::parse_bytes(&bytes).is_err());
        }

        let mut boundary_frame = frame_payload(0);
        boundary_frame[0..4].copy_from_slice(&3u32.to_le_bytes());
        boundary_frame[8..12].copy_from_slice(&1i32.to_le_bytes());
        boundary_frame[24..28].copy_from_slice(&2u32.to_le_bytes());
        let mut boundary_bytes = record(0, &[]);
        boundary_bytes.extend_from_slice(&record(1, &boundary_frame));
        let boundary_array = DofrArray::parse_bytes(&boundary_bytes).unwrap();
        let DofrPayload::Frame(frame) = boundary_array.get(1).unwrap().payload().unwrap() else {
            panic!("expected boundary frame payload");
        };
        assert_eq!(frame.divider().units(), DofrDividerUnits::Relative);
        assert_eq!(frame.t_cols(), 1);
        assert_eq!(frame.frame_kind(), DofrFrameKind::Nil);
        assert_eq!(frame.scroll(), DofrScrollType::No);

        let mut boundary_splitter = [0u8; DOFR_FSN_SPBD_SIZE];
        boundary_splitter[..4].copy_from_slice(&31_680i32.to_le_bytes());
        let mut boundary_splitter_bytes = record(0, &[]);
        boundary_splitter_bytes.extend_from_slice(&record(5, &boundary_splitter));
        let boundary_splitter_array = DofrArray::parse_bytes(&boundary_splitter_bytes).unwrap();
        let DofrPayload::FrameSplitter(splitter) =
            boundary_splitter_array.get(1).unwrap().payload().unwrap()
        else {
            panic!("expected boundary splitter payload");
        };
        assert_eq!(splitter.width_twips(), 31_680);

        let mut payload = frame_payload(2);
        payload[28..32].copy_from_slice(&u32::MAX.to_le_bytes());
        payload[32..36].copy_from_slice(&u32::MAX.to_le_bytes());
        let mut bytes = record(0, &[]);
        bytes.extend_from_slice(&record(1, &payload));
        bytes.extend_from_slice(&record(2, &u32::MAX.to_le_bytes()));
        bytes.extend_from_slice(&record(2, &0u32.to_le_bytes()));
        let array = DofrArray::parse_bytes(&bytes).unwrap();
        let DofrPayload::Frame(frame) = array.get(1).unwrap().payload().unwrap() else {
            panic!("expected frame payload");
        };
        assert_eq!(frame.unused_flags(), 0x3FFF_FFFF);
        assert_eq!(frame.unused2(), u32::MAX);
        let DofrPayload::ChildMarker(marker) = array.get(2).unwrap().payload().unwrap() else {
            panic!("expected marker payload");
        };
        assert_eq!(marker.unused(), 0x7FFF_FFFF);
    }

    #[test]
    fn rejects_required_zero_bits_in_splitter_and_list_style() {
        let mut splitter = [0u8; DOFR_FSN_SPBD_SIZE];
        splitter[8..12].copy_from_slice(&u32::MAX.to_le_bytes());
        let mut splitter_bytes = record(0, &[]);
        splitter_bytes.extend_from_slice(&record(5, &splitter));
        assert!(DofrArray::parse_bytes(&splitter_bytes).is_err());

        let mut list_style = Vec::new();
        list_style.extend_from_slice(&1i32.to_le_bytes());
        list_style.extend_from_slice(&0xE000_0000u32.to_le_bytes());
        let mut list_style_bytes = record(6, &list_style);
        assert!(DofrArray::parse_bytes(&list_style_bytes).is_err());

        list_style[4..8].copy_from_slice(&0x1000_0003u32.to_le_bytes());
        list_style_bytes = record(6, &list_style);
        let array = DofrArray::parse_bytes(&list_style_bytes).unwrap();
        let DofrPayload::ListStyles(styles) = array.get(0).unwrap().payload().unwrap() else {
            panic!("expected list-style payload");
        };
        assert_eq!(styles.get(0).unwrap().unused(), 0);
    }

    #[test]
    fn names_and_file_names_require_and_follow_the_most_recent_frame() {
        let mut before_frame = record(0, &[]);
        before_frame.extend_from_slice(&record(3, &empty_xstz()));
        before_frame.extend_from_slice(&record(4, &empty_xstz()));
        assert!(DofrArray::parse_bytes(&before_frame).is_err());

        let mut bytes = frame_set();
        let mut second_frame = frame_payload(1);
        second_frame[8..12].copy_from_slice(&(-1i32).to_le_bytes());
        bytes.extend_from_slice(&record(1, &second_frame));
        bytes.extend_from_slice(&record(0xCAFE, &[0xA5, 0x5A]));
        bytes.extend_from_slice(&record(3, &empty_xstz()));
        bytes.extend_from_slice(&record(4, &empty_xstz()));
        let array = DofrArray::parse_bytes(&bytes).unwrap();
        assert_eq!(array.get(3).unwrap().kind(), DofrType::Unknown(0xCAFE));
        assert!(matches!(
            array.get(4).unwrap().payload().unwrap(),
            DofrPayload::FrameName(_)
        ));
        assert!(matches!(
            array.get(5).unwrap().payload().unwrap(),
            DofrPayload::FrameFileName(_)
        ));
    }

    #[test]
    fn child_push_accepts_name_and_file_name_associations() {
        for name_kind in [3u32, 4u32] {
            let mut bytes = frame_set();
            bytes.extend_from_slice(&record(name_kind, &empty_xstz()));
            bytes.extend_from_slice(&record(2, &1u32.to_le_bytes()));
            bytes.extend_from_slice(&record(1, &frame_payload(2)));
            bytes.extend_from_slice(&record(2, &0u32.to_le_bytes()));
            let array = DofrArray::parse_bytes(&bytes).unwrap();
            assert_eq!(array.len(), 6);
            assert_eq!(array.get(2).unwrap().kind(), DofrType::from_raw(name_kind));
        }
    }

    #[test]
    fn replacement_is_atomic_source_checked_and_reversible() {
        let source_bytes = frame_set();
        let source = DofrArray::parse_bytes(&source_bytes).unwrap();
        let mut transaction = source.transaction();
        let replacement = record(1, &frame_payload(1));
        transaction.replace_record_bytes(1, &replacement).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().revert(commit.snapshot()).unwrap().bytes(),
            source.bytes()
        );

        let mut stale = source.transaction();
        assert!(
            stale
                .replace_record_bytes(1, &record(1, &[0; DOFR_FSN_SIZE - 1]))
                .is_err()
        );
        assert_eq!(stale.snapshot().unwrap().bytes(), source.bytes());
        let mut changed_source = source_bytes.clone();
        let last = changed_source.len() - 1;
        changed_source[last] ^= 0xFF;
        let changed = DofrArray::parse_bytes(&changed_source).unwrap();
        assert!(commit.patch().apply(&changed).is_err());
    }

    #[test]
    fn applies_only_to_the_declared_fib_range() {
        let source = frame_set();
        let mut replacement = source.clone();
        let frame_start = record(0, &[]).len();
        replacement[frame_start + 4..frame_start + 8].copy_from_slice(&1u32.to_le_bytes());
        let source_array = DofrArray::parse_bytes(&source).unwrap();
        let replacement_array = DofrArray::parse_bytes(&replacement).unwrap();
        let mut transaction = source_array.transaction();
        transaction
            .replace_record_bytes(1, replacement_array.get(1).unwrap().bytes())
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();

        let index = FIB_INDEX_RG_DOFR;
        let offset = 3u32;
        let mut fib_data = vec![0; 154 + (index + 1) * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&((index + 1) as u16).to_le_bytes());
        let pointer = 154 + index * 8;
        fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        fib_data[pointer + 4..pointer + 8].copy_from_slice(&(source.len() as u32).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let mut table = vec![0xCC; offset as usize + source.len() + 2];
        table[offset as usize..offset as usize + source.len()].copy_from_slice(&source);
        patch.apply_to_table_stream(&fib, &mut table).unwrap();
        assert_eq!(
            &table[offset as usize..offset as usize + source.len()],
            replacement.as_slice()
        );
        patch.revert_to_table_stream(&fib, &mut table).unwrap();
        assert_eq!(
            &table[offset as usize..offset as usize + source.len()],
            source.as_slice()
        );
        table[offset as usize] ^= 1;
        assert!(patch.apply_to_table_stream(&fib, &mut table).is_err());
    }
}
