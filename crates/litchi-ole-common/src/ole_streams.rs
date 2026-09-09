//! Bounded, inert readers and source-checked editors for the OLEDS OLE2
//! presentation streams.
//!
//! The structures in this module are the stream payloads used below an OLE2
//! object storage: `\x02OlePres###` presentation streams, their optional
//! `TOCENTRY` values, and the `\x01Ole10Native` native-data stream.  They are
//! deliberately format-neutral.  Presentation bytes and native bytes remain
//! opaque; this module does not decode images, metafiles, device modes, or
//! native documents, and it never opens, resolves, or activates them.
//!
//! Every parser checks encoded sizes before retaining a variable-length field.
//! Parsed snapshots share their source allocation across clones.  Edits are
//! staged against an immutable snapshot and publish only after the candidate
//! is bounded, encoded, and reparsed.  Unknown reserved fields and trailing
//! producer bytes are copied through a changed presentation edit.

use litchi_cfb::OleError;
use std::fmt;
use std::ops::Range;
use std::sync::Arc;

/// The OLEDS name prefix for presentation streams.
pub const PRESENTATION_STREAM_PREFIX: &str = "\u{0002}OlePres";
/// The OLEDS stream name for converted OLE1.0 native data.
pub const NATIVE_STREAM_NAME: &str = "\u{0001}Ole10Native";
/// OLEDS `NANI` TOC signature.
pub const TOC_SIGNATURE: u32 = 0x494e_414e;
/// Standard clipboard format `CF_BITMAP`.
pub const CF_BITMAP: u32 = 0x0002;
/// Standard clipboard format `CF_METAFILEPICT`.
pub const CF_METAFILEPICT: u32 = 0x0003;
/// Standard clipboard format `CF_DIB`.
pub const CF_DIB: u32 = 0x0008;
/// Standard clipboard format `CF_ENHMETAFILE`.
pub const CF_ENHMETAFILE: u32 = 0x000e;

// [MS-OLEDS] 2.3.4 applies the 0x0201 bound to the primary presentation
// format.  TOCENTRY uses the same wire grammar but does not inherit that
// presentation-only bound; its registered format is bounded by the caller's
// stream limit instead.
const MAX_PRIMARY_REGISTERED_FORMAT_CHARS: usize = 0x0201;
const MAX_PRESENTATION_INDEX: usize = 999;
const DEFAULT_MAX_TOC_ENTRIES: usize = 999;
const DEFAULT_MAX_BYTES: usize = 64 * 1024 * 1024;
const MIN_TOC_ENTRY_BYTES: usize = 40;
// The fixed DEVMODEA fields described by [MS-OLEDS] 2.1.6 occupy 156 bytes.
// dmSize excludes any private driver bytes that follow those fields.
const DEVMODEA_PUBLIC_SIZE: usize = 156;

/// Resource ceilings for one OLEDS stream projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum serialized size of one complete stream.
    pub max_bytes: usize,
    /// Maximum opaque presentation/native payload bytes.
    pub max_data_bytes: usize,
    /// Maximum TOC entries retained from one presentation stream.
    ///
    /// This is a caller policy for one stream and is independent of the
    /// OLEDS storage-level limit on presentation stream names.
    pub max_toc_entries: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_BYTES,
            max_data_bytes: DEFAULT_MAX_BYTES,
            max_toc_entries: DEFAULT_MAX_TOC_ENTRIES,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self, OleError> {
        if self.max_bytes == 0 || self.max_data_bytes == 0 || self.max_toc_entries == 0 {
            return Err(invalid("OLEDS stream limits must be non-zero"));
        }
        Ok(self)
    }
}

/// A standard or registered OLE clipboard format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ClipboardFormat {
    /// No format marker.  This form is accepted for a TOCENTRY, but not for
    /// the primary OLEPresentationStream.
    None,
    /// A standard numeric clipboard format identifier.
    Standard(u32),
    /// A registered ANSI format, including its terminating NUL byte.
    ///
    /// The bytes are intentionally retained instead of decoded through an
    /// arbitrary code page.  This keeps producer data lossless and inert.
    Registered(Arc<[u8]>),
}

impl ClipboardFormat {
    /// Creates a registered-format value from its raw, NUL-terminated ANSI
    /// bytes.
    pub fn registered(bytes: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        let bytes = bytes.into();
        validate_registered_bytes_bounded(&bytes, DEFAULT_MAX_BYTES)?;
        Ok(Self::Registered(bytes.into()))
    }

    /// Returns the standard identifier, if this is a standard format.
    #[must_use]
    pub const fn standard_id(&self) -> Option<u32> {
        match self {
            Self::Standard(value) => Some(*value),
            Self::None | Self::Registered(_) => None,
        }
    }

    /// Borrows the registered ANSI bytes, including the terminating NUL.
    #[must_use]
    pub fn registered_bytes(&self) -> Option<&[u8]> {
        match self {
            Self::Registered(bytes) => Some(bytes),
            Self::None | Self::Standard(_) => None,
        }
    }

    /// Whether this format is the standard CF_METAFILEPICT format.
    #[must_use]
    pub const fn is_metafile_pict(&self) -> bool {
        matches!(self, Self::Standard(CF_METAFILEPICT))
    }

    fn encoded_len(&self) -> Result<usize, OleError> {
        match self {
            Self::None => Ok(4),
            Self::Standard(_) => Ok(8),
            Self::Registered(bytes) => {
                validate_registered_bytes(bytes)?;
                4usize
                    .checked_add(bytes.len())
                    .ok_or_else(|| invalid("registered clipboard format size overflows"))
            },
        }
    }

    fn encode_with_marker(&self, marker: u32, output: &mut Vec<u8>) -> Result<(), OleError> {
        match self {
            Self::None => {
                if marker != 0 {
                    return Err(invalid("empty clipboard format marker is invalid"));
                }
                output.extend_from_slice(&0u32.to_le_bytes());
            },
            Self::Standard(value) => {
                if !matches!(marker, u32::MAX | 0xffff_fffe) {
                    return Err(invalid("standard clipboard format marker is invalid"));
                }
                output.extend_from_slice(&marker.to_le_bytes());
                output.extend_from_slice(&value.to_le_bytes());
            },
            Self::Registered(bytes) => {
                validate_registered_bytes(bytes)?;
                let count = u32::try_from(bytes.len())
                    .map_err(|_| invalid("registered clipboard format exceeds u32"))?;
                if marker != count {
                    return Err(invalid(
                        "registered clipboard format marker does not match its length",
                    ));
                }
                output.extend_from_slice(&marker.to_le_bytes());
                output.extend_from_slice(bytes);
            },
        }
        Ok(())
    }
}

fn default_format_marker(format: &ClipboardFormat) -> u32 {
    match format {
        ClipboardFormat::None => 0,
        ClipboardFormat::Standard(_) => u32::MAX,
        ClipboardFormat::Registered(bytes) => u32::try_from(bytes.len()).unwrap_or(u32::MAX),
    }
}

fn validate_format_marker(format: &ClipboardFormat, marker: u32) -> Result<(), OleError> {
    match format {
        ClipboardFormat::None if marker == 0 => Ok(()),
        ClipboardFormat::Standard(_) if matches!(marker, u32::MAX | 0xffff_fffe) => Ok(()),
        ClipboardFormat::Registered(bytes) if u32::try_from(bytes.len()).ok() == Some(marker) => {
            validate_registered_bytes(bytes)
        },
        ClipboardFormat::None => Err(invalid("empty clipboard format marker is invalid")),
        ClipboardFormat::Standard(_) => Err(invalid("standard clipboard format marker is invalid")),
        ClipboardFormat::Registered(_) => Err(invalid(
            "registered clipboard format marker does not match its length",
        )),
    }
}

fn validate_primary_format(format: &ClipboardFormat) -> Result<(), OleError> {
    if let ClipboardFormat::Registered(bytes) = format {
        validate_registered_bytes_bounded(bytes, MAX_PRIMARY_REGISTERED_FORMAT_CHARS)?;
    }
    Ok(())
}

/// An opaque OLEDS target-device byte sequence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TargetDevice(Arc<[u8]>);

impl TargetDevice {
    /// Creates a target-device value from an opaque DVTARGETDEVICE payload.
    ///
    /// An empty value represents an absent target device.  Non-empty values
    /// must retain at least the fixed DVTARGETDEVICE header.
    pub fn from_bytes(bytes: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        let bytes = bytes.into();
        if bytes.len() > DEFAULT_MAX_BYTES {
            return Err(invalid("target device exceeds the default stream limit"));
        }
        validate_target_device(&bytes, "target device")?;
        Ok(Self(bytes.into()))
    }

    /// Borrows the exact target-device bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.0
    }

    /// Returns whether no target-device field is present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Revision(u64);

impl Revision {
    fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

/// A stable identity for one exact OLEDS stream payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct StreamRevision(u64);

impl StreamRevision {
    /// Returns the raw source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

impl From<Revision> for StreamRevision {
    fn from(value: Revision) -> Self {
        Self(value.0)
    }
}

struct Reader<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len() - self.offset
    }

    fn position(&self) -> usize {
        self.offset
    }

    fn u32(&mut self, field: &str) -> Result<u32, OleError> {
        let bytes = self.take(4, field)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn i32(&mut self, field: &str) -> Result<i32, OleError> {
        let bytes = self.take(4, field)?;
        Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn array<const N: usize>(&mut self, field: &str) -> Result<[u8; N], OleError> {
        let bytes = self.take(N, field)?;
        let mut result = [0; N];
        result.copy_from_slice(bytes);
        Ok(result)
    }

    fn take(&mut self, count: usize, field: &str) -> Result<&'a [u8], OleError> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(format!("{field} size overflows")))?;
        if end > self.bytes.len() {
            return Err(invalid(format!("{field} is truncated")));
        }
        let start = self.offset;
        self.offset = end;
        Ok(&self.bytes[start..end])
    }

    fn sized_range(
        &mut self,
        encoded_size: u32,
        field: &str,
        includes_header: bool,
        required: bool,
    ) -> Result<Option<Range<usize>>, OleError> {
        let size = usize::try_from(encoded_size)
            .map_err(|_| invalid(format!("{field} size exceeds this platform")))?;
        let payload = if includes_header {
            if size == 0 {
                if required {
                    return Err(invalid(format!("{field} must be present")));
                }
                return Ok(None);
            }
            if size < 4 {
                return Err(invalid(format!(
                    "{field} size is smaller than its size field"
                )));
            }
            size - 4
        } else {
            if size == 0 {
                if required {
                    return Err(invalid(format!("{field} must be present")));
                }
                return Ok(None);
            }
            size
        };
        let start = self.position();
        self.take(payload, field)?;
        Ok(Some(start..self.position()))
    }
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

fn validate_registered_bytes(bytes: &[u8]) -> Result<(), OleError> {
    if bytes.is_empty() || bytes.last() != Some(&0) {
        return Err(invalid(
            "registered clipboard format must be NUL terminated",
        ));
    }
    if bytes[..bytes.len() - 1].contains(&0) {
        return Err(invalid(
            "registered clipboard format contains an embedded NUL",
        ));
    }
    Ok(())
}

fn validate_registered_bytes_bounded(bytes: &[u8], max_chars: usize) -> Result<(), OleError> {
    validate_registered_bytes(bytes)?;
    if bytes.len() > max_chars {
        return Err(invalid(
            "registered clipboard format exceeds the configured bound",
        ));
    }
    Ok(())
}

fn parse_clipboard_format(
    reader: &mut Reader<'_>,
    required: bool,
    max_registered_chars: usize,
) -> Result<(ClipboardFormat, u32), OleError> {
    let marker = reader.u32("clipboard format marker")?;
    match marker {
        0 if !required => Ok((ClipboardFormat::None, marker)),
        0 => Err(invalid("primary clipboard format marker must be non-zero")),
        u32::MAX | 0xffff_fffe => Ok((
            ClipboardFormat::Standard(reader.u32("standard clipboard format")?),
            marker,
        )),
        value => {
            let length = usize::try_from(value)
                .map_err(|_| invalid("registered clipboard format size exceeds this platform"))?;
            if length > max_registered_chars {
                return Err(invalid(
                    "registered clipboard format exceeds the configured bound",
                ));
            }
            let bytes = reader.take(length, "registered clipboard format")?;
            validate_registered_bytes_bounded(bytes, max_registered_chars)?;
            Ok((ClipboardFormat::Registered(Arc::from(bytes)), marker))
        },
    }
}

fn checked_len_add(total: &mut usize, value: usize, field: &str) -> Result<(), OleError> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| invalid(format!("{field} serialized size overflows")))?;
    Ok(())
}

fn ensure_u32(value: usize, field: &str) -> Result<u32, OleError> {
    u32::try_from(value).map_err(|_| invalid(format!("{field} exceeds u32")))
}

fn reserve_output(
    output: &mut Vec<u8>,
    size: usize,
    resource: &'static str,
) -> Result<(), OleError> {
    output
        .try_reserve_exact(size)
        .map_err(|source| OleError::Allocation { resource, source })
}

/// A source-backed byte field with an optional copy-on-write replacement.
///
/// Parsed large payloads therefore retain only a range into the one stream
/// allocation.  A replacement allocates exactly once when the caller edits
/// that field, and cloned transactions continue to share it through `Arc`.
#[derive(Debug, Clone, PartialEq, Eq)]
struct BytesField {
    source: Arc<[u8]>,
    range: Range<usize>,
    replacement: Option<Arc<[u8]>>,
}

impl BytesField {
    fn empty() -> Self {
        Self {
            source: Arc::from([]),
            range: 0..0,
            replacement: None,
        }
    }

    fn from_source(source: Arc<[u8]>, range: Range<usize>) -> Self {
        Self {
            source,
            range,
            replacement: None,
        }
    }

    fn from_owned(bytes: Vec<u8>) -> Self {
        let replacement: Arc<[u8]> = bytes.into();
        Self {
            source: Arc::from([]),
            range: 0..0,
            replacement: Some(replacement),
        }
    }

    fn bytes(&self) -> &[u8] {
        self.replacement
            .as_deref()
            .unwrap_or_else(|| &self.source[self.range.clone()])
    }

    fn len(&self) -> usize {
        self.bytes().len()
    }

    fn is_empty(&self) -> bool {
        self.bytes().is_empty()
    }

    fn replace(&mut self, bytes: Vec<u8>) -> bool {
        if self.bytes() == bytes.as_slice() {
            return false;
        }
        if &self.source[self.range.clone()] == bytes.as_slice() {
            self.replacement = None;
        } else {
            self.replacement = Some(bytes.into());
        }
        true
    }
}

fn validate_target_device(bytes: &[u8], field: &str) -> Result<(), OleError> {
    if bytes.is_empty() {
        return Ok(());
    }
    if bytes.len() < 8 {
        return Err(invalid(format!(
            "{field} is shorter than the DVTARGETDEVICE offset header"
        )));
    }
    let offsets = [
        u16::from_le_bytes([bytes[0], bytes[1]]),
        u16::from_le_bytes([bytes[2], bytes[3]]),
        u16::from_le_bytes([bytes[4], bytes[5]]),
        u16::from_le_bytes([bytes[6], bytes[7]]),
    ];
    let mut ranges = Vec::new();
    for (index, offset) in offsets.into_iter().enumerate() {
        if offset == 0 {
            continue;
        }
        let start = usize::from(offset);
        if start < 8 || start >= bytes.len() {
            return Err(invalid(format!(
                "{field} field {index} offset is outside the target-device bytes"
            )));
        }
        let end = if index < 3 {
            bytes[start..]
                .iter()
                .position(|byte| *byte == 0)
                .map(|position| start + position + 1)
                .ok_or_else(|| {
                    invalid(format!("{field} ANSI field {index} is not NUL terminated"))
                })?
        } else {
            // DEVMODEA keeps dmSize and dmDriverExtra at offsets 68 and 70:
            // two 32-byte ANSI names precede the fixed header.  Validate the
            // complete public structure before retaining the opaque field.
            if bytes.len() < start + 72 {
                return Err(invalid(format!("{field} DEVMODEA header is truncated")));
            }
            let dm_size = usize::from(u16::from_le_bytes([bytes[start + 68], bytes[start + 69]]));
            let dm_driver_extra =
                usize::from(u16::from_le_bytes([bytes[start + 70], bytes[start + 71]]));
            if dm_size < DEVMODEA_PUBLIC_SIZE {
                return Err(invalid(format!(
                    "{field} DEVMODEA dmSize is smaller than its public structure"
                )));
            }
            let public_end = start
                .checked_add(dm_size)
                .ok_or_else(|| invalid(format!("{field} DEVMODEA size overflows")))?;
            public_end
                .checked_add(dm_driver_extra)
                .filter(|end| *end <= bytes.len())
                .ok_or_else(|| {
                    invalid(format!("{field} DEVMODEA extends past the target device"))
                })?
        };
        ranges.push((start, end));
    }
    ranges.sort_unstable_by_key(|(start, _)| *start);
    for pair in ranges.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(invalid(format!("{field} fields overlap")));
        }
    }
    Ok(())
}

/// A typed OLEDS TOCENTRY value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TocEntry {
    wire: BytesField,
    format: ClipboardFormat,
    format_marker: u32,
    target_device: BytesField,
    aspect: u32,
    lindex: u32,
    tymed: u32,
    reserved1: [u8; 12],
    advf: u32,
    reserved2: u32,
    dirty: bool,
}

impl TocEntry {
    fn content_eq(&self, other: &Self) -> bool {
        self.wire == other.wire
            && self.format == other.format
            && self.format_marker == other.format_marker
            && self.target_device == other.target_device
            && self.aspect == other.aspect
            && self.lindex == other.lindex
            && self.tymed == other.tymed
            && self.reserved1 == other.reserved1
            && self.advf == other.advf
            && self.reserved2 == other.reserved2
    }

    /// Parses one TOCENTRY from a complete byte slice.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        if bytes.len() > DEFAULT_MAX_BYTES {
            return Err(invalid("TOCENTRY exceeds the default stream limit"));
        }
        let snapshot = Self::parse_shared(Arc::<[u8]>::from(bytes))?;
        Ok(snapshot)
    }

    /// Parses one TOCENTRY while sharing the supplied source allocation.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self, OleError> {
        if bytes.len() > DEFAULT_MAX_BYTES {
            return Err(invalid("TOCENTRY exceeds the default stream limit"));
        }
        let mut reader = Reader::new(&bytes);
        let (format, format_marker) =
            parse_clipboard_format(&mut reader, false, DEFAULT_MAX_BYTES)?;
        let target_size = reader.u32("TOCENTRY target-device size")?;
        let target = reader.sized_range(target_size, "TOCENTRY target device", false, false)?;
        if let Some(range) = target.as_ref() {
            validate_target_device(&bytes[range.clone()], "TOCENTRY target device")?;
        }
        let aspect = reader.u32("TOCENTRY aspect")?;
        let lindex = reader.u32("TOCENTRY lindex")?;
        let tymed = reader.u32("TOCENTRY tymed")?;
        let reserved1 = reader.array::<12>("TOCENTRY reserved bytes")?;
        let advf = reader.u32("TOCENTRY advf")?;
        let reserved2 = reader.u32("TOCENTRY reserved2")?;
        if reader.remaining() != 0 {
            return Err(invalid("TOCENTRY contains trailing bytes"));
        }
        Ok(Self {
            wire: BytesField::from_source(Arc::clone(&bytes), 0..bytes.len()),
            format,
            format_marker,
            target_device: target.map_or_else(BytesField::empty, |range| {
                BytesField::from_source(Arc::clone(&bytes), range)
            }),
            aspect,
            lindex,
            tymed,
            reserved1,
            advf,
            reserved2,
            dirty: false,
        })
    }

    /// Creates a TOCENTRY from typed fields using the OLEDS wire defaults.
    pub fn new(format: ClipboardFormat) -> Result<Self, OleError> {
        let format_marker = default_format_marker(&format);
        let entry = Self {
            wire: BytesField::empty(),
            format,
            format_marker,
            target_device: BytesField::empty(),
            aspect: 0,
            lindex: u32::MAX,
            tymed: 0,
            reserved1: [0; 12],
            advf: 0,
            reserved2: 0,
            dirty: true,
        };
        let bytes = entry.to_bytes()?;
        Self::parse_shared(bytes.into())
    }

    /// Exact source bytes, when the entry is unedited.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.wire.bytes()
    }

    /// The clipboard format descriptor.
    #[must_use]
    pub fn clipboard_format(&self) -> &ClipboardFormat {
        &self.format
    }

    /// The exact MarkerOrLength value retained for the clipboard format.
    #[must_use]
    pub const fn clipboard_format_marker(&self) -> u32 {
        self.format_marker
    }

    /// Replaces the clipboard format.
    pub fn set_clipboard_format(&mut self, value: ClipboardFormat) -> Result<(), OleError> {
        value.encoded_len()?;
        if self.format != value {
            self.format_marker = match (&value, self.format_marker) {
                (ClipboardFormat::Standard(_), marker @ (u32::MAX | 0xffff_fffe)) => marker,
                _ => default_format_marker(&value),
            };
            self.format = value;
            self.dirty = true;
        }
        Ok(())
    }

    /// Borrows the opaque target-device bytes.
    #[must_use]
    pub fn target_device(&self) -> &[u8] {
        self.target_device.bytes()
    }

    /// Replaces or clears the target-device bytes.
    pub fn set_target_device(&mut self, value: impl Into<Vec<u8>>) -> Result<(), OleError> {
        self.set_target_device_with_limits(value, Limits::default())
    }

    /// Replaces or clears the target-device bytes under explicit limits.
    pub fn set_target_device_with_limits(
        &mut self,
        value: impl Into<Vec<u8>>,
        limits: Limits,
    ) -> Result<(), OleError> {
        let limits = limits.validate()?;
        let value = value.into();
        if value.len() > limits.max_bytes {
            return Err(invalid(
                "TOCENTRY target device exceeds the configured limit",
            ));
        }
        validate_target_device(&value, "TOCENTRY target device")?;
        if self.target_device.replace(value) {
            self.dirty = true;
        }
        Ok(())
    }

    /// The implementation-specific rendering aspect hint.
    #[must_use]
    pub const fn aspect(&self) -> u32 {
        self.aspect
    }

    /// Sets the rendering aspect hint.
    pub const fn set_aspect(&mut self, value: u32) {
        if self.aspect != value {
            self.aspect = value;
            self.dirty = true;
        }
    }

    /// The implementation-specific rendering index hint.
    #[must_use]
    pub const fn lindex(&self) -> u32 {
        self.lindex
    }

    /// Sets the rendering index hint.
    pub const fn set_lindex(&mut self, value: u32) {
        if self.lindex != value {
            self.lindex = value;
            self.dirty = true;
        }
    }

    /// The uninterpreted TYMED field.
    #[must_use]
    pub const fn tymed(&self) -> u32 {
        self.tymed
    }

    /// Sets the uninterpreted TYMED field.
    pub const fn set_tymed(&mut self, value: u32) {
        if self.tymed != value {
            self.tymed = value;
            self.dirty = true;
        }
    }

    /// The exact reserved bytes between TYMED and Advf.
    #[must_use]
    pub const fn reserved1(&self) -> &[u8; 12] {
        &self.reserved1
    }

    /// The implementation-specific Advf hint.
    #[must_use]
    pub const fn advf(&self) -> u32 {
        self.advf
    }

    /// Sets the Advf hint.
    pub const fn set_advf(&mut self, value: u32) {
        if self.advf != value {
            self.advf = value;
            self.dirty = true;
        }
    }

    /// The exact reserved trailing field.
    #[must_use]
    pub const fn reserved2(&self) -> u32 {
        self.reserved2
    }

    /// Sets the reserved trailing field while retaining its wire position.
    pub const fn set_reserved2(&mut self, value: u32) {
        if self.reserved2 != value {
            self.reserved2 = value;
            self.dirty = true;
        }
    }

    fn encoded_len(&self) -> Result<usize, OleError> {
        let mut length = self.format.encoded_len()?;
        checked_len_add(&mut length, 4, "TOCENTRY target-device size")?;
        checked_len_add(
            &mut length,
            self.target_device.len(),
            "TOCENTRY target device",
        )?;
        checked_len_add(&mut length, 4 * 3 + 12 + 4 + 4, "TOCENTRY fixed fields")?;
        Ok(length)
    }

    fn validate(&self, limits: Limits) -> Result<(), OleError> {
        self.format.encoded_len()?;
        validate_format_marker(&self.format, self.format_marker)?;
        if self.target_device.len() > limits.max_bytes {
            return Err(invalid("TOCENTRY target device exceeds the stream limit"));
        }
        if !self.target_device.is_empty() && self.target_device.len() < 4 {
            return Err(invalid("TOCENTRY target device is truncated"));
        }
        if self.encoded_len()? > limits.max_bytes {
            return Err(invalid("TOCENTRY exceeds the configured stream limit"));
        }
        Ok(())
    }

    /// Serializes this entry, preserving source bytes when it is unedited.
    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        self.to_bytes_with_limits(Limits::default())
    }

    /// Serializes this entry under explicit resource limits.
    pub fn to_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>, OleError> {
        let limits = limits.validate()?;
        self.validate(limits)?;
        if !self.dirty {
            if self.wire.len() > limits.max_bytes {
                return Err(invalid("TOCENTRY exceeds the configured stream limit"));
            }
            return Ok(self.wire.bytes().to_vec());
        }
        let length = self.encoded_len()?;
        let mut output = Vec::new();
        reserve_output(&mut output, length, "OLEDS TOCENTRY")?;
        self.encode_into(&mut output, limits)?;
        debug_assert_eq!(output.len(), length);
        Ok(output)
    }

    fn encode_into(&self, output: &mut Vec<u8>, limits: Limits) -> Result<(), OleError> {
        self.validate(limits)?;
        self.format.encode_with_marker(self.format_marker, output)?;
        output.extend_from_slice(
            &ensure_u32(self.target_device.len(), "TOCENTRY target device")?.to_le_bytes(),
        );
        output.extend_from_slice(self.target_device.bytes());
        output.extend_from_slice(&self.aspect.to_le_bytes());
        output.extend_from_slice(&self.lindex.to_le_bytes());
        output.extend_from_slice(&self.tymed.to_le_bytes());
        output.extend_from_slice(&self.reserved1);
        output.extend_from_slice(&self.advf.to_le_bytes());
        output.extend_from_slice(&self.reserved2.to_le_bytes());
        Ok(())
    }
}

/// A typed OLEDS `OLEPresentationStream` (`\x02OlePres###`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OlePresentationStream {
    wire: Arc<[u8]>,
    format: ClipboardFormat,
    format_marker: u32,
    target_device: BytesField,
    aspect: u32,
    lindex: u32,
    advf: u32,
    reserved1: u32,
    width: i32,
    height: i32,
    data: BytesField,
    reserved2: Option<[u8; 18]>,
    toc_signature: Option<u32>,
    toc_count: Option<u32>,
    toc_entries: Vec<TocEntry>,
    unknown_tail: BytesField,
    dirty: bool,
}

impl OlePresentationStream {
    fn content_eq(&self, other: &Self) -> bool {
        self.wire == other.wire
            && self.format == other.format
            && self.format_marker == other.format_marker
            && self.target_device == other.target_device
            && self.aspect == other.aspect
            && self.lindex == other.lindex
            && self.advf == other.advf
            && self.reserved1 == other.reserved1
            && self.width == other.width
            && self.height == other.height
            && self.data == other.data
            && self.reserved2 == other.reserved2
            && self.toc_signature == other.toc_signature
            && self.toc_count == other.toc_count
            && self.unknown_tail == other.unknown_tail
            && self.toc_entries.len() == other.toc_entries.len()
            && self
                .toc_entries
                .iter()
                .zip(&other.toc_entries)
                .all(|(left, right)| left.content_eq(right))
    }

    /// Parses one complete OLEPresentationStream with default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parses one complete OLEPresentationStream under explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(invalid(
                "OLEPresentationStream exceeds the configured limit",
            ));
        }
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses one complete OLEPresentationStream without copying an existing
    /// source allocation.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(invalid(
                "OLEPresentationStream exceeds the configured limit",
            ));
        }
        let mut reader = Reader::new(&bytes);
        let (format, format_marker) =
            parse_clipboard_format(&mut reader, true, MAX_PRIMARY_REGISTERED_FORMAT_CHARS)?;
        if matches!(format, ClipboardFormat::Standard(CF_BITMAP)) {
            return Err(invalid("OLEPresentationStream cannot use CF_BITMAP"));
        }
        let target_size = reader.u32("OLEPresentationStream target-device size")?;
        let target = reader.sized_range(
            target_size,
            "OLEPresentationStream target device",
            true,
            true,
        )?;
        if let Some(range) = target.as_ref() {
            validate_target_device(&bytes[range.clone()], "OLEPresentationStream target device")?;
        }
        let aspect = reader.u32("OLEPresentationStream aspect")?;
        let lindex = reader.u32("OLEPresentationStream lindex")?;
        let advf = reader.u32("OLEPresentationStream advf")?;
        let reserved1 = reader.u32("OLEPresentationStream reserved1")?;
        let width = reader.i32("OLEPresentationStream width")?;
        let height = reader.i32("OLEPresentationStream height")?;
        let data_size = reader.u32("OLEPresentationStream data size")?;
        let data_len = usize::try_from(data_size)
            .map_err(|_| invalid("OLEPresentationStream data size exceeds this platform"))?;
        if data_len > limits.max_data_bytes {
            return Err(invalid(
                "OLEPresentationStream data exceeds the configured limit",
            ));
        }
        let data_start = reader.position();
        reader.take(data_len, "OLEPresentationStream data")?;
        let data = BytesField::from_source(Arc::clone(&bytes), data_start..reader.position());

        let reserved2 = if format.is_metafile_pict() {
            Some(reader.array::<18>("OLEPresentationStream metafile reserved bytes")?)
        } else {
            None
        };

        let mut toc_signature = None;
        let mut toc_count = None;
        let mut toc_entries = Vec::new();
        let unknown_tail_start = if reader.remaining() == 0 {
            bytes.len()
        } else {
            toc_signature = Some(reader.u32("OLEPresentationStream TOC signature")?);
            let count = reader.u32("OLEPresentationStream TOC count")?;
            toc_count = Some(count);
            if toc_signature == Some(TOC_SIGNATURE) {
                let count_usize = usize::try_from(count).map_err(|_| {
                    invalid("OLEPresentationStream TOC count exceeds this platform")
                })?;
                if count_usize > limits.max_toc_entries {
                    return Err(invalid(
                        "OLEPresentationStream TOC count exceeds the configured limit",
                    ));
                }
                if count_usize > reader.remaining() / MIN_TOC_ENTRY_BYTES {
                    return Err(invalid("OLEPresentationStream TOC entries are truncated"));
                }
                toc_count = Some(count);
                toc_entries
                    .try_reserve_exact(count_usize)
                    .map_err(|source| OleError::Allocation {
                        resource: "OLEDS TOC entries",
                        source,
                    })?;
                for index in 0..count_usize {
                    let start = reader.position();
                    let entry = parse_toc_entry_from_reader(&mut reader, &bytes, limits.max_bytes)?;
                    if entry.encoded_len()? > limits.max_bytes {
                        return Err(invalid(format!(
                            "OLEDS TOC entry {index} exceeds the configured limit"
                        )));
                    }
                    toc_entries.push(entry);
                    if reader.position() <= start {
                        return Err(invalid("OLEPresentationStream TOC entry has zero width"));
                    }
                }
            } else if count != 0 {
                return Err(invalid(
                    "non-NANI presentation TOC signatures must have a zero count",
                ));
            }
            reader.position()
        };

        let source = Arc::clone(&bytes);
        let wire_len = bytes.len();
        let stream = Self {
            wire: bytes,
            format,
            format_marker,
            target_device: target.map_or_else(BytesField::empty, |range| {
                BytesField::from_source(Arc::clone(&source), range)
            }),
            aspect,
            lindex,
            advf,
            reserved1,
            width,
            height,
            data,
            reserved2,
            toc_signature,
            toc_count,
            toc_entries,
            unknown_tail: BytesField::from_source(source, unknown_tail_start..wire_len),
            dirty: false,
        };
        stream.validate(limits)?;
        Ok(stream)
    }

    /// Creates a minimal presentation stream with opaque presentation data.
    pub fn new(format: ClipboardFormat, data: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        let data = data.into();
        validate_primary_format(&format)?;
        if matches!(
            format,
            ClipboardFormat::None | ClipboardFormat::Standard(CF_BITMAP)
        ) {
            return Err(invalid("primary OLEPresentationStream format is not valid"));
        }
        if data.len() > Limits::default().max_data_bytes {
            return Err(invalid("presentation data exceeds the default limit"));
        }
        let format_marker = default_format_marker(&format);
        let mut stream = Self {
            wire: Arc::from([]),
            format,
            format_marker,
            target_device: BytesField::empty(),
            aspect: 1,
            lindex: u32::MAX,
            advf: 0,
            reserved1: 0,
            width: 0,
            height: 0,
            data: BytesField::from_owned(data),
            reserved2: None,
            toc_signature: None,
            toc_count: None,
            toc_entries: Vec::new(),
            unknown_tail: BytesField::empty(),
            dirty: true,
        };
        if stream.format.is_metafile_pict() {
            stream.reserved2 = Some([0; 18]);
        }
        let bytes = stream.to_bytes_with_limits(Limits::default())?;
        Self::parse_shared(bytes.into(), Limits::default())
    }

    /// Exact source bytes, when this value has not been edited.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.wire
    }

    /// The clipboard format descriptor.
    #[must_use]
    pub fn clipboard_format(&self) -> &ClipboardFormat {
        &self.format
    }

    /// The exact MarkerOrLength value retained for the clipboard format.
    #[must_use]
    pub const fn clipboard_format_marker(&self) -> u32 {
        self.format_marker
    }

    /// Replaces the primary clipboard format.
    pub fn set_clipboard_format(&mut self, value: ClipboardFormat) -> Result<(), OleError> {
        validate_primary_format(&value)?;
        if matches!(
            value,
            ClipboardFormat::None | ClipboardFormat::Standard(CF_BITMAP)
        ) {
            return Err(invalid("primary OLEPresentationStream format is not valid"));
        }
        value.encoded_len()?;
        if self.format != value {
            self.format_marker = match (&value, self.format_marker) {
                (ClipboardFormat::Standard(_), marker @ (u32::MAX | 0xffff_fffe)) => marker,
                _ => default_format_marker(&value),
            };
            self.format = value;
            if self.format.is_metafile_pict() && self.reserved2.is_none() {
                self.reserved2 = Some([0; 18]);
            }
            if !self.format.is_metafile_pict() {
                self.reserved2 = None;
            }
            self.dirty = true;
        }
        Ok(())
    }

    /// Borrows the opaque DVTARGETDEVICE bytes.
    #[must_use]
    pub fn target_device(&self) -> &[u8] {
        self.target_device.bytes()
    }

    /// Replaces or clears the DVTARGETDEVICE bytes.
    pub fn set_target_device(&mut self, value: impl Into<Vec<u8>>) -> Result<(), OleError> {
        self.set_target_device_with_limits(value, Limits::default())
    }

    /// Replaces or clears the DVTARGETDEVICE bytes under explicit limits.
    pub fn set_target_device_with_limits(
        &mut self,
        value: impl Into<Vec<u8>>,
        limits: Limits,
    ) -> Result<(), OleError> {
        let limits = limits.validate()?;
        let value = value.into();
        if value.len() > limits.max_bytes {
            return Err(invalid(
                "presentation target device exceeds the configured limit",
            ));
        }
        validate_target_device(&value, "OLEPresentationStream target device")?;
        if self.target_device.replace(value) {
            self.dirty = true;
        }
        Ok(())
    }

    /// The implementation-specific rendering aspect hint.
    #[must_use]
    pub const fn aspect(&self) -> u32 {
        self.aspect
    }

    /// Sets the rendering aspect hint.
    pub const fn set_aspect(&mut self, value: u32) {
        if self.aspect != value {
            self.aspect = value;
            self.dirty = true;
        }
    }

    /// The implementation-specific rendering index hint.
    #[must_use]
    pub const fn lindex(&self) -> u32 {
        self.lindex
    }

    /// Sets the rendering index hint.
    pub const fn set_lindex(&mut self, value: u32) {
        if self.lindex != value {
            self.lindex = value;
            self.dirty = true;
        }
    }

    /// The implementation-specific Advf hint.
    #[must_use]
    pub const fn advf(&self) -> u32 {
        self.advf
    }

    /// Sets the Advf hint.
    pub const fn set_advf(&mut self, value: u32) {
        if self.advf != value {
            self.advf = value;
            self.dirty = true;
        }
    }

    /// The exact uninterpreted reserved value.
    #[must_use]
    pub const fn reserved1(&self) -> u32 {
        self.reserved1
    }

    /// Sets the reserved value while retaining its wire position.
    pub const fn set_reserved1(&mut self, value: u32) {
        if self.reserved1 != value {
            self.reserved1 = value;
            self.dirty = true;
        }
    }

    /// Presentation width in pixels, represented by the OLEDS signed `long`.
    #[must_use]
    pub const fn width(&self) -> i32 {
        self.width
    }

    /// Sets presentation width in pixels using the OLEDS signed `long` range.
    pub const fn set_width(&mut self, value: i32) {
        if self.width != value {
            self.width = value;
            self.dirty = true;
        }
    }

    /// Presentation height in pixels, represented by the OLEDS signed `long`.
    #[must_use]
    pub const fn height(&self) -> i32 {
        self.height
    }

    /// Sets presentation height in pixels using the OLEDS signed `long` range.
    pub const fn set_height(&mut self, value: i32) {
        if self.height != value {
            self.height = value;
            self.dirty = true;
        }
    }

    /// Borrows opaque presentation data.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        self.data.bytes()
    }

    /// Replaces opaque presentation data.
    pub fn set_data(&mut self, value: impl Into<Vec<u8>>) -> Result<(), OleError> {
        self.set_data_with_limits(value, Limits::default())
    }

    /// Replaces opaque presentation data under explicit limits.
    pub fn set_data_with_limits(
        &mut self,
        value: impl Into<Vec<u8>>,
        limits: Limits,
    ) -> Result<(), OleError> {
        let limits = limits.validate()?;
        let data = value.into();
        if data.len() > limits.max_data_bytes {
            return Err(invalid("presentation data exceeds the configured limit"));
        }
        if self.data.replace(data) {
            self.dirty = true;
        }
        Ok(())
    }

    /// The optional 18-byte CF_METAFILEPICT reserved field.
    #[must_use]
    pub fn reserved2(&self) -> Option<&[u8; 18]> {
        self.reserved2.as_ref()
    }

    /// The optional TOC signature.
    ///
    /// When present, the wire stream also carries a count.  A NANI signature
    /// permits that count's entries; another signature is retained as an
    /// unknown TOC marker and must carry a zero count.
    #[must_use]
    pub const fn toc_signature(&self) -> Option<u32> {
        self.toc_signature
    }

    /// The optional TOC count from the source stream.
    ///
    /// A non-NANI signature has a count of zero and no typed entries.
    #[must_use]
    pub const fn toc_count(&self) -> Option<u32> {
        self.toc_count
    }

    /// Borrows parsed TOC entries in source order.
    #[must_use]
    pub fn toc_entries(&self) -> &[TocEntry] {
        &self.toc_entries
    }

    /// Replaces one TOC entry atomically within this in-memory projection.
    pub fn update_toc_entry<F>(&mut self, index: usize, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut TocEntry) -> Result<(), OleError>,
    {
        self.update_toc_entry_with_limits(index, Limits::default(), edit)
    }

    /// Replaces one TOC entry atomically under explicit limits.
    pub fn update_toc_entry_with_limits<F>(
        &mut self,
        index: usize,
        limits: Limits,
        edit: F,
    ) -> Result<(), OleError>
    where
        F: FnOnce(&mut TocEntry) -> Result<(), OleError>,
    {
        let limits = limits.validate()?;
        let mut candidate =
            self.toc_entries.get(index).cloned().ok_or_else(|| {
                invalid("OLEPresentationStream TOC entry index is outside the table")
            })?;
        let before = candidate.clone();
        edit(&mut candidate)?;
        candidate.validate(limits)?;
        if !candidate.content_eq(&before) {
            self.toc_entries[index] = candidate;
            self.toc_count = Some(ensure_u32(
                self.toc_entries.len(),
                "presentation TOC count",
            )?);
            self.toc_signature = Some(TOC_SIGNATURE);
            self.dirty = true;
        }
        Ok(())
    }

    /// Inserts a TOC entry and authorizes a NANI TOC table when needed.
    pub fn insert_toc_entry(&mut self, index: usize, entry: TocEntry) -> Result<(), OleError> {
        self.insert_toc_entry_with_limits(index, entry, Limits::default())
    }

    /// Inserts a TOC entry under explicit stream limits.
    ///
    /// A table created by this operation uses the OLEDS `NANI` signature and
    /// an entry count matching the resulting table.  Existing unknown tail
    /// bytes remain attached to the source-backed presentation stream.
    pub fn insert_toc_entry_with_limits(
        &mut self,
        index: usize,
        entry: TocEntry,
        limits: Limits,
    ) -> Result<(), OleError> {
        let limits = limits.validate()?;
        if index > self.toc_entries.len() {
            return Err(invalid(
                "OLEPresentationStream TOC insertion index is outside the table",
            ));
        }
        entry.validate(limits)?;
        let mut candidate = self.clone();
        candidate.toc_entries.insert(index, entry);
        candidate.toc_signature = Some(TOC_SIGNATURE);
        candidate.toc_count = Some(ensure_u32(
            candidate.toc_entries.len(),
            "presentation TOC count",
        )?);
        candidate.dirty = true;
        candidate.validate(limits)?;
        *self = candidate;
        Ok(())
    }

    /// Removes one TOC entry while retaining a valid NANI table.
    pub fn remove_toc_entry(&mut self, index: usize) -> Result<TocEntry, OleError> {
        self.remove_toc_entry_with_limits(index, Limits::default())
    }

    /// Removes one TOC entry under explicit stream limits.
    ///
    /// Removing the final entry leaves an explicit `NANI` table with a zero
    /// count, which keeps the operation's table grammar unambiguous.
    pub fn remove_toc_entry_with_limits(
        &mut self,
        index: usize,
        limits: Limits,
    ) -> Result<TocEntry, OleError> {
        let limits = limits.validate()?;
        if index >= self.toc_entries.len() {
            return Err(invalid(
                "OLEPresentationStream TOC removal index is outside the table",
            ));
        }
        let mut candidate = self.clone();
        let removed = candidate.toc_entries.remove(index);
        candidate.toc_signature = Some(TOC_SIGNATURE);
        candidate.toc_count = Some(ensure_u32(
            candidate.toc_entries.len(),
            "presentation TOC count",
        )?);
        candidate.dirty = true;
        candidate.validate(limits)?;
        *self = candidate;
        Ok(removed)
    }

    /// The unparsed bytes after the recognized TOC fields.
    #[must_use]
    pub fn unknown_tail(&self) -> &[u8] {
        self.unknown_tail.bytes()
    }

    fn encoded_len(&self, limits: Limits) -> Result<usize, OleError> {
        let mut length = self.format.encoded_len()?;
        let target_size = self
            .target_device
            .len()
            .checked_add(4)
            .ok_or_else(|| invalid("presentation target-device size overflows"))?;
        checked_len_add(&mut length, 4, "presentation target-device size")?;
        checked_len_add(&mut length, target_size - 4, "presentation target device")?;
        checked_len_add(&mut length, 28, "presentation fixed fields")?;
        if self.data.len() > limits.max_data_bytes {
            return Err(invalid("presentation data exceeds the configured limit"));
        }
        checked_len_add(&mut length, self.data.len(), "presentation data")?;
        if self.format.is_metafile_pict() {
            if self.reserved2.is_none() {
                return Err(invalid("CF_METAFILEPICT presentation lacks Reserved2"));
            }
            checked_len_add(&mut length, 18, "presentation Reserved2")?;
        } else if self.reserved2.is_some() {
            return Err(invalid(
                "presentation Reserved2 is only valid for CF_METAFILEPICT",
            ));
        }
        if let Some(signature) = self.toc_signature {
            checked_len_add(&mut length, 4, "presentation TOC signature")?;
            self.toc_count_for_encoding(limits)?;
            checked_len_add(&mut length, 4, "presentation TOC count")?;
            if signature == TOC_SIGNATURE {
                for entry in &self.toc_entries {
                    entry.validate(limits)?;
                    checked_len_add(&mut length, entry.encoded_len()?, "presentation TOC entry")?;
                }
            }
            checked_len_add(
                &mut length,
                self.unknown_tail.len(),
                "presentation unknown tail",
            )?;
        } else if !self.toc_entries.is_empty() || self.toc_count.is_some() {
            return Err(invalid("presentation TOC entries have no signature"));
        } else {
            checked_len_add(
                &mut length,
                self.unknown_tail.len(),
                "presentation unknown tail",
            )?;
        }
        if length > limits.max_bytes {
            return Err(invalid(
                "OLEPresentationStream exceeds the configured limit",
            ));
        }
        Ok(length)
    }

    fn toc_count_for_encoding(&self, limits: Limits) -> Result<u32, OleError> {
        let Some(signature) = self.toc_signature else {
            if self.toc_count.is_some() || !self.toc_entries.is_empty() {
                return Err(invalid("presentation TOC entries have no signature"));
            }
            return Ok(0);
        };
        let count = self
            .toc_count
            .ok_or_else(|| invalid("presentation TOC signature has no count"))?;
        if signature == TOC_SIGNATURE {
            if self.toc_entries.len() > limits.max_toc_entries {
                return Err(invalid(
                    "presentation TOC count exceeds the configured limit",
                ));
            }
            let actual = ensure_u32(self.toc_entries.len(), "presentation TOC count")?;
            if count != actual {
                return Err(invalid("presentation TOC count does not match its entries"));
            }
            Ok(count)
        } else if count == 0 && self.toc_entries.is_empty() {
            Ok(count)
        } else {
            Err(invalid(
                "non-NANI presentation TOC signatures must have a zero count",
            ))
        }
    }

    fn validate(&self, limits: Limits) -> Result<(), OleError> {
        validate_primary_format(&self.format)?;
        self.format.encoded_len()?;
        validate_format_marker(&self.format, self.format_marker)?;
        if self.target_device.len() > limits.max_bytes {
            return Err(invalid(
                "presentation target device exceeds the configured limit",
            ));
        }
        if !self.target_device.is_empty() && self.target_device.len() < 4 {
            return Err(invalid("presentation target device is truncated"));
        }
        self.toc_count_for_encoding(limits)?;
        self.encoded_len(limits).map(|_| ())
    }

    /// Serializes this stream under the default resource limits.
    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        self.to_bytes_with_limits(Limits::default())
    }

    /// Serializes this stream under explicit resource limits.
    pub fn to_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>, OleError> {
        let limits = limits.validate()?;
        self.validate(limits)?;
        if !self.dirty {
            if self.wire.len() > limits.max_bytes {
                return Err(invalid(
                    "OLEPresentationStream exceeds the configured limit",
                ));
            }
            return Ok(self.wire.to_vec());
        }
        let length = self.encoded_len(limits)?;
        let mut output = Vec::new();
        reserve_output(&mut output, length, "OLEDS presentation stream")?;
        self.format
            .encode_with_marker(self.format_marker, &mut output)?;
        let target_size = self
            .target_device
            .len()
            .checked_add(4)
            .ok_or_else(|| invalid("presentation target-device size overflows"))?;
        output.extend_from_slice(
            &ensure_u32(target_size, "presentation target-device size")?.to_le_bytes(),
        );
        output.extend_from_slice(self.target_device.bytes());
        output.extend_from_slice(&self.aspect.to_le_bytes());
        output.extend_from_slice(&self.lindex.to_le_bytes());
        output.extend_from_slice(&self.advf.to_le_bytes());
        output.extend_from_slice(&self.reserved1.to_le_bytes());
        output.extend_from_slice(&self.width.to_le_bytes());
        output.extend_from_slice(&self.height.to_le_bytes());
        output.extend_from_slice(&ensure_u32(self.data.len(), "presentation data")?.to_le_bytes());
        output.extend_from_slice(self.data.bytes());
        if let Some(reserved2) = self.reserved2 {
            output.extend_from_slice(&reserved2);
        }
        if let Some(signature) = self.toc_signature {
            output.extend_from_slice(&signature.to_le_bytes());
            let count = self.toc_count_for_encoding(limits)?;
            output.extend_from_slice(&count.to_le_bytes());
            if signature == TOC_SIGNATURE {
                for entry in &self.toc_entries {
                    if entry.dirty {
                        entry.encode_into(&mut output, limits)?;
                    } else {
                        entry.validate(limits)?;
                        output.extend_from_slice(entry.wire.bytes());
                    }
                }
            }
            output.extend_from_slice(self.unknown_tail.bytes());
        } else {
            output.extend_from_slice(self.unknown_tail.bytes());
        }
        debug_assert_eq!(output.len(), length);
        Ok(output)
    }
}

fn parse_toc_entry_from_reader(
    reader: &mut Reader<'_>,
    source: &Arc<[u8]>,
    max_registered_chars: usize,
) -> Result<TocEntry, OleError> {
    let start = reader.position();
    let (format, format_marker) = parse_clipboard_format(reader, false, max_registered_chars)?;
    let target_size = reader.u32("TOCENTRY target-device size")?;
    let target = reader.sized_range(target_size, "TOCENTRY target device", false, false)?;
    if let Some(range) = target.as_ref() {
        validate_target_device(&source[range.clone()], "TOCENTRY target device")?;
    }
    let aspect = reader.u32("TOCENTRY aspect")?;
    let lindex = reader.u32("TOCENTRY lindex")?;
    let tymed = reader.u32("TOCENTRY tymed")?;
    let reserved1 = reader.array::<12>("TOCENTRY reserved bytes")?;
    let advf = reader.u32("TOCENTRY advf")?;
    let reserved2 = reader.u32("TOCENTRY reserved2")?;
    let end = reader.position();
    Ok(TocEntry {
        wire: BytesField::from_source(Arc::clone(source), start..end),
        format,
        format_marker,
        target_device: target.map_or_else(BytesField::empty, |range| {
            BytesField::from_source(Arc::clone(source), range)
        }),
        aspect,
        lindex,
        tymed,
        reserved1,
        advf,
        reserved2,
        dirty: false,
    })
}

/// A cheap immutable snapshot of an OLEPresentationStream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationSnapshot {
    stream: OlePresentationStream,
    limits: Limits,
    revision: StreamRevision,
}

impl PresentationSnapshot {
    /// Parses a source presentation stream with default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parses a source presentation stream under explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(invalid(
                "OLEPresentationStream exceeds the configured limit",
            ));
        }
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses a source presentation stream without copying its allocation.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        let stream = OlePresentationStream::parse_shared(bytes, limits)?;
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }

    /// Captures an already parsed stream as a source snapshot.
    pub fn from_stream(stream: OlePresentationStream, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        stream.validate(limits)?;
        if stream.bytes().len() > limits.max_bytes {
            return Err(invalid(
                "OLEPresentationStream exceeds the configured limit",
            ));
        }
        if stream.dirty {
            let bytes = stream.to_bytes_with_limits(limits)?;
            return Self::parse_shared(bytes.into(), limits);
        }
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }

    /// Borrows the typed presentation projection.
    #[must_use]
    pub const fn stream(&self) -> &OlePresentationStream {
        &self.stream
    }

    /// Exact source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.stream.wire)
    }

    /// Source identity used by reversible patches.
    #[must_use]
    pub const fn revision(&self) -> StreamRevision {
        self.revision
    }

    /// Source fingerprint used by reversible patches.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Limits retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Starts an isolated transaction.
    #[must_use]
    pub fn edit(&self) -> PresentationTransaction {
        PresentationTransaction {
            source: self.clone(),
            candidate: self.stream.clone(),
        }
    }

    fn patch_to(&self, after: &Self) -> PresentationPatch {
        PresentationPatch::new(self.clone(), after.clone())
    }
}

impl std::ops::Deref for PresentationSnapshot {
    type Target = OlePresentationStream;

    fn deref(&self) -> &Self::Target {
        self.stream()
    }
}

/// A failure-atomic edit over one presentation snapshot.
#[derive(Debug, Clone)]
pub struct PresentationTransaction {
    source: PresentationSnapshot,
    candidate: OlePresentationStream,
}

impl PresentationTransaction {
    /// Borrows the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &PresentationSnapshot {
        &self.source
    }

    /// Borrows the current typed candidate.
    #[must_use]
    pub const fn stream(&self) -> &OlePresentationStream {
        &self.candidate
    }

    /// Whether the candidate differs from its source projection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.candidate.content_eq(&self.source.stream)
    }

    /// Applies one callback to a cloned candidate and validates it before
    /// publication.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut OlePresentationStream) -> Result<(), OleError>,
    {
        let mut candidate = self.candidate.clone();
        edit(&mut candidate)?;
        candidate.validate(self.source.limits)?;
        self.candidate = candidate;
        Ok(self)
    }

    /// Stages an opaque data replacement.
    pub fn set_data(&mut self, value: impl Into<Vec<u8>>) -> Result<&mut Self, OleError> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_data_with_limits(value, limits))
    }

    /// Stages a primary clipboard-format replacement.
    pub fn set_clipboard_format(&mut self, value: ClipboardFormat) -> Result<&mut Self, OleError> {
        self.update(|stream| stream.set_clipboard_format(value))
    }

    /// Stages a target-device replacement.
    pub fn set_target_device(&mut self, value: impl Into<Vec<u8>>) -> Result<&mut Self, OleError> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_target_device_with_limits(value, limits))
    }

    /// Stages an aspect-hint replacement.
    pub fn set_aspect(&mut self, value: u32) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            stream.set_aspect(value);
            Ok(())
        })
    }

    /// Stages an Lindex-hint replacement.
    pub fn set_lindex(&mut self, value: u32) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            stream.set_lindex(value);
            Ok(())
        })
    }

    /// Stages an Advf-hint replacement.
    pub fn set_advf(&mut self, value: u32) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            stream.set_advf(value);
            Ok(())
        })
    }

    /// Stages a width replacement.
    pub fn set_width(&mut self, value: i32) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            stream.set_width(value);
            Ok(())
        })
    }

    /// Stages a height replacement.
    pub fn set_height(&mut self, value: i32) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            stream.set_height(value);
            Ok(())
        })
    }

    /// Stages a TOC entry edit.
    pub fn update_toc_entry<F>(&mut self, index: usize, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut TocEntry) -> Result<(), OleError>,
    {
        let limits = self.source.limits;
        self.update(|stream| stream.update_toc_entry_with_limits(index, limits, edit))
    }

    /// Stages insertion of a TOC entry.
    pub fn insert_toc_entry(
        &mut self,
        index: usize,
        entry: TocEntry,
    ) -> Result<&mut Self, OleError> {
        let limits = self.source.limits;
        self.update(|stream| stream.insert_toc_entry_with_limits(index, entry, limits))
    }

    /// Stages removal of a TOC entry.
    pub fn remove_toc_entry(&mut self, index: usize) -> Result<TocEntry, OleError> {
        let limits = self.source.limits;
        let mut removed = None;
        self.update(|stream| {
            removed = Some(stream.remove_toc_entry_with_limits(index, limits)?);
            Ok(())
        })?;
        removed.ok_or_else(|| invalid("OLEPresentationStream TOC removal produced no entry"))
    }

    /// Projects the current candidate as a checked snapshot.
    pub fn snapshot(&self) -> Result<PresentationSnapshot, OleError> {
        self.materialize()
    }

    /// Discards this transaction and returns its source snapshot.
    #[must_use]
    pub fn rollback(self) -> PresentationSnapshot {
        self.source
    }

    /// Validates and publishes this transaction.
    pub fn commit(self) -> Result<PresentationCommit, OleError> {
        let snapshot = self.materialize()?;
        let patch = self.source.patch_to(&snapshot);
        Ok(PresentationCommit { snapshot, patch })
    }

    fn materialize(&self) -> Result<PresentationSnapshot, OleError> {
        self.candidate.validate(self.source.limits)?;
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let bytes = self.candidate.to_bytes_with_limits(self.source.limits)?;
        PresentationSnapshot::parse_shared(bytes.into(), self.source.limits)
    }
}

/// A successful presentation transaction publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationCommit {
    snapshot: PresentationSnapshot,
    patch: PresentationPatch,
}

impl PresentationCommit {
    /// Whether the exact stream bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &PresentationSnapshot {
        &self.snapshot
    }

    /// Borrows the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &PresentationPatch {
        &self.patch
    }

    /// Consumes this result into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> PresentationSnapshot {
        self.snapshot
    }

    /// Consumes this result into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> PresentationPatch {
        self.patch
    }

    /// Splits this publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (PresentationSnapshot, PresentationPatch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, exact-source-checked presentation replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationPatch {
    base: StreamRevision,
    target: StreamRevision,
    before: PresentationSnapshot,
    after: PresentationSnapshot,
    changed: bool,
}

impl PresentationPatch {
    fn new(before: PresentationSnapshot, after: PresentationSnapshot) -> Self {
        let changed = before.bytes() != after.bytes();
        Self {
            base: before.revision,
            target: after.revision,
            before,
            after,
            changed,
        }
    }

    /// Expected source revision.
    #[must_use]
    pub const fn base(&self) -> StreamRevision {
        self.base
    }

    /// Resulting revision.
    #[must_use]
    pub const fn target(&self) -> StreamRevision {
        self.target
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Borrows the exact source snapshot retained by this patch.
    #[must_use]
    pub const fn source(&self) -> &PresentationSnapshot {
        &self.before
    }

    /// Borrows the exact replacement snapshot retained by this patch.
    #[must_use]
    pub const fn replacement(&self) -> &PresentationSnapshot {
        &self.after
    }

    /// Whether the replacement is byte-for-byte empty.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        !self.changed
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &PresentationSnapshot) -> Result<PresentationSnapshot, OleError> {
        if source.revision != self.base || source.bytes() != self.before.bytes() {
            return Err(invalid(
                "OLEDS presentation patch source does not match its base",
            ));
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }
}

/// A typed OLEDS `OLENativeStream` (`\x01Ole10Native`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleNativeStream {
    wire: Arc<[u8]>,
    data: BytesField,
    dirty: bool,
}

impl OleNativeStream {
    fn content_eq(&self, other: &Self) -> bool {
        self.wire == other.wire && self.data == other.data
    }

    /// Parses a native stream under default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parses a native stream under explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(invalid("OLENativeStream exceeds the configured limit"));
        }
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses a native stream without copying an existing source allocation.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes || bytes.len() < 4 {
            return Err(invalid(
                "OLENativeStream is truncated or exceeds the configured limit",
            ));
        }
        let mut reader = Reader::new(&bytes);
        let size = usize::try_from(reader.u32("OLENativeStream native-data size")?)
            .map_err(|_| invalid("OLENativeStream native-data size exceeds this platform"))?;
        if size > limits.max_data_bytes {
            return Err(invalid("OLENativeStream data exceeds the configured limit"));
        }
        let data_start = reader.position();
        reader.take(size, "OLENativeStream native data")?;
        if reader.remaining() != 0 {
            return Err(invalid(
                "OLENativeStream has trailing bytes after NativeData",
            ));
        }
        let data_end = reader.position();
        let source = Arc::clone(&bytes);
        Ok(Self {
            wire: bytes,
            data: BytesField::from_source(source, data_start..data_end),
            dirty: false,
        })
    }

    /// Creates a native stream from opaque bytes.
    pub fn new(data: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        let data = data.into();
        if data.len() > Limits::default().max_data_bytes {
            return Err(invalid("native data exceeds the default limit"));
        }
        let mut output = Vec::new();
        reserve_output(&mut output, 4 + data.len(), "OLEDS native stream")?;
        output.extend_from_slice(&ensure_u32(data.len(), "native data")?.to_le_bytes());
        output.extend_from_slice(&data);
        let wire: Arc<[u8]> = output.into();
        Ok(Self {
            data: BytesField::from_source(Arc::clone(&wire), 4..wire.len()),
            wire,
            dirty: false,
        })
    }

    /// Exact source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.wire
    }

    /// Borrows opaque native data.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        self.data.bytes()
    }

    /// Replaces opaque native data.
    pub fn set_data(&mut self, value: impl Into<Vec<u8>>) -> Result<(), OleError> {
        self.set_data_with_limits(value, Limits::default())
    }

    /// Replaces opaque native data under explicit limits.
    pub fn set_data_with_limits(
        &mut self,
        value: impl Into<Vec<u8>>,
        limits: Limits,
    ) -> Result<(), OleError> {
        let limits = limits.validate()?;
        let value = value.into();
        if value.len() > limits.max_data_bytes {
            return Err(invalid("native data exceeds the configured limit"));
        }
        if self.data.replace(value) {
            self.dirty = true;
        }
        Ok(())
    }

    fn encoded_len(&self, limits: Limits) -> Result<usize, OleError> {
        if self.data.len() > limits.max_data_bytes {
            return Err(invalid("native data exceeds the configured limit"));
        }
        let length = 4usize
            .checked_add(self.data.len())
            .ok_or_else(|| invalid("native stream serialized size overflows"))?;
        if length > limits.max_bytes {
            return Err(invalid("OLENativeStream exceeds the configured limit"));
        }
        Ok(length)
    }

    /// Serializes this stream under the default resource limits.
    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        self.to_bytes_with_limits(Limits::default())
    }

    /// Serializes this stream under explicit resource limits.
    pub fn to_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>, OleError> {
        let limits = limits.validate()?;
        self.encoded_len(limits)?;
        if !self.dirty {
            if self.wire.len() > limits.max_bytes {
                return Err(invalid("OLENativeStream exceeds the configured limit"));
            }
            return Ok(self.wire.to_vec());
        }
        let length = self.encoded_len(limits)?;
        let mut output = Vec::new();
        reserve_output(&mut output, length, "OLEDS native stream")?;
        output.extend_from_slice(&ensure_u32(self.data.len(), "native data")?.to_le_bytes());
        output.extend_from_slice(self.data.bytes());
        Ok(output)
    }
}

/// A cheap immutable snapshot of an OLENativeStream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeSnapshot {
    stream: OleNativeStream,
    limits: Limits,
    revision: StreamRevision,
}

impl NativeSnapshot {
    /// Parses a native stream with default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parses a native stream with explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(invalid("OLENativeStream exceeds the configured limit"));
        }
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses a native stream without copying the source allocation.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        let stream = OleNativeStream::parse_shared(bytes, limits)?;
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }

    /// Borrows the typed native stream.
    #[must_use]
    pub const fn stream(&self) -> &OleNativeStream {
        &self.stream
    }

    /// Exact source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Shared ownership of exact source bytes.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.stream.wire)
    }

    /// Source identity used by patches.
    #[must_use]
    pub const fn revision(&self) -> StreamRevision {
        self.revision
    }

    /// Source fingerprint used by patches.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Limits retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Starts an isolated native-data edit.
    #[must_use]
    pub fn edit(&self) -> NativeTransaction {
        NativeTransaction {
            source: self.clone(),
            candidate: self.stream.clone(),
        }
    }

    /// Captures an already parsed native stream as a source snapshot.
    pub fn from_stream(stream: OleNativeStream, limits: Limits) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        stream.encoded_len(limits)?;
        if stream.bytes().len() > limits.max_bytes {
            return Err(invalid("OLENativeStream exceeds the configured limit"));
        }
        if stream.dirty {
            let bytes = stream.to_bytes_with_limits(limits)?;
            return Self::parse_shared(bytes.into(), limits);
        }
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }
}

impl std::ops::Deref for NativeSnapshot {
    type Target = OleNativeStream;

    fn deref(&self) -> &Self::Target {
        self.stream()
    }
}

/// A failure-atomic edit over one native snapshot.
#[derive(Debug, Clone)]
pub struct NativeTransaction {
    source: NativeSnapshot,
    candidate: OleNativeStream,
}

impl NativeTransaction {
    /// Borrows the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &NativeSnapshot {
        &self.source
    }

    /// Borrows the current candidate.
    #[must_use]
    pub const fn stream(&self) -> &OleNativeStream {
        &self.candidate
    }

    /// Whether the candidate differs from its source projection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.candidate.content_eq(&self.source.stream)
    }

    /// Applies a custom edit to a cloned candidate and validates it before
    /// publishing it into this transaction.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut OleNativeStream) -> Result<(), OleError>,
    {
        let mut candidate = self.candidate.clone();
        edit(&mut candidate)?;
        candidate.encoded_len(self.source.limits)?;
        self.candidate = candidate;
        Ok(self)
    }

    /// Stages an opaque native-data replacement.
    pub fn set_data(&mut self, value: impl Into<Vec<u8>>) -> Result<&mut Self, OleError> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_data_with_limits(value, limits))
    }

    /// Projects the current candidate as a checked snapshot.
    pub fn snapshot(&self) -> Result<NativeSnapshot, OleError> {
        self.materialize()
    }

    /// Discards this transaction and returns its source.
    #[must_use]
    pub fn rollback(self) -> NativeSnapshot {
        self.source
    }

    /// Validates and publishes this transaction.
    pub fn commit(self) -> Result<NativeCommit, OleError> {
        let snapshot = self.materialize()?;
        let patch = NativePatch::new(self.source, snapshot.clone());
        Ok(NativeCommit { snapshot, patch })
    }

    fn materialize(&self) -> Result<NativeSnapshot, OleError> {
        self.candidate.encoded_len(self.source.limits)?;
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let bytes = self.candidate.to_bytes_with_limits(self.source.limits)?;
        NativeSnapshot::parse_shared(bytes.into(), self.source.limits)
    }
}

/// A successful native-data transaction publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeCommit {
    snapshot: NativeSnapshot,
    patch: NativePatch,
}

impl NativeCommit {
    /// Whether exact source bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &NativeSnapshot {
        &self.snapshot
    }

    /// Borrows the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &NativePatch {
        &self.patch
    }

    /// Consumes this result into its published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> NativeSnapshot {
        self.snapshot
    }

    /// Consumes this result into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> NativePatch {
        self.patch
    }

    /// Splits this publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (NativeSnapshot, NativePatch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, exact-source-checked native-data replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativePatch {
    base: StreamRevision,
    target: StreamRevision,
    before: NativeSnapshot,
    after: NativeSnapshot,
    changed: bool,
}

impl NativePatch {
    fn new(before: NativeSnapshot, after: NativeSnapshot) -> Self {
        let changed = before.bytes() != after.bytes();
        Self {
            base: before.revision,
            target: after.revision,
            before,
            after,
            changed,
        }
    }

    /// Expected source revision.
    #[must_use]
    pub const fn base(&self) -> StreamRevision {
        self.base
    }

    /// Resulting revision.
    #[must_use]
    pub const fn target(&self) -> StreamRevision {
        self.target
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Borrows the exact source snapshot retained by this patch.
    #[must_use]
    pub const fn source(&self) -> &NativeSnapshot {
        &self.before
    }

    /// Borrows the exact replacement snapshot retained by this patch.
    #[must_use]
    pub const fn replacement(&self) -> &NativeSnapshot {
        &self.after
    }

    /// Whether this patch is a byte-for-byte no-op.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        !self.changed
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &NativeSnapshot) -> Result<NativeSnapshot, OleError> {
        if source.revision != self.base || source.bytes() != self.before.bytes() {
            return Err(invalid("OLEDS native patch source does not match its base"));
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }
}

/// Alias following the stream's specification name.
pub type OLEPresentationStream = OlePresentationStream;
/// Concise alias for [`OlePresentationStream`].
pub type PresentationStream = OlePresentationStream;
/// Alias following the stream's specification name.
pub type OLENativeStream = OleNativeStream;
/// Concise alias for [`OleNativeStream`].
pub type NativeStream = OleNativeStream;

/// Alias for the presentation snapshot.
pub type Snapshot = PresentationSnapshot;
/// Alias for the native snapshot.
pub type NativeDataSnapshot = NativeSnapshot;

/// Presentation-stream namespaced exports.
pub mod presentation {
    pub use super::{
        ClipboardFormat, Limits, OLEPresentationStream, OlePresentationStream, PresentationCommit,
        PresentationPatch, PresentationSnapshot, PresentationStream, PresentationTransaction,
        StreamRevision, TargetDevice, TocEntry,
    };
}

/// Native-stream namespaced exports.
pub mod native {
    pub use super::{
        Limits, NativeCommit, NativePatch, NativeSnapshot, NativeStream, NativeTransaction,
        OLENativeStream, OleNativeStream, StreamRevision,
    };
}

/// TOCENTRY namespaced exports.
pub mod toc_entry {
    pub use super::{ClipboardFormat, TargetDevice, TocEntry};
}

/// Parses a named `\x02OlePres###` stream and checks its OLEDS stream name.
pub fn parse_named_presentation(
    name: &str,
    bytes: Arc<[u8]>,
    limits: Limits,
) -> Result<PresentationSnapshot, OleError> {
    presentation_index(name)?;
    PresentationSnapshot::parse_shared(bytes, limits)
}

/// Parses the named `\x01Ole10Native` stream.
pub fn parse_named_native(
    name: &str,
    bytes: Arc<[u8]>,
    limits: Limits,
) -> Result<NativeSnapshot, OleError> {
    if name != NATIVE_STREAM_NAME {
        return Err(invalid("stream name is not \\x01Ole10Native"));
    }
    NativeSnapshot::parse_shared(bytes, limits)
}

/// Returns the numeric presentation index from a valid `\x02OlePres###` name.
pub fn presentation_index(name: &str) -> Result<usize, OleError> {
    let suffix = name
        .strip_prefix(PRESENTATION_STREAM_PREFIX)
        .ok_or_else(|| invalid("stream name is not an OLEDS presentation stream"))?;
    if suffix.len() != 3 || !suffix.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(
            "OLEDS presentation stream name must end in three digits",
        ));
    }
    let index = suffix
        .parse::<usize>()
        .map_err(|_| invalid("OLEDS presentation stream index is invalid"))?;
    if index > MAX_PRESENTATION_INDEX {
        return Err(invalid("OLEDS presentation stream index exceeds 999"));
    }
    Ok(index)
}

/// Builds a canonical OLEDS presentation stream name for an index.
pub fn presentation_name(index: usize) -> Result<String, OleError> {
    if index > MAX_PRESENTATION_INDEX {
        return Err(invalid("OLEDS presentation stream index exceeds 999"));
    }
    Ok(format!("{PRESENTATION_STREAM_PREFIX}{index:03}"))
}

impl fmt::Display for ClipboardFormat {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => formatter.write_str("none"),
            Self::Standard(value) => write!(formatter, "standard({value:#x})"),
            Self::Registered(bytes) => write!(formatter, "registered({} bytes)", bytes.len()),
        }
    }
}
