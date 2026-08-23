#![allow(
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines an intentional opaque or future-variant fallback to this codec boundary"
)]

//! BIFF12 record framing.

use std::fmt;
use std::io::{self, Read};

use super::{Error, LimitResource, Result, Stage};

/// Largest value representable by the four-byte BIFF12 record-size field.
pub const MAX_WIRE_PAYLOAD: usize = 0x0fff_ffff;

/// Largest value representable by the four-byte `XLWideString` unit count.
pub const MAX_WIRE_STRING_UNITS: usize = u32::MAX as usize;

/// Finite resource limits for one raw operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    payload: usize,
    string_units: usize,
}

impl Limits {
    /// Safe default for ordinary document processing.
    pub const DEFAULT: Self = Self {
        payload: 64 * 1024 * 1024,
        string_units: 1_048_576,
    };

    /// Construct explicit payload and UTF-16 code-unit limits.
    #[must_use]
    pub const fn new(payload: usize, string_units: usize) -> Self {
        Self {
            payload,
            string_units,
        }
    }

    /// Construct and validate explicit raw ceilings.
    pub const fn try_new(payload: usize, string_units: usize) -> Result<Self> {
        Self::new(payload, string_units).validate()
    }

    /// Validate the physical bounds of this raw limit profile.
    pub const fn validate(self) -> Result<Self> {
        if self.payload > MAX_WIRE_PAYLOAD {
            return Err(Error::InvalidLimit {
                resource: LimitResource::Payload,
                value: self.payload,
                maximum: MAX_WIRE_PAYLOAD,
            });
        }
        if self.string_units > MAX_WIRE_STRING_UNITS {
            return Err(Error::InvalidLimit {
                resource: LimitResource::StringUnits,
                value: self.string_units,
                maximum: MAX_WIRE_STRING_UNITS,
            });
        }
        Ok(self)
    }

    /// Maximum accepted bytes in one record payload.
    #[must_use]
    pub const fn payload(self) -> usize {
        self.payload
    }

    /// Maximum accepted UTF-16 code units in one string.
    #[must_use]
    pub const fn string_units(self) -> usize {
        self.string_units
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Validated 14-bit BIFF12 record kind.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Kind(pub(super) u16);

impl Kind {
    /// Largest legal record kind.
    pub const MAX: u16 = 0x3fff;

    /// Validate a numeric record kind.
    pub const fn new(value: u16) -> Result<Self> {
        if value <= Self::MAX {
            Ok(Self(value))
        } else {
            Err(Error::KindOutOfRange { value })
        }
    }

    /// Return the numeric wire value.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl TryFrom<u16> for Kind {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        Self::new(value)
    }
}

impl From<Kind> for u16 {
    fn from(value: Kind) -> Self {
        value.get()
    }
}

impl fmt::Display for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "0x{:04X}", self.get())
    }
}

impl fmt::LowerHex for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::LowerHex::fmt(&self.get(), formatter)
    }
}

impl fmt::UpperHex for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::UpperHex::fmt(&self.get(), formatter)
    }
}

/// Validated BIFF12 record header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Header {
    /// Typed record kind.
    kind: Kind,
    /// Declared payload length.
    len: usize,
}

impl Header {
    /// Typed record kind.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Declared payload length.
    #[must_use]
    pub const fn len(self) -> usize {
        self.len
    }

    /// Read one header from a stream using explicit limits.
    ///
    /// `Ok(None)` means the stream was already at a clean record boundary.
    /// Once any header byte is consumed, premature EOF is a typed truncation.
    pub fn read<R: Read>(reader: &mut R, limits: Limits) -> Result<Option<Self>> {
        limits.validate()?;
        let Some(first) = read_first(reader)? else {
            return Ok(None);
        };
        let (kind, kind_len) = read_kind_from(reader, first, 0)?;
        let len = read_len_from(reader, limits, kind_len)?;
        Ok(Some(Self { kind, len }))
    }

    /// Decode one header from the start of a byte slice.
    pub fn parse(input: &[u8], limits: Limits) -> Result<(Self, usize)> {
        limits.validate()?;
        let first = input.first().copied().ok_or(Error::Truncated {
            stage: Stage::Kind,
            offset: 0,
            needed: 1,
            available: 0,
        })?;
        let (kind, kind_len) = parse_kind(input, first)?;
        let (len, len_len) = parse_len(input, kind_len, limits)?;
        let consumed = kind_len.checked_add(len_len).ok_or(Error::LengthOverflow {
            what: "record header",
            length: usize::MAX,
        })?;
        Ok((Self { kind, len }, consumed))
    }

    /// Whether the payload is empty.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.len == 0
    }
}

/// One borrowed BIFF12 record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Record<'a> {
    /// Validated header.
    header: Header,
    /// Payload lent directly from the input slice.
    payload: &'a [u8],
    offset: usize,
}

impl<'a> Record<'a> {
    /// Typed record kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.header.kind()
    }

    /// Validated payload length.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.header.len()
    }

    /// Whether the payload is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.header.is_empty()
    }

    /// Payload lent directly from the input slice.
    #[must_use]
    pub const fn payload(&self) -> &'a [u8] {
        self.payload
    }

    /// Byte offset of the record header in the source slice.
    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }
}

/// Zero-copy iterator over a complete BIFF12 record stream.
pub struct Records<'a> {
    input: &'a [u8],
    offset: usize,
    limits: Limits,
    failed: bool,
}

impl<'a> Records<'a> {
    /// Iterate with safe default limits.
    #[must_use]
    pub const fn new(input: &'a [u8]) -> Self {
        Self::with_limits(input, Limits::DEFAULT)
    }

    /// Iterate with explicit finite limits.
    #[must_use]
    pub const fn with_limits(input: &'a [u8], limits: Limits) -> Self {
        Self {
            input,
            offset: 0,
            limits,
            failed: false,
        }
    }

    /// Construct an iterator after validating its raw limit profile.
    pub fn try_with_limits(input: &'a [u8], limits: Limits) -> Result<Self> {
        match limits.validate() {
            Ok(limits) => Ok(Self::with_limits(input, limits)),
            Err(error) => Err(error),
        }
    }

    /// Current byte offset at the next record boundary.
    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }

    fn fail(&mut self, error: Error) -> Option<Result<Record<'a>>> {
        self.failed = true;
        Some(Err(error))
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<Record<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed {
            return None;
        }
        if let Err(error) = self.limits.validate() {
            return self.fail(error);
        }
        if self.offset == self.input.len() {
            return None;
        }
        let tail = match self.input.get(self.offset..) {
            Some(tail) => tail,
            None => {
                return self.fail(Error::Truncated {
                    stage: Stage::Kind,
                    offset: self.offset,
                    needed: 1,
                    available: 0,
                });
            },
        };
        let (header, header_len) = match Header::parse(tail, self.limits) {
            Ok(value) => value,
            Err(error) => return self.fail(shift_error(error, self.offset)),
        };
        let payload_start = match self.offset.checked_add(header_len) {
            Some(value) => value,
            None => {
                return self.fail(Error::LengthOverflow {
                    what: "record offset",
                    length: usize::MAX,
                });
            },
        };
        let payload_end = match payload_start.checked_add(header.len()) {
            Some(value) => value,
            None => {
                return self.fail(Error::LengthOverflow {
                    what: "record payload",
                    length: header.len(),
                });
            },
        };
        let payload = match self.input.get(payload_start..payload_end) {
            Some(payload) => payload,
            None => {
                let available = self.input.len().saturating_sub(payload_start);
                return self.fail(Error::Truncated {
                    stage: Stage::Payload,
                    offset: payload_start,
                    needed: header.len(),
                    available,
                });
            },
        };
        let offset = self.offset;
        self.offset = payload_end;
        Some(Ok(Record {
            header,
            payload,
            offset,
        }))
    }
}

fn read_first(reader: &mut impl Read) -> Result<Option<u8>> {
    let mut byte = [0_u8; 1];
    loop {
        match reader.read(&mut byte) {
            Ok(0) => return Ok(None),
            Ok(_) => return Ok(Some(byte[0])),
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => return Err(Error::Io(error)),
        }
    }
}

fn read_required(reader: &mut impl Read, stage: Stage, offset: usize) -> Result<u8> {
    let mut byte = [0_u8; 1];
    reader.read_exact(&mut byte).map_err(|error| {
        if error.kind() == io::ErrorKind::UnexpectedEof {
            Error::Truncated {
                stage,
                offset,
                needed: 1,
                available: 0,
            }
        } else {
            Error::Io(error)
        }
    })?;
    Ok(byte[0])
}

fn read_kind_from(reader: &mut impl Read, first: u8, offset: usize) -> Result<(Kind, usize)> {
    if first & 0x80 == 0 {
        return Ok((Kind(u16::from(first)), 1));
    }
    let second = read_required(reader, Stage::Kind, offset.saturating_add(1))?;
    if second & 0x80 != 0 {
        return Err(Error::InvalidKind { offset });
    }
    let value = u16::from(first & 0x7f) | (u16::from(second) << 7);
    if value < 0x80 {
        return Err(Error::InvalidKind { offset });
    }
    Ok((Kind(value), 2))
}

fn read_len_from(reader: &mut impl Read, limits: Limits, offset: usize) -> Result<usize> {
    let mut value = 0_usize;
    for index in 0..4_usize {
        let byte = read_required(reader, Stage::Length, offset.saturating_add(index))?;
        value |= usize::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 || index == 3 {
            if value > limits.payload() {
                return Err(Error::PayloadLimit {
                    length: value,
                    limit: limits.payload(),
                    offset,
                });
            }
            return Ok(value);
        }
    }
    Err(Error::Truncated {
        stage: Stage::Length,
        offset,
        needed: 1,
        available: 0,
    })
}

fn parse_kind(input: &[u8], first: u8) -> Result<(Kind, usize)> {
    if first & 0x80 == 0 {
        return Ok((Kind(u16::from(first)), 1));
    }
    let second = input.get(1).copied().ok_or(Error::Truncated {
        stage: Stage::Kind,
        offset: 1,
        needed: 1,
        available: 0,
    })?;
    if second & 0x80 != 0 {
        return Err(Error::InvalidKind { offset: 0 });
    }
    let value = u16::from(first & 0x7f) | (u16::from(second) << 7);
    if value < 0x80 {
        return Err(Error::InvalidKind { offset: 0 });
    }
    Ok((Kind(value), 2))
}

fn parse_len(input: &[u8], start: usize, limits: Limits) -> Result<(usize, usize)> {
    let mut value = 0_usize;
    for index in 0..4_usize {
        let position = start.checked_add(index).ok_or(Error::LengthOverflow {
            what: "record header",
            length: usize::MAX,
        })?;
        let byte = input.get(position).copied().ok_or(Error::Truncated {
            stage: Stage::Length,
            offset: position,
            needed: 1,
            available: input.len().saturating_sub(position),
        })?;
        value |= usize::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 || index == 3 {
            if value > limits.payload() {
                return Err(Error::PayloadLimit {
                    length: value,
                    limit: limits.payload(),
                    offset: start,
                });
            }
            return Ok((value, index + 1));
        }
    }
    Err(Error::Truncated {
        stage: Stage::Length,
        offset: start,
        needed: 1,
        available: 0,
    })
}

fn shift_error(error: Error, base: usize) -> Error {
    match error {
        Error::Truncated {
            stage,
            offset,
            needed,
            available,
        } => Error::Truncated {
            stage,
            offset: base.saturating_add(offset),
            needed,
            available,
        },
        Error::InvalidKind { offset } => Error::InvalidKind {
            offset: base.saturating_add(offset),
        },
        Error::PayloadLimit {
            length,
            limit,
            offset,
        } => Error::PayloadLimit {
            length,
            limit,
            offset: base.saturating_add(offset),
        },
        other => other,
    }
}
