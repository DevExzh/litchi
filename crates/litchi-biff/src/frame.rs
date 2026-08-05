use std::fmt;

use crate::limits::{HEADER_BYTES, as_u64};
use crate::stream::Records;
use crate::{Error, Limits, Resource, Result};

/// A checked two-byte BIFF record identifier.
///
/// The frame layer intentionally does not enumerate record kinds. Unknown
/// values are valid preservation data and are interpreted by the owning format
/// crate. Values read from the wire are therefore always representable; the
/// fallible conversions protect callers moving wider integer values into this
/// exact wire domain.
#[derive(Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Kind(u16);

impl Kind {
    /// Creates a kind from its exact on-wire unsigned 16-bit value.
    #[must_use]
    pub const fn from_wire(value: u16) -> Self {
        Self(value)
    }

    /// Converts a wider unsigned value without truncating it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidKind`] when `value` cannot fit in two bytes.
    pub fn try_from_u64(value: u64) -> Result<Self> {
        u16::try_from(value)
            .map(Self::from_wire)
            .map_err(|_error| Error::InvalidKind { value })
    }

    /// Returns the exact on-wire value.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl TryFrom<u64> for Kind {
    type Error = Error;

    fn try_from(value: u64) -> Result<Self> {
        Self::try_from_u64(value)
    }
}

impl TryFrom<usize> for Kind {
    type Error = Error;

    fn try_from(value: usize) -> Result<Self> {
        Self::try_from_u64(as_u64(value))
    }
}

impl From<u16> for Kind {
    fn from(value: u16) -> Self {
        Self::from_wire(value)
    }
}

impl From<Kind> for u16 {
    fn from(value: Kind) -> Self {
        value.get()
    }
}

impl fmt::Debug for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "Kind(0x{:04X})", self.0)
    }
}

impl fmt::Display for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "0x{:04X}", self.0)
    }
}

/// A zero-copy view of one complete BIFF record frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecordRef<'a> {
    pub(crate) kind: Kind,
    pub(crate) payload: &'a [u8],
    pub(crate) encoded: &'a [u8],
    pub(crate) offset: usize,
}

impl<'a> RecordRef<'a> {
    /// Returns the record identifier.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Borrows the payload, excluding the four-byte frame header.
    #[must_use]
    pub const fn payload(self) -> &'a [u8] {
        self.payload
    }

    /// Borrows the exact header and payload bytes from the source stream.
    #[must_use]
    pub const fn encoded(self) -> &'a [u8] {
        self.encoded
    }

    /// Returns the header offset in the source stream.
    #[must_use]
    pub const fn offset(self) -> usize {
        self.offset
    }

    /// Copies this frame into an owned lossless record.
    ///
    /// # Errors
    ///
    /// Returns a limit or allocation error before the copy is attempted when
    /// the configured output budget cannot accommodate this frame.
    pub fn own(self) -> Result<Record> {
        self.own_with(Limits::default())
    }

    /// Copies this frame using explicit output and payload limits.
    ///
    /// # Errors
    ///
    /// Returns a limit or allocation error before the copy is attempted when
    /// the configured budget cannot accommodate this frame.
    pub fn own_with(self, limits: Limits) -> Result<Record> {
        let validated_limits = limits.validate()?;
        ensure_payload_limit(self.payload.len(), validated_limits.max_record_bytes)?;
        ensure_output_limit(self.encoded.len(), validated_limits.max_output_bytes)?;

        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.encoded.len())
            .map_err(|_error| Error::Allocation {
                resource: Resource::RecordFrame,
            })?;
        bytes.extend_from_slice(self.encoded);
        Ok(Record { bytes })
    }
}

/// A move-owned, lossless BIFF record frame.
///
/// The constructor accepts exactly one complete frame. The original header,
/// payload bytes, and their order are retained without semantic rewriting.
#[derive(Debug, PartialEq, Eq)]
pub struct Record {
    bytes: Vec<u8>,
}

impl Record {
    /// Takes ownership of and validates one frame under default limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error for an empty, truncated, oversized, or multi-frame
    /// input, or when the input cannot be retained under the default budget.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership of and validates one frame under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error for an empty, truncated, oversized, or multi-frame
    /// input, or when the input cannot be retained under `limits`.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        ensure_input_limit(bytes.len(), validated_limits.max_input_bytes)?;
        ensure_output_limit(bytes.len(), validated_limits.max_output_bytes)?;

        {
            let mut records = Records::validated(&bytes, validated_limits);
            match records.next().ok_or(Error::EmptyRecord)? {
                Ok(_) => {},
                Err(error) => return Err(error),
            }
            match records.next() {
                None => {},
                Some(Ok(second)) => {
                    return Err(Error::MultipleRecords {
                        offset: second.offset(),
                    });
                },
                Some(Err(error)) => return Err(error),
            }
        }

        Ok(Self { bytes })
    }

    /// Returns the record identifier.
    #[must_use]
    pub fn kind(&self) -> Kind {
        frame_kind(&self.bytes)
    }

    /// Borrows the exact encoded frame.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrows this owned frame as a zero-copy record view at offset zero.
    #[must_use]
    pub fn as_ref(&self) -> RecordRef<'_> {
        let payload = self.bytes.get(HEADER_BYTES..).unwrap_or(&[]);
        RecordRef {
            kind: self.kind(),
            payload,
            encoded: &self.bytes,
            offset: 0,
        }
    }

    /// Returns the encoded frame length, including its header.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    /// Reports whether the encoded frame is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    /// Recovers the original frame allocation without copying.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

impl AsRef<[u8]> for Record {
    fn as_ref(&self) -> &[u8] {
        self.as_bytes()
    }
}

fn ensure_payload_limit(payload_len: usize, maximum: usize) -> Result<()> {
    if payload_len > maximum {
        return Err(Error::LimitExceeded {
            resource: Resource::RecordBytes,
            observed: as_u64(payload_len),
            maximum: as_u64(maximum),
        });
    }
    Ok(())
}

fn ensure_input_limit(stream_len: usize, maximum: usize) -> Result<()> {
    if stream_len > maximum {
        return Err(Error::LimitExceeded {
            resource: Resource::InputBytes,
            observed: as_u64(stream_len),
            maximum: as_u64(maximum),
        });
    }
    Ok(())
}

fn ensure_output_limit(output_len: usize, maximum: usize) -> Result<()> {
    if output_len > maximum {
        return Err(Error::LimitExceeded {
            resource: Resource::OutputBytes,
            observed: as_u64(output_len),
            maximum: as_u64(maximum),
        });
    }
    Ok(())
}

fn frame_kind(bytes: &[u8]) -> Kind {
    let Some((&low, rest)) = bytes.split_first() else {
        return Kind::from_wire(0);
    };
    let Some(&high) = rest.first() else {
        return Kind::from_wire(u16::from_le_bytes([low, 0]));
    };
    Kind::from_wire(u16::from_le_bytes([low, high]))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn frame(kind: u16, payload: &[u8]) -> Vec<u8> {
        let Ok(length) = u16::try_from(payload.len()) else {
            return Vec::new();
        };
        let mut bytes = Vec::with_capacity(HEADER_BYTES + payload.len());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    #[test]
    fn kind_keeps_wire_values_and_rejects_wider_values() {
        let kind = Kind::from_wire(u16::MAX);
        assert_eq!(kind.get(), u16::MAX);
        assert_eq!(Kind::try_from(u64::from(u16::MAX)), Ok(kind));
        assert_eq!(Kind::try_from(usize::from(u16::MAX)), Ok(kind));
        assert!(matches!(
            Kind::try_from(u64::from(u16::MAX) + 1),
            Err(Error::InvalidKind { value }) if value == u64::from(u16::MAX) + 1
        ));
    }

    #[test]
    fn borrowed_view_lends_exact_parts_and_can_be_owned() {
        let bytes = frame(0x7777, &[0xAA, 0xBB, 0xCC]);
        let mut records = Records::new(&bytes);
        let Some(Ok(view)) = records.next() else {
            panic!("valid frame did not parse");
        };
        assert_eq!(view.kind(), Kind::from_wire(0x7777));
        assert_eq!(view.payload(), &[0xAA, 0xBB, 0xCC]);
        assert_eq!(view.encoded(), bytes.as_slice());
        assert_eq!(view.offset(), 0);
        assert_eq!(view.encoded().len(), bytes.len());

        let Ok(owned) = view.own() else {
            panic!("valid frame did not become owned");
        };
        assert_eq!(owned.as_bytes(), bytes.as_slice());
        assert_eq!(owned.as_ref(), view);
        assert_eq!(owned.into_bytes(), bytes);
    }

    #[test]
    fn owned_constructor_rejects_empty_multiple_and_truncated_input() {
        assert!(matches!(Record::open(Vec::new()), Err(Error::EmptyRecord)));

        let mut multiple = frame(1, &[1]);
        multiple.extend_from_slice(&frame(2, &[2]));
        assert!(matches!(
            Record::open(multiple),
            Err(Error::MultipleRecords { offset: 5 })
        ));

        assert!(matches!(
            Record::open(vec![1, 2, 1, 0]),
            Err(Error::TruncatedPayload {
                offset: 0,
                kind,
                declared: 1,
                available: 0
            }) if kind == Kind::from_wire(0x0201)
        ));
    }

    #[test]
    fn ownership_limits_are_checked_before_copying() {
        let bytes = frame(1, &[1, 2]);
        assert!(matches!(
            Record::with_limits(
                bytes.clone(),
                Limits {
                    max_output_bytes: bytes.len() - 1,
                    ..Limits::default()
                }
            ),
            Err(Error::LimitExceeded {
                resource: Resource::OutputBytes,
                ..
            })
        ));

        let mut records = Records::new(&bytes);
        let Some(Ok(view)) = records.next() else {
            panic!("valid frame did not parse");
        };
        assert!(matches!(
            view.own_with(Limits {
                max_record_bytes: 1,
                ..Limits::default()
            }),
            Err(Error::LimitExceeded {
                resource: Resource::RecordBytes,
                ..
            })
        ));
    }
}
