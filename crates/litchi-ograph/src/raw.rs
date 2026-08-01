//! Checked BIFF record framing with borrowed payloads.

use std::fmt;
use std::iter::FusedIterator;

use crate::limits::as_u64;
use crate::{Error, Limits, Result};

const HEADER_BYTES: usize = 4;

/// A strongly typed BIFF record identifier.
#[derive(Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Kind(u16);

impl Kind {
    /// Creates an identifier from its on-disk value.
    pub const fn new(value: u16) -> Self {
        Self(value)
    }

    /// Returns the on-disk value.
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl From<u16> for Kind {
    fn from(value: u16) -> Self {
        Self::new(value)
    }
}

impl From<Kind> for u16 {
    fn from(value: Kind) -> Self {
        value.get()
    }
}

impl fmt::Debug for Kind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Kind({:#06X})", self.0)
    }
}

impl fmt::Display for Kind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:#06X}", self.0)
    }
}

/// A zero-copy view of one complete BIFF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecordRef<'a> {
    kind: Kind,
    payload: &'a [u8],
    encoded: &'a [u8],
    offset: usize,
}

impl<'a> RecordRef<'a> {
    /// Record identifier.
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Borrowed payload bytes, excluding the four-byte header.
    pub const fn payload(self) -> &'a [u8] {
        self.payload
    }

    /// Borrowed header and payload exactly as they appeared in the input.
    pub const fn encoded(self) -> &'a [u8] {
        self.encoded
    }

    /// Byte offset of the record header in the input.
    pub const fn offset(self) -> usize {
        self.offset
    }
}

/// Bounded, allocation-free iterator over BIFF record frames.
#[derive(Debug)]
pub struct Records<'a> {
    bytes: &'a [u8],
    offset: usize,
    count: usize,
    limits: Limits,
    done: bool,
}

impl<'a> Records<'a> {
    /// Parses with conservative default resource limits.
    pub fn new(bytes: &'a [u8]) -> Self {
        Self {
            bytes,
            offset: 0,
            count: 0,
            limits: Limits::default(),
            done: false,
        }
    }

    /// Parses using explicitly configured resource limits.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        Ok(Self {
            bytes,
            offset: 0,
            count: 0,
            limits: limits.validate()?,
            done: false,
        })
    }

    fn fail(&mut self, error: Error) -> Option<Result<RecordRef<'a>>> {
        self.done = true;
        Some(Err(error))
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<RecordRef<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done || self.offset == self.bytes.len() {
            self.done = true;
            return None;
        }

        let next_count = match self.count.checked_add(1) {
            Some(value) => value,
            None => {
                return self.fail(Error::SizeOverflow {
                    resource: "record count",
                });
            },
        };
        if next_count > self.limits.max_records {
            return self.fail(Error::LimitExceeded {
                resource: "record count",
                observed: as_u64(next_count),
                maximum: as_u64(self.limits.max_records),
            });
        }

        let available = self.bytes.len().saturating_sub(self.offset);
        if available < HEADER_BYTES {
            return self.fail(Error::TruncatedHeader {
                offset: self.offset,
                available,
            });
        }

        let header_end = match self.offset.checked_add(HEADER_BYTES) {
            Some(value) => value,
            None => {
                return self.fail(Error::SizeOverflow {
                    resource: "record header",
                });
            },
        };
        let Some(header) = self.bytes.get(self.offset..header_end) else {
            return self.fail(Error::TruncatedHeader {
                offset: self.offset,
                available,
            });
        };
        let kind = Kind::new(u16::from_le_bytes([header[0], header[1]]));
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        if payload_len > self.limits.max_record_bytes {
            return self.fail(Error::LimitExceeded {
                resource: "record bytes",
                observed: as_u64(payload_len),
                maximum: as_u64(self.limits.max_record_bytes),
            });
        }

        let end = match header_end.checked_add(payload_len) {
            Some(value) => value,
            None => {
                return self.fail(Error::SizeOverflow {
                    resource: "record payload",
                });
            },
        };
        let Some(encoded) = self.bytes.get(self.offset..end) else {
            return self.fail(Error::TruncatedPayload {
                offset: self.offset,
                kind: kind.get(),
                declared: payload_len,
                available: self.bytes.len().saturating_sub(header_end),
            });
        };
        let Some(payload) = self.bytes.get(header_end..end) else {
            return self.fail(Error::TruncatedPayload {
                offset: self.offset,
                kind: kind.get(),
                declared: payload_len,
                available: self.bytes.len().saturating_sub(header_end),
            });
        };

        let record = RecordRef {
            kind,
            payload,
            encoded,
            offset: self.offset,
        };
        self.offset = end;
        self.count = next_count;
        Some(Ok(record))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        if self.done || self.offset == self.bytes.len() {
            (0, Some(0))
        } else {
            (1, None)
        }
    }
}

impl FusedIterator for Records<'_> {}

/// Bounded BIFF stream encoder.
#[derive(Debug)]
pub struct Encoder {
    bytes: Vec<u8>,
    count: usize,
    limits: Limits,
}

impl Encoder {
    /// Creates an empty encoder with conservative default resource limits.
    pub fn new() -> Self {
        Self {
            bytes: Vec::new(),
            count: 0,
            limits: Limits::default(),
        }
    }

    /// Creates an empty encoder with explicitly configured limits.
    pub fn with_limits(limits: Limits) -> Result<Self> {
        Ok(Self {
            bytes: Vec::new(),
            count: 0,
            limits: limits.validate()?,
        })
    }

    /// Appends one checked record.
    pub fn push(&mut self, kind: Kind, payload: &[u8]) -> Result<()> {
        self.check_record(payload.len())?;
        let len = u16::try_from(payload.len()).map_err(|_| Error::LimitExceeded {
            resource: "record bytes",
            observed: as_u64(payload.len()),
            maximum: u64::from(u16::MAX),
        })?;
        let growth = HEADER_BYTES
            .checked_add(payload.len())
            .ok_or(Error::SizeOverflow {
                resource: "encoded record",
            })?;
        self.reserve(growth)?;
        self.bytes.extend_from_slice(&kind.get().to_le_bytes());
        self.bytes.extend_from_slice(&len.to_le_bytes());
        self.bytes.extend_from_slice(payload);
        self.count += 1;
        Ok(())
    }

    /// Copies an already validated frame exactly, preserving its raw header,
    /// unknown payload, and relative order.
    pub fn push_ref(&mut self, record: RecordRef<'_>) -> Result<()> {
        self.check_record(record.payload.len())?;
        self.reserve(record.encoded.len())?;
        self.bytes.extend_from_slice(record.encoded);
        self.count += 1;
        Ok(())
    }

    /// Borrows the bytes encoded so far.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Completes encoding without another allocation or copy.
    pub fn finish(self) -> Vec<u8> {
        self.bytes
    }

    fn check_record(&self, payload_len: usize) -> Result<()> {
        let next_count = self.count.checked_add(1).ok_or(Error::SizeOverflow {
            resource: "record count",
        })?;
        if next_count > self.limits.max_records {
            return Err(Error::LimitExceeded {
                resource: "record count",
                observed: as_u64(next_count),
                maximum: as_u64(self.limits.max_records),
            });
        }
        if payload_len > self.limits.max_record_bytes {
            return Err(Error::LimitExceeded {
                resource: "record bytes",
                observed: as_u64(payload_len),
                maximum: as_u64(self.limits.max_record_bytes),
            });
        }
        Ok(())
    }

    fn reserve(&mut self, growth: usize) -> Result<()> {
        let new_len = self
            .bytes
            .len()
            .checked_add(growth)
            .ok_or(Error::SizeOverflow {
                resource: "encoded stream",
            })?;
        if new_len > self.limits.max_output_bytes {
            return Err(Error::LimitExceeded {
                resource: "output bytes",
                observed: as_u64(new_len),
                maximum: as_u64(self.limits.max_output_bytes),
            });
        }
        self.bytes
            .try_reserve(growth)
            .map_err(|_| Error::Allocation {
                resource: "encoded stream",
            })
    }
}

impl Default for Encoder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    #[test]
    fn preserves_unknown_records_and_order_exactly() {
        let mut input = record(0x7777, &[0xAA, 0xBB]);
        input.extend_from_slice(&record(0x000A, &[]));

        let mut encoder = Encoder::new();
        for item in Records::new(&input) {
            encoder
                .push_ref(item.expect("valid fixture"))
                .expect("encode");
        }
        assert_eq!(encoder.finish(), input);
    }

    #[test]
    fn reports_truncated_header_and_payload_without_panicking() {
        let header = Records::new(&[1, 2, 3]).next().expect("one error");
        assert!(matches!(header, Err(Error::TruncatedHeader { .. })));

        let payload = Records::new(&[0x09, 0x08, 0x04, 0x00, 1, 2])
            .next()
            .expect("one error");
        assert!(matches!(payload, Err(Error::TruncatedPayload { .. })));
    }

    #[test]
    fn enforces_record_count_payload_and_output_limits() {
        let count_limits = Limits {
            max_records: 1,
            ..Limits::default()
        };
        let mut input = record(1, &[]);
        input.extend_from_slice(&record(2, &[]));
        let mut records = Records::with_limits(&input, count_limits).expect("limits");
        assert!(records.next().expect("first").is_ok());
        assert!(matches!(
            records.next().expect("limit"),
            Err(Error::LimitExceeded {
                resource: "record count",
                ..
            })
        ));
        assert!(records.next().is_none());

        let payload_limits = Limits {
            max_record_bytes: 1,
            ..Limits::default()
        };
        let input = record(1, &[1, 2]);
        assert!(matches!(
            Records::with_limits(&input, payload_limits)
                .expect("limits")
                .next()
                .expect("limit"),
            Err(Error::LimitExceeded {
                resource: "record bytes",
                ..
            })
        ));

        let output_limits = Limits {
            max_output_bytes: 5,
            ..Limits::default()
        };
        let mut encoder = Encoder::with_limits(output_limits).expect("limits");
        let error = encoder.push(Kind::new(1), &[1, 2]).expect_err("too large");
        assert!(matches!(
            error,
            Error::LimitExceeded {
                resource: "output bytes",
                ..
            }
        ));
        assert!(encoder.as_bytes().is_empty());
    }

    #[test]
    fn rejects_invalid_limit_configuration() {
        let limits = Limits {
            max_record_bytes: crate::limits::MAX_BIFF_RECORD_BYTES + 1,
            ..Limits::default()
        };
        assert!(matches!(
            Records::with_limits(&[], limits),
            Err(Error::InvalidLimit {
                resource: "record bytes",
                ..
            })
        ));
    }
}
