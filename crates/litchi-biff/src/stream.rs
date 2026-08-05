use std::iter::FusedIterator;

use crate::frame::{Kind, RecordRef};
use crate::limits::{HEADER_BYTES, as_u64};
use crate::{Error, Limits, Resource, Result};

/// Allocation-free traversal of complete BIFF frames in a borrowed stream.
///
/// Each successful item borrows the original header and payload. The iterator
/// stops after its first malformed frame and is fused thereafter. An empty
/// input is a valid empty collection; a non-empty input must end on a complete
/// frame boundary.
#[derive(Debug)]
pub struct Records<'a> {
    bytes: &'a [u8],
    offset: usize,
    count: usize,
    limits: Limits,
    pending: Option<Error>,
    done: bool,
}

impl<'a> Records<'a> {
    /// Creates a borrowed parser with conservative default limits.
    #[must_use]
    pub fn new(bytes: &'a [u8]) -> Self {
        let limits = Limits::default();
        Self::from_parts(bytes, limits)
    }

    /// Creates a borrowed parser with explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for an invalid configuration or
    /// [`Error::LimitExceeded`] when the complete input exceeds the stream
    /// ceiling. No input allocation is performed.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        ensure_input_limit(bytes.len(), validated_limits.max_input_bytes)?;
        Ok(Self::from_parts(bytes, validated_limits))
    }

    pub(crate) fn validated(bytes: &'a [u8], limits: Limits) -> Self {
        Self::from_parts(bytes, limits)
    }

    fn from_parts(bytes: &'a [u8], limits: Limits) -> Self {
        let pending = (bytes.len() > limits.max_input_bytes).then_some(Error::LimitExceeded {
            resource: Resource::InputBytes,
            observed: as_u64(bytes.len()),
            maximum: as_u64(limits.max_input_bytes),
        });
        Self {
            bytes,
            offset: 0,
            count: 0,
            limits,
            pending,
            done: false,
        }
    }

    fn fail(&mut self, error: Error) -> Error {
        self.done = true;
        error
    }

    fn next_count(&mut self) -> Result<usize> {
        let next_count = self.count.checked_add(1).ok_or(Error::SizeOverflow {
            resource: Resource::RecordCount,
        })?;
        if next_count > self.limits.max_records {
            return Err(Error::LimitExceeded {
                resource: Resource::RecordCount,
                observed: as_u64(next_count),
                maximum: as_u64(self.limits.max_records),
            });
        }
        Ok(next_count)
    }

    fn read_one(&mut self) -> Result<RecordRef<'a>> {
        let available = self.bytes.len().saturating_sub(self.offset);
        if available < HEADER_BYTES {
            return Err(Error::TruncatedHeader {
                offset: self.offset,
                available,
            });
        }

        let header_end = self
            .offset
            .checked_add(HEADER_BYTES)
            .ok_or(Error::SizeOverflow {
                resource: Resource::RecordHeader,
            })?;
        let Some(header) = self.bytes.get(self.offset..header_end) else {
            return Err(Error::TruncatedHeader {
                offset: self.offset,
                available,
            });
        };
        let [kind_lo, kind_hi, length_lo, length_hi] = header else {
            return Err(Error::TruncatedHeader {
                offset: self.offset,
                available,
            });
        };
        let kind = Kind::from_wire(u16::from_le_bytes([*kind_lo, *kind_hi]));
        let payload_len = usize::from(u16::from_le_bytes([*length_lo, *length_hi]));
        if payload_len > self.limits.max_record_bytes {
            return Err(Error::LimitExceeded {
                resource: Resource::RecordBytes,
                observed: as_u64(payload_len),
                maximum: as_u64(self.limits.max_record_bytes),
            });
        }

        let end = header_end
            .checked_add(payload_len)
            .ok_or(Error::SizeOverflow {
                resource: Resource::RecordPayload,
            })?;
        let Some(encoded) = self.bytes.get(self.offset..end) else {
            return Err(Error::TruncatedPayload {
                offset: self.offset,
                kind,
                declared: payload_len,
                available: self.bytes.len().saturating_sub(header_end),
            });
        };
        let Some(payload) = self.bytes.get(header_end..end) else {
            return Err(Error::TruncatedPayload {
                offset: self.offset,
                kind,
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
        Ok(record)
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<RecordRef<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        if let Some(error) = self.pending.take() {
            self.done = true;
            return Some(Err(error));
        }
        if self.offset == self.bytes.len() {
            self.done = true;
            return None;
        }

        let next_count = match self.next_count() {
            Ok(value) => value,
            Err(error) => return Some(Err(self.fail(error))),
        };
        let record = match self.read_one() {
            Ok(value) => value,
            Err(error) => return Some(Err(self.fail(error))),
        };
        self.count = next_count;
        Some(Ok(record))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        if self.done || (self.pending.is_none() && self.offset == self.bytes.len()) {
            (0, Some(0))
        } else if self.pending.is_some() {
            (1, Some(1))
        } else {
            (1, None)
        }
    }
}

impl FusedIterator for Records<'_> {}

/// A bounded encoder for new or preserved BIFF frames.
#[derive(Debug)]
pub struct Encoder {
    bytes: Vec<u8>,
    count: usize,
    limits: Limits,
}

impl Encoder {
    /// Creates an empty encoder with conservative default limits.
    #[must_use]
    pub fn new() -> Self {
        Self {
            bytes: Vec::new(),
            count: 0,
            limits: Limits::default(),
        }
    }

    /// Creates an empty encoder with explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when the payload ceiling exceeds the
    /// physical BIFF bound.
    pub fn with_limits(limits: Limits) -> Result<Self> {
        Ok(Self {
            bytes: Vec::new(),
            count: 0,
            limits: limits.validate()?,
        })
    }

    /// Appends one checked kind and payload without partial output on failure.
    ///
    /// # Errors
    ///
    /// Returns a limit, arithmetic, or allocation error before changing the
    /// encoded length.
    pub fn push(&mut self, kind: Kind, payload: &[u8]) -> Result<()> {
        self.check_record(payload.len())?;
        let payload_len = u16::try_from(payload.len()).map_err(|_error| Error::LimitExceeded {
            resource: Resource::RecordBytes,
            observed: as_u64(payload.len()),
            maximum: as_u64(crate::MAX_RECORD_BYTES),
        })?;
        let growth = HEADER_BYTES
            .checked_add(payload.len())
            .ok_or(Error::SizeOverflow {
                resource: Resource::EncodedRecord,
            })?;
        self.reserve(growth)?;
        self.bytes.extend_from_slice(&kind.get().to_le_bytes());
        self.bytes.extend_from_slice(&payload_len.to_le_bytes());
        self.bytes.extend_from_slice(payload);
        self.count += 1;
        Ok(())
    }

    /// Appends an already validated borrowed frame byte-for-byte.
    ///
    /// # Errors
    ///
    /// Returns a limit, arithmetic, or allocation error before changing the
    /// encoded length.
    pub fn push_ref(&mut self, record: RecordRef<'_>) -> Result<()> {
        self.check_record(record.payload.len())?;
        self.reserve(record.encoded.len())?;
        self.bytes.extend_from_slice(record.encoded);
        self.count += 1;
        Ok(())
    }

    /// Borrows the bytes encoded so far.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the number of frames accepted so far.
    #[must_use]
    pub const fn count(&self) -> usize {
        self.count
    }

    /// Returns the number of encoded bytes accepted so far.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    /// Reports whether no frames have been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    /// Completes encoding without another allocation or copy.
    #[must_use]
    pub fn finish(self) -> Vec<u8> {
        self.bytes
    }

    fn check_record(&self, payload_len: usize) -> Result<()> {
        let next_count = self.count.checked_add(1).ok_or(Error::SizeOverflow {
            resource: Resource::RecordCount,
        })?;
        if next_count > self.limits.max_records {
            return Err(Error::LimitExceeded {
                resource: Resource::RecordCount,
                observed: as_u64(next_count),
                maximum: as_u64(self.limits.max_records),
            });
        }
        if payload_len > self.limits.max_record_bytes {
            return Err(Error::LimitExceeded {
                resource: Resource::RecordBytes,
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
                resource: Resource::EncodedStream,
            })?;
        if new_len > self.limits.max_output_bytes {
            return Err(Error::LimitExceeded {
                resource: Resource::OutputBytes,
                observed: as_u64(new_len),
                maximum: as_u64(self.limits.max_output_bytes),
            });
        }
        self.bytes
            .try_reserve(growth)
            .map_err(|_error| Error::Allocation {
                resource: Resource::EncodedStream,
            })
    }
}

impl Default for Encoder {
    fn default() -> Self {
        Self::new()
    }
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

    fn append_frame(output: &mut Vec<u8>, kind: u16, payload: &[u8]) {
        output.extend_from_slice(&frame(kind, payload));
    }

    #[test]
    fn parses_empty_stream_and_fuses_after_end() {
        let mut records = Records::new(&[]);
        assert_eq!(records.next(), None);
        assert_eq!(records.next(), None);
        assert_eq!(records.size_hint(), (0, Some(0)));
    }

    #[test]
    fn preserves_offsets_unknown_kinds_and_order() {
        let mut input = Vec::new();
        append_frame(&mut input, 0x7777, &[0xAA, 0xBB]);
        append_frame(&mut input, 0x000A, &[]);
        append_frame(&mut input, u16::MAX, &[0xCC]);

        let mut kinds = Vec::new();
        let mut offsets = Vec::new();
        let mut encoder = Encoder::new();
        for item in Records::new(&input) {
            let Ok(record) = item else {
                panic!("valid generated stream did not parse");
            };
            kinds.push(record.kind());
            offsets.push(record.offset());
            assert!(encoder.push_ref(record).is_ok());
        }
        assert_eq!(
            kinds,
            vec![
                Kind::from_wire(0x7777),
                Kind::from_wire(0x000A),
                Kind::from_wire(u16::MAX)
            ]
        );
        assert_eq!(offsets, vec![0, 6, 10]);
        assert_eq!(encoder.finish(), input);
    }

    #[test]
    fn rejects_each_nonempty_truncation_of_a_frame() {
        let complete = frame(0x1234, &[1, 2, 3, 4]);
        for cut in 1..complete.len() {
            let mut records = Records::new(&complete[..cut]);
            let Some(Err(error)) = records.next() else {
                panic!("truncated prefix unexpectedly parsed");
            };
            if cut < HEADER_BYTES {
                assert!(matches!(error, Error::TruncatedHeader { offset: 0, .. }));
            } else {
                assert!(matches!(error, Error::TruncatedPayload { offset: 0, .. }));
            }
            assert_eq!(records.next(), None);
        }
    }

    #[test]
    fn rejects_input_limits_before_traversal() {
        let bytes = frame(1, &[1]);
        assert!(matches!(
            Records::with_limits(
                &bytes,
                Limits {
                    max_input_bytes: bytes.len() - 1,
                    ..Limits::default()
                }
            ),
            Err(Error::LimitExceeded {
                resource: Resource::InputBytes,
                ..
            })
        ));

        let mut records = Records::with_limits(
            &bytes,
            Limits {
                max_records: 0,
                ..Limits::default()
            },
        )
        .unwrap_or_else(|_| Records::new(&[]));
        let Some(Err(Error::LimitExceeded {
            resource: Resource::RecordCount,
            ..
        })) = records.next()
        else {
            panic!("record count limit was not enforced");
        };
        assert_eq!(records.next(), None);
    }

    #[test]
    fn rejects_invalid_payload_limits_and_oversized_records() {
        let invalid = Limits {
            max_record_bytes: crate::MAX_RECORD_BYTES + 1,
            ..Limits::default()
        };
        assert!(matches!(
            Records::with_limits(&[], invalid),
            Err(Error::InvalidLimit {
                resource: Resource::RecordBytes,
                ..
            })
        ));

        let too_large = vec![0x01, 0x00, 0x21, 0x20];
        let mut records = Records::new(&too_large);
        assert!(matches!(
            records.next(),
            Some(Err(Error::LimitExceeded {
                resource: Resource::RecordBytes,
                observed: 8225,
                maximum: 8224
            }))
        ));
        assert_eq!(records.next(), None);
    }

    #[test]
    fn encoder_failures_are_atomic_and_finish_is_zero_copy() {
        let mut encoder = Encoder::with_limits(Limits {
            max_records: 1,
            max_output_bytes: 5,
            ..Limits::default()
        })
        .unwrap_or_else(|_| Encoder::new());
        assert!(encoder.push(Kind::from_wire(1), &[1]).is_ok());
        let before = encoder.as_bytes().to_vec();
        assert!(matches!(
            encoder.push(Kind::from_wire(2), &[]),
            Err(Error::LimitExceeded {
                resource: Resource::RecordCount,
                ..
            })
        ));
        assert_eq!(encoder.as_bytes(), before.as_slice());
        assert_eq!(encoder.count(), 1);
        assert_eq!(encoder.len(), 5);
        assert!(!encoder.is_empty());
        assert_eq!(encoder.finish(), before);

        let mut output_limited = Encoder::with_limits(Limits {
            max_output_bytes: HEADER_BYTES - 1,
            ..Limits::default()
        })
        .unwrap_or_else(|_| Encoder::new());
        assert!(matches!(
            output_limited.push(Kind::from_wire(3), &[]),
            Err(Error::LimitExceeded {
                resource: Resource::OutputBytes,
                ..
            })
        ));
        assert!(output_limited.is_empty());
    }

    #[test]
    fn encoder_accepts_the_spec_payload_boundary_and_rejects_the_next_byte() {
        let payload = vec![0xA5; crate::MAX_RECORD_BYTES];
        let mut encoder = Encoder::new();
        assert!(encoder.push(Kind::from_wire(1), &payload).is_ok());
        assert_eq!(encoder.len(), HEADER_BYTES + crate::MAX_RECORD_BYTES);

        let mut restricted = Encoder::with_limits(Limits {
            max_record_bytes: crate::MAX_RECORD_BYTES - 1,
            ..Limits::default()
        })
        .unwrap_or_else(|_| Encoder::new());
        assert!(matches!(
            restricted.push(Kind::from_wire(1), &payload),
            Err(Error::LimitExceeded {
                resource: Resource::RecordBytes,
                observed: 8224,
                maximum: 8223
            })
        ));
    }

    #[test]
    fn deterministic_property_round_trip_preserves_every_generated_frame() {
        let mut state = 0xC0FF_EE12_u64;
        let mut input = Vec::new();
        let mut expected = Vec::new();
        for _ in 0..512 {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            let kind = u16::try_from((state >> 16) & u64::from(u16::MAX)).unwrap_or(0);
            let payload_len = usize::try_from((state >> 48) & 0x3F).unwrap_or(0);
            state = state.rotate_left(17);
            let mut payload = Vec::with_capacity(payload_len);
            for index in 0..payload_len {
                state = state
                    .wrapping_mul(2_862_933_555_777_941_757)
                    .wrapping_add(3_037_000_493);
                let byte =
                    u8::try_from((state ^ u64::try_from(index).unwrap_or(0)) & 0xFF).unwrap_or(0);
                payload.push(byte);
            }
            let encoded = frame(kind, &payload);
            input.extend_from_slice(&encoded);
            expected.push((Kind::from_wire(kind), payload, encoded));
        }

        let mut index = 0usize;
        for item in Records::new(&input) {
            let Ok(record) = item else {
                panic!("deterministic generated stream did not parse");
            };
            let Some((kind, payload, encoded)) = expected.get(index) else {
                panic!("parser produced too many records");
            };
            assert_eq!(record.kind(), *kind);
            assert_eq!(record.payload(), payload.as_slice());
            assert_eq!(record.encoded(), encoded.as_slice());
            index += 1;
        }
        assert_eq!(index, expected.len());
    }
}
