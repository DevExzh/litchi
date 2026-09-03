//! Strict borrowed ingress for the Keynote movie `TSP.DataReference` edge.
//!
//! A movie keeps its video and poster assets in small nested data-reference
//! messages.  This codec validates that selected message without decoding the
//! surrounding `TSD.MovieArchive`, then cross-checks the scalar through the
//! private Buffa lazy view.  The caller-owned source slice remains the only
//! preservation representation; this module never re-encodes it.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict pass intentionally precedes the private Buffa cross-check."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_data_reference_generated::LitchiIwaDataReferenceProjection::DataReferenceLazyView;

const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const IDENTIFIER_FIELD: u32 = 1;

/// Finite limits for one nested Keynote movie data-reference payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Construct an explicit bytes/fields/work/nesting policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }

    /// Build a finite policy sized for one already-borrowed source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(2).max(1),
            8,
        )
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact finite resource consumption of one data-reference decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

impl DecodeReport {
    /// Number of fields visited by the strict source pass.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa bytes charged for the decode.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest message depth observed.  The root is depth one.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
}

/// Strict data-reference decode failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    /// Return the exact resource observation for a limit failure.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    const fn invalid() -> Self {
        Self { limit: None }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid Keynote movie data reference")
    }
}

impl std::error::Error for DecodeError {}

/// Resource classification for a bounded data-reference decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or Buffa message bytes exceed the configured ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Strict field visits exceed the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus Buffa work exceeds the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured recursion is outside the finite supported range.
    Nesting { observed: u32, maximum: u32 },
}

/// Borrowed semantic facts from one valid `TSP.DataReference` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataReferenceSnapshot<'source> {
    identifier: u64,
    raw: &'source [u8],
}

impl<'source> DataReferenceSnapshot<'source> {
    /// Return the required non-zero native media identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Return the exact caller-owned payload that was validated.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Decode one nested `TSP.DataReference` with strict canonical wire checks.
pub fn decode_data_reference<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<DataReferenceSnapshot<'source>, DecodeError> {
    decode_data_reference_with_report(source, options).map(|(snapshot, _report)| snapshot)
}

/// Decode one data reference and return exact finite resource consumption.
pub fn decode_data_reference_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(DataReferenceSnapshot<'source>, DecodeReport), DecodeError> {
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.max_fields == 0 {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: 1,
            maximum: options.max_fields,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }

    let strict = strict_data_reference(source, options)?;
    let work_bytes = source
        .len()
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let view: DataReferenceLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if !view.has_identifier() || view.identifier != strict.identifier {
        return Err(DecodeError::invalid());
    }
    Ok((
        DataReferenceSnapshot {
            identifier: strict.identifier,
            raw: source,
        },
        DecodeReport {
            fields: strict.fields,
            work_bytes,
            max_depth: strict.max_depth,
        },
    ))
}

#[derive(Debug, Clone, Copy)]
struct StrictDataReference {
    identifier: u64,
    fields: usize,
    max_depth: u32,
}

#[derive(Debug, Clone, Copy)]
struct StrictState {
    identifier: Option<u64>,
    fields: usize,
    max_depth: u32,
}

impl StrictState {
    const fn new() -> Self {
        Self {
            identifier: None,
            fields: 0,
            max_depth: 1,
        }
    }

    fn field(&mut self, options: DecodeOptions) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed,
                maximum: options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }
}

fn strict_data_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<StrictDataReference, DecodeError> {
    let mut input = source;
    let mut state = StrictState::new();
    strict_message(&mut input, options, 1, None, &mut state)?;
    Ok(StrictDataReference {
        identifier: state.identifier.ok_or_else(DecodeError::invalid)?,
        fields: state.fields,
        max_depth: state.max_depth,
    })
}

fn strict_message(
    input: &mut &[u8],
    options: DecodeOptions,
    depth: u32,
    expected_end_group: Option<u32>,
    state: &mut StrictState,
) -> Result<(), DecodeError> {
    state.max_depth = state.max_depth.max(depth);
    loop {
        if input.is_empty() {
            return expected_end_group
                .is_none()
                .then_some(())
                .ok_or_else(DecodeError::invalid);
        }
        let field = parse_field(input)?;

        // Count end-group tags too. This keeps the strict report aligned with
        // the complete wire walk rather than charging only payload fields.
        state.field(options)?;

        if field.wire == 4 {
            if expected_end_group == Some(field.number) {
                return Ok(());
            }
            // An end-group at the root, or an end-group for a different
            // opener, is malformed even when all surrounding bytes are valid.
            return Err(DecodeError::invalid());
        }

        if expected_end_group.is_none() && field.number == IDENTIFIER_FIELD {
            if state.identifier.is_some() || field.wire != 0 {
                return Err(DecodeError::invalid());
            }
            let value = field.varint.ok_or_else(DecodeError::invalid)?;
            if value == 0 {
                return Err(DecodeError::invalid());
            }
            state.identifier = Some(value);
        }

        if field.wire == 3 {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            if child_depth > options.recursion_limit {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: child_depth,
                    maximum: options.recursion_limit,
                }));
            }
            strict_message(input, options, child_depth, Some(field.number), state)?;
        }
    }
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    varint: Option<u64>,
    _bytes: Option<&'source [u8]>,
}

fn parse_field<'source>(input: &mut &'source [u8]) -> Result<Field<'source>, DecodeError> {
    let tag = take_varint(input)?;
    let number = u32::try_from(tag >> 3).map_err(|_error| DecodeError::invalid())?;
    let wire = u8::try_from(tag & 7).map_err(|_error| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let mut field = Field {
        number,
        wire,
        varint: None,
        _bytes: None,
    };
    match wire {
        0 => field.varint = Some(take_varint(input)?),
        1 => {
            let _ = take(input, 8)?;
        },
        2 => {
            let length =
                usize::try_from(take_varint(input)?).map_err(|_error| DecodeError::invalid())?;
            field._bytes = Some(take(input, length)?);
        },
        5 => {
            let _ = take(input, 4)?;
        },
        // Unknown groups are skipped by Buffa and retained by the caller-owned
        // raw source. The strict message walk validates their framing below.
        3 | 4 => {},
        _ => return Err(DecodeError::invalid()),
    }
    Ok(field)
}

fn take<'source>(input: &mut &'source [u8], count: usize) -> Result<&'source [u8], DecodeError> {
    if input.len() < count {
        return Err(DecodeError::invalid());
    }
    let (value, rest) = input.split_at(count);
    *input = rest;
    Ok(value)
}

fn take_varint(input: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *input;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(DecodeError::invalid)?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if encoded_varint_len(value) != consumed {
                return Err(DecodeError::invalid());
            }
            *input = &original[consumed..];
            return Ok(value);
        }
    }
    Err(DecodeError::invalid())
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use buffa::Message as _;

    const OPTIONS: DecodeOptions = DecodeOptions::new(64, 8, 128, 8);

    #[test]
    fn canonical_data_reference_cross_checks_lazy_view_and_borrows_raw() {
        let source = [0x08, 0x96, 0x01, 0x98, 0x06, 0x07];
        let (snapshot, report) = decode_data_reference_with_report(&source, OPTIONS).unwrap();
        assert_eq!(snapshot.identifier(), 150);
        assert_eq!(snapshot.raw(), source);
        assert_eq!(report.fields(), 2);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.max_depth(), 1);
    }

    #[test]
    fn malformed_or_noncanonical_selected_data_reference_is_rejected() {
        for source in [
            &[][..],
            &[0x08, 0x00][..],
            &[0x08, 0x81, 0x00][..],
            &[0x08, 0x01, 0x08, 0x02][..],
            &[0x10, 0x01][..],
            &[0x08, 0x01, 0x0b][..],
        ] {
            assert!(
                decode_data_reference(source, OPTIONS).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn unknown_scalar_bytes_survive_the_borrowed_projection() {
        let source = [0x08, 0x07, 0x1a, 0x01, 0xff, 0x25, 0x01, 0x02, 0x03, 0x04];
        let view = decode_data_reference(&source, OPTIONS).unwrap();
        assert_eq!(view.identifier(), 7);
        assert_eq!(view.raw(), source);
    }

    #[test]
    fn unknown_groups_round_trip_without_reencoding_or_projection() {
        // Field 10 is an unknown group. Its nested field 1 must remain
        // opaque: only the root DataReference.identifier is selected.
        let source = [0x53, 0x08, 0x2a, 0x54, 0x08, 0x07];
        let (snapshot, report) = decode_data_reference_with_report(&source, OPTIONS).unwrap();
        assert_eq!(snapshot.identifier(), 7);
        assert_eq!(snapshot.raw(), source);
        assert_eq!(report.fields(), 4);
        assert_eq!(report.max_depth(), 2);
    }

    #[test]
    fn malformed_unknown_groups_and_group_depth_are_rejected() {
        for source in [
            &[0x53, 0x5c, 0x08, 0x07][..],       // mismatched end-group number
            &[0x53, 0x08, 0x01, 0x08, 0x07][..], // unterminated group
            &[0x54, 0x08, 0x07][..],             // stray end-group
            &[0x0b, 0x0c, 0x08, 0x07][..],       // selected field with group wire
        ] {
            assert!(
                decode_data_reference(source, OPTIONS).is_err(),
                "{source:?}"
            );
        }

        let nested = [0x53, 0x5b, 0x5c, 0x54, 0x08, 0x07];
        assert_eq!(
            decode_data_reference(&nested, DecodeOptions::new(64, 8, 128, 2))
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 3,
                maximum: 2,
            })
        );
    }

    #[test]
    fn limits_are_checked_before_lazy_decode() {
        let source = [0x08, 0x07];
        assert_eq!(
            decode_data_reference(&source, DecodeOptions::new(1, 8, 128, 8))
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: 2,
                maximum: 1
            })
        );
        assert_eq!(
            decode_data_reference(&source, DecodeOptions::new(64, 8, 3, 8))
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Work {
                observed: 4,
                maximum: 3
            })
        );
    }

    #[test]
    fn generated_data_reference_fixture_matches_lazy_message_shape() {
        let source = crate::buffa_data_reference_generated::LitchiIwaDataReferenceProjection::DataReference {
            identifier: 9,
        }
        .try_encode_to_vec()
        .unwrap();
        assert_eq!(
            decode_data_reference(&source, OPTIONS)
                .unwrap()
                .identifier(),
            9
        );
    }
}
