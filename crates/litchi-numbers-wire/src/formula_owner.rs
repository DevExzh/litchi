//! Generated-free Numbers formula-owner identities and the empty table owner.
//!
//! A table's formula-owner UUID derives from the table UUID with the byte
//! order native `CFUUIDArchive` values use. This module keeps that arithmetic
//! and the canonical empty `TSCE.FormulaOwnerDependenciesArchive` envelope
//! together, so table creation and duplication need no generated schema
//! value. [`encode_empty_table`](crate::formula_owner::encode_empty_table) emits, byte for byte, the owner the legacy
//! host built and encoded with Prost (record 0368 parity test).

use core::fmt;

use litchi_iwa_common::varint::{encode_varint_into, encoded_len};

use crate::{LimitKind, Limits};

/// The native owner kind of a table's primary calculation-engine owner.
pub const TABLE_OWNER_KIND: u32 = 1;

/// A UUID in the two 64-bit words of a native `TSP.UUID`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Uuid {
    lower: u64,
    upper: u64,
}

impl Uuid {
    /// A UUID from its `TSP.UUID` words.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    /// The lower word.
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    /// The upper word.
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }

    /// Whether both words are zero.
    #[must_use]
    pub const fn is_nil(self) -> bool {
        self.lower == 0 && self.upper == 0
    }

    /// The native formula-owner UUID of the table with this UUID.
    ///
    /// Native Numbers stores the table owner's identity with the bytes of
    /// each UUID half reversed and the halves swapped.
    #[must_use]
    pub const fn for_table(self) -> Self {
        Self {
            lower: self.upper.swap_bytes(),
            upper: self.lower.swap_bytes(),
        }
    }

    /// The four words of a native `CFUUIDArchive`.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the CFUUID words are the four 32-bit quarters of the UUID"
    )]
    pub const fn cfuuid_words(self) -> [u32; 4] {
        [
            self.lower as u32,
            (self.lower >> 32) as u32,
            self.upper as u32,
            (self.upper >> 32) as u32,
        ]
    }

    /// The identity of native `CFUUIDArchive` words.
    #[must_use]
    pub const fn from_cfuuid_words(words: [u32; 4]) -> Self {
        Self {
            lower: words[0] as u64 | ((words[1] as u64) << 32),
            upper: words[2] as u64 | ((words[3] as u64) << 32),
        }
    }

    /// Translate a derived owner UUID from the `source` table owner to the
    /// `replacement` one, keeping its 128-bit wrapping offset (the native
    /// duplication rule for a table's derived owners).
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the 128-bit sum is split back into its two 64-bit words"
    )]
    pub const fn translate_from(self, source: Self, replacement: Self) -> Self {
        let owner = ((self.upper as u128) << 64) | self.lower as u128;
        let source = ((source.upper as u128) << 64) | source.lower as u128;
        let replacement = ((replacement.upper as u128) << 64) | replacement.lower as u128;
        let translated = replacement.wrapping_add(owner.wrapping_sub(source));
        Self {
            lower: translated as u64,
            upper: (translated >> 64) as u64,
        }
    }
}

/// A failure encoding an empty formula owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A required identity (table UUID, table-info reference or internal
    /// owner identifier) is zero.
    InvalidIdentity,
    /// The exact encoded owner exceeds the caller's budget.
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The exact encoded size.
        observed: usize,
        /// The caller-selected maximum.
        maximum: usize,
    },
    /// The exact output allocation could not be reserved.
    Allocation {
        /// The bytes requested.
        requested: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidIdentity => formatter.write_str("formula-owner identity is zero"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "formula-owner {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { requested } => {
                write!(
                    formatter,
                    "could not allocate {requested} formula-owner bytes"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// The empty sub-messages of a table owner, in field order, before the
/// table-info reference (field 11).
const LEADING_EMPTY_FIELDS: [u32; 2] = [4, 5];
/// The empty volatile-dependency sets (fields of field 6).
const VOLATILE_FIELDS: [u32; 6] = [1, 2, 3, 4, 5, 7];
/// The empty spanning dependencies and cell errors between the volatile and
/// whole-owner dependencies and the reference.
const SPANNING_FIELDS: [u32; 2] = [7, 8];
/// The empty sub-messages after the reference.
const TRAILING_EMPTY_FIELDS: [u32; 4] = [13, 14, 15, 16];

/// Encode the canonical empty `FormulaOwnerDependenciesArchive` of a table.
///
/// The owner carries the table's derived owner UUID, `internal_owner_id`,
/// owner kind [`TABLE_OWNER_KIND`], every empty dependency envelope (the six
/// volatile sets and the whole-owner dependent-cell set included) and a
/// reference to `table_info`, in field order. The exact size is computed and
/// checked against `limits` before the output is reserved.
///
/// # Errors
///
/// [`Error::InvalidIdentity`] for a nil table UUID or a zero reference or
/// internal identifier; [`Error::Limit`] when the owner exceeds the output
/// budget; [`Error::Allocation`] when the reservation fails.
pub fn encode_empty_table(
    table_uuid: Uuid,
    table_info: u64,
    internal_owner_id: u32,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    if table_uuid.is_nil() || table_info == 0 || internal_owner_id == 0 {
        return Err(Error::InvalidIdentity);
    }
    let owner_uuid = table_uuid.for_table();
    let uuid_len =
        varint_field_len(1, owner_uuid.lower()) + varint_field_len(2, owner_uuid.upper());
    let volatile_len = VOLATILE_FIELDS
        .iter()
        .map(|field| bytes_field_len(*field, 0))
        .sum();
    let whole_len = bytes_field_len(1, 0);
    let reference_len = varint_field_len(1, table_info);
    let output_len = bytes_field_len(1, uuid_len)
        + varint_field_len(2, u64::from(internal_owner_id))
        + varint_field_len(3, u64::from(TABLE_OWNER_KIND))
        + LEADING_EMPTY_FIELDS
            .iter()
            .map(|field| bytes_field_len(*field, 0))
            .sum::<usize>()
        + bytes_field_len(6, volatile_len)
        + SPANNING_FIELDS
            .iter()
            .map(|field| bytes_field_len(*field, 0))
            .sum::<usize>()
        + bytes_field_len(9, whole_len)
        + bytes_field_len(10, 0)
        + bytes_field_len(11, reference_len)
        + TRAILING_EMPTY_FIELDS
            .iter()
            .map(|field| bytes_field_len(*field, 0))
            .sum::<usize>();
    if output_len > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: output_len,
            maximum: limits.max_output_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_allocation| Error::Allocation {
            requested: output_len,
        })?;
    key(&mut output, 1, 2);
    length(&mut output, uuid_len);
    varint_field(&mut output, 1, owner_uuid.lower());
    varint_field(&mut output, 2, owner_uuid.upper());
    varint_field(&mut output, 2, u64::from(internal_owner_id));
    varint_field(&mut output, 3, u64::from(TABLE_OWNER_KIND));
    for field in LEADING_EMPTY_FIELDS {
        empty_field(&mut output, field);
    }
    key(&mut output, 6, 2);
    length(&mut output, volatile_len);
    for field in VOLATILE_FIELDS {
        empty_field(&mut output, field);
    }
    for field in SPANNING_FIELDS {
        empty_field(&mut output, field);
    }
    key(&mut output, 9, 2);
    length(&mut output, whole_len);
    empty_field(&mut output, 1);
    empty_field(&mut output, 10);
    key(&mut output, 11, 2);
    length(&mut output, reference_len);
    varint_field(&mut output, 1, table_info);
    for field in TRAILING_EMPTY_FIELDS {
        empty_field(&mut output, field);
    }
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn key(output: &mut Vec<u8>, field: u32, wire_type: u8) {
    encode_varint_into(output, u64::from(field) << 3 | u64::from(wire_type));
}

fn length(output: &mut Vec<u8>, value: usize) {
    encode_varint_into(output, value as u64);
}

fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    key(output, field, 0);
    encode_varint_into(output, value);
}

fn empty_field(output: &mut Vec<u8>, field: u32) {
    key(output, field, 2);
    output.push(0);
}

const fn key_len(field: u32, wire_type: u8) -> usize {
    encoded_len((field as u64) << 3 | wire_type as u64)
}

const fn varint_field_len(field: u32, value: u64) -> usize {
    key_len(field, 0) + encoded_len(value)
}

const fn bytes_field_len(field: u32, payload: usize) -> usize {
    key_len(field, 2) + encoded_len(payload as u64) + payload
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_owner_uuid_uses_native_byte_order() {
        let table = Uuid::new(0xa1f5_067b_8152_4a6a, 0x68b3_84c7_9e7f_4b90);
        assert_eq!(
            table.for_table(),
            Uuid::new(0x904b_7f9e_c784_b368, 0x6a4a_5281_7b06_f5a1)
        );
        assert_eq!(
            table.for_table().cfuuid_words(),
            [0xc784_b368, 0x904b_7f9e, 0x7b06_f5a1, 0x6a4a_5281]
        );
        assert_eq!(
            Uuid::from_cfuuid_words(table.for_table().cfuuid_words()),
            table.for_table()
        );
    }

    #[test]
    fn owner_family_translation_keeps_wrapping_offsets() {
        let source = Uuid::new(0x1000, 0x2000);
        let replacement = Uuid::new(0x9000, 0xa000);
        assert_eq!(
            Uuid::new(0x1005, 0x1fff).translate_from(source, replacement),
            Uuid::new(0x9005, 0x9fff)
        );
        assert_eq!(
            Uuid::new(0, 0).translate_from(Uuid::new(1, 0), Uuid::new(0, 0)),
            Uuid::new(u64::MAX, u64::MAX)
        );
    }

    #[test]
    fn empty_owner_has_every_native_envelope_in_field_order() {
        let encoded = encode_empty_table(
            Uuid::new(0x0011_2233_4455_6677, 0x8899_aabb_ccdd_eeff),
            0x1234,
            7,
            &Limits::DEFAULT,
        )
        .unwrap();
        // volatile sets 1..5 and 7, whole-owner dependent cells, reference
        let volatile = [
            0x32, 12, 0x0a, 0, 0x12, 0, 0x1a, 0, 0x22, 0, 0x2a, 0, 0x3a, 0,
        ];
        assert!(
            encoded
                .windows(volatile.len())
                .any(|window| window == volatile)
        );
        assert!(
            encoded
                .windows(4)
                .any(|window| window == [0x4a, 2, 0x0a, 0])
        );
        assert!(
            encoded
                .windows(5)
                .any(|window| window == [0x5a, 3, 0x08, 0xb4, 0x24])
        );
        assert!(encoded.ends_with(&[0x6a, 0, 0x72, 0, 0x7a, 0, 0x82, 0x01, 0]));
        for (table, info, internal) in [
            (Uuid::new(0, 0), 1, 1),
            (Uuid::new(1, 2), 0, 1),
            (Uuid::new(1, 2), 3, 0),
        ] {
            assert_eq!(
                encode_empty_table(table, info, internal, &Limits::DEFAULT),
                Err(Error::InvalidIdentity)
            );
        }
    }

    #[test]
    fn output_limit_is_checked_before_reservation() {
        let limits = Limits::DEFAULT.with_output_bytes(1).unwrap();
        assert!(matches!(
            encode_empty_table(Uuid::new(1, 2), 3, 4, &limits),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                maximum: 1,
                ..
            })
        ));
    }
}
