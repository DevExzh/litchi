//! Schema-driven merge of a replacement message into a stored one: the
//! fields a schema knows are replaced (singular messages recursively), every
//! unknown field keeps its bytes and position (the migration host's
//! `replace_known_wire_fields_preserving_unknown_with_schema`, record 0375a).

use super::{Budget, Error, Field, Limits, allocate, push_varint, read_varint, visit_fields};

/// The wire shape of one known field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum KnownKind {
    /// A varint, which must be canonical.
    Varint,
    /// A fixed32.
    Fixed32,
    /// A fixed64.
    Fixed64,
    /// A length-delimited value merged as a unit.
    Bytes,
    /// A nested message merged recursively.
    Message(&'static [KnownField]),
}

/// One field a schema knows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct KnownField {
    /// The field number.
    pub number: u32,
    /// Its wire shape.
    pub kind: KnownKind,
    /// Whether it is repeated.
    pub repeated: bool,
}

const fn known(number: u32, kind: KnownKind) -> KnownField {
    KnownField {
        number,
        kind,
        repeated: false,
    }
}

/// The schemas of the merged native messages.
pub mod schema {
    use super::{KnownField, KnownKind, known};

    const COLOR: &[KnownField] = &[
        known(1, KnownKind::Varint),
        known(3, KnownKind::Fixed32),
        known(4, KnownKind::Fixed32),
        known(5, KnownKind::Fixed32),
        known(6, KnownKind::Fixed32),
        known(7, KnownKind::Fixed32),
        known(8, KnownKind::Fixed32),
        known(9, KnownKind::Fixed32),
        known(10, KnownKind::Fixed32),
        known(11, KnownKind::Fixed32),
        known(12, KnownKind::Varint),
    ];
    const REFERENCE: &[KnownField] = &[
        known(1, KnownKind::Varint),
        known(2, KnownKind::Varint),
        known(3, KnownKind::Varint),
    ];
    const REFERENCE_DICTIONARY_ENTRY: &[KnownField] = &[
        known(1, KnownKind::Message(REFERENCE)),
        known(2, KnownKind::Message(REFERENCE)),
    ];
    const REFERENCE_DICTIONARY: &[KnownField] = &[KnownField {
        number: 1,
        kind: KnownKind::Message(REFERENCE_DICTIONARY_ENTRY),
        repeated: true,
    }];
    const STROKE_PATTERN: &[KnownField] = &[
        known(1, KnownKind::Varint),
        known(2, KnownKind::Fixed32),
        known(3, KnownKind::Varint),
        KnownField {
            number: 4,
            kind: KnownKind::Fixed32,
            repeated: true,
        },
    ];
    const SMART_STROKE: &[KnownField] = &[
        known(2, KnownKind::Bytes),
        known(3, KnownKind::Varint),
        known(4, KnownKind::Message(REFERENCE_DICTIONARY)),
        known(5, KnownKind::Fixed64),
    ];
    const FRAME: &[KnownField] = &[known(2, KnownKind::Bytes), known(3, KnownKind::Fixed32)];
    const PATTERNED_STROKE: &[KnownField] = &[known(2, KnownKind::Bytes)];

    const DATA_REFERENCE: &[KnownField] = &[known(1, KnownKind::Varint)];
    const POINT: &[KnownField] = &[known(1, KnownKind::Fixed32), known(2, KnownKind::Fixed32)];
    const GRADIENT_STOP: &[KnownField] = &[
        known(1, KnownKind::Message(COLOR)),
        known(2, KnownKind::Fixed32),
        known(3, KnownKind::Fixed32),
    ];
    const ANGLE_GRADIENT: &[KnownField] = &[known(2, KnownKind::Fixed32)];
    const TRANSFORM_GRADIENT: &[KnownField] = &[
        known(1, KnownKind::Message(POINT)),
        known(2, KnownKind::Message(POINT)),
        known(3, KnownKind::Message(POINT)),
    ];
    const GRADIENT: &[KnownField] = &[
        known(1, KnownKind::Varint),
        KnownField {
            number: 2,
            kind: KnownKind::Message(GRADIENT_STOP),
            repeated: true,
        },
        known(3, KnownKind::Fixed32),
        known(4, KnownKind::Varint),
        known(5, KnownKind::Message(ANGLE_GRADIENT)),
        known(6, KnownKind::Message(TRANSFORM_GRADIENT)),
    ];
    const IMAGE_FILL: &[KnownField] = &[
        known(1, KnownKind::Message(REFERENCE)),
        known(2, KnownKind::Varint),
        known(3, KnownKind::Message(COLOR)),
        known(4, KnownKind::Message(POINT)),
        known(5, KnownKind::Message(REFERENCE)),
        known(6, KnownKind::Message(DATA_REFERENCE)),
        known(7, KnownKind::Message(DATA_REFERENCE)),
        known(8, KnownKind::Varint),
        known(9, KnownKind::Message(COLOR)),
    ];

    /// `TSD.FillArchive` (the host's chart background-fill schema).
    pub const FILL: &[KnownField] = &[
        known(1, KnownKind::Message(COLOR)),
        known(2, KnownKind::Message(GRADIENT)),
        known(3, KnownKind::Message(IMAGE_FILL)),
    ];

    const CONTACT_SHADOW: &[KnownField] =
        &[known(2, KnownKind::Fixed32), known(4, KnownKind::Fixed32)];
    const CURVED_SHADOW: &[KnownField] = &[known(1, KnownKind::Fixed32)];
    const EMPTY: &[KnownField] = &[];

    /// `TSD.ShadowArchive` (the host's legend-shadow strict schema).
    pub const SHADOW: &[KnownField] = &[
        known(1, KnownKind::Message(COLOR)),
        known(2, KnownKind::Fixed32),
        known(3, KnownKind::Fixed32),
        known(4, KnownKind::Varint),
        known(5, KnownKind::Fixed32),
        known(6, KnownKind::Varint),
        known(7, KnownKind::Varint),
        known(8, KnownKind::Message(EMPTY)),
        known(9, KnownKind::Message(CONTACT_SHADOW)),
        known(10, KnownKind::Message(CURVED_SHADOW)),
    ];

    /// `TSD.StrokeArchive`.
    pub const STROKE: &[KnownField] = &[
        known(1, KnownKind::Message(COLOR)),
        known(2, KnownKind::Fixed32),
        known(3, KnownKind::Varint),
        known(4, KnownKind::Varint),
        known(5, KnownKind::Fixed32),
        known(6, KnownKind::Message(STROKE_PATTERN)),
        known(7, KnownKind::Message(SMART_STROKE)),
        known(8, KnownKind::Message(FRAME)),
        known(9, KnownKind::Message(PATTERNED_STROKE)),
    ];
}

fn fields(data: &[u8], budget: &mut Budget) -> Result<Vec<Field>, Error> {
    let mut found = Vec::new();
    visit_fields(data, budget, |field| {
        found.push(field);
        Ok(())
    })?;
    Ok(found)
}

const fn expected_wire_type(kind: KnownKind) -> u8 {
    match kind {
        KnownKind::Varint => 0,
        KnownKind::Fixed64 => 1,
        KnownKind::Bytes | KnownKind::Message(_) => 2,
        KnownKind::Fixed32 => 5,
    }
}

fn canonical_varint(data: &[u8], field: &Field) -> Result<(), Error> {
    let mut offset = field.key_end;
    let value = read_varint(data, &mut offset)?;
    if offset - field.key_end != super::varint_len(value) {
        return Err(Error::Malformed {
            reason: "a known varint is not canonical",
        });
    }
    Ok(())
}

fn canonical_length(field: &Field) -> Result<(), Error> {
    if field.wire_type == 2
        && field.payload_start - field.key_end
            != super::varint_len((field.end - field.payload_start) as u64)
    {
        return Err(Error::Malformed {
            reason: "a known length prefix is not canonical",
        });
    }
    Ok(())
}

/// Check the known fields of `data`: singular fields at most once, the
/// schema's wire type, canonical framing and varints, nested messages
/// recursively.
///
/// # Errors
///
/// [`Error::DuplicateField`], [`Error::WireType`] or [`Error::Malformed`].
pub fn validate(data: &[u8], schema: &[KnownField], limits: Limits) -> Result<(), Error> {
    validate_with(data, schema, &mut Budget::new(limits))
}

pub(crate) fn validate_with(
    data: &[u8],
    schema: &[KnownField],
    budget: &mut Budget,
) -> Result<(), Error> {
    let list = fields(data, budget)?;
    for (index, field) in list.iter().enumerate() {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number) else {
            continue;
        };
        if !spec.repeated && list[..index].iter().any(|seen| seen.number == field.number) {
            return Err(Error::DuplicateField {
                field: field.number,
            });
        }
        let expected = expected_wire_type(spec.kind);
        if field.wire_type != expected {
            return Err(Error::WireType {
                field: field.number,
                expected,
                found: field.wire_type,
            });
        }
        canonical_length(field)?;
        if matches!(spec.kind, KnownKind::Varint) {
            canonical_varint(data, field)?;
        }
        if let KnownKind::Message(nested) = spec.kind {
            validate_with(field.payload(data), nested, budget)?;
        }
    }
    Ok(())
}

fn raw<'a>(data: &'a [u8], field: &Field) -> &'a [u8] {
    &data[field.start..field.end]
}

/// Merge `replacement` into `existing` (see the module documentation).
///
/// # Errors
///
/// The [`validate`] errors for either side; limit and allocation errors.
pub fn merge(
    existing: &[u8],
    replacement: &[u8],
    schema: &[KnownField],
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    let mut budget = Budget::new(limits);
    validate_with(existing, schema, &mut budget)?;
    validate_with(replacement, schema, &mut budget)?;
    if existing == replacement {
        return Ok(existing.to_vec());
    }
    let existing_fields = fields(existing, &mut budget)?;
    let replacement_fields = fields(replacement, &mut budget)?;
    let mut output = allocate(existing.len() + replacement.len(), limits)?;
    let is_known = |number: u32| schema.iter().any(|spec| spec.number == number);
    let first_known = existing_fields
        .iter()
        .position(|field| is_known(field.number));
    let overlap = existing_fields.iter().any(|field| {
        is_known(field.number)
            && replacement_fields
                .iter()
                .any(|candidate| candidate.number == field.number)
    });
    if !overlap {
        let mut inserted = false;
        for (index, field) in existing_fields.iter().enumerate() {
            if is_known(field.number) {
                if first_known == Some(index) && !inserted {
                    output.extend_from_slice(replacement);
                    inserted = true;
                }
                continue;
            }
            output.extend_from_slice(raw(existing, field));
        }
        if !inserted {
            output.extend_from_slice(replacement);
        }
        return Ok(output);
    }
    let mut emitted: Vec<u32> = Vec::new();
    for field in &existing_fields {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number) else {
            output.extend_from_slice(raw(existing, field));
            continue;
        };
        if emitted.contains(&field.number) {
            continue;
        }
        emitted.push(field.number);
        let Some(replacement_field) = replacement_fields
            .iter()
            .find(|candidate| candidate.number == field.number)
        else {
            continue;
        };
        if spec.repeated {
            for candidate in replacement_fields
                .iter()
                .filter(|candidate| candidate.number == field.number)
            {
                output.extend_from_slice(raw(replacement, candidate));
            }
            continue;
        }
        if let KnownKind::Message(nested) = spec.kind {
            let merged = merge(
                field.payload(existing),
                replacement_field.payload(replacement),
                nested,
                limits,
            )?;
            output.extend_from_slice(&existing[field.start..field.key_end]);
            push_varint(&mut output, merged.len() as u64);
            output.extend_from_slice(&merged);
        } else {
            output.extend_from_slice(raw(replacement, replacement_field));
        }
    }
    for field in &replacement_fields {
        let matched = existing_fields
            .iter()
            .any(|candidate| candidate.number == field.number);
        if !matched || !is_known(field.number) {
            output.extend_from_slice(raw(replacement, field));
        }
    }
    Ok(output)
}
