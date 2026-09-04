//! Bounded generated-free admission for Numbers table-model candidates.

use super::*;
use litchi_iwa_common::LimitKind;
use litchi_iwa_protos::numbers_table_cell_storage_codec as table_model_codec;

const MAX_FIELDS: usize = litchi_iwa_common::WireLimits::MAX_FIELDS;
const MAX_INPUT_BYTES: usize = litchi_iwa_common::WireLimits::MAX_INPUT_BYTES;
const MAX_WORK: usize = litchi_iwa_common::WireLimits::MAX_REWRITE_WORK;
const MAX_REFERENCES: usize = litchi_numbers::MAX_REFERENCES;
const MAX_TEXT_BYTES: usize = litchi_numbers::DEFAULT_MAX_TEXT_BYTES;
const RECURSION_LIMIT: u32 = 64;

/// Result of probing one type-gated table-model candidate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CandidateProbe {
    Valid,
    NotModel,
    Malformed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ProjectionKind {
    Strict,
    SparseCompatibility,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CandidateClassification {
    Projection(ProjectionKind),
    NotModel,
    Malformed,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct ShapeScanReport {
    scanned_bytes: usize,
    fields: usize,
    max_depth: usize,
}

#[derive(Debug, Clone, Copy)]
struct ScannedField<'source> {
    number: u32,
    wire_type: u8,
    payload: &'source [u8],
    canonical_key: bool,
    canonical_length: bool,
    end: usize,
}

struct ShapeScanner {
    base: ProbeBudget,
    report: ShapeScanReport,
}

impl ShapeScanner {
    const fn new(base: ProbeBudget) -> Self {
        Self {
            base,
            report: ShapeScanReport {
                scanned_bytes: 0,
                fields: 0,
                max_depth: 0,
            },
        }
    }

    fn charge_message(&mut self, bytes: usize, depth: usize) -> litchi_iwa_common::Result<()> {
        let scanned_bytes = self.report.scanned_bytes.checked_add(bytes).ok_or(
            litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: usize::MAX,
                limit: MAX_INPUT_BYTES,
            },
        )?;
        let observed_input = self.base.input_bytes.saturating_add(scanned_bytes);
        if observed_input > MAX_INPUT_BYTES {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: observed_input,
                limit: MAX_INPUT_BYTES,
            });
        }
        let observed_work = self.base.work.saturating_add(scanned_bytes);
        if observed_work > MAX_WORK {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed: observed_work,
                limit: MAX_WORK,
            });
        }
        if depth > RECURSION_LIMIT as usize {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: depth,
                limit: RECURSION_LIMIT as usize,
            });
        }
        self.report.scanned_bytes = scanned_bytes;
        self.report.max_depth = self.report.max_depth.max(depth);
        Ok(())
    }

    fn charge_field(&mut self) -> litchi_iwa_common::Result<()> {
        let fields =
            self.report
                .fields
                .checked_add(1)
                .ok_or(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Fields,
                    observed: usize::MAX,
                    limit: MAX_FIELDS,
                })?;
        let observed = self.base.fields.saturating_add(fields);
        if observed > MAX_FIELDS {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: MAX_FIELDS,
            });
        }
        self.report.fields = fields;
        Ok(())
    }
}

/// Aggregate finite ledger for every candidate in one discovery operation.
///
/// Historical Numbers fixtures rely on generated proto2 defaults, so this
/// narrow probe deliberately uses the codec's compatibility envelope. It is
/// not the complete-state authority: after exact candidate selection, callers
/// that need the full model still decode that one source into the existing
/// owned Prost value before mutation or publication.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct ProbeBudget {
    input_bytes: usize,
    fields: usize,
    work: usize,
    references: usize,
    text_bytes: usize,
}

impl ProbeBudget {
    pub(crate) const fn new() -> Self {
        Self {
            input_bytes: 0,
            fields: 0,
            work: 0,
            references: 0,
            text_bytes: 0,
        }
    }

    pub(crate) fn options(self, source: &[u8]) -> table_model_codec::DecodeOptions {
        table_model_codec::DecodeOptions::new(
            source
                .len()
                .max(1)
                .min(MAX_INPUT_BYTES.saturating_sub(self.input_bytes)),
            MAX_FIELDS.saturating_sub(self.fields),
            MAX_WORK.saturating_sub(self.work),
            RECURSION_LIMIT,
            MAX_REFERENCES.saturating_sub(self.references),
            MAX_TEXT_BYTES.saturating_sub(self.text_bytes),
        )
    }

    fn charge_typed(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: litchi_iwa_common::LimitKind,
    ) -> Result<()> {
        let observed = current.checked_add(amount).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-model candidate {kind} counter overflows host usize"
            ))
        })?;
        if observed > maximum {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit: maximum,
            }));
        }
        *current = observed;
        Ok(())
    }

    fn charge_untyped(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        name: &str,
    ) -> Result<()> {
        let observed = current.checked_add(amount).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-model candidate {name} counter overflows host usize"
            ))
        })?;
        if observed > maximum {
            return Err(Error::InvalidFormat(format!(
                "Numbers table-model candidate {name} limit exceeded: observed {observed}, limit {maximum}"
            )));
        }
        *current = observed;
        Ok(())
    }

    pub(crate) fn charge_report(&mut self, report: table_model_codec::DecodeReport) -> Result<()> {
        let mut next = *self;
        Self::charge_typed(
            &mut next.input_bytes,
            report.source_bytes(),
            MAX_INPUT_BYTES,
            litchi_iwa_common::LimitKind::InputBytes,
        )?;
        Self::charge_typed(
            &mut next.fields,
            report.fields(),
            MAX_FIELDS,
            litchi_iwa_common::LimitKind::Fields,
        )?;
        Self::charge_typed(
            &mut next.work,
            report.work_bytes(),
            MAX_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )?;
        Self::charge_untyped(
            &mut next.references,
            report.references(),
            MAX_REFERENCES,
            "reference",
        )?;
        Self::charge_untyped(
            &mut next.text_bytes,
            report.text_bytes(),
            MAX_TEXT_BYTES,
            "text-byte",
        )?;
        *self = next;
        Ok(())
    }

    fn charge_shape_scan(&mut self, report: ShapeScanReport) -> Result<()> {
        let mut next = *self;
        Self::charge_typed(
            &mut next.input_bytes,
            report.scanned_bytes,
            MAX_INPUT_BYTES,
            LimitKind::InputBytes,
        )?;
        Self::charge_typed(
            &mut next.fields,
            report.fields,
            MAX_FIELDS,
            LimitKind::Fields,
        )?;
        Self::charge_typed(
            &mut next.work,
            report.scanned_bytes,
            MAX_WORK,
            LimitKind::RewriteWork,
        )?;
        *self = next;
        Ok(())
    }

    fn charge_malformed(&mut self, source: &[u8]) -> Result<()> {
        let mut next = *self;
        Self::charge_typed(
            &mut next.input_bytes,
            source.len(),
            MAX_INPUT_BYTES,
            litchi_iwa_common::LimitKind::InputBytes,
        )?;
        let conservative_work = source.len().max(1).checked_mul(32).ok_or_else(|| {
            Error::InvalidFormat(
                "Numbers table-model malformed-candidate work overflows host usize".to_owned(),
            )
        })?;
        Self::charge_typed(
            &mut next.work,
            conservative_work,
            MAX_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )?;
        *self = next;
        Ok(())
    }
}

fn parse_shape_key(
    source: &[u8],
    offset: usize,
) -> litchi_iwa_common::Result<(u32, u8, usize, bool)> {
    let (key, width) = litchi_iwa_common::decode_varint_from_bytes(&source[offset..])
        .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?;
    let key_end = offset.checked_add(width).ok_or_else(|| {
        litchi_iwa_common::Error::InvalidFormat("protobuf key offset overflow".to_owned())
    })?;
    let number = u32::try_from(key >> 3).map_err(|error| {
        litchi_iwa_common::Error::InvalidFormat(format!(
            "protobuf field number does not fit u32: {error}"
        ))
    })?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "invalid protobuf field number {number}"
        )));
    }
    let wire_type = u8::try_from(key & 7).map_err(|error| {
        litchi_iwa_common::Error::InvalidFormat(format!(
            "protobuf wire type does not fit u8: {error}"
        ))
    })?;
    if wire_type > 5 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "invalid protobuf wire type {wire_type}"
        )));
    }
    Ok((
        number,
        wire_type,
        key_end,
        width == litchi_iwa_common::varint::encoded_len(key),
    ))
}

fn parse_shape_scalar(
    source: &[u8],
    key_end: usize,
    wire_type: u8,
) -> litchi_iwa_common::Result<(usize, usize, bool)> {
    let (payload_start, end, canonical_length) = match wire_type {
        0 => {
            let (_, width) = litchi_iwa_common::decode_varint_from_bytes(&source[key_end..])
                .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?;
            (key_end, key_end.checked_add(width), true)
        },
        1 => (key_end, key_end.checked_add(8), true),
        2 => {
            let (encoded_length, width) =
                litchi_iwa_common::decode_varint_from_bytes(&source[key_end..])
                    .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?;
            let payload_start = key_end.checked_add(width).ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat(
                    "protobuf length-prefix overflow".to_owned(),
                )
            })?;
            let length = usize::try_from(encoded_length).map_err(|error| {
                litchi_iwa_common::Error::InvalidFormat(format!(
                    "protobuf field length exceeds usize: {error}"
                ))
            })?;
            (
                payload_start,
                payload_start.checked_add(length),
                width == litchi_iwa_common::varint::encoded_len(encoded_length),
            )
        },
        5 => (key_end, key_end.checked_add(4), true),
        3 | 4 => {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "protobuf group was parsed as a scalar".to_owned(),
            ));
        },
        _ => unreachable!("wire type was bounded by parse_shape_key"),
    };
    let end = end.ok_or_else(|| {
        litchi_iwa_common::Error::InvalidFormat("protobuf field range overflow".to_owned())
    })?;
    if end > source.len() {
        return Err(litchi_iwa_common::Error::InvalidFormat(
            "truncated protobuf field".to_owned(),
        ));
    }
    Ok((payload_start, end, canonical_length))
}

fn skip_shape_group(
    source: &[u8],
    mut offset: usize,
    group_number: u32,
    depth: usize,
    scanner: &mut ShapeScanner,
) -> litchi_iwa_common::Result<usize> {
    if depth > RECURSION_LIMIT as usize {
        return Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed: depth,
            limit: RECURSION_LIMIT as usize,
        });
    }
    scanner.report.max_depth = scanner.report.max_depth.max(depth);
    while offset < source.len() {
        let (number, wire_type, key_end, _) = parse_shape_key(source, offset)?;
        scanner.charge_field()?;
        if wire_type == 4 {
            if number != group_number {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "protobuf end-group field {number} does not match start-group field {group_number}"
                )));
            }
            return Ok(key_end);
        }
        offset = if wire_type == 3 {
            skip_shape_group(source, key_end, number, depth + 1, scanner)?
        } else {
            parse_shape_scalar(source, key_end, wire_type)?.1
        };
    }
    Err(litchi_iwa_common::Error::InvalidFormat(format!(
        "protobuf start-group field {group_number} is not terminated"
    )))
}

fn parse_shape_field<'source>(
    source: &'source [u8],
    offset: usize,
    depth: usize,
    scanner: &mut ShapeScanner,
) -> litchi_iwa_common::Result<ScannedField<'source>> {
    let (number, wire_type, key_end, canonical_key) = parse_shape_key(source, offset)?;
    scanner.charge_field()?;
    if wire_type == 4 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "unexpected protobuf end-group field {number}"
        )));
    }
    if wire_type == 3 {
        let end = skip_shape_group(source, key_end, number, depth + 1, scanner)?;
        return Ok(ScannedField {
            number,
            wire_type,
            payload: &source[key_end..end],
            canonical_key,
            canonical_length: true,
            end,
        });
    }
    let (payload_start, end, canonical_length) = parse_shape_scalar(source, key_end, wire_type)?;
    Ok(ScannedField {
        number,
        wire_type,
        payload: &source[payload_start..end],
        canonical_key,
        canonical_length,
        end,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReferenceIdentifierShape {
    Default,
    Nonzero,
    Malformed,
}

fn classify_reference_identifier(
    source: &[u8],
    depth: usize,
    scanner: &mut ShapeScanner,
) -> litchi_iwa_common::Result<ReferenceIdentifierShape> {
    if source.is_empty() {
        return Ok(ReferenceIdentifierShape::Default);
    }
    scanner.charge_message(source.len(), depth)?;
    let mut offset = 0;
    let mut identifier = None;
    let mut has_extra_fields = false;
    while offset < source.len() {
        let field = parse_shape_field(source, offset, depth, scanner)?;
        offset = field.end;
        if field.number != 1 {
            has_extra_fields = true;
            continue;
        }
        if field.wire_type != 0 || !field.canonical_key {
            return Ok(ReferenceIdentifierShape::Malformed);
        }
        let Ok((value, width)) = litchi_iwa_common::decode_varint_from_bytes(field.payload) else {
            return Ok(ReferenceIdentifierShape::Malformed);
        };
        if width != field.payload.len()
            || width != litchi_iwa_common::varint::encoded_len(value)
            || identifier.replace(value).is_some()
        {
            return Ok(ReferenceIdentifierShape::Malformed);
        }
    }
    Ok(match identifier {
        Some(0) if !has_extra_fields => ReferenceIdentifierShape::Default,
        Some(0) => ReferenceIdentifierShape::Malformed,
        Some(_) => ReferenceIdentifierShape::Nonzero,
        None => ReferenceIdentifierShape::Malformed,
    })
}

fn classify_candidate(
    message_type: u32,
    source: &[u8],
    budget: &mut ProbeBudget,
) -> Result<CandidateClassification> {
    const LEGACY_MODEL_TYPE: u32 = 6_000;
    const CANONICAL_MODEL_TYPE: u32 = 6_001;
    const DATA_STORE: u8 = 1 << 0;
    const ROW_COUNT: u8 = 1 << 1;
    const COLUMN_COUNT: u8 = 1 << 2;
    const TABLE_NAME: u8 = 1 << 3;
    const REQUIRED: u8 = DATA_STORE | ROW_COUNT | COLUMN_COUNT | TABLE_NAME;

    if !matches!(message_type, LEGACY_MODEL_TYPE | CANONICAL_MODEL_TYPE) {
        return Ok(CandidateClassification::NotModel);
    }

    let mut present = 0_u8;
    let mut invalid_selected = false;
    let mut dense_style_reference = false;
    let mut seen_known = 0_u128;
    let mut scanner = ShapeScanner::new(*budget);
    let scan = (|| -> litchi_iwa_common::Result<()> {
        scanner.charge_message(source.len(), 0)?;
        let mut offset = 0;
        while offset < source.len() {
            let field = parse_shape_field(source, offset, 0, &mut scanner)?;
            offset = field.end;
            if let Some(valid_wire) = valid_known_root_wire_type(field.number, field.wire_type) {
                let bit = 1_u128 << field.number;
                let repeated = matches!(field.number, 90..=92);
                let canonical_framing =
                    field.canonical_key && (field.wire_type != 2 || field.canonical_length);
                if !valid_wire || !canonical_framing || (!repeated && seen_known & bit != 0) {
                    invalid_selected = true;
                }
                seen_known |= bit;
                if field.wire_type == 0 {
                    let canonical = litchi_iwa_common::decode_varint_from_bytes(field.payload)
                        .is_ok_and(|(value, width)| {
                            width == field.payload.len()
                                && width == litchi_iwa_common::varint::encoded_len(value)
                        });
                    invalid_selected |= !canonical;
                }
            }
            let (bit, expected_wire_type) = match field.number {
                4 => (DATA_STORE, 2),
                6 => (ROW_COUNT, 0),
                7 => (COLUMN_COUNT, 0),
                8 => (TABLE_NAME, 2),
                _ => continue,
            };
            if field.wire_type != expected_wire_type || present & bit != 0 {
                invalid_selected = true;
                continue;
            }
            if matches!(field.number, 6 | 7) {
                let canonical = litchi_iwa_common::decode_varint_from_bytes(field.payload)
                    .is_ok_and(|(value, width)| {
                        width == field.payload.len() && u32::try_from(value).is_ok()
                    });
                invalid_selected |= !canonical;
            } else if field.number == 8 && std::str::from_utf8(field.payload).is_err() {
                invalid_selected = true;
            } else if field.number == 4 {
                scanner.charge_message(field.payload.len(), 1)?;
                let mut nested_offset = 0;
                let mut seen_metadata_references = 0_u8;
                while nested_offset < field.payload.len() {
                    let nested = parse_shape_field(field.payload, nested_offset, 1, &mut scanner)?;
                    nested_offset = nested.end;
                    // HeaderStorage and zero/default metadata references are
                    // valid generated proto2 compatibility data. The style
                    // table reference is the native envelope's ownership
                    // marker: old/editor-created models often materialise a
                    // column-header reference while retaining a generated
                    // zero style reference. Such a model must stay on the
                    // compatibility projection; merely seeing field 2 must
                    // not silently upgrade it to strict decoding.
                    let bit = match nested.number {
                        2 => 1,
                        5 => 2,
                        _ => continue,
                    };
                    if nested.wire_type != 2 || seen_metadata_references & bit != 0 {
                        invalid_selected = true;
                        continue;
                    }
                    seen_metadata_references |= bit;
                    match classify_reference_identifier(nested.payload, 2, &mut scanner)? {
                        ReferenceIdentifierShape::Default => {},
                        ReferenceIdentifierShape::Nonzero if nested.number == 5 => {
                            dense_style_reference = true;
                        },
                        ReferenceIdentifierShape::Nonzero => {},
                        ReferenceIdentifierShape::Malformed => invalid_selected = true,
                    }
                }
            }
            present |= bit;
        }
        Ok(())
    })();

    match scan {
        Ok(()) => budget.charge_shape_scan(scanner.report)?,
        Err(error @ litchi_iwa_common::Error::LimitExceeded { .. })
        | Err(error @ litchi_iwa_common::Error::Allocation { .. })
        | Err(error @ litchi_iwa_common::Error::InvalidLimit { .. }) => {
            return Err(Error::IwaCommon(error));
        },
        Err(litchi_iwa_common::Error::InvalidFormat(_))
            if message_type == LEGACY_MODEL_TYPE && present != REQUIRED =>
        {
            budget.charge_malformed(source)?;
            return Ok(CandidateClassification::NotModel);
        },
        Err(litchi_iwa_common::Error::InvalidFormat(_)) => {
            budget.charge_malformed(source)?;
            return Ok(CandidateClassification::Malformed);
        },
    }

    if message_type == LEGACY_MODEL_TYPE && present != REQUIRED {
        let model_dimension_signature = ROW_COUNT | COLUMN_COUNT;
        return Ok(
            if invalid_selected && present & model_dimension_signature == model_dimension_signature
            {
                CandidateClassification::Malformed
            } else {
                CandidateClassification::NotModel
            },
        );
    }
    if invalid_selected || present != REQUIRED {
        return Ok(CandidateClassification::Malformed);
    }
    Ok(CandidateClassification::Projection(
        if dense_style_reference {
            ProjectionKind::Strict
        } else {
            ProjectionKind::SparseCompatibility
        },
    ))
}

fn valid_known_root_wire_type(number: u32, wire_type: u8) -> Option<bool> {
    let valid = match number {
        1
        | 3..=5
        | 8
        | 18..=21
        | 23..=27
        | 30
        | 34..=36
        | 38
        | 39
        | 43..=49
        | 52
        | 60..=89
        | 93 => wire_type == 2,
        6 | 7 | 9..=15 | 22 | 28 | 29 | 31 | 32 | 37 | 40..=42 | 50 | 51 => wire_type == 0,
        16 | 17 | 33 => wire_type == 1,
        90..=92 => matches!(wire_type, 0 | 2),
        _ => return None,
    };
    Some(valid)
}

fn combined_observed(current: usize, observed: usize, context: &str) -> Result<usize> {
    current.checked_add(observed).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-model {context} counter overflows host usize"
        ))
    })
}

pub(crate) fn map_resource_error(
    error: table_model_codec::DecodeError,
    budget: ProbeBudget,
) -> Result<Error> {
    use table_model_codec::DecodeLimit;

    let mapped = match error.resource_limit() {
        Some(DecodeLimit::Bytes {
            observed,
            maximum: _,
        }) => Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::InputBytes,
            observed: combined_observed(budget.input_bytes, observed, "input-byte")?,
            limit: MAX_INPUT_BYTES,
        }),
        Some(DecodeLimit::Fields { observed, .. }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed: combined_observed(budget.fields, observed, "field")?,
                limit: MAX_FIELDS,
            })
        },
        Some(DecodeLimit::Work { observed, .. }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: combined_observed(budget.work, observed, "work")?,
                limit: MAX_WORK,
            })
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Nesting,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                limit: usize::try_from(maximum).unwrap_or(usize::MAX),
            })
        },
        Some(DecodeLimit::Allocation { requested }) => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table-model projection",
                amount: requested,
            })
        },
        Some(DecodeLimit::References { observed, .. }) => Error::InvalidFormat(format!(
            "Numbers table-model candidate reference limit exceeded: observed {}, limit {MAX_REFERENCES}",
            combined_observed(budget.references, observed, "reference")?
        )),
        Some(DecodeLimit::Text { observed, .. }) => Error::InvalidFormat(format!(
            "Numbers table-model candidate text-byte limit exceeded: observed {}, limit {MAX_TEXT_BYTES}",
            combined_observed(budget.text_bytes, observed, "text-byte")?
        )),
        Some(DecodeLimit::Retained { .. }) | None => Error::InvalidFormat(
            "Numbers table-model candidate exceeded an unsupported projection limit".to_owned(),
        ),
        Some(_) => Error::InvalidFormat(
            "Numbers table-model candidate exceeded an unknown projection limit".to_owned(),
        ),
    };
    Ok(mapped)
}

pub(super) fn probe_candidate(
    message_type: u32,
    source: &[u8],
    budget: &mut ProbeBudget,
) -> Result<CandidateProbe> {
    let projection = match classify_candidate(message_type, source, budget)? {
        CandidateClassification::Projection(projection) => projection,
        CandidateClassification::NotModel => return Ok(CandidateProbe::NotModel),
        CandidateClassification::Malformed => return Ok(CandidateProbe::Malformed),
    };
    let decoded = match projection {
        ProjectionKind::Strict => {
            table_model_codec::decode_table_model_with_report(source, budget.options(source))
        },
        ProjectionKind::SparseCompatibility => {
            table_model_codec::decode_table_model_compatibility_with_report(
                source,
                budget.options(source),
            )
        },
    };
    match decoded {
        Ok((_snapshot, report)) => {
            budget.charge_report(report)?;
            Ok(CandidateProbe::Valid)
        },
        Err(error) if error.resource_limit().is_some() => Err(map_resource_error(error, *budget)?),
        Err(_error) => {
            budget.charge_malformed(source)?;
            Ok(CandidateProbe::Malformed)
        },
    }
}

/// Select one bounded table-model candidate from an archive object.
///
/// Type 6000 is shared with modern `TableInfoArchive`. A canonical type-6001
/// payload is therefore authoritative whenever it is present. Legacy aliases
/// are considered only when the object has no canonical message, and only
/// after their exact model signature passes the generated-free projection.
pub(crate) fn select_candidate(
    messages: &[RawMessage],
    budget: &mut ProbeBudget,
    error: impl Fn(&str) -> Error,
) -> Result<Option<usize>> {
    const LEGACY_MODEL_TYPE: u32 = 6_000;
    const CANONICAL_MODEL_TYPE: u32 = 6_001;

    let has_canonical = messages
        .iter()
        .any(|message| message.type_ == CANONICAL_MODEL_TYPE);
    let mut selected = None;
    for (index, message) in messages.iter().enumerate().filter(|(_, message)| {
        message.type_ == CANONICAL_MODEL_TYPE
            || (!has_canonical && message.type_ == LEGACY_MODEL_TYPE)
    }) {
        match probe_candidate(message.type_, message.data.as_slice(), budget)? {
            CandidateProbe::Valid if selected.replace(index).is_some() => {
                return Err(error("has multiple Numbers table model payloads"));
            },
            CandidateProbe::Valid | CandidateProbe::NotModel => {},
            CandidateProbe::Malformed => {
                return Err(error("contains a malformed Numbers table model payload"));
            },
        }
    }
    Ok(selected)
}

#[cfg(test)]
mod tests {
    use super::*;

    const SPARSE_MODEL: &[u8] = &[0x22, 0x00, 0x30, 0x00, 0x38, 0x00, 0x42, 0x00];

    #[test]
    fn sparse_generated_defaults_are_a_bounded_compatibility_candidate() {
        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(6_001, SPARSE_MODEL, &mut budget).unwrap(),
            CandidateProbe::Valid
        );
        assert!(budget.fields >= 8);
    }

    #[test]
    fn legacy_table_info_is_not_a_table_model_candidate() {
        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(6_000, &[0x12, 0x00], &mut budget).unwrap(),
            CandidateProbe::NotModel
        );

        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(
                6_000,
                &[0x20, 0x00, 0x30, 0x00, 0x38, 0x00, 0x42, 0x00],
                &mut budget,
            )
            .unwrap(),
            CandidateProbe::Malformed
        );
    }

    #[test]
    fn canonical_candidate_requires_the_selected_model_signature() {
        for payload in [
            vec![],
            [SPARSE_MODEL, &[0x18, 0x00]].concat(),
            [SPARSE_MODEL, &[0x42, 0x00]].concat(),
        ] {
            let mut budget = ProbeBudget::new();
            assert_eq!(
                probe_candidate(6_001, payload.as_slice(), &mut budget).unwrap(),
                CandidateProbe::Malformed
            );
        }
    }

    #[test]
    fn balanced_unknown_groups_remain_opaque_compatibility_data() {
        let payload = [SPARSE_MODEL, &[0xa3, 0x06, 0x08, 0x07, 0xa4, 0x06]].concat();
        for message_type in [6_000, 6_001] {
            let mut budget = ProbeBudget::new();
            assert_eq!(
                probe_candidate(message_type, payload.as_slice(), &mut budget).unwrap(),
                CandidateProbe::Valid
            );
        }
    }

    #[test]
    fn default_data_store_metadata_does_not_force_dense_projection() {
        let store = [
            0x0a, 0x01, 0xff, // opaque HeaderStorage metadata
            0x12, 0x00, // empty column-header Reference
            0x2a, 0x02, 0x08, 0x00, // explicit zero style Reference
        ];
        let mut model = vec![0x22, u8::try_from(store.len()).unwrap()];
        model.extend_from_slice(&store);
        model.extend_from_slice(&[0x30, 0x00, 0x38, 0x00, 0x42, 0x00]);

        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(6_001, model.as_slice(), &mut budget).unwrap(),
            CandidateProbe::Valid
        );
    }

    #[test]
    fn nonzero_sparse_data_store_metadata_keeps_compatibility_projection() {
        // A generated/editor-created model can materialise just the selected
        // column-header route while omitting the remaining proto2-required
        // DataStore metadata.  The non-zero route must not force a strict
        // decode of those intentionally omitted defaults.
        let store = [
            0x12, 0x02, 0x08, 0x2b, // column-header Reference { identifier: 43 }
        ];
        let mut model = vec![0x22, u8::try_from(store.len()).unwrap()];
        model.extend_from_slice(&store);
        model.extend_from_slice(&[0x30, 0x00, 0x38, 0x00, 0x42, 0x00]);

        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(6_001, model.as_slice(), &mut budget).unwrap(),
            CandidateProbe::Valid
        );
    }

    #[test]
    fn malformed_sparse_metadata_references_do_not_bypass_strict_admission() {
        for store in [
            vec![0x10, 0x00],             // selected column-header field has the wrong wire type
            vec![0x12, 0x00, 0x12, 0x00], // duplicate selected reference
            vec![0x12, 0x01, 0x80],       // malformed nested reference
            vec![0x12, 0x04, 0x08, 0x00, 0x10, 0x01], // zero identifier plus trailing data
        ] {
            let mut model = vec![0x22, u8::try_from(store.len()).unwrap()];
            model.extend_from_slice(&store);
            model.extend_from_slice(&[0x30, 0x00, 0x38, 0x00, 0x42, 0x00]);

            let mut budget = ProbeBudget::new();
            assert_eq!(
                probe_candidate(6_001, model.as_slice(), &mut budget).unwrap(),
                CandidateProbe::Malformed
            );
        }
    }

    #[test]
    fn malformed_candidate_is_charged_without_becoming_valid() {
        let mut budget = ProbeBudget::new();
        assert_eq!(
            probe_candidate(6_001, &[0x80], &mut budget).unwrap(),
            CandidateProbe::Malformed
        );
        assert!(budget.input_bytes >= 1);
        assert!(budget.work >= 32);
    }

    #[test]
    fn exhausted_aggregate_field_budget_is_typed() {
        let mut budget = ProbeBudget {
            fields: MAX_FIELDS,
            ..ProbeBudget::new()
        };
        let error = probe_candidate(6_001, &[0x42, 0x01, b'T'], &mut budget)
            .expect_err("one field exceeds the residual aggregate ceiling");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: MAX_FIELDS,
            }) if observed > MAX_FIELDS
        ));
    }

    #[test]
    fn malformed_candidate_cannot_overrun_aggregate_work() {
        let mut budget = ProbeBudget {
            work: MAX_WORK,
            ..ProbeBudget::new()
        };
        let error = probe_candidate(6_001, &[0x80], &mut budget)
            .expect_err("malformed work exceeds the aggregate ceiling");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed,
                limit: MAX_WORK,
            }) if observed > MAX_WORK
        ));
    }

    #[test]
    fn successful_candidates_share_one_atomic_field_budget() {
        let mut budget = ProbeBudget {
            fields: MAX_FIELDS - 8,
            ..ProbeBudget::new()
        };

        assert_eq!(
            probe_candidate(6_001, SPARSE_MODEL, &mut budget).unwrap(),
            CandidateProbe::Valid
        );
        assert_eq!(budget.fields, MAX_FIELDS);

        let before = budget;
        let error = probe_candidate(6_001, SPARSE_MODEL, &mut budget)
            .expect_err("later valid candidate exceeds the aggregate field budget");
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: MAX_FIELDS,
            }) if observed > MAX_FIELDS
        ));
        assert_eq!(budget, before);
    }
}
