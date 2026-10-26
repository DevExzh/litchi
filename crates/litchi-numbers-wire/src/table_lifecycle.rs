//! Source-preserving object-reference rewrites for the Numbers table graph
//! lifecycle (duplicate and remove), shared by the three table owners.
//!
//! Ported in record 0368 from the donor `numbers_table_lifecycle_codec`
//! (record 0363 W0-B). The rewrite changes only the identifiers of
//! schema-selected `TSP.Reference` envelopes and, for a model, its UUID and
//! name; every other field, unknown field and byte of framing is copied
//! verbatim. One caller-selected [`Limits`] governs a whole call: input and
//! output bytes, parsed fields per message and the aggregate number of
//! visited fields ([`LimitKind::Items`]).

use core::fmt;
use std::cell::Cell;

use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into, encoded_len};
use litchi_iwa_common::wire::{WireFieldView, WireView};

use crate::{LimitKind, Limits};

/// One native object-reference replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ReferenceRemap {
    /// The existing native object identifier.
    pub old: u64,
    /// The fresh native object identifier.
    pub new: u64,
}

/// The object-reference paths of `TST.TableModelArchive`.
///
/// A path ending in a repeated field rewrites every occurrence.
pub const TABLE_MODEL_REFERENCE_PATHS: &[&[u32]] = &[
    &[4, 1, 2],
    &[4, 2],
    &[4, 3, 1, 2],
    &[4, 4],
    &[4, 5],
    &[4, 6],
    &[4, 11],
    &[4, 12],
    &[4, 13],
    &[4, 15],
    &[4, 16],
    &[4, 17],
    &[4, 18],
    &[4, 19],
    &[4, 20],
    &[4, 21],
    &[4, 22],
    &[46],
    &[49],
];

/// The object-reference paths of `TST.TableInfoArchive`.
pub const TABLE_INFO_REFERENCE_PATHS: &[&[u32]] = &[
    &[1, 2],
    &[1, 6],
    &[1, 9],
    &[2],
    &[3],
    &[4],
    &[5],
    &[6],
    &[15],
    &[17],
];

/// The segment-reference path of `TST.TableDataList`.
const TABLE_DATA_LIST_REFERENCE_PATHS: &[&[u32]] = &[&[4]];

/// A typed table-lifecycle rewrite failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The source is not valid wire, a selected field has the wrong wire
    /// type, a reference identifier is missing, repeated or malformed, or a
    /// path is empty.
    InvalidSource,
    /// A caller-selected budget was exceeded.
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The observed or requested amount.
        observed: usize,
        /// The caller-selected maximum.
        maximum: usize,
    },
    /// An output reservation failed.
    Allocation {
        /// The bytes requested.
        requested: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSource => formatter.write_str("invalid Numbers table lifecycle wire"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "Numbers table lifecycle {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { requested } => write!(
                formatter,
                "could not allocate {requested} Numbers table lifecycle bytes"
            ),
        }
    }
}

impl std::error::Error for Error {}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Self::Limit {
                kind: match kind {
                    litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
                    litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                    litchi_iwa_common::LimitKind::Fields => LimitKind::Fields,
                    _ => LimitKind::Items,
                },
                observed,
                maximum: limit,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => {
                Self::Allocation { requested: amount }
            },
            _ => Self::InvalidSource,
        }
    }
}

/// Result of a table-lifecycle rewrite.
pub type Result<T> = std::result::Result<T, Error>;

/// Rewrite a model's UUID (field 1), name (field 8) and owned references.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the model has no single UUID or name field
/// or a selected reference is malformed; [`Error::Limit`] /
/// [`Error::Allocation`] when a budget or reservation fails.
pub fn rewrite_table_model(
    source: &[u8],
    table_uuid: &[u8],
    table_name: &[u8],
    remap: &[ReferenceRemap],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let work = Cell::new(0);
    let output = replace_singular(source, 1, table_uuid, limits, &work)?;
    let output = replace_singular(&output, 8, table_name, limits, &work)?;
    remap_paths(&output, TABLE_MODEL_REFERENCE_PATHS, remap, limits, &work)
}

/// Rewrite the owned references of a table-info payload.
///
/// # Errors
///
/// As [`remap_reference_paths`].
pub fn rewrite_table_info(
    source: &[u8],
    remap: &[ReferenceRemap],
    limits: &Limits,
) -> Result<Vec<u8>> {
    remap_reference_paths(source, TABLE_INFO_REFERENCE_PATHS, remap, limits)
}

/// Rewrite the segment references of one `TST.TableDataList` payload.
///
/// # Errors
///
/// As [`remap_reference_paths`].
pub fn rewrite_table_data_list(
    source: &[u8],
    remap: &[ReferenceRemap],
    limits: &Limits,
) -> Result<Vec<u8>> {
    remap_reference_paths(source, TABLE_DATA_LIST_REFERENCE_PATHS, remap, limits)
}

/// Rewrite the identifier of one raw `TSP.Reference` envelope; an unmapped
/// identifier returns the source bytes.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing, repeated or non-varint
/// identifier; [`Error::Limit`] / [`Error::Allocation`].
pub fn rewrite_reference(
    source: &[u8],
    remap: &[ReferenceRemap],
    limits: &Limits,
) -> Result<Vec<u8>> {
    rewrite_reference_counted(source, remap, limits, &Cell::new(0))
}

/// Rewrite every selected reference envelope at every occurrence of each
/// path. Absent optional paths are preserved; a present malformed path fails
/// before any output is returned.
///
/// # Errors
///
/// [`Error::InvalidSource`] for an empty path or malformed selected field;
/// [`Error::Limit`] / [`Error::Allocation`].
pub fn remap_reference_paths(
    source: &[u8],
    paths: &[&[u32]],
    remap: &[ReferenceRemap],
    limits: &Limits,
) -> Result<Vec<u8>> {
    remap_paths(source, paths, remap, limits, &Cell::new(0))
}

fn remap_paths(
    source: &[u8],
    paths: &[&[u32]],
    remap: &[ReferenceRemap],
    limits: &Limits,
    work: &Cell<usize>,
) -> Result<Vec<u8>> {
    let mut output = copy(source, limits)?;
    for path in paths {
        output = rewrite_path(&output, path, remap, limits, work)?;
    }
    Ok(output)
}

fn view<'a>(source: &'a [u8], limits: &Limits, work: &Cell<usize>) -> Result<WireView<'a>> {
    if source.len() > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: source.len(),
            maximum: limits.max_input_bytes(),
        });
    }
    let view = WireView::parse(source, limits.wire())?;
    let visited = work.get().saturating_add(view.fields().count());
    if visited > limits.max_items() {
        return Err(Error::Limit {
            kind: LimitKind::Items,
            observed: visited,
            maximum: limits.max_items(),
        });
    }
    work.set(visited);
    Ok(view)
}

fn rewrite_reference_counted(
    source: &[u8],
    remap: &[ReferenceRemap],
    limits: &Limits,
    work: &Cell<usize>,
) -> Result<Vec<u8>> {
    let view = view(source, limits, work)?;
    let mut selected: Option<WireFieldView<'_>> = None;
    for field in view.fields() {
        if field.number() != 1 {
            continue;
        }
        if selected.is_some() || field.wire_type() != 0 {
            return Err(Error::InvalidSource);
        }
        selected = Some(field);
    }
    let field = selected.ok_or(Error::InvalidSource)?;
    let (identifier, length) =
        decode_varint_from_bytes(field.payload()).map_err(|_error| Error::InvalidSource)?;
    if length != field.payload().len() {
        return Err(Error::InvalidSource);
    }
    let Some(replacement) = remap
        .iter()
        .find(|candidate| candidate.old == identifier)
        .map(|candidate| candidate.new)
    else {
        return copy(source, limits);
    };
    let mut output = Vec::new();
    let key = field.key();
    let start = offset_of(source, field.raw());
    let end = start + field.raw().len();
    reserve(
        &mut output,
        source.len() - field.payload().len() + encoded_len(replacement),
        limits,
    )?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(key);
    encode_varint_into(&mut output, replacement);
    output.extend_from_slice(&source[end..]);
    Ok(output)
}

fn rewrite_path(
    source: &[u8],
    path: &[u32],
    remap: &[ReferenceRemap],
    limits: &Limits,
    work: &Cell<usize>,
) -> Result<Vec<u8>> {
    let (&head, tail) = path.split_first().ok_or(Error::InvalidSource)?;
    let view = view(source, limits, work)?;
    let mut replacements = Vec::new();
    for (index, field) in view.fields().enumerate() {
        if field.number() != head || field.wire_type() != 2 {
            continue;
        }
        let payload = if tail.is_empty() {
            rewrite_reference_counted(field.payload(), remap, limits, work)?
        } else {
            rewrite_path(field.payload(), tail, remap, limits, work)?
        };
        if payload.as_slice() != field.payload() {
            replacements.push((index, payload));
        }
    }
    if replacements.is_empty() {
        return copy(source, limits);
    }
    let mut output_len = source.len();
    for (index, payload) in &replacements {
        let field = view.get(*index).ok_or(Error::InvalidSource)?;
        output_len = output_len - (field.raw().len() - field.key().len())
            + encoded_len(payload.len() as u64)
            + payload.len();
    }
    let mut output = Vec::new();
    reserve(&mut output, output_len, limits)?;
    let mut next = replacements.iter().peekable();
    for (index, field) in view.fields().enumerate() {
        match next.peek() {
            Some((replacement_index, payload)) if *replacement_index == index => {
                output.extend_from_slice(field.key());
                encode_varint_into(&mut output, payload.len() as u64);
                output.extend_from_slice(payload);
                next.next();
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    Ok(output)
}

fn replace_singular(
    source: &[u8],
    number: u32,
    payload: &[u8],
    limits: &Limits,
    work: &Cell<usize>,
) -> Result<Vec<u8>> {
    let view = view(source, limits, work)?;
    let mut selected = None;
    for (index, field) in view.fields().enumerate() {
        if field.number() == number {
            if selected.is_some() || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            selected = Some(index);
        }
    }
    let field = view
        .get(selected.ok_or(Error::InvalidSource)?)
        .ok_or(Error::InvalidSource)?;
    if field.payload() == payload {
        return copy(source, limits);
    }
    // An equal-length payload keeps its original length prefix.
    let prefix_len = field.raw().len() - field.key().len() - field.payload().len();
    let start = offset_of(source, field.raw());
    let end = start + field.raw().len();
    let same_length = field.payload().len() == payload.len();
    let length_len = if same_length {
        prefix_len
    } else {
        encoded_len(payload.len() as u64)
    };
    let mut output = Vec::new();
    reserve(
        &mut output,
        source.len() - field.raw().len() + field.key().len() + length_len + payload.len(),
        limits,
    )?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(field.key());
    if same_length {
        let key_end = start + field.key().len();
        output.extend_from_slice(&source[key_end..key_end + prefix_len]);
    } else {
        encode_varint_into(&mut output, payload.len() as u64);
    }
    output.extend_from_slice(payload);
    output.extend_from_slice(&source[end..]);
    Ok(output)
}

/// The offset of a sub-slice the wire view borrowed from `source`.
fn offset_of(source: &[u8], part: &[u8]) -> usize {
    part.as_ptr() as usize - source.as_ptr() as usize
}

fn reserve(output: &mut Vec<u8>, capacity: usize, limits: &Limits) -> Result<()> {
    if capacity > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: capacity,
            maximum: limits.max_output_bytes(),
        });
    }
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Error::Allocation {
            requested: capacity,
        })
}

fn copy(source: &[u8], limits: &Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, source.len(), limits)?;
    output.extend_from_slice(source);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, field: u32, value: u64) {
        encode_varint_into(output, u64::from(field) << 3);
        encode_varint_into(output, value);
    }

    fn bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
        encode_varint_into(output, u64::from(field) << 3 | 2);
        encode_varint_into(output, payload.len() as u64);
        output.extend_from_slice(payload);
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        varint(&mut output, 1, identifier);
        output
    }

    fn model() -> Vec<u8> {
        let mut nested = Vec::new();
        bytes(&mut nested, 2, &reference(7));
        varint(&mut nested, 99, 123);
        let mut source = Vec::new();
        bytes(&mut source, 1, b"old-uuid");
        bytes(&mut source, 8, b"Old");
        bytes(&mut source, 4, &nested);
        varint(&mut source, 98, 456);
        source
    }

    #[test]
    fn model_reference_paths_rewrite_only_selected_envelopes() {
        let rewritten = rewrite_table_model(
            &model(),
            b"00000000-0000-4000-8000-000000000007",
            b"Copy",
            &[ReferenceRemap { old: 7, new: 9 }],
            &Limits::DEFAULT,
        )
        .unwrap();
        let mut nested = Vec::new();
        bytes(&mut nested, 2, &reference(9));
        varint(&mut nested, 99, 123);
        let mut expected = Vec::new();
        bytes(&mut expected, 1, b"00000000-0000-4000-8000-000000000007");
        bytes(&mut expected, 8, b"Copy");
        bytes(&mut expected, 4, &nested);
        varint(&mut expected, 98, 456);
        assert_eq!(rewritten, expected);
    }

    #[test]
    fn malformed_reference_is_rejected_before_rewrite() {
        let mut malformed = Vec::new();
        bytes(&mut malformed, 1, &[7]);
        let mut source = Vec::new();
        bytes(&mut source, 1, b"uuid");
        bytes(&mut source, 8, b"name");
        let mut nested = Vec::new();
        bytes(&mut nested, 2, &malformed);
        bytes(&mut source, 4, &nested);
        assert_eq!(
            rewrite_table_model(&source, b"uuid", b"name", &[], &Limits::DEFAULT),
            Err(Error::InvalidSource)
        );
        assert_eq!(
            remap_reference_paths(&source, &[&[]], &[], &Limits::DEFAULT),
            Err(Error::InvalidSource)
        );
    }

    #[test]
    fn unmapped_reference_and_unchanged_model_are_byte_exact() {
        let source = reference(7);
        assert_eq!(
            rewrite_reference(&source, &[], &Limits::DEFAULT).unwrap(),
            source
        );
        let model = model();
        assert_eq!(
            rewrite_table_model(&model, b"old-uuid", b"Old", &[], &Limits::DEFAULT).unwrap(),
            model
        );
    }

    #[test]
    fn budgets_are_checked() {
        let model = model();
        let narrow = Limits::DEFAULT.with_items(3).unwrap();
        assert!(matches!(
            rewrite_table_model(&model, b"u", b"n", &[], &narrow),
            Err(Error::Limit {
                kind: LimitKind::Items,
                ..
            })
        ));
        let small = Limits::DEFAULT.with_output_bytes(4).unwrap();
        assert!(matches!(
            rewrite_table_info(&model, &[], &small),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
        let mut list = Vec::new();
        bytes(&mut list, 4, &reference(5));
        let mut expected = Vec::new();
        bytes(&mut expected, 4, &reference(500));
        assert_eq!(
            rewrite_table_data_list(
                &list,
                &[ReferenceRemap { old: 5, new: 500 }],
                &Limits::DEFAULT
            )
            .unwrap(),
            expected
        );
    }
}
