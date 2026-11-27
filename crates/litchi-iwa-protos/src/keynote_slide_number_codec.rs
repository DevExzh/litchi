//! Strict, generated-free snapshots for Keynote slide-number records.
//!
//! Buffa's private lazy projection is compiled alongside this module; this
//! boundary keeps the only repeated native table as caller-owned raw bytes.

use crate::buffa_keynote_slide_number_generated::LitchiIwaProjection as projection;
use buffa::DecodeOptions as BuffaDecodeOptions;
use std::cell::Cell;
use std::{fmt, str};

const NODE_VISIBILITY_FIELD: u32 = 18;
const STORAGE_KIND_FIELD: u32 = 1;
const STORAGE_TEXT_FIELD: u32 = 3;
const STORAGE_ATTACHMENT_TABLE_FIELD: u32 = 9;
const STORAGE_IN_DOCUMENT_FIELD: u32 = 10;
const ATTACHMENT_TABLE_ENTRIES_FIELD: u32 = 1;
const ATTACHMENT_CHARACTER_INDEX_FIELD: u32 = 1;
const ATTACHMENT_OBJECT_FIELD: u32 = 2;
const SLIDE_NUMBER_ATTACHMENT_SUPER_FIELD: u32 = 1;
const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;
const TEXTUAL_KIND_FIELD: u32 = 2;
const MAX_RECURSION: u32 = 64;

/// Finite policy shared by each focused decoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}
impl DecodeOptions {
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
}

/// Exact successful decoder consumption, suitable for transaction budget merge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}
impl DecodeReport {
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
}

/// Typed resource-policy failure observation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}

/// Content-free strict wire failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    message: &'static str,
    limit: Option<DecodeLimit>,
}
impl DecodeError {
    const fn plain(message: &'static str) -> Self {
        Self {
            message,
            limit: None,
        }
    }
    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            message: "slide-number resource limit exceeded",
            limit: Some(limit),
        }
    }
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }
}
impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.message)
    }
}
impl std::error::Error for DecodeError {}

/// Visibility selector observed from `KN.SlideNodeArchive` field 18.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNumberNodeSnapshot {
    visibility: Option<bool>,
}
impl SlideNumberNodeSnapshot {
    #[must_use]
    pub const fn visibility(self) -> Option<bool> {
        self.visibility
    }
}

/// Scalar storage facts; the borrowed table payload is validated but not owned.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNumberStorageSnapshot<'a> {
    kind: Option<i32>,
    in_document: Option<bool>,
    attachment_table: Option<&'a [u8]>,
}
impl<'a> SlideNumberStorageSnapshot<'a> {
    #[must_use]
    pub const fn kind(self) -> Option<i32> {
        self.kind
    }
    #[must_use]
    pub const fn in_document(self) -> Option<bool> {
        self.in_document
    }
    #[must_use]
    pub const fn attachment_table(self) -> Option<&'a [u8]> {
        self.attachment_table
    }
}

/// Borrowed textual facts from a slide-number attachment's required super.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNumberAttachmentSnapshot<'a> {
    string_equivalent: Option<&'a str>,
    kind: Option<i32>,
}
impl<'a> SlideNumberAttachmentSnapshot<'a> {
    #[must_use]
    pub const fn string_equivalent(self) -> Option<&'a str> {
        self.string_equivalent
    }
    #[must_use]
    pub const fn kind(self) -> Option<i32> {
        self.kind
    }
}

#[allow(
    non_snake_case,
    reason = "Retains concise internal error construction."
)]
const fn DecodeError(message: &'static str) -> DecodeError {
    DecodeError::plain(message)
}

pub fn decode_slide_number_node(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SlideNumberNodeSnapshot, DecodeError> {
    Ok(decode_slide_number_node_with_report(source, options)?.0)
}
pub fn decode_slide_number_node_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SlideNumberNodeSnapshot, DecodeReport), DecodeError> {
    let aggregate = AggregateBudget::new(options);
    aggregate.message(source.len())?;
    let mut p = Parser::new(source, options, &aggregate, 1)?;
    let mut visibility = None;
    while let Some(f) = p.field()? {
        if f.number == NODE_VISIBILITY_FIELD {
            singular(
                &mut visibility,
                "duplicate KN.SlideNodeArchive.isSlideNumberVisible",
            )?;
            visibility = Some(boolean(f.varint()?)?);
        }
    }
    let strict = SlideNumberNodeSnapshot { visibility };
    let view: projection::SlideNumberNodeArchiveLazyView<'_> = buffa(options, source)?;
    if view.is_slide_number_visible != strict.visibility {
        return Err(DecodeError("Buffa slide-number node projection disagrees"));
    }
    Ok((strict, aggregate.report()))
}

pub fn decode_slide_number_storage(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SlideNumberStorageSnapshot<'_>, DecodeError> {
    Ok(decode_slide_number_storage_with_report(source, options)?.0)
}
pub fn decode_slide_number_storage_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SlideNumberStorageSnapshot<'_>, DecodeReport), DecodeError> {
    let aggregate = AggregateBudget::new(options);
    aggregate.message(source.len())?;
    let mut p = Parser::new(source, options, &aggregate, 1)?;
    let mut kind = None;
    let mut in_document = None;
    let mut attachment_table = None;
    let mut text_fragment = None;
    while let Some(f) = p.field()? {
        match f.number {
            STORAGE_KIND_FIELD => {
                singular(&mut kind, "duplicate TSWP.StorageArchive.kind")?;
                kind = Some(int32(f.varint()?)?);
            },
            STORAGE_ATTACHMENT_TABLE_FIELD => {
                singular(
                    &mut attachment_table,
                    "duplicate TSWP.StorageArchive.table_attachment",
                )?;
                let table = f.bytes()?;
                validate_table(table, p.options, &aggregate)?;
                attachment_table = Some(table);
            },
            STORAGE_TEXT_FIELD => {
                singular(&mut text_fragment, "duplicate TSWP.StorageArchive.text")?;
                let fragment = str::from_utf8(f.bytes()?)
                    .map_err(|_utf8_error| DecodeError("invalid UTF-8 storage text"))?;
                if fragment != "\u{fffc}" {
                    return Err(DecodeError(
                        "slide-number storage text is not one object replacement character",
                    ));
                }
                text_fragment = Some(());
            },
            STORAGE_IN_DOCUMENT_FIELD => {
                singular(
                    &mut in_document,
                    "duplicate TSWP.StorageArchive.in_document",
                )?;
                in_document = Some(boolean(f.varint()?)?);
            },
            _ => {},
        }
    }
    if text_fragment.is_none() {
        return Err(DecodeError("missing TSWP.StorageArchive.text"));
    }
    if attachment_table.is_none() {
        return Err(DecodeError("missing TSWP.StorageArchive.table_attachment"));
    }
    let strict = SlideNumberStorageSnapshot {
        kind,
        in_document,
        attachment_table,
    };
    let view: projection::SlideNumberStorageArchiveLazyView<'_> = buffa(options, source)?;
    if view.kind != strict.kind
        || view.in_document != strict.in_document
        || view.attachment_table != strict.attachment_table
    {
        return Err(DecodeError(
            "Buffa slide-number storage projection disagrees",
        ));
    }
    Ok((strict, aggregate.report()))
}

pub fn decode_slide_number_attachment(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SlideNumberAttachmentSnapshot<'_>, DecodeError> {
    Ok(decode_slide_number_attachment_with_report(source, options)?.0)
}
pub fn decode_slide_number_attachment_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SlideNumberAttachmentSnapshot<'_>, DecodeReport), DecodeError> {
    let aggregate = AggregateBudget::new(options);
    aggregate.message(source.len())?;
    let mut p = Parser::new(source, options, &aggregate, 1)?;
    let mut super_payload = None;
    while let Some(f) = p.field()? {
        if f.number == SLIDE_NUMBER_ATTACHMENT_SUPER_FIELD {
            singular(
                &mut super_payload,
                "duplicate KN.SlideNumberAttachmentArchive.super",
            )?;
            super_payload = Some(f.bytes()?);
        }
    }
    let textual_payload =
        super_payload.ok_or(DecodeError("missing KN.SlideNumberAttachmentArchive.super"))?;
    aggregate.message(textual_payload.len())?;
    let mut text = Parser::new(textual_payload, options, &aggregate, 2)?;
    let mut string_equivalent = None;
    let mut kind = None;
    while let Some(f) = text.field()? {
        match f.number {
            TEXTUAL_STRING_EQUIVALENT_FIELD => {
                singular(
                    &mut string_equivalent,
                    "duplicate TSWP.TextualAttachmentArchive.string_equivalent",
                )?;
                string_equivalent = Some(
                    str::from_utf8(f.bytes()?)
                        .map_err(|_utf8_error| DecodeError("invalid UTF-8 textual attachment"))?,
                );
            },
            TEXTUAL_KIND_FIELD => {
                singular(&mut kind, "duplicate TSWP.TextualAttachmentArchive.kind")?;
                kind = Some(int32(f.varint()?)?);
            },
            _ => {},
        }
    }
    let strict = SlideNumberAttachmentSnapshot {
        string_equivalent,
        kind,
    };
    let view: projection::SlideNumberAttachmentArchiveLazyView<'_> = buffa(options, source)?;
    let super_view = view
        .super_
        .get()
        .map_err(|_buffa_error| DecodeError("Buffa slide-number attachment projection failed"))?
        .ok_or(DecodeError("Buffa slide-number attachment super absent"))?;
    if super_view.string_equivalent != strict.string_equivalent || super_view.kind != strict.kind {
        return Err(DecodeError(
            "Buffa slide-number attachment projection disagrees",
        ));
    }
    Ok((strict, aggregate.report()))
}

fn buffa<'a, T: buffa::LazyMessageView<'a>>(
    options: DecodeOptions,
    source: &'a [u8],
) -> Result<T, DecodeError> {
    BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(0)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_buffa_error| DecodeError("Buffa slide-number projection failed"))
}

fn singular<T>(value: &mut Option<T>, error: &'static str) -> Result<(), DecodeError> {
    if value.is_some() {
        Err(DecodeError(error))
    } else {
        Ok(())
    }
}
fn boolean(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError("non-canonical bool")),
    }
}
fn int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(result) = i32::try_from(value) {
        return Ok(result);
    }
    if value >= 0xffff_ffff_8000_0000 {
        let signed = i64::from_ne_bytes(value.to_ne_bytes());
        return i32::try_from(signed)
            .map_err(|_conversion_error| DecodeError("non-canonical int32"));
    }
    Err(DecodeError("non-canonical int32"))
}

fn validate_table(
    source: &[u8],
    options: DecodeOptions,
    aggregate: &AggregateBudget,
) -> Result<(), DecodeError> {
    aggregate.message(source.len())?;
    let mut table = Parser::new(source, options, aggregate, 2)?;
    let mut entries = 0usize;
    while let Some(f) = table.field()? {
        if f.number == ATTACHMENT_TABLE_ENTRIES_FIELD {
            entries = entries
                .checked_add(1)
                .ok_or(DecodeError("attachment entry count overflow"))?;
            if entries != 1 {
                return Err(DecodeError("multiple slide-number attachment entries"));
            }
            let entry_payload = f.bytes()?;
            aggregate.message(entry_payload.len())?;
            let mut entry = Parser::new(entry_payload, options, aggregate, 3)?;
            let mut character = None;
            let mut object = None::<()>;
            while let Some(field) = entry.field()? {
                match field.number {
                    ATTACHMENT_CHARACTER_INDEX_FIELD => {
                        singular(&mut character, "duplicate ObjectAttribute.character_index")?;
                        let index = field.varint()?;
                        if index != 0 {
                            return Err(DecodeError(
                                "slide-number attachment character index is not zero",
                            ));
                        }
                        character = Some(index);
                    },
                    ATTACHMENT_OBJECT_FIELD => {
                        singular(&mut object, "duplicate ObjectAttribute.object")?;
                        validate_reference(field.bytes()?, options, aggregate)?;
                        object = Some(());
                    },
                    _ => {},
                }
            }
            if character.is_none() {
                return Err(DecodeError("missing ObjectAttribute.character_index"));
            }
            if object.is_none() {
                return Err(DecodeError("missing ObjectAttribute.object"));
            }
        }
    }
    if entries != 1 {
        return Err(DecodeError("missing slide-number attachment entry"));
    }
    Ok(())
}
fn validate_reference(
    source: &[u8],
    options: DecodeOptions,
    aggregate: &AggregateBudget,
) -> Result<(), DecodeError> {
    aggregate.message(source.len())?;
    let mut p = Parser::new(source, options, aggregate, 4)?;
    let mut id = None;
    let mut deprecated_type = None;
    let mut external = None;
    while let Some(f) = p.field()? {
        match f.number {
            1 => {
                singular(&mut id, "duplicate TSP.Reference.identifier")?;
                let value = f.varint()?;
                if value == 0 {
                    return Err(DecodeError("zero slide-number attachment identifier"));
                }
                id = Some(value);
            },
            2 => {
                singular(
                    &mut deprecated_type,
                    "duplicate TSP.Reference.deprecated_type",
                )?;
                deprecated_type = Some(f.varint()?);
            },
            3 => {
                singular(
                    &mut external,
                    "duplicate TSP.Reference.deprecated_is_external",
                )?;
                if boolean(f.varint()?)? {
                    return Err(DecodeError("external slide-number attachment reference"));
                }
                external = Some(());
            },
            _ => {},
        }
    }
    id.ok_or(DecodeError("missing TSP.Reference.identifier"))
        .map(|_| ())
}

#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: u8,
    value: &'a [u8],
    varint: u64,
}
impl<'a> Field<'a> {
    fn varint(self) -> Result<u64, DecodeError> {
        if self.wire == 0 {
            Ok(self.varint)
        } else {
            Err(DecodeError("wire type mismatch"))
        }
    }
    fn bytes(self) -> Result<&'a [u8], DecodeError> {
        if self.wire == 2 {
            Ok(self.value)
        } else {
            Err(DecodeError("wire type mismatch"))
        }
    }
}
struct Parser<'a, 'budget> {
    remaining: &'a [u8],
    options: DecodeOptions,
    aggregate: &'budget AggregateBudget,
}

struct AggregateBudget {
    fields: Cell<usize>,
    work: Cell<usize>,
    max_depth: Cell<u32>,
    max_fields: usize,
    maximum: usize,
}
impl AggregateBudget {
    fn new(options: DecodeOptions) -> Self {
        Self {
            fields: Cell::new(0),
            work: Cell::new(0),
            max_depth: Cell::new(0),
            max_fields: options.max_fields,
            maximum: options.max_work_bytes,
        }
    }
    fn message(&self, bytes: usize) -> Result<(), DecodeError> {
        let charge = bytes
            .checked_mul(2)
            .ok_or(DecodeError("work byte overflow"))?;
        let work = self
            .work
            .get()
            .checked_add(charge)
            .ok_or(DecodeError("work byte overflow"))?;
        if work > self.maximum {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: work,
                maximum: self.maximum,
            }));
        }
        self.work.set(work);
        Ok(())
    }
    fn field(&self) -> Result<(), DecodeError> {
        let fields = self
            .fields
            .get()
            .checked_add(1)
            .ok_or(DecodeError("field count overflow"))?;
        if fields > self.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: fields,
                maximum: self.max_fields,
            }));
        }
        self.fields.set(fields);
        Ok(())
    }
    fn depth(&self, depth: u32, maximum: u32) -> Result<(), DecodeError> {
        if depth > maximum {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum,
            }));
        }
        self.max_depth.set(self.max_depth.get().max(depth));
        Ok(())
    }
    fn report(&self) -> DecodeReport {
        DecodeReport {
            fields: self.fields.get(),
            work_bytes: self.work.get(),
            max_depth: self.max_depth.get(),
        }
    }
}
impl<'a, 'budget> Parser<'a, 'budget> {
    fn new(
        source: &'a [u8],
        options: DecodeOptions,
        aggregate: &'budget AggregateBudget,
        depth: u32,
    ) -> Result<Self, DecodeError> {
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limit(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        aggregate.depth(depth, options.recursion_limit)?;
        Ok(Self {
            remaining: source,
            options,
            aggregate,
        })
    }
    fn field(&mut self) -> Result<Option<Field<'a>>, DecodeError> {
        if self.remaining.is_empty() {
            return Ok(None);
        }
        self.aggregate.field()?;
        let tag = take(&mut self.remaining)?;
        if tag >> 3 == 0 || tag >> 3 > 0x1fff_ffff {
            return Err(DecodeError("invalid field number"));
        }
        let wire = (tag & 7) as u8;
        let number = u32::try_from(tag >> 3)
            .map_err(|_conversion_error| DecodeError("invalid field number"))?;
        match wire {
            0 => {
                let value = take(&mut self.remaining)?;
                Ok(Some(Field {
                    number,
                    wire,
                    value: &[],
                    varint: value,
                }))
            },
            2 => {
                let n = usize::try_from(take(&mut self.remaining)?)
                    .map_err(|_conversion_error| DecodeError("length overflow"))?;
                if self.remaining.len() < n {
                    return Err(DecodeError("truncated field"));
                }
                let (value, rest) = self.remaining.split_at(n);
                self.remaining = rest;
                Ok(Some(Field {
                    number,
                    wire,
                    value,
                    varint: 0,
                }))
            },
            _ => Err(DecodeError("unsupported wire type")),
        }
    }
}
fn take(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let mut value = 0u64;
    for index in 0..10 {
        let byte = *source.get(index).ok_or(DecodeError("truncated varint"))?;
        if index == 9 && byte > 1 {
            return Err(DecodeError("varint too long"));
        }
        value |= u64::from(byte & 127) << (index * 7);
        if byte & 128 == 0 {
            if (if value == 0 {
                1
            } else {
                (64 - value.leading_zeros()).div_ceil(7) as usize
            }) != index + 1
            {
                return Err(DecodeError("non-canonical varint"));
            }
            *source = &source[index + 1..];
            return Ok(value);
        }
    }
    Err(DecodeError("truncated varint"))
}

#[cfg(test)]
mod tests {
    use super::*;
    const OPTIONS: DecodeOptions = DecodeOptions::new(1024, 32, 4096, 4);
    #[test]
    fn node_visibility_is_optional() {
        assert_eq!(
            decode_slide_number_node(&[0x90, 1, 1], OPTIONS)
                .unwrap()
                .visibility(),
            Some(true)
        );
    }
    #[test]
    fn storage_validates_repeated_table_without_retaining_it() {
        let raw = storage_fixture();
        assert_eq!(
            decode_slide_number_storage(&raw, OPTIONS)
                .unwrap()
                .attachment_table(),
            Some(&[0x0a, 6, 0x08, 0, 0x12, 2, 0x08, 1][..])
        );
    }

    #[test]
    fn storage_report_accounts_for_every_nested_closure_message() {
        let source = storage_fixture();
        let (snapshot, report) =
            decode_slide_number_storage_with_report(&source, DecodeOptions::new(1024, 6, 62, 4))
                .unwrap();
        assert!(snapshot.attachment_table().is_some());
        assert_eq!(report.fields(), 6);
        assert_eq!(report.work_bytes(), 62);
        assert_eq!(report.max_depth(), 4);
    }

    #[test]
    fn nested_storage_field_limit_is_global_and_exact() {
        let error = decode_slide_number_storage_with_report(
            &storage_fixture(),
            DecodeOptions::new(1024, 5, 62, 4),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 6,
                maximum: 5
            })
        );
    }

    #[test]
    fn nested_storage_work_limit_is_global_and_exact() {
        let error = decode_slide_number_storage_with_report(
            &storage_fixture(),
            DecodeOptions::new(1024, 6, 61, 4),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Work {
                observed: 62,
                maximum: 61
            })
        );
    }

    #[test]
    fn nested_storage_depth_limit_is_global_and_exact() {
        let error = decode_slide_number_storage_with_report(
            &storage_fixture(),
            DecodeOptions::new(1024, 6, 62, 3),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 4,
                maximum: 3
            })
        );
    }

    fn storage_fixture() -> [u8; 15] {
        [
            0x1a, 3, 0xef, 0xbf, 0xbc, 0x4a, 8, 0x0a, 6, 0x08, 0, 0x12, 2, 0x08, 1,
        ]
    }

    #[test]
    fn byte_limit_has_an_exact_boundary() {
        let exact = DecodeOptions::new(3, 8, 6, 4);
        assert!(decode_slide_number_node(&[0x90, 1, 1], exact).is_ok());
        let short = DecodeOptions::new(2, 8, 6, 4);
        assert!(decode_slide_number_node(&[0x90, 1, 1], short).is_err());
    }

    #[test]
    fn malformed_selected_node_field_is_rejected_before_projection() {
        assert!(decode_slide_number_node(&[0x90, 1, 2], OPTIONS).is_err());
        assert!(decode_slide_number_node(&[0x90, 1, 1, 0x90, 1, 1], OPTIONS).is_err());
    }
}
