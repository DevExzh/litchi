//! Strict, source-preserving Buffa views for iWork annotation authors.
//!
//! The package owner supplies all author policy (including generated names,
//! public identifiers, and archive object identities).  This module owns only
//! the wire schema, finite resource accounting, and the small borrowed
//! author/storage values needed by the three format adapters.  The original
//! payload remains the preservation representation; storage rewrites copy it
//! and alter only one selected repeated reference field.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight, borrowed projection, and bounded writer are kept together."
)]

use core::{fmt, str};

use buffa::{DecodeOptions as BuffaDecodeOptions, ViewEncode as _};

use crate::buffa_annotation_author_generated::LitchiIwaAnnotationAuthorProjection as projection;

const AUTHOR_NAME_FIELD: u32 = 1;
const AUTHOR_COLOR_FIELD: u32 = 2;
const AUTHOR_PUBLIC_ID_FIELD: u32 = 3;
const AUTHOR_IS_PUBLIC_FIELD: u32 = 4;
const AUTHOR_PUBLIC_IDS_FIELD: u32 = 5;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_CYAN_FIELD: u32 = 7;
const COLOR_MAGENTA_FIELD: u32 = 8;
const COLOR_YELLOW_FIELD: u32 = 9;
const COLOR_BLACK_FIELD: u32 = 10;
const COLOR_WHITE_FIELD: u32 = 11;
const COLOR_RGBSPACE_FIELD: u32 = 12;
const STORAGE_AUTHOR_FIELD: u32 = 1;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;
const MAX_RECURSION_LIMIT: u32 = 64;
const DEFAULT_RECURSION_LIMIT: u32 = 16;
const MAX_DEFAULT_FIELDS: usize = 4096;
const MAX_DEFAULT_REFERENCES: usize = 1024;
const MAX_DEFAULT_TEXT_BYTES: usize = 256 * 1024;
const MAX_DEFAULT_WORK_BYTES: usize = 4 * 1024 * 1024;
const MAX_DEFAULT_OUTPUT_BYTES: usize = 512 * 1024;
const MAX_DEFAULT_ALLOCATIONS: usize = 4096;

/// Finite limits for one author or author-storage payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
    max_text_bytes: usize,
    max_allocations: usize,
}

impl DecodeOptions {
    /// Construct an explicit finite wire and allocation policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
        max_text_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
            max_text_bytes,
            max_allocations,
        }
    }

    /// Build a bounded profile from a known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes.saturating_add(64),
            bytes.saturating_mul(8).max(16),
            bytes.saturating_mul(64).max(128),
            DEFAULT_RECURSION_LIMIT,
            MAX_DEFAULT_REFERENCES,
            bytes.saturating_mul(8).max(64),
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    /// Return a conservative profile for an empty canonical payload.
    #[must_use]
    pub const fn default_profile() -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_FIELDS,
            MAX_DEFAULT_WORK_BYTES,
            DEFAULT_RECURSION_LIMIT,
            MAX_DEFAULT_REFERENCES,
            MAX_DEFAULT_TEXT_BYTES,
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    /// Replace the source/output message-byte ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Replace the aggregate field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the reference-count ceiling.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    /// Replace the text-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }

    /// Replace the logical allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            // Buffa accounts repeated borrowed strings and deferred message
            // fragments through this shared budget. Zero would reject every
            // valid repeated author ID/reference before the handwritten
            // preflight gets a chance to account the same work.
            .with_element_memory_limit(self.max_work_bytes)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact strict consumption of one author/storage decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    text_bytes: usize,
    allocations: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.source_bytes
    }

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

    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Typed bounded decode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    NonFiniteColor(&'static str),
    Invalid,
}

/// Finite resource category exceeded by strict decoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    References { observed: usize, maximum: usize },
    Text { observed: usize, maximum: usize },
    Allocations { observed: usize, maximum: usize },
    Scratch { observed: usize, maximum: usize },
    Retained { observed: usize, maximum: usize },
}

impl DecodeError {
    const fn invalid() -> Self {
        Self {
            kind: DecodeErrorKind::Invalid,
        }
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    const fn nonfinite_color(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonFiniteColor(field),
        }
    }

    /// Return the finite resource failure, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the missing required field, if any.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }

    /// Return the duplicated singular field, if any.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the canonical-wire reason, if any.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return the invalid UTF-8 field, if any.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::InvalidUtf8(field) => Some(field),
            _ => None,
        }
    }

    /// Return the non-finite color field, if any.
    #[must_use]
    pub const fn nonfinite_color_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonFiniteColor(field) => Some(field),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author output-byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author field limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author work limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::References { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author reference limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author text limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author allocation limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author scratch limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author retained-byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::NonFiniteColor(field) => write!(formatter, "{field} is non-finite"),
            DecodeErrorKind::Invalid => {
                formatter.write_str("invalid iWork annotation-author payload")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge => Self::resource(DecodeLimit::Bytes {
                observed: 0,
                maximum: 0,
            }),
            buffa::DecodeError::RecursionLimitExceeded => Self::resource(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION_LIMIT,
            }),
            other => Self {
                kind: DecodeErrorKind::Wire(other),
            },
        }
    }
}

/// Lossless scalar color facts from an annotation author.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColorSnapshot {
    model: i32,
    red: Option<f32>,
    green: Option<f32>,
    blue: Option<f32>,
    alpha: Option<f32>,
    cyan: Option<f32>,
    magenta: Option<f32>,
    yellow: Option<f32>,
    black: Option<f32>,
    white: Option<f32>,
    rgbspace: Option<i32>,
}

impl ColorSnapshot {
    #[must_use]
    pub const fn model(self) -> i32 {
        self.model
    }
    #[must_use]
    pub const fn red(self) -> Option<f32> {
        self.red
    }
    #[must_use]
    pub const fn green(self) -> Option<f32> {
        self.green
    }
    #[must_use]
    pub const fn blue(self) -> Option<f32> {
        self.blue
    }
    #[must_use]
    pub const fn alpha(self) -> Option<f32> {
        self.alpha
    }
    #[must_use]
    pub const fn cyan(self) -> Option<f32> {
        self.cyan
    }
    #[must_use]
    pub const fn magenta(self) -> Option<f32> {
        self.magenta
    }
    #[must_use]
    pub const fn yellow(self) -> Option<f32> {
        self.yellow
    }
    #[must_use]
    pub const fn black(self) -> Option<f32> {
        self.black
    }
    #[must_use]
    pub const fn white(self) -> Option<f32> {
        self.white
    }
    #[must_use]
    pub const fn rgbspace(self) -> Option<i32> {
        self.rgbspace
    }
}

/// Borrowed scalar facts from one `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AuthorReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl AuthorReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Concise neutral spelling for one author-storage reference.
pub type AuthorReference = AuthorReferenceSnapshot;

/// Borrowed annotation-author facts. Strings point into the caller payload;
/// repeated public IDs are retained as borrowed references only.
#[derive(Debug, Clone, PartialEq)]
pub struct AnnotationAuthorSnapshot<'source> {
    name: Option<&'source str>,
    color: Option<ColorSnapshot>,
    public_id: Option<&'source str>,
    is_public_author: Option<bool>,
    public_ids: Vec<&'source str>,
}

impl<'source> AnnotationAuthorSnapshot<'source> {
    #[must_use]
    pub const fn name(&self) -> Option<&'source str> {
        self.name
    }

    #[must_use]
    pub const fn color(&self) -> Option<ColorSnapshot> {
        self.color
    }

    #[must_use]
    pub const fn public_id(&self) -> Option<&'source str> {
        self.public_id
    }

    #[must_use]
    pub const fn is_public_author(&self) -> Option<bool> {
        self.is_public_author
    }

    /// Iterate source-ordered repeated `public_ids` without cloning strings.
    pub fn public_ids(&self) -> impl Iterator<Item = &'source str> + '_ {
        self.public_ids.iter().copied()
    }

    #[must_use]
    pub const fn public_id_count(&self) -> usize {
        self.public_ids.len()
    }
}

/// Concise neutral spelling for an annotation-author snapshot.
pub type AuthorSnapshot<'source> = AnnotationAuthorSnapshot<'source>;

/// Borrowed author-storage facts in source order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnnotationAuthorStorageSnapshot {
    author_refs: Vec<AuthorReferenceSnapshot>,
}

impl AnnotationAuthorStorageSnapshot {
    /// Iterate source-ordered `annotation_author` references.
    pub fn author_refs(&self) -> impl Iterator<Item = AuthorReferenceSnapshot> + '_ {
        self.author_refs.iter().copied()
    }

    /// Compatibility alias for callers that call the field `references`.
    pub fn references(&self) -> impl Iterator<Item = AuthorReferenceSnapshot> + '_ {
        self.author_refs()
    }

    /// Alias emphasizing the native field name.
    pub fn author_references(&self) -> impl Iterator<Item = AuthorReferenceSnapshot> + '_ {
        self.author_refs()
    }

    #[must_use]
    pub const fn author_ref_count(&self) -> usize {
        self.author_refs.len()
    }
}

/// Concise neutral spelling for author storage.
pub type AuthorStorageSnapshot = AnnotationAuthorStorageSnapshot;

/// Decode one annotation-author payload.
pub fn decode_annotation_author(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AnnotationAuthorSnapshot<'_>, DecodeError> {
    Ok(decode_annotation_author_with_report(source, options)?.0)
}

/// Decode one annotation-author payload and return strict resource usage.
pub fn decode_annotation_author_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(AnnotationAuthorSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let (snapshot, raw_color) = decode_author_strict(source, &mut budget, 1)?;
    parity_author(source, &snapshot, raw_color, &mut budget)?;
    Ok((snapshot, budget.report()))
}

/// Decode one annotation-author storage payload.
pub fn decode_annotation_author_storage(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AnnotationAuthorStorageSnapshot, DecodeError> {
    Ok(decode_annotation_author_storage_with_report(source, options)?.0)
}

/// Decode author storage and return strict resource usage.
pub fn decode_annotation_author_storage_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(AnnotationAuthorStorageSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_storage_strict(source, &mut budget, 1)?.0;
    parity_storage(source, &snapshot, &mut budget)?;
    Ok((snapshot, budget.report()))
}

/// Compatibility spelling for archive-oriented package owners.
pub fn decode_annotation_author_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AnnotationAuthorSnapshot<'_>, DecodeError> {
    decode_annotation_author(source, options)
}

/// Compatibility spelling for archive-oriented package owners.
pub fn decode_annotation_author_storage_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AnnotationAuthorStorageSnapshot, DecodeError> {
    decode_annotation_author_storage(source, options)
}

/// Concise neutral spelling for an author payload reader.
pub fn decode_author(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AuthorSnapshot<'_>, DecodeError> {
    decode_annotation_author(source, options)
}

/// Concise neutral spelling for an author-storage reader.
pub fn decode_author_storage(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AuthorStorageSnapshot, DecodeError> {
    decode_annotation_author_storage(source, options)
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct ColorValues {
    model: i32,
    red: Option<f32>,
    green: Option<f32>,
    blue: Option<f32>,
    alpha: Option<f32>,
    cyan: Option<f32>,
    magenta: Option<f32>,
    yellow: Option<f32>,
    black: Option<f32>,
    white: Option<f32>,
    rgbspace: Option<i32>,
}

#[derive(Clone, Copy, Debug)]
struct Field<'source> {
    number: u32,
    wire_type: u8,
    value: Value<'source>,
    key_canonical: bool,
    value_canonical: bool,
    length_canonical: bool,
    start: usize,
    end: usize,
}

#[derive(Clone, Copy, Debug)]
enum Value<'source> {
    Varint(u64),
    Fixed64,
    Bytes(&'source [u8]),
    Group,
    Fixed32(u32),
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

impl<'source> Field<'source> {
    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            Value::Bytes(value) if self.wire_type == 2 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Varint(value) if self.wire_type == 0 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn fixed32(self) -> Result<u32, DecodeError> {
        match self.value {
            Value::Fixed32(value) if self.wire_type == 5 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }
}

fn decode_author_strict<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(AnnotationAuthorSnapshot<'source>, Option<&'source [u8]>), DecodeError> {
    budget.message(source, depth)?;
    let fields = parse_fields(source, budget, depth)?;
    let mut name = None;
    let mut color = None;
    let mut raw_color = None;
    let mut public_id = None;
    let mut is_public_author = None;
    let mut public_ids = Vec::new();
    for field in fields {
        validate_wire_field(field)?;
        match field.number {
            AUTHOR_NAME_FIELD => {
                require_wire(field, 2, "TSK.AnnotationAuthorArchive.name")?;
                if name.is_some() {
                    return Err(DecodeError::duplicate("TSK.AnnotationAuthorArchive.name"));
                }
                name = Some(strict_string(
                    field.bytes()?,
                    budget,
                    "TSK.AnnotationAuthorArchive.name",
                )?);
            },
            AUTHOR_COLOR_FIELD => {
                require_wire(field, 2, "TSK.AnnotationAuthorArchive.color")?;
                if color.is_some() {
                    return Err(DecodeError::duplicate("TSK.AnnotationAuthorArchive.color"));
                }
                let raw = field.bytes()?;
                let value = decode_color_strict(
                    raw,
                    budget,
                    depth.checked_add(1).ok_or_else(DecodeError::invalid)?,
                )?;
                raw_color = Some(raw);
                color = Some(value);
            },
            AUTHOR_PUBLIC_ID_FIELD => {
                require_wire(field, 2, "TSK.AnnotationAuthorArchive.public_id")?;
                if public_id.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSK.AnnotationAuthorArchive.public_id",
                    ));
                }
                public_id = Some(strict_string(
                    field.bytes()?,
                    budget,
                    "TSK.AnnotationAuthorArchive.public_id",
                )?);
            },
            AUTHOR_IS_PUBLIC_FIELD => {
                require_wire(field, 0, "TSK.AnnotationAuthorArchive.is_public_author")?;
                if is_public_author.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSK.AnnotationAuthorArchive.is_public_author",
                    ));
                }
                is_public_author = Some(canonical_bool(field.varint()?)?);
            },
            AUTHOR_PUBLIC_IDS_FIELD => {
                require_wire(field, 2, "TSK.AnnotationAuthorArchive.public_ids")?;
                let value = strict_string(
                    field.bytes()?,
                    budget,
                    "TSK.AnnotationAuthorArchive.public_ids",
                )?;
                reserve_push(&mut public_ids, value, budget)?;
            },
            _ => {},
        }
    }
    Ok((
        AnnotationAuthorSnapshot {
            name,
            color: color.map(color_snapshot),
            public_id,
            is_public_author,
            public_ids,
        },
        raw_color,
    ))
}

fn decode_color_strict(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ColorValues, DecodeError> {
    budget.message(source, depth)?;
    let fields = parse_fields(source, budget, depth)?;
    let mut model = None;
    let mut red = None;
    let mut green = None;
    let mut blue = None;
    let mut alpha = None;
    let mut cyan = None;
    let mut magenta = None;
    let mut yellow = None;
    let mut black = None;
    let mut white = None;
    let mut rgbspace = None;
    for field in fields {
        validate_wire_field(field)?;
        match field.number {
            COLOR_MODEL_FIELD => {
                require_wire(field, 0, "TSP.Color.model")?;
                if model.replace(canonical_int32(field.varint()?)?).is_some() {
                    return Err(DecodeError::duplicate("TSP.Color.model"));
                }
            },
            COLOR_RED_FIELD => assign_color(&mut red, field, "TSP.Color.r")?,
            COLOR_GREEN_FIELD => assign_color(&mut green, field, "TSP.Color.g")?,
            COLOR_BLUE_FIELD => assign_color(&mut blue, field, "TSP.Color.b")?,
            COLOR_ALPHA_FIELD => assign_color(&mut alpha, field, "TSP.Color.a")?,
            COLOR_CYAN_FIELD => assign_color(&mut cyan, field, "TSP.Color.c")?,
            COLOR_MAGENTA_FIELD => assign_color(&mut magenta, field, "TSP.Color.m")?,
            COLOR_YELLOW_FIELD => assign_color(&mut yellow, field, "TSP.Color.y")?,
            COLOR_BLACK_FIELD => assign_color(&mut black, field, "TSP.Color.k")?,
            COLOR_WHITE_FIELD => assign_color(&mut white, field, "TSP.Color.w")?,
            COLOR_RGBSPACE_FIELD => {
                require_wire(field, 0, "TSP.Color.rgbspace")?;
                if rgbspace
                    .replace(canonical_int32(field.varint()?)?)
                    .is_some()
                {
                    return Err(DecodeError::duplicate("TSP.Color.rgbspace"));
                }
            },
            _ => {},
        }
    }
    Ok(ColorValues {
        model: model.ok_or_else(|| DecodeError::missing("TSP.Color.model"))?,
        red,
        green,
        blue,
        alpha,
        cyan,
        magenta,
        yellow,
        black,
        white,
        rgbspace,
    })
}

fn assign_color(
    destination: &mut Option<f32>,
    field: Field<'_>,
    name: &'static str,
) -> Result<(), DecodeError> {
    require_wire(field, 5, name)?;
    if destination.is_some() {
        return Err(DecodeError::duplicate(name));
    }
    let value = f32::from_bits(field.fixed32()?);
    if !value.is_finite() {
        return Err(DecodeError::nonfinite_color(name));
    }
    *destination = Some(value);
    Ok(())
}

fn decode_storage_strict<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(AnnotationAuthorStorageSnapshot, Vec<Field<'source>>), DecodeError> {
    budget.message(source, depth)?;
    let fields = parse_fields(source, budget, depth)?;
    let mut author_refs = Vec::new();
    for field in &fields {
        validate_wire_field(*field)?;
        if field.number != STORAGE_AUTHOR_FIELD {
            continue;
        }
        require_wire(
            *field,
            2,
            "TSK.AnnotationAuthorStorageArchive.annotation_author",
        )?;
        let raw = field.bytes()?;
        budget.reference(raw.len())?;
        let reference = decode_reference_strict(
            raw,
            budget,
            depth.checked_add(1).ok_or_else(DecodeError::invalid)?,
        )?;
        reserve_push(&mut author_refs, reference, budget)?;
    }
    Ok((AnnotationAuthorStorageSnapshot { author_refs }, fields))
}

fn decode_reference_strict(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<AuthorReferenceSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let fields = parse_fields(source, budget, depth)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    for field in fields {
        validate_wire_field(field)?;
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                require_wire(field, 0, "TSP.Reference.identifier")?;
                if identifier.replace(field.varint()?).is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                require_wire(field, 0, "TSP.Reference.deprecated_type")?;
                if deprecated_type
                    .replace(canonical_int32(field.varint()?)?)
                    .is_some()
                {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                require_wire(field, 0, "TSP.Reference.deprecated_is_external")?;
                if deprecated_is_external
                    .replace(canonical_bool(field.varint()?)?)
                    .is_some()
                {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
            },
            _ => {},
        }
    }
    Ok(AuthorReferenceSnapshot {
        identifier: identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn parity_author(
    source: &[u8],
    strict: &AnnotationAuthorSnapshot<'_>,
    raw_color: Option<&[u8]>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.message(source, 1)?;
    let view: projection::AnnotationAuthorArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let color = match (view.color.get().map_err(DecodeError::from)?, raw_color) {
        (Some(view), Some(raw)) => {
            budget.message(raw, 2)?;
            Some(parity_color(
                &view,
                strict.color.ok_or_else(DecodeError::invalid)?,
            )?)
        },
        (None, None) => None,
        _ => return Err(DecodeError::invalid()),
    };
    if view.name != strict.name
        || view.public_id != strict.public_id
        || view.is_public_author != strict.is_public_author
        || view
            .public_ids
            .iter()
            .copied()
            .ne(strict.public_ids.iter().copied())
        || color != strict.color
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parity_color(
    view: &projection::ColorLazyView<'_>,
    strict: ColorSnapshot,
) -> Result<ColorSnapshot, DecodeError> {
    if !view.has_model() {
        return Err(DecodeError::missing("TSP.Color.model"));
    }
    let projected = ColorSnapshot {
        model: view.model,
        red: view.r,
        green: view.g,
        blue: view.b,
        alpha: view.a,
        cyan: view.c,
        magenta: view.m,
        yellow: view.y,
        black: view.k,
        white: view.w,
        rgbspace: view.rgbspace,
    };
    if projected != strict {
        return Err(DecodeError::invalid());
    }
    Ok(projected)
}

fn parity_storage(
    source: &[u8],
    strict: &AnnotationAuthorStorageSnapshot,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.message(source, 1)?;
    let view: projection::AnnotationAuthorStorageArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let mut strict_iter = strict.author_refs.iter();
    for item in &view.annotation_author {
        let reference = item.map_err(DecodeError::from)?;
        let actual = force_reference(&reference)?;
        let expected = strict_iter.next().ok_or_else(DecodeError::invalid)?;
        if actual != *expected {
            return Err(DecodeError::invalid());
        }
    }
    if strict_iter.next().is_some() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn force_reference(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<AuthorReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing("TSP.Reference.identifier"));
    }
    Ok(AuthorReferenceSnapshot {
        identifier: view.identifier,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
    })
}

fn color_snapshot(value: ColorValues) -> ColorSnapshot {
    ColorSnapshot {
        model: value.model,
        red: value.red,
        green: value.green,
        blue: value.blue,
        alpha: value.alpha,
        cyan: value.cyan,
        magenta: value.magenta,
        yellow: value.yellow,
        black: value.black,
        white: value.white,
        rgbspace: value.rgbspace,
    }
}

fn strict_string<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    field: &'static str,
) -> Result<&'source str, DecodeError> {
    let value = str::from_utf8(source).map_err(|_error| DecodeError::utf8(field))?;
    budget.text(source.len())?;
    Ok(value)
}

fn reserve_push<T>(values: &mut Vec<T>, value: T, budget: &mut Budget) -> Result<(), DecodeError> {
    if values.try_reserve(1).is_err() {
        return Err(DecodeError::resource(DecodeLimit::Allocations {
            observed: values.len().saturating_add(1),
            maximum: budget.options.max_allocations,
        }));
    }
    budget.allocation()?;
    values.push(value);
    Ok(())
}

fn validate_wire_field(field: Field<'_>) -> Result<(), DecodeError> {
    if !field.key_canonical {
        return Err(DecodeError::noncanonical("field key"));
    }
    match field.wire_type {
        0 if !field.value_canonical => Err(DecodeError::noncanonical("varint value")),
        2 if !field.length_canonical => Err(DecodeError::noncanonical("length value")),
        0 | 1 | 2 | 3 | 5 => Ok(()),
        _ => Err(buffa::DecodeError::InvalidWireType(u32::from(field.wire_type)).into()),
    }
}

fn require_wire(field: Field<'_>, wire_type: u8, name: &'static str) -> Result<(), DecodeError> {
    if field.wire_type != wire_type {
        return Err(DecodeError::invalid());
    }
    if !field.key_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    if wire_type == 0 && !field.value_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    if wire_type == 2 && !field.length_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    Ok(())
}

fn parse_fields<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<Field<'source>>, DecodeError> {
    let mut remaining = source;
    let mut fields = Vec::new();
    while !remaining.is_empty() {
        let before = remaining.len();
        let item = parse_field(&mut remaining, budget, depth)?;
        let end = source.len().saturating_sub(remaining.len());
        let start = source.len().saturating_sub(before);
        match item {
            Some(ParseItem::Field(mut field)) => {
                field.start = start;
                field.end = end;
                reserve_push(&mut fields, field, budget)?;
            },
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => break,
        }
    }
    Ok(fields)
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.field()?;
    budget.observe_depth(depth)?;
    let (tag, key_canonical) = take_varint(source)?;
    let number =
        u32::try_from(tag >> 3).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire_type = u8::try_from(tag & 7).map_err(|_error| DecodeError::invalid())?;
    let mut value_canonical = true;
    let mut length_canonical = true;
    let value = match wire_type {
        0 => {
            let (value, canonical) = take_varint(source)?;
            value_canonical = canonical;
            Value::Varint(value)
        },
        1 => {
            take_exact(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let (length, canonical) = take_varint(source)?;
            length_canonical = canonical;
            let length =
                usize::try_from(length).map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            Value::Bytes(take_exact(source, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, budget, child_depth)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => Value::Fixed32(u32::from_le_bytes(
            take_exact(source, 4)?
                .try_into()
                .map_err(|_error| DecodeError::invalid())?,
        )),
        _ => return Err(buffa::DecodeError::InvalidWireType(u32::from(wire_type)).into()),
    };
    Ok(Some(ParseItem::Field(Field {
        number,
        wire_type,
        value,
        key_canonical,
        value_canonical,
        length_canonical,
        start: 0,
        end: 0,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    amount: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(amount);
    *source = remaining;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index.checked_add(1).ok_or_else(DecodeError::invalid)?;
            *source = &original[consumed..];
            return Ok((value, encoded_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

const fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(value) = i32::try_from(value) {
        return Ok(value);
    }
    if value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical("int32 scalar is out of range"));
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
        .map_err(|_error| DecodeError::noncanonical("int32 scalar is out of range"))
}

struct Budget {
    options: DecodeOptions,
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    text_bytes: usize,
    allocations: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes =
            usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes || source.len() > options.max_message_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len().max(options.max_message_bytes),
                maximum: hard_bytes.min(options.max_message_bytes),
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }));
        }
        Ok(Self {
            options,
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            text_bytes: 0,
            allocations: 0,
        })
    }

    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        if source.len() > self.options.max_message_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: self.options.max_message_bytes,
            }));
        }
        self.observe_depth(depth)?;
        self.work(source.len())
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn reference(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .references
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_references {
            return Err(DecodeError::resource(DecodeLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references = observed;
        self.work(bytes)
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        // Text is charged as work as well as a separate logical ceiling so a
        // caller cannot hide a large borrowed string behind a small field.
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_text_bytes {
            return Err(DecodeError::resource(DecodeLimit::Text {
                observed,
                maximum: self.options.max_text_bytes,
            }));
        }
        self.text_bytes = observed;
        self.work(bytes)
    }

    fn allocation(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .allocations
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_allocations {
            return Err(DecodeError::resource(DecodeLimit::Allocations {
                observed,
                maximum: self.options.max_allocations,
            }));
        }
        self.allocations = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            text_bytes: self.text_bytes,
            allocations: self.allocations,
        }
    }
}

/// A color value supplied by a package owner for a generated author.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AuthorColorWrite {
    model: i32,
    red: Option<f32>,
    green: Option<f32>,
    blue: Option<f32>,
    alpha: Option<f32>,
    cyan: Option<f32>,
    magenta: Option<f32>,
    yellow: Option<f32>,
    black: Option<f32>,
    white: Option<f32>,
    rgbspace: Option<i32>,
}

/// Concise neutral spelling for a color supplied by the package owner.
pub type ColorWrite = AuthorColorWrite;

impl AuthorColorWrite {
    /// Construct a color from all native optional scalar presences.
    #[must_use]
    #[allow(
        clippy::too_many_arguments,
        reason = "The fields mirror the native color schema."
    )]
    pub const fn new(
        model: i32,
        red: Option<f32>,
        green: Option<f32>,
        blue: Option<f32>,
        alpha: Option<f32>,
        cyan: Option<f32>,
        magenta: Option<f32>,
        yellow: Option<f32>,
        black: Option<f32>,
        white: Option<f32>,
        rgbspace: Option<i32>,
    ) -> Self {
        Self {
            model,
            red,
            green,
            blue,
            alpha,
            cyan,
            magenta,
            yellow,
            black,
            white,
            rgbspace,
        }
    }

    /// Copy a decoded color into a write value without changing presence.
    #[must_use]
    pub const fn from_snapshot(color: ColorSnapshot) -> Self {
        Self::new(
            color.model,
            color.red,
            color.green,
            color.blue,
            color.alpha,
            color.cyan,
            color.magenta,
            color.yellow,
            color.black,
            color.white,
            color.rgbspace,
        )
    }
}

/// Bounded author values supplied by a package owner.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AnnotationAuthorWrite<'source> {
    name: Option<&'source str>,
    color: Option<AuthorColorWrite>,
    public_id: Option<&'source str>,
    is_public_author: Option<bool>,
    public_ids: &'source [&'source str],
}

/// Concise neutral spelling for a package-owned author write.
pub type AuthorWrite<'source> = AnnotationAuthorWrite<'source>;

impl<'source> AnnotationAuthorWrite<'source> {
    /// Build an author write while preserving every optional native field.
    #[must_use]
    pub const fn new(
        name: Option<&'source str>,
        color: Option<AuthorColorWrite>,
        public_id: Option<&'source str>,
        is_public_author: Option<bool>,
        public_ids: &'source [&'source str],
    ) -> Self {
        Self {
            name,
            color,
            public_id,
            is_public_author,
            public_ids,
        }
    }
}

/// Canonical author-storage values supplied by a package owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnnotationAuthorStorageWrite<'source> {
    author_identifiers: &'source [u64],
}

/// Concise neutral spelling for a package-owned storage write.
pub type AuthorStorageWrite<'source> = AnnotationAuthorStorageWrite<'source>;

impl<'source> AnnotationAuthorStorageWrite<'source> {
    #[must_use]
    pub const fn new(author_identifiers: &'source [u64]) -> Self {
        Self { author_identifiers }
    }
}

/// Finite output policy for Buffa-generated author/storage writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_references: usize,
    max_text_bytes: usize,
    max_allocations: usize,
}

impl EncodeOptions {
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_references: usize,
        max_text_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_references,
            max_text_bytes,
            max_allocations,
        }
    }

    #[must_use]
    pub const fn for_author(_write: &AnnotationAuthorWrite<'_>) -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_FIELDS,
            MAX_DEFAULT_WORK_BYTES,
            0,
            MAX_DEFAULT_TEXT_BYTES,
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    #[must_use]
    pub const fn for_storage(write: &AnnotationAuthorStorageWrite<'_>) -> Self {
        Self::new(
            MAX_DEFAULT_OUTPUT_BYTES,
            MAX_DEFAULT_FIELDS,
            MAX_DEFAULT_WORK_BYTES,
            write.author_identifiers.len().saturating_add(1),
            0,
            MAX_DEFAULT_ALLOCATIONS,
        )
    }

    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }
}

/// Exact output accounting for a canonical author/storage write.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    references: usize,
    text_bytes: usize,
    allocations: usize,
}

impl EncodeReport {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Typed bounded encode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncodeError {
    kind: EncodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum EncodeErrorKind {
    Wire(buffa::EncodeError),
    Resource(EncodeLimit),
    InvalidInput(&'static str),
    Allocation { amount: usize },
    Verification,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    References { observed: usize, maximum: usize },
    Text { observed: usize, maximum: usize },
    Allocations { observed: usize, maximum: usize },
}

impl EncodeError {
    const fn invalid(field: &'static str) -> Self {
        Self {
            kind: EncodeErrorKind::InvalidInput(field),
        }
    }

    const fn resource(limit: EncodeLimit) -> Self {
        Self {
            kind: EncodeErrorKind::Resource(limit),
        }
    }
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            EncodeErrorKind::Wire(error) => error.fmt(formatter),
            EncodeErrorKind::Resource(EncodeLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author output limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::Resource(EncodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author field limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::Resource(EncodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author work limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::Resource(EncodeLimit::References { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author reference limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::Resource(EncodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author text limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::Resource(EncodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "iWork annotation-author allocation limit exceeded: observed {observed}, maximum {maximum}"
            ),
            EncodeErrorKind::InvalidInput(field) => {
                write!(formatter, "invalid iWork annotation-author input: {field}")
            },
            EncodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate iWork annotation-author output for {amount} bytes"
            ),
            EncodeErrorKind::Verification => {
                formatter.write_str("iWork annotation-author Buffa output verification failed")
            },
        }
    }
}

impl std::error::Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self {
            kind: EncodeErrorKind::Wire(error),
        }
    }
}

/// Encode a canonical author payload through a private Buffa view.
pub fn encode_annotation_author(
    write: &AnnotationAuthorWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    Ok(encode_annotation_author_with_report(write, options)?.0)
}

/// Compatibility spelling for package-owned generated authors.
pub fn canonical_annotation_author(
    write: &AnnotationAuthorWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    encode_annotation_author(write, options)
}

/// Encode a canonical author and return exact output accounting.
pub fn encode_annotation_author_with_report(
    write: &AnnotationAuthorWrite<'_>,
    options: EncodeOptions,
) -> Result<(Vec<u8>, EncodeReport), EncodeError> {
    validate_author_write(write)?;
    let report = author_report(write)?;
    preflight_encode(report, options)?;
    let view = author_view(write);
    let measured = usize::try_from(view.try_encoded_len()?).map_err(|_error| EncodeError {
        kind: EncodeErrorKind::Verification,
    })?;
    if measured != report.output_bytes {
        return Err(EncodeError {
            kind: EncodeErrorKind::Verification,
        });
    }
    encode_view(view, report.output_bytes, options.max_output_bytes).map(|bytes| (bytes, report))
}

/// Encode canonical author storage through a private Buffa view.
pub fn encode_annotation_author_storage(
    write: &AnnotationAuthorStorageWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    Ok(encode_annotation_author_storage_with_report(write, options)?.0)
}

/// Compatibility spelling for a canonical storage payload.
pub fn canonical_annotation_author_storage(
    write: &AnnotationAuthorStorageWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    encode_annotation_author_storage(write, options)
}

/// Concise neutral spelling for a package-owned canonical author.
pub fn canonical_author(
    write: &AuthorWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    encode_annotation_author(write, options)
}

/// Concise neutral spelling for a package-owned canonical author storage.
pub fn canonical_author_storage(
    write: &AuthorStorageWrite<'_>,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    encode_annotation_author_storage(write, options)
}

/// Encode canonical author storage and return exact output accounting.
pub fn encode_annotation_author_storage_with_report(
    write: &AnnotationAuthorStorageWrite<'_>,
    options: EncodeOptions,
) -> Result<(Vec<u8>, EncodeReport), EncodeError> {
    validate_storage_write(write)?;
    let report = storage_report(write)?;
    preflight_encode(report, options)?;
    let view = storage_view(write);
    let measured = usize::try_from(view.try_encoded_len()?).map_err(|_error| EncodeError {
        kind: EncodeErrorKind::Verification,
    })?;
    if measured != report.output_bytes {
        return Err(EncodeError {
            kind: EncodeErrorKind::Verification,
        });
    }
    encode_view(view, report.output_bytes, options.max_output_bytes).map(|bytes| (bytes, report))
}

fn encode_view<'source, V>(
    view: V,
    output_bytes: usize,
    maximum: usize,
) -> Result<Vec<u8>, EncodeError>
where
    V: buffa::ViewEncode<'source>,
{
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_bytes)
        .map_err(|_| EncodeErrorKind::Allocation {
            amount: output_bytes,
        })
        .map_err(|kind| EncodeError { kind })?;
    let max = u32::try_from(maximum).unwrap_or(u32::MAX);
    let encoded = view.try_encode_bounded(max, &mut output)?;
    if usize::try_from(encoded).ok() != Some(output_bytes) || output.len() != output_bytes {
        return Err(EncodeError {
            kind: EncodeErrorKind::Verification,
        });
    }
    Ok(output)
}

fn author_view<'source>(
    write: &AnnotationAuthorWrite<'source>,
) -> projection::AnnotationAuthorArchiveView<'source> {
    projection::AnnotationAuthorArchiveView {
        name: write.name,
        color: write
            .color
            .map(color_view)
            .map(buffa::MessageFieldView::set)
            .unwrap_or_else(buffa::MessageFieldView::unset),
        public_id: write.public_id,
        is_public_author: write.is_public_author,
        public_ids: write.public_ids.iter().copied().collect(),
    }
}

fn color_view<'source>(write: AuthorColorWrite) -> projection::ColorView<'source> {
    projection::ColorView {
        model: write.model,
        r: write.red,
        g: write.green,
        b: write.blue,
        a: write.alpha,
        c: write.cyan,
        m: write.magenta,
        y: write.yellow,
        k: write.black,
        w: write.white,
        rgbspace: write.rgbspace,
        ..Default::default()
    }
}

fn storage_view<'source>(
    write: &AnnotationAuthorStorageWrite<'source>,
) -> projection::AnnotationAuthorStorageArchiveView<'source> {
    projection::AnnotationAuthorStorageArchiveView {
        annotation_author: write
            .author_identifiers
            .iter()
            .copied()
            .map(reference_view)
            .collect(),
    }
}

fn reference_view<'source>(identifier: u64) -> projection::ReferenceView<'source> {
    projection::ReferenceView {
        identifier,
        ..Default::default()
    }
}

fn validate_author_write(write: &AnnotationAuthorWrite<'_>) -> Result<(), EncodeError> {
    for value in [write.name, write.public_id]
        .into_iter()
        .flatten()
        .chain(write.public_ids.iter().copied())
    {
        if value.len() > MAX_DEFAULT_TEXT_BYTES {
            return Err(EncodeError::invalid("text exceeds default bound"));
        }
    }
    if let Some(color) = write.color {
        for value in [
            color.red,
            color.green,
            color.blue,
            color.alpha,
            color.cyan,
            color.magenta,
            color.yellow,
            color.black,
            color.white,
        ]
        .into_iter()
        .flatten()
        {
            if !value.is_finite() {
                return Err(EncodeError::invalid("color channel is non-finite"));
            }
        }
    }
    Ok(())
}

fn validate_storage_write(write: &AnnotationAuthorStorageWrite<'_>) -> Result<(), EncodeError> {
    if write.author_identifiers.contains(&0) {
        return Err(EncodeError::invalid("author identifier is zero"));
    }
    Ok(())
}

fn author_report(write: &AnnotationAuthorWrite<'_>) -> Result<EncodeReport, EncodeError> {
    let color_bytes = write
        .color
        .map(color_encoded_len)
        .unwrap_or(Some(0))
        .ok_or_else(|| EncodeError::invalid("author color length overflow"))?;
    let mut output_bytes = 0usize;
    let mut fields = 0usize;
    let mut text_bytes = 0usize;
    for (field, value) in [
        (AUTHOR_NAME_FIELD, write.name),
        (AUTHOR_PUBLIC_ID_FIELD, write.public_id),
    ] {
        if let Some(value) = value {
            output_bytes = output_bytes
                .checked_add(
                    length_field_len(field, value.len())
                        .ok_or_else(|| EncodeError::invalid("author output length overflow"))?,
                )
                .ok_or_else(|| EncodeError::invalid("author output length overflow"))?;
            fields += 1;
            text_bytes = text_bytes
                .checked_add(value.len())
                .ok_or_else(|| EncodeError::invalid("author text length overflow"))?;
        }
    }
    if let Some(color) = write.color {
        output_bytes = output_bytes
            .checked_add(
                length_field_len(AUTHOR_COLOR_FIELD, color_bytes)
                    .ok_or_else(|| EncodeError::invalid("author color length overflow"))?,
            )
            .ok_or_else(|| EncodeError::invalid("author output length overflow"))?;
        fields += 1;
        fields = fields
            .checked_add(color_fields(color))
            .ok_or_else(|| EncodeError::invalid("author field count overflow"))?;
    }
    if let Some(value) = write.is_public_author {
        output_bytes = output_bytes
            .checked_add(varint_field_len(AUTHOR_IS_PUBLIC_FIELD, u64::from(value)))
            .ok_or_else(|| EncodeError::invalid("author output length overflow"))?;
        fields += 1;
    }
    for value in write.public_ids.iter() {
        output_bytes = output_bytes
            .checked_add(
                length_field_len(AUTHOR_PUBLIC_IDS_FIELD, value.len())
                    .ok_or_else(|| EncodeError::invalid("author output length overflow"))?,
            )
            .ok_or_else(|| EncodeError::invalid("author output length overflow"))?;
        fields += 1;
        text_bytes = text_bytes
            .checked_add(value.len())
            .ok_or_else(|| EncodeError::invalid("author text length overflow"))?;
    }
    let work_bytes = output_bytes
        .checked_mul(4)
        .ok_or_else(|| EncodeError::invalid("author work length overflow"))?
        .max(1);
    let allocations = usize::from(output_bytes != 0)
        .checked_add(usize::from(write.color.is_some()))
        .and_then(|value| value.checked_add(usize::from(!write.public_ids.is_empty())))
        .ok_or_else(|| EncodeError::invalid("author allocation count overflow"))?;
    Ok(EncodeReport {
        output_bytes,
        fields,
        work_bytes,
        references: 0,
        text_bytes,
        allocations,
    })
}

fn color_fields(color: AuthorColorWrite) -> usize {
    1 + [
        color.red,
        color.green,
        color.blue,
        color.alpha,
        color.cyan,
        color.magenta,
        color.yellow,
        color.black,
        color.white,
    ]
    .into_iter()
    .filter(Option::is_some)
    .count()
        + usize::from(color.rgbspace.is_some())
}

fn color_encoded_len(color: AuthorColorWrite) -> Option<usize> {
    let mut length = varint_field_len(COLOR_MODEL_FIELD, encode_int32(color.model));
    for (field, value) in [
        (COLOR_RED_FIELD, color.red),
        (COLOR_GREEN_FIELD, color.green),
        (COLOR_BLUE_FIELD, color.blue),
        (COLOR_ALPHA_FIELD, color.alpha),
        (COLOR_CYAN_FIELD, color.cyan),
        (COLOR_MAGENTA_FIELD, color.magenta),
        (COLOR_YELLOW_FIELD, color.yellow),
        (COLOR_BLACK_FIELD, color.black),
        (COLOR_WHITE_FIELD, color.white),
    ] {
        if value.is_some() {
            length = length.checked_add(fixed32_field_len(field))?;
        }
    }
    if let Some(value) = color.rgbspace {
        length = length.checked_add(varint_field_len(COLOR_RGBSPACE_FIELD, encode_int32(value)))?;
    }
    Some(length)
}

fn storage_report(write: &AnnotationAuthorStorageWrite<'_>) -> Result<EncodeReport, EncodeError> {
    let mut output_bytes = 0usize;
    for identifier in write.author_identifiers.iter().copied() {
        let reference = varint_field_len(REFERENCE_IDENTIFIER_FIELD, identifier);
        output_bytes = output_bytes
            .checked_add(
                length_field_len(STORAGE_AUTHOR_FIELD, reference)
                    .ok_or_else(|| EncodeError::invalid("storage output length overflow"))?,
            )
            .ok_or_else(|| EncodeError::invalid("storage output length overflow"))?;
    }
    let references = write.author_identifiers.len();
    let fields = references
        .checked_mul(2)
        .ok_or_else(|| EncodeError::invalid("storage field count overflow"))?;
    let work_bytes = output_bytes
        .checked_mul(4)
        .ok_or_else(|| EncodeError::invalid("storage work length overflow"))?
        .max(1);
    let allocations = usize::from(output_bytes != 0)
        .checked_add(usize::from(!write.author_identifiers.is_empty()))
        .ok_or_else(|| EncodeError::invalid("storage allocation count overflow"))?;
    Ok(EncodeReport {
        output_bytes,
        fields,
        work_bytes,
        references,
        text_bytes: 0,
        allocations,
    })
}

fn preflight_encode(report: EncodeReport, options: EncodeOptions) -> Result<(), EncodeError> {
    let checks = [
        (report.output_bytes > options.max_output_bytes).then_some(EncodeLimit::OutputBytes {
            observed: report.output_bytes,
            maximum: options.max_output_bytes,
        }),
        (report.fields > options.max_fields).then_some(EncodeLimit::Fields {
            observed: report.fields,
            maximum: options.max_fields,
        }),
        (report.work_bytes > options.max_work_bytes).then_some(EncodeLimit::Work {
            observed: report.work_bytes,
            maximum: options.max_work_bytes,
        }),
        (report.references > options.max_references).then_some(EncodeLimit::References {
            observed: report.references,
            maximum: options.max_references,
        }),
        (report.text_bytes > options.max_text_bytes).then_some(EncodeLimit::Text {
            observed: report.text_bytes,
            maximum: options.max_text_bytes,
        }),
        (report.allocations > options.max_allocations).then_some(EncodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        }),
    ];
    if let Some(limit) = checks.into_iter().flatten().next() {
        return Err(EncodeError::resource(limit));
    }
    Ok(())
}

fn encode_int32(value: i32) -> u64 {
    if value < 0 {
        u64::from_ne_bytes(i64::from(value).to_ne_bytes())
    } else {
        u64::try_from(value).unwrap_or_default()
    }
}

fn key_len(field: u32, wire_type: u8) -> Option<usize> {
    varint_size((u64::from(field) << 3) | u64::from(wire_type))
}

fn varint_size(mut value: u64) -> Option<usize> {
    let mut size = 1usize;
    while value >= 0x80 {
        value >>= 7;
        size = size.checked_add(1)?;
    }
    Some(size)
}

fn varint_field_len(field: u32, value: u64) -> usize {
    key_len(field, 0)
        .unwrap_or(usize::MAX)
        .saturating_add(varint_size(value).unwrap_or(usize::MAX))
}

fn fixed32_field_len(field: u32) -> usize {
    key_len(field, 5).unwrap_or(usize::MAX).saturating_add(4)
}

fn length_field_len(field: u32, inner: usize) -> Option<usize> {
    key_len(field, 2)?
        .checked_add(varint_size(u64::try_from(inner).ok()?)?)?
        .checked_add(inner)
}

/// One source-order storage reference rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnnotationAuthorStorageRewrite {
    /// Append one canonical `TSP.Reference` after all source fields.
    Append { identifier: u64 },
    /// Remove one source-order reference after checking its current identity.
    Remove {
        ordinal: usize,
        expected_identifier: u64,
    },
}

/// Concise neutral spelling for one storage rewrite.
pub type AuthorStorageRewrite = AnnotationAuthorStorageRewrite;

impl AnnotationAuthorStorageRewrite {
    #[must_use]
    pub const fn append(identifier: u64) -> Self {
        Self::Append { identifier }
    }

    #[must_use]
    pub const fn remove(ordinal: usize, expected_identifier: u64) -> Self {
        Self::Remove {
            ordinal,
            expected_identifier,
        }
    }
}

/// Resource reservation replayed before a storage rewrite allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    allocations: usize,
    scratch_bytes: usize,
    retained_bytes: usize,
}

impl RewriteExecutionRequirements {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
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
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            allocations: self.allocations,
            scratch_bytes: self.scratch_bytes,
            retained_bytes: self.retained_bytes,
        }
    }

    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-provided storage rewrite ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub references: usize,
    pub allocations: usize,
    pub scratch_bytes: usize,
    pub retained_bytes: usize,
}

impl RewriteExecutionLimits {
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            input_bytes: usize::MAX,
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            references: usize::MAX,
            allocations: usize::MAX,
            scratch_bytes: usize::MAX,
            retained_bytes: usize::MAX,
        }
    }
}

/// Exact accounting returned after a storage rewrite executes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    requirements: RewriteExecutionRequirements,
    changed: bool,
}

impl RewriteReport {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }
    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.requirements.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.requirements.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.requirements.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.requirements.work_bytes
    }
    #[must_use]
    pub const fn references(self) -> usize {
        self.requirements.references
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.requirements.allocations
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.requirements.scratch_bytes
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.requirements.retained_bytes
    }
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Prepared source-bound author-storage rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedAnnotationAuthorStorageRewrite<'source> {
    source: &'source [u8],
    options: DecodeOptions,
    edit: AnnotationAuthorStorageRewrite,
    removed_range: Option<(usize, usize)>,
    requirements: RewriteExecutionRequirements,
    source_report: DecodeReport,
}

impl<'source> PreparedAnnotationAuthorStorageRewrite<'source> {
    #[must_use]
    pub const fn requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute only after the caller has replayed the prepared limits.
    pub fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
        check_rewrite_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_error| {
                DecodeError::resource(DecodeLimit::Allocations {
                    observed: self.requirements.allocations,
                    maximum: limits.allocations,
                })
            })?;
        match self.removed_range {
            Some((start, end)) => {
                output.extend_from_slice(&self.source[..start]);
                output.extend_from_slice(&self.source[end..]);
            },
            None => {
                output.extend_from_slice(self.source);
                let identifier = match self.edit {
                    AnnotationAuthorStorageRewrite::Append { identifier } => identifier,
                    AnnotationAuthorStorageRewrite::Remove { .. } => {
                        return Err(DecodeError::invalid());
                    },
                };
                append_reference_field(&mut output, identifier)?;
            },
        }
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        let changed = self.source != output;
        let candidate_options = self
            .options
            .with_max_message_bytes(self.options.max_message_bytes.max(output.len()));
        let (candidate, result) =
            decode_annotation_author_storage_with_report(&output, candidate_options)?;
        let expected = expected_storage_after(
            self.source,
            self.edit,
            self.removed_range,
            candidate_options,
        )?;
        if candidate != expected {
            return Err(DecodeError::invalid());
        }
        Ok((
            output,
            RewriteReport {
                source: self.source_report,
                result,
                requirements: self.requirements,
                changed,
            },
        ))
    }
}

/// Prepare a bounded storage append/remove without allocating candidate bytes.
pub fn prepare_annotation_author_storage_rewrite<'source>(
    source: &'source [u8],
    edit: AnnotationAuthorStorageRewrite,
    options: DecodeOptions,
) -> Result<PreparedAnnotationAuthorStorageRewrite<'source>, DecodeError> {
    if matches!(edit, AnnotationAuthorStorageRewrite::Append { identifier } if identifier == 0) {
        return Err(DecodeError::invalid());
    }
    let mut budget = Budget::new(source, options)?;
    let (snapshot, fields) = decode_storage_strict(source, &mut budget, 1)?;
    parity_storage(source, &snapshot, &mut budget)?;
    let references = snapshot.author_refs.len();
    let removed_range = match edit {
        AnnotationAuthorStorageRewrite::Append { .. } => None,
        AnnotationAuthorStorageRewrite::Remove {
            ordinal,
            expected_identifier,
        } => {
            let mut current = 0usize;
            let mut selected = None;
            for field in &fields {
                if field.number != STORAGE_AUTHOR_FIELD {
                    continue;
                }
                if current == ordinal {
                    let actual = snapshot
                        .author_refs
                        .get(ordinal)
                        .ok_or_else(DecodeError::invalid)?
                        .identifier();
                    if actual != expected_identifier {
                        return Err(DecodeError::invalid());
                    }
                    selected = Some((field.start, field.end));
                    break;
                }
                current = current.checked_add(1).ok_or_else(DecodeError::invalid)?;
            }
            selected.ok_or_else(DecodeError::invalid).map(Some)?
        },
    };
    let output_bytes = match (edit, removed_range) {
        (AnnotationAuthorStorageRewrite::Append { identifier }, None) => source
            .len()
            .checked_add(
                length_field_len(
                    STORAGE_AUTHOR_FIELD,
                    varint_field_len(REFERENCE_IDENTIFIER_FIELD, identifier),
                )
                .ok_or_else(DecodeError::invalid)?,
            )
            .ok_or_else(DecodeError::invalid)?,
        (AnnotationAuthorStorageRewrite::Remove { .. }, Some((start, end))) => source
            .len()
            .checked_sub(end.checked_sub(start).ok_or_else(DecodeError::invalid)?)
            .ok_or_else(DecodeError::invalid)?,
        _ => return Err(DecodeError::invalid()),
    };
    let fields = budget
        .fields
        .checked_add(
            output_bytes
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?,
        )
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = budget
        .work_bytes
        .checked_add(
            output_bytes
                .checked_mul(2)
                .ok_or_else(DecodeError::invalid)?,
        )
        .ok_or_else(DecodeError::invalid)?;
    let references = match edit {
        AnnotationAuthorStorageRewrite::Append { .. } => {
            references.checked_add(1).ok_or_else(DecodeError::invalid)?
        },
        AnnotationAuthorStorageRewrite::Remove { .. } => {
            references.checked_sub(1).ok_or_else(DecodeError::invalid)?
        },
    };
    let allocations = rewrite_allocation_upper_bound(budget.allocations, output_bytes, edit)?;
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        max_depth: budget.max_depth,
        references,
        allocations,
        scratch_bytes: 0,
        retained_bytes: output_bytes,
    };
    ensure_rewrite_requirements_limits(requirements, options)?;
    Ok(PreparedAnnotationAuthorStorageRewrite {
        source,
        options,
        edit,
        removed_range,
        requirements,
        source_report: budget.report(),
    })
}

/// Concise neutral spelling for a prepared author-storage rewrite.
pub fn prepare_author_storage_rewrite<'source>(
    source: &'source [u8],
    edit: AuthorStorageRewrite,
    options: DecodeOptions,
) -> Result<PreparedAnnotationAuthorStorageRewrite<'source>, DecodeError> {
    prepare_annotation_author_storage_rewrite(source, edit, options)
}

/// Execute a storage append/remove with an unrestricted local ceiling.
pub fn rewrite_annotation_author_storage(
    source: &[u8],
    edit: AnnotationAuthorStorageRewrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_annotation_author_storage_rewrite(source, edit, options)?;
    let limits = RewriteExecutionLimits {
        input_bytes: usize::MAX,
        output_bytes: usize::MAX,
        fields: usize::MAX,
        work_bytes: usize::MAX,
        max_depth: u32::MAX,
        references: usize::MAX,
        allocations: usize::MAX,
        scratch_bytes: usize::MAX,
        retained_bytes: usize::MAX,
    };
    Ok(prepared.execute(limits)?.0)
}

/// Concise neutral spelling for an author-storage rewrite.
pub fn rewrite_author_storage(
    source: &[u8],
    edit: AuthorStorageRewrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_annotation_author_storage(source, edit, options)
}

fn check_rewrite_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.input_bytes > limits.input_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: requirements.input_bytes,
            maximum: limits.input_bytes,
        }));
    }
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::resource(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.references > limits.references {
        return Err(DecodeError::resource(DecodeLimit::References {
            observed: requirements.references,
            maximum: limits.references,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::resource(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::resource(DecodeLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::resource(DecodeLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    Ok(())
}

/// Count the allocations required by the output buffer and both verification
/// passes. The strict source decoder charges each `Vec::try_reserve` push;
/// candidate and expected snapshots are bounded by the source count plus their
/// one logical append. Removal can only shrink those vectors, so the source
/// count is a safe upper bound for both phases.
fn rewrite_allocation_upper_bound(
    source_allocations: usize,
    output_bytes: usize,
    edit: AnnotationAuthorStorageRewrite,
) -> Result<usize, DecodeError> {
    let output_buffer = usize::from(output_bytes != 0);
    let candidate = match edit {
        AnnotationAuthorStorageRewrite::Append { .. } => source_allocations
            .checked_add(3)
            .ok_or_else(DecodeError::invalid)?,
        AnnotationAuthorStorageRewrite::Remove { .. } => source_allocations,
    };
    let expected = match edit {
        AnnotationAuthorStorageRewrite::Append { .. } => source_allocations
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?,
        AnnotationAuthorStorageRewrite::Remove { .. } => source_allocations,
    };
    output_buffer
        .checked_add(candidate)
        .and_then(|value| value.checked_add(expected))
        .ok_or_else(DecodeError::invalid)
}

fn ensure_rewrite_requirements_limits(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.input_bytes > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: requirements.input_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.output_bytes > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.fields > options.max_fields {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if requirements.references > options.max_references {
        return Err(DecodeError::resource(DecodeLimit::References {
            observed: requirements.references,
            maximum: options.max_references,
        }));
    }
    if requirements.allocations > options.max_allocations {
        return Err(DecodeError::resource(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: options.max_allocations,
        }));
    }
    Ok(())
}

fn expected_storage_after(
    source: &[u8],
    edit: AnnotationAuthorStorageRewrite,
    removed_range: Option<(usize, usize)>,
    options: DecodeOptions,
) -> Result<AnnotationAuthorStorageSnapshot, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let (mut snapshot, _) = decode_storage_strict(source, &mut budget, 1)?;
    match (edit, removed_range) {
        (AnnotationAuthorStorageRewrite::Append { identifier }, None) => {
            reserve_push(
                &mut snapshot.author_refs,
                AuthorReferenceSnapshot {
                    identifier,
                    deprecated_type: None,
                    deprecated_is_external: None,
                },
                &mut budget,
            )?;
        },
        (AnnotationAuthorStorageRewrite::Remove { ordinal, .. }, Some(_)) => {
            if ordinal >= snapshot.author_refs.len() {
                return Err(DecodeError::invalid());
            }
            snapshot.author_refs.remove(ordinal);
        },
        _ => return Err(DecodeError::invalid()),
    }
    Ok(snapshot)
}

fn append_reference_field(output: &mut Vec<u8>, identifier: u64) -> Result<(), DecodeError> {
    let inner = varint_field_len(REFERENCE_IDENTIFIER_FIELD, identifier);
    let length = u64::try_from(inner).map_err(|_error| DecodeError::invalid())?;
    push_varint(output, (u64::from(STORAGE_AUTHOR_FIELD) << 3) | 2);
    push_varint(output, length);
    push_varint(output, u64::from(REFERENCE_IDENTIFIER_FIELD) << 3);
    push_varint(output, identifier);
    Ok(())
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{tsk, tsp};
    use prost::Message as _;

    fn color() -> AuthorColorWrite {
        AuthorColorWrite::new(
            1,
            Some(0.2),
            Some(0.3),
            Some(0.4),
            Some(1.0),
            None,
            None,
            None,
            None,
            None,
            Some(1),
        )
    }

    fn author() -> AnnotationAuthorWrite<'static> {
        static IDS: [&str; 2] = ["public-a", "public-b"];
        AnnotationAuthorWrite::new(
            Some("Ada"),
            Some(color()),
            Some("public-a"),
            Some(false),
            &IDS,
        )
    }

    #[test]
    fn author_round_trip_preserves_optional_values() {
        let write = author();
        let (bytes, report) =
            encode_annotation_author_with_report(&write, EncodeOptions::for_author(&write))
                .unwrap();
        assert_eq!(report.allocations(), 3);
        let snapshot = decode_annotation_author(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(snapshot.name(), Some("Ada"));
        assert_eq!(snapshot.public_id(), Some("public-a"));
        assert_eq!(snapshot.is_public_author(), Some(false));
        assert_eq!(
            snapshot.public_ids().collect::<Vec<_>>(),
            ["public-a", "public-b"]
        );
        assert_eq!(snapshot.color().unwrap().red(), Some(0.2));
    }

    #[test]
    fn encode_reports_charge_only_materialized_view_allocations() {
        static EMPTY: [&str; 0] = [];
        let empty_author = AnnotationAuthorWrite::new(None, None, None, None, &EMPTY);
        let (empty_author_bytes, empty_author_report) = encode_annotation_author_with_report(
            &empty_author,
            EncodeOptions::for_author(&empty_author),
        )
        .unwrap();
        assert!(empty_author_bytes.is_empty());
        assert_eq!(empty_author_report.allocations(), 0);

        let identifiers = [7, 9];
        let storage = AnnotationAuthorStorageWrite::new(&identifiers);
        let (_, storage_report) = encode_annotation_author_storage_with_report(
            &storage,
            EncodeOptions::for_storage(&storage),
        )
        .unwrap();
        assert_eq!(storage_report.allocations(), 2);

        let empty_storage = AnnotationAuthorStorageWrite::new(&EMPTY_U64);
        let (_, empty_storage_report) = encode_annotation_author_storage_with_report(
            &empty_storage,
            EncodeOptions::for_storage(&empty_storage),
        )
        .unwrap();
        assert_eq!(empty_storage_report.allocations(), 0);
    }

    #[test]
    fn storage_append_remove_preserves_unknown_bytes() {
        let source = [
            0x0a, 0x02, 0x08, 0x07, 0x7a, 0x01, 0x41, 0x0a, 0x02, 0x08, 0x09,
        ];
        let options = DecodeOptions::for_source(&source);
        let appended = rewrite_annotation_author_storage(
            &source,
            AnnotationAuthorStorageRewrite::append(11),
            options,
        )
        .unwrap();
        assert!(appended.starts_with(&source));
        assert_eq!(appended[2..5], source[2..5]);
        let removed = rewrite_annotation_author_storage(
            &appended,
            AnnotationAuthorStorageRewrite::remove(0, 7),
            DecodeOptions::for_source(&appended),
        )
        .unwrap();
        assert_eq!(
            removed,
            [
                0x7a, 0x01, 0x41, 0x0a, 0x02, 0x08, 0x09, 0x0a, 0x02, 0x08, 0x0b,
            ]
        );
    }

    #[test]
    fn duplicate_singular_and_invalid_color_are_rejected() {
        let duplicate = [0x0a, 0x01, b'a', 0x0a, 0x01, b'b'];
        let error = decode_annotation_author(&duplicate, DecodeOptions::for_source(&duplicate))
            .expect_err("duplicate name");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSK.AnnotationAuthorArchive.name")
        );

        let nonfinite = [0x12, 0x05, 0x08, 0x01, 0x1d, 0, 0, 0xc0, 0x7f];
        assert!(
            decode_annotation_author(&nonfinite, DecodeOptions::for_source(&nonfinite)).is_err()
        );
    }

    #[test]
    fn unknown_fields_are_accepted_but_known_wire_and_text_remain_strict() {
        let unknown = [0x7a, 0x01, b'x'];
        let snapshot =
            decode_annotation_author(&unknown, DecodeOptions::for_source(&unknown)).unwrap();
        assert_eq!(snapshot.name(), None);

        let invalid_utf8 = [0x0a, 0x01, 0xff];
        let error =
            decode_annotation_author(&invalid_utf8, DecodeOptions::for_source(&invalid_utf8))
                .expect_err("invalid author name UTF-8");
        assert_eq!(
            error.invalid_utf8_field(),
            Some("TSK.AnnotationAuthorArchive.name")
        );

        let bad_wire = [0x08, 0x01];
        assert!(decode_annotation_author(&bad_wire, DecodeOptions::for_source(&bad_wire)).is_err());
    }

    #[test]
    fn buffa_author_view_matches_prost_for_every_native_color_model() {
        let colors = [
            tsp::Color {
                model: 1,
                r: Some(0.2),
                g: Some(0.3),
                b: Some(0.4),
                a: Some(1.0),
                rgbspace: Some(1),
                ..Default::default()
            },
            tsp::Color {
                model: 2,
                c: Some(0.1),
                m: Some(0.2),
                y: Some(0.3),
                k: Some(0.4),
                a: Some(0.9),
                ..Default::default()
            },
            tsp::Color {
                model: 3,
                w: Some(0.75),
                a: Some(0.8),
                ..Default::default()
            },
        ];
        for (index, color) in colors.into_iter().enumerate() {
            let author = tsk::AnnotationAuthorArchive {
                name: Some(format!("native-author-{index}")),
                color: Some(color.clone()),
                public_id: Some(format!("native-public-{index}")),
                is_public_author: Some(index % 2 == 0),
                public_ids: vec![format!("native-public-{index}"), "shared".to_owned()],
            };
            let source = author.encode_to_vec();
            let prost_author = tsk::AnnotationAuthorArchive::decode(source.as_slice()).unwrap();
            assert_eq!(prost_author, author);
            let public_id = author.public_id.as_deref();
            let public_ids = [author.public_ids[0].as_str(), author.public_ids[1].as_str()];
            let write = AnnotationAuthorWrite::new(
                author.name.as_deref(),
                author.color.as_ref().map(|value| {
                    AuthorColorWrite::new(
                        value.model,
                        value.r,
                        value.g,
                        value.b,
                        value.a,
                        value.c,
                        value.m,
                        value.y,
                        value.k,
                        value.w,
                        value.rgbspace,
                    )
                }),
                public_id,
                author.is_public_author,
                &public_ids,
            );
            let encoded = encode_annotation_author(&write, EncodeOptions::for_author(&write))
                .expect("Buffa author encoding");
            assert_eq!(encoded, source);
            let snapshot =
                decode_annotation_author(&source, DecodeOptions::for_source(&source)).unwrap();
            assert_eq!(snapshot.name(), author.name.as_deref());
            assert_eq!(snapshot.public_id(), author.public_id.as_deref());
            assert_eq!(snapshot.is_public_author(), author.is_public_author);
            assert_eq!(
                snapshot.public_ids().collect::<Vec<_>>(),
                [author.public_ids[0].as_str(), author.public_ids[1].as_str()]
            );
            let actual = snapshot.color().expect("native color");
            assert_eq!(actual.model(), color.model);
            assert_eq!(actual.red(), color.r);
            assert_eq!(actual.green(), color.g);
            assert_eq!(actual.blue(), color.b);
            assert_eq!(actual.alpha(), color.a);
            assert_eq!(actual.cyan(), color.c);
            assert_eq!(actual.magenta(), color.m);
            assert_eq!(actual.yellow(), color.y);
            assert_eq!(actual.black(), color.k);
            assert_eq!(actual.white(), color.w);
            assert_eq!(actual.rgbspace(), color.rgbspace);
        }
    }

    #[test]
    fn decode_and_encode_limits_are_inclusive_at_the_reported_boundary() {
        let write = author();
        let (source, encode_report) =
            encode_annotation_author_with_report(&write, EncodeOptions::for_author(&write))
                .unwrap();
        let exact_encode = EncodeOptions::new(
            encode_report.output_bytes(),
            encode_report.fields(),
            encode_report.work_bytes(),
            encode_report.references(),
            encode_report.text_bytes(),
            encode_report.allocations(),
        );
        let (encoded_again, _) =
            encode_annotation_author_with_report(&write, exact_encode).unwrap();
        assert_eq!(encoded_again, source);
        assert!(matches!(
            encode_annotation_author(
                &write,
                exact_encode.with_max_output_bytes(encode_report.output_bytes() - 1),
            ),
            Err(EncodeError {
                kind: EncodeErrorKind::Resource(EncodeLimit::OutputBytes { .. })
            })
        ));

        let (_, decode_report) =
            decode_annotation_author_with_report(&source, DecodeOptions::for_source(&source))
                .unwrap();
        let exact_decode = DecodeOptions::new(
            source.len(),
            decode_report.fields(),
            decode_report.work_bytes(),
            decode_report.max_depth().max(1),
            decode_report.references(),
            decode_report.text_bytes(),
            decode_report.allocations(),
        );
        decode_annotation_author(&source, exact_decode).unwrap();
        assert!(matches!(
            decode_annotation_author(
                &source,
                exact_decode.with_max_work_bytes(decode_report.work_bytes() - 1),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Work { .. })
            })
        ));
        assert!(matches!(
            decode_annotation_author(
                &source,
                exact_decode.with_max_fields(decode_report.fields() - 1),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Fields { .. })
            })
        ));
        assert!(matches!(
            decode_annotation_author(
                &source,
                exact_decode.with_max_text_bytes(decode_report.text_bytes() - 1),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Text { .. })
            })
        ));
    }

    #[test]
    fn prepared_storage_rewrite_accepts_exact_requirements_only() {
        let source = [0x0a, 0x02, 0x08, 0x07, 0x7a, 0x01, 0x41];
        let options = DecodeOptions::for_source(&source);
        let prepared = prepare_annotation_author_storage_rewrite(
            &source,
            AnnotationAuthorStorageRewrite::append(11),
            options,
        )
        .unwrap();
        let requirements = prepared.requirements();
        assert!(requirements.allocations() > 1);
        let (candidate, report) = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(
            candidate,
            [
                0x0a, 0x02, 0x08, 0x07, 0x7a, 0x01, 0x41, 0x0a, 0x02, 0x08, 0x0b
            ]
        );
        assert_eq!(report.output_bytes(), requirements.output_bytes());

        let prepared = prepare_annotation_author_storage_rewrite(
            &source,
            AnnotationAuthorStorageRewrite::append(11),
            options,
        )
        .unwrap();
        let mut insufficient = prepared.requirements().exact();
        insufficient.output_bytes -= 1;
        assert!(matches!(
            prepared.execute(insufficient),
            Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::OutputBytes { .. })
            })
        ));

        let prepared = prepare_annotation_author_storage_rewrite(
            &source,
            AnnotationAuthorStorageRewrite::append(11),
            options,
        )
        .unwrap();
        let mut insufficient = prepared.requirements().exact();
        insufficient.allocations -= 1;
        assert!(matches!(
            prepared.execute(insufficient),
            Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Allocations { .. })
            })
        ));

        let removed = prepare_annotation_author_storage_rewrite(
            &source,
            AnnotationAuthorStorageRewrite::remove(0, 7),
            options,
        )
        .unwrap();
        assert_eq!(removed.requirements().references(), 0);
    }

    const EMPTY_U64: [u64; 0] = [];
}
