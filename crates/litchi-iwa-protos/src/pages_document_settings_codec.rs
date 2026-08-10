//! Strict, generated-free Pages document-settings projection.
//!
//! The caller retains both root and settings wire payloads. This module only
//! publishes bounded scalar facts after raw preflight agrees with Buffa.

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_body_generated::LitchiIwaProjection as projection;

const ROOT_SUPER: u32 = 15;
const ROOT_SETTINGS: u32 = 7;
const ROOT_BODY_STORAGE: u32 = 4;
const ROOT_INITIAL_SECTION: u32 = 5;
const REFERENCE_IDENTIFIER: u32 = 1;
const REFERENCE_TYPE: u32 = 2;
const REFERENCE_EXTERNAL: u32 = 3;
const SETTINGS_BODY: u32 = 1;
const SETTINGS_HEADERS: u32 = 2;
const SETTINGS_FOOTERS: u32 = 3;
const SETTINGS_HYPHENATION: u32 = 9;
const SETTINGS_USE_LIGATURES: u32 = 10;
const SETTINGS_FOOTNOTE_KIND: u32 = 30;
const SETTINGS_FOOTNOTE_FORMAT: u32 = 31;
const SETTINGS_FOOTNOTE_NUMBERING: u32 = 32;
const SETTINGS_FOOTNOTE_GAP: u32 = 33;
const SETTINGS_FACING_PAGES: u32 = 34;
const MAX_RECURSION: u32 = 64;

/// Finite aggregate policy for a root document and its Settings object.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}
impl DecodeOptions {
    #[must_use]
    pub const fn new(bytes: usize, fields: usize, work: usize, recursion: u32) -> Self {
        Self {
            max_message_bytes: bytes,
            max_fields: fields,
            max_work_bytes: work,
            recursion_limit: recursion,
        }
    }
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact local reference to the document's Settings archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: NonZeroU64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}
impl ReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
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

/// Presence-preserving selected `TP.SettingsArchive` Boolean options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentOptionsSnapshot {
    body: Option<bool>,
    headers: Option<bool>,
    footers: Option<bool>,
    hyphenation: Option<bool>,
    use_ligatures: Option<bool>,
    facing_pages: Option<bool>,
}
impl DocumentOptionsSnapshot {
    #[must_use]
    pub const fn body(self) -> Option<bool> {
        self.body
    }
    #[must_use]
    pub const fn headers(self) -> Option<bool> {
        self.headers
    }
    #[must_use]
    pub const fn footers(self) -> Option<bool> {
        self.footers
    }
    #[must_use]
    pub const fn hyphenation(self) -> Option<bool> {
        self.hyphenation
    }
    #[must_use]
    pub const fn use_ligatures(self) -> Option<bool> {
        self.use_ligatures
    }
    #[must_use]
    pub const fn facing_pages(self) -> Option<bool> {
        self.facing_pages
    }
}

/// Generated-free settings facts joined to the exact root reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentSettingsSnapshot {
    settings_reference: ReferenceSnapshot,
    document_options: DocumentOptionsSnapshot,
    footnote_kind: Option<i32>,
    footnote_format: Option<i32>,
    footnote_numbering: Option<i32>,
    footnote_gap: Option<i32>,
}
impl DocumentSettingsSnapshot {
    #[must_use]
    pub const fn settings_reference(self) -> ReferenceSnapshot {
        self.settings_reference
    }
    #[must_use]
    pub const fn document_options(self) -> DocumentOptionsSnapshot {
        self.document_options
    }
    #[must_use]
    pub const fn footnote_kind(self) -> Option<i32> {
        self.footnote_kind
    }
    #[must_use]
    pub const fn footnote_format(self) -> Option<i32> {
        self.footnote_format
    }
    #[must_use]
    pub const fn footnote_numbering(self) -> Option<i32> {
        self.footnote_numbering
    }
    #[must_use]
    pub const fn footnote_gap(self) -> Option<i32> {
        self.footnote_gap
    }
}

/// Content-free wire byte/nesting limit failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}

/// Strict preflight or projection failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError(ErrorKind);
#[derive(Debug, Clone, PartialEq, Eq)]
enum ErrorKind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    Missing(&'static str),
    Duplicate(&'static str),
    NonCanonical(&'static str),
    Field { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Projection,
}
impl DecodeError {
    const fn resource(x: WireResourceLimit) -> Self {
        Self(ErrorKind::Resource(x))
    }
    const fn missing(x: &'static str) -> Self {
        Self(ErrorKind::Missing(x))
    }
    const fn duplicate(x: &'static str) -> Self {
        Self(ErrorKind::Duplicate(x))
    }
    const fn canonical(x: &'static str) -> Self {
        Self(ErrorKind::NonCanonical(x))
    }
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        if let ErrorKind::Resource(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        if let ErrorKind::Missing(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        if let ErrorKind::Duplicate(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        if let ErrorKind::NonCanonical(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        if let ErrorKind::Field { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        if let ErrorKind::Work { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }
}
impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.0 {
            ErrorKind::Wire(e) => e.fmt(f),
            ErrorKind::Resource(WireResourceLimit::Bytes { .. }) => {
                f.write_str("Pages document-settings byte limit exceeded")
            },
            ErrorKind::Resource(WireResourceLimit::Nesting { .. }) => {
                f.write_str("Pages document-settings nesting limit exceeded")
            },
            ErrorKind::Missing(x) => write!(f, "missing required field {x}"),
            ErrorKind::Duplicate(x) => write!(f, "duplicate singular field {x}"),
            ErrorKind::NonCanonical(x) => write!(f, "non-canonical protobuf representation: {x}"),
            ErrorKind::Field { observed, maximum } => {
                write!(f, "visited {observed} fields; maximum is {maximum}")
            },
            ErrorKind::Work { observed, maximum } => {
                write!(f, "requires {observed} work bytes; maximum is {maximum}")
            },
            ErrorKind::Projection => {
                f.write_str("strict preflight disagrees with Buffa projection")
            },
        }
    }
}
impl std::error::Error for DecodeError {}
impl From<buffa::DecodeError> for DecodeError {
    fn from(e: buffa::DecodeError) -> Self {
        match e {
            buffa::DecodeError::MessageTooLarge | buffa::DecodeError::RecursionLimitExceeded => {
                Self(ErrorKind::Projection)
            },
            x => Self(ErrorKind::Wire(x)),
        }
    }
}

/// Decode and cross-check only the root document's local Settings reference.
pub fn decode_document_settings_reference(
    root: &[u8],
    options: DecodeOptions,
) -> Result<ReferenceSnapshot, DecodeError> {
    validate(root.len(), options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_root(root, options, &mut budget)?;
    project_root(root, options, strict)
}

/// Decode a root document and its separate Settings payload under one aggregate budget.
pub fn decode_document_settings(
    root: &[u8],
    settings: &[u8],
    options: DecodeOptions,
) -> Result<DocumentSettingsSnapshot, DecodeError> {
    let total = root
        .len()
        .checked_add(settings.len())
        .ok_or_else(|| DecodeError(ErrorKind::Projection))?;
    validate(total, options)?;
    let mut budget = Budget::new(options);
    let reference = preflight_root(root, options, &mut budget)?;
    let strict = preflight_settings(settings, options, &mut budget)?;
    let reference = project_root(root, options, reference)?;
    let projected = project_settings(settings, options)?;
    if projected != strict {
        return Err(DecodeError(ErrorKind::Projection));
    }
    Ok(DocumentSettingsSnapshot {
        settings_reference: reference,
        document_options: strict.options,
        footnote_kind: strict.footnote_kind,
        footnote_format: strict.footnote_format,
        footnote_numbering: strict.footnote_numbering,
        footnote_gap: strict.footnote_gap,
    })
}

fn validate(observed: usize, o: DecodeOptions) -> Result<(), DecodeError> {
    let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_| DecodeError(ErrorKind::Projection))?;
    if o.max_message_bytes > hard {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: o.max_message_bytes,
            maximum: hard,
        }));
    }
    if observed > o.max_message_bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed,
            maximum: o.max_message_bytes,
        }));
    }
    if o.recursion_limit == 0 || o.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::resource(WireResourceLimit::Nesting {
            observed: o.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}
struct Budget {
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
    max_recursion: u32,
}
impl Budget {
    const fn new(o: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work: 0,
            max_fields: o.max_fields,
            max_work: o.max_work_bytes,
            max_recursion: o.recursion_limit,
        }
    }
    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(|| DecodeError(ErrorKind::Projection))?;
        if observed > self.max_fields {
            return Err(DecodeError(ErrorKind::Field {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }
    fn message(&mut self, n: usize) -> Result<(), DecodeError> {
        let observed = n
            .checked_mul(2)
            .and_then(|cost| self.work.checked_add(cost))
            .ok_or_else(|| DecodeError(ErrorKind::Projection))?;
        if observed > self.max_work {
            return Err(DecodeError(ErrorKind::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.work = observed;
        Ok(())
    }
    const fn nesting(&self) -> DecodeError {
        DecodeError::resource(WireResourceLimit::Nesting {
            observed: self.max_recursion.saturating_add(1),
            maximum: self.max_recursion,
        })
    }
}

fn preflight_root(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    b.message(source.len())?;
    let mut super_seen = false;
    let mut settings = None;
    let mut body = false;
    let mut section = false;
    let mut rest = source;
    while let Some(field) = next(&mut rest, o.recursion_limit - 1, b)? {
        match field.number {
            ROOT_SUPER => {
                if super_seen {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.super"));
                }
                super_seen = true;
                let _ = field.bytes()?;
            },
            ROOT_SETTINGS => {
                if settings.is_some() {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.settings"));
                }
                settings = Some(preflight_reference(field.bytes()?, o, b)?);
            },
            ROOT_BODY_STORAGE => {
                if body {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.body_storage"));
                }
                body = true;
                let _ = field.bytes()?;
            },
            ROOT_INITIAL_SECTION => {
                if section {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.section"));
                }
                section = true;
                let _ = field.bytes()?;
            },
            _ => {},
        }
    }
    if !super_seen {
        return Err(DecodeError::missing("TP.DocumentArchive.super"));
    }
    settings.ok_or_else(|| DecodeError::missing("TP.DocumentArchive.settings"))
}
fn preflight_reference(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    b.message(source.len())?;
    let mut identifier = None;
    let mut kind = None;
    let mut external = None;
    let mut rest = source;
    let depth = o
        .recursion_limit
        .checked_sub(2)
        .ok_or_else(|| b.nesting())?;
    while let Some(field) = next(&mut rest, depth, b)? {
        match field.number {
            REFERENCE_IDENTIFIER => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                identifier = NonZeroU64::new(field.varint()?)
                    .ok_or_else(|| DecodeError::canonical("reference identifier is zero"))
                    .map(Some)?;
            },
            REFERENCE_TYPE => {
                if kind.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                kind = Some(int32(field.varint()?)?);
            },
            REFERENCE_EXTERNAL => {
                if external.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                external = Some(boolean(field.varint()?)?);
            },
            _ => {},
        }
    }
    if external == Some(true) {
        return Err(DecodeError::canonical("settings reference must be local"));
    }
    Ok(ReferenceSnapshot {
        identifier: identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
        deprecated_type: kind,
        deprecated_is_external: external,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StrictSettings {
    options: DocumentOptionsSnapshot,
    footnote_kind: Option<i32>,
    footnote_format: Option<i32>,
    footnote_numbering: Option<i32>,
    footnote_gap: Option<i32>,
}
fn preflight_settings(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
) -> Result<StrictSettings, DecodeError> {
    b.message(source.len())?;
    let mut x = StrictSettings {
        options: DocumentOptionsSnapshot {
            body: None,
            headers: None,
            footers: None,
            hyphenation: None,
            use_ligatures: None,
            facing_pages: None,
        },
        footnote_kind: None,
        footnote_format: None,
        footnote_numbering: None,
        footnote_gap: None,
    };
    let mut rest = source;
    while let Some(field) = next(&mut rest, o.recursion_limit - 1, b)? {
        match field.number {
            SETTINGS_BODY => set_bool(&mut x.options.body, field, "TP.SettingsArchive.body")?,
            SETTINGS_HEADERS => {
                set_bool(&mut x.options.headers, field, "TP.SettingsArchive.headers")?
            },
            SETTINGS_FOOTERS => {
                set_bool(&mut x.options.footers, field, "TP.SettingsArchive.footers")?
            },
            SETTINGS_HYPHENATION => set_bool(
                &mut x.options.hyphenation,
                field,
                "TP.SettingsArchive.hyphenation",
            )?,
            SETTINGS_USE_LIGATURES => set_bool(
                &mut x.options.use_ligatures,
                field,
                "TP.SettingsArchive.use_ligatures",
            )?,
            SETTINGS_FACING_PAGES => set_bool(
                &mut x.options.facing_pages,
                field,
                "TP.SettingsArchive.facing_pages",
            )?,
            SETTINGS_FOOTNOTE_KIND => set_i32(
                &mut x.footnote_kind,
                field,
                "TP.SettingsArchive.footnote_kind",
            )?,
            SETTINGS_FOOTNOTE_FORMAT => set_i32(
                &mut x.footnote_format,
                field,
                "TP.SettingsArchive.footnote_format",
            )?,
            SETTINGS_FOOTNOTE_NUMBERING => set_i32(
                &mut x.footnote_numbering,
                field,
                "TP.SettingsArchive.footnote_numbering",
            )?,
            SETTINGS_FOOTNOTE_GAP => set_i32(
                &mut x.footnote_gap,
                field,
                "TP.SettingsArchive.footnote_gap",
            )?,
            _ => {},
        }
    }
    Ok(x)
}
fn set_bool(
    slot: &mut Option<bool>,
    field: Field<'_>,
    name: &'static str,
) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::duplicate(name));
    }
    *slot = Some(boolean(field.varint()?)?);
    Ok(())
}
fn set_i32(
    slot: &mut Option<i32>,
    field: Field<'_>,
    name: &'static str,
) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::duplicate(name));
    }
    *slot = Some(int32(field.varint()?)?);
    Ok(())
}
fn boolean(x: u64) -> Result<bool, DecodeError> {
    match x {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::canonical("bool scalar is not zero or one")),
    }
}
fn int32(x: u64) -> Result<i32, DecodeError> {
    if x <= i32::MAX as u64 || x >= i32::MIN as i64 as u64 {
        Ok(x as i32)
    } else {
        Err(DecodeError::canonical("int32 scalar is not sign-extended"))
    }
}

fn project_root(
    source: &[u8],
    o: DecodeOptions,
    strict: ReferenceSnapshot,
) -> Result<ReferenceSnapshot, DecodeError> {
    let view: projection::PagesDocumentBodyArchiveLazyView<'_> =
        o.buffa().decode_lazy_view(source)?;
    let reference = view
        .settings
        .get()?
        .ok_or_else(|| DecodeError(ErrorKind::Projection))?;
    let projected = ReferenceSnapshot {
        identifier: NonZeroU64::new(reference.identifier)
            .ok_or_else(|| DecodeError(ErrorKind::Projection))?,
        deprecated_type: reference.deprecated_type,
        deprecated_is_external: reference.deprecated_is_external,
    };
    if projected != strict {
        return Err(DecodeError(ErrorKind::Projection));
    }
    Ok(strict)
}
fn project_settings(source: &[u8], o: DecodeOptions) -> Result<StrictSettings, DecodeError> {
    let view: projection::PagesSettingsArchiveLazyView<'_> = o.buffa().decode_lazy_view(source)?;
    Ok(StrictSettings {
        options: DocumentOptionsSnapshot {
            body: view.body,
            headers: view.headers,
            footers: view.footers,
            hyphenation: view.hyphenation,
            use_ligatures: view.use_ligatures,
            facing_pages: view.facing_pages,
        },
        footnote_kind: view.footnote_kind,
        footnote_format: view.footnote_format,
        footnote_numbering: view.footnote_numbering,
        footnote_gap: view.footnote_gap,
    })
}

#[derive(Clone, Copy)]
enum Value<'a> {
    Varint(u64),
    Bytes(&'a [u8]),
    Other,
}
#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: buffa::encoding::WireType,
    value: Value<'a>,
}
impl<'a> Field<'a> {
    fn varint(self) -> Result<u64, DecodeError> {
        if self.wire != buffa::encoding::WireType::Varint {
            return Err(wire(
                self.number,
                self.wire,
                buffa::encoding::WireType::Varint,
            ));
        }
        if let Value::Varint(x) = self.value {
            Ok(x)
        } else {
            Err(DecodeError(ErrorKind::Projection))
        }
    }
    fn bytes(self) -> Result<&'a [u8], DecodeError> {
        if self.wire != buffa::encoding::WireType::LengthDelimited {
            return Err(wire(
                self.number,
                self.wire,
                buffa::encoding::WireType::LengthDelimited,
            ));
        }
        if let Value::Bytes(x) = self.value {
            Ok(x)
        } else {
            Err(DecodeError(ErrorKind::Projection))
        }
    }
}
fn wire(
    number: u32,
    actual: buffa::encoding::WireType,
    expected: buffa::encoding::WireType,
) -> DecodeError {
    buffa::DecodeError::WireTypeMismatch {
        field_number: number,
        expected: expected as u8,
        actual: actual as u8,
    }
    .into()
}
fn next<'a>(
    s: &mut &'a [u8],
    depth: u32,
    b: &mut Budget,
) -> Result<Option<Field<'a>>, DecodeError> {
    if s.is_empty() {
        return Ok(None);
    }
    let (tag, canonical) = varint(s)?;
    if !canonical {
        return Err(DecodeError::canonical("protobuf field key"));
    }
    b.field()?;
    let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw >> 3;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire = buffa::encoding::WireType::from_u32(raw & 7)?;
    let value = match wire {
        buffa::encoding::WireType::Varint => {
            let (x, c) = varint(s)?;
            if !c {
                return Err(DecodeError::canonical("protobuf varint value"));
            }
            Value::Varint(x)
        },
        buffa::encoding::WireType::Fixed32 => {
            take(s, 4)?;
            Value::Other
        },
        buffa::encoding::WireType::Fixed64 => {
            take(s, 8)?;
            Value::Other
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (n, c) = varint(s)?;
            if !c {
                return Err(DecodeError::canonical("length-delimited size"));
            }
            Value::Bytes(take(
                s,
                usize::try_from(n).map_err(|_| buffa::DecodeError::MessageTooLarge)?,
            )?)
        },
        buffa::encoding::WireType::StartGroup => {
            skip_group(
                s,
                number,
                depth.checked_sub(1).ok_or_else(|| b.nesting())?,
                b,
            )?;
            Value::Other
        },
        buffa::encoding::WireType::EndGroup => {
            return Err(buffa::DecodeError::InvalidEndGroup(number).into());
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw & 7).into()),
    };
    Ok(Some(Field {
        number,
        wire,
        value,
    }))
}
fn skip_group(s: &mut &[u8], number: u32, depth: u32, b: &mut Budget) -> Result<(), DecodeError> {
    loop {
        if s.is_empty() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        }
        let before = *s;
        let (tag, c) = varint(s)?;
        if !c {
            return Err(DecodeError::canonical("protobuf field key"));
        }
        let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
        if raw >> 3 == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        if raw & 7 == 4 {
            b.field()?;
            if raw >> 3 == number {
                return Ok(());
            }
            return Err(buffa::DecodeError::InvalidEndGroup(raw >> 3).into());
        }
        *s = before;
        let _ = next(s, depth, b)?;
    }
}
fn varint(s: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *s;
    let mut value = 0;
    for index in 0..10 {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 127) << (index * 7);
        if byte & 128 == 0 {
            *s = &original[index + 1..];
            let mut n = value;
            let mut len = 1;
            while n >= 128 {
                n >>= 7;
                len += 1;
            }
            return Ok((value, len == index + 1));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}
fn take<'a>(s: &mut &'a [u8], n: usize) -> Result<&'a [u8], DecodeError> {
    if s.len() < n {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (head, tail) = s.split_at(n);
    *s = tail;
    Ok(head)
}

#[cfg(test)]
mod tests {
    use super::*;
    use prost::Message as _;
    fn options(root: &[u8], settings: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            root.len() + settings.len(),
            64,
            (root.len() + settings.len()) * 4,
            4,
        )
    }
    fn fixture() -> (Vec<u8>, Vec<u8>) {
        let root = crate::tp::DocumentArchive {
            super_: crate::tsa::DocumentArchive::default(),
            settings: Some(crate::tsp::Reference {
                identifier: 7,
                deprecated_type: Some(-2),
                deprecated_is_external: Some(false),
            }),
            ..Default::default()
        }
        .encode_to_vec();
        let settings = crate::tp::SettingsArchive {
            body: Some(false),
            headers: Some(true),
            footers: Some(false),
            hyphenation: Some(true),
            use_ligatures: Some(false),
            facing_pages: Some(true),
            footnote_kind: Some(-1),
            footnote_format: Some(2),
            footnote_numbering: Some(-3),
            footnote_gap: Some(4),
            ..Default::default()
        }
        .encode_to_vec();
        (root, settings)
    }
    #[test]
    fn prost_parity_and_root_only() -> Result<(), Box<dyn std::error::Error>> {
        let (root, settings) = fixture();
        let o = options(&root, &settings);
        let r = decode_document_settings_reference(&root, o)?;
        assert_eq!(r.identifier().get(), 7);
        assert_eq!(r.deprecated_type(), Some(-2));
        let x = decode_document_settings(&root, &settings, o)?;
        assert_eq!(x.settings_reference(), r);
        assert_eq!(x.document_options().facing_pages(), Some(true));
        assert_eq!(x.document_options().body(), Some(false));
        assert_eq!(x.footnote_kind(), Some(-1));
        assert_eq!(x.footnote_gap(), Some(4));
        Ok(())
    }
    #[test]
    fn rejects_root_reference_and_scalar_violations() {
        let (root, settings) = fixture();
        let o = options(&root, &settings);
        let duplicate = [0x7a, 0, 0x3a, 2, 8, 1, 0x3a, 2, 8, 2];
        assert_eq!(
            decode_document_settings_reference(&duplicate, DecodeOptions::new(10, 10, 40, 4))
                .expect_err("duplicate")
                .duplicate_singular_field(),
            Some("TP.DocumentArchive.settings")
        );
        let external = [0x7a, 0, 0x3a, 4, 8, 1, 0x18, 1];
        assert_eq!(
            decode_document_settings_reference(&external, DecodeOptions::new(8, 8, 32, 4))
                .expect_err("external")
                .noncanonical_reason(),
            Some("settings reference must be local")
        );
        let bad_bool = [0x08, 2];
        assert_eq!(
            decode_document_settings(
                &root,
                &bad_bool,
                DecodeOptions::new(root.len() + 2, 64, (root.len() + 2) * 4, 4)
            )
            .expect_err("bool")
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        let bad_i32 = [0xf0, 1, 0x80, 0x80, 0x80, 0x80, 0x08];
        assert_eq!(
            decode_document_settings(
                &root,
                &bad_i32,
                DecodeOptions::new(
                    root.len() + bad_i32.len(),
                    64,
                    (root.len() + bad_i32.len()) * 4,
                    4
                )
            )
            .expect_err("i32")
            .noncanonical_reason(),
            Some("int32 scalar is not sign-extended")
        );
        let _ = o;
    }
    #[test]
    fn rejects_required_duplicate_and_wrong_wire_envelopes() {
        let missing_super = [0x3a, 2, 8, 1];
        assert_eq!(
            decode_document_settings_reference(
                &missing_super,
                DecodeOptions::new(missing_super.len(), 8, 16, 4)
            )
            .expect_err("super")
            .missing_required_field(),
            Some("TP.DocumentArchive.super")
        );
        let zero_identifier = [0x7a, 0, 0x3a, 2, 8, 0];
        assert_eq!(
            decode_document_settings_reference(
                &zero_identifier,
                DecodeOptions::new(zero_identifier.len(), 8, 32, 4)
            )
            .expect_err("zero identifier")
            .noncanonical_reason(),
            Some("reference identifier is zero")
        );
        let wrong_wire_reference = [0x7a, 0, 0x38, 1];
        assert!(
            decode_document_settings_reference(
                &wrong_wire_reference,
                DecodeOptions::new(wrong_wire_reference.len(), 8, 16, 4)
            )
            .is_err()
        );

        let root = [0x7a, 0, 0x3a, 2, 8, 1];
        let duplicate_bool = [0x08, 0, 0x08, 1];
        assert_eq!(
            decode_document_settings(
                &root,
                &duplicate_bool,
                DecodeOptions::new(root.len() + duplicate_bool.len(), 16, 32, 4)
            )
            .expect_err("duplicate bool")
            .duplicate_singular_field(),
            Some("TP.SettingsArchive.body")
        );
        let wrong_wire_bool = [0x0d, 0, 0, 0, 0];
        let wrong_wire_int32 = [0xf5, 0x01, 0, 0, 0, 0];
        for settings in [wrong_wire_bool.as_slice(), wrong_wire_int32.as_slice()] {
            assert!(
                decode_document_settings(
                    &root,
                    settings,
                    DecodeOptions::new(root.len() + settings.len(), 16, 64, 4)
                )
                .is_err()
            );
        }
    }
    #[test]
    fn aggregate_limits_are_exact() {
        let (root, settings) = fixture();
        let n = root.len() + settings.len();
        let baseline = DecodeOptions::new(n, 64, usize::MAX, 4);
        let mut budget = Budget::new(baseline);
        let _ = preflight_root(&root, baseline, &mut budget).expect("root preflight");
        let _ = preflight_settings(&settings, baseline, &mut budget).expect("settings preflight");
        let work = budget.work;
        assert!(
            decode_document_settings(&root, &settings, DecodeOptions::new(n, 64, work, 4)).is_ok()
        );
        assert_eq!(
            decode_document_settings(&root, &settings, DecodeOptions::new(n - 1, 64, work, 4))
                .expect_err("bytes")
                .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: n,
                maximum: n - 1
            })
        );
        assert_eq!(
            decode_document_settings(&root, &settings, DecodeOptions::new(n, 0, work, 4))
                .expect_err("fields")
                .field_limit_values(),
            Some((1, 0))
        );
        assert_eq!(
            decode_document_settings(&root, &settings, DecodeOptions::new(n, 64, work - 1, 4))
                .expect_err("work")
                .work_limit_values(),
            Some((work, work - 1))
        );
        assert_eq!(
            decode_document_settings(&root, &settings, DecodeOptions::new(n, 64, work, 1))
                .expect_err("nesting")
                .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1
            })
        );
    }
}
