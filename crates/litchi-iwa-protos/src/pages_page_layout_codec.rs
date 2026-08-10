//! Strict private Buffa projection for Pages document page-layout scalars.
//!
//! Raw source bytes remain the sole preservation representation. This module
//! validates canonical selected fields before forcing the private lazy view.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_body_generated::LitchiIwaProjection as projection;

const SUPER: u32 = 15;
const BODY_STORAGE: u32 = 4;
const INITIAL_SECTION: u32 = 5;
const WIDTH: u32 = 30;
const HEIGHT: u32 = 31;
const LEFT: u32 = 32;
const RIGHT: u32 = 33;
const TOP: u32 = 34;
const BOTTOM: u32 = 35;
const HEADER: u32 = 36;
const FOOTER: u32 = 37;
const SCALE: u32 = 38;
const VERTICAL: u32 = 39;
const ORIENTATION: u32 = 42;
const MAX_RECURSION: u32 = 64;

/// Finite resource policy for a single Pages document layout payload.
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

/// Presence-preserving native page-layout scalar facts.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageLayoutSnapshot {
    page_width: Option<f32>,
    page_height: Option<f32>,
    left_margin: Option<f32>,
    right_margin: Option<f32>,
    top_margin: Option<f32>,
    bottom_margin: Option<f32>,
    header_margin: Option<f32>,
    footer_margin: Option<f32>,
    page_scale: Option<f32>,
    lays_out_body_vertically: Option<bool>,
    orientation: Option<u32>,
}

impl PageLayoutSnapshot {
    #[must_use]
    pub const fn page_width(self) -> Option<f32> {
        self.page_width
    }
    #[must_use]
    pub const fn page_height(self) -> Option<f32> {
        self.page_height
    }
    #[must_use]
    pub const fn left_margin(self) -> Option<f32> {
        self.left_margin
    }
    #[must_use]
    pub const fn right_margin(self) -> Option<f32> {
        self.right_margin
    }
    #[must_use]
    pub const fn top_margin(self) -> Option<f32> {
        self.top_margin
    }
    #[must_use]
    pub const fn bottom_margin(self) -> Option<f32> {
        self.bottom_margin
    }
    #[must_use]
    pub const fn header_margin(self) -> Option<f32> {
        self.header_margin
    }
    #[must_use]
    pub const fn footer_margin(self) -> Option<f32> {
        self.footer_margin
    }
    #[must_use]
    pub const fn page_scale(self) -> Option<f32> {
        self.page_scale
    }
    #[must_use]
    pub const fn lays_out_body_vertically(self) -> Option<bool> {
        self.lays_out_body_vertically
    }
    #[must_use]
    pub const fn orientation(self) -> Option<u32> {
        self.orientation
    }
}

/// Content-free byte/nesting resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}

/// Failure from strict page-layout preflight or the Buffa cross-check.
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
    const fn resource(limit: WireResourceLimit) -> Self {
        Self(ErrorKind::Resource(limit))
    }
    const fn missing(name: &'static str) -> Self {
        Self(ErrorKind::Missing(name))
    }
    const fn duplicate(name: &'static str) -> Self {
        Self(ErrorKind::Duplicate(name))
    }
    const fn noncanonical(reason: &'static str) -> Self {
        Self(ErrorKind::NonCanonical(reason))
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
            ErrorKind::Wire(x) => x.fmt(f),
            ErrorKind::Resource(WireResourceLimit::Bytes { .. }) => {
                f.write_str("Pages page-layout wire byte limit exceeded")
            },
            ErrorKind::Resource(WireResourceLimit::Nesting { .. }) => {
                f.write_str("Pages page-layout wire nesting limit exceeded")
            },
            ErrorKind::Missing(x) => write!(f, "missing required field {x}"),
            ErrorKind::Duplicate(x) => write!(f, "duplicate singular field {x}"),
            ErrorKind::NonCanonical(x) => write!(f, "non-canonical protobuf representation: {x}"),
            ErrorKind::Field { observed, maximum } => write!(
                f,
                "Pages page-layout visited {observed} fields; maximum is {maximum}"
            ),
            ErrorKind::Work { observed, maximum } => write!(
                f,
                "Pages page-layout requires {observed} work bytes; maximum is {maximum}"
            ),
            ErrorKind::Projection => {
                f.write_str("Pages page-layout strict preflight disagrees with Buffa projection")
            },
        }
    }
}
impl std::error::Error for DecodeError {}
impl From<buffa::DecodeError> for DecodeError {
    fn from(e: buffa::DecodeError) -> Self {
        match e {
            // Strict preflight measures these before Buffa runs. Buffa does not
            // expose an exact content-free observation for its later failure.
            buffa::DecodeError::MessageTooLarge | buffa::DecodeError::RecursionLimitExceeded => {
                Self(ErrorKind::Projection)
            },
            x => Self(ErrorKind::Wire(x)),
        }
    }
}

/// Decode presence-preserving page-layout scalars from one `TP.DocumentArchive`.
pub fn decode_page_layout(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PageLayoutSnapshot, DecodeError> {
    validate(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight(source, options, &mut budget)?;
    let view: projection::PagesDocumentBodyArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let projected = PageLayoutSnapshot {
        page_width: view.page_width,
        page_height: view.page_height,
        left_margin: view.left_margin,
        right_margin: view.right_margin,
        top_margin: view.top_margin,
        bottom_margin: view.bottom_margin,
        header_margin: view.header_margin,
        footer_margin: view.footer_margin,
        page_scale: view.page_scale,
        lays_out_body_vertically: view.lays_out_body_vertically,
        orientation: view.orientation,
    };
    if projected != strict {
        return Err(DecodeError(ErrorKind::Projection));
    }
    Ok(strict)
}

fn validate(source: &[u8], o: DecodeOptions) -> Result<(), DecodeError> {
    let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_| DecodeError(ErrorKind::Projection))?;
    if o.max_message_bytes > hard {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: o.max_message_bytes,
            maximum: hard,
        }));
    }
    if source.len() > o.max_message_bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: source.len(),
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
        let n = self.fields.saturating_add(1);
        if n > self.max_fields {
            return Err(DecodeError(ErrorKind::Field {
                observed: n,
                maximum: self.max_fields,
            }));
        }
        self.fields = n;
        Ok(())
    }
    fn message(&mut self, n: usize) -> Result<(), DecodeError> {
        let observed = self.work.saturating_add(n.saturating_mul(2));
        if observed > self.max_work {
            return Err(DecodeError(ErrorKind::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.work = observed;
        Ok(())
    }
    const fn nesting_limit(&self) -> DecodeError {
        DecodeError::resource(WireResourceLimit::Nesting {
            observed: self.max_recursion.saturating_add(1),
            maximum: self.max_recursion,
        })
    }
}

fn preflight(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
) -> Result<PageLayoutSnapshot, DecodeError> {
    b.message(source.len())?;
    let mut result = PageLayoutSnapshot {
        page_width: None,
        page_height: None,
        left_margin: None,
        right_margin: None,
        top_margin: None,
        bottom_margin: None,
        header_margin: None,
        footer_margin: None,
        page_scale: None,
        lays_out_body_vertically: None,
        orientation: None,
    };
    let mut body_storage_seen = false;
    let mut initial_section_seen = false;
    let mut super_seen = false;
    let mut rest = source;
    // Buffa counts the root message as one recursion level. The raw walker
    // starts inside that root, so it has one fewer nested group level left.
    while let Some(f) = next(&mut rest, o.recursion_limit - 1, b)? {
        match f.number {
            BODY_STORAGE => {
                if body_storage_seen {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.body_storage"));
                }
                body_storage_seen = true;
                let _opaque_reference = f.bytes()?;
            },
            INITIAL_SECTION => {
                if initial_section_seen {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.initial_section"));
                }
                initial_section_seen = true;
                let _opaque_reference = f.bytes()?;
            },
            SUPER => {
                if super_seen {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.super"));
                }
                super_seen = true;
                f.bytes()?;
            },
            WIDTH => set_float(&mut result.page_width, f, "TP.DocumentArchive.page_width")?,
            HEIGHT => set_float(&mut result.page_height, f, "TP.DocumentArchive.page_height")?,
            LEFT => set_float(&mut result.left_margin, f, "TP.DocumentArchive.left_margin")?,
            RIGHT => set_float(
                &mut result.right_margin,
                f,
                "TP.DocumentArchive.right_margin",
            )?,
            TOP => set_float(&mut result.top_margin, f, "TP.DocumentArchive.top_margin")?,
            BOTTOM => set_float(
                &mut result.bottom_margin,
                f,
                "TP.DocumentArchive.bottom_margin",
            )?,
            HEADER => set_float(
                &mut result.header_margin,
                f,
                "TP.DocumentArchive.header_margin",
            )?,
            FOOTER => set_float(
                &mut result.footer_margin,
                f,
                "TP.DocumentArchive.footer_margin",
            )?,
            SCALE => set_float(&mut result.page_scale, f, "TP.DocumentArchive.page_scale")?,
            VERTICAL => {
                if result.lays_out_body_vertically.is_some() {
                    return Err(DecodeError::duplicate(
                        "TP.DocumentArchive.lays_out_body_vertically",
                    ));
                }
                result.lays_out_body_vertically = Some(match f.varint()? {
                    0 => false,
                    1 => true,
                    _ => return Err(DecodeError::noncanonical("bool scalar is not zero or one")),
                });
            },
            ORIENTATION => {
                if result.orientation.is_some() {
                    return Err(DecodeError::duplicate("TP.DocumentArchive.orientation"));
                }
                result.orientation = Some(
                    u32::try_from(f.varint()?)
                        .map_err(|_| DecodeError::noncanonical("uint32 scalar is out of range"))?,
                );
            },
            _ => {},
        }
    }
    if !super_seen {
        return Err(DecodeError::missing("TP.DocumentArchive.super"));
    }
    Ok(result)
}
fn set_float(slot: &mut Option<f32>, f: Field<'_>, name: &'static str) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::duplicate(name));
    }
    *slot = Some(f.fixed32()?);
    Ok(())
}

#[derive(Clone, Copy)]
enum Value<'a> {
    Varint(u64),
    Fixed32(f32),
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
    fn fixed32(self) -> Result<f32, DecodeError> {
        if self.wire != buffa::encoding::WireType::Fixed32 {
            return Err(wire(
                self.number,
                self.wire,
                buffa::encoding::WireType::Fixed32,
            ));
        }
        if let Value::Fixed32(x) = self.value {
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
    n: u32,
    actual: buffa::encoding::WireType,
    expected: buffa::encoding::WireType,
) -> DecodeError {
    buffa::DecodeError::WireTypeMismatch {
        field_number: n,
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
    };
    let (tag, canon) = varint(s)?;
    if !canon {
        return Err(DecodeError::noncanonical("protobuf field key"));
    };
    b.field()?;
    let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw >> 3;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    };
    let wire_type = buffa::encoding::WireType::from_u32(raw & 7)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (x, c) = varint(s)?;
            if !c {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            };
            Value::Varint(x)
        },
        buffa::encoding::WireType::Fixed32 => {
            let x = take(s, 4)?;
            Value::Fixed32(f32::from_le_bytes(
                x.try_into()
                    .map_err(|_| buffa::DecodeError::UnexpectedEof)?,
            ))
        },
        buffa::encoding::WireType::Fixed64 => {
            take(s, 8)?;
            Value::Other
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (n, c) = varint(s)?;
            if !c {
                return Err(DecodeError::noncanonical("length-delimited size"));
            };
            let n = usize::try_from(n).map_err(|_| buffa::DecodeError::MessageTooLarge)?;
            Value::Bytes(take(s, n)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child = depth.checked_sub(1).ok_or_else(|| b.nesting_limit())?;
            skip_group(s, number, child, b)?;
            Value::Other
        },
        buffa::encoding::WireType::EndGroup => {
            return Err(buffa::DecodeError::InvalidEndGroup(number).into());
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw & 7).into()),
    };
    Ok(Some(Field {
        number,
        wire: wire_type,
        value,
    }))
}
fn skip_group(s: &mut &[u8], number: u32, depth: u32, b: &mut Budget) -> Result<(), DecodeError> {
    loop {
        if s.is_empty() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        };
        let before = *s;
        let (tag, c) = varint(s)?;
        if !c {
            return Err(DecodeError::noncanonical("protobuf field key"));
        };
        let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
        if raw >> 3 == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        };
        if raw & 7 == 4 {
            b.field()?;
            if raw >> 3 == number {
                return Ok(());
            }
            return Err(buffa::DecodeError::InvalidEndGroup(raw >> 3).into());
        };
        *s = before;
        let _ = next(s, depth, b)?;
    }
}
fn varint(s: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let orig = *s;
    let mut v = 0;
    for i in 0..10 {
        let byte = *orig.get(i).ok_or(buffa::DecodeError::UnexpectedEof)?;
        if i == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        };
        v |= u64::from(byte & 127) << (i * 7);
        if byte & 128 == 0 {
            *s = &orig[i + 1..];
            let mut n = v;
            let mut len = 1;
            while n >= 128 {
                n >>= 7;
                len += 1
            }
            return Ok((v, len == i + 1));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}
fn take<'a>(s: &mut &'a [u8], n: usize) -> Result<&'a [u8], DecodeError> {
    if s.len() < n {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    };
    let (a, z) = s.split_at(n);
    *s = z;
    Ok(a)
}

#[cfg(test)]
mod tests {
    use super::*;
    use prost::Message as _;
    fn opts(s: &[u8], r: u32) -> DecodeOptions {
        DecodeOptions::new(
            s.len().max(1),
            s.len().max(1),
            s.len().saturating_mul(2).max(1),
            r,
        )
    }
    #[test]
    fn prost_parity_and_presence() -> Result<(), Box<dyn std::error::Error>> {
        let s = crate::tp::DocumentArchive {
            super_: crate::tsa::DocumentArchive::default(),
            page_width: Some(612.0),
            page_height: Some(792.0),
            left_margin: Some(72.0),
            lays_out_body_vertically: Some(false),
            orientation: Some(1),
            ..Default::default()
        }
        .encode_to_vec();
        let x = decode_page_layout(&s, opts(&s, 2))?;
        assert_eq!(x.page_width(), Some(612.0));
        assert_eq!(x.page_height(), Some(792.0));
        assert_eq!(x.left_margin(), Some(72.0));
        assert_eq!(x.lays_out_body_vertically(), Some(false));
        assert_eq!(x.orientation(), Some(1));
        assert_eq!(x.right_margin(), None);
        Ok(())
    }
    #[test]
    fn rejects_wrong_wire_and_noncanonical_selected_scalars() {
        let width_varint = [0x7a, 0, 0xf0, 1, 0];
        let vertical_fixed32 = [0x7a, 0, 0xbd, 2, 0, 0, 0, 0];
        let orientation_fixed32 = [0x7a, 0, 0xd5, 2, 0, 0, 0, 0];
        for source in [
            width_varint.as_slice(),
            vertical_fixed32.as_slice(),
            orientation_fixed32.as_slice(),
        ] {
            assert!(decode_page_layout(source, opts(source, 2)).is_err());
        }
        let dup = [0x7a, 0, 0xf5, 1, 0, 0, 0, 0, 0xf5, 1, 0, 0, 0, 0];
        assert_eq!(
            decode_page_layout(&dup, opts(&dup, 2))
                .expect_err("duplicate")
                .duplicate_singular_field(),
            Some("TP.DocumentArchive.page_width")
        );
        let bad = [0x7a, 0, 0xb8, 2, 2];
        assert_eq!(
            decode_page_layout(&bad, opts(&bad, 2))
                .expect_err("bool")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn rejects_duplicate_or_wrong_wire_lazy_singular_envelopes() {
        let duplicate_body_storage = [0x7a, 0, 0x22, 0, 0x22, 0];
        let duplicate_initial_section = [0x7a, 0, 0x2a, 0, 0x2a, 0];
        assert_eq!(
            decode_page_layout(&duplicate_body_storage, opts(&duplicate_body_storage, 2))
                .expect_err("duplicate body storage")
                .duplicate_singular_field(),
            Some("TP.DocumentArchive.body_storage")
        );
        assert_eq!(
            decode_page_layout(
                &duplicate_initial_section,
                opts(&duplicate_initial_section, 2)
            )
            .expect_err("duplicate initial section")
            .duplicate_singular_field(),
            Some("TP.DocumentArchive.initial_section")
        );
        let wrong_wire_body_storage = [0x7a, 0, 0x20, 0];
        assert!(
            decode_page_layout(&wrong_wire_body_storage, opts(&wrong_wire_body_storage, 2))
                .is_err()
        );
    }

    #[test]
    fn accepts_unknown_fields_of_every_wire_type_including_groups() {
        let source = [
            0x7a, 0x00, // required super
            0xa0, 0x06, 0x01, // unknown field 100, varint
            0xa9, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, // field 101, fixed64
            0xb2, 0x06, 0x02, 0xaa, 0xbb, // field 102, length-delimited
            0xbd, 0x06, 0, 0, 0, 0, // field 103, fixed32
            0xc3, 0x06, 0x08, 0x01, 0xc4, 0x06, // field 104 group
        ];
        assert_eq!(
            decode_page_layout(&source, opts(&source, 2)).expect("unknown fields are opaque"),
            PageLayoutSnapshot {
                page_width: None,
                page_height: None,
                left_margin: None,
                right_margin: None,
                top_margin: None,
                bottom_margin: None,
                header_margin: None,
                footer_margin: None,
                page_scale: None,
                lays_out_body_vertically: None,
                orientation: None,
            }
        );
    }

    #[test]
    fn rejects_malformed_groups() {
        let missing_end = [0x7a, 0, 0x2b];
        let mismatched_end = [0x7a, 0, 0x2b, 0x34];
        for source in [missing_end.as_slice(), mismatched_end.as_slice()] {
            assert!(decode_page_layout(source, opts(source, 2)).is_err());
        }
    }

    #[test]
    fn resource_limits_are_exact_at_boundary_and_one_over() {
        let src = [0x7a, 0];
        assert!(decode_page_layout(&src, DecodeOptions::new(2, 1, 4, 1)).is_ok());
        assert_eq!(
            decode_page_layout(&src, DecodeOptions::new(1, 2, 4, 2))
                .expect_err("bytes")
                .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: 2,
                maximum: 1
            })
        );
        assert_eq!(
            decode_page_layout(&src, DecodeOptions::new(2, 0, 4, 1))
                .expect_err("fields")
                .field_limit_values(),
            Some((1, 0))
        );
        assert_eq!(
            decode_page_layout(&src, DecodeOptions::new(2, 1, 3, 1))
                .expect_err("work")
                .work_limit_values(),
            Some((4, 3))
        );

        let nested_groups = [
            0x7a, 0, // super
            0xa3, 0x06, // field 100 start group
            0xab, 0x06, // field 101 start group
            0xac, 0x06, // field 101 end group
            0xa4, 0x06, // field 100 end group
        ];
        assert!(decode_page_layout(&nested_groups, DecodeOptions::new(10, 5, 20, 3)).is_ok());
        assert_eq!(
            decode_page_layout(&nested_groups, DecodeOptions::new(10, 5, 20, 2))
                .expect_err("nesting")
                .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 3,
                maximum: 2
            })
        );
    }
}
