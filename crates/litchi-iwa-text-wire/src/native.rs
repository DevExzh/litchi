//! Private bounded wire models of the TSWP/TSS/TSD/TSP messages the
//! storage engines read and write.
//!
//! Each model decodes with protobuf merge semantics (a repeated singular
//! scalar keeps its last value, a repeated singular message merges, a
//! repeated field appends) and encodes canonically in ascending field
//! order, which is the order the generated Prost oracle writes. Only the
//! fields the engines interpret are modelled; every other field is ignored
//! on decode, and the engines decide by an exact field-set check whether a
//! payload is fully understood before they rewrite it.

use litchi_iwa_common::{LimitKind, WireLimits, wire::WireView};

/// A malformed payload or an exhausted resource inside the native layer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NativeError {
    /// The payload is malformed or carries a value the engine cannot
    /// represent.
    Invalid(&'static str),
    /// A configured resource was exhausted.
    Limit {
        kind: LimitKind,
        observed: usize,
        maximum: usize,
    },
    /// A bounded allocation failed.
    Allocation { amount: usize },
}

pub(crate) type NativeResult<T> = Result<T, NativeError>;

pub(crate) fn admit(kind: LimitKind, observed: usize, maximum: usize) -> NativeResult<()> {
    if observed > maximum {
        return Err(NativeError::Limit {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn from_common(error: litchi_iwa_common::Error) -> NativeError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => NativeError::Limit {
            kind,
            observed,
            maximum: limit,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => NativeError::Allocation { amount },
        _ => NativeError::Invalid("malformed protobuf wire"),
    }
}

/// Map an error of a common wire rewrite helper.
pub(crate) fn common(error: litchi_iwa_common::Error) -> NativeError {
    from_common(error)
}

/// One operation's aggregate scan budget: input bytes, fields and work are
/// accumulated across every payload the operation visits, so no nested or
/// chained payload resets the allowance.
pub(crate) struct Budget {
    limits: WireLimits,
    bytes: usize,
    fields: usize,
    work: usize,
}

impl Budget {
    pub(crate) const fn new(limits: WireLimits) -> Self {
        Self {
            limits,
            bytes: 0,
            fields: 0,
            work: 0,
        }
    }

    pub(crate) const fn limits(&self) -> WireLimits {
        self.limits
    }

    pub(crate) fn view<'a>(&mut self, data: &'a [u8], depth: usize) -> NativeResult<WireView<'a>> {
        admit(LimitKind::Nesting, depth, self.limits.max_nesting())?;
        self.bytes = self.bytes.checked_add(data.len()).unwrap_or(usize::MAX);
        admit(
            LimitKind::InputBytes,
            self.bytes,
            self.limits.max_input_bytes(),
        )?;
        let fields_left = self.limits.max_fields().saturating_sub(self.fields);
        let work_left = self.limits.max_rewrite_work().saturating_sub(self.work);
        let work_limited = work_left < fields_left;
        let remaining = fields_left.min(work_left);
        if remaining == 0 && !data.is_empty() {
            return Err(self.field_failure(work_limited, 1));
        }
        let policy = self
            .limits
            .with_fields(remaining.max(1))
            .map_err(from_common)?;
        let view = WireView::parse(data, policy).map_err(|error| match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                ..
            } => self.field_failure(work_limited, observed),
            other => from_common(other),
        })?;
        self.fields += view.len();
        self.work += view.len();
        Ok(view)
    }

    fn field_failure(&self, work_limited: bool, additional: usize) -> NativeError {
        let (kind, prior, maximum) = if work_limited {
            (
                LimitKind::RewriteWork,
                self.work,
                self.limits.max_rewrite_work(),
            )
        } else {
            (LimitKind::Fields, self.fields, self.limits.max_fields())
        };
        NativeError::Limit {
            kind,
            observed: prior.saturating_add(additional),
            maximum,
        }
    }

    /// Admit a payload handed to a common rewrite helper (which applies its
    /// own default ceilings) against this operation's input budget.
    pub(crate) fn admit_bytes(&mut self, length: usize) -> NativeResult<()> {
        self.bytes = self.bytes.checked_add(length).unwrap_or(usize::MAX);
        admit(
            LimitKind::InputBytes,
            self.bytes,
            self.limits.max_input_bytes(),
        )
    }
}

const VARINT: u8 = 0;
const FIXED64: u8 = 1;
const LENGTH_DELIMITED: u8 = 2;
const FIXED32: u8 = 5;

fn varint(payload: &[u8]) -> NativeResult<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_| NativeError::Invalid("malformed protobuf varint"))?;
    if length != payload.len() {
        return Err(NativeError::Invalid("malformed protobuf varint"));
    }
    Ok(value)
}

/// One decoded field: its number, wire type and payload bytes.
#[derive(Clone, Copy)]
pub(crate) struct Field<'a> {
    pub(crate) number: u32,
    wire_type: u8,
    payload: &'a [u8],
}

impl<'a> Field<'a> {
    fn expect(self, wire_type: u8) -> NativeResult<&'a [u8]> {
        if self.wire_type != wire_type {
            return Err(NativeError::Invalid(
                "protobuf field has the wrong wire type",
            ));
        }
        Ok(self.payload)
    }

    pub(crate) fn uint64(self) -> NativeResult<u64> {
        varint(self.expect(VARINT)?)
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "protobuf uint32 decoding truncates a wider varint"
    )]
    pub(crate) fn uint32(self) -> NativeResult<u32> {
        Ok(self.uint64()? as u32)
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "protobuf int32 decoding truncates the sign-extended varint"
    )]
    pub(crate) fn int32(self) -> NativeResult<i32> {
        Ok(self.uint64()? as i32)
    }

    pub(crate) fn bool(self) -> NativeResult<bool> {
        Ok(self.uint64()? != 0)
    }

    pub(crate) fn float(self) -> NativeResult<f32> {
        let bytes: [u8; 4] = self
            .expect(FIXED32)?
            .try_into()
            .map_err(|_| NativeError::Invalid("malformed protobuf fixed32"))?;
        Ok(f32::from_le_bytes(bytes))
    }

    pub(crate) fn bytes(self) -> NativeResult<&'a [u8]> {
        self.expect(LENGTH_DELIMITED)
    }

    pub(crate) fn string(self, budget: &Budget) -> NativeResult<String> {
        let text = std::str::from_utf8(self.bytes()?)
            .map_err(|_| NativeError::Invalid("protobuf string is not UTF-8"))?;
        owned_string(text, budget)
    }

    /// Decode a repeated float that may be packed or unpacked.
    pub(crate) fn floats(self, output: &mut Vec<f32>) -> NativeResult<()> {
        match self.wire_type {
            FIXED32 => {
                reserve(output, 1)?;
                output.push(self.float()?);
            },
            LENGTH_DELIMITED => {
                if self.payload.len() % 4 != 0 {
                    return Err(NativeError::Invalid("malformed packed protobuf floats"));
                }
                reserve(output, self.payload.len() / 4)?;
                for chunk in self.payload.chunks_exact(4) {
                    let bytes: [u8; 4] = chunk
                        .try_into()
                        .map_err(|_| NativeError::Invalid("malformed packed protobuf floats"))?;
                    output.push(f32::from_le_bytes(bytes));
                }
            },
            _ => {
                return Err(NativeError::Invalid(
                    "protobuf field has the wrong wire type",
                ));
            },
        }
        Ok(())
    }
}

pub(crate) fn owned_string(text: &str, budget: &Budget) -> NativeResult<String> {
    admit(
        LimitKind::OutputBytes,
        text.len(),
        budget.limits().max_output_bytes(),
    )?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| NativeError::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

pub(crate) fn reserve<T>(values: &mut Vec<T>, additional: usize) -> NativeResult<()> {
    let amount = values.len().checked_add(additional).unwrap_or(usize::MAX);
    values
        .try_reserve(additional)
        .map_err(|_| NativeError::Allocation { amount })
}

/// Visit every field of `data` in wire order.
pub(crate) fn fields<'a>(
    data: &'a [u8],
    depth: usize,
    budget: &mut Budget,
    mut visit: impl FnMut(Field<'a>, &mut Budget) -> NativeResult<()>,
) -> NativeResult<()> {
    let view = budget.view(data, depth)?;
    for field in view.fields() {
        let wire_type = field.wire_type();
        if !matches!(wire_type, VARINT | FIXED64 | LENGTH_DELIMITED | FIXED32) {
            return Err(NativeError::Invalid("protobuf groups are not supported"));
        }
        visit(
            Field {
                number: field.number(),
                wire_type,
                payload: field.payload(),
            },
            budget,
        )?;
    }
    Ok(())
}

/// The field numbers of `data` in wire order (a repeated number repeats).
pub(crate) fn field_numbers(data: &[u8], budget: &mut Budget) -> NativeResult<Vec<u32>> {
    let mut numbers = Vec::new();
    fields(data, 0, budget, |field, _| {
        reserve(&mut numbers, 1)?;
        numbers.push(field.number);
        Ok(())
    })?;
    Ok(numbers)
}

/// Whether `data` holds exactly the `expected` field numbers (as a
/// multiset, in any order).
pub(crate) fn has_exact_fields(
    data: &[u8],
    expected: &[u32],
    budget: &mut Budget,
) -> NativeResult<bool> {
    let mut actual = field_numbers(data, budget)?;
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

/// Whether field `number` occurs once, as a canonical one-byte boolean.
pub(crate) fn has_canonical_bool_field(
    data: &[u8],
    number: u32,
    budget: &mut Budget,
) -> NativeResult<bool> {
    let mut found = Vec::new();
    fields(data, 0, budget, |field, _| {
        if field.number == number {
            reserve(&mut found, 1)?;
            found.push((field.wire_type, field.payload));
        }
        Ok(())
    })?;
    let [(wire_type, payload)] = found.as_slice() else {
        return Ok(false);
    };
    Ok(*wire_type == VARINT && matches!(payload, [0] | [1]))
}

/// The payload of the single occurrence of length-delimited field `number`.
pub(crate) fn required_payload<'a>(
    data: &'a [u8],
    number: u32,
    budget: &mut Budget,
) -> NativeResult<&'a [u8]> {
    let mut found: Option<&'a [u8]> = None;
    let mut count = 0usize;
    fields(data, 0, budget, |field, _| {
        if field.number == number {
            found = Some(field.bytes()?);
            count += 1;
        }
        Ok(())
    })?;
    match (found, count) {
        (Some(payload), 1) => Ok(payload),
        _ => Err(NativeError::Invalid(
            "a required nested message must occur exactly once",
        )),
    }
}

/// Canonical protobuf writer with a bounded output.
pub(crate) struct Writer {
    output: Vec<u8>,
    maximum: usize,
}

impl Writer {
    pub(crate) const fn new(limits: WireLimits) -> Self {
        Self {
            output: Vec::new(),
            maximum: limits.max_output_bytes(),
        }
    }

    fn reserve(&mut self, additional: usize) -> NativeResult<()> {
        let observed = self
            .output
            .len()
            .checked_add(additional)
            .unwrap_or(usize::MAX);
        admit(LimitKind::OutputBytes, observed, self.maximum)?;
        self.output
            .try_reserve(additional)
            .map_err(|_| NativeError::Allocation { amount: observed })
    }

    fn raw_varint(&mut self, mut value: u64) -> NativeResult<()> {
        self.reserve(litchi_iwa_common::varint::encoded_len(value))?;
        loop {
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the low seven bits fit in a byte"
            )]
            let byte = (value & 0x7f) as u8;
            value >>= 7;
            if value == 0 {
                self.output.push(byte);
                return Ok(());
            }
            self.output.push(byte | 0x80);
        }
    }

    fn key(&mut self, number: u32, wire_type: u8) -> NativeResult<()> {
        self.raw_varint((u64::from(number) << 3) | u64::from(wire_type))
    }

    pub(crate) fn uint64(&mut self, number: u32, value: u64) -> NativeResult<()> {
        self.key(number, VARINT)?;
        self.raw_varint(value)
    }

    pub(crate) fn uint32(&mut self, number: u32, value: u32) -> NativeResult<()> {
        self.uint64(number, u64::from(value))
    }

    #[allow(
        clippy::cast_sign_loss,
        reason = "protobuf int32 encodes the sign-extended 64-bit value"
    )]
    pub(crate) fn int32(&mut self, number: u32, value: i32) -> NativeResult<()> {
        self.uint64(number, i64::from(value) as u64)
    }

    pub(crate) fn bool(&mut self, number: u32, value: bool) -> NativeResult<()> {
        self.uint64(number, u64::from(value))
    }

    pub(crate) fn float(&mut self, number: u32, value: f32) -> NativeResult<()> {
        self.key(number, FIXED32)?;
        self.reserve(4)?;
        self.output.extend_from_slice(&value.to_le_bytes());
        Ok(())
    }

    pub(crate) fn bytes(&mut self, number: u32, value: &[u8]) -> NativeResult<()> {
        self.key(number, LENGTH_DELIMITED)?;
        self.raw_varint(value.len() as u64)?;
        self.reserve(value.len())?;
        self.output.extend_from_slice(value);
        Ok(())
    }

    /// Write a nested message produced by `encode`.
    pub(crate) fn message(
        &mut self,
        number: u32,
        encode: impl FnOnce(&mut Self) -> NativeResult<()>,
    ) -> NativeResult<()> {
        let mut nested = Self {
            output: Vec::new(),
            maximum: self.maximum,
        };
        encode(&mut nested)?;
        self.bytes(number, &nested.output)
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.output
    }
}

macro_rules! optional {
    ($writer:expr, $method:ident, $number:expr, $value:expr) => {
        if let Some(value) = $value {
            $writer.$method($number, value)?;
        }
    };
}

/// `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) struct Reference {
    pub(crate) identifier: u64,
}

impl Reference {
    pub(crate) fn merge(
        &mut self,
        data: &[u8],
        depth: usize,
        budget: &mut Budget,
    ) -> NativeResult<()> {
        fields(data, depth, budget, |field, _| {
            match field.number {
                1 => self.identifier = field.uint64()?,
                2 => {
                    field.int32()?;
                },
                3 => {
                    field.bool()?;
                },
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(self, writer: &mut Writer) -> NativeResult<()> {
        writer.uint64(1, self.identifier)
    }
}

/// `TSP.Point`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(crate) struct Point {
    pub(crate) x: f32,
    pub(crate) y: f32,
}

impl Point {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, _| {
            match field.number {
                1 => self.x = field.float()?,
                2 => self.y = field.float()?,
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(self, writer: &mut Writer) -> NativeResult<()> {
        writer.float(1, self.x)?;
        writer.float(2, self.y)
    }
}

/// `TSP.Color`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
#[allow(clippy::struct_field_names, reason = "the schema's channel names")]
pub(crate) struct Color {
    pub(crate) model: i32,
    pub(crate) r: Option<f32>,
    pub(crate) g: Option<f32>,
    pub(crate) b: Option<f32>,
    pub(crate) a: Option<f32>,
    pub(crate) c: Option<f32>,
    pub(crate) m: Option<f32>,
    pub(crate) y: Option<f32>,
    pub(crate) k: Option<f32>,
    pub(crate) w: Option<f32>,
    pub(crate) rgbspace: Option<i32>,
}

impl Color {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, _| {
            match field.number {
                1 => self.model = field.int32()?,
                3 => self.r = Some(field.float()?),
                4 => self.g = Some(field.float()?),
                5 => self.b = Some(field.float()?),
                6 => self.a = Some(field.float()?),
                7 => self.c = Some(field.float()?),
                8 => self.m = Some(field.float()?),
                9 => self.y = Some(field.float()?),
                10 => self.k = Some(field.float()?),
                11 => self.w = Some(field.float()?),
                12 => self.rgbspace = Some(field.int32()?),
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        writer.int32(1, self.model)?;
        optional!(writer, float, 3, self.r);
        optional!(writer, float, 4, self.g);
        optional!(writer, float, 5, self.b);
        optional!(writer, float, 6, self.a);
        optional!(writer, float, 7, self.c);
        optional!(writer, float, 8, self.m);
        optional!(writer, float, 9, self.y);
        optional!(writer, float, 10, self.k);
        optional!(writer, float, 11, self.w);
        optional!(writer, int32, 12, self.rgbspace);
        Ok(())
    }
}

/// `TSD.StrokePatternArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct StrokePattern {
    pub(crate) kind: Option<i32>,
    pub(crate) phase: Option<f32>,
    pub(crate) count: Option<u32>,
    pub(crate) pattern: Vec<f32>,
}

impl StrokePattern {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, _| {
            match field.number {
                1 => self.kind = Some(field.int32()?),
                2 => self.phase = Some(field.float()?),
                3 => self.count = Some(field.uint32()?),
                4 => field.floats(&mut self.pattern)?,
                _ => {},
            }
            Ok(())
        })
    }

    fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        optional!(writer, int32, 1, self.kind);
        optional!(writer, float, 2, self.phase);
        optional!(writer, uint32, 3, self.count);
        for value in &self.pattern {
            writer.float(4, *value)?;
        }
        Ok(())
    }
}

/// `TSD.StrokeArchive` (smart, frame and patterned strokes are kept only
/// as presence: the engines refuse them).
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct Stroke {
    pub(crate) color: Option<Color>,
    pub(crate) width: Option<f32>,
    pub(crate) cap: Option<i32>,
    pub(crate) join: Option<i32>,
    pub(crate) miter_limit: Option<f32>,
    pub(crate) pattern: Option<StrokePattern>,
    pub(crate) has_special: bool,
}

impl Stroke {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.color.get_or_insert_with(Color::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                2 => self.width = Some(field.float()?),
                3 => self.cap = Some(field.int32()?),
                4 => self.join = Some(field.int32()?),
                5 => self.miter_limit = Some(field.float()?),
                6 => self
                    .pattern
                    .get_or_insert_with(StrokePattern::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                7..=9 => {
                    field.bytes()?;
                    self.has_special = true;
                },
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        if let Some(color) = &self.color {
            writer.message(1, |writer| color.encode(writer))?;
        }
        optional!(writer, float, 2, self.width);
        optional!(writer, int32, 3, self.cap);
        optional!(writer, int32, 4, self.join);
        optional!(writer, float, 5, self.miter_limit);
        if let Some(pattern) = &self.pattern {
            writer.message(6, |writer| pattern.encode(writer))?;
        }
        Ok(())
    }
}

/// `TSD.ShadowArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(crate) struct Shadow {
    pub(crate) color: Option<Color>,
    pub(crate) angle: Option<f32>,
    pub(crate) offset: Option<f32>,
    pub(crate) radius: Option<i32>,
    pub(crate) opacity: Option<f32>,
    pub(crate) is_enabled: Option<bool>,
    pub(crate) kind: Option<i32>,
    pub(crate) drop: bool,
    pub(crate) contact: Option<(Option<f32>, Option<f32>)>,
    pub(crate) curved: Option<Option<f32>>,
}

impl Shadow {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.color.get_or_insert_with(Color::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                2 => self.angle = Some(field.float()?),
                3 => self.offset = Some(field.float()?),
                4 => self.radius = Some(field.int32()?),
                5 => self.opacity = Some(field.float()?),
                6 => self.is_enabled = Some(field.bool()?),
                7 => self.kind = Some(field.int32()?),
                8 => {
                    field.bytes()?;
                    self.drop = true;
                },
                9 => {
                    let contact = self.contact.get_or_insert((None, None));
                    fields(field.bytes()?, depth + 1, budget, |inner, _| {
                        match inner.number {
                            2 => contact.0 = Some(inner.float()?),
                            4 => contact.1 = Some(inner.float()?),
                            _ => {},
                        }
                        Ok(())
                    })?;
                },
                10 => {
                    let curved = self.curved.get_or_insert(None);
                    fields(field.bytes()?, depth + 1, budget, |inner, _| {
                        if inner.number == 1 {
                            *curved = Some(inner.float()?);
                        }
                        Ok(())
                    })?;
                },
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        if let Some(color) = &self.color {
            writer.message(1, |writer| color.encode(writer))?;
        }
        optional!(writer, float, 2, self.angle);
        optional!(writer, float, 3, self.offset);
        optional!(writer, int32, 4, self.radius);
        optional!(writer, float, 5, self.opacity);
        optional!(writer, bool, 6, self.is_enabled);
        optional!(writer, int32, 7, self.kind);
        if self.drop {
            writer.bytes(8, &[])?;
        }
        if let Some((height, offset)) = self.contact {
            writer.message(9, |writer| {
                optional!(writer, float, 2, height);
                optional!(writer, float, 4, offset);
                Ok(())
            })?;
        }
        if let Some(curve) = self.curved {
            writer.message(10, |writer| {
                optional!(writer, float, 1, curve);
                Ok(())
            })?;
        }
        Ok(())
    }
}

/// `TSD.FillArchive` (gradient and image fills are kept only as presence).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(crate) struct Fill {
    pub(crate) color: Option<Color>,
    pub(crate) has_gradient: bool,
    pub(crate) has_image: bool,
}

impl Fill {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.color.get_or_insert_with(Color::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                2 => {
                    field.bytes()?;
                    self.has_gradient = true;
                },
                3 => {
                    field.bytes()?;
                    self.has_image = true;
                },
                _ => {},
            }
            Ok(())
        })
    }

    fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        if let Some(color) = &self.color {
            writer.message(1, |writer| color.encode(writer))?;
        }
        Ok(())
    }
}

/// `TSWP.LineSpacingArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(crate) struct LineSpacing {
    pub(crate) mode: Option<i32>,
    pub(crate) amount: Option<f32>,
    pub(crate) baseline_rule: Option<f32>,
}

impl LineSpacing {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, _| {
            match field.number {
                1 => self.mode = Some(field.int32()?),
                2 => self.amount = Some(field.float()?),
                3 => self.baseline_rule = Some(field.float()?),
                _ => {},
            }
            Ok(())
        })
    }

    fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        optional!(writer, int32, 1, self.mode);
        optional!(writer, float, 2, self.amount);
        optional!(writer, float, 3, self.baseline_rule);
        Ok(())
    }
}

/// `TSWP.TabArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct Tab {
    pub(crate) position: Option<f32>,
    pub(crate) alignment: Option<i32>,
    pub(crate) leader: Option<String>,
}

/// `TSWP.TabsArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct Tabs {
    pub(crate) tabs: Vec<Tab>,
}

impl Tabs {
    fn merge(&mut self, data: &[u8], depth: usize, budget: &mut Budget) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            if field.number == 1 {
                let mut tab = Tab::default();
                fields(field.bytes()?, depth + 1, budget, |inner, budget| {
                    match inner.number {
                        1 => tab.position = Some(inner.float()?),
                        2 => tab.alignment = Some(inner.int32()?),
                        3 => tab.leader = Some(inner.string(budget)?),
                        _ => {},
                    }
                    Ok(())
                })?;
                reserve(&mut self.tabs, 1)?;
                self.tabs.push(tab);
            }
            Ok(())
        })
    }

    fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        for tab in &self.tabs {
            writer.message(1, |writer| {
                optional!(writer, float, 1, tab.position);
                optional!(writer, int32, 2, tab.alignment);
                if let Some(leader) = &tab.leader {
                    writer.bytes(3, leader.as_bytes())?;
                }
                Ok(())
            })?;
        }
        Ok(())
    }
}

/// The `TSWP.CharacterStylePropertiesArchive` fields the engines read or
/// write.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct CharacterProperties {
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) font_size: Option<f32>,
    pub(crate) font_name_null: Option<bool>,
    pub(crate) font_name: Option<String>,
    pub(crate) font_color_null: Option<bool>,
    pub(crate) font_color: Option<Color>,
    pub(crate) language_null: Option<bool>,
    pub(crate) language: Option<String>,
    pub(crate) superscript: Option<i32>,
    pub(crate) underline: Option<i32>,
    pub(crate) strikethru: Option<i32>,
    pub(crate) capitalization: Option<i32>,
    pub(crate) baseline_shift: Option<f32>,
    pub(crate) ligatures: Option<i32>,
    pub(crate) shadow_null: Option<bool>,
    pub(crate) shadow: Option<Shadow>,
    pub(crate) background_color_null: Option<bool>,
    pub(crate) background_color: Option<Color>,
    pub(crate) tracking: Option<f32>,
    pub(crate) writing_direction: Option<i32>,
    pub(crate) capitalization_uses_linguistics: Option<bool>,
    pub(crate) tsd_stroke_null: Option<bool>,
    pub(crate) tsd_stroke: Option<Stroke>,
    pub(crate) tsd_fill_null: Option<bool>,
    pub(crate) tsd_fill: Option<Fill>,
}

impl CharacterProperties {
    pub(crate) fn merge(
        &mut self,
        data: &[u8],
        depth: usize,
        budget: &mut Budget,
    ) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.bold = Some(field.bool()?),
                2 => self.italic = Some(field.bool()?),
                3 => self.font_size = Some(field.float()?),
                4 => self.font_name_null = Some(field.bool()?),
                5 => self.font_name = Some(field.string(budget)?),
                6 => self.font_color_null = Some(field.bool()?),
                7 => self.font_color.get_or_insert_with(Color::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                8 => self.language_null = Some(field.bool()?),
                9 => self.language = Some(field.string(budget)?),
                10 => self.superscript = Some(field.int32()?),
                11 => self.underline = Some(field.int32()?),
                12 => self.strikethru = Some(field.int32()?),
                13 => self.capitalization = Some(field.int32()?),
                14 => self.baseline_shift = Some(field.float()?),
                16 => self.ligatures = Some(field.int32()?),
                20 => self.shadow_null = Some(field.bool()?),
                21 => self.shadow.get_or_insert_with(Shadow::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                25 => self.background_color_null = Some(field.bool()?),
                26 => self
                    .background_color
                    .get_or_insert_with(Color::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                27 => self.tracking = Some(field.float()?),
                35 => self.writing_direction = Some(field.int32()?),
                41 => self.capitalization_uses_linguistics = Some(field.bool()?),
                43 => self.tsd_stroke_null = Some(field.bool()?),
                44 => self.tsd_stroke.get_or_insert_with(Stroke::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                45 => self.tsd_fill_null = Some(field.bool()?),
                46 => self.tsd_fill.get_or_insert_with(Fill::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        optional!(writer, bool, 1, self.bold);
        optional!(writer, bool, 2, self.italic);
        optional!(writer, float, 3, self.font_size);
        optional!(writer, bool, 4, self.font_name_null);
        if let Some(name) = &self.font_name {
            writer.bytes(5, name.as_bytes())?;
        }
        optional!(writer, bool, 6, self.font_color_null);
        if let Some(color) = &self.font_color {
            writer.message(7, |writer| color.encode(writer))?;
        }
        optional!(writer, bool, 8, self.language_null);
        if let Some(language) = &self.language {
            writer.bytes(9, language.as_bytes())?;
        }
        optional!(writer, int32, 10, self.superscript);
        optional!(writer, int32, 11, self.underline);
        optional!(writer, int32, 12, self.strikethru);
        optional!(writer, int32, 13, self.capitalization);
        optional!(writer, float, 14, self.baseline_shift);
        optional!(writer, int32, 16, self.ligatures);
        optional!(writer, bool, 20, self.shadow_null);
        if let Some(shadow) = &self.shadow {
            writer.message(21, |writer| shadow.encode(writer))?;
        }
        optional!(writer, bool, 25, self.background_color_null);
        if let Some(color) = &self.background_color {
            writer.message(26, |writer| color.encode(writer))?;
        }
        optional!(writer, float, 27, self.tracking);
        optional!(writer, int32, 35, self.writing_direction);
        optional!(writer, bool, 41, self.capitalization_uses_linguistics);
        optional!(writer, bool, 43, self.tsd_stroke_null);
        if let Some(stroke) = &self.tsd_stroke {
            writer.message(44, |writer| stroke.encode(writer))?;
        }
        optional!(writer, bool, 45, self.tsd_fill_null);
        if let Some(fill) = &self.tsd_fill {
            writer.message(46, |writer| fill.encode(writer))?;
        }
        Ok(())
    }
}

/// The `TSWP.ParagraphStylePropertiesArchive` fields the engines read or
/// write.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct ParagraphProperties {
    pub(crate) alignment: Option<i32>,
    pub(crate) decimal_tab_null: Option<bool>,
    pub(crate) decimal_tab: Option<String>,
    pub(crate) default_tab_stops: Option<f32>,
    pub(crate) fill_null: Option<bool>,
    pub(crate) fill: Option<Color>,
    pub(crate) first_line_indent: Option<f32>,
    pub(crate) hyphenate: Option<bool>,
    pub(crate) keep_lines_together: Option<bool>,
    pub(crate) keep_with_next: Option<bool>,
    pub(crate) left_indent: Option<f32>,
    pub(crate) line_spacing_null: Option<bool>,
    pub(crate) line_spacing: Option<LineSpacing>,
    pub(crate) page_break_before: Option<bool>,
    pub(crate) deprecated_borders: Option<i32>,
    pub(crate) historical_rule_offset_null: Option<bool>,
    pub(crate) historical_rule_offset: Option<Point>,
    pub(crate) rule_width: Option<f32>,
    pub(crate) right_indent: Option<f32>,
    pub(crate) space_after: Option<f32>,
    pub(crate) space_before: Option<f32>,
    pub(crate) tabs_null: Option<bool>,
    pub(crate) tabs: Option<Tabs>,
    pub(crate) widow_control: Option<bool>,
    pub(crate) stroke_null: Option<bool>,
    pub(crate) stroke: Option<Stroke>,
    pub(crate) list_style_null: Option<bool>,
    pub(crate) list_style: Option<Reference>,
    pub(crate) following_style_null: Option<bool>,
    pub(crate) following_style: Option<Reference>,
    pub(crate) border_positions: Option<i32>,
    pub(crate) rounded_corners: Option<bool>,
}

impl ParagraphProperties {
    pub(crate) fn merge(
        &mut self,
        data: &[u8],
        depth: usize,
        budget: &mut Budget,
    ) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.alignment = Some(field.int32()?),
                2 => self.decimal_tab_null = Some(field.bool()?),
                3 => self.decimal_tab = Some(field.string(budget)?),
                4 => self.default_tab_stops = Some(field.float()?),
                5 => self.fill_null = Some(field.bool()?),
                6 => self.fill.get_or_insert_with(Color::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                7 => self.first_line_indent = Some(field.float()?),
                8 => self.hyphenate = Some(field.bool()?),
                9 => self.keep_lines_together = Some(field.bool()?),
                10 => self.keep_with_next = Some(field.bool()?),
                11 => self.left_indent = Some(field.float()?),
                12 => self.line_spacing_null = Some(field.bool()?),
                13 => self
                    .line_spacing
                    .get_or_insert_with(LineSpacing::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                14 => self.page_break_before = Some(field.bool()?),
                15 => self.deprecated_borders = Some(field.int32()?),
                16 => self.historical_rule_offset_null = Some(field.bool()?),
                17 => self
                    .historical_rule_offset
                    .get_or_insert_with(Point::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                18 => self.rule_width = Some(field.float()?),
                19 => self.right_indent = Some(field.float()?),
                20 => self.space_after = Some(field.float()?),
                21 => self.space_before = Some(field.float()?),
                24 => self.tabs_null = Some(field.bool()?),
                25 => self.tabs.get_or_insert_with(Tabs::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                26 => self.widow_control = Some(field.bool()?),
                31 => self.stroke_null = Some(field.bool()?),
                32 => self.stroke.get_or_insert_with(Stroke::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                39 => self.list_style_null = Some(field.bool()?),
                40 => self
                    .list_style
                    .get_or_insert_with(Reference::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                41 => self.following_style_null = Some(field.bool()?),
                42 => self
                    .following_style
                    .get_or_insert_with(Reference::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                45 => self.border_positions = Some(field.int32()?),
                46 => self.rounded_corners = Some(field.bool()?),
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        optional!(writer, int32, 1, self.alignment);
        optional!(writer, bool, 2, self.decimal_tab_null);
        if let Some(tab) = &self.decimal_tab {
            writer.bytes(3, tab.as_bytes())?;
        }
        optional!(writer, float, 4, self.default_tab_stops);
        optional!(writer, bool, 5, self.fill_null);
        if let Some(color) = &self.fill {
            writer.message(6, |writer| color.encode(writer))?;
        }
        optional!(writer, float, 7, self.first_line_indent);
        optional!(writer, bool, 8, self.hyphenate);
        optional!(writer, bool, 9, self.keep_lines_together);
        optional!(writer, bool, 10, self.keep_with_next);
        optional!(writer, float, 11, self.left_indent);
        optional!(writer, bool, 12, self.line_spacing_null);
        if let Some(spacing) = &self.line_spacing {
            writer.message(13, |writer| spacing.encode(writer))?;
        }
        optional!(writer, bool, 14, self.page_break_before);
        optional!(writer, int32, 15, self.deprecated_borders);
        optional!(writer, bool, 16, self.historical_rule_offset_null);
        if let Some(point) = self.historical_rule_offset {
            writer.message(17, |writer| point.encode(writer))?;
        }
        optional!(writer, float, 18, self.rule_width);
        optional!(writer, float, 19, self.right_indent);
        optional!(writer, float, 20, self.space_after);
        optional!(writer, float, 21, self.space_before);
        optional!(writer, bool, 24, self.tabs_null);
        if let Some(tabs) = &self.tabs {
            writer.message(25, |writer| tabs.encode(writer))?;
        }
        optional!(writer, bool, 26, self.widow_control);
        optional!(writer, bool, 31, self.stroke_null);
        if let Some(stroke) = &self.stroke {
            writer.message(32, |writer| stroke.encode(writer))?;
        }
        optional!(writer, bool, 39, self.list_style_null);
        if let Some(reference) = self.list_style {
            writer.message(40, |writer| reference.encode(writer))?;
        }
        optional!(writer, bool, 41, self.following_style_null);
        if let Some(reference) = self.following_style {
            writer.message(42, |writer| reference.encode(writer))?;
        }
        optional!(writer, int32, 45, self.border_positions);
        optional!(writer, bool, 46, self.rounded_corners);
        Ok(())
    }
}

/// `TSS.StyleArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct StyleBase {
    pub(crate) name: Option<String>,
    pub(crate) style_identifier: Option<String>,
    pub(crate) parent: Option<Reference>,
    pub(crate) is_variation: Option<bool>,
    pub(crate) stylesheet: Option<Reference>,
}

impl StyleBase {
    pub(crate) fn merge(
        &mut self,
        data: &[u8],
        depth: usize,
        budget: &mut Budget,
    ) -> NativeResult<()> {
        fields(data, depth, budget, |field, budget| {
            match field.number {
                1 => self.name = Some(field.string(budget)?),
                2 => self.style_identifier = Some(field.string(budget)?),
                3 => self.parent.get_or_insert_with(Reference::default).merge(
                    field.bytes()?,
                    depth + 1,
                    budget,
                )?,
                4 => self.is_variation = Some(field.bool()?),
                5 => self
                    .stylesheet
                    .get_or_insert_with(Reference::default)
                    .merge(field.bytes()?, depth + 1, budget)?,
                _ => {},
            }
            Ok(())
        })
    }

    pub(crate) fn encode(&self, writer: &mut Writer) -> NativeResult<()> {
        if let Some(name) = &self.name {
            writer.bytes(1, name.as_bytes())?;
        }
        if let Some(identifier) = &self.style_identifier {
            writer.bytes(2, identifier.as_bytes())?;
        }
        if let Some(parent) = self.parent {
            writer.message(3, |writer| parent.encode(writer))?;
        }
        optional!(writer, bool, 4, self.is_variation);
        if let Some(stylesheet) = self.stylesheet {
            writer.message(5, |writer| stylesheet.encode(writer))?;
        }
        Ok(())
    }
}

/// `TSWP.ParagraphStyleArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct ParagraphStyle {
    pub(crate) base: StyleBase,
    pub(crate) override_count: Option<u32>,
    pub(crate) character: Option<CharacterProperties>,
    pub(crate) paragraph: Option<ParagraphProperties>,
}

impl ParagraphStyle {
    pub(crate) fn decode(data: &[u8], budget: &mut Budget) -> NativeResult<Self> {
        let mut style = Self::default();
        fields(data, 0, budget, |field, budget| {
            match field.number {
                1 => style.base.merge(field.bytes()?, 1, budget)?,
                10 => style.override_count = Some(field.uint32()?),
                11 => style
                    .character
                    .get_or_insert_with(CharacterProperties::default)
                    .merge(field.bytes()?, 1, budget)?,
                12 => style
                    .paragraph
                    .get_or_insert_with(ParagraphProperties::default)
                    .merge(field.bytes()?, 1, budget)?,
                _ => {},
            }
            Ok(())
        })?;
        Ok(style)
    }

    pub(crate) fn encode(&self, limits: WireLimits) -> NativeResult<Vec<u8>> {
        let mut writer = Writer::new(limits);
        writer.message(1, |writer| self.base.encode(writer))?;
        optional!(writer, uint32, 10, self.override_count);
        if let Some(character) = &self.character {
            writer.message(11, |writer| character.encode(writer))?;
        }
        if let Some(paragraph) = &self.paragraph {
            writer.message(12, |writer| paragraph.encode(writer))?;
        }
        Ok(writer.into_bytes())
    }
}
