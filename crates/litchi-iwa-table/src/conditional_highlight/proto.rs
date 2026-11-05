//! A prost-free mirror of the schema subset conditional highlights encode and
//! decode (record 0370d).
//!
//! The leaf has no protobuf runtime. These plain structs keep the generated
//! names the host editor's conditional-highlight code uses, so its node
//! builders port unchanged, and encode in field-number order exactly as the
//! generated encoder does: a required field is always written, an optional
//! one when it is set, a repeated one per element. Decoding reads the fields
//! these routes inspect; any other field of a message whose emptiness a
//! route checks marks it `other`, so a native graph carrying overrides the
//! routes do not model is refused rather than misread.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::struct_field_names,
    reason = "the mirror keeps the generated schema's names and order"
)]

use litchi_iwa_common::WireLimits;

/// A malformed payload for the mirrored schema.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DecodeError;

type Result<T> = std::result::Result<T, DecodeError>;

/// One encoded message under construction.
#[derive(Default)]
pub(super) struct Writer(Vec<u8>);

impl Writer {
    fn key(&mut self, number: u32, wire_type: u8) {
        self.raw_varint((u64::from(number) << 3) | u64::from(wire_type));
    }

    fn raw_varint(&mut self, mut value: u64) {
        while value >= 0x80 {
            self.0.push((value as u8) | 0x80);
            value >>= 7;
        }
        self.0.push(value as u8);
    }

    pub(super) fn uint(&mut self, number: u32, value: u64) {
        self.key(number, 0);
        self.raw_varint(value);
    }

    /// An `int32` (negative values sign-extend to ten bytes, as protobuf
    /// requires).
    pub(super) fn int32(&mut self, number: u32, value: i32) {
        self.uint(number, i64::from(value) as u64);
    }

    pub(super) fn boolean(&mut self, number: u32, value: bool) {
        self.uint(number, u64::from(value));
    }

    pub(super) fn double(&mut self, number: u32, value: f64) {
        self.key(number, 1);
        self.0.extend_from_slice(&value.to_le_bytes());
    }

    pub(super) fn float(&mut self, number: u32, value: f32) {
        self.key(number, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
    }

    pub(super) fn bytes(&mut self, number: u32, value: &[u8]) {
        self.key(number, 2);
        self.raw_varint(value.len() as u64);
        self.0.extend_from_slice(value);
    }

    pub(super) fn message(&mut self, number: u32, value: &impl Encode) {
        self.bytes(number, &value.encode_to_vec());
    }

    pub(super) fn into_bytes(self) -> Vec<u8> {
        self.0
    }
}

/// A mirrored message's encoder.
pub(super) trait Encode {
    fn encode(&self, writer: &mut Writer);

    fn encode_to_vec(&self) -> Vec<u8> {
        let mut writer = Writer::default();
        self.encode(&mut writer);
        writer.into_bytes()
    }
}

/// One decoded field.
pub(super) struct Field<'a> {
    pub(super) number: u32,
    wire_type: u8,
    payload: &'a [u8],
}

impl<'a> Field<'a> {
    pub(super) fn varint(&self) -> Result<u64> {
        if self.wire_type != 0 {
            return Err(DecodeError);
        }
        litchi_iwa_common::varint::decode_varint_from_bytes(self.payload)
            .map(|(value, _)| value)
            .map_err(|_error| DecodeError)
    }

    pub(super) fn uint32(&self) -> Result<u32> {
        u32::try_from(self.varint()?).map_err(|_error| DecodeError)
    }

    pub(super) fn int32(&self) -> Result<i32> {
        Ok(self.varint()? as i64 as i32)
    }

    pub(super) fn boolean(&self) -> Result<bool> {
        Ok(self.varint()? != 0)
    }

    pub(super) fn double(&self) -> Result<f64> {
        let bytes: [u8; 8] = (self.wire_type == 1)
            .then_some(self.payload)
            .and_then(|payload| payload.try_into().ok())
            .ok_or(DecodeError)?;
        Ok(f64::from_le_bytes(bytes))
    }

    pub(super) fn float(&self) -> Result<f32> {
        let bytes: [u8; 4] = (self.wire_type == 5)
            .then_some(self.payload)
            .and_then(|payload| payload.try_into().ok())
            .ok_or(DecodeError)?;
        Ok(f32::from_le_bytes(bytes))
    }

    pub(super) fn bytes(&self) -> Result<&'a [u8]> {
        if self.wire_type != 2 {
            return Err(DecodeError);
        }
        Ok(self.payload)
    }

    pub(super) fn string(&self) -> Result<String> {
        std::str::from_utf8(self.bytes()?)
            .map(str::to_owned)
            .map_err(|_error| DecodeError)
    }

    pub(super) fn message<M: Decode>(&self) -> Result<M> {
        M::decode(self.bytes()?)
    }
}

/// The fields of `data`, in order.
pub(super) fn fields(data: &[u8]) -> Result<Vec<Field<'_>>> {
    let view = litchi_iwa_common::wire::parse_wire_view(data, WireLimits::DEFAULT)
        .map_err(|_error| DecodeError)?;
    Ok(view
        .fields()
        .map(|field| Field {
            number: field.number(),
            wire_type: field.wire_type(),
            payload: field.payload(),
        })
        .collect())
}

/// A mirrored message's decoder.
pub(super) trait Decode: Sized {
    fn decode(data: &[u8]) -> Result<Self>;
}

pub(crate) mod tsp {
    use super::{Decode, DecodeError, Encode, Result, Writer, fields};

    #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
    pub(crate) struct Uuid {
        pub(crate) lower: u64,
        pub(crate) upper: u64,
    }

    impl Encode for Uuid {
        fn encode(&self, writer: &mut Writer) {
            writer.uint(1, self.lower);
            writer.uint(2, self.upper);
        }
    }

    impl Decode for Uuid {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut uuid = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => uuid.lower = field.varint()?,
                    2 => uuid.upper = field.varint()?,
                    _ => {},
                }
            }
            Ok(uuid)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq, Eq)]
    pub(crate) struct CfuuidArchive {
        pub(crate) uuid_bytes: Option<Vec<u8>>,
        pub(crate) uuid_w0: Option<u32>,
        pub(crate) uuid_w1: Option<u32>,
        pub(crate) uuid_w2: Option<u32>,
        pub(crate) uuid_w3: Option<u32>,
    }

    impl Encode for CfuuidArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(bytes) = &self.uuid_bytes {
                writer.bytes(1, bytes);
            }
            for (number, value) in [
                (2, self.uuid_w0),
                (3, self.uuid_w1),
                (4, self.uuid_w2),
                (5, self.uuid_w3),
            ] {
                if let Some(value) = value {
                    writer.uint(number, u64::from(value));
                }
            }
        }
    }

    impl Decode for CfuuidArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut uuid = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => uuid.uuid_bytes = Some(field.bytes()?.to_vec()),
                    2 => uuid.uuid_w0 = Some(field.uint32()?),
                    3 => uuid.uuid_w1 = Some(field.uint32()?),
                    4 => uuid.uuid_w2 = Some(field.uint32()?),
                    5 => uuid.uuid_w3 = Some(field.uint32()?),
                    _ => {},
                }
            }
            Ok(uuid)
        }
    }

    #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
    pub(crate) struct Reference {
        pub(crate) identifier: u64,
    }

    impl Encode for Reference {
        fn encode(&self, writer: &mut Writer) {
            writer.uint(1, self.identifier);
        }
    }

    impl Decode for Reference {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut reference = Self::default();
            for field in fields(data)? {
                if field.number == 1 {
                    reference.identifier = field.varint()?;
                }
            }
            Ok(reference)
        }
    }

    /// `TSP.Color` (RGB model only is produced).
    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct Color {
        pub(crate) model: i32,
        pub(crate) r: Option<f32>,
        pub(crate) g: Option<f32>,
        pub(crate) b: Option<f32>,
        pub(crate) a: Option<f32>,
        /// Any of the CMYK or white channels.
        pub(crate) other_channels: bool,
        pub(crate) rgbspace: Option<i32>,
    }

    impl Encode for Color {
        fn encode(&self, writer: &mut Writer) {
            writer.int32(1, self.model);
            for (number, value) in [(3, self.r), (4, self.g), (5, self.b), (6, self.a)] {
                if let Some(value) = value {
                    writer.float(number, value);
                }
            }
            if let Some(space) = self.rgbspace {
                writer.int32(12, space);
            }
        }
    }

    impl Decode for Color {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut color = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => color.model = field.int32()?,
                    3 => color.r = Some(field.float()?),
                    4 => color.g = Some(field.float()?),
                    5 => color.b = Some(field.float()?),
                    6 => color.a = Some(field.float()?),
                    7..=11 => color.other_channels = true,
                    12 => color.rgbspace = Some(field.int32()?),
                    _ => {},
                }
            }
            Ok(color)
        }
    }

    impl From<DecodeError> for crate::conditional_highlight::Error {
        fn from(_error: DecodeError) -> Self {
            Self::InvalidSource
        }
    }
}

pub(crate) mod tsd {
    use super::{Decode, Encode, Result, Writer, fields, tsp};

    /// `TSD.FillArchive`: a colour fill (gradients and images are only
    /// detected).
    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct FillArchive {
        pub(crate) color: Option<tsp::Color>,
        pub(crate) other: bool,
    }

    impl Encode for FillArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(color) = &self.color {
                writer.message(1, color);
            }
        }
    }

    impl Decode for FillArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut fill = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => fill.color = Some(field.message()?),
                    _ => fill.other = true,
                }
            }
            Ok(fill)
        }
    }
}

pub(crate) mod tswp {
    use super::{Decode, Encode, Result, Writer, fields, tsp};

    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct CharacterStylePropertiesArchive {
        pub(crate) bold: Option<bool>,
        pub(crate) font_color: Option<tsp::Color>,
        pub(crate) other: bool,
    }

    impl Encode for CharacterStylePropertiesArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(bold) = self.bold {
                writer.boolean(1, bold);
            }
            if let Some(color) = &self.font_color {
                writer.message(7, color);
            }
        }
    }

    impl Decode for CharacterStylePropertiesArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut properties = Self::default();
            let mut text_fill = None;
            for field in fields(data)? {
                match field.number {
                    1 => properties.bold = Some(field.boolean()?),
                    7 => properties.font_color = Some(field.message()?),
                    46 => text_fill = Some(field.message::<super::tsd::FillArchive>()?),
                    _ => properties.other = true,
                }
            }
            // Numbers' preset rule styles ("Red Fill", record 0432) repeat the
            // font colour as the text's colour fill: the same override, twice.
            if text_fill.is_some_and(|fill| fill.other || fill.color != properties.font_color) {
                properties.other = true;
            }
            Ok(properties)
        }
    }

    /// `TSWP.ParagraphStyleArchive` with an empty `super` style.
    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct ParagraphStyleArchive {
        pub(crate) override_count: Option<u32>,
        pub(crate) char_properties: Option<CharacterStylePropertiesArchive>,
        /// Whether `para_properties` is present and sets any field.
        pub(crate) para_properties_set: bool,
    }

    impl Encode for ParagraphStyleArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.bytes(1, &[]);
            if let Some(count) = self.override_count {
                writer.uint(10, u64::from(count));
            }
            if let Some(properties) = &self.char_properties {
                writer.message(11, properties);
            }
        }
    }

    impl Decode for ParagraphStyleArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut style = Self::default();
            for field in fields(data)? {
                match field.number {
                    10 => style.override_count = Some(field.uint32()?),
                    11 => style.char_properties = Some(field.message()?),
                    12 => style.para_properties_set |= !fields(field.bytes()?)?.is_empty(),
                    _ => {},
                }
            }
            Ok(style)
        }
    }
}

pub(crate) mod tst {
    use super::{Decode, Encode, Result, Writer, fields, tsce, tsd, tsp};

    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct CellStylePropertiesArchive {
        pub(crate) cell_fill: Option<tsd::FillArchive>,
        pub(crate) other: bool,
    }

    impl Encode for CellStylePropertiesArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(fill) = &self.cell_fill {
                writer.message(1, fill);
            }
        }
    }

    impl Decode for CellStylePropertiesArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut properties = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => properties.cell_fill = Some(field.message()?),
                    _ => properties.other = true,
                }
            }
            Ok(properties)
        }
    }

    /// `TST.CellStyleArchive` with an empty `super` style.
    #[derive(Debug, Clone, Copy, Default, PartialEq)]
    pub(crate) struct CellStyleArchive {
        pub(crate) override_count: Option<u32>,
        pub(crate) cell_properties: Option<CellStylePropertiesArchive>,
    }

    impl Encode for CellStyleArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.bytes(1, &[]);
            if let Some(count) = self.override_count {
                writer.uint(10, u64::from(count));
            }
            if let Some(properties) = &self.cell_properties {
                writer.message(11, properties);
            }
        }
    }

    impl Decode for CellStyleArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut style = Self::default();
            for field in fields(data)? {
                match field.number {
                    10 => style.override_count = Some(field.uint32()?),
                    11 => style.cell_properties = Some(field.message()?),
                    _ => {},
                }
            }
            Ok(style)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct FormulaPredArgDataArchive {
        pub(crate) double_value: Option<f64>,
        pub(crate) decimal_low: Option<u64>,
        pub(crate) decimal_high: Option<u64>,
        pub(crate) string_value: Option<String>,
        pub(crate) date_value: Option<f64>,
    }

    impl Encode for FormulaPredArgDataArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(value) = self.double_value {
                writer.double(1, value);
            }
            if let Some(value) = self.decimal_low {
                writer.uint(2, value);
            }
            if let Some(value) = self.decimal_high {
                writer.uint(3, value);
            }
            if let Some(value) = &self.string_value {
                writer.bytes(4, value.as_bytes());
            }
            if let Some(value) = self.date_value {
                writer.double(5, value);
            }
        }
    }

    impl Decode for FormulaPredArgDataArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut value = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => value.double_value = Some(field.double()?),
                    2 => value.decimal_low = Some(field.varint()?),
                    3 => value.decimal_high = Some(field.varint()?),
                    4 => value.string_value = Some(field.string()?),
                    5 => value.date_value = Some(field.double()?),
                    _ => {},
                }
            }
            Ok(value)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct FormulaPredArgArchive {
        pub(crate) arg_type: i32,
        pub(crate) arg_value: Option<FormulaPredArgDataArchive>,
        pub(crate) relative_cell_ref: Option<tsce::RelativeCellRefArchive>,
        pub(crate) preserve_row: Option<bool>,
        pub(crate) preserve_column: Option<bool>,
    }

    impl Encode for FormulaPredArgArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.int32(1, self.arg_type);
            if let Some(value) = &self.arg_value {
                writer.message(2, value);
            }
            if let Some(reference) = &self.relative_cell_ref {
                writer.message(4, reference);
            }
            if let Some(value) = self.preserve_row {
                writer.boolean(8, value);
            }
            if let Some(value) = self.preserve_column {
                writer.boolean(9, value);
            }
        }
    }

    impl Decode for FormulaPredArgArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut argument = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => argument.arg_type = field.int32()?,
                    2 => argument.arg_value = Some(field.message()?),
                    4 => argument.relative_cell_ref = Some(field.message()?),
                    8 => argument.preserve_row = Some(field.boolean()?),
                    9 => argument.preserve_column = Some(field.boolean()?),
                    _ => {},
                }
            }
            Ok(argument)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct FormulaPredicateArchive {
        pub(crate) predicate_type: i32,
        pub(crate) qualifier1: i32,
        pub(crate) qualifier2: i32,
        pub(crate) param_value0: Option<FormulaPredArgArchive>,
        pub(crate) param_value1: Option<FormulaPredArgArchive>,
        pub(crate) param_value2: Option<FormulaPredArgArchive>,
        pub(crate) formula: Option<tsce::FormulaArchive>,
        pub(crate) for_conditional_style: Option<bool>,
    }

    impl Encode for FormulaPredicateArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.int32(1, self.predicate_type);
            writer.int32(2, self.qualifier1);
            writer.int32(3, self.qualifier2);
            for (number, value) in [
                (4, &self.param_value0),
                (5, &self.param_value1),
                (6, &self.param_value2),
            ] {
                if let Some(value) = value {
                    writer.message(number, value);
                }
            }
            if let Some(formula) = &self.formula {
                writer.message(7, formula);
            }
            if let Some(value) = self.for_conditional_style {
                writer.boolean(8, value);
            }
        }
    }

    impl Decode for FormulaPredicateArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut predicate = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => predicate.predicate_type = field.int32()?,
                    2 => predicate.qualifier1 = field.int32()?,
                    3 => predicate.qualifier2 = field.int32()?,
                    4 => predicate.param_value0 = Some(field.message()?),
                    5 => predicate.param_value1 = Some(field.message()?),
                    6 => predicate.param_value2 = Some(field.message()?),
                    7 => predicate.formula = Some(field.message()?),
                    8 => predicate.for_conditional_style = Some(field.boolean()?),
                    _ => {},
                }
            }
            Ok(predicate)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct FormulaPredicatePrePivotArchive {
        pub(crate) formula: tsce::FormulaArchive,
        pub(crate) predicate_type: i32,
        pub(crate) qualifier1: i32,
        pub(crate) qualifier2: i32,
        pub(crate) param_index1: i32,
        pub(crate) param_index2: i32,
        pub(crate) param_index0: i32,
    }

    impl Encode for FormulaPredicatePrePivotArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.message(1, &self.formula);
            writer.int32(2, self.predicate_type);
            writer.int32(3, self.qualifier1);
            writer.int32(4, self.qualifier2);
            writer.int32(5, self.param_index1);
            writer.int32(6, self.param_index2);
            writer.int32(7, self.param_index0);
        }
    }

    impl Decode for FormulaPredicatePrePivotArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut predicate = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => predicate.formula = field.message()?,
                    2 => predicate.predicate_type = field.int32()?,
                    3 => predicate.qualifier1 = field.int32()?,
                    4 => predicate.qualifier2 = field.int32()?,
                    5 => predicate.param_index1 = field.int32()?,
                    6 => predicate.param_index2 = field.int32()?,
                    7 => predicate.param_index0 = field.int32()?,
                    _ => {},
                }
            }
            Ok(predicate)
        }
    }

    pub(crate) mod conditional_style_set_archive {
        use super::super::{Decode, Encode, Result, Writer, fields};
        use super::{FormulaPredicateArchive, FormulaPredicatePrePivotArchive, tsp};

        #[derive(Debug, Clone, Default, PartialEq)]
        pub(crate) struct ConditionalStyleRulePrePivot {
            pub(crate) predicate: FormulaPredicatePrePivotArchive,
            pub(crate) cell_style: tsp::Reference,
            pub(crate) text_style: tsp::Reference,
        }

        impl Encode for ConditionalStyleRulePrePivot {
            fn encode(&self, writer: &mut Writer) {
                writer.message(1, &self.predicate);
                writer.message(2, &self.cell_style);
                writer.message(3, &self.text_style);
            }
        }

        impl Decode for ConditionalStyleRulePrePivot {
            fn decode(data: &[u8]) -> Result<Self> {
                let mut rule = Self::default();
                for field in fields(data)? {
                    match field.number {
                        1 => rule.predicate = field.message()?,
                        2 => rule.cell_style = field.message()?,
                        3 => rule.text_style = field.message()?,
                        _ => {},
                    }
                }
                Ok(rule)
            }
        }

        #[derive(Debug, Clone, Default, PartialEq)]
        pub(crate) struct ConditionalStyleRule {
            pub(crate) predicate: Option<FormulaPredicateArchive>,
            pub(crate) cell_style: tsp::Reference,
            pub(crate) text_style: tsp::Reference,
        }

        impl Encode for ConditionalStyleRule {
            fn encode(&self, writer: &mut Writer) {
                if let Some(predicate) = &self.predicate {
                    writer.message(1, predicate);
                }
                writer.message(2, &self.cell_style);
                writer.message(3, &self.text_style);
            }
        }

        impl Decode for ConditionalStyleRule {
            fn decode(data: &[u8]) -> Result<Self> {
                let mut rule = Self::default();
                for field in fields(data)? {
                    match field.number {
                        1 => rule.predicate = Some(field.message()?),
                        2 => rule.cell_style = field.message()?,
                        3 => rule.text_style = field.message()?,
                        _ => {},
                    }
                }
                Ok(rule)
            }
        }

        #[derive(Debug, Clone, Default, PartialEq)]
        pub(crate) struct ConditionalStyleRules {
            pub(crate) rule: Vec<ConditionalStyleRule>,
        }

        impl Encode for ConditionalStyleRules {
            fn encode(&self, writer: &mut Writer) {
                for rule in &self.rule {
                    writer.message(1, rule);
                }
            }
        }

        impl Decode for ConditionalStyleRules {
            fn decode(data: &[u8]) -> Result<Self> {
                let mut rules = Self::default();
                for field in fields(data)? {
                    if field.number == 1 {
                        rules.rule.push(field.message()?);
                    }
                }
                Ok(rules)
            }
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct ConditionalStyleSetArchive {
        pub(crate) rule_count: u32,
        pub(crate) rules_prepivot: Vec<conditional_style_set_archive::ConditionalStyleRulePrePivot>,
        pub(crate) rules: Option<conditional_style_set_archive::ConditionalStyleRules>,
    }

    impl Encode for ConditionalStyleSetArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.uint(1, u64::from(self.rule_count));
            for rule in &self.rules_prepivot {
                writer.message(2, rule);
            }
            if let Some(rules) = &self.rules {
                writer.message(3, rules);
            }
        }
    }

    impl Decode for ConditionalStyleSetArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut set = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => set.rule_count = field.uint32()?,
                    2 => set.rules_prepivot.push(field.message()?),
                    3 => set.rules = Some(field.message()?),
                    _ => {},
                }
            }
            Ok(set)
        }
    }
}

pub(crate) mod tsce {
    use super::{Decode, Encode, Result, Writer, fields, tsp};

    #[derive(Debug, Clone, Default, PartialEq, Eq)]
    pub(crate) struct RelativeCellRefArchive {
        pub(crate) relative_row_offset: Option<i32>,
        pub(crate) relative_column_offset: Option<i32>,
        pub(crate) table_uid: Option<tsp::Uuid>,
        pub(crate) preserve_column: Option<bool>,
        pub(crate) preserve_row: Option<bool>,
        pub(crate) is_spanning_column: Option<bool>,
        pub(crate) is_spanning_row: Option<bool>,
    }

    impl Encode for RelativeCellRefArchive {
        fn encode(&self, writer: &mut Writer) {
            if let Some(value) = self.relative_row_offset {
                writer.int32(1, value);
            }
            if let Some(value) = self.relative_column_offset {
                writer.int32(2, value);
            }
            if let Some(uid) = &self.table_uid {
                writer.message(3, uid);
            }
            for (number, value) in [
                (4, self.preserve_column),
                (5, self.preserve_row),
                (6, self.is_spanning_column),
                (7, self.is_spanning_row),
            ] {
                if let Some(value) = value {
                    writer.boolean(number, value);
                }
            }
        }
    }

    impl Decode for RelativeCellRefArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut reference = Self::default();
            for field in fields(data)? {
                match field.number {
                    1 => reference.relative_row_offset = Some(field.int32()?),
                    2 => reference.relative_column_offset = Some(field.int32()?),
                    3 => reference.table_uid = Some(field.message()?),
                    4 => reference.preserve_column = Some(field.boolean()?),
                    5 => reference.preserve_row = Some(field.boolean()?),
                    6 => reference.is_spanning_column = Some(field.boolean()?),
                    7 => reference.is_spanning_row = Some(field.boolean()?),
                    _ => {},
                }
            }
            Ok(reference)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct FormulaArchive {
        pub(crate) ast_node_array: AstNodeArrayArchive,
    }

    impl Encode for FormulaArchive {
        fn encode(&self, writer: &mut Writer) {
            writer.message(1, &self.ast_node_array);
        }
    }

    impl Decode for FormulaArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut formula = Self::default();
            for field in fields(data)? {
                if field.number == 1 {
                    formula.ast_node_array = field.message()?;
                }
            }
            Ok(formula)
        }
    }

    #[derive(Debug, Clone, Default, PartialEq)]
    pub(crate) struct AstNodeArrayArchive {
        pub(crate) ast_node: Vec<ast_node_array_archive::AstNodeArchive>,
    }

    impl Encode for AstNodeArrayArchive {
        fn encode(&self, writer: &mut Writer) {
            for node in &self.ast_node {
                writer.message(1, node);
            }
        }
    }

    impl Decode for AstNodeArrayArchive {
        fn decode(data: &[u8]) -> Result<Self> {
            let mut array = Self::default();
            for field in fields(data)? {
                if field.number == 1 {
                    array.ast_node.push(field.message()?);
                }
            }
            Ok(array)
        }
    }

    pub(crate) mod ast_node_array_archive {
        use super::super::{Decode, Encode, Result, Writer, fields};
        use super::tsp;

        /// `TSCE.ASTNodeArrayArchive.ASTNodeType`, as the generated enum
        /// names the values these routes use.
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        #[repr(i32)]
        #[allow(clippy::enum_variant_names, reason = "the generated enum's names")]
        pub(crate) enum AstNodeType {
            AdditionNode = 1,
            SubtractionNode = 2,
            MultiplicationNode = 3,
            GreaterThanNode = 7,
            GreaterThanOrEqualToNode = 8,
            LessThanNode = 9,
            LessThanOrEqualToNode = 10,
            EqualToNode = 11,
            NotEqualToNode = 12,
            FunctionNode = 16,
            NumberNode = 17,
            BooleanNode = 18,
            StringNode = 19,
            DateNode = 20,
            DurationNode = 21,
            BeginThunkNode = 34,
            EndThunkNode = 35,
            LinkedCellRefNode = 63,
        }

        #[derive(Debug, Clone, Default, PartialEq, Eq)]
        pub(crate) struct AstCrossTableReferenceExtraInfoArchive {
            pub(crate) table_id: tsp::CfuuidArchive,
        }

        impl Encode for AstCrossTableReferenceExtraInfoArchive {
            fn encode(&self, writer: &mut Writer) {
                writer.message(1, &self.table_id);
            }
        }

        impl Decode for AstCrossTableReferenceExtraInfoArchive {
            fn decode(data: &[u8]) -> Result<Self> {
                let mut info = Self::default();
                for field in fields(data)? {
                    if field.number == 1 {
                        info.table_id = field.message()?;
                    }
                }
                Ok(info)
            }
        }

        /// One AST node; `other` records any field these routes do not
        /// write.
        #[derive(Debug, Clone, Default, PartialEq)]
        pub(crate) struct AstNodeArchive {
            pub(crate) ast_node_type: i32,
            pub(crate) ast_function_node_index: Option<u32>,
            pub(crate) ast_function_node_num_args: Option<u32>,
            pub(crate) ast_number_node_number: Option<f64>,
            pub(crate) ast_boolean_node_boolean: Option<bool>,
            pub(crate) ast_string_node_string: Option<String>,
            pub(crate) ast_date_node_date_num: Option<f64>,
            pub(crate) ast_duration_node_unit_num: Option<f64>,
            pub(crate) ast_duration_node_unit: Option<i32>,
            pub(crate) ast_number_node_decimal_low: Option<u64>,
            pub(crate) ast_number_node_decimal_high: Option<u64>,
            pub(crate) ast_date_node_suppress_date_format: Option<bool>,
            pub(crate) ast_date_node_suppress_time_format: Option<bool>,
            pub(crate) ast_date_node_date_time_format: Option<String>,
            pub(crate) ast_duration_node_style: Option<u32>,
            pub(crate) ast_duration_node_duration_unit_largest: Option<u32>,
            pub(crate) ast_duration_node_duration_unit_smallest: Option<u32>,
            pub(crate) ast_cross_table_reference_extra_info:
                Option<AstCrossTableReferenceExtraInfoArchive>,
            pub(crate) ast_duration_node_use_automatic_units: Option<bool>,
            pub(crate) other: bool,
        }

        impl Encode for AstNodeArchive {
            fn encode(&self, writer: &mut Writer) {
                writer.int32(1, self.ast_node_type);
                if let Some(value) = self.ast_function_node_index {
                    writer.uint(2, u64::from(value));
                }
                if let Some(value) = self.ast_function_node_num_args {
                    writer.uint(3, u64::from(value));
                }
                if let Some(value) = self.ast_number_node_number {
                    writer.double(4, value);
                }
                if let Some(value) = self.ast_boolean_node_boolean {
                    writer.boolean(5, value);
                }
                if let Some(value) = &self.ast_string_node_string {
                    writer.bytes(6, value.as_bytes());
                }
                if let Some(value) = self.ast_date_node_date_num {
                    writer.double(7, value);
                }
                if let Some(value) = self.ast_duration_node_unit_num {
                    writer.double(8, value);
                }
                if let Some(value) = self.ast_duration_node_unit {
                    writer.int32(9, value);
                }
                if let Some(value) = self.ast_date_node_suppress_date_format {
                    writer.boolean(19, value);
                }
                if let Some(value) = self.ast_date_node_suppress_time_format {
                    writer.boolean(20, value);
                }
                if let Some(value) = &self.ast_date_node_date_time_format {
                    writer.bytes(21, value.as_bytes());
                }
                if let Some(value) = self.ast_duration_node_style {
                    writer.uint(22, u64::from(value));
                }
                if let Some(value) = self.ast_duration_node_duration_unit_largest {
                    writer.uint(23, u64::from(value));
                }
                if let Some(value) = self.ast_duration_node_duration_unit_smallest {
                    writer.uint(24, u64::from(value));
                }
                if let Some(info) = &self.ast_cross_table_reference_extra_info {
                    writer.message(28, info);
                }
                if let Some(value) = self.ast_duration_node_use_automatic_units {
                    writer.boolean(29, value);
                }
                if let Some(value) = self.ast_number_node_decimal_low {
                    writer.uint(42, value);
                }
                if let Some(value) = self.ast_number_node_decimal_high {
                    writer.uint(43, value);
                }
            }
        }

        impl Decode for AstNodeArchive {
            fn decode(data: &[u8]) -> Result<Self> {
                let mut node = Self::default();
                for field in fields(data)? {
                    match field.number {
                        1 => node.ast_node_type = field.int32()?,
                        2 => node.ast_function_node_index = Some(field.uint32()?),
                        3 => node.ast_function_node_num_args = Some(field.uint32()?),
                        4 => node.ast_number_node_number = Some(field.double()?),
                        5 => node.ast_boolean_node_boolean = Some(field.boolean()?),
                        6 => node.ast_string_node_string = Some(field.string()?),
                        7 => node.ast_date_node_date_num = Some(field.double()?),
                        8 => node.ast_duration_node_unit_num = Some(field.double()?),
                        9 => node.ast_duration_node_unit = Some(field.int32()?),
                        19 => node.ast_date_node_suppress_date_format = Some(field.boolean()?),
                        20 => node.ast_date_node_suppress_time_format = Some(field.boolean()?),
                        21 => node.ast_date_node_date_time_format = Some(field.string()?),
                        22 => node.ast_duration_node_style = Some(field.uint32()?),
                        23 => node.ast_duration_node_duration_unit_largest = Some(field.uint32()?),
                        24 => {
                            node.ast_duration_node_duration_unit_smallest = Some(field.uint32()?);
                        },
                        28 => node.ast_cross_table_reference_extra_info = Some(field.message()?),
                        29 => node.ast_duration_node_use_automatic_units = Some(field.boolean()?),
                        42 => node.ast_number_node_decimal_low = Some(field.varint()?),
                        43 => node.ast_number_node_decimal_high = Some(field.varint()?),
                        _ => node.other = true,
                    }
                }
                Ok(node)
            }
        }
    }
}
