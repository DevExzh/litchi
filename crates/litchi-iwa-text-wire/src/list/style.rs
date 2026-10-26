//! The `TSWP.ListStyleArchive` (2023) model: decode, canonical presets,
//! variations and in-place level-array patches.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_varint_field, rewrite_repeated_fixed32_fields,
    rewrite_repeated_length_delimited_fields, rewrite_repeated_varint_fields,
};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLabelColor, ParagraphListNumberFormat,
    ParagraphListNumberPunctuation, ParagraphListNumberSequence,
};

use crate::native::{self, Budget, Field, NativeError, NativeResult, Writer};
use crate::values::{color_from_native, color_to_native};

pub(super) const LEVELS: usize = 9;
const OVERRIDE_COUNT_FIELD: u32 = 10;
const LABEL_TYPES_FIELD: u32 = 11;
const TEXT_INDENTS_FIELD: u32 = 12;
const INDENTS_FIELD: u32 = 13;
const GEOMETRIES_FIELD: u32 = 14;
const NUMBER_TYPES_FIELD: u32 = 15;
const STRINGS_FIELD: u32 = 16;
const IMAGES_FIELD: u32 = 17;
const SHADOW_NULL_FIELD: u32 = 18;
const SHADOW_FIELD: u32 = 19;
const FONT_COLOR_NULL_FIELD: u32 = 20;
const FONT_COLOR_FIELD: u32 = 21;
const FONT_NAME_NULL_FIELD: u32 = 22;
const FONT_NAME_FIELD: u32 = 23;
const WRITING_DIRECTION_FIELD: u32 = 24;
const TIERED_NUMBERS_FIELD: u32 = 25;
const FONT_EM_POINTS: f32 = 11.0;
const NONE_INDENT_STEP_POINTS: f32 = 36.0;
const BULLET_INDENT_STEP_POINTS: f32 = 9.0;
const NUMBER_INDENT_STEP_POINTS: f32 = 18.0;
const BULLET_BASELINE_OFFSET_POINTS: f32 = -1.0;
pub(super) const DEFAULT_LABEL_SCALE: f32 = 1.0;
const BULLET_GLYPH: &str = "•";
const LABEL_NONE: i32 = 0;
const LABEL_STRING: i32 = 2;
const LABEL_NUMBER: i32 = 3;
const NUMBER_DECIMAL: i32 = 0;
const NUMBER_CIRCLED: i32 = 48;
const NUMBER_HEBREW_BIBLICAL_STANDARD: i32 = 61;
const NUMBER_TYPE_MAX: i32 = 64;

/// `TSWP.ListStyleArchive.LabelGeometry`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(super) struct Geometry {
    pub(super) scale: Option<f32>,
    pub(super) baseline_offset: Option<f32>,
    pub(super) scale_with_text: Option<bool>,
}

impl Geometry {
    fn encode(self, writer: &mut Writer) -> NativeResult<()> {
        if let Some(scale) = self.scale {
            writer.float(1, scale)?;
        }
        if let Some(offset) = self.baseline_offset {
            writer.float(2, offset)?;
        }
        if let Some(scale_with_text) = self.scale_with_text {
            writer.bool(3, scale_with_text)?;
        }
        Ok(())
    }
}

/// The fields of a list style the engine reads.
#[derive(Debug, Clone, PartialEq, Default)]
pub(super) struct ListStyle {
    pub(super) base: native::StyleBase,
    pub(super) override_count: Option<u32>,
    pub(super) label_types: Vec<i32>,
    pub(super) text_indents: Vec<f32>,
    pub(super) indents: Vec<f32>,
    pub(super) geometries: Vec<Geometry>,
    pub(super) number_types: Vec<i32>,
    pub(super) strings: Vec<String>,
    pub(super) images: usize,
    pub(super) shadow_null: Option<bool>,
    pub(super) shadow: bool,
    pub(super) font_color_null: Option<bool>,
    pub(super) font_color: Option<native::Color>,
    pub(super) font_name_null: Option<bool>,
    pub(super) font_name: Option<String>,
    pub(super) writing_direction: Option<i32>,
    pub(super) tiered_numbers: Vec<bool>,
}

fn varints(field: Field<'_>, output: &mut Vec<i32>) -> NativeResult<()> {
    if let Ok(packed) = field.bytes() {
        let mut offset = 0;
        while offset < packed.len() {
            let (value, length) =
                litchi_iwa_common::varint::decode_varint_from_bytes(&packed[offset..])
                    .map_err(|_| NativeError::Invalid("malformed packed protobuf varints"))?;
            native::reserve(output, 1)?;
            #[allow(clippy::cast_possible_truncation, reason = "protobuf int32 decoding")]
            output.push(value as i32);
            offset += length;
        }
        return Ok(());
    }
    native::reserve(output, 1)?;
    output.push(field.int32()?);
    Ok(())
}

impl ListStyle {
    pub(super) fn decode(data: &[u8], budget: &mut Budget) -> NativeResult<Self> {
        let mut style = Self::default();
        native::fields(data, 0, budget, |field, budget| {
            match field.number {
                1 => style.base.merge(field.bytes()?, 1, budget)?,
                OVERRIDE_COUNT_FIELD => style.override_count = Some(field.uint32()?),
                LABEL_TYPES_FIELD => varints(field, &mut style.label_types)?,
                TEXT_INDENTS_FIELD => field.floats(&mut style.text_indents)?,
                INDENTS_FIELD => field.floats(&mut style.indents)?,
                GEOMETRIES_FIELD => {
                    let mut geometry = Geometry::default();
                    native::fields(field.bytes()?, 1, budget, |inner, _| {
                        match inner.number {
                            1 => geometry.scale = Some(inner.float()?),
                            2 => geometry.baseline_offset = Some(inner.float()?),
                            3 => geometry.scale_with_text = Some(inner.bool()?),
                            _ => {},
                        }
                        Ok(())
                    })?;
                    native::reserve(&mut style.geometries, 1)?;
                    style.geometries.push(geometry);
                },
                NUMBER_TYPES_FIELD => varints(field, &mut style.number_types)?,
                STRINGS_FIELD => {
                    let value = field.string(budget)?;
                    native::reserve(&mut style.strings, 1)?;
                    style.strings.push(value);
                },
                IMAGES_FIELD => {
                    field.bytes()?;
                    style.images += 1;
                },
                SHADOW_NULL_FIELD => style.shadow_null = Some(field.bool()?),
                SHADOW_FIELD => {
                    field.bytes()?;
                    style.shadow = true;
                },
                FONT_COLOR_NULL_FIELD => style.font_color_null = Some(field.bool()?),
                FONT_COLOR_FIELD => {
                    let mut color = style.font_color.unwrap_or_default();
                    color.merge(field.bytes()?, 1, budget)?;
                    style.font_color = Some(color);
                },
                FONT_NAME_NULL_FIELD => style.font_name_null = Some(field.bool()?),
                FONT_NAME_FIELD => style.font_name = Some(field.string(budget)?),
                WRITING_DIRECTION_FIELD => style.writing_direction = Some(field.int32()?),
                TIERED_NUMBERS_FIELD => {
                    let mut values = Vec::new();
                    varints(field, &mut values)?;
                    native::reserve(&mut style.tiered_numbers, values.len())?;
                    style
                        .tiered_numbers
                        .extend(values.into_iter().map(|value| value != 0));
                },
                _ => {},
            }
            Ok(())
        })?;
        Ok(style)
    }

    pub(super) fn parent(&self) -> Option<u64> {
        self.base
            .parent
            .map(|parent| parent.identifier)
            .filter(|identifier| *identifier != 0)
    }

    fn encode_arrays(&self, writer: &mut Writer) -> NativeResult<()> {
        optional_uint32(writer, OVERRIDE_COUNT_FIELD, self.override_count)?;
        for value in &self.label_types {
            writer.int32(LABEL_TYPES_FIELD, *value)?;
        }
        for value in &self.text_indents {
            writer.float(TEXT_INDENTS_FIELD, *value)?;
        }
        for value in &self.indents {
            writer.float(INDENTS_FIELD, *value)?;
        }
        for geometry in &self.geometries {
            writer.message(GEOMETRIES_FIELD, |writer| geometry.encode(writer))?;
        }
        for value in &self.number_types {
            writer.int32(NUMBER_TYPES_FIELD, *value)?;
        }
        for value in &self.strings {
            writer.bytes(STRINGS_FIELD, value.as_bytes())?;
        }
        if let Some(null) = self.font_color_null {
            writer.bool(FONT_COLOR_NULL_FIELD, null)?;
        }
        if let Some(color) = &self.font_color {
            writer.message(FONT_COLOR_FIELD, |writer| color.encode(writer))?;
        }
        for value in &self.tiered_numbers {
            writer.bool(TIERED_NUMBERS_FIELD, *value)?;
        }
        Ok(())
    }

    pub(super) fn encode(&self, limits: WireLimits) -> NativeResult<Vec<u8>> {
        let mut writer = Writer::new(limits);
        writer.message(1, |writer| self.base.encode(writer))?;
        self.encode_arrays(&mut writer)?;
        Ok(writer.into_bytes())
    }
}

fn optional_uint32(writer: &mut Writer, number: u32, value: Option<u32>) -> NativeResult<()> {
    if let Some(value) = value {
        writer.uint32(number, value)?;
    }
    Ok(())
}

fn level_indents(step: f32) -> Vec<f32> {
    #[allow(clippy::cast_precision_loss, reason = "nine small level indexes")]
    (0..LEVELS).map(|level| level as f32 * step).collect()
}

fn repeated_geometry(baseline: impl Fn(usize) -> f32) -> Vec<Geometry> {
    (0..LEVELS)
        .map(|level| Geometry {
            scale: Some(DEFAULT_LABEL_SCALE),
            baseline_offset: Some(baseline(level)),
            scale_with_text: Some(true),
        })
        .collect()
}

pub(super) fn canonical(preset: ParagraphList) -> ListStyle {
    match preset {
        ParagraphList::None => ListStyle {
            override_count: Some(4),
            label_types: vec![LABEL_NONE; LEVELS],
            text_indents: vec![0.0; LEVELS],
            indents: level_indents(NONE_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|_| 0.0),
            ..ListStyle::default()
        },
        ParagraphList::Bullet => ListStyle {
            override_count: Some(5),
            label_types: vec![LABEL_STRING; LEVELS],
            text_indents: vec![BULLET_INDENT_STEP_POINTS / FONT_EM_POINTS; LEVELS],
            indents: level_indents(BULLET_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|level| {
                if level == 0 {
                    0.0
                } else {
                    BULLET_BASELINE_OFFSET_POINTS
                }
            }),
            strings: vec![BULLET_GLYPH.to_owned(); LEVELS],
            ..ListStyle::default()
        },
        ParagraphList::Numbered => ListStyle {
            override_count: Some(6),
            label_types: vec![LABEL_NUMBER; LEVELS],
            text_indents: vec![NUMBER_INDENT_STEP_POINTS / FONT_EM_POINTS; LEVELS],
            indents: level_indents(NUMBER_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|_| 0.0),
            number_types: vec![NUMBER_DECIMAL; LEVELS],
            tiered_numbers: vec![false; LEVELS],
            ..ListStyle::default()
        },
    }
}

pub(super) fn matches_preset(style: &ListStyle, preset: ParagraphList) -> bool {
    if preset == ParagraphList::None
        && style.base.is_variation != Some(true)
        && (style.label_types.is_empty()
            || style.label_types.iter().all(|label| *label == LABEL_NONE))
    {
        return style.strings.is_empty() && style.number_types.is_empty() && style.images == 0;
    }
    let expected = canonical(preset);
    style.override_count == expected.override_count
        && style.label_types == expected.label_types
        && style.text_indents == expected.text_indents
        && style.indents == expected.indents
        && style.geometries == expected.geometries
        && style.number_types == expected.number_types
        && style.strings == expected.strings
        && style.images == expected.images
        && style.shadow_null == expected.shadow_null
        && style.shadow == expected.shadow
        && style.font_color_null == expected.font_color_null
        && style.font_color == expected.font_color
        && style.font_name_null == expected.font_name_null
        && style.font_name == expected.font_name
        && style.writing_direction == expected.writing_direction
        && style.tiered_numbers == expected.tiered_numbers
}

pub(super) fn preset_of(style: &ListStyle) -> Option<ParagraphList> {
    [
        ParagraphList::None,
        ParagraphList::Bullet,
        ParagraphList::Numbered,
    ]
    .into_iter()
    .find(|preset| matches_preset(style, *preset))
}

pub(super) const fn preset_name(preset: ParagraphList) -> &'static str {
    match preset {
        ParagraphList::None => "None",
        ParagraphList::Bullet => "Bullet",
        ParagraphList::Numbered => "Numbered",
    }
}

/// The named root preset object of a stylesheet.
pub(super) fn preset_payload(
    stylesheet: u64,
    preset: ParagraphList,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    let mut style = canonical(preset);
    style.base = native::StyleBase {
        name: Some(preset_name(preset).to_owned()),
        stylesheet: Some(native::Reference {
            identifier: stylesheet,
        }),
        ..native::StyleBase::default()
    };
    style.encode(limits)
}

/// A one-property variation of `parent`.
pub(super) fn variation(
    parent: u64,
    stylesheet: u64,
    mut overrides: ListStyle,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    overrides.base = native::StyleBase {
        parent: Some(native::Reference { identifier: parent }),
        is_variation: Some(true),
        stylesheet: Some(native::Reference {
            identifier: stylesheet,
        }),
        ..native::StyleBase::default()
    };
    overrides.encode(limits)
}

pub(super) fn validate_strings(strings: &[String]) -> NativeResult<()> {
    if strings.len() != LEVELS || strings.iter().any(String::is_empty) {
        return Err(NativeError::Invalid(
            "text-bullet list style must define nine nonempty markers",
        ));
    }
    Ok(())
}

pub(super) fn validate_geometries(geometries: &[Geometry]) -> NativeResult<()> {
    if geometries.len() != LEVELS {
        return Err(NativeError::Invalid(
            "list style must define nine geometries",
        ));
    }
    for geometry in geometries {
        let scale = geometry.scale.unwrap_or(DEFAULT_LABEL_SCALE);
        let baseline = geometry.baseline_offset.unwrap_or(0.0);
        if !scale.is_finite() || scale <= 0.0 || !baseline.is_finite() {
            return Err(NativeError::Invalid("list style contains invalid geometry"));
        }
        if geometry.scale_with_text == Some(false) {
            return Err(NativeError::Invalid(
                "list style contains an absolute-size geometry",
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_floats(values: &[f32]) -> NativeResult<()> {
    if values.len() != LEVELS
        || values
            .iter()
            .any(|value| !value.is_finite() || *value < 0.0)
    {
        return Err(NativeError::Invalid(
            "list style must define nine finite nonnegative indents",
        ));
    }
    Ok(())
}

pub(super) fn validate_number_types(values: &[i32]) -> NativeResult<()> {
    if values.len() != LEVELS {
        return Err(NativeError::Invalid(
            "numbered list style must define nine formats",
        ));
    }
    for value in values {
        number_format_from_native(*value)?;
    }
    Ok(())
}

pub(super) fn validate_tiered(values: &[bool]) -> NativeResult<()> {
    if values.len() != LEVELS {
        return Err(NativeError::Invalid(
            "numbered list style must define nine tiering values",
        ));
    }
    Ok(())
}

const fn sequence_base(sequence: ParagraphListNumberSequence) -> i32 {
    match sequence {
        ParagraphListNumberSequence::Decimal => 0,
        ParagraphListNumberSequence::RomanUppercase => 3,
        ParagraphListNumberSequence::RomanLowercase => 6,
        ParagraphListNumberSequence::LatinUppercase => 9,
        ParagraphListNumberSequence::LatinLowercase => 12,
        ParagraphListNumberSequence::JapaneseIdeographic => 15,
        ParagraphListNumberSequence::JapaneseHiragana => 18,
        ParagraphListNumberSequence::JapaneseKatakana => 21,
        ParagraphListNumberSequence::JapaneseHiraganaIroha => 24,
        ParagraphListNumberSequence::JapaneseKatakanaIroha => 27,
        ParagraphListNumberSequence::SimplifiedChineseIdeographic => 30,
        ParagraphListNumberSequence::TraditionalChineseIdeographic => 33,
        ParagraphListNumberSequence::FormalJapaneseIdeographic => 36,
        ParagraphListNumberSequence::FormalSimplifiedChineseIdeographic => 39,
        ParagraphListNumberSequence::FormalTraditionalChineseIdeographic => 42,
        ParagraphListNumberSequence::KoreanAlphabet => 45,
        ParagraphListNumberSequence::ArabicIndic => 49,
        ParagraphListNumberSequence::ArabicAlphabet => 52,
        ParagraphListNumberSequence::ArabicAbjad => 55,
        ParagraphListNumberSequence::HebrewAlphabet => 58,
        ParagraphListNumberSequence::HebrewBiblical => 62,
    }
}

pub(super) const fn number_format_to_native(format: ParagraphListNumberFormat) -> i32 {
    match format {
        ParagraphListNumberFormat::Circled => NUMBER_CIRCLED,
        ParagraphListNumberFormat::HebrewBiblicalStandard => NUMBER_HEBREW_BIBLICAL_STANDARD,
        ParagraphListNumberFormat::Affixed {
            sequence,
            punctuation,
        } => {
            sequence_base(sequence)
                + match punctuation {
                    ParagraphListNumberPunctuation::Period => 0,
                    ParagraphListNumberPunctuation::Parentheses => 1,
                    ParagraphListNumberPunctuation::RightParenthesis => 2,
                }
        },
    }
}

pub(super) fn number_format_from_native(value: i32) -> NativeResult<ParagraphListNumberFormat> {
    if !(0..=NUMBER_TYPE_MAX).contains(&value) {
        return Err(NativeError::Invalid(
            "native numbered-list format is unknown",
        ));
    }
    if value == NUMBER_CIRCLED {
        return Ok(ParagraphListNumberFormat::Circled);
    }
    if value == NUMBER_HEBREW_BIBLICAL_STANDARD {
        return Ok(ParagraphListNumberFormat::HebrewBiblicalStandard);
    }
    for sequence in ParagraphListNumberSequence::ALL {
        for punctuation in ParagraphListNumberPunctuation::ALL {
            let format = ParagraphListNumberFormat::affixed(sequence, punctuation);
            if number_format_to_native(format) == value {
                return Ok(format);
            }
        }
    }
    Err(NativeError::Invalid(
        "native numbered-list format has no public representation",
    ))
}

pub(super) fn label_color_from_native(
    style: &ListStyle,
) -> NativeResult<Option<ParagraphListLabelColor>> {
    if style.font_color_null == Some(true) {
        return Ok(Some(ParagraphListLabelColor::Automatic));
    }
    if let Some(color) = style.font_color.as_ref() {
        return Ok(Some(ParagraphListLabelColor::Explicit(color_from_native(
            color,
        )?)));
    }
    if style.font_color_null == Some(false) {
        return Err(NativeError::Invalid(
            "list style enables a missing label color",
        ));
    }
    Ok(None)
}

pub(super) fn label_color_to_native(
    color: ParagraphListLabelColor,
) -> (Option<bool>, Option<native::Color>) {
    match color {
        ParagraphListLabelColor::Automatic => (Some(true), None),
        ParagraphListLabelColor::Explicit(color) => (None, Some(color_to_native(color))),
    }
}

/// The property arrays a list-style patch writes.
#[derive(Debug, Clone, PartialEq)]
pub(super) enum Arrays {
    Strings(Vec<String>),
    Geometries(Vec<Geometry>),
    Indentation(Vec<f32>, Vec<f32>),
    LabelColor(ParagraphListLabelColor),
    NumberTypes(Vec<i32>),
    Tiered(Vec<bool>),
}

impl Arrays {
    pub(super) fn overrides(&self) -> ListStyle {
        match self.clone() {
            Self::Strings(strings) => ListStyle {
                override_count: Some(1),
                strings,
                ..ListStyle::default()
            },
            Self::Geometries(geometries) => ListStyle {
                override_count: Some(1),
                geometries,
                ..ListStyle::default()
            },
            Self::Indentation(indents, text_indents) => ListStyle {
                override_count: Some(2),
                indents,
                text_indents,
                ..ListStyle::default()
            },
            Self::LabelColor(color) => {
                let (font_color_null, font_color) = label_color_to_native(color);
                ListStyle {
                    override_count: Some(1),
                    font_color_null,
                    font_color,
                    ..ListStyle::default()
                }
            },
            Self::NumberTypes(number_types) => ListStyle {
                override_count: Some(1),
                number_types,
                ..ListStyle::default()
            },
            Self::Tiered(tiered_numbers) => ListStyle {
                override_count: Some(1),
                tiered_numbers,
                ..ListStyle::default()
            },
        }
    }
}

fn override_count(data: &[u8], current: &ListStyle, delta: i64) -> NativeResult<Vec<u8>> {
    let count = i64::from(current.override_count.unwrap_or(0)) + delta;
    let count = u32::try_from(count)
        .map_err(|_| NativeError::Invalid("list style override count is out of range"))?;
    patch_varint_field(
        data,
        OVERRIDE_COUNT_FIELD,
        current.override_count.is_some(),
        Some(u64::from(count)),
    )
    .map_err(native::common)
}

fn geometry_payloads(geometries: &[Geometry], limits: WireLimits) -> NativeResult<Vec<Vec<u8>>> {
    geometries
        .iter()
        .map(|geometry| {
            let mut writer = Writer::new(limits);
            geometry.encode(&mut writer)?;
            Ok(writer.into_bytes())
        })
        .collect()
}

/// Replace one property's arrays in a style payload in place (the host's
/// `replace_direct_*`), counting a newly present override.
pub(super) fn with_arrays(
    data: &[u8],
    current: &ListStyle,
    arrays: &Arrays,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    let wire = native::common;
    Ok(match arrays {
        Arrays::Strings(strings) => {
            let encoded: Vec<Vec<u8>> = strings.iter().map(|s| s.as_bytes().to_vec()).collect();
            let patched = rewrite_repeated_length_delimited_fields(data, STRINGS_FIELD, &encoded)
                .map_err(wire)?;
            if current.strings.is_empty() {
                override_count(&patched, current, 1)?
            } else {
                patched
            }
        },
        Arrays::Geometries(geometries) => {
            let encoded = geometry_payloads(geometries, limits)?;
            let patched =
                rewrite_repeated_length_delimited_fields(data, GEOMETRIES_FIELD, &encoded)
                    .map_err(wire)?;
            if current.geometries.is_empty() {
                override_count(&patched, current, 1)?
            } else {
                patched
            }
        },
        Arrays::Indentation(indents, text_indents) => {
            let missing =
                i64::from(current.indents.is_empty()) + i64::from(current.text_indents.is_empty());
            let indents: Vec<u32> = indents.iter().map(|value| value.to_bits()).collect();
            let text_indents: Vec<u32> = text_indents.iter().map(|value| value.to_bits()).collect();
            let patched =
                rewrite_repeated_fixed32_fields(data, INDENTS_FIELD, &indents).map_err(wire)?;
            let patched =
                rewrite_repeated_fixed32_fields(&patched, TEXT_INDENTS_FIELD, &text_indents)
                    .map_err(wire)?;
            if missing == 0 {
                patched
            } else {
                override_count(&patched, current, missing)?
            }
        },
        Arrays::LabelColor(color) => {
            let had = current.font_color_null.is_some() || current.font_color.is_some();
            let (null, color) = label_color_to_native(*color);
            let patched = patch_varint_field(
                data,
                FONT_COLOR_NULL_FIELD,
                current.font_color_null.is_some(),
                null.map(u64::from),
            )
            .map_err(wire)?;
            let encoded = color
                .map(|color| {
                    let mut writer = Writer::new(limits);
                    color.encode(&mut writer)?;
                    Ok::<_, NativeError>(writer.into_bytes())
                })
                .transpose()?;
            let patched = patch_length_delimited_field(
                &patched,
                FONT_COLOR_FIELD,
                current.font_color.is_some(),
                encoded.as_deref(),
            )
            .map_err(wire)?;
            if had {
                patched
            } else {
                override_count(&patched, current, 1)?
            }
        },
        Arrays::NumberTypes(values) => {
            let encoded = values
                .iter()
                .map(|value| u64::try_from(*value))
                .collect::<Result<Vec<_>, _>>()
                .map_err(|_| {
                    NativeError::Invalid("list style contains a negative number format")
                })?;
            let patched =
                rewrite_repeated_varint_fields(data, NUMBER_TYPES_FIELD, &encoded).map_err(wire)?;
            if current.number_types.is_empty() {
                override_count(&patched, current, 1)?
            } else {
                patched
            }
        },
        Arrays::Tiered(values) => {
            let encoded: Vec<u64> = values.iter().copied().map(u64::from).collect();
            let patched = rewrite_repeated_varint_fields(data, TIERED_NUMBERS_FIELD, &encoded)
                .map_err(wire)?;
            if current.tiered_numbers.is_empty() {
                override_count(&patched, current, 1)?
            } else {
                patched
            }
        },
    })
}

/// The property arrays a removal clears.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Direct {
    Geometries,
    Indentation,
    LabelColor,
    NumberTypes,
    Tiered,
}

/// Remove one property's direct arrays (the host's `remove_direct_*`);
/// `None` when the style holds none.
pub(super) fn without(
    data: &[u8],
    current: &ListStyle,
    direct: Direct,
) -> NativeResult<Option<Vec<u8>>> {
    let wire = native::common;
    Ok(Some(match direct {
        Direct::Geometries => {
            if current.geometries.is_empty() {
                return Ok(None);
            }
            let patched = rewrite_repeated_length_delimited_fields(data, GEOMETRIES_FIELD, &[])
                .map_err(wire)?;
            override_count(&patched, current, -1)?
        },
        Direct::Indentation => {
            let removed = i64::from(!current.indents.is_empty())
                + i64::from(!current.text_indents.is_empty());
            if removed == 0 {
                return Ok(None);
            }
            let patched =
                rewrite_repeated_fixed32_fields(data, INDENTS_FIELD, &[]).map_err(wire)?;
            let patched =
                rewrite_repeated_fixed32_fields(&patched, TEXT_INDENTS_FIELD, &[]).map_err(wire)?;
            override_count(&patched, current, -removed)?
        },
        Direct::LabelColor => {
            if current.font_color_null.is_none() && current.font_color.is_none() {
                return Ok(None);
            }
            let patched = patch_varint_field(
                data,
                FONT_COLOR_NULL_FIELD,
                current.font_color_null.is_some(),
                None,
            )
            .map_err(wire)?;
            let patched = patch_length_delimited_field(
                &patched,
                FONT_COLOR_FIELD,
                current.font_color.is_some(),
                None,
            )
            .map_err(wire)?;
            override_count(&patched, current, -1)?
        },
        Direct::NumberTypes => {
            if current.number_types.is_empty() {
                return Ok(None);
            }
            let patched =
                rewrite_repeated_varint_fields(data, NUMBER_TYPES_FIELD, &[]).map_err(wire)?;
            override_count(&patched, current, -1)?
        },
        Direct::Tiered => {
            if current.tiered_numbers.is_empty() {
                return Ok(None);
            }
            let patched =
                rewrite_repeated_varint_fields(data, TIERED_NUMBERS_FIELD, &[]).map_err(wire)?;
            override_count(&patched, current, -1)?
        },
    }))
}
