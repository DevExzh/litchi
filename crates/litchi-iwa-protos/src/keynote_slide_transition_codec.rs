//! Private-type Buffa projection for Keynote slide transitions.
//!
//! Strict handwritten parsing establishes singularity, wire types, and
//! canonical encodings before Buffa observes a selected transition field.
//! Buffa then supplies a bounded borrowed lazy-view cross-check.  The source
//! bytes remain the only preservation and rewrite representation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes its low-level wire iterator."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_slide_transition_generated::LitchiIwaProjection as projection;

const SLIDE_TRANSITION_FIELD: u32 = 4;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;
const ATTRIBUTES_ANIMATION_FIELD: u32 = 8;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite resource profile for a single Keynote slide-transition payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted slide payload.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            // Unknown fields stay in caller-owned source and this projection
            // does not retain them, but Buffa must still be permitted to skip
            // a finite number of future native fields. Every field consumes at
            // least one source byte, so the message-byte cap is a hard upper
            // bound on this count without allocating an unknown-field vector.
            .with_unknown_field_limit(self.max_message_bytes)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(DecodeError::recursion_limit)?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Borrowed, generated-type-free Keynote slide-transition projection.
///
/// The required native transition envelope is always represented. Every
/// optional scalar preserves source presence rather than applying Keynote's UI
/// defaults; `settings.animation == None` is the legacy/no-modern-animation
/// state.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SlideTransitionSnapshot<'source> {
    /// Modern transition settings inside the required native envelope.
    pub settings: TransitionSettingsSnapshot<'source>,
}

/// Modern `KN.TransitionAttributesArchive` fields.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TransitionSettingsSnapshot<'source> {
    /// Whether any deprecated database transition field (1--7) is present.
    ///
    /// Those fields remain opaque because modern editing must not normalize or
    /// silently erase a legacy-only transition representation.
    pub has_legacy_database_fields: bool,
    /// Optional detailed animation settings.
    pub animation: Option<AnimationSnapshot<'source>>,
    /// Native custom twist (`field 9`).
    pub custom_twist: Option<f32>,
    /// Native mosaic size (`field 10`).
    pub custom_mosaic_size: Option<u32>,
    /// Native mosaic type (`field 11`).
    pub custom_mosaic_type: Option<u32>,
    /// Native bounce setting (`field 12`).
    pub custom_bounce: Option<bool>,
    /// Native magic-move unmatched-object fade setting (`field 13`).
    pub custom_magic_move_fade_unmatched_objects: Option<bool>,
    /// Native timing-curve enum discriminant, including unknown values.
    pub custom_timing_curve: Option<i32>,
    /// Native text-delivery enum discriminant, including unknown values.
    pub custom_text_delivery_type: Option<i32>,
    /// Native motion-blur setting (`field 17`).
    pub custom_motion_blur: Option<bool>,
    /// Native travel distance (`field 18`).
    pub custom_travel_distance: Option<f32>,
}

/// All selected `KN.AnimationAttributesArchive` fields.
///
/// Color and path-source values are opaque, borrowed nested protobuf payloads.
/// They are intentionally not decoded or re-encoded by this projection.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AnimationSnapshot<'source> {
    pub animation_type: Option<&'source str>,
    pub effect: Option<&'source str>,
    pub duration: Option<f64>,
    pub direction: Option<u32>,
    pub delay: Option<f64>,
    pub is_automatic: Option<bool>,
    pub color: Option<&'source [u8]>,
    pub custom_effect_timing_curve_1: Option<&'source [u8]>,
    pub custom_effect_timing_curve_2: Option<&'source [u8]>,
    pub custom_effect_timing_curve_3: Option<&'source [u8]>,
    pub random_number_seed: Option<u32>,
    pub custom_detail: Option<f64>,
    pub custom_effect_timing_curve_theme_name_1: Option<&'source str>,
    pub custom_effect_timing_curve_theme_name_2: Option<&'source str>,
    pub custom_effect_timing_curve_theme_name_3: Option<&'source str>,
    pub writing_direction_is_rtl: Option<bool>,
}

/// Failure from strict Keynote transition preflight or the private projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Projection,
}

impl DecodeError {
    fn recursion_limit() -> Self {
        buffa::DecodeError::RecursionLimitExceeded.into()
    }

    const fn missing_required(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate_singular(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Singular known field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Required known field absent in the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Stable explanation for canonicality failures.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote slide-transition strict preflight disagrees with Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

/// Decode selected transition attributes from a complete `KN.SlideArchive`.
///
/// Strict preflight runs before the private Buffa lazy view, so malformed
/// selected fields cannot use Buffa's more permissive last-one-wins behavior.
pub fn decode_slide_transition<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<SlideTransitionSnapshot<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let strict = preflight_slide_transition(source, options)?;
    let view: projection::KeynoteSlideTransitionArchiveLazyView<'source> =
        options.buffa().decode_lazy_view(source)?;
    let projected = force_slide_transition_projection(&view)?;
    if !same_projected_transition(&projected, &strict) {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode the required `KN.SlideNodeArchive.hasTransition` flag.
pub fn decode_slide_node_has_transition(
    source: &[u8],
    options: DecodeOptions,
) -> Result<bool, DecodeError> {
    validate_decode_input(source, options)?;
    let strict = preflight_slide_node_transition(source, options)?;
    let view: projection::KeynoteSlideNodeTransitionArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_has_transition() || view.has_transition != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    const MAX_RECURSION_LIMIT: u32 = 64;
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes
        || source.len() > options.max_message_bytes
    {
        return Err(buffa::DecodeError::MessageTooLarge.into());
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit());
    }
    Ok(())
}

fn preflight_slide_transition(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SlideTransitionSnapshot<'_>, DecodeError> {
    let mut transition = None;
    for field_result in StrictFields::new(source, options.recursion_limit) {
        let field = field_result?;
        if field.number == SLIDE_TRANSITION_FIELD {
            if transition.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "KN.SlideArchive.transition",
                ));
            }
            transition = Some(preflight_transition(
                field.length_delimited()?,
                options.descend()?,
            )?);
        }
    }
    Ok(SlideTransitionSnapshot {
        settings: transition
            .ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.transition"))?,
    })
}

fn preflight_transition(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TransitionSettingsSnapshot<'_>, DecodeError> {
    let mut attributes = None;
    for field_result in StrictFields::new(source, options.recursion_limit) {
        let field = field_result?;
        if field.number == TRANSITION_ATTRIBUTES_FIELD {
            if attributes.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "KN.TransitionArchive.attributes",
                ));
            }
            attributes = Some(preflight_transition_attributes(
                field.length_delimited()?,
                options.descend()?,
            )?);
        }
    }
    attributes.ok_or_else(|| DecodeError::missing_required("KN.TransitionArchive.attributes"))
}

fn preflight_transition_attributes(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TransitionSettingsSnapshot<'_>, DecodeError> {
    let mut seen = 0u32;
    let mut animation = None;
    let mut custom_twist = None;
    let mut custom_mosaic_size = None;
    let mut custom_mosaic_type = None;
    let mut custom_bounce = None;
    let mut custom_magic_move_fade_unmatched_objects = None;
    let mut custom_timing_curve = None;
    let mut custom_text_delivery_type = None;
    let mut custom_motion_blur = None;
    let mut custom_travel_distance = None;
    let mut has_legacy_database_fields = false;
    for field_result in StrictFields::new(source, options.recursion_limit) {
        let field = field_result?;
        match field.number {
            1..=7 => has_legacy_database_fields = true,
            ATTRIBUTES_ANIMATION_FIELD => {
                mark_singular(
                    &mut seen,
                    8,
                    "KN.TransitionAttributesArchive.animationAttributes",
                )?;
                animation = Some(preflight_animation(
                    field.length_delimited()?,
                    options.descend()?,
                )?);
            },
            9 => {
                mark_singular(&mut seen, 9, "KN.TransitionAttributesArchive.custom_twist")?;
                custom_twist = Some(f32::from_bits(field.fixed32_bits()?));
            },
            10 => {
                mark_singular(
                    &mut seen,
                    10,
                    "KN.TransitionAttributesArchive.custom_mosaic_size",
                )?;
                custom_mosaic_size =
                    Some(u32::try_from(field.varint()?).map_err(|_error| {
                        DecodeError::noncanonical("uint32 scalar exceeds u32")
                    })?);
            },
            11 => {
                mark_singular(
                    &mut seen,
                    11,
                    "KN.TransitionAttributesArchive.custom_mosaic_type",
                )?;
                custom_mosaic_type =
                    Some(u32::try_from(field.varint()?).map_err(|_error| {
                        DecodeError::noncanonical("uint32 scalar exceeds u32")
                    })?);
            },
            12 => {
                mark_singular(
                    &mut seen,
                    12,
                    "KN.TransitionAttributesArchive.custom_bounce",
                )?;
                custom_bounce = Some(require_canonical_bool(field.varint()?)?);
            },
            13 => {
                mark_singular(
                    &mut seen,
                    13,
                    "KN.TransitionAttributesArchive.custom_magic_move_fade_unmatched_objects",
                )?;
                custom_magic_move_fade_unmatched_objects =
                    Some(require_canonical_bool(field.varint()?)?);
            },
            15 => {
                mark_singular(
                    &mut seen,
                    15,
                    "KN.TransitionAttributesArchive.custom_timing_curve",
                )?;
                custom_timing_curve = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            16 => {
                mark_singular(
                    &mut seen,
                    16,
                    "KN.TransitionAttributesArchive.custom_text_delivery_type",
                )?;
                custom_text_delivery_type =
                    Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            17 => {
                mark_singular(
                    &mut seen,
                    17,
                    "KN.TransitionAttributesArchive.custom_motion_blur",
                )?;
                custom_motion_blur = Some(require_canonical_bool(field.varint()?)?);
            },
            18 => {
                mark_singular(
                    &mut seen,
                    18,
                    "KN.TransitionAttributesArchive.custom_travel_distance",
                )?;
                custom_travel_distance = Some(f32::from_bits(field.fixed32_bits()?));
            },
            _ => {},
        }
    }
    Ok(TransitionSettingsSnapshot {
        has_legacy_database_fields,
        animation,
        custom_twist,
        custom_mosaic_size,
        custom_mosaic_type,
        custom_bounce,
        custom_magic_move_fade_unmatched_objects,
        custom_timing_curve,
        custom_text_delivery_type,
        custom_motion_blur,
        custom_travel_distance,
    })
}

fn preflight_animation(
    source: &[u8],
    options: DecodeOptions,
) -> Result<AnimationSnapshot<'_>, DecodeError> {
    let mut seen = 0u32;
    let mut result = AnimationSnapshot {
        animation_type: None,
        effect: None,
        duration: None,
        direction: None,
        delay: None,
        is_automatic: None,
        color: None,
        custom_effect_timing_curve_1: None,
        custom_effect_timing_curve_2: None,
        custom_effect_timing_curve_3: None,
        random_number_seed: None,
        custom_detail: None,
        custom_effect_timing_curve_theme_name_1: None,
        custom_effect_timing_curve_theme_name_2: None,
        custom_effect_timing_curve_theme_name_3: None,
        writing_direction_is_rtl: None,
    };
    for field_result in StrictFields::new(source, options.recursion_limit) {
        let field = field_result?;
        let number = field.number;
        if !(1..=16).contains(&number) {
            continue;
        }
        mark_singular(&mut seen, number, animation_field_name(number))?;
        match number {
            1 => result.animation_type = Some(utf8(field.length_delimited()?)?),
            2 => result.effect = Some(utf8(field.length_delimited()?)?),
            3 => result.duration = Some(f64::from_bits(field.fixed64_bits()?)),
            4 => {
                result.direction =
                    Some(u32::try_from(field.varint()?).map_err(|_error| {
                        DecodeError::noncanonical("uint32 scalar exceeds u32")
                    })?);
            },
            5 => result.delay = Some(f64::from_bits(field.fixed64_bits()?)),
            6 => result.is_automatic = Some(require_canonical_bool(field.varint()?)?),
            7 => result.color = Some(field.length_delimited()?),
            8 => result.custom_effect_timing_curve_1 = Some(field.length_delimited()?),
            9 => result.custom_effect_timing_curve_2 = Some(field.length_delimited()?),
            10 => result.custom_effect_timing_curve_3 = Some(field.length_delimited()?),
            11 => {
                result.random_number_seed =
                    Some(u32::try_from(field.varint()?).map_err(|_error| {
                        DecodeError::noncanonical("uint32 scalar exceeds u32")
                    })?);
            },
            12 => result.custom_detail = Some(f64::from_bits(field.fixed64_bits()?)),
            13 => {
                result.custom_effect_timing_curve_theme_name_1 =
                    Some(utf8(field.length_delimited()?)?);
            },
            14 => {
                result.custom_effect_timing_curve_theme_name_2 =
                    Some(utf8(field.length_delimited()?)?);
            },
            15 => {
                result.custom_effect_timing_curve_theme_name_3 =
                    Some(utf8(field.length_delimited()?)?);
            },
            16 => result.writing_direction_is_rtl = Some(require_canonical_bool(field.varint()?)?),
            _ => unreachable!("range checked above"),
        }
    }
    Ok(result)
}

fn preflight_slide_node_transition(
    source: &[u8],
    options: DecodeOptions,
) -> Result<bool, DecodeError> {
    let mut has_transition = None;
    for field_result in StrictFields::new(source, options.recursion_limit) {
        let field = field_result?;
        if field.number == 7 {
            if has_transition.is_some() {
                return Err(DecodeError::duplicate_singular(
                    "KN.SlideNodeArchive.hasTransition",
                ));
            }
            has_transition = Some(require_canonical_bool(field.varint()?)?);
        }
    }
    has_transition.ok_or_else(|| DecodeError::missing_required("KN.SlideNodeArchive.hasTransition"))
}

fn force_slide_transition_projection<'source>(
    view: &projection::KeynoteSlideTransitionArchiveLazyView<'source>,
) -> Result<SlideTransitionSnapshot<'source>, DecodeError> {
    let transition = view
        .transition
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.SlideArchive.transition"))?;
    let attributes = transition
        .attributes
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.TransitionArchive.attributes"))?;
    let animation = attributes
        .animation_attributes
        .get()?
        .map(|animation| AnimationSnapshot {
            animation_type: animation.animation_type,
            effect: animation.effect,
            duration: animation.duration,
            direction: animation.direction,
            delay: animation.delay,
            is_automatic: animation.is_automatic,
            color: animation.color,
            custom_effect_timing_curve_1: animation.custom_effect_timing_curve_1,
            custom_effect_timing_curve_2: animation.custom_effect_timing_curve_2,
            custom_effect_timing_curve_3: animation.custom_effect_timing_curve_3,
            random_number_seed: animation.random_number_seed,
            custom_detail: animation.custom_detail,
            custom_effect_timing_curve_theme_name_1: animation
                .custom_effect_timing_curve_theme_name_1,
            custom_effect_timing_curve_theme_name_2: animation
                .custom_effect_timing_curve_theme_name_2,
            custom_effect_timing_curve_theme_name_3: animation
                .custom_effect_timing_curve_theme_name_3,
            writing_direction_is_rtl: animation.writing_direction_is_rtl,
        });
    Ok(SlideTransitionSnapshot {
        settings: TransitionSettingsSnapshot {
            has_legacy_database_fields: false,
            animation,
            custom_twist: attributes.custom_twist,
            custom_mosaic_size: attributes.custom_mosaic_size,
            custom_mosaic_type: attributes.custom_mosaic_type,
            custom_bounce: attributes.custom_bounce,
            custom_magic_move_fade_unmatched_objects: attributes
                .custom_magic_move_fade_unmatched_objects,
            custom_timing_curve: attributes.custom_timing_curve,
            custom_text_delivery_type: attributes.custom_text_delivery_type,
            custom_motion_blur: attributes.custom_motion_blur,
            custom_travel_distance: attributes.custom_travel_distance,
        },
    })
}

fn same_projected_transition(
    projected_snapshot: &SlideTransitionSnapshot<'_>,
    strict_snapshot: &SlideTransitionSnapshot<'_>,
) -> bool {
    let projected_settings = projected_snapshot.settings;
    let strict_settings = strict_snapshot.settings;
    same_animation(projected_settings.animation, strict_settings.animation)
        && projected_settings.custom_twist.map(f32::to_bits)
            == strict_settings.custom_twist.map(f32::to_bits)
        && projected_settings.custom_mosaic_size == strict_settings.custom_mosaic_size
        && projected_settings.custom_mosaic_type == strict_settings.custom_mosaic_type
        && projected_settings.custom_bounce == strict_settings.custom_bounce
        && projected_settings.custom_magic_move_fade_unmatched_objects
            == strict_settings.custom_magic_move_fade_unmatched_objects
        && projected_settings.custom_timing_curve == strict_settings.custom_timing_curve
        && projected_settings.custom_text_delivery_type == strict_settings.custom_text_delivery_type
        && projected_settings.custom_motion_blur == strict_settings.custom_motion_blur
        && projected_settings.custom_travel_distance.map(f32::to_bits)
            == strict_settings.custom_travel_distance.map(f32::to_bits)
}

fn same_animation(
    projected_animation: Option<AnimationSnapshot<'_>>,
    strict_animation: Option<AnimationSnapshot<'_>>,
) -> bool {
    match (projected_animation, strict_animation) {
        (None, None) => true,
        (Some(projected_value), Some(strict_value)) => {
            projected_value.animation_type == strict_value.animation_type
                && projected_value.effect == strict_value.effect
                && projected_value.duration.map(f64::to_bits)
                    == strict_value.duration.map(f64::to_bits)
                && projected_value.direction == strict_value.direction
                && projected_value.delay.map(f64::to_bits) == strict_value.delay.map(f64::to_bits)
                && projected_value.is_automatic == strict_value.is_automatic
                && projected_value.color == strict_value.color
                && projected_value.custom_effect_timing_curve_1
                    == strict_value.custom_effect_timing_curve_1
                && projected_value.custom_effect_timing_curve_2
                    == strict_value.custom_effect_timing_curve_2
                && projected_value.custom_effect_timing_curve_3
                    == strict_value.custom_effect_timing_curve_3
                && projected_value.random_number_seed == strict_value.random_number_seed
                && projected_value.custom_detail.map(f64::to_bits)
                    == strict_value.custom_detail.map(f64::to_bits)
                && projected_value.custom_effect_timing_curve_theme_name_1
                    == strict_value.custom_effect_timing_curve_theme_name_1
                && projected_value.custom_effect_timing_curve_theme_name_2
                    == strict_value.custom_effect_timing_curve_theme_name_2
                && projected_value.custom_effect_timing_curve_theme_name_3
                    == strict_value.custom_effect_timing_curve_theme_name_3
                && projected_value.writing_direction_is_rtl == strict_value.writing_direction_is_rtl
        },
        (None, Some(_)) | (Some(_), None) => false,
    }
}

fn animation_field_name(number: u32) -> &'static str {
    match number {
        1 => "KN.AnimationAttributesArchive.animation_type",
        2 => "KN.AnimationAttributesArchive.effect",
        3 => "KN.AnimationAttributesArchive.duration",
        4 => "KN.AnimationAttributesArchive.direction",
        5 => "KN.AnimationAttributesArchive.delay",
        6 => "KN.AnimationAttributesArchive.is_automatic",
        7 => "KN.AnimationAttributesArchive.color",
        8 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_1",
        9 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_2",
        10 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_3",
        11 => "KN.AnimationAttributesArchive.random_number_seed",
        12 => "KN.AnimationAttributesArchive.custom_detail",
        13 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_theme_name_1",
        14 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_theme_name_2",
        15 => "KN.AnimationAttributesArchive.custom_effect_timing_curve_theme_name_3",
        16 => "KN.AnimationAttributesArchive.writing_direction_is_rtl",
        _ => unreachable!("only selected animation fields call this helper"),
    }
}

fn mark_singular(seen: &mut u32, field_number: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u32
        .checked_shl(field_number)
        .ok_or_else(|| DecodeError::noncanonical("known field number exceeds presence mask"))?;
    if *seen & bit != 0 {
        return Err(DecodeError::duplicate_singular(name));
    }
    *seen |= bit;
    Ok(())
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn require_canonical_int32(value: u64) -> Result<u64, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    Ok(value)
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Strict preflight proved the u64 is a canonical sign-extended int32."
)]
fn decode_int32(value: u64) -> i32 {
    value as i32
}

fn utf8(source: &[u8]) -> Result<&str, DecodeError> {
    str::from_utf8(source).map_err(|_error| DecodeError::noncanonical("string field is not UTF-8"))
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64(u64),
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32(u32),
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        if self.wire_type != expected {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: expected as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Varint)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn fixed32_bits(self) -> Result<u32, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed32)?;
        match self.value {
            StrictValue::Fixed32(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group => Err(DecodeError::projection()),
        }
    }

    fn fixed64_bits(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed64)?;
        match self.value {
            StrictValue::Fixed64(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }
}

#[derive(Clone, Debug)]
struct StrictFields<'source> {
    remaining: &'source [u8],
    recursion_limit: u32,
}

impl<'source> StrictFields<'source> {
    const fn new(source: &'source [u8], recursion_limit: u32) -> Self {
        Self {
            remaining: source,
            recursion_limit,
        }
    }

    fn next_field(&mut self) -> Result<Option<StrictField<'source>>, DecodeError> {
        match parse_strict_field(&mut self.remaining, self.recursion_limit)? {
            Some(ParseItem::Field(field)) => Ok(Some(field)),
            Some(ParseItem::EndGroup(number)) => {
                Err(buffa::DecodeError::InvalidEndGroup(number).into())
            },
            None => Ok(None),
        }
    }
}

impl<'source> Iterator for StrictFields<'source> {
    type Item = Result<StrictField<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.next_field() {
            Ok(Some(field)) => Some(Ok(field)),
            Ok(None) => None,
            Err(error) => {
                self.remaining = &[];
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for StrictFields<'_> {}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let (value, canonical_value) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            (StrictValue::Varint(value), canonical)
        },
        buffa::encoding::WireType::Fixed64 => {
            let bytes = take_exact(source, 8)?;
            let bits = u64::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_error| buffa::DecodeError::UnexpectedEof)?,
            );
            (StrictValue::Fixed64(bits), true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                canonical,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(DecodeError::recursion_limit)?;
            skip_strict_group(source, field_number, child_limit)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            let bytes = take_exact(source, 4)?;
            let bits = u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_error| buffa::DecodeError::UnexpectedEof)?,
            );
            (StrictValue::Fixed32(bits), true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        canonical_key,
        canonical_value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    recursion_limit: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
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
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::cast_possible_truncation,
        clippy::expect_used,
        clippy::large_types_passed_by_value,
        clippy::shadow_unrelated,
        reason = "Focused adversarial tests deliberately use direct assertions and compact field-number tables."
    )]

    use super::{
        DecodeError, DecodeOptions, decode_slide_node_has_transition, decode_slide_transition,
    };

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), 8)
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "Each emitted varint byte intentionally contains only its low seven bits."
    )]
    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn push_overlong_varint(value: u64, output: &mut Vec<u8>) {
        let start = output.len();
        push_varint(value, output);
        *output.last_mut().expect("canonical varint has one byte") |= 0x80;
        output.push(0);
        debug_assert!(output.len() >= start + 2);
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn fixed32_field(number: u32, value: f32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(number) << 3) | 5, &mut output);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn fixed64_field(number: u32, value: f64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(number) << 3) | 1, &mut output);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn length_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::with_capacity(payload.len() + 8);
        push_varint((u64::from(number) << 3) | 2, &mut output);
        push_varint(payload.len() as u64, &mut output);
        output.extend_from_slice(payload);
        output
    }

    fn slide(attributes: &[u8]) -> Vec<u8> {
        let transition = length_field(2, attributes);
        length_field(4, &transition)
    }

    fn with_animation(animation: &[u8]) -> Vec<u8> {
        slide(&length_field(8, animation))
    }

    fn error<T>(result: Result<T, DecodeError>) -> DecodeError {
        match result {
            Err(error) => error,
            Ok(_) => panic!("malformed selected transition field unexpectedly decoded"),
        }
    }

    #[test]
    fn projects_every_modern_field_and_borrows_opaque_payloads()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut animation = Vec::new();
        animation.extend(length_field(1, b"dissolve"));
        animation.extend(length_field(2, b"in"));
        animation.extend(fixed64_field(3, 1.25));
        animation.extend(varint_field(4, 9));
        animation.extend(fixed64_field(5, 0.5));
        animation.extend(varint_field(6, 1));
        animation.extend(length_field(7, &[0xff, 0x00]));
        animation.extend(length_field(8, &[0x80]));
        animation.extend(length_field(9, &[0x81]));
        animation.extend(length_field(10, &[0x82]));
        animation.extend(varint_field(11, 42));
        animation.extend(fixed64_field(12, 3.5));
        animation.extend(length_field(13, b"curve-a"));
        animation.extend(length_field(14, b"curve-b"));
        animation.extend(length_field(15, b"curve-c"));
        animation.extend(varint_field(16, 0));
        let mut attributes = length_field(8, &animation);
        attributes.extend(fixed32_field(9, 2.5));
        attributes.extend(varint_field(10, 12));
        attributes.extend(varint_field(11, 13));
        attributes.extend(varint_field(12, 1));
        attributes.extend(varint_field(13, 0));
        attributes.extend(varint_field(15, 5));
        attributes.extend(varint_field(16, u64::MAX));
        attributes.extend(varint_field(17, 1));
        attributes.extend(fixed32_field(18, 88.0));
        let source = slide(&attributes);

        let snapshot = decode_slide_transition(&source, options(&source))?;
        let settings = snapshot.settings;
        assert!(!settings.has_legacy_database_fields);
        assert_eq!(settings.custom_twist, Some(2.5));
        assert_eq!(settings.custom_mosaic_size, Some(12));
        assert_eq!(settings.custom_mosaic_type, Some(13));
        assert_eq!(settings.custom_bounce, Some(true));
        assert_eq!(
            settings.custom_magic_move_fade_unmatched_objects,
            Some(false)
        );
        assert_eq!(settings.custom_timing_curve, Some(5));
        assert_eq!(settings.custom_text_delivery_type, Some(-1));
        assert_eq!(settings.custom_motion_blur, Some(true));
        assert_eq!(settings.custom_travel_distance, Some(88.0));
        let animation = settings.animation.expect("animation is present");
        assert_eq!(animation.animation_type, Some("dissolve"));
        assert_eq!(animation.effect, Some("in"));
        assert_eq!(animation.duration, Some(1.25));
        assert_eq!(animation.direction, Some(9));
        assert_eq!(animation.delay, Some(0.5));
        assert_eq!(animation.is_automatic, Some(true));
        assert_eq!(animation.color, Some(&[0xff, 0x00][..]));
        assert_eq!(animation.custom_effect_timing_curve_1, Some(&[0x80][..]));
        assert_eq!(animation.custom_effect_timing_curve_2, Some(&[0x81][..]));
        assert_eq!(animation.custom_effect_timing_curve_3, Some(&[0x82][..]));
        assert_eq!(animation.random_number_seed, Some(42));
        assert_eq!(animation.custom_detail, Some(3.5));
        assert_eq!(
            animation.custom_effect_timing_curve_theme_name_1,
            Some("curve-a")
        );
        assert_eq!(
            animation.custom_effect_timing_curve_theme_name_2,
            Some("curve-b")
        );
        assert_eq!(
            animation.custom_effect_timing_curve_theme_name_3,
            Some("curve-c")
        );
        assert_eq!(animation.writing_direction_is_rtl, Some(false));
        let color = animation.color.expect("color payload is present");
        assert!(core::ptr::eq(
            color.as_ptr(),
            source[color_offset(&source, color)..].as_ptr()
        ));
        Ok(())
    }

    fn color_offset(source: &[u8], color: &[u8]) -> usize {
        let start = source.as_ptr() as usize;
        let value = color.as_ptr() as usize;
        value.checked_sub(start).expect("borrow points into source")
    }

    #[test]
    fn every_selected_animation_and_custom_field_preserves_presence()
    -> Result<(), Box<dyn std::error::Error>> {
        let animation_fields = [
            length_field(1, b"type"),
            length_field(2, b"effect"),
            fixed64_field(3, 1.0),
            varint_field(4, 1),
            fixed64_field(5, 2.0),
            varint_field(6, 0),
            length_field(7, &[1]),
            length_field(8, &[2]),
            length_field(9, &[3]),
            length_field(10, &[4]),
            varint_field(11, 5),
            fixed64_field(12, 6.0),
            length_field(13, b"one"),
            length_field(14, b"two"),
            length_field(15, b"three"),
            varint_field(16, 1),
        ];
        for (index, field) in animation_fields.iter().enumerate() {
            let source = with_animation(field);
            let animation = decode_slide_transition(&source, options(&source))?
                .settings
                .animation
                .expect("one selected field retains an animation envelope");
            assert!(animation_present(animation, index as u32 + 1));
        }
        let custom_fields = [
            fixed32_field(9, 1.0),
            varint_field(10, 1),
            varint_field(11, 1),
            varint_field(12, 0),
            varint_field(13, 1),
            varint_field(15, 1),
            varint_field(16, u64::MAX),
            varint_field(17, 0),
            fixed32_field(18, 2.0),
        ];
        for (index, field) in custom_fields.iter().enumerate() {
            let source = slide(field);
            let settings = decode_slide_transition(&source, options(&source))?.settings;
            assert!(custom_present(
                settings,
                [9, 10, 11, 12, 13, 15, 16, 17, 18][index]
            ));
            assert!(settings.animation.is_none());
        }
        let legacy = slide(&length_field(1, b"legacy"));
        let settings = decode_slide_transition(&legacy, options(&legacy))?.settings;
        assert!(settings.has_legacy_database_fields);
        assert!(settings.animation.is_none());
        Ok(())
    }

    #[test]
    fn unknown_nested_fields_are_skipped_without_retention_or_normalization()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut animation = length_field(1, b"known");
        animation.extend(varint_field(92, 77));
        let mut attributes = length_field(8, &animation);
        attributes.extend(length_field(91, &[0xff, 0x00]));
        let source = slide(&attributes);
        let snapshot = decode_slide_transition(&source, options(&source))?;
        assert_eq!(
            snapshot
                .settings
                .animation
                .expect("known animation is retained")
                .animation_type,
            Some("known")
        );
        Ok(())
    }

    fn animation_present(animation: super::AnimationSnapshot<'_>, field: u32) -> bool {
        match field {
            1 => animation.animation_type.is_some(),
            2 => animation.effect.is_some(),
            3 => animation.duration.is_some(),
            4 => animation.direction.is_some(),
            5 => animation.delay.is_some(),
            6 => animation.is_automatic.is_some(),
            7 => animation.color.is_some(),
            8 => animation.custom_effect_timing_curve_1.is_some(),
            9 => animation.custom_effect_timing_curve_2.is_some(),
            10 => animation.custom_effect_timing_curve_3.is_some(),
            11 => animation.random_number_seed.is_some(),
            12 => animation.custom_detail.is_some(),
            13 => animation.custom_effect_timing_curve_theme_name_1.is_some(),
            14 => animation.custom_effect_timing_curve_theme_name_2.is_some(),
            15 => animation.custom_effect_timing_curve_theme_name_3.is_some(),
            16 => animation.writing_direction_is_rtl.is_some(),
            _ => false,
        }
    }

    fn custom_present(settings: super::TransitionSettingsSnapshot<'_>, field: u32) -> bool {
        match field {
            9 => settings.custom_twist.is_some(),
            10 => settings.custom_mosaic_size.is_some(),
            11 => settings.custom_mosaic_type.is_some(),
            12 => settings.custom_bounce.is_some(),
            13 => settings.custom_magic_move_fade_unmatched_objects.is_some(),
            15 => settings.custom_timing_curve.is_some(),
            16 => settings.custom_text_delivery_type.is_some(),
            17 => settings.custom_motion_blur.is_some(),
            18 => settings.custom_travel_distance.is_some(),
            _ => false,
        }
    }

    #[test]
    fn strict_preflight_rejects_required_duplicate_wrong_wire_and_noncanonical_selected_fields() {
        let missing_slide = error(decode_slide_transition(&[], DecodeOptions::new(1, 8)));
        assert_eq!(
            missing_slide.missing_required_field(),
            Some("KN.SlideArchive.transition")
        );
        let missing_attributes = error(decode_slide_transition(
            &length_field(4, &[]),
            DecodeOptions::new(2, 8),
        ));
        assert_eq!(
            missing_attributes.missing_required_field(),
            Some("KN.TransitionArchive.attributes")
        );

        let mut duplicate_animation = length_field(2, b"one");
        duplicate_animation.extend(length_field(2, b"two"));
        let duplicate = error(decode_slide_transition(
            &with_animation(&duplicate_animation),
            options(&with_animation(&duplicate_animation)),
        ));
        assert_eq!(
            duplicate.duplicate_singular_field(),
            Some("KN.AnimationAttributesArchive.effect")
        );

        let wrong_animation = with_animation(&varint_field(2, 1));
        assert!(
            !error(decode_slide_transition(
                &wrong_animation,
                options(&wrong_animation)
            ))
            .to_string()
            .is_empty()
        );
        let wrong_custom = slide(&length_field(9, b"not a float"));
        assert!(
            !error(decode_slide_transition(
                &wrong_custom,
                options(&wrong_custom)
            ))
            .to_string()
            .is_empty()
        );

        let mut noncanonical_key = Vec::new();
        push_overlong_varint((4_u64 << 3) | 2, &mut noncanonical_key);
        let transition = length_field(2, &[]);
        push_varint(transition.len() as u64, &mut noncanonical_key);
        noncanonical_key.extend_from_slice(&transition);
        assert_eq!(
            error(decode_slide_transition(
                &noncanonical_key,
                options(&noncanonical_key)
            ))
            .noncanonical_reason(),
            Some("protobuf field key")
        );

        let mut noncanonical_transition_size = Vec::new();
        push_varint((4_u64 << 3) | 2, &mut noncanonical_transition_size);
        push_overlong_varint(transition.len() as u64, &mut noncanonical_transition_size);
        noncanonical_transition_size.extend_from_slice(&transition);
        assert_eq!(
            error(decode_slide_transition(
                &noncanonical_transition_size,
                options(&noncanonical_transition_size)
            ))
            .noncanonical_reason(),
            Some("length-delimited size")
        );

        let mut noncanonical_direction = Vec::new();
        push_varint(4_u64 << 3, &mut noncanonical_direction);
        push_overlong_varint(1, &mut noncanonical_direction);
        let source = with_animation(&noncanonical_direction);
        assert_eq!(
            error(decode_slide_transition(&source, options(&source))).noncanonical_reason(),
            Some("protobuf varint value")
        );

        let bad_bool = with_animation(&varint_field(6, 2));
        assert_eq!(
            error(decode_slide_transition(&bad_bool, options(&bad_bool))).noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn node_flag_is_required_canonical_and_cross_checked() -> Result<(), Box<dyn std::error::Error>>
    {
        let source = varint_field(7, 1);
        assert!(decode_slide_node_has_transition(&source, options(&source))?);
        let absent = error(decode_slide_node_has_transition(
            &[],
            DecodeOptions::new(1, 8),
        ));
        assert_eq!(
            absent.missing_required_field(),
            Some("KN.SlideNodeArchive.hasTransition")
        );
        let mut duplicate = varint_field(7, 0);
        duplicate.extend(varint_field(7, 1));
        assert_eq!(
            error(decode_slide_node_has_transition(
                &duplicate,
                options(&duplicate)
            ))
            .duplicate_singular_field(),
            Some("KN.SlideNodeArchive.hasTransition")
        );
        let invalid = varint_field(7, 2);
        assert_eq!(
            error(decode_slide_node_has_transition(
                &invalid,
                options(&invalid)
            ))
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        Ok(())
    }

    #[test]
    fn finite_limits_are_enforced_before_projection() {
        let source = slide(&[]);
        assert!(
            !error(decode_slide_transition(
                &source,
                DecodeOptions::new(source.len() - 1, 8)
            ))
            .to_string()
            .is_empty()
        );
        assert!(
            !error(decode_slide_transition(
                &source,
                DecodeOptions::new(source.len(), 0)
            ))
            .to_string()
            .is_empty()
        );
        let animation = with_animation(&length_field(1, b"x"));
        assert!(
            !error(decode_slide_transition(
                &animation,
                DecodeOptions::new(animation.len(), 2)
            ))
            .to_string()
            .is_empty()
        );
    }
}
