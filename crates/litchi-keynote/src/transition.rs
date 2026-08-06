//! Keynote transition vocabulary independent of archive and wire models.

use crate::{Error, Result};

const NONE_EFFECT: &str = "none";
const DISSOLVE_EFFECT: &str = "apple:dissolve";
const MAGIC_MOVE_EFFECT: &str = "apple:magic-move-implied-motion-path";
const LINEAR_ACCELERATION: i32 = 1;
const EASE_IN_ACCELERATION: i32 = 2;
const EASE_OUT_ACCELERATION: i32 = 3;
const EASE_IN_OUT_ACCELERATION: i32 = 4;
const CUSTOM_ACCELERATION: i32 = 5;
const BY_OBJECT_DELIVERY: i32 = 1;
const BY_WORD_DELIVERY: i32 = 2;
const BY_CHARACTER_DELIVERY: i32 = 3;
const BY_LINE_DELIVERY: i32 = 4;
const TIMING_CURVE_SLOT_COUNT: usize = 3;

/// Maximum UTF-8 byte length of a transition identifier or theme name.
pub const MAX_IDENTIFIER_BYTES: usize = 64 * 1024;
/// Maximum size of one opaque color or timing-curve payload.
pub const MAX_OPAQUE_PAYLOAD_BYTES: usize = 4 * 1024 * 1024;

/// One of the fixed timing-curve slots carried by a Keynote transition.
///
/// A closed slot type keeps indexed access panic-free while preserving the
/// native three-slot layout without allocating a collection for the index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TimingCurveSlot {
    /// The first timing-curve slot.
    First,
    /// The second timing-curve slot.
    Second,
    /// The third timing-curve slot.
    Third,
}

impl TimingCurveSlot {
    const fn index(self) -> usize {
        match self {
            Self::First => 0,
            Self::Second => 1,
            Self::Third => 2,
        }
    }

    /// All native timing-curve slots in wire order.
    pub const ALL: [Self; TIMING_CURVE_SLOT_COUNT] = [Self::First, Self::Second, Self::Third];
}

/// A slide-transition effect understood by Keynote.
///
/// Named variants cover identifiers verified in native Keynote documents.
/// [`Effect::Unknown`] preserves identifiers introduced by future releases
/// without reducing the semantic API to an untyped string.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Effect {
    /// Keynote's native “No Transition Effect” representation.
    None,
    /// A dissolve transition.
    Dissolve,
    /// Keynote's Magic Move transition.
    MagicMove,
    /// An identifier introduced by a newer Keynote release.
    Unknown { identifier: Box<str> },
}

impl Effect {
    /// Decode a native effect identifier without discarding unknown values.
    ///
    /// # Errors
    ///
    /// Returns a typed error when an unknown identifier is empty, contains a
    /// NUL byte, or exceeds the bounded semantic storage budget.
    pub fn from_identifier(identifier: &str) -> Result<Self> {
        match identifier {
            NONE_EFFECT => Ok(Self::None),
            DISSOLVE_EFFECT => Ok(Self::Dissolve),
            MAGIC_MOVE_EFFECT => Ok(Self::MagicMove),
            other => Self::unknown(other),
        }
    }

    /// Construct an unknown effect identifier while rejecting non-canonical
    /// known identifiers and strings that cannot be written to native text
    /// fields.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalEffect`] when `identifier` names a known
    /// effect, or [`Error::NulString`] when it contains a NUL byte.
    pub fn unknown(identifier: impl AsRef<str>) -> Result<Self> {
        let identifier_text = identifier.as_ref();
        Self::validate_identifier(identifier_text)?;
        if matches!(
            identifier_text,
            NONE_EFFECT | DISSOLVE_EFFECT | MAGIC_MOVE_EFFECT
        ) {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(Self::Unknown {
            identifier: identifier_text.into(),
        })
    }

    /// Return the native Keynote effect identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::None => NONE_EFFECT,
            Self::Dissolve => DISSOLVE_EFFECT,
            Self::MagicMove => MAGIC_MOVE_EFFECT,
            Self::Unknown { identifier } => identifier,
        }
    }

    /// Return whether the value uses a named variant for a known identifier.
    #[must_use]
    pub fn is_canonical(&self) -> bool {
        !matches!(
            self,
            Self::Unknown { identifier }
                if matches!(
                    &**identifier,
                    NONE_EFFECT | DISSOLVE_EFFECT | MAGIC_MOVE_EFFECT
                )
        )
    }

    /// Validate the effect before it is published through a native adapter.
    ///
    /// Unknown identifiers are valid and remain lossless. The unknown form is
    /// rejected only when it shadows a named native identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NulString`] when the identifier contains a NUL byte or
    /// [`Error::NonCanonicalEffect`] when an unknown value shadows a named
    /// native effect.
    pub fn validate(&self) -> Result<()> {
        Self::validate_identifier(self.identifier())?;
        if !self.is_canonical() {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(())
    }

    fn validate_identifier(identifier: &str) -> Result<()> {
        if identifier.is_empty() {
            return Err(Error::EmptyIdentifier);
        }
        if identifier.len() > MAX_IDENTIFIER_BYTES {
            return Err(Error::IdentifierTooLarge);
        }
        if identifier.contains('\0') {
            return Err(Error::NulString);
        }
        Ok(())
    }
}

/// Effect-specific direction discriminator stored by Keynote.
///
/// The native value is retained directly so unknown effect-specific values
/// remain lossless without increasing the value beyond four bytes.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Direction(u32);

impl Direction {
    /// Wrap a native effect-specific direction value.
    #[must_use]
    pub const fn from_native(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific direction value.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Effect-specific mosaic layout discriminator stored by Keynote.
///
/// The native value is retained directly so unknown effect-specific values
/// remain lossless without increasing the value beyond four bytes.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MosaicType(u32);

impl MosaicType {
    /// Wrap a native effect-specific mosaic layout value.
    #[must_use]
    pub const fn from_native(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific mosaic layout value.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Recognized transition acceleration curves.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AccelerationKind {
    /// Constant speed; displayed as “None” in Keynote.
    Linear,
    /// Ease into the transition.
    EaseIn,
    /// Ease out of the transition.
    EaseOut,
    /// Ease into and out of the transition.
    EaseInOut,
    /// A custom timing curve.
    Custom,
}

/// Acceleration curve used by transitions such as Magic Move.
///
/// Known curves have named associated constants, while future native values
/// remain lossless in the same four-byte representation.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Acceleration(i32);

impl Acceleration {
    /// Constant-speed transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Linear: Self = Self(LINEAR_ACCELERATION);
    /// Ease into the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseIn: Self = Self(EASE_IN_ACCELERATION);
    /// Ease out of the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseOut: Self = Self(EASE_OUT_ACCELERATION);
    /// Ease into and out of the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseInOut: Self = Self(EASE_IN_OUT_ACCELERATION);
    /// Use a custom timing curve.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Custom: Self = Self(CUSTOM_ACCELERATION);

    /// Decode a native timing-curve value without discarding unknown values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native transition timing-curve value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized timing-curve kind, if known.
    #[must_use]
    pub const fn kind(self) -> Option<AccelerationKind> {
        match self.0 {
            LINEAR_ACCELERATION => Some(AccelerationKind::Linear),
            EASE_IN_ACCELERATION => Some(AccelerationKind::EaseIn),
            EASE_OUT_ACCELERATION => Some(AccelerationKind::EaseOut),
            EASE_IN_OUT_ACCELERATION => Some(AccelerationKind::EaseInOut),
            CUSTOM_ACCELERATION => Some(AccelerationKind::Custom),
            _ => None,
        }
    }
}

impl std::fmt::Debug for Acceleration {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.kind() {
            Some(kind) => kind.fmt(formatter),
            None => formatter.debug_tuple("Unsupported").field(&self.0).finish(),
        }
    }
}

/// Recognized ways to deliver matching text during a transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextDeliveryKind {
    /// Deliver by object.
    ByObject,
    /// Deliver by word.
    ByWord,
    /// Deliver by character.
    ByCharacter,
    /// Deliver by line.
    ByLine,
}

/// How matching text is delivered during a Keynote transition.
///
/// Known delivery modes have named associated constants, while future native
/// values remain lossless in the same four-byte representation.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextDelivery(i32);

impl TextDelivery {
    /// Deliver by object.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByObject: Self = Self(BY_OBJECT_DELIVERY);
    /// Deliver by word.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByWord: Self = Self(BY_WORD_DELIVERY);
    /// Deliver by character.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByCharacter: Self = Self(BY_CHARACTER_DELIVERY);
    /// Deliver by line.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByLine: Self = Self(BY_LINE_DELIVERY);

    /// Decode a native text-delivery value without discarding unknown values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native transition text-delivery value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized delivery kind, if known.
    #[must_use]
    pub const fn kind(self) -> Option<TextDeliveryKind> {
        match self.0 {
            BY_OBJECT_DELIVERY => Some(TextDeliveryKind::ByObject),
            BY_WORD_DELIVERY => Some(TextDeliveryKind::ByWord),
            BY_CHARACTER_DELIVERY => Some(TextDeliveryKind::ByCharacter),
            BY_LINE_DELIVERY => Some(TextDeliveryKind::ByLine),
            _ => None,
        }
    }
}

impl std::fmt::Debug for TextDelivery {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.kind() {
            Some(kind) => kind.fmt(formatter),
            None => formatter.debug_tuple("Unsupported").field(&self.0).finish(),
        }
    }
}

/// Lossless animation-level values attached to a slide transition.
///
/// The payload fields intentionally remain opaque. Their protobuf structure,
/// validation, and wire-preserving patching belong to the IWA adapter; this
/// crate only owns bounded native-free storage and semantic scalar checks.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct AnimationParameters {
    /// Exact native color payload, when one is present.
    color_payload: Option<Box<[u8]>>,
    /// Exact native timing-curve payloads for Keynote's three curve slots.
    ///
    /// This is a fixed-size array because the native schema has exactly three
    /// slots. Keeping the slots inline avoids an allocation and capacity
    /// overhead for the overwhelmingly common all-absent value.
    timing_curve_payloads: [Option<Box<[u8]>>; TIMING_CURVE_SLOT_COUNT],
    /// Native random seed used by effects with randomized motion.
    random_number_seed: Option<u32>,
    /// Native effect detail value.
    detail: Option<f64>,
    /// Native theme names for the three timing-curve slots.
    timing_curve_theme_names: [Option<Box<str>>; TIMING_CURVE_SLOT_COUNT],
    /// Whether transition text should use right-to-left writing direction.
    writing_direction_is_rtl: Option<bool>,
}

impl AnimationParameters {
    /// Construct empty animation parameters.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            color_payload: None,
            timing_curve_payloads: [None, None, None],
            random_number_seed: None,
            detail: None,
            timing_curve_theme_names: [None, None, None],
            writing_direction_is_rtl: None,
        }
    }

    /// Validate semantic values without interpreting opaque native payloads.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDetail`] for a non-finite detail value or
    /// [`Error::NulString`] for a theme name containing a NUL byte.
    pub fn validate(&self) -> Result<()> {
        if self.detail.is_some_and(|value| !value.is_finite()) {
            return Err(Error::InvalidDetail);
        }
        if self
            .color_payload
            .as_deref()
            .is_some_and(|payload| payload.len() > MAX_OPAQUE_PAYLOAD_BYTES)
            || self
                .timing_curve_payloads
                .iter()
                .flatten()
                .any(|payload| payload.len() > MAX_OPAQUE_PAYLOAD_BYTES)
        {
            return Err(Error::PayloadTooLarge);
        }
        if self
            .timing_curve_theme_names
            .iter()
            .flatten()
            .any(|name| name.len() > MAX_IDENTIFIER_BYTES || name.contains('\0'))
        {
            return if self
                .timing_curve_theme_names
                .iter()
                .flatten()
                .any(|name| name.len() > MAX_IDENTIFIER_BYTES)
            {
                Err(Error::IdentifierTooLarge)
            } else {
                Err(Error::NulString)
            };
        }
        Ok(())
    }

    /// Borrow an optional color payload without copying it.
    #[must_use]
    pub fn color_payload(&self) -> Option<&[u8]> {
        self.color_payload.as_deref()
    }

    /// Replace or clear the optional color payload.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PayloadTooLarge`] when `payload` exceeds the bounded
    /// semantic storage budget.
    pub fn set_color_payload(&mut self, payload: Option<&[u8]>) -> Result<()> {
        self.color_payload = bounded_payload(payload)?;
        Ok(())
    }

    /// Return the optional payload for one fixed timing-curve slot.
    #[must_use]
    pub fn timing_curve_payload(&self, slot: TimingCurveSlot) -> Option<&[u8]> {
        self.timing_curve_payloads[slot.index()].as_deref()
    }

    /// Borrow all timing-curve payload slots in native wire order.
    #[must_use]
    pub fn timing_curve_payloads(&self) -> [Option<&[u8]>; TIMING_CURVE_SLOT_COUNT] {
        std::array::from_fn(|index| self.timing_curve_payloads[index].as_deref())
    }

    /// Replace or clear one fixed timing-curve payload.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PayloadTooLarge`] when `payload` exceeds the bounded
    /// semantic storage budget.
    pub fn set_timing_curve_payload(
        &mut self,
        slot: TimingCurveSlot,
        payload: Option<&[u8]>,
    ) -> Result<()> {
        self.timing_curve_payloads[slot.index()] = bounded_payload(payload)?;
        Ok(())
    }

    /// Return the native random seed, when present.
    #[must_use]
    pub const fn random_number_seed(&self) -> Option<u32> {
        self.random_number_seed
    }

    /// Replace or clear the native random seed.
    pub const fn set_random_number_seed(&mut self, value: Option<u32>) {
        self.random_number_seed = value;
    }

    /// Return the optional finite effect detail value.
    #[must_use]
    pub const fn detail(&self) -> Option<f64> {
        self.detail
    }

    /// Replace or clear the effect detail value after validating finiteness.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDetail`] when `value` is non-finite.
    pub const fn set_detail(&mut self, value: Option<f64>) -> Result<()> {
        if let Some(detail) = value
            && !detail.is_finite()
        {
            return Err(Error::InvalidDetail);
        }
        self.detail = value;
        Ok(())
    }

    /// Return the optional theme name for one fixed timing-curve slot.
    #[must_use]
    pub fn timing_curve_theme_name(&self, slot: TimingCurveSlot) -> Option<&str> {
        self.timing_curve_theme_names[slot.index()].as_deref()
    }

    /// Borrow all timing-curve theme names in native wire order.
    #[must_use]
    pub fn timing_curve_theme_names(&self) -> [Option<&str>; TIMING_CURVE_SLOT_COUNT] {
        std::array::from_fn(|index| self.timing_curve_theme_names[index].as_deref())
    }

    /// Replace or clear one timing-curve theme name.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IdentifierTooLarge`] or [`Error::NulString`] when the
    /// candidate name cannot be represented by the semantic text field.
    pub fn set_timing_curve_theme_name(
        &mut self,
        slot: TimingCurveSlot,
        value: Option<&str>,
    ) -> Result<()> {
        self.timing_curve_theme_names[slot.index()] = bounded_text(value)?;
        Ok(())
    }

    /// Return whether transition text uses right-to-left writing direction.
    #[must_use]
    pub const fn writing_direction_is_rtl(&self) -> Option<bool> {
        self.writing_direction_is_rtl
    }

    /// Replace or clear the right-to-left writing-direction flag.
    pub const fn set_writing_direction_is_rtl(&mut self, value: Option<bool>) {
        self.writing_direction_is_rtl = value;
    }
}

/// Lossless effect-specific values shared by Keynote transitions.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct CustomParameters {
    /// Native twist amount, when supplied by the effect.
    twist: Option<f32>,
    /// Native mosaic size, when supplied by the effect.
    mosaic_size: Option<u32>,
    /// Native mosaic layout discriminator.
    mosaic_type: Option<MosaicType>,
    /// Whether a bounce effect is enabled.
    bounce: Option<bool>,
    /// Whether Magic Move fades unmatched objects.
    magic_move_fade_unmatched_objects: Option<bool>,
    /// Native timing-curve discriminator.
    acceleration: Option<Acceleration>,
    /// Native matching-text delivery discriminator.
    text_delivery: Option<TextDelivery>,
    /// Whether motion blur is enabled.
    motion_blur: Option<bool>,
    /// Native travel distance, when supplied by the effect.
    travel_distance: Option<f32>,
}

impl CustomParameters {
    /// Construct empty effect-specific parameters.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            twist: None,
            mosaic_size: None,
            mosaic_type: None,
            bounce: None,
            magic_move_fade_unmatched_objects: None,
            acceleration: None,
            text_delivery: None,
            motion_blur: None,
            travel_distance: None,
        }
    }

    /// Return the optional finite twist amount.
    #[must_use]
    pub const fn twist(&self) -> Option<f32> {
        self.twist
    }

    /// Replace or clear the twist amount after validating finiteness.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] when `value` is non-finite.
    pub fn set_twist(&mut self, value: Option<f32>) -> Result<()> {
        validate_custom_float(value)?;
        self.twist = value;
        Ok(())
    }

    /// Return the optional mosaic size.
    #[must_use]
    pub const fn mosaic_size(&self) -> Option<u32> {
        self.mosaic_size
    }

    /// Replace or clear the mosaic size.
    pub const fn set_mosaic_size(&mut self, value: Option<u32>) {
        self.mosaic_size = value;
    }

    /// Return the optional mosaic layout discriminator.
    #[must_use]
    pub const fn mosaic_type(&self) -> Option<MosaicType> {
        self.mosaic_type
    }

    /// Replace or clear the mosaic layout discriminator.
    pub const fn set_mosaic_type(&mut self, value: Option<MosaicType>) {
        self.mosaic_type = value;
    }

    /// Return whether bounce is enabled.
    #[must_use]
    pub const fn bounce(&self) -> Option<bool> {
        self.bounce
    }

    /// Replace or clear the bounce flag.
    pub const fn set_bounce(&mut self, value: Option<bool>) {
        self.bounce = value;
    }

    /// Return whether unmatched Magic Move objects fade.
    #[must_use]
    pub const fn magic_move_fade_unmatched_objects(&self) -> Option<bool> {
        self.magic_move_fade_unmatched_objects
    }

    /// Replace or clear the unmatched-object fade flag.
    pub const fn set_magic_move_fade_unmatched_objects(&mut self, value: Option<bool>) {
        self.magic_move_fade_unmatched_objects = value;
    }

    /// Return the optional timing-curve discriminator.
    #[must_use]
    pub const fn acceleration(&self) -> Option<Acceleration> {
        self.acceleration
    }

    /// Replace or clear the timing-curve discriminator.
    pub const fn set_acceleration(&mut self, value: Option<Acceleration>) {
        self.acceleration = value;
    }

    /// Return the optional matching-text delivery discriminator.
    #[must_use]
    pub const fn text_delivery(&self) -> Option<TextDelivery> {
        self.text_delivery
    }

    /// Replace or clear the matching-text delivery discriminator.
    pub const fn set_text_delivery(&mut self, value: Option<TextDelivery>) {
        self.text_delivery = value;
    }

    /// Return whether motion blur is enabled.
    #[must_use]
    pub const fn motion_blur(&self) -> Option<bool> {
        self.motion_blur
    }

    /// Replace or clear the motion-blur flag.
    pub const fn set_motion_blur(&mut self, value: Option<bool>) {
        self.motion_blur = value;
    }

    /// Return the optional finite travel distance.
    #[must_use]
    pub const fn travel_distance(&self) -> Option<f32> {
        self.travel_distance
    }

    /// Replace or clear the travel distance after validating finiteness.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] when `value` is non-finite.
    pub fn set_travel_distance(&mut self, value: Option<f32>) -> Result<()> {
        validate_custom_float(value)?;
        self.travel_distance = value;
        Ok(())
    }

    /// Validate all custom floating-point values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] when a custom value is NaN or
    /// infinite. Unknown integer discriminators remain valid and lossless.
    pub fn validate(&self) -> Result<()> {
        if self
            .twist
            .into_iter()
            .chain(self.travel_distance)
            .any(|value| !value.is_finite())
        {
            return Err(Error::InvalidCustomFloat);
        }
        Ok(())
    }
}

/// Lossless, archive-free semantic settings for one Keynote slide transition.
///
/// The fields retain native presence and unknown values, while the owned
/// string and payload fields use exact-size boxed storage. Call [`Self::validate`]
/// before passing settings to an IWA adapter or publishing them from an edit.
#[derive(Debug, Clone, PartialEq)]
pub struct Settings {
    /// Native animation name, when present.
    animation_type: Option<Box<str>>,
    /// Semantic transition effect.
    effect: Option<Effect>,
    /// Transition duration in seconds.
    duration: Option<f64>,
    /// Effect-specific direction discriminator.
    direction: Option<Direction>,
    /// Delay before the transition starts, in seconds.
    delay: Option<f64>,
    /// Whether the transition starts automatically.
    is_automatic: Option<bool>,
    /// Animation-level opaque payloads and semantic values.
    animation_parameters: AnimationParameters,
    /// Effect-specific scalar values.
    custom_parameters: CustomParameters,
}

impl Settings {
    /// Construct empty transition settings.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            animation_type: None,
            effect: None,
            duration: None,
            direction: None,
            delay: None,
            is_automatic: None,
            animation_parameters: AnimationParameters::new(),
            custom_parameters: CustomParameters::new(),
        }
    }

    /// Start a checked transition-settings builder.
    #[must_use]
    pub fn builder() -> SettingsBuilder {
        SettingsBuilder::default()
    }

    /// Return the optional native animation name.
    #[must_use]
    pub fn animation_type(&self) -> Option<&str> {
        self.animation_type.as_deref()
    }

    /// Replace or clear the native animation name after validating its text
    /// and bounded storage budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IdentifierTooLarge`] or [`Error::NulString`] when the
    /// candidate name cannot be represented by the semantic text field.
    pub fn set_animation_type(&mut self, value: Option<&str>) -> Result<()> {
        self.animation_type = bounded_text(value)?;
        Ok(())
    }

    /// Return the optional semantic transition effect.
    #[must_use]
    pub fn effect(&self) -> Option<&Effect> {
        self.effect.as_ref()
    }

    /// Replace or clear the transition effect after validating its canonical
    /// representation.
    ///
    /// # Errors
    ///
    /// Returns a typed effect validation error when `value` is not a valid
    /// semantic effect.
    pub fn set_effect(&mut self, value: Option<Effect>) -> Result<()> {
        if let Some(effect) = &value {
            effect.validate()?;
        }
        self.effect = value;
        Ok(())
    }

    /// Return the optional transition duration in seconds.
    #[must_use]
    pub const fn duration(&self) -> Option<f64> {
        self.duration
    }

    /// Replace or clear the transition duration after validating it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDuration`] when `value` is non-finite or
    /// negative.
    pub fn set_duration(&mut self, value: Option<f64>) -> Result<()> {
        validate_non_negative(value, Error::InvalidDuration)?;
        self.duration = value;
        Ok(())
    }

    /// Return the optional effect-specific direction.
    #[must_use]
    pub const fn direction(&self) -> Option<Direction> {
        self.direction
    }

    /// Replace or clear the effect-specific direction.
    pub const fn set_direction(&mut self, value: Option<Direction>) {
        self.direction = value;
    }

    /// Return the optional transition delay in seconds.
    #[must_use]
    pub const fn delay(&self) -> Option<f64> {
        self.delay
    }

    /// Replace or clear the transition delay after validating it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDelay`] when `value` is non-finite or negative.
    pub fn set_delay(&mut self, value: Option<f64>) -> Result<()> {
        validate_non_negative(value, Error::InvalidDelay)?;
        self.delay = value;
        Ok(())
    }

    /// Return whether the transition starts automatically.
    #[must_use]
    pub const fn is_automatic(&self) -> Option<bool> {
        self.is_automatic
    }

    /// Replace or clear the automatic-start flag.
    pub const fn set_is_automatic(&mut self, value: Option<bool>) {
        self.is_automatic = value;
    }

    /// Borrow the animation-level parameters.
    #[must_use]
    pub const fn animation_parameters(&self) -> &AnimationParameters {
        &self.animation_parameters
    }

    /// Replace the animation-level parameters.
    ///
    /// # Errors
    ///
    /// Returns a typed animation-parameter validation error when `value` is
    /// not a valid semantic value.
    pub fn set_animation_parameters(&mut self, value: AnimationParameters) -> Result<()> {
        value.validate()?;
        self.animation_parameters = value;
        Ok(())
    }

    /// Borrow the effect-specific parameters.
    #[must_use]
    pub const fn custom_parameters(&self) -> &CustomParameters {
        &self.custom_parameters
    }

    /// Replace the effect-specific parameters.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] when `value` contains a non-finite
    /// custom floating-point value.
    pub fn set_custom_parameters(&mut self, value: CustomParameters) -> Result<()> {
        value.validate()?;
        self.custom_parameters = value;
        Ok(())
    }

    /// Return whether this setting describes a visible transition effect.
    #[must_use]
    pub fn has_effect(&self) -> bool {
        !matches!(self.effect, None | Some(Effect::None))
    }

    /// Validate every semantic field without decoding opaque payloads.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when a duration or delay is non-finite or
    /// negative, a custom value or detail is non-finite, a string contains a
    /// NUL byte, or an effect uses a non-canonical known identifier.
    pub fn validate(&self) -> Result<()> {
        if self
            .duration
            .is_some_and(|value| !value.is_finite() || value < 0.0)
        {
            return Err(Error::InvalidDuration);
        }
        if self
            .delay
            .is_some_and(|value| !value.is_finite() || value < 0.0)
        {
            return Err(Error::InvalidDelay);
        }
        if self
            .animation_type
            .as_deref()
            .is_some_and(|value| value.len() > MAX_IDENTIFIER_BYTES || value.contains('\0'))
        {
            return if self
                .animation_type
                .as_deref()
                .is_some_and(|value| value.len() > MAX_IDENTIFIER_BYTES)
            {
                Err(Error::IdentifierTooLarge)
            } else {
                Err(Error::NulString)
            };
        }
        if let Some(effect) = &self.effect {
            effect.validate()?;
        }
        self.animation_parameters.validate()?;
        self.custom_parameters.validate()
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self::new()
    }
}

/// Checked builder for [`Settings`].
#[derive(Debug, Clone, Default)]
pub struct SettingsBuilder {
    settings: Settings,
}

impl SettingsBuilder {
    /// Set or clear the native animation name.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IdentifierTooLarge`] or [`Error::NulString`] for an
    /// invalid candidate name.
    pub fn animation_type(mut self, value: Option<&str>) -> Result<Self> {
        self.settings.set_animation_type(value)?;
        Ok(self)
    }

    /// Set or clear the semantic transition effect.
    ///
    /// # Errors
    ///
    /// Returns a typed effect validation error for an invalid candidate.
    pub fn effect(mut self, value: Option<Effect>) -> Result<Self> {
        self.settings.set_effect(value)?;
        Ok(self)
    }

    /// Set or clear the transition duration in seconds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDuration`] for a non-finite or negative value.
    pub fn duration(mut self, value: Option<f64>) -> Result<Self> {
        self.settings.set_duration(value)?;
        Ok(self)
    }

    /// Set or clear the effect-specific direction.
    #[must_use]
    pub fn direction(mut self, value: Option<Direction>) -> Self {
        self.settings.set_direction(value);
        self
    }

    /// Set or clear the transition delay in seconds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDelay`] for a non-finite or negative value.
    pub fn delay(mut self, value: Option<f64>) -> Result<Self> {
        self.settings.set_delay(value)?;
        Ok(self)
    }

    /// Set or clear the automatic-start flag.
    #[must_use]
    pub fn is_automatic(mut self, value: Option<bool>) -> Self {
        self.settings.set_is_automatic(value);
        self
    }

    /// Set the animation-level parameters.
    ///
    /// # Errors
    ///
    /// Returns a typed animation-parameter validation error for an invalid
    /// candidate.
    pub fn animation_parameters(mut self, value: AnimationParameters) -> Result<Self> {
        self.settings.set_animation_parameters(value)?;
        Ok(self)
    }

    /// Set the effect-specific parameters.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] for a non-finite custom value.
    pub fn custom_parameters(mut self, value: CustomParameters) -> Result<Self> {
        self.settings.set_custom_parameters(value)?;
        Ok(self)
    }

    /// Finish the builder with a validated semantic value.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic validation error if any candidate was not
    /// representable by the transition model.
    pub fn build(self) -> Result<Settings> {
        self.settings.validate()?;
        Ok(self.settings)
    }
}

fn bounded_payload(candidate: Option<&[u8]>) -> Result<Option<Box<[u8]>>> {
    let Some(bytes) = candidate else {
        return Ok(None);
    };
    if bytes.len() > MAX_OPAQUE_PAYLOAD_BYTES {
        return Err(Error::PayloadTooLarge);
    }
    Ok(Some(bytes.to_vec().into_boxed_slice()))
}

fn bounded_text(candidate: Option<&str>) -> Result<Option<Box<str>>> {
    let Some(text) = candidate else {
        return Ok(None);
    };
    if text.len() > MAX_IDENTIFIER_BYTES {
        return Err(Error::IdentifierTooLarge);
    }
    if text.contains('\0') {
        return Err(Error::NulString);
    }
    Ok(Some(text.into()))
}

fn validate_custom_float(candidate: Option<f32>) -> Result<()> {
    if candidate.is_some_and(|number| !number.is_finite()) {
        return Err(Error::InvalidCustomFloat);
    }
    Ok(())
}

fn validate_non_negative(candidate: Option<f64>, error: Error) -> Result<()> {
    if candidate.is_some_and(|number| !number.is_finite() || number < 0.0) {
        return Err(error);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{
        AnimationParameters, CustomParameters, Effect, Settings, TextDelivery, TimingCurveSlot,
    };
    use crate::Error;

    #[test]
    fn effects_map_native_identifiers_losslessly() {
        let future = Effect::unknown("com.example.future").unwrap();
        for effect in [Effect::None, Effect::Dissolve, Effect::MagicMove, future] {
            assert_eq!(
                Effect::from_identifier(effect.identifier()).unwrap(),
                effect
            );
        }
    }

    #[test]
    fn effect_constructors_reject_non_canonical_unknown_values() {
        for identifier in [
            "none",
            "apple:dissolve",
            "apple:magic-move-implied-motion-path",
        ] {
            assert_eq!(Effect::unknown(identifier), Err(Error::NonCanonicalEffect));
        }
        assert!(
            Effect::unknown("com.example.future")
                .unwrap()
                .is_canonical()
        );
    }

    #[test]
    fn transition_owned_handles_are_compact() {
        assert_eq!(
            size_of::<Option<Box<[u8]>>>(),
            size_of::<Option<Box<str>>>(),
        );
        assert_eq!(size_of::<Option<Box<[u8]>>>(), 2 * size_of::<usize>(),);
        assert_eq!(size_of::<CustomParameters>() % 4, 0);
    }

    #[test]
    fn transition_unknown_values_and_opaque_payloads_are_lossless() {
        let mut animation_parameters = AnimationParameters::new();
        animation_parameters
            .set_color_payload(Some(&[0xff, 0x00, 0x7f]))
            .unwrap();
        animation_parameters
            .set_timing_curve_payload(TimingCurveSlot::Second, Some(&[0xff]))
            .unwrap();
        animation_parameters
            .set_timing_curve_theme_name(TimingCurveSlot::First, Some("Future Curve"))
            .unwrap();

        let mut custom_parameters = CustomParameters::new();
        custom_parameters.set_acceleration(Some(super::Acceleration::from_native(99)));
        custom_parameters.set_text_delivery(Some(TextDelivery::from_native(-7)));
        custom_parameters.set_mosaic_type(Some(super::MosaicType::from_native(u32::MAX)));

        let settings = Settings::builder()
            .animation_type(Some("future-transition"))
            .unwrap()
            .effect(Some(Effect::from_identifier("com.example.future").unwrap()))
            .unwrap()
            .animation_parameters(animation_parameters)
            .unwrap()
            .custom_parameters(custom_parameters)
            .unwrap()
            .build()
            .unwrap();

        assert_eq!(settings.validate(), Ok(()));
        assert_eq!(
            settings.effect().map(Effect::identifier),
            Some("com.example.future")
        );
        assert_eq!(
            settings.animation_parameters().color_payload(),
            Some(&[0xff, 0x00, 0x7f][..])
        );
        assert_eq!(
            settings
                .animation_parameters()
                .timing_curve_payload(TimingCurveSlot::Second),
            Some(&[0xff][..])
        );
        assert_eq!(
            settings
                .animation_parameters()
                .timing_curve_theme_name(TimingCurveSlot::First),
            Some("Future Curve")
        );
    }

    #[test]
    fn checked_setters_reject_invalid_candidates_before_mutation() {
        let mut settings = Settings::new();
        assert_eq!(
            settings.set_duration(Some(-0.1)),
            Err(Error::InvalidDuration)
        );
        assert_eq!(settings.duration(), None);
        assert_eq!(settings.set_delay(Some(f64::NAN)), Err(Error::InvalidDelay));
        assert_eq!(settings.delay(), None);
        assert_eq!(
            settings.set_animation_type(Some("bad\0name")),
            Err(Error::NulString)
        );
        assert_eq!(settings.animation_type(), None);

        let mut custom_parameters = CustomParameters::new();
        assert_eq!(
            custom_parameters.set_twist(Some(f32::INFINITY)),
            Err(Error::InvalidCustomFloat)
        );
        assert_eq!(custom_parameters.twist(), None);
        assert_eq!(
            custom_parameters.set_travel_distance(Some(f32::NAN)),
            Err(Error::InvalidCustomFloat)
        );

        let mut animation_parameters = AnimationParameters::new();
        assert_eq!(
            animation_parameters.set_detail(Some(f64::NEG_INFINITY)),
            Err(Error::InvalidDetail)
        );
        assert_eq!(animation_parameters.detail(), None);
        assert_eq!(
            animation_parameters
                .set_timing_curve_theme_name(TimingCurveSlot::Third, Some("bad\0name"),),
            Err(Error::NulString)
        );
        assert_eq!(
            animation_parameters.timing_curve_theme_name(TimingCurveSlot::Third),
            None
        );
    }

    #[test]
    fn transition_owned_storage_is_bounded_before_publication() {
        assert_eq!(Effect::unknown(""), Err(Error::EmptyIdentifier),);

        let mut settings = Settings::new();
        let long_identifier = "x".repeat(super::MAX_IDENTIFIER_BYTES + 1);
        assert_eq!(
            settings.set_animation_type(Some(&long_identifier)),
            Err(Error::IdentifierTooLarge)
        );

        let mut animation_parameters = AnimationParameters::new();
        let long_payload = vec![0; super::MAX_OPAQUE_PAYLOAD_BYTES + 1];
        assert_eq!(
            animation_parameters.set_color_payload(Some(&long_payload)),
            Err(Error::PayloadTooLarge)
        );
    }

    #[test]
    fn effect_constructors_validate_canonical_and_nul_rules() {
        assert_eq!(Effect::unknown("none"), Err(Error::NonCanonicalEffect));
        assert_eq!(Effect::unknown("future\0effect"), Err(Error::NulString));
        assert_eq!(
            Effect::unknown("com.example.future")
                .as_ref()
                .map(Effect::identifier),
            Ok("com.example.future")
        );
    }

    #[test]
    fn transition_scalars_are_compact_and_lossless() {
        assert_eq!(size_of::<super::Direction>(), 4);
        assert_eq!(size_of::<super::MosaicType>(), 4);
        assert_eq!(size_of::<super::Acceleration>(), 4);
        assert_eq!(size_of::<TextDelivery>(), 4);
        assert_eq!(super::Acceleration::EaseInOut.native_value(), 4);
        assert_eq!(super::Acceleration::from_native(19).native_value(), 19);
        assert_eq!(
            super::Acceleration::from_native(i32::MIN).native_value(),
            i32::MIN
        );
        assert_eq!(
            super::Acceleration::from_native(i32::MAX).native_value(),
            i32::MAX
        );
        assert_eq!(TextDelivery::from_native(-1).native_value(), -1);
        assert_eq!(TextDelivery::from_native(i32::MAX).native_value(), i32::MAX);
        assert_eq!(
            super::Direction::from_native(u32::MAX).native_value(),
            u32::MAX
        );
        assert_eq!(super::Acceleration::from_native(19).kind(), None);
        assert_eq!(
            TextDelivery::ByCharacter.kind(),
            Some(super::TextDeliveryKind::ByCharacter)
        );
    }
}
