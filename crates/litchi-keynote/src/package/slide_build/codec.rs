//! Wire codec of `KN.BuildArchive` / `KN.BuildChunkArchive` for the
//! slide-build family (a port of the migration host's `builds.rs`, record
//! 0372). Fields this codec does not model keep their bytes on a patch.

use litchi_iwa_common::WireLimits;

use super::{Settings, Stage};
use crate::Seconds;
use crate::build::{
    Acceleration, Action, Blink, Bounce, Effect, Emphasis, Flip, FlipDirection,
    HorizontalDirection, Jiggle, JiggleIntensity, Keyboard, KeyboardDirection, Motion, Node,
    NodeKind, ObjectEffect, Opacity, Path, Point, Pop, Pulse, Rotation, RotationDirection, Scale,
    Start, Subpath, SwooshDirection, TimingCurve, UnknownText,
};
use crate::package::structure::StructureError;
use crate::package::structure::wire::{Fields, Writer};

// KN.BuildArchive
const BUILD_DRAWABLE: u32 = 1;
const BUILD_DELIVERY: u32 = 2;
const BUILD_DURATION: u32 = 3;
const BUILD_ATTRIBUTES: u32 = 4;
const BUILD_CHUNK_ID_SEED: u32 = 5;
// KN.BuildAttributesArchive
const ATTR_DATABASE_ANIMATION_TYPE: u32 = 1;
const ATTR_DATABASE_EFFECT: u32 = 2;
const ATTR_DATABASE_DIRECTION: u32 = 3;
const ATTR_EVENT_TRIGGER: u32 = 4;
const ATTR_DATABASE_DELAY: u32 = 5;
const ATTR_DATABASE_DURATION: u32 = 8;
const ATTR_ROTATION_ANGLE: u32 = 9;
const ATTR_ROTATION_DIRECTION: u32 = 10;
const ATTR_SCALE_SIZE: u32 = 11;
const ATTR_COLOR_ALPHA: u32 = 12;
const ATTR_ACCELERATION: u32 = 13;
const ATTR_CHART_ROTATION: u32 = 17;
const ATTR_ANIMATION: u32 = 18;
const ATTR_CUSTOM_BOUNCE: u32 = 19;
const ATTR_TEXT_DELIVERY: u32 = 20;
const ATTR_DELIVERY_OPTION: u32 = 21;
const ATTR_MOTION_PATH: u32 = 22;
const ATTR_DECAY: u32 = 23;
const ATTR_REPEAT_COUNT: u32 = 24;
const ATTR_ACTION_SCALE: u32 = 25;
const ATTR_JIGGLE: u32 = 26;
const ATTR_MOTION_BLUR: u32 = 29;
const ATTR_INCLUDE_ENDPOINTS: u32 = 30;
const ATTR_SHINE: u32 = 33;
const ATTR_SCALE_AMOUNT: u32 = 34;
const ATTR_TRAVEL_DISTANCE: u32 = 35;
const ATTR_CURSOR: u32 = 36;
const ATTR_ALIGN_TO_PATH: u32 = 37;
// KN.AnimationAttributesArchive
const ANIM_TYPE: u32 = 1;
const ANIM_EFFECT: u32 = 2;
const ANIM_DURATION: u32 = 3;
const ANIM_DIRECTION: u32 = 4;
const ANIM_DELAY: u32 = 5;
const ANIM_AUTOMATIC: u32 = 6;
const ANIM_TIMING_CURVE: u32 = 8;
const ANIM_RANDOM_SEED: u32 = 11;
const ANIM_TIMING_CURVE_THEME: u32 = 13;
const ANIM_RTL: u32 = 16;
// KN.BuildChunkArchive
const CHUNK_BUILD: u32 = 1;
const CHUNK_DELAY: u32 = 3;
const CHUNK_DURATION: u32 = 4;
const CHUNK_AUTOMATIC: u32 = 5;
const CHUNK_REFERENT: u32 = 6;
const CHUNK_IDENTIFIER: u32 = 7;
const CHUNK_BUILD_ID: u32 = 8;
// TSD.PathSourceArchive / EditableBezierPathSourceArchive
const PATH_HORIZONTAL_FLIP: u32 = 1;
const PATH_VERTICAL_FLIP: u32 = 2;
const PATH_EDITABLE: u32 = 8;
const EDITABLE_SUBPATHS: u32 = 1;
const EDITABLE_NATURAL_SIZE: u32 = 2;

const APPEAR: &str = "apple:bc-appear";
const DISSOLVE: &str = "apple:bc-dissolve";
const DEFAULT_DELIVERY: &str = "All at Once";
const TEXT_DELIVERY_BY_OBJECT: i32 = 1;
const DELIVERY_OPTION_FORWARD: i32 = 1;
const CLOCKWISE: i32 = 31;
const COUNTERCLOCKWISE: i32 = 32;

type Result<T> = std::result::Result<T, StructureError>;

fn invalid() -> StructureError {
    StructureError::Unexpected("Keynote build payload")
}

/// Raw native values a build carries beside its typed settings; kept on a
/// patch, defaulted (as the host's constructors do) on creation.
#[derive(Debug, Clone, PartialEq)]
pub(in crate::package) struct Native {
    pub(super) delivery: Box<str>,
    /// The native effect identifier, kept when the typed effect is the one
    /// it decodes to.
    pub(super) effect_identifier: Option<Box<str>>,
    pub(super) direction: Option<u32>,
    pub(super) text_delivery: Option<i32>,
    pub(super) delivery_option: Option<i32>,
    pub(super) event_trigger: Option<u32>,
    pub(super) bounce: Option<bool>,
    pub(super) motion_blur: Option<bool>,
    pub(super) include_endpoints: Option<bool>,
    pub(super) shine: Option<bool>,
    pub(super) scale_amount: Option<f64>,
    pub(super) travel_distance: Option<f64>,
}

impl Native {
    /// The host constructors' defaults for `stage` and `effect`.
    pub(super) fn defaults(stage: &Stage, effect: &Effect) -> Self {
        let text = matches!(stage, Stage::In | Stage::Out)
            && matches!(
                effect,
                Effect::Appear | Effect::Dissolve | Effect::Object(ObjectEffect::Dissolve)
            );
        Self {
            delivery: DEFAULT_DELIVERY.into(),
            effect_identifier: None,
            direction: None,
            text_delivery: text.then_some(TEXT_DELIVERY_BY_OBJECT),
            delivery_option: text.then_some(DELIVERY_OPTION_FORWARD),
            event_trigger: Some(1),
            bounce: None,
            motion_blur: None,
            include_endpoints: None,
            shine: None,
            scale_amount: None,
            travel_distance: None,
        }
    }
}

/// The native identifier of a typed effect.
pub(super) fn effect_identifier(effect: &Effect, native: &Native) -> String {
    if let Some(raw) = &native.effect_identifier
        && decode_effect_class(raw) == effect_class(effect)
    {
        return raw.to_string();
    }
    match effect {
        Effect::Appear => APPEAR.to_owned(),
        Effect::Dissolve => DISSOLVE.to_owned(),
        other => other.identifier().to_owned(),
    }
}

/// A comparison class of effects that share a native identifier family.
fn effect_class(effect: &Effect) -> String {
    match effect {
        Effect::Appear => "appear".to_owned(),
        Effect::Dissolve => "dissolve".to_owned(),
        other => other.identifier().to_owned(),
    }
}

fn decode_effect_class(identifier: &str) -> String {
    match Effect::from_identifier(identifier) {
        Ok(Effect::Appear) => "appear".to_owned(),
        Ok(Effect::Dissolve) => "dissolve".to_owned(),
        _ => identifier.to_owned(),
    }
}

/// The native direction a typed effect writes, or the raw one it keeps.
pub(super) fn direction(effect: &Effect, native: &Native) -> Option<u32> {
    match effect {
        Effect::Object(object) => match object {
            ObjectEffect::Dissolve
            | ObjectEffect::Shimmer
            | ObjectEffect::Swoosh(SwooshDirection::Center) => None,
            ObjectEffect::Skid(HorizontalDirection::LeftToRight)
            | ObjectEffect::Swoosh(SwooshDirection::FromLeft)
            | ObjectEffect::Trace(HorizontalDirection::LeftToRight) => Some(11),
            ObjectEffect::Skid(HorizontalDirection::RightToLeft)
            | ObjectEffect::Swoosh(SwooshDirection::FromRight)
            | ObjectEffect::Trace(HorizontalDirection::RightToLeft) => Some(12),
        },
        Effect::Emphasis(Emphasis::Flip(flip)) => Some(match flip.direction() {
            FlipDirection::LeftToRight => 11,
            FlipDirection::RightToLeft => 12,
        }),
        Effect::Emphasis(_) | Effect::Action(_) => None,
        Effect::Keyboard(keyboard) => Some(match keyboard.direction() {
            KeyboardDirection::Forward => 111,
            KeyboardDirection::Backward => 112,
        }),
        _ => native.direction,
    }
}

/// Whether the effect is a typed action or emphasis (host
/// `is_typed_action_effect`).
pub(super) fn is_typed_action(effect: &Effect) -> bool {
    matches!(
        effect,
        Effect::Action(Action::Rotate(_) | Action::Scale(_) | Action::Opacity(_) | Action::Move(_))
            | Effect::Emphasis(
                Emphasis::Blink(_)
                    | Emphasis::Bounce(_)
                    | Emphasis::Flip(_)
                    | Emphasis::Jiggle(_)
                    | Emphasis::Pop(_)
                    | Emphasis::Pulse(_)
            )
    )
}

/// The acceleration of a typed action, if the effect is one.
pub(super) fn action_acceleration(effect: &Effect) -> Option<Acceleration> {
    match effect {
        Effect::Action(Action::Rotate(rotation)) => Some(rotation.acceleration()),
        Effect::Action(Action::Scale(scale)) => Some(scale.acceleration()),
        Effect::Action(Action::Opacity(opacity)) => Some(opacity.acceleration()),
        Effect::Action(Action::Move(motion)) => Some(motion.acceleration()),
        _ => None,
    }
}

fn stage_text(stage: &Stage) -> &str {
    match stage {
        Stage::In => "In",
        Stage::Out => "Out",
        Stage::Action => "Action",
        Stage::Unknown(text) => text.as_str(),
    }
}

fn point_writer(x: f64, y: f64) -> Writer {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "TSP.Point stores 32-bit floats"
    )]
    Writer::new().float(1, x as f32).float(2, y as f32)
}

/// Encode a motion or timing path as `TSD.PathSourceArchive`.
pub(super) fn path_source(path: &Path) -> Writer {
    let mut editable = Writer::new();
    for subpath in path.subpaths() {
        let mut encoded = Writer::new();
        for node in subpath.nodes() {
            let kind = match node.kind() {
                NodeKind::Sharp => 1,
                NodeKind::Bezier => 2,
                NodeKind::Smooth => 3,
            };
            let point = |p: Point| point_writer(p.x(), p.y());
            encoded = encoded.message(
                1,
                &Writer::new()
                    .message(1, &point(node.in_control_point()))
                    .message(2, &point(node.point()))
                    .message(3, &point(node.out_control_point()))
                    .varint(4, kind),
            );
        }
        encoded = encoded.boolean(2, subpath.closed());
        editable = editable.message(EDITABLE_SUBPATHS, &encoded);
    }
    editable = editable.message(
        EDITABLE_NATURAL_SIZE,
        &point_writer(path.natural_width(), path.natural_height()),
    );
    Writer::new()
        .boolean(PATH_HORIZONTAL_FLIP, path.horizontal_flip())
        .boolean(PATH_VERTICAL_FLIP, path.vertical_flip())
        .message(PATH_EDITABLE, &editable)
}

fn read_point(payload: &[u8], wire: WireLimits) -> Result<Point> {
    let fields = Fields::parse(payload, wire)?;
    let x = fields
        .unique(1)?
        .map(|f| f.float())
        .transpose()?
        .unwrap_or(0.0);
    let y = fields
        .unique(2)?
        .map(|f| f.float())
        .transpose()?
        .unwrap_or(0.0);
    Point::new(f64::from(x), f64::from(y)).map_err(|_| invalid())
}

/// Decode a `TSD.PathSourceArchive` holding only an editable Bézier path.
pub(super) fn read_path(payload: &[u8], wire: WireLimits) -> Result<Option<Path>> {
    let fields = Fields::parse(payload, wire)?;
    for other in [3, 4, 5, 6, 7] {
        if fields.unique(other)?.is_some() {
            return Ok(None);
        }
    }
    let Some(editable) = fields.unique(PATH_EDITABLE)? else {
        return Ok(None);
    };
    let flag = |number| -> Result<bool> {
        Ok(fields
            .unique(number)?
            .map(|f| f.varint())
            .transpose()?
            .unwrap_or(0)
            != 0)
    };
    let editable = Fields::parse(editable.message()?, wire)?;
    let Some(size) = editable.unique(EDITABLE_NATURAL_SIZE)? else {
        return Ok(None);
    };
    let size = read_point(size.message()?, wire)?;
    let mut subpaths = Vec::new();
    for subpath in editable.all(EDITABLE_SUBPATHS) {
        let subpath = Fields::parse(subpath.message()?, wire)?;
        let mut nodes = Vec::new();
        for node in subpath.all(1) {
            let node = Fields::parse(node.message()?, wire)?;
            let point = |number| -> Result<Point> {
                read_point(node.unique(number)?.ok_or_else(invalid)?.message()?, wire)
            };
            let kind = match node.unique(4)?.map(|f| f.varint()).transpose()? {
                Some(1) => NodeKind::Sharp,
                Some(2) => NodeKind::Bezier,
                Some(3) => NodeKind::Smooth,
                _ => return Ok(None),
            };
            nodes.push(Node::new(point(1)?, point(2)?, point(3)?, kind));
        }
        let closed = subpath
            .unique(2)?
            .map(|f| f.varint())
            .transpose()?
            .unwrap_or(0)
            != 0;
        match Subpath::new(nodes, closed) {
            Ok(subpath) => subpaths.push(subpath),
            Err(_) => return Ok(None),
        }
    }
    Ok(Path::new(
        subpaths,
        size.x(),
        size.y(),
        flag(PATH_HORIZONTAL_FLIP)?,
        flag(PATH_VERTICAL_FLIP)?,
    )
    .ok())
}

fn double(fields: &Fields, number: u32) -> Result<Option<f64>> {
    fields
        .unique(number)?
        .map(|field| {
            let payload = field.payload();
            let bytes: [u8; 8] = payload.try_into().map_err(|_| invalid())?;
            Ok(f64::from_le_bytes(bytes))
        })
        .transpose()
}

fn varint(fields: &Fields, number: u32) -> Result<Option<u64>> {
    fields
        .unique(number)?
        .map(|field| field.varint())
        .transpose()
}

fn string(fields: &Fields, number: u32) -> Result<Option<String>> {
    fields
        .unique(number)?
        .map(|field| {
            std::str::from_utf8(field.message()?)
                .map(str::to_owned)
                .map_err(|_| invalid())
        })
        .transpose()
}

fn as_u32(value: Option<u64>) -> Option<u32> {
    value.and_then(|value| u32::try_from(value).ok())
}

#[allow(clippy::cast_possible_truncation, reason = "protobuf int32 varints")]
fn as_i32(value: Option<u64>) -> Option<i32> {
    value.map(|value| value as i32)
}

/// One timing chunk's start fields.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct Chunk {
    pub(super) delay: Option<f64>,
    pub(super) duration: Option<f64>,
    pub(super) automatic: Option<bool>,
    pub(super) referent: Option<bool>,
}

pub(super) fn read_chunk(payload: &[u8], wire: WireLimits) -> Result<(Option<u64>, Chunk)> {
    let fields = Fields::parse(payload, wire)?;
    let build = fields
        .unique(CHUNK_BUILD)?
        .map(|field| crate::package::structure::wire::reference_identifier(field.message()?, wire))
        .transpose()?;
    Ok((
        build,
        Chunk {
            delay: double(&fields, CHUNK_DELAY)?,
            duration: double(&fields, CHUNK_DURATION)?,
            automatic: varint(&fields, CHUNK_AUTOMATIC)?.map(|value| value != 0),
            referent: varint(&fields, CHUNK_REFERENT)?.map(|value| value != 0),
        },
    ))
}

/// The drawable a build animates.
pub(super) fn build_drawable(payload: &[u8], wire: WireLimits) -> Result<Option<u64>> {
    let fields = Fields::parse(payload, wire)?;
    fields
        .unique(BUILD_DRAWABLE)?
        .map(|field| crate::package::structure::wire::reference_identifier(field.message()?, wire))
        .transpose()
}

fn typed(result: crate::Result<Effect>) -> Option<Effect> {
    result.ok()
}

/// Decode a build into settings (host `build_settings`).
pub(super) fn read_settings(
    payload: &[u8],
    first_chunk: Option<Chunk>,
    starts_slide_events: bool,
    wire: WireLimits,
) -> Result<Settings> {
    let build = Fields::parse(payload, wire)?;
    let delivery = string(&build, BUILD_DELIVERY)?.unwrap_or_default();
    let attributes = Fields::parse(
        build
            .unique(BUILD_ATTRIBUTES)?
            .ok_or_else(invalid)?
            .message()?,
        wire,
    )?;
    let animation = attributes
        .unique(ATTR_ANIMATION)?
        .map(|field| Fields::parse(field.message()?, wire))
        .transpose()?;
    let from_animation = |number| -> Result<Option<String>> {
        Ok(match &animation {
            Some(animation) => string(animation, number)?,
            None => None,
        })
    };
    let animation_type = match from_animation(ANIM_TYPE)? {
        Some(value) => value,
        None => string(&attributes, ATTR_DATABASE_ANIMATION_TYPE)?.unwrap_or_default(),
    };
    let effect_text = match from_animation(ANIM_EFFECT)? {
        Some(value) => value,
        None => string(&attributes, ATTR_DATABASE_EFFECT)?.unwrap_or_default(),
    };
    let direction = match &animation {
        Some(animation) => varint(animation, ANIM_DIRECTION)?,
        None => None,
    }
    .or(varint(&attributes, ATTR_DATABASE_DIRECTION)?);
    let direction = as_u32(direction);
    let acceleration =
        as_i32(varint(&attributes, ATTR_ACCELERATION)?).map(Acceleration::from_native);
    let stage = match animation_type.as_str() {
        "In" => Stage::In,
        "Out" => Stage::Out,
        "Action" => Stage::Action,
        other => Stage::Unknown(UnknownText::new(other).map_err(|_| invalid())?),
    };
    let repeat = as_u32(varint(&attributes, ATTR_REPEAT_COUNT)?);
    let decay = varint(&attributes, ATTR_DECAY)?.map(|value| value != 0);
    let action_scale = double(&attributes, ATTR_ACTION_SCALE)?;
    let in_out = matches!(stage, Stage::In | Stage::Out);
    let effect = match effect_text.as_str() {
        "apple:action-rotation" => (|| {
            let degrees = double(&attributes, ATTR_ROTATION_ANGLE).ok()??;
            let direction = match as_i32(varint(&attributes, ATTR_ROTATION_DIRECTION).ok()?)? {
                CLOCKWISE => RotationDirection::Clockwise,
                COUNTERCLOCKWISE => RotationDirection::Counterclockwise,
                _ => return None,
            };
            Some(Effect::Action(Action::Rotate(
                Rotation::new(degrees, direction, acceleration?).ok()?,
            )))
        })(),
        "apple:action-scale" => double(&attributes, ATTR_SCALE_SIZE)?
            .zip(acceleration)
            .and_then(|(factor, acceleration)| Scale::new(factor, acceleration).ok())
            .map(|scale| Effect::Action(Action::Scale(scale))),
        "apple:action-opacity" => double(&attributes, ATTR_COLOR_ALPHA)?
            .zip(acceleration)
            .and_then(|(percent, acceleration)| Opacity::new(percent, acceleration).ok())
            .map(|opacity| Effect::Action(Action::Opacity(opacity))),
        "apple:action-motion-path" => match (attributes.unique(ATTR_MOTION_PATH)?, acceleration) {
            (Some(path), Some(acceleration)) => read_path(path.message()?, wire)?.map(|path| {
                let align = varint(&attributes, ATTR_ALIGN_TO_PATH).ok().flatten() == Some(1);
                Effect::Action(Action::Move(Motion::new(path, align, acceleration)))
            }),
            _ => None,
        },
        "apple:action-blink" => repeat
            .zip(decay)
            .and_then(|(count, fade)| Blink::new(count, fade).ok())
            .map(|blink| Effect::Emphasis(Emphasis::Blink(blink))),
        "apple:action-bounce" => repeat
            .zip(decay)
            .and_then(|(count, decay)| Bounce::new(count, decay).ok())
            .map(|bounce| Effect::Emphasis(Emphasis::Bounce(bounce))),
        "apple:action-flip" => repeat
            .zip(direction.and_then(|value| match value {
                11 => Some(FlipDirection::LeftToRight),
                12 => Some(FlipDirection::RightToLeft),
                _ => None,
            }))
            .and_then(|(count, direction)| Flip::new(count, direction).ok())
            .map(|flip| Effect::Emphasis(Emphasis::Flip(flip))),
        "apple:action-jiggle" => as_i32(varint(&attributes, ATTR_JIGGLE)?)
            .and_then(|value| match value {
                0 => Some(JiggleIntensity::Small),
                1 => Some(JiggleIntensity::Medium),
                2 => Some(JiggleIntensity::Large),
                _ => None,
            })
            .map(|intensity| Effect::Emphasis(Emphasis::Jiggle(Jiggle::new(intensity)))),
        "apple:action-pop" => action_scale
            .and_then(|scale| Pop::new(scale).ok())
            .map(|pop| Effect::Emphasis(Emphasis::Pop(pop))),
        "apple:action-pulse" => repeat
            .zip(action_scale)
            .and_then(|(count, scale)| Pulse::new(count, scale).ok())
            .map(|pulse| Effect::Emphasis(Emphasis::Pulse(pulse))),
        "apple:keyboard" if in_out => direction
            .and_then(|value| match value {
                111 => Some(KeyboardDirection::Forward),
                112 => Some(KeyboardDirection::Backward),
                _ => None,
            })
            .zip(varint(&attributes, ATTR_CURSOR)?.map(|value| value != 0))
            .map(|(direction, cursor)| Effect::Keyboard(Keyboard::new(direction, cursor))),
        text if in_out => {
            object_effect(text, direction)
                .map(Effect::Object)
                .or_else(|| match typed(Effect::from_identifier(text)) {
                    Some(Effect::Unknown(_)) | None => None,
                    Some(effect) => Some(effect),
                })
        },
        _ => None,
    };
    let effect = match effect {
        Some(effect) => effect,
        None if effect_text.is_empty() => {
            return Err(invalid());
        },
        None => Effect::Unknown(UnknownText::new(effect_text.as_str()).map_err(|_| invalid())?),
    };
    let timing_curve = match (&animation, acceleration) {
        (Some(animation), Some(acceleration))
            if acceleration == Acceleration::Custom && action_acceleration(&effect).is_some() =>
        {
            match animation.unique(ANIM_TIMING_CURVE)? {
                Some(curve) => read_path(curve.message()?, wire)?
                    .and_then(|path| TimingCurve::from_path(path).ok()),
                None => None,
            }
        },
        _ => None,
    };
    let animation_double = |number| -> Result<Option<f64>> {
        Ok(match &animation {
            Some(animation) => double(animation, number)?,
            None => None,
        })
    };
    let duration = animation_double(ANIM_DURATION)?
        .or(double(&attributes, ATTR_DATABASE_DURATION)?)
        .or(double(&build, BUILD_DURATION)?)
        .or(first_chunk.and_then(|chunk| chunk.duration))
        .unwrap_or_default();
    let delay = animation_double(ANIM_DELAY)?
        .or(double(&attributes, ATTR_DATABASE_DELAY)?)
        .or(first_chunk.and_then(|chunk| chunk.delay))
        .unwrap_or_default();
    let automatic = first_chunk
        .and_then(|chunk| chunk.automatic)
        .or(match &animation {
            Some(animation) => varint(animation, ANIM_AUTOMATIC)?.map(|value| value != 0),
            None => None,
        })
        .unwrap_or(false);
    let referent = first_chunk.and_then(|chunk| chunk.referent).unwrap_or(true);
    let start = match (automatic, referent, starts_slide_events) {
        (false, _, _) => Start::OnClick,
        (true, false, _) => Start::WithPrevious,
        (true, true, true) => Start::AfterTransition,
        (true, true, false) => Start::AfterPrevious,
    };
    let flag = |number| -> Result<Option<bool>> {
        Ok(varint(&attributes, number)?.map(|value| value != 0))
    };
    Ok(Settings {
        stage,
        effect,
        duration: Seconds::new(duration).map_err(|_| invalid())?,
        delay: Seconds::new(delay).map_err(|_| invalid())?,
        start,
        timing_curve,
        native: Native {
            delivery: delivery.into(),
            effect_identifier: Some(effect_text.into()),
            direction,
            text_delivery: as_i32(varint(&attributes, ATTR_TEXT_DELIVERY)?),
            delivery_option: as_i32(varint(&attributes, ATTR_DELIVERY_OPTION)?),
            event_trigger: as_u32(varint(&attributes, ATTR_EVENT_TRIGGER)?),
            bounce: flag(ATTR_CUSTOM_BOUNCE)?,
            motion_blur: flag(ATTR_MOTION_BLUR)?,
            include_endpoints: flag(ATTR_INCLUDE_ENDPOINTS)?,
            shine: flag(ATTR_SHINE)?,
            scale_amount: double(&attributes, ATTR_SCALE_AMOUNT)?,
            travel_distance: double(&attributes, ATTR_TRAVEL_DISTANCE)?,
        },
    })
}

fn object_effect(identifier: &str, direction: Option<u32>) -> Option<ObjectEffect> {
    use HorizontalDirection::{LeftToRight, RightToLeft};
    Some(match (identifier, direction) {
        ("apple:dissolve character", None) => ObjectEffect::Dissolve,
        ("com.apple.iWork.Keynote.KLNShimmer", None) => ObjectEffect::Shimmer,
        ("com.apple.iWork.Keynote.KNBuildSkidByCharacter", None | Some(11)) => {
            ObjectEffect::Skid(LeftToRight)
        },
        ("com.apple.iWork.Keynote.KNBuildSkidByCharacter", Some(12)) => {
            ObjectEffect::Skid(RightToLeft)
        },
        ("com.apple.iWork.Keynote.BLTSwoosh", None) => {
            ObjectEffect::Swoosh(SwooshDirection::Center)
        },
        ("com.apple.iWork.Keynote.BLTSwoosh", Some(11)) => {
            ObjectEffect::Swoosh(SwooshDirection::FromLeft)
        },
        ("com.apple.iWork.Keynote.BLTSwoosh", Some(12)) => {
            ObjectEffect::Swoosh(SwooshDirection::FromRight)
        },
        ("com.apple.iWork.Keynote.Trace", None | Some(11)) => ObjectEffect::Trace(LeftToRight),
        ("com.apple.iWork.Keynote.Trace", Some(12)) => ObjectEffect::Trace(RightToLeft),
        _ => return None,
    })
}

/// The typed action/emphasis attribute values a settings value writes.
struct Typed {
    rotation: Option<(f64, i32)>,
    scale: Option<f64>,
    alpha: Option<f64>,
    motion: Option<(Path, bool)>,
    acceleration: Option<i32>,
    decay: Option<bool>,
    repeat: Option<u32>,
    action_scale: Option<f64>,
    jiggle: Option<i32>,
    cursor: Option<bool>,
}

fn typed_values(effect: &Effect) -> Typed {
    let mut typed = Typed {
        rotation: None,
        scale: None,
        alpha: None,
        motion: None,
        acceleration: action_acceleration(effect).map(Acceleration::native_value),
        decay: None,
        repeat: None,
        action_scale: None,
        jiggle: None,
        cursor: None,
    };
    match effect {
        Effect::Action(Action::Rotate(rotation)) => {
            let direction = match rotation.direction() {
                RotationDirection::Clockwise => CLOCKWISE,
                RotationDirection::Counterclockwise => COUNTERCLOCKWISE,
            };
            typed.rotation = Some((rotation.total_degrees(), direction));
        },
        Effect::Action(Action::Scale(scale)) => typed.scale = Some(scale.factor()),
        Effect::Action(Action::Opacity(opacity)) => typed.alpha = Some(opacity.percent()),
        Effect::Action(Action::Move(motion)) => {
            typed.motion = Some((motion.path().clone(), motion.align_to_path()));
        },
        Effect::Emphasis(Emphasis::Blink(blink)) => {
            typed.decay = Some(blink.fade());
            typed.repeat = Some(blink.repeat_count());
        },
        Effect::Emphasis(Emphasis::Bounce(bounce)) => {
            typed.decay = Some(bounce.decay());
            typed.repeat = Some(bounce.repeat_count());
        },
        Effect::Emphasis(Emphasis::Flip(flip)) => typed.repeat = Some(flip.repeat_count()),
        Effect::Emphasis(Emphasis::Jiggle(jiggle)) => {
            typed.jiggle = Some(match jiggle.intensity() {
                JiggleIntensity::Small => 0,
                JiggleIntensity::Medium => 1,
                JiggleIntensity::Large => 2,
            });
        },
        Effect::Emphasis(Emphasis::Pop(pop)) => typed.action_scale = Some(pop.scale_percent()),
        Effect::Emphasis(Emphasis::Pulse(pulse)) => {
            typed.repeat = Some(pulse.repeat_count());
            typed.action_scale = Some(pulse.scale_percent());
        },
        Effect::Keyboard(keyboard) => typed.cursor = Some(keyboard.show_cursor()),
        _ => {},
    }
    typed
}

/// Encode a fresh build (host `new_build_archive`).
pub(super) fn encode_build(drawable: u64, settings: &Settings, seed: u32) -> Vec<u8> {
    let native = &settings.native;
    let typed = typed_values(&settings.effect);
    let mut animation = Writer::new()
        .string(ANIM_TYPE, stage_text(&settings.stage))
        .string(ANIM_EFFECT, &effect_identifier(&settings.effect, native))
        .double(ANIM_DURATION, settings.duration.as_f64());
    if let Some(direction) = direction(&settings.effect, native) {
        animation = animation.varint(ANIM_DIRECTION, u64::from(direction));
    }
    animation = animation.double(ANIM_DELAY, settings.delay.as_f64());
    if let Some(curve) = &settings.timing_curve {
        animation = animation.message(ANIM_TIMING_CURVE, &path_source(curve.path()));
    }
    animation = animation
        .varint(ANIM_RANDOM_SEED, u64::from(seed))
        .boolean(ANIM_RTL, false);
    let mut attributes = Writer::new();
    if let Some(trigger) = native.event_trigger {
        attributes = attributes.varint(ATTR_EVENT_TRIGGER, u64::from(trigger));
    }
    if let Some((angle, direction)) = typed.rotation {
        attributes = attributes
            .double(ATTR_ROTATION_ANGLE, angle)
            .varint(ATTR_ROTATION_DIRECTION, u64::from(direction.unsigned_abs()));
    }
    if let Some(scale) = typed.scale {
        attributes = attributes.double(ATTR_SCALE_SIZE, scale);
    }
    if let Some(alpha) = typed.alpha {
        attributes = attributes.double(ATTR_COLOR_ALPHA, alpha);
    }
    if let Some(acceleration) = typed.acceleration {
        attributes = attributes.varint(ATTR_ACCELERATION, int_varint(acceleration));
    }
    attributes = attributes.double(ATTR_CHART_ROTATION, 60.0);
    attributes = attributes.message(ATTR_ANIMATION, &animation);
    if let Some(bounce) = native.bounce {
        attributes = attributes.boolean(ATTR_CUSTOM_BOUNCE, bounce);
    }
    if let Some(value) = native.text_delivery {
        attributes = attributes.varint(ATTR_TEXT_DELIVERY, int_varint(value));
    }
    if let Some(value) = native.delivery_option {
        attributes = attributes.varint(ATTR_DELIVERY_OPTION, int_varint(value));
    }
    if let Some((path, _)) = &typed.motion {
        attributes = attributes.message(ATTR_MOTION_PATH, &path_source(path));
    }
    if let Some(decay) = typed.decay {
        attributes = attributes.boolean(ATTR_DECAY, decay);
    }
    if let Some(repeat) = typed.repeat {
        attributes = attributes.varint(ATTR_REPEAT_COUNT, u64::from(repeat));
    }
    if let Some(scale) = typed.action_scale {
        attributes = attributes.double(ATTR_ACTION_SCALE, scale);
    }
    if let Some(jiggle) = typed.jiggle {
        attributes = attributes.varint(ATTR_JIGGLE, int_varint(jiggle));
    }
    for (number, value) in [
        (ATTR_MOTION_BLUR, native.motion_blur),
        (ATTR_INCLUDE_ENDPOINTS, native.include_endpoints),
        (ATTR_SHINE, native.shine),
    ] {
        if let Some(value) = value {
            attributes = attributes.boolean(number, value);
        }
    }
    for (number, value) in [
        (ATTR_SCALE_AMOUNT, native.scale_amount),
        (ATTR_TRAVEL_DISTANCE, native.travel_distance),
    ] {
        if let Some(value) = value {
            attributes = attributes.double(number, value);
        }
    }
    if let Some(cursor) = typed.cursor {
        attributes = attributes.boolean(ATTR_CURSOR, cursor);
    }
    if typed.motion.as_ref().is_some_and(|(_, align)| *align) {
        attributes = attributes.boolean(ATTR_ALIGN_TO_PATH, true);
    }
    Writer::new()
        .reference(BUILD_DRAWABLE, drawable)
        .string(BUILD_DELIVERY, &native.delivery)
        .double(BUILD_DURATION, 0.0)
        .message(BUILD_ATTRIBUTES, &attributes)
        .varint(BUILD_CHUNK_ID_SEED, 1)
        .finish()
}

#[allow(clippy::cast_sign_loss, reason = "protobuf int32 varint encoding")]
fn int_varint(value: i32) -> u64 {
    i64::from(value) as u64
}

/// `(automatic, referent)` of a start (host `build_start_fields`).
pub(super) const fn start_fields(start: Start) -> (bool, bool) {
    match start {
        Start::OnClick => (false, true),
        Start::AfterTransition | Start::AfterPrevious => (true, true),
        Start::WithPrevious => (true, false),
    }
}

/// Encode a fresh timing chunk (host `new_build_chunk`).
pub(super) fn encode_chunk(build: u64, uuid: (u64, u64), settings: &Settings) -> Vec<u8> {
    let (automatic, referent) = start_fields(settings.start);
    let uuid = Writer::new().varint(1, uuid.0).varint(2, uuid.1);
    Writer::new()
        .reference(CHUNK_BUILD, build)
        .double(CHUNK_DELAY, settings.delay.as_f64())
        .double(CHUNK_DURATION, settings.duration.as_f64())
        .boolean(CHUNK_AUTOMATIC, automatic)
        .boolean(CHUNK_REFERENT, referent)
        .message(
            CHUNK_IDENTIFIER,
            &Writer::new().message(1, &uuid).varint(2, 1),
        )
        .message(CHUNK_BUILD_ID, &uuid)
        .finish()
}

fn set_or_remove(fields: &mut Fields, number: u32, value: Option<FieldValue<'_>>) {
    match value {
        None => {
            fields.remove(number);
        },
        Some(FieldValue::Varint(value)) => fields.set_varint(number, value),
        Some(FieldValue::Double(value)) => fields.set_double(number, value),
        Some(FieldValue::Bytes(value)) => fields.set_message(number, value),
    }
}

enum FieldValue<'a> {
    Varint(u64),
    Double(f64),
    Bytes(&'a [u8]),
}

fn keep_or(fields: &mut Fields, number: u32, value: Option<FieldValue<'_>>) {
    // The host's patch writes a value or clears the field; same here.
    set_or_remove(fields, number, value);
}

/// Patch an existing build with `settings` (host
/// `patch_build_settings_wire`): typed fields are written or cleared,
/// everything else keeps its bytes.
pub(super) fn patch_build(
    original: &[u8],
    settings: &Settings,
    wire: WireLimits,
) -> Result<Vec<u8>> {
    let native = &settings.native;
    let mut build = Fields::parse(original, wire)?;
    let attributes_raw = build
        .unique(BUILD_ATTRIBUTES)?
        .ok_or_else(invalid)?
        .message()?
        .to_vec();
    let mut attributes = Fields::parse(&attributes_raw, wire)?;
    let animation_raw = attributes
        .unique(ATTR_ANIMATION)?
        .ok_or_else(invalid)?
        .message()?
        .to_vec();
    let mut animation = Fields::parse(&animation_raw, wire)?;
    build.set_message(BUILD_DELIVERY, native.delivery.as_bytes());
    keep_or(
        &mut attributes,
        ATTR_EVENT_TRIGGER,
        native
            .event_trigger
            .map(|v| FieldValue::Varint(u64::from(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_TEXT_DELIVERY,
        native
            .text_delivery
            .map(|v| FieldValue::Varint(int_varint(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_DELIVERY_OPTION,
        native
            .delivery_option
            .map(|v| FieldValue::Varint(int_varint(v))),
    );
    let opaque_action =
        matches!(settings.stage, Stage::Action) && !is_typed_action(&settings.effect);
    let typed = typed_values(&settings.effect);
    if !opaque_action {
        keep_or(
            &mut attributes,
            ATTR_ROTATION_ANGLE,
            typed.rotation.map(|(a, _)| FieldValue::Double(a)),
        );
        keep_or(
            &mut attributes,
            ATTR_ROTATION_DIRECTION,
            typed
                .rotation
                .map(|(_, d)| FieldValue::Varint(int_varint(d))),
        );
        keep_or(
            &mut attributes,
            ATTR_SCALE_SIZE,
            typed.scale.map(FieldValue::Double),
        );
        keep_or(
            &mut attributes,
            ATTR_COLOR_ALPHA,
            typed.alpha.map(FieldValue::Double),
        );
        keep_or(
            &mut attributes,
            ATTR_ACCELERATION,
            typed
                .acceleration
                .map(|v| FieldValue::Varint(int_varint(v))),
        );
        let path = typed
            .motion
            .as_ref()
            .map(|(path, _)| path_source(path).finish());
        keep_or(
            &mut attributes,
            ATTR_MOTION_PATH,
            path.as_deref().map(FieldValue::Bytes),
        );
        let had_align = varint(&attributes, ATTR_ALIGN_TO_PATH)?;
        let align = typed.motion.as_ref().and_then(|(_, align)| {
            if *align {
                Some(1)
            } else if had_align == Some(0) {
                Some(0)
            } else {
                None
            }
        });
        keep_or(
            &mut attributes,
            ATTR_ALIGN_TO_PATH,
            align.map(FieldValue::Varint),
        );
        keep_or(
            &mut attributes,
            ATTR_DECAY,
            typed.decay.map(|v| FieldValue::Varint(u64::from(v))),
        );
        keep_or(
            &mut attributes,
            ATTR_REPEAT_COUNT,
            typed.repeat.map(|v| FieldValue::Varint(u64::from(v))),
        );
        keep_or(
            &mut attributes,
            ATTR_ACTION_SCALE,
            typed.action_scale.map(FieldValue::Double),
        );
        keep_or(
            &mut attributes,
            ATTR_JIGGLE,
            typed.jiggle.map(|v| FieldValue::Varint(int_varint(v))),
        );
    }
    keep_or(
        &mut attributes,
        ATTR_CUSTOM_BOUNCE,
        native.bounce.map(|v| FieldValue::Varint(u64::from(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_MOTION_BLUR,
        native.motion_blur.map(|v| FieldValue::Varint(u64::from(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_INCLUDE_ENDPOINTS,
        native
            .include_endpoints
            .map(|v| FieldValue::Varint(u64::from(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_SHINE,
        native.shine.map(|v| FieldValue::Varint(u64::from(v))),
    );
    keep_or(
        &mut attributes,
        ATTR_SCALE_AMOUNT,
        native.scale_amount.map(FieldValue::Double),
    );
    keep_or(
        &mut attributes,
        ATTR_TRAVEL_DISTANCE,
        native.travel_distance.map(FieldValue::Double),
    );
    let original_keyboard = string(&animation, ANIM_EFFECT)?.as_deref() == Some("apple:keyboard");
    if original_keyboard || matches!(settings.effect, Effect::Keyboard(_)) {
        keep_or(
            &mut attributes,
            ATTR_CURSOR,
            typed.cursor.map(|v| FieldValue::Varint(u64::from(v))),
        );
    }
    animation.set_message(ANIM_TYPE, stage_text(&settings.stage).as_bytes());
    let effect = effect_identifier(&settings.effect, native);
    animation.set_message(ANIM_EFFECT, effect.as_bytes());
    animation.set_double(ANIM_DURATION, settings.duration.as_f64());
    keep_or(
        &mut animation,
        ANIM_DIRECTION,
        direction(&settings.effect, native).map(|v| FieldValue::Varint(u64::from(v))),
    );
    animation.set_double(ANIM_DELAY, settings.delay.as_f64());
    if action_acceleration(&settings.effect) == Some(Acceleration::Custom) {
        if let Some(curve) = &settings.timing_curve {
            animation.set_message(ANIM_TIMING_CURVE, &path_source(curve.path()).finish());
            animation.remove(ANIM_TIMING_CURVE_THEME);
        }
    } else {
        animation.remove(ANIM_TIMING_CURVE);
        animation.remove(ANIM_TIMING_CURVE_THEME);
    }
    attributes.set_message(ATTR_ANIMATION, &animation.encode());
    build.set_message(BUILD_ATTRIBUTES, &attributes.encode());
    Ok(build.encode())
}

/// Patch a timing chunk's start, delay and duration.
pub(super) fn patch_chunk(
    original: &[u8],
    settings: &Settings,
    wire: WireLimits,
) -> Result<Vec<u8>> {
    let (automatic, referent) = start_fields(settings.start);
    let mut chunk = Fields::parse(original, wire)?;
    chunk.set_double(CHUNK_DELAY, settings.delay.as_f64());
    chunk.set_double(CHUNK_DURATION, settings.duration.as_f64());
    chunk.set_varint(CHUNK_AUTOMATIC, u64::from(automatic));
    chunk.set_varint(CHUNK_REFERENT, u64::from(referent));
    Ok(chunk.encode())
}
