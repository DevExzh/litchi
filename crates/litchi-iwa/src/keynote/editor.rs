//! Semantic text editing for existing Keynote slides.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::{
    DrawableCommentInfo, DrawableCommentReplyInfo, IWorkDrawableCommentEditor, IWorkDrawableInfo,
};
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    add_component_object_uuids, component_uuid_identifiers, next_object_identifier,
    release_package_identifier_suffix, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::protobuf::{kn, tsd, tsp, tswp};
use crate::shapes::{
    DrawableGeometry, DrawableProperties, set_shape_geometry, set_shape_properties, shape_geometry,
    shape_properties,
};
use crate::text::{IWorkTextEditor, TextStorageInfo};
use crate::wire::{
    append_repeated_length_delimited_field, patch_fixed32_field, patch_fixed64_field,
    patch_length_delimited_field, patch_nested_fixed32_field, patch_nested_fixed64_field,
    patch_nested_length_delimited_field, patch_nested_varint_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const ROTATE_ACTION_EFFECT: &str = "apple:action-rotation";
const SCALE_ACTION_EFFECT: &str = "apple:action-scale";
const OPACITY_ACTION_EFFECT: &str = "apple:action-opacity";
const MOVE_ACTION_EFFECT: &str = "apple:action-motion-path";
const BLINK_ACTION_EFFECT: &str = "apple:action-blink";
const BOUNCE_ACTION_EFFECT: &str = "apple:action-bounce";
const FLIP_ACTION_EFFECT: &str = "apple:action-flip";
const JIGGLE_ACTION_EFFECT: &str = "apple:action-jiggle";
const POP_ACTION_EFFECT: &str = "apple:action-pop";
const PULSE_ACTION_EFFECT: &str = "apple:action-pulse";
const KEYBOARD_BUILD_EFFECT: &str = "apple:keyboard";
const SHIMMER_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.KLNShimmer";
const SKID_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.KNBuildSkidByCharacter";
const SWOOSH_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.BLTSwoosh";
const TRACE_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.Trace";
const TEXT_BOX_DUPLICATE_OFFSET: f32 = 10.0;

/// Writable text targets resolved from one slide in presentation order.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideInfo {
    pub index: usize,
    pub node_id: u64,
    pub slide_id: u64,
    pub name: Option<String>,
    pub is_skipped: bool,
    pub is_slide_number_visible: Option<bool>,
    pub transition: Option<KeynoteTransitionSettings>,
    pub title_storage_id: Option<u64>,
    pub title: Option<String>,
    pub body_storage_id: Option<u64>,
    pub body: Option<String>,
    pub notes_storage_id: Option<u64>,
    pub notes: Option<String>,
}

/// Semantic role of a writable text-bearing drawable owned by a slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSlideTextRole {
    Title,
    Body,
    TextBox,
}

/// A writable text storage owned by a drawable on one Keynote slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeynoteSlideTextInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub role: KeynoteSlideTextRole,
    pub storage: TextStorageInfo,
}

/// A Keynote text box removed from a slide with its final text state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedKeynoteTextBox {
    pub text: KeynoteSlideTextInfo,
}

#[derive(Debug, Clone)]
struct KeynoteTextBoxGraph {
    slide_id: u64,
    archive_name: String,
    drawable_id: u64,
    storage_id: u64,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

/// Modern transition fields embedded in a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteTransitionSettings {
    pub animation_type: Option<String>,
    pub effect: Option<String>,
    pub duration: Option<f64>,
    pub direction: Option<u32>,
    pub delay: Option<f64>,
    pub is_automatic: Option<bool>,
    /// Modern animation-level parameters, including byte-exact native color
    /// and timing-curve protobuf payloads.
    pub animation_parameters: KeynoteTransitionAnimationParameters,
    /// Native effect-specific parameters whose effect semantics are not yet
    /// promoted to typed transition variants.
    pub custom_parameters: KeynoteTransitionCustomParameters,
}

/// Lossless parameters stored inside a transition's modern animation archive.
///
/// Color and timing curves are kept as encoded native protobuf payloads. This
/// permits arbitrary current and future path-source variants to round-trip
/// without exposing private generated protobuf types in the public API.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct KeynoteTransitionAnimationParameters {
    pub color_payload: Option<Vec<u8>>,
    pub timing_curve_payloads: [Option<Vec<u8>>; 3],
    pub random_number_seed: Option<u32>,
    pub detail: Option<f64>,
    pub timing_curve_theme_names: [Option<String>; 3],
    pub writing_direction_is_rtl: Option<bool>,
}

/// Lossless native parameters shared by Keynote transition effects.
///
/// These values intentionally retain their protobuf-level representation so
/// callers can round-trip effects that do not yet have a typed semantic API.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct KeynoteTransitionCustomParameters {
    pub twist: Option<f32>,
    pub mosaic_size: Option<u32>,
    pub mosaic_type: Option<u32>,
    pub bounce: Option<bool>,
    pub magic_move_fade_unmatched_objects: Option<bool>,
    pub timing_curve: Option<i32>,
    pub text_delivery_type: Option<i32>,
    pub motion_blur: Option<bool>,
    pub travel_distance: Option<f32>,
}

impl KeynoteTransitionSettings {
    fn from_native(attributes: &kn::TransitionAttributesArchive) -> Option<Self> {
        let animation = attributes.animation_attributes.as_ref()?;
        Some(Self {
            animation_type: animation.animation_type.clone(),
            effect: animation.effect.clone(),
            duration: animation.duration,
            direction: animation.direction,
            delay: animation.delay,
            is_automatic: animation.is_automatic,
            animation_parameters: KeynoteTransitionAnimationParameters {
                color_payload: None,
                timing_curve_payloads: [None, None, None],
                random_number_seed: animation.random_number_seed,
                detail: animation.custom_detail,
                timing_curve_theme_names: [
                    animation.custom_effect_timing_curve_theme_name_1.clone(),
                    animation.custom_effect_timing_curve_theme_name_2.clone(),
                    animation.custom_effect_timing_curve_theme_name_3.clone(),
                ],
                writing_direction_is_rtl: animation.writing_direction_is_rtl,
            },
            custom_parameters: KeynoteTransitionCustomParameters {
                twist: attributes.custom_twist,
                mosaic_size: attributes.custom_mosaic_size,
                mosaic_type: attributes.custom_mosaic_type,
                bounce: attributes.custom_bounce,
                magic_move_fade_unmatched_objects: attributes
                    .custom_magic_move_fade_unmatched_objects,
                timing_curve: attributes.custom_timing_curve,
                text_delivery_type: attributes.custom_text_delivery_type,
                motion_blur: attributes.custom_motion_blur,
                travel_distance: attributes.custom_travel_distance,
            },
        })
    }
}

fn optional_length_delimited_payload(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
    let payloads = repeated_length_delimited_payloads(data, field_number)?;
    if payloads.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} occurs more than once"
        )));
    }
    Ok(payloads.into_iter().next())
}

fn required_length_delimited_payload<'a>(
    data: &'a [u8],
    field_number: u32,
    context: &str,
) -> Result<&'a [u8]> {
    optional_length_delimited_payload(data, field_number)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{context} is missing protobuf field {field_number}"
        ))
    })
}

fn transition_settings_from_wire(
    original: &[u8],
    attributes: &kn::TransitionAttributesArchive,
) -> Result<Option<KeynoteTransitionSettings>> {
    let Some(mut settings) = KeynoteTransitionSettings::from_native(attributes) else {
        return Ok(None);
    };
    let transition = required_length_delimited_payload(original, 4, "Keynote slide transition")?;
    let attributes_wire =
        required_length_delimited_payload(transition, 2, "Keynote transition attributes")?;
    let animation = required_length_delimited_payload(
        attributes_wire,
        8,
        "Keynote modern transition attributes",
    )?;
    settings.animation_parameters.color_payload =
        optional_length_delimited_payload(animation, 7)?.map(Vec::from);
    for (index, field_number) in [8, 9, 10].into_iter().enumerate() {
        settings.animation_parameters.timing_curve_payloads[index] =
            optional_length_delimited_payload(animation, field_number)?.map(Vec::from);
    }
    Ok(Some(settings))
}

/// Native start relationship for one Keynote build event.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteBuildStart {
    /// Advance to this build with a presenter click.
    OnClick,
    /// Start automatically after the slide transition.
    AfterTransition,
    /// Start concurrently with the preceding build event.
    WithPrevious,
    /// Start after the preceding build event completes.
    AfterPrevious,
}

/// Speed curve for a Keynote action build.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteBuildAcceleration {
    None,
    EaseIn,
    EaseOut,
    EaseInOut,
    Custom,
}

/// Rotation direction for a Keynote Rotate action build.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteRotationDirection {
    Clockwise,
    Counterclockwise,
}

/// Typed parameters for Keynote's object Rotate action.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteRotationAction {
    /// Total rotation in degrees. For example, two full turns plus 90° is 810°.
    pub total_degrees: f64,
    pub direction: KeynoteRotationDirection,
    pub acceleration: KeynoteBuildAcceleration,
}

/// Typed parameters for Keynote's object Scale action.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteScaleAction {
    /// Final size as a factor of the object's original size (`1.5` is 150%).
    pub scale_factor: f64,
    pub acceleration: KeynoteBuildAcceleration,
}

/// Typed parameters for Keynote's object Opacity action.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteOpacityAction {
    /// Final opacity in Keynote percentage points (`37.0` is 37%).
    pub opacity_percent: f64,
    pub acceleration: KeynoteBuildAcceleration,
}

/// Node kind used by Keynote's editable Bézier motion paths.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteMotionPathNodeType {
    Sharp,
    Bezier,
    Smooth,
}

/// A point in a Keynote motion path, measured in slide points from its origin.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteMotionPathPoint {
    pub x: f32,
    pub y: f32,
}

impl KeynoteMotionPathPoint {
    pub const fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }
}

/// One editable node and its absolute-in-path Bézier control points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteMotionPathNode {
    pub in_control_point: KeynoteMotionPathPoint,
    pub point: KeynoteMotionPathPoint,
    pub out_control_point: KeynoteMotionPathPoint,
    pub node_type: KeynoteMotionPathNodeType,
}

impl KeynoteMotionPathNode {
    /// A sharp node whose control points coincide with the node.
    pub const fn sharp(x: f32, y: f32) -> Self {
        let point = KeynoteMotionPathPoint::new(x, y);
        Self {
            in_control_point: point,
            point,
            out_control_point: point,
            node_type: KeynoteMotionPathNodeType::Sharp,
        }
    }
}

/// One continuous or closed subpath in a Keynote motion path.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteMotionSubpath {
    pub nodes: Vec<KeynoteMotionPathNode>,
    pub closed: bool,
}

/// Editable motion geometry stored inline in a Keynote Move action.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteMotionPath {
    pub subpaths: Vec<KeynoteMotionSubpath>,
    pub natural_width: f32,
    pub natural_height: f32,
    pub horizontal_flip: bool,
    pub vertical_flip: bool,
}

impl KeynoteMotionPath {
    /// A native-compatible straight path from the object's current position.
    pub fn straight(delta_x: f32, delta_y: f32) -> Self {
        Self {
            subpaths: vec![KeynoteMotionSubpath {
                nodes: vec![
                    KeynoteMotionPathNode::sharp(0.0, 0.0),
                    KeynoteMotionPathNode::sharp(delta_x, delta_y),
                ],
                closed: false,
            }],
            natural_width: delta_x.abs(),
            natural_height: delta_y.abs(),
            horizontal_flip: false,
            vertical_flip: false,
        }
    }

    /// Recompute conservative natural bounds from every node and control point.
    pub fn recalculate_natural_size(&mut self) {
        let mut points = self
            .subpaths
            .iter()
            .flat_map(|subpath| &subpath.nodes)
            .flat_map(|node| [node.in_control_point, node.point, node.out_control_point]);
        let Some(first) = points.next() else {
            self.natural_width = 0.0;
            self.natural_height = 0.0;
            return;
        };
        let (mut min_x, mut max_x, mut min_y, mut max_y) = (first.x, first.x, first.y, first.y);
        for point in points {
            min_x = min_x.min(point.x);
            max_x = max_x.max(point.x);
            min_y = min_y.min(point.y);
            max_y = max_y.max(point.y);
        }
        self.natural_width = max_x - min_x;
        self.natural_height = max_y - min_y;
    }
}

/// Typed parameters for Keynote's object Move action.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteMoveAction {
    pub path: KeynoteMotionPath,
    pub align_to_path: bool,
    pub acceleration: KeynoteBuildAcceleration,
}

/// Horizontal direction used by Keynote's Flip emphasis action.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteFlipDirection {
    LeftToRight,
    RightToLeft,
}

/// Intensity used by Keynote's Jiggle emphasis action.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteJiggleIntensity {
    Small,
    Medium,
    Large,
}

/// Typed parameters for Keynote's object-emphasis actions.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum KeynoteEmphasisAction {
    Blink {
        repeat_count: u32,
        fade: bool,
    },
    Bounce {
        repeat_count: u32,
        decay: bool,
    },
    Flip {
        repeat_count: u32,
        direction: KeynoteFlipDirection,
    },
    Jiggle {
        intensity: KeynoteJiggleIntensity,
    },
    Pop {
        scale_percent: f64,
    },
    Pulse {
        repeat_count: u32,
        scale_percent: f64,
    },
}

/// Text traversal direction used by Keynote's Keyboard build effect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteKeyboardDirection {
    Forward,
    Backward,
}

/// Typed parameters for Keynote's Keyboard build-in/build-out effect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct KeynoteKeyboardBuild {
    pub direction: KeynoteKeyboardDirection,
    pub show_cursor: bool,
}

/// Horizontal traversal used by Keynote's Skid and Trace builds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteHorizontalBuildDirection {
    LeftToRight,
    RightToLeft,
}

/// Origin used by Keynote's Swoosh build.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSwooshDirection {
    Center,
    FromLeft,
    FromRight,
}

/// Typed native Build In / Build Out effects with no opaque parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteObjectBuildEffect {
    Shimmer,
    Skid {
        direction: KeynoteHorizontalBuildDirection,
    },
    Swoosh {
        direction: KeynoteSwooshDirection,
    },
    Trace {
        direction: KeynoteHorizontalBuildDirection,
    },
}

/// Raw native parameters shared by Keynote build effects that do not yet have
/// a dedicated semantic model.
///
/// The fields retain their protobuf meaning and presence. This permits
/// lossless CRUD for newer or effect-specific builds while their user-facing
/// semantics are still being mapped.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct KeynoteBuildCustomParameters {
    pub bounce: Option<bool>,
    pub motion_blur: Option<bool>,
    pub include_endpoints: Option<bool>,
    pub shine: Option<bool>,
    pub scale_amount: Option<f64>,
    pub travel_distance: Option<f64>,
}

impl KeynoteBuildCustomParameters {
    fn is_empty(self) -> bool {
        self == Self::default()
    }
}

/// Editable fields for an all-at-once Keynote object build.
///
/// Keynote effect identifiers are database strings such as
/// `apple:bc-appear` or `apple:dissolve character`. Keeping them as strings
/// preserves effects introduced by newer Keynote releases.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteBuildSettings {
    pub delivery: String,
    pub animation_type: String,
    pub effect: String,
    pub duration: f64,
    /// Delay before an `AfterTransition` or `AfterPrevious` build starts.
    pub delay: f64,
    pub start: KeynoteBuildStart,
    pub direction: Option<u32>,
    /// Raw `BuildAttributesTextDelivery` value for forward compatibility.
    pub text_delivery: Option<i32>,
    /// Raw `BuildAttributesDeliveryOption` value for forward compatibility.
    pub delivery_option: Option<i32>,
    pub event_trigger: Option<u32>,
    /// Present for Keynote's native `apple:action-rotation` action effect.
    pub rotation: Option<KeynoteRotationAction>,
    /// Present for Keynote's native `apple:action-scale` action effect.
    pub scale: Option<KeynoteScaleAction>,
    /// Present for Keynote's native `apple:action-opacity` action effect.
    pub opacity: Option<KeynoteOpacityAction>,
    /// Present for Keynote's native `apple:action-motion-path` action effect.
    pub move_action: Option<KeynoteMoveAction>,
    /// Present for Keynote's Blink/Bounce/Flip/Jiggle/Pop/Pulse actions.
    pub emphasis: Option<KeynoteEmphasisAction>,
    /// Present for Keynote's native `apple:keyboard` build-in/build-out effect.
    pub keyboard: Option<KeynoteKeyboardBuild>,
    /// Present for typed Shimmer, Skid, Swoosh, and Trace builds.
    pub object_effect: Option<KeynoteObjectBuildEffect>,
    /// Raw parameters for native effects without a dedicated typed model.
    pub custom_parameters: KeynoteBuildCustomParameters,
}

impl KeynoteBuildSettings {
    /// Native-compatible object-level Appear build-in settings.
    pub fn appear_in() -> Self {
        Self {
            delivery: "All at Once".to_owned(),
            animation_type: "In".to_owned(),
            effect: "apple:bc-appear".to_owned(),
            duration: 1.0,
            delay: 0.0,
            start: KeynoteBuildStart::OnClick,
            direction: None,
            text_delivery: Some(
                kn::build_attributes_archive::BuildAttributesTextDelivery::KTextDeliveryByObject
                    as i32,
            ),
            delivery_option: Some(
                kn::build_attributes_archive::BuildAttributesDeliveryOption::KDeliveryOptionForward
                    as i32,
            ),
            event_trigger: Some(1),
            rotation: None,
            scale: None,
            opacity: None,
            move_action: None,
            emphasis: None,
            keyboard: None,
            object_effect: None,
            custom_parameters: KeynoteBuildCustomParameters::default(),
        }
    }

    /// Native-compatible object-level Appear build-out settings.
    pub fn appear_out() -> Self {
        Self {
            animation_type: "Out".to_owned(),
            ..Self::appear_in()
        }
    }

    fn object_build(animation_type: &str, effect: KeynoteObjectBuildEffect, duration: f64) -> Self {
        Self {
            animation_type: animation_type.to_owned(),
            effect: object_build_effect_identifier(effect).to_owned(),
            duration,
            direction: native_object_build_direction(effect),
            text_delivery: None,
            delivery_option: None,
            object_effect: Some(effect),
            ..Self::appear_in()
        }
    }

    /// Native-compatible Shimmer Build In.
    pub fn shimmer_in() -> Self {
        Self::object_build("In", KeynoteObjectBuildEffect::Shimmer, 1.5)
    }

    /// Native-compatible Shimmer Build Out.
    pub fn shimmer_out() -> Self {
        Self::object_build("Out", KeynoteObjectBuildEffect::Shimmer, 1.5)
    }

    /// Native-compatible Skid Build In.
    pub fn skid_in(direction: KeynoteHorizontalBuildDirection) -> Self {
        Self::object_build("In", KeynoteObjectBuildEffect::Skid { direction }, 1.25)
    }

    /// Native-compatible Skid Build Out.
    pub fn skid_out(direction: KeynoteHorizontalBuildDirection) -> Self {
        Self::object_build("Out", KeynoteObjectBuildEffect::Skid { direction }, 1.25)
    }

    /// Native-compatible Swoosh Build In.
    pub fn swoosh_in(direction: KeynoteSwooshDirection) -> Self {
        Self::object_build("In", KeynoteObjectBuildEffect::Swoosh { direction }, 1.0)
    }

    /// Native-compatible Swoosh Build Out.
    pub fn swoosh_out(direction: KeynoteSwooshDirection) -> Self {
        Self::object_build("Out", KeynoteObjectBuildEffect::Swoosh { direction }, 1.0)
    }

    /// Native-compatible Trace Build In.
    pub fn trace_in(direction: KeynoteHorizontalBuildDirection) -> Self {
        Self::object_build("In", KeynoteObjectBuildEffect::Trace { direction }, 2.0)
    }

    /// Native-compatible Trace Build Out.
    pub fn trace_out(direction: KeynoteHorizontalBuildDirection) -> Self {
        Self::object_build("Out", KeynoteObjectBuildEffect::Trace { direction }, 2.0)
    }

    /// Native-compatible Rotate action with Keynote's default ease-in/out curve.
    pub fn rotate_action(total_degrees: f64, direction: KeynoteRotationDirection) -> Self {
        Self {
            animation_type: "Action".to_owned(),
            effect: ROTATE_ACTION_EFFECT.to_owned(),
            text_delivery: None,
            delivery_option: None,
            rotation: Some(KeynoteRotationAction {
                total_degrees,
                direction,
                acceleration: KeynoteBuildAcceleration::EaseInOut,
            }),
            scale: None,
            opacity: None,
            move_action: None,
            ..Self::appear_in()
        }
    }

    /// Native-compatible Scale action with Keynote's default ease-in/out curve.
    pub fn scale_action(scale_factor: f64) -> Self {
        Self {
            animation_type: "Action".to_owned(),
            effect: SCALE_ACTION_EFFECT.to_owned(),
            text_delivery: None,
            delivery_option: None,
            rotation: None,
            scale: Some(KeynoteScaleAction {
                scale_factor,
                acceleration: KeynoteBuildAcceleration::EaseInOut,
            }),
            opacity: None,
            move_action: None,
            ..Self::appear_in()
        }
    }

    /// Native-compatible Opacity action with Keynote's default ease-in/out curve.
    pub fn opacity_action(opacity_percent: f64) -> Self {
        Self {
            animation_type: "Action".to_owned(),
            effect: OPACITY_ACTION_EFFECT.to_owned(),
            text_delivery: None,
            delivery_option: None,
            rotation: None,
            scale: None,
            opacity: Some(KeynoteOpacityAction {
                opacity_percent,
                acceleration: KeynoteBuildAcceleration::EaseInOut,
            }),
            move_action: None,
            ..Self::appear_in()
        }
    }

    /// Native-compatible straight Move action with Keynote's default curve.
    pub fn move_action(delta_x: f32, delta_y: f32) -> Self {
        Self::move_along_path(KeynoteMotionPath::straight(delta_x, delta_y))
    }

    /// Native-compatible Move action along an arbitrary editable Bézier path.
    pub fn move_along_path(path: KeynoteMotionPath) -> Self {
        Self {
            animation_type: "Action".to_owned(),
            effect: MOVE_ACTION_EFFECT.to_owned(),
            text_delivery: None,
            delivery_option: None,
            rotation: None,
            scale: None,
            opacity: None,
            move_action: Some(KeynoteMoveAction {
                path,
                align_to_path: false,
                acceleration: KeynoteBuildAcceleration::EaseInOut,
            }),
            ..Self::appear_in()
        }
    }

    /// Native-compatible Blink emphasis action.
    pub fn blink_action(repeat_count: u32, fade: bool) -> Self {
        Self::emphasis_action(
            BLINK_ACTION_EFFECT,
            KeynoteEmphasisAction::Blink { repeat_count, fade },
        )
    }

    /// Native-compatible Bounce emphasis action.
    pub fn bounce_action(repeat_count: u32, decay: bool) -> Self {
        Self::emphasis_action(
            BOUNCE_ACTION_EFFECT,
            KeynoteEmphasisAction::Bounce {
                repeat_count,
                decay,
            },
        )
    }

    /// Native-compatible Flip emphasis action.
    pub fn flip_action(repeat_count: u32, direction: KeynoteFlipDirection) -> Self {
        Self::emphasis_action(
            FLIP_ACTION_EFFECT,
            KeynoteEmphasisAction::Flip {
                repeat_count,
                direction,
            },
        )
    }

    /// Native-compatible Jiggle emphasis action.
    pub fn jiggle_action(intensity: KeynoteJiggleIntensity) -> Self {
        Self::emphasis_action(
            JIGGLE_ACTION_EFFECT,
            KeynoteEmphasisAction::Jiggle { intensity },
        )
    }

    /// Native-compatible Pop emphasis action.
    pub fn pop_action(scale_percent: f64) -> Self {
        let mut settings = Self::emphasis_action(
            POP_ACTION_EFFECT,
            KeynoteEmphasisAction::Pop { scale_percent },
        );
        settings.duration = 0.5;
        settings
    }

    /// Native-compatible Pulse emphasis action.
    pub fn pulse_action(repeat_count: u32, scale_percent: f64) -> Self {
        Self::emphasis_action(
            PULSE_ACTION_EFFECT,
            KeynoteEmphasisAction::Pulse {
                repeat_count,
                scale_percent,
            },
        )
    }

    /// Native-compatible Keyboard build-in settings.
    pub fn keyboard_in(direction: KeynoteKeyboardDirection, show_cursor: bool) -> Self {
        Self::keyboard_build("In", direction, show_cursor)
    }

    /// Native-compatible Keyboard build-out settings.
    pub fn keyboard_out(direction: KeynoteKeyboardDirection, show_cursor: bool) -> Self {
        Self::keyboard_build("Out", direction, show_cursor)
    }

    fn keyboard_build(
        animation_type: &str,
        direction: KeynoteKeyboardDirection,
        show_cursor: bool,
    ) -> Self {
        Self {
            animation_type: animation_type.to_owned(),
            effect: KEYBOARD_BUILD_EFFECT.to_owned(),
            duration: 3.0,
            direction: Some(native_keyboard_direction(direction)),
            text_delivery: None,
            delivery_option: None,
            keyboard: Some(KeynoteKeyboardBuild {
                direction,
                show_cursor,
            }),
            ..Self::appear_in()
        }
    }

    fn emphasis_action(effect: &str, emphasis: KeynoteEmphasisAction) -> Self {
        let direction = match emphasis {
            KeynoteEmphasisAction::Flip { direction, .. } => Some(native_flip_direction(direction)),
            _ => None,
        };
        Self {
            animation_type: "Action".to_owned(),
            effect: effect.to_owned(),
            direction,
            text_delivery: None,
            delivery_option: None,
            rotation: None,
            scale: None,
            opacity: None,
            move_action: None,
            emphasis: Some(emphasis),
            ..Self::appear_in()
        }
    }
}

/// One timing chunk owned by a Keynote build.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteBuildChunkInfo {
    pub object_id: u64,
    pub delay: Option<f64>,
    pub duration: Option<f64>,
    pub automatic: Option<bool>,
    pub referent: Option<bool>,
    pub chunk_id: Option<i32>,
}

/// A drawable build and its timing chunks in slide build order.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteBuildInfo {
    pub slide_index: usize,
    pub object_id: u64,
    pub drawable_object_id: u64,
    pub settings: KeynoteBuildSettings,
    pub chunks: Vec<KeynoteBuildChunkInfo>,
}

/// Writable presentation-level behavior stored in `KN.ShowArchive`.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteShowSettings {
    pub width: f32,
    pub height: f32,
    pub slide_numbers_visible: Option<bool>,
    pub loop_presentation: Option<bool>,
    /// Raw `KNShowMode` value for forward compatibility.
    pub mode: Option<i32>,
    pub autoplay_transition_delay: Option<f64>,
    pub autoplay_build_delay: Option<f64>,
    pub idle_timer_active: Option<bool>,
    pub idle_timer_delay: Option<f64>,
    pub automatically_plays_upon_open: Option<bool>,
}

impl From<&kn::ShowArchive> for KeynoteShowSettings {
    fn from(show: &kn::ShowArchive) -> Self {
        Self {
            width: show.size.width,
            height: show.size.height,
            slide_numbers_visible: show.slide_numbers_visible,
            loop_presentation: show.loop_presentation,
            mode: show.mode,
            autoplay_transition_delay: show.autoplay_transition_delay,
            autoplay_build_delay: show.autoplay_build_delay,
            idle_timer_active: show.idle_timer_active,
            idle_timer_delay: show.idle_timer_delay,
            automatically_plays_upon_open: show.automatically_plays_upon_open,
        }
    }
}

/// Transactional editor for title and body placeholders in an existing Keynote package.
#[derive(Debug, Clone)]
pub struct KeynoteEditor {
    text: IWorkTextEditor,
}

impl KeynoteEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        let editor = Self {
            text: IWorkTextEditor::from_package(package),
        };
        editor.slides()?;
        Ok(editor)
    }

    pub fn slides(&self) -> Result<Vec<KeynoteSlideInfo>> {
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;

        let mut slides = Vec::with_capacity(show.slide_tree.slides.len());
        for (index, node_reference) in show.slide_tree.slides.into_iter().enumerate() {
            let node: kn::SlideNodeArchive =
                graph.decode(node_reference.identifier, "KN.SlideNodeArchive")?;
            let slide_reference = node.slide.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide node {} has no slide reference",
                    node_reference.identifier
                ))
            })?;
            let slide: kn::SlideArchive =
                graph.decode(slide_reference.identifier, "KN.SlideArchive")?;
            let title_storage_id = slide
                .title_placeholder
                .map(|reference| graph.drawable_storage(reference.identifier))
                .transpose()?
                .flatten();
            let body_storage_id = slide
                .body_placeholder
                .map(|reference| graph.drawable_storage(reference.identifier))
                .transpose()?
                .flatten();
            let title = title_storage_id
                .map(|identifier| graph.storage_text(identifier))
                .transpose()?;
            let body = body_storage_id
                .map(|identifier| graph.storage_text(identifier))
                .transpose()?;
            let notes_storage_id = slide
                .note
                .map(|reference| {
                    graph
                        .decode::<kn::NoteArchive>(reference.identifier, "KN.NoteArchive")
                        .map(|note| note.contained_storage.identifier)
                })
                .transpose()?;
            let notes = notes_storage_id
                .map(|identifier| graph.storage_text(identifier))
                .transpose()?;
            let transition = if slide.transition.attributes.animation_attributes.is_some() {
                let original =
                    graph.message_data_type(slide_reference.identifier, 5, "KN.SlideArchive")?;
                let transition =
                    transition_settings_from_wire(original, &slide.transition.attributes)?;
                validate_transition_wire(original, &slide.transition.attributes)?;
                transition
            } else {
                None
            };
            slides.push(KeynoteSlideInfo {
                index,
                node_id: node_reference.identifier,
                slide_id: slide_reference.identifier,
                name: slide.name.filter(|name| !name.is_empty()),
                is_skipped: node.is_skipped,
                is_slide_number_visible: node.is_slide_number_visible,
                transition,
                title_storage_id,
                title,
                body_storage_id,
                body,
                notes_storage_id,
                notes,
            });
        }
        Ok(slides)
    }

    /// List every writable text storage owned by a drawable on one slide.
    ///
    /// This includes title and body placeholders plus ordinary text boxes in
    /// slide ownership order. Speaker notes are exposed separately on
    /// [`KeynoteSlideInfo`] because they are owned by `KN.NoteArchive` rather
    /// than a drawable.
    pub fn slide_text_storages(&self, slide_index: usize) -> Result<Vec<KeynoteSlideTextInfo>> {
        let slides = self.slides()?;
        slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let mut storage_owners = HashMap::<u64, (usize, u64)>::new();
        let mut result = Vec::new();
        for (owner_slide_index, owner) in slides.iter().enumerate() {
            let slide: kn::SlideArchive = graph.decode(owner.slide_id, "KN.SlideArchive")?;
            let title = slide
                .title_placeholder
                .as_ref()
                .map(|reference| reference.identifier);
            let body = slide
                .body_placeholder
                .as_ref()
                .map(|reference| reference.identifier);
            let mut seen_drawables = HashSet::with_capacity(slide.owned_drawables.len());
            for reference in &slide.owned_drawables {
                let drawable_id = reference.identifier;
                if !seen_drawables.insert(drawable_id) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {owner_slide_index} repeats owned drawable {drawable_id}"
                    )));
                }
                let Some(storage_id) = graph.drawable_storage(drawable_id)? else {
                    continue;
                };
                if let Some((previous_slide, previous_drawable)) =
                    storage_owners.insert(storage_id, (owner_slide_index, drawable_id))
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {previous_slide} drawable {previous_drawable} and slide {owner_slide_index} drawable {drawable_id} share owned text storage {storage_id}"
                    )));
                }
                if owner_slide_index != slide_index {
                    continue;
                }
                let storage = self.text.storage(storage_id).map_err(|error| {
                    Error::InvalidFormat(format!(
                        "Keynote slide {slide_index} drawable {drawable_id} owns invalid text storage {storage_id}: {error}"
                    ))
                })?;
                let role = if title == Some(drawable_id) {
                    KeynoteSlideTextRole::Title
                } else if body == Some(drawable_id) {
                    KeynoteSlideTextRole::Body
                } else {
                    KeynoteSlideTextRole::TextBox
                };
                result.push(KeynoteSlideTextInfo {
                    slide_index,
                    drawable_object_id: drawable_id,
                    role,
                    storage,
                });
            }
        }
        Ok(result)
    }

    /// Replace a UTF-16 range in a slide-owned text box or placeholder.
    pub fn replace_slide_text_storage(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.require_slide_text_storage(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.replace_text(storage_id, range, replacement)?;
        Self::from_package(staged.package().clone())?;
        self.text = staged;
        Ok(())
    }

    /// Replace all text in a slide-owned text box or placeholder.
    pub fn set_slide_text_storage(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.require_slide_text_storage(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text(storage_id, replacement)?;
        Self::from_package(staged.package().clone())?;
        self.text = staged;
        Ok(())
    }

    /// Clear a slide-owned text box or placeholder without deleting its shape.
    pub fn clear_slide_text_storage(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.set_slide_text_storage(slide_index, drawable_object_id, "")
    }

    /// Read the geometry of an ordinary text box owned by one slide.
    pub fn slide_text_box_geometry(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        shape_geometry(self.package(), &graph.archive_name, drawable_object_id)
    }

    /// Update position, size, flags, and rotation on a slide-owned ordinary text box.
    pub fn set_slide_text_box_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_text_box_geometry(slide_index, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Keynote text-box geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties from an ordinary slide-owned text box.
    pub fn slide_text_box_properties(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        shape_properties(self.package(), &graph.archive_name, drawable_object_id)
    }

    /// Update shared drawable properties on an ordinary slide-owned text box.
    ///
    /// Keynote may disable aspect-ratio constraints for auto-sized text boxes
    /// even though the corresponding property is present in the archive.
    pub fn set_slide_text_box_properties(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_properties(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_text_box_properties(slide_index, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Keynote text-box properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate an ordinary slide-owned text box with independent storage.
    ///
    /// The shape, stand-in title and caption, and writable storage are cloned
    /// into the same slide component with fresh object identifiers and UUIDs.
    /// The clone is appended to both drawable ownership lists and offset by ten
    /// points so it remains independently selectable in Keynote.
    pub fn duplicate_slide_text_box(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
        text: &str,
    ) -> Result<KeynoteSlideTextInfo> {
        let source = self.text_box_graph(slide_index, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Keynote text-box graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Keynote text-box object {identifier} is missing"))
                })?;
                clone_slide_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = remap[&source.drawable_id];
        let new_storage_id = remap[&source.storage_id];
        offset_keynote_text_box(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            TEXT_BOX_DUPLICATE_OFFSET,
        )?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.set_text(new_storage_id, text)?;
        staged = text_editor.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Keynote text-box graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(&mut staged, source.slide_id, &new_uuid_object_ids)?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .slide_text_storages(slide_index)?
            .into_iter()
            .find(|item| item.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote text-box duplication failed validation".to_owned())
            })?;
        let created_graph = verified.text_box_graph(slide_index, new_drawable_id)?;
        if created.role != KeynoteSlideTextRole::TextBox
            || created.storage.object_id != new_storage_id
            || created.storage.text != text
            || created_graph.object_ids.len() != source.object_ids.len()
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Remove an ordinary slide-owned text box and its private object graph.
    pub fn remove_slide_text_box(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteTextBox> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let text = self
            .slide_text_storages(slide_index)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote text box {drawable_object_id} lost its writable storage"
                ))
            })?;

        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &graph.archive_name,
            graph.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        for identifier in &graph.object_ids {
            remove_object(&mut staged, &graph.archive_name, *identifier)?;
        }
        for identifier in &graph.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote text-box object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, graph.slide_id, &graph.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_package(staged)?;
        if verified
            .slide_text_storages(slide_index)?
            .iter()
            .any(|item| item.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteTextBox { text })
    }

    /// List supported direct-comment drawables owned by one slide.
    pub fn slide_drawables(&self, slide_index: usize) -> Result<Vec<IWorkDrawableInfo>> {
        let owned = self.slide_owned_drawable_ids(slide_index)?;
        let mut drawables = IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .drawables()?
            .into_iter()
            .filter(|drawable| owned.contains(&drawable.object_id))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.object_id);
        Ok(drawables)
    }

    /// Read a comment attached directly to a drawable owned by one slide.
    pub fn slide_drawable_comment(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<DrawableCommentInfo>> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .comment(drawable_object_id)
    }

    /// Create or replace a direct drawable comment on one slide.
    pub fn set_slide_drawable_comment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.set_comment(drawable_object_id, text)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Delete a direct drawable comment on one slide.
    pub fn clear_slide_drawable_comment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Read direct replies in a comment thread on one slide drawable.
    pub fn slide_drawable_comment_replies(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<DrawableCommentReplyInfo>> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .replies(drawable_object_id)
    }

    /// Add a reply to a direct comment on one slide drawable.
    pub fn add_slide_drawable_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.add_reply(drawable_object_id, text)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(reply_id)
    }

    /// Update a direct reply, returning its current storage identifier.
    pub fn set_slide_drawable_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.set_reply(drawable_object_id, reply_storage_object_id, text)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(reply_id)
    }

    /// Remove a direct reply from a comment on one slide drawable.
    pub fn remove_slide_drawable_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.remove_reply(drawable_object_id, reply_storage_object_id)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    pub fn show_settings(&self) -> Result<KeynoteShowSettings> {
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
        Ok(KeynoteShowSettings::from(&show))
    }

    pub fn set_show_settings(&mut self, settings: KeynoteShowSettings) -> Result<()> {
        validate_show_settings(&settings)?;
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let archive_name = graph.archive_name(show_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(show_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == 2 && kn::ShowArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show object {show_id} has no ShowArchive payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let show = kn::ShowArchive::decode(original)?;
            let data = patch_show_settings_wire(original, &show, &settings)?;
            let verified = kn::ShowArchive::decode(data.as_slice())?;
            if KeynoteShowSettings::from(&verified) != settings {
                return Err(Error::InvalidFormat(
                    "Keynote show-settings wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.show_settings()? != settings {
            return Err(Error::InvalidFormat(
                "Keynote show settings failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    pub fn set_slide_skipped(&mut self, slide_index: usize, skipped: bool) -> Result<()> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let node_id = slide.node_id;
        let graph = ObjectGraph::read(self.text.package())?;
        let archive_name = graph.archive_name(node_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(node_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| kn::SlideNodeArchive::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote slide node {node_id} has no SlideNodeArchive payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let data = patch_varint_field(original, 4, true, Some(u64::from(skipped)))?;
            let node = kn::SlideNodeArchive::decode(data.as_slice())?;
            if node.is_skipped != skipped {
                return Err(Error::InvalidFormat(
                    "Keynote slide skip-state wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slides()?
            .get(slide_index)
            .map(|slide| slide.is_skipped)
            != Some(skipped)
        {
            return Err(Error::InvalidFormat(
                "Keynote slide skip-state update failed validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// List semantic object builds and their timing chunks for one slide.
    #[allow(deprecated)]
    pub fn slide_builds(&self, slide_index: usize) -> Result<Vec<KeynoteBuildInfo>> {
        let slides = self.slides()?;
        let slide_info = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let slide: kn::SlideArchive = graph.decode(slide_info.slide_id, "KN.SlideArchive")?;
        let archive_name = graph.archive_name(slide_info.slide_id)?;
        let owned_drawables = slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<HashSet<_>>();
        if owned_drawables.len() != slide.owned_drawables.len() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_index} repeats an owned drawable"
            )));
        }

        let mut build_ids = HashSet::with_capacity(slide.builds.len());
        for reference in &slide.builds {
            if !build_ids.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} repeats build {}",
                    reference.identifier
                )));
            }
            if graph.archive_name(reference.identifier)? != archive_name {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build {} is outside slide component {archive_name}",
                    reference.identifier
                )));
            }
        }

        let mut chunks_by_build = HashMap::<u64, Vec<KeynoteBuildChunkInfo>>::new();
        let mut chunk_ids = HashSet::with_capacity(slide.build_chunks.len());
        for reference in &slide.build_chunks {
            let chunk_id = reference.identifier;
            if !chunk_ids.insert(chunk_id) {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} repeats build chunk {chunk_id}"
                )));
            }
            if graph.archive_name(chunk_id)? != archive_name {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build chunk {chunk_id} is outside slide component {archive_name}"
                )));
            }
            let chunk: kn::BuildChunkArchive =
                graph.decode_type(chunk_id, BUILD_CHUNK_MESSAGE_TYPE, "KN.BuildChunkArchive")?;
            let build_id = chunk
                .build
                .as_ref()
                .map(|reference| reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote build chunk {chunk_id} has no build reference"
                    ))
                })?;
            if !build_ids.contains(&build_id) {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build chunk {chunk_id} refers to unowned build {build_id}"
                )));
            }
            let identifier = chunk.build_chunk_identifier.as_ref();
            if let (Some(build_uuid), Some(identifier_uuid)) = (
                chunk.build_id.as_ref(),
                identifier.and_then(|value| value.build_id.as_ref()),
            ) && build_uuid != identifier_uuid
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build chunk {chunk_id} has inconsistent build UUIDs"
                )));
            }
            chunks_by_build
                .entry(build_id)
                .or_default()
                .push(KeynoteBuildChunkInfo {
                    object_id: chunk_id,
                    delay: chunk.delay,
                    duration: chunk.duration,
                    automatic: chunk.automatic,
                    referent: chunk.referent,
                    chunk_id: identifier.and_then(|value| value.build_chunk_id),
                });
        }

        let first_chunk_id = slide
            .build_chunks
            .first()
            .map(|reference| reference.identifier);
        let mut builds = Vec::with_capacity(slide.builds.len());
        for reference in &slide.builds {
            let build_id = reference.identifier;
            let build: kn::BuildArchive =
                graph.decode_type(build_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")?;
            let drawable_object_id = build
                .drawable
                .as_ref()
                .map(|reference| reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote build {build_id} has no drawable reference"
                    ))
                })?;
            if !owned_drawables.contains(&drawable_object_id) {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build {build_id} targets drawable {drawable_object_id} outside slide {slide_index}"
                )));
            }
            let chunks = chunks_by_build.remove(&build_id).unwrap_or_default();
            let starts_slide_events = chunks
                .first()
                .is_some_and(|chunk| Some(chunk.object_id) == first_chunk_id);
            let settings = build_settings(&build, &chunks, starts_slide_events);
            let wire = graph.message_data_type(build_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")?;
            if validate_custom_build_parameters_wire(wire, &build).is_err() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build {build_id} has invalid custom-parameter wire data"
                )));
            }
            let typed_wire_is_valid = match settings.effect.as_str() {
                effect if is_typed_action_effect(effect) => {
                    validate_typed_action_wire(wire, &build).is_ok()
                },
                KEYBOARD_BUILD_EFFECT => validate_keyboard_build_wire(wire, &build).is_ok(),
                effect if is_object_build_effect(effect) => {
                    validate_object_build_wire(wire, &build).is_ok()
                },
                _ => true,
            };
            if is_typed_build_effect(&settings.effect)
                && (validate_build_settings(&settings).is_err() || !typed_wire_is_valid)
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build {build_id} has invalid typed parameters"
                )));
            }
            builds.push(KeynoteBuildInfo {
                slide_index,
                object_id: build_id,
                drawable_object_id,
                settings,
                chunks,
            });
        }
        if !chunks_by_build.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_index} contains orphaned build chunks"
            )));
        }
        let chunk_positions = slide
            .build_chunks
            .iter()
            .enumerate()
            .map(|(index, reference)| (reference.identifier, index))
            .collect::<HashMap<_, _>>();
        for build in &builds {
            let Some(first_chunk) = build.chunks.first() else {
                continue;
            };
            let event_index = chunk_positions[&first_chunk.object_id];
            if validate_build_start_position(build.settings.start, event_index).is_err() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build {} has {:?} at invalid event index {event_index}",
                    build.object_id, build.settings.start
                )));
            }
        }
        Ok(builds)
    }

    /// Add a native-compatible, all-at-once object build and timing chunk.
    pub fn add_slide_build(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        settings: KeynoteBuildSettings,
    ) -> Result<KeynoteBuildInfo> {
        validate_build_settings(&settings)?;
        if typed_action_acceleration(&settings) == Some(KeynoteBuildAcceleration::Custom) {
            return Err(Error::ParseError(
                "Creating a custom-curve Keynote action is not yet supported".to_owned(),
            ));
        }
        if settings.animation_type == "Action" && !is_typed_action_effect(&settings.effect) {
            return Err(Error::ParseError(
                "Creating an unsupported Keynote action without its native parameters is not supported"
                    .to_owned(),
            ));
        }
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        let slides = self.slides()?;
        let slide = &slides[slide_index];
        let existing = self.slide_builds(slide_index)?;
        let existing_chunk_count = existing
            .iter()
            .map(|build| build.chunks.len())
            .sum::<usize>();
        validate_build_start_position(settings.start, existing_chunk_count)?;
        let graph = ObjectGraph::read(self.package())?;
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let build_id = next_object_identifier(self.package())?;
        let chunk_id = build_id
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let (build_uuid, random_number_seed) = new_build_uuid_and_seed();
        let build = new_build_archive(drawable_object_id, &settings, random_number_seed);
        let chunk = new_build_chunk(build_id, build_uuid, &settings);

        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            archive.insert_object(ArchiveObject::new(
                build_id,
                vec![RawMessage {
                    type_: BUILD_MESSAGE_TYPE,
                    data: build.encode_to_vec(),
                }],
            )?)?;
            archive.insert_object(ArchiveObject::new(
                chunk_id,
                vec![RawMessage {
                    type_: BUILD_CHUNK_MESSAGE_TYPE,
                    data: chunk.encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })?;
        patch_slide_build_references(
            &mut staged,
            &archive_name,
            slide.slide_id,
            &[],
            &[],
            &[(build_id, chunk_id)],
        )?;
        add_component_object_uuids(&mut staged, slide.slide_id, &[build_id])?;
        set_package_last_object_identifier(&mut staged, chunk_id)?;
        patch_slide_build_cache(&mut staged, &graph, slide.node_id, existing_chunk_count + 1)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slide_builds(slide_index)?
            .into_iter()
            .find(|build| build.object_id == build_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote build creation failed validation".to_owned())
            })?;
        if created.drawable_object_id != drawable_object_id
            || created.settings != settings
            || created.chunks.len() != 1
            || created.chunks[0].object_id != chunk_id
        {
            return Err(Error::InvalidFormat(
                "Keynote build creation failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(created)
    }

    /// Update an existing build and every timing chunk it owns.
    pub fn set_slide_build(
        &mut self,
        slide_index: usize,
        build_object_id: u64,
        settings: KeynoteBuildSettings,
    ) -> Result<()> {
        validate_build_settings(&settings)?;
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let build = self
            .slide_builds(slide_index)?
            .into_iter()
            .find(|build| build.object_id == build_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote build {build_object_id} is not owned by slide {slide_index}"
                ))
            })?;
        if typed_action_acceleration(&settings) == Some(KeynoteBuildAcceleration::Custom)
            && typed_action_acceleration(&build.settings) != Some(KeynoteBuildAcceleration::Custom)
        {
            return Err(Error::ParseError(
                "Changing a Keynote action to a custom curve requires native curve data".to_owned(),
            ));
        }
        if build.settings.animation_type == "Action"
            && !is_typed_action_effect(&build.settings.effect)
            && (settings.animation_type != build.settings.animation_type
                || settings.effect != build.settings.effect)
        {
            return Err(Error::ParseError(
                "Changing an unsupported Keynote action could discard opaque native parameters"
                    .to_owned(),
            ));
        }
        if settings.animation_type == "Action"
            && !is_typed_action_effect(&settings.effect)
            && (build.settings.animation_type != settings.animation_type
                || build.settings.effect != settings.effect)
        {
            return Err(Error::ParseError(
                "Changing to an unsupported Keynote action without its native parameters is not supported"
                    .to_owned(),
            ));
        }
        if build.chunks.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "Keynote build {build_object_id} has no timing chunks"
            )));
        }
        let graph = ObjectGraph::read(self.package())?;
        let native_slide: kn::SlideArchive = graph.decode(slide.slide_id, "KN.SlideArchive")?;
        let build_chunk_ids = build
            .chunks
            .iter()
            .map(|chunk| chunk.object_id)
            .collect::<HashSet<_>>();
        let event_index = native_slide
            .build_chunks
            .iter()
            .position(|reference| build_chunk_ids.contains(&reference.identifier))
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote build {build_object_id} has no ordered timing chunk"
                ))
            })?;
        validate_build_start_position(settings.start, event_index)?;
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(build_object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote build {build_object_id} is missing"))
            })?;
            normalize_and_patch_build_object(object, &settings)?;
            for chunk in &build.chunks {
                let object = archive.object_mut(chunk.object_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote build chunk {} is missing",
                        chunk.object_id
                    ))
                })?;
                patch_build_chunk_object(object, &settings)?;
            }
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let updated = verified
            .slide_builds(slide_index)?
            .into_iter()
            .find(|candidate| candidate.object_id == build_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote build update lost the build".to_owned())
            })?;
        if updated.settings != settings {
            return Err(Error::InvalidFormat(
                "Keynote build update failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Replace the object-build order for a slide with an exact permutation.
    ///
    /// Every build currently owned by the slide must appear exactly once. Timing
    /// chunks belonging to a build move with it while retaining their internal
    /// order. Slides whose chunks interleave multiple builds are rejected because
    /// grouping those events would silently change their native timing semantics.
    pub fn reorder_slide_builds(
        &mut self,
        slide_index: usize,
        build_object_ids: &[u64],
    ) -> Result<()> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let builds = self.slide_builds(slide_index)?;
        if build_object_ids.len() != builds.len() {
            return Err(Error::ParseError(format!(
                "Keynote slide {slide_index} build order contains {} identifiers, expected {}",
                build_object_ids.len(),
                builds.len()
            )));
        }
        let expected = builds
            .iter()
            .map(|build| build.object_id)
            .collect::<HashSet<_>>();
        let requested = build_object_ids.iter().copied().collect::<HashSet<_>>();
        if requested.len() != build_object_ids.len() || requested != expected {
            return Err(Error::ParseError(format!(
                "Keynote slide {slide_index} build order must be an exact permutation of its builds"
            )));
        }

        let graph = ObjectGraph::read(self.package())?;
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let native_slide: kn::SlideArchive = graph.decode(slide.slide_id, "KN.SlideArchive")?;
        let chunk_owner = builds
            .iter()
            .flat_map(|build| {
                build
                    .chunks
                    .iter()
                    .map(move |chunk| (chunk.object_id, build.object_id))
            })
            .collect::<HashMap<_, _>>();
        let mut closed_builds = HashSet::new();
        let mut active_build = None;
        for reference in &native_slide.build_chunks {
            let owner = chunk_owner.get(&reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote build chunk {} has no resolved owner",
                    reference.identifier
                ))
            })?;
            if active_build != Some(*owner) {
                if closed_builds.contains(owner) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_index} interleaves timing chunks for build {owner}"
                    )));
                }
                if let Some(previous) = active_build.replace(*owner) {
                    closed_builds.insert(previous);
                }
            }
        }

        let chunks_by_build = builds
            .iter()
            .map(|build| {
                (
                    build.object_id,
                    build
                        .chunks
                        .iter()
                        .map(|chunk| chunk.object_id)
                        .collect::<Vec<_>>(),
                )
            })
            .collect::<HashMap<_, _>>();
        let ordered_chunk_ids = build_object_ids
            .iter()
            .flat_map(|identifier| chunks_by_build[identifier].iter().copied())
            .collect::<Vec<_>>();
        let builds_by_identifier = builds
            .iter()
            .map(|build| (build.object_id, build))
            .collect::<HashMap<_, _>>();
        let mut event_index = 0usize;
        for identifier in build_object_ids {
            let build = builds_by_identifier[identifier];
            if !build.chunks.is_empty() {
                validate_build_start_position(build.settings.start, event_index)?;
                event_index = event_index.checked_add(build.chunks.len()).ok_or_else(|| {
                    Error::InvalidFormat("Keynote build-event count overflow".to_owned())
                })?;
            }
        }

        let mut staged = self.package().clone();
        patch_slide_build_order_references(
            &mut staged,
            &archive_name,
            slide.slide_id,
            build_object_ids,
            &ordered_chunk_ids,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let verified_ids = verified
            .slide_builds(slide_index)?
            .iter()
            .map(|build| build.object_id)
            .collect::<Vec<_>>();
        if verified_ids != build_object_ids {
            return Err(Error::InvalidFormat(
                "Keynote build reorder failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Move one object build to a zero-based position in the slide build order.
    pub fn move_slide_build(
        &mut self,
        slide_index: usize,
        build_object_id: u64,
        target_index: usize,
    ) -> Result<()> {
        let builds = self.slide_builds(slide_index)?;
        if target_index >= builds.len() {
            return Err(Error::ParseError(format!(
                "Keynote build target index {target_index} is out of range for {} builds",
                builds.len()
            )));
        }
        let mut identifiers = builds
            .iter()
            .map(|build| build.object_id)
            .collect::<Vec<_>>();
        let source_index = identifiers
            .iter()
            .position(|identifier| *identifier == build_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote build {build_object_id} is not owned by slide {slide_index}"
                ))
            })?;
        let identifier = identifiers.remove(source_index);
        identifiers.insert(target_index, identifier);
        self.reorder_slide_builds(slide_index, &identifiers)
    }

    /// Remove a build, all of its timing chunks, and its component UUID entry.
    pub fn remove_slide_build(
        &mut self,
        slide_index: usize,
        build_object_id: u64,
    ) -> Result<KeynoteBuildInfo> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let builds = self.slide_builds(slide_index)?;
        let removed = builds
            .iter()
            .find(|build| build.object_id == build_object_id)
            .cloned()
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote build {build_object_id} is not owned by slide {slide_index}"
                ))
            })?;
        let remaining_chunk_count = builds
            .iter()
            .filter(|build| build.object_id != build_object_id)
            .map(|build| build.chunks.len())
            .sum::<usize>();
        let mut event_index = 0usize;
        for build in builds
            .iter()
            .filter(|build| build.object_id != build_object_id)
        {
            if !build.chunks.is_empty() {
                validate_build_start_position(build.settings.start, event_index)?;
                event_index = event_index.checked_add(build.chunks.len()).ok_or_else(|| {
                    Error::InvalidFormat("Keynote build-event count overflow".to_owned())
                })?;
            }
        }
        let graph = ObjectGraph::read(self.package())?;
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let chunk_ids = removed
            .chunks
            .iter()
            .map(|chunk| chunk.object_id)
            .collect::<Vec<_>>();
        let mut staged = self.package().clone();
        patch_slide_build_references(
            &mut staged,
            &archive_name,
            slide.slide_id,
            &[build_object_id],
            &chunk_ids,
            &[],
        )?;
        for identifier in &chunk_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote build chunk {identifier} remains referenced outside its slide"
                )));
            }
            remove_object(&mut staged, &archive_name, *identifier)?;
        }
        if package_references_object(&staged, build_object_id)? {
            return Err(Error::InvalidFormat(format!(
                "Keynote build {build_object_id} remains referenced outside its slide"
            )));
        }
        remove_object(&mut staged, &archive_name, build_object_id)?;
        if component_uuid_identifiers(&staged, slide.slide_id)?
            .is_some_and(|identifiers| identifiers.contains(&build_object_id))
        {
            remove_component_object_uuids(&mut staged, slide.slide_id, &[build_object_id])?;
        }
        patch_slide_build_cache(&mut staged, &graph, slide.node_id, remaining_chunk_count)?;
        let mut released = chunk_ids;
        released.push(build_object_id);
        release_package_identifier_suffix(&mut staged, &released)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_builds(slide_index)?
            .iter()
            .any(|build| build.object_id == build_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote build removal failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(removed)
    }

    /// Replace the standard transition fields for a slide with a modern
    /// `AnimationAttributesArchive`. Legacy-only transitions are rejected so
    /// that the editor never guesses which representation should take precedence.
    pub fn set_slide_transition(
        &mut self,
        slide_index: usize,
        settings: KeynoteTransitionSettings,
    ) -> Result<()> {
        validate_transition_settings(&settings)?;
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        if slide.transition.is_none() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_index} has no modern transition attributes"
            )));
        }
        let slide_id = slide.slide_id;
        let graph = ObjectGraph::read(self.text.package())?;
        let archive_name = graph.archive_name(slide_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(slide_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == 5 && kn::SlideArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote slide object {slide_id} has no SlideArchive payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let slide = kn::SlideArchive::decode(original)?;
            let attributes = &slide.transition.attributes;
            if attributes.animation_attributes.is_none() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} has no modern transition attributes"
                )));
            }
            let data = patch_transition_settings_wire(original, attributes, &settings)?;
            let verified = kn::SlideArchive::decode(data.as_slice())?;
            let verified = transition_settings_from_wire(&data, &verified.transition.attributes)?;
            if verified.as_ref() != Some(&settings) {
                return Err(Error::InvalidFormat(
                    "Keynote transition wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slides()?
            .get(slide_index)
            .and_then(|slide| slide.transition.as_ref())
            != Some(&settings)
        {
            return Err(Error::InvalidFormat(
                "Keynote transition update failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    pub fn set_slide_title(&mut self, slide_index: usize, replacement: &str) -> Result<()> {
        let storage_id = self.slide_storage(slide_index, true)?;
        self.text.set_text(storage_id, replacement)
    }

    pub fn replace_slide_title(
        &mut self,
        slide_index: usize,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.slide_storage(slide_index, true)?;
        self.text.replace_text(storage_id, range, replacement)
    }

    pub fn clear_slide_title(&mut self, slide_index: usize) -> Result<()> {
        self.set_slide_title(slide_index, "")
    }

    pub fn set_slide_body(&mut self, slide_index: usize, replacement: &str) -> Result<()> {
        let storage_id = self.slide_storage(slide_index, false)?;
        self.text.set_text(storage_id, replacement)
    }

    pub fn replace_slide_body(
        &mut self,
        slide_index: usize,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.slide_storage(slide_index, false)?;
        self.text.replace_text(storage_id, range, replacement)
    }

    pub fn clear_slide_body(&mut self, slide_index: usize) -> Result<()> {
        self.set_slide_body(slide_index, "")
    }

    pub fn set_slide_notes(&mut self, slide_index: usize, replacement: &str) -> Result<()> {
        let storage_id = self.slide_notes_storage(slide_index)?;
        self.text.set_text(storage_id, replacement)
    }

    pub fn replace_slide_notes(
        &mut self,
        slide_index: usize,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.slide_notes_storage(slide_index)?;
        self.text.replace_text(storage_id, range, replacement)
    }

    pub fn clear_slide_notes(&mut self, slide_index: usize) -> Result<()> {
        self.set_slide_notes(slide_index, "")
    }

    /// Set or clear a slide's optional navigator name.
    pub fn set_slide_name(&mut self, slide_index: usize, name: Option<&str>) -> Result<()> {
        if name.is_some_and(|name| name.contains('\0')) {
            return Err(Error::ParseError(
                "Keynote slide names cannot contain NUL".to_owned(),
            ));
        }
        let normalized = name.filter(|name| !name.is_empty()).map(str::to_owned);
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.text.package())?;
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let slide_id = slide.slide_id;
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(slide_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == 5 && kn::SlideArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote slide object {slide_id} has no SlideArchive payload"
                    ))
                })?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let value = kn::SlideArchive::decode(original)?;
            let data = patch_length_delimited_field(
                original,
                10,
                value.name.is_some(),
                normalized.as_deref().map(str::as_bytes),
            )?;
            let verified = kn::SlideArchive::decode(data.as_slice())?;
            if verified.name.as_ref() != normalized.as_ref() {
                return Err(Error::InvalidFormat(
                    "Keynote slide-name wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slides()?
            .get(slide_index)
            .and_then(|slide| slide.name.as_ref())
            != normalized.as_ref()
        {
            return Err(Error::InvalidFormat(
                "Keynote slide rename failed validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Move a slide to another zero-based position in the presentation.
    pub fn move_slide(&mut self, from: usize, to: usize) -> Result<()> {
        let slides = self.slides()?;
        if from >= slides.len() || to >= slides.len() {
            return Err(Error::ParseError(format!(
                "Keynote slide move {from} -> {to} is out of range for {} slides",
                slides.len()
            )));
        }
        if from == to {
            return Ok(());
        }

        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let show_archive = graph.archive_name(show_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&show_archive, |archive| {
            let object = archive.object_mut(show_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| kn::ShowArchive::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show object {show_id} has no ShowArchive payload"
                    ))
                })?;
            let show = kn::ShowArchive::decode(object.messages[message_index].data.as_slice())?;
            if show.slide_tree.slides.len() != slides.len() {
                return Err(Error::InvalidFormat(
                    "Keynote slide tree changed during move".to_owned(),
                ));
            }
            let mut desired = show
                .slide_tree
                .slides
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let identifier = desired.remove(from);
            desired.insert(to, identifier);
            let message_type = object.messages[message_index].type_;
            let data = rewrite_show_slide_references(
                object.messages[message_index].data.as_slice(),
                &show.slide_tree.slides,
                &desired,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;

        let bytes = staged.to_bytes()?;
        let verified = KeynoteEditor::from_bytes(&bytes)?;
        let expected_id = slides[from].slide_id;
        if verified.slides()?.get(to).map(|slide| slide.slide_id) != Some(expected_id) {
            return Err(Error::InvalidFormat(
                "Keynote slide move failed validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }

    /// Duplicate a slide immediately after its source using a guarded graph clone.
    ///
    /// All object IDs internal to the dedicated slide component are remapped.
    /// External style and media references are shared, while stale thumbnails
    /// are removed from the new slide node so Keynote can regenerate them.
    pub fn duplicate_slide(&mut self, slide_index: usize) -> Result<KeynoteSlideInfo> {
        let slides = self.slides()?;
        let source = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.text.package())?;
        let source_archive_name = graph.archive_name(source.slide_id)?.to_owned();
        let expected_archive_name = format!("Index/Slide-{}.iwa", source.slide_id);
        if source_archive_name != expected_archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {} is not stored in its dedicated component {expected_archive_name}",
                source.slide_id
            )));
        }
        let source_archive = self.text.package().archive(&source_archive_name)?;
        let mut next_identifier = graph
            .objects
            .keys()
            .copied()
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let new_node_id = next_identifier;
        next_identifier = next_identifier
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let mut remap = HashMap::with_capacity(source_archive.objects.len());
        for object in &source_archive.objects {
            let old_identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Object in {source_archive_name} has no identifier"))
            })?;
            remap.insert(old_identifier, next_identifier);
            next_identifier = next_identifier
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        }
        let new_slide_id = *remap.get(&source.slide_id).ok_or_else(|| {
            Error::InvalidFormat(
                "Keynote slide component does not contain its slide object".to_owned(),
            )
        })?;
        let cloned_objects = source_archive
            .objects
            .iter()
            .map(|object| clone_slide_object(object, &remap))
            .collect::<Result<Vec<_>>>()?;
        let mut staged = self.text.package().clone();
        let new_archive_name = format!("Index/Slide-{new_slide_id}.iwa");
        if staged.contains_entry(&new_archive_name) {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide component {new_archive_name} already exists"
            )));
        }
        staged.replace_archive(
            &new_archive_name,
            &Archive {
                objects: cloned_objects,
            },
        )?;

        let node_archive_name = graph.archive_name(source.node_id)?.to_owned();
        let node_archive = self.text.package().archive(&node_archive_name)?;
        let source_node = node_archive.object(source.node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {} is missing", source.node_id))
        })?;
        let new_node = clone_slide_node(source_node, new_node_id, source.slide_id, new_slide_id)?;
        staged.update_archive(&node_archive_name, |archive| {
            archive.insert_object(new_node)?;
            Ok(())
        })?;

        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let show_archive_name = graph.archive_name(show_id)?.to_owned();
        staged.update_archive(&show_archive_name, |archive| {
            let object = archive.object_mut(show_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| kn::ShowArchive::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show object {show_id} has no ShowArchive payload"
                    ))
                })?;
            let show = kn::ShowArchive::decode(object.messages[message_index].data.as_slice())?;
            let source_position = show
                .slide_tree
                .slides
                .iter()
                .position(|reference| reference.identifier == source.node_id)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show does not reference source node {}",
                        source.node_id
                    ))
                })?;
            let mut desired = show
                .slide_tree
                .slides
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            desired.insert(source_position + 1, new_node_id);
            let message_type = object.messages[message_index].type_;
            let data = rewrite_show_slide_references(
                object.messages[message_index].data.as_slice(),
                &show.slide_tree.slides,
                &desired,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            let references =
                &mut object.archive_info.message_infos[message_index].object_references;
            if !references.contains(&new_node_id) {
                references.push(new_node_id);
            }
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                if field.object_references.contains(&source.node_id)
                    && !field.object_references.contains(&new_node_id)
                {
                    field.object_references.push(new_node_id);
                }
            }
            Ok(())
        })?;

        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slides()?
            .into_iter()
            .find(|slide| slide.slide_id == new_slide_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote slide duplication failed validation".to_owned())
            })?;
        if created.index != slide_index + 1 {
            return Err(Error::InvalidFormat(
                "Keynote duplicated slide has the wrong order".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(created)
    }

    /// Remove a slide and its slide-tree node.
    ///
    /// A dedicated `Index/Slide-<id>.iwa` component is removed in one operation;
    /// otherwise only the slide object is removed so unrelated colocated objects
    /// are preserved. Removing the final slide is rejected.
    pub fn remove_slide(&mut self, slide_index: usize) -> Result<KeynoteSlideInfo> {
        let slides = self.slides()?;
        if slides.len() <= 1 {
            return Err(Error::ParseError(
                "Cannot remove the final Keynote slide".to_owned(),
            ));
        }
        let removed = slides.get(slide_index).cloned().ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;

        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let show_archive = graph.archive_name(show_id)?.to_owned();
        let node_archive = graph.archive_name(removed.node_id)?.to_owned();
        let slide_archive = graph.archive_name(removed.slide_id)?.to_owned();
        let mut staged = self.text.package().clone();

        staged.update_archive(&show_archive, |archive| {
            let object = archive.object_mut(show_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| kn::ShowArchive::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show object {show_id} has no ShowArchive payload"
                    ))
                })?;
            let show = kn::ShowArchive::decode(object.messages[message_index].data.as_slice())?;
            let desired = show
                .slide_tree
                .slides
                .iter()
                .filter(|reference| reference.identifier != removed.node_id)
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            if desired.len() + 1 != show.slide_tree.slides.len() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote show does not contain slide node {} exactly once",
                    removed.node_id
                )));
            }
            let message_type = object.messages[message_index].type_;
            let data = rewrite_show_slide_references(
                object.messages[message_index].data.as_slice(),
                &show.slide_tree.slides,
                &desired,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            object.archive_info.message_infos[message_index]
                .object_references
                .retain(|&identifier| identifier != removed.node_id);
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                field
                    .object_references
                    .retain(|&identifier| identifier != removed.node_id);
            }
            Ok(())
        })?;

        remove_object(&mut staged, &node_archive, removed.node_id)?;
        let dedicated_slide_archive = format!("Index/Slide-{}.iwa", removed.slide_id);
        if slide_archive == dedicated_slide_archive {
            staged.remove_entry(&slide_archive).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {slide_archive} is missing"
                ))
            })?;
        } else {
            remove_object(&mut staged, &slide_archive, removed.slide_id)?;
        }

        let bytes = staged.to_bytes()?;
        let verified = KeynoteEditor::from_bytes(&bytes)?;
        if verified.slides()?.len() + 1 != slides.len() {
            return Err(Error::InvalidFormat(
                "Keynote slide deletion failed validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(removed)
    }

    pub fn package(&self) -> &IWorkPackage {
        self.text.package()
    }

    /// List metadata-backed media reachable from this presentation package.
    pub fn media_assets(&self) -> Result<Vec<EmbeddedMediaAsset>> {
        reachable_embedded_assets(self.package(), [1])
    }

    /// List media reachable from one slide node and its owned object graph.
    pub fn slide_media_assets(&self, slide_index: usize) -> Result<Vec<EmbeddedMediaAsset>> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        reachable_embedded_assets(self.package(), [slide.node_id, slide.slide_id])
    }

    pub fn extract_media(&self, data_identifier: u64) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Keynote object graph"
            )));
        }
        IWorkMediaEditor::from_package(self.package().clone())?.extract(data_identifier)
    }

    /// Replace a referenced materialized asset without changing its data identifier.
    pub fn replace_media(&mut self, data_identifier: u64, replacement: &[u8]) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Keynote object graph"
            )));
        }
        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let old = media.replace(data_identifier, replacement)?;
        let staged = media.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(old)
    }

    pub fn into_package(self) -> IWorkPackage {
        self.text.into_package()
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.text.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.text.save(path)
    }

    fn require_slide_text_storage(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<u64> {
        self.slide_text_storages(slide_index)?
            .into_iter()
            .find(|text| text.drawable_object_id == drawable_object_id)
            .map(|text| text.storage.object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "drawable object {drawable_object_id} does not own writable text on Keynote slide {slide_index}"
                ))
            })
    }

    #[allow(deprecated)]
    fn text_box_graph(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<KeynoteTextBoxGraph> {
        let slides = self.slides()?;
        let slide_info = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let text = self
            .slide_text_storages(slide_index)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "drawable object {drawable_object_id} does not own writable text on Keynote slide {slide_index}"
                ))
            })?;
        if text.role != KeynoteSlideTextRole::TextBox {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide drawable {drawable_object_id} is a {:?} placeholder, not an ordinary text box",
                text.role
            )));
        }

        let object_graph = ObjectGraph::read(self.package())?;
        let slide: kn::SlideArchive =
            object_graph.decode(slide_info.slide_id, "KN.SlideArchive")?;
        for (name, references) in [
            ("owned_drawables", &slide.owned_drawables),
            ("drawables_z_order", &slide.drawables_z_order),
        ] {
            let matches = references
                .iter()
                .filter(|reference| reference.identifier == drawable_object_id)
                .count();
            if matches != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {} {name} must contain text box {drawable_object_id} exactly once",
                    slide_info.slide_id
                )));
            }
        }

        let archive_name = object_graph.archive_name(slide_info.slide_id)?.to_owned();
        if object_graph.archive_name(drawable_object_id)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote text box {drawable_object_id} is outside slide component {archive_name}"
            )));
        }
        let archive = self.package().archive(&archive_name)?;
        let drawable = archive.object(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote text box {drawable_object_id} is missing"))
        })?;
        let shape_messages = drawable
            .messages
            .iter()
            .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if shape_messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote text box {drawable_object_id} must have exactly one shape payload"
            )));
        }
        let shape = tswp::ShapeInfoArchive::decode(shape_messages[0].data.as_slice())?;
        if shape.is_text_box != Some(true) {
            return Err(Error::InvalidFormat(format!(
                "Keynote drawable {drawable_object_id} is not marked as a text box"
            )));
        }
        let storage_id = shape
            .owned_storage
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote text box {drawable_object_id} has no owned storage"
                ))
            })?;
        if shape
            .deprecated_storage
            .as_ref()
            .map(|reference| reference.identifier)
            != Some(storage_id)
            || storage_id != text.storage.object_id
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote text box {drawable_object_id} has inconsistent storage ownership"
            )));
        }
        let title_id = shape
            .super_
            .super_
            .title
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote text box {drawable_object_id} has no title stand-in"
                ))
            })?;
        let caption_id = shape
            .super_
            .super_
            .caption
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote text box {drawable_object_id} has no caption stand-in"
                ))
            })?;
        let required = [drawable_object_id, caption_id, title_id, storage_id]
            .into_iter()
            .collect::<HashSet<_>>();
        if required.len() != 4 {
            return Err(Error::InvalidFormat(format!(
                "Keynote text box {drawable_object_id} has aliased private objects"
            )));
        }
        for (identifier, label) in [(caption_id, "caption"), (title_id, "title")] {
            if object_graph.archive_name(identifier)? != archive_name {
                return Err(Error::InvalidFormat(format!(
                    "Keynote text-box {label} object {identifier} is outside {archive_name}"
                )));
            }
            let object = archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote text-box {label} object {identifier} is missing"
                ))
            })?;
            if object
                .messages
                .iter()
                .filter(|message| message.type_ == STANDIN_CAPTION_MESSAGE_TYPE)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote text-box {label} object {identifier} must have exactly one stand-in payload"
                )));
            }
        }
        if object_graph.archive_name(storage_id)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote text-box storage {storage_id} is outside {archive_name}"
            )));
        }
        let storage = archive.object(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote text-box storage {storage_id} is missing"))
        })?;
        if storage
            .messages
            .iter()
            .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote text-box storage {storage_id} must have exactly one writable payload"
            )));
        }
        let object_ids = archive
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .filter(|identifier| required.contains(identifier))
            .collect::<Vec<_>>();
        if object_ids.len() != 4 {
            return Err(Error::InvalidFormat(format!(
                "Keynote text box {drawable_object_id} has an incomplete private graph"
            )));
        }
        for name in self.package().iwa_entry_names() {
            for object in self.package().archive(name)?.objects {
                let owner = object.archive_info.identifier.ok_or_else(|| {
                    Error::Archive(format!("Object in {name} has no archive identifier"))
                })?;
                if required.contains(&owner) || owner == slide_info.slide_id {
                    continue;
                }
                if object.archive_info.message_infos.iter().any(|info| {
                    info.object_references
                        .iter()
                        .chain(
                            info.field_infos
                                .iter()
                                .flat_map(|field| &field.object_references),
                        )
                        .any(|identifier| required.contains(identifier))
                }) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote text-box graph {drawable_object_id} is referenced by external object {owner}"
                    )));
                }
            }
        }
        let uuid_object_ids = component_uuid_identifiers(self.package(), slide_info.slide_id)?
            .map(|mapped| {
                if required.iter().all(|identifier| mapped.contains(identifier)) {
                    Ok(object_ids.clone())
                } else {
                    Err(Error::InvalidFormat(format!(
                        "Keynote slide {} UUID map does not cover text-box graph {drawable_object_id}",
                        slide_info.slide_id
                    )))
                }
            })
            .transpose()?
            .unwrap_or_default();
        Ok(KeynoteTextBoxGraph {
            slide_id: slide_info.slide_id,
            archive_name,
            drawable_id: drawable_object_id,
            storage_id,
            object_ids,
            uuid_object_ids,
        })
    }

    fn slide_storage(&self, slide_index: usize, title: bool) -> Result<u64> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let (kind, storage) = if title {
            ("title", slide.title_storage_id)
        } else {
            ("body", slide.body_storage_id)
        };
        storage.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide {slide_index} has no writable {kind} placeholder storage"
            ))
        })
    }

    fn slide_notes_storage(&self, slide_index: usize) -> Result<u64> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        slide.notes_storage_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide {slide_index} has no writable speaker-notes storage"
            ))
        })
    }

    fn slide_owned_drawable_ids(&self, slide_index: usize) -> Result<Vec<u64>> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let archive: kn::SlideArchive = graph.decode(slide.slide_id, "KN.SlideArchive")?;
        Ok(archive
            .owned_drawables
            .into_iter()
            .map(|drawable| drawable.identifier)
            .collect())
    }

    fn require_slide_drawable(&self, slide_index: usize, drawable_object_id: u64) -> Result<()> {
        if !self
            .slide_owned_drawable_ids(slide_index)?
            .contains(&drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "drawable object {drawable_object_id} is not owned by Keynote slide {slide_index}"
            )));
        }
        if !self
            .slide_drawables(slide_index)?
            .iter()
            .any(|drawable| drawable.object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide drawable {drawable_object_id} has no supported direct drawable payload"
            )));
        }
        Ok(())
    }
}

mod builds;
mod slide_graph;

use builds::*;
use slide_graph::*;
#[cfg(test)]
mod tests;
