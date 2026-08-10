//! Semantic text editing for existing Keynote slides.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::path::Path;
use std::rc::Rc;

use litchi_iwa_common::comment::{
    DrawableComment, DrawableId, DrawableInfo, DrawableReply, StorageId,
};
use litchi_iwa_text::columns::Columns;
use litchi_iwa_text::paragraph::drop_cap::{DropCap, Placement};
use litchi_iwa_text::position::TextPosition;
use prost::Message;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::IWorkDrawableCommentEditor;
use crate::media::MediaAssetId;
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    add_component_external_reference, add_component_link, add_component_object_uuids,
    clone_component_registration, component_identifier_for_entry,
    component_identifier_for_object_uuid, component_uuid_identifiers, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, remove_component_registration,
    set_package_last_object_identifier,
};
use crate::protobuf::{kn, tsd, tsp, tss, tswp};
use crate::shapes::{
    DrawableGeometry, DrawableProperties, RgbaColor, reset_shape_text_columns,
    reset_shape_text_layout, set_shape_geometry, set_shape_properties, set_shape_text_columns,
    set_shape_text_layout, shape_geometry, shape_properties, shape_text_columns, shape_text_layout,
};
use crate::text::layout::Layout;
use crate::text::{
    Alignment, Background as TextAppearanceBackground, Borders, IWorkTextEditor, Indents,
    LineSpacing, Outline, ParagraphBackground, ParagraphDecimalTabCharacter,
    ParagraphDefaultTabInterval, ParagraphFlow, ParagraphList, ParagraphListBullet,
    ParagraphListBulletGeometry, ParagraphListIndentation, ParagraphListLabelColor,
    ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumberFormat,
    ParagraphListNumberScale, ParagraphListNumberTiering, ParagraphListNumbering,
    ParagraphTabStops, ParagraphWritingDirection, Shadow, Spacing, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextComment, TextCommentBody, TextCommentId,
    TextCommentReply, TextCommentReplyBody, TextCommentReplyId, TextDecorations, TextFont,
    TextHighlight, TextHighlightId, TextHyperlink, TextHyperlinkId, TextHyperlinkTarget,
    TextLanguage, TextLanguageRun, TextLigatures, TextRange, TextScript, TextStorageId,
    TextStorageInfo, TextStyle,
};
use crate::wire::{
    append_repeated_length_delimited_field, parse_wire_fields, patch_fixed32_field,
    patch_fixed64_field, patch_length_delimited_field, patch_nested_fixed32_field,
    patch_nested_fixed64_field, patch_nested_length_delimited_field, patch_nested_varint_field,
    patch_varint_field, remove_repeated_length_delimited_field_where,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field, transform_length_delimited_fields_at_path,
};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};
use litchi_iwa_index::ObjectId;
use litchi_keynote::build as semantic_build;
pub use litchi_keynote::build::{Acceleration as BuildAcceleration, Start as BuildStart};
use litchi_keynote::transition::Settings as TransitionSettings;

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
const DISSOLVE_BUILD_EFFECT: &str = "apple:dissolve character";
const SHIMMER_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.KLNShimmer";
const SKID_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.KNBuildSkidByCharacter";
const SWOOSH_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.BLTSwoosh";
const TRACE_BUILD_EFFECT: &str = "com.apple.iWork.Keynote.Trace";
const DRAWABLE_DUPLICATE_OFFSET: f32 = 10.0;
const TABLE_DUPLICATE_OFFSET: f32 = DRAWABLE_DUPLICATE_OFFSET;
// These caches are deliberately operation-scoped.  Keeping only the compact
// slide ownership data lets `slide_text_storages` reuse the decode performed
// by `slides` without retaining decoded protobufs on a mutable editor.
const MAX_OPERATION_CACHED_SLIDES: usize = 128;
const MAX_OPERATION_CACHED_DRAWABLES_PER_SLIDE: usize = 512;
const MAX_OPERATION_CACHED_DRAWABLE_STORAGES: usize = 1_024;

/// Writable text targets resolved from one slide in presentation order.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideInfo {
    pub index: usize,
    pub node_id: u64,
    pub slide_id: u64,
    pub name: Option<String>,
    /// Theme layout currently selected for this slide.
    ///
    /// Legacy slides without a template relationship report `None`.
    pub layout: Option<KeynoteSlideLayoutInfo>,
    pub is_skipped: bool,
    pub is_slide_number_visible: Option<bool>,
    /// Whether the layout-provided title placeholder participates in this slide.
    ///
    /// `None` means the selected layout has no title placeholder. Hidden
    /// placeholders retain their storage, so [`Self::title`] remains readable.
    pub is_title_visible: Option<bool>,
    /// Whether the layout-provided body placeholder participates in this slide.
    ///
    /// `None` means the selected layout has no body placeholder. Hidden
    /// placeholders retain their storage, so [`Self::body`] remains readable.
    pub is_body_visible: Option<bool>,
    pub transition: Option<TransitionSettings>,
    pub title_storage_id: Option<TextStorageId>,
    pub title: Option<String>,
    pub body_storage_id: Option<TextStorageId>,
    pub body: Option<String>,
    pub notes_storage_id: Option<TextStorageId>,
    pub notes: Option<String>,
}

/// Stable identity of a slide layout in the presentation theme.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct KeynoteSlideLayoutId(ObjectId);

impl KeynoteSlideLayoutId {
    /// Construct a layout identity, rejecting Keynote's null reference sentinel.
    pub const fn new(raw: u64) -> Option<Self> {
        match ObjectId::new(raw) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    pub fn as_u64(self) -> u64 {
        self.0.get()
    }
}

/// A theme-provided Keynote slide layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeynoteSlideLayoutInfo {
    pub id: KeynoteSlideLayoutId,
    pub name: String,
    pub is_default: bool,
}

/// Semantic role of a writable text-bearing drawable owned by a slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSlideTextRole {
    Title,
    Body,
    TextBox,
    /// Writable text owned by an ordinary non-text-box shape.
    Shape,
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
    storage_id: TextStorageId,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
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
    pub acceleration: BuildAcceleration,
}

/// Typed parameters for Keynote's object Scale action.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteScaleAction {
    /// Final size as a factor of the object's original size (`1.5` is 150%).
    pub scale_factor: f64,
    pub acceleration: BuildAcceleration,
}

/// Typed parameters for Keynote's object Opacity action.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteOpacityAction {
    /// Final opacity in Keynote percentage points (`37.0` is 37%).
    pub opacity_percent: f64,
    pub acceleration: BuildAcceleration,
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

/// Editable normalized timing curve used by a custom Keynote action acceleration.
///
/// The path starts at `(0, 0)` and ends at `(1, 1)`. Its coordinates express
/// elapsed time on the horizontal axis and animation progress on the vertical
/// axis, matching Keynote's native custom-curve payload.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteBuildTimingCurve {
    pub path: KeynoteMotionPath,
}

impl KeynoteBuildTimingCurve {
    /// Construct a one-segment cubic Bézier timing curve.
    ///
    /// The control points are normalized to the timing-curve coordinate space;
    /// callers can use [`Self::from_path`] for a multi-segment curve.
    pub fn cubic(
        first_control_point: KeynoteMotionPathPoint,
        second_control_point: KeynoteMotionPathPoint,
    ) -> Self {
        let origin = KeynoteMotionPathPoint::new(0.0, 0.0);
        let destination = KeynoteMotionPathPoint::new(1.0, 1.0);
        Self {
            path: KeynoteMotionPath {
                subpaths: vec![KeynoteMotionSubpath {
                    nodes: vec![
                        KeynoteMotionPathNode {
                            in_control_point: origin,
                            point: origin,
                            out_control_point: first_control_point,
                            node_type: KeynoteMotionPathNodeType::Bezier,
                        },
                        KeynoteMotionPathNode {
                            in_control_point: second_control_point,
                            point: destination,
                            out_control_point: destination,
                            node_type: KeynoteMotionPathNodeType::Bezier,
                        },
                    ],
                    closed: false,
                }],
                natural_width: 1.0,
                natural_height: 1.0,
                horizontal_flip: false,
                vertical_flip: false,
            },
        }
    }

    /// Construct a linear timing curve.
    pub fn linear() -> Self {
        Self::cubic(
            KeynoteMotionPathPoint::new(0.0, 0.0),
            KeynoteMotionPathPoint::new(1.0, 1.0),
        )
    }

    /// Wrap an explicitly constructed native-compatible timing path.
    pub fn from_path(path: KeynoteMotionPath) -> Self {
        Self { path }
    }
}

/// Typed parameters for Keynote's object Move action.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteMoveAction {
    pub path: KeynoteMotionPath,
    pub align_to_path: bool,
    pub acceleration: BuildAcceleration,
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
    Dissolve,
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
    pub(crate) delivery: String,
    pub(crate) animation_type: String,
    pub(crate) effect: String,
    pub(crate) duration: f64,
    /// Delay before an `AfterTransition` or `AfterPrevious` build starts.
    pub(crate) delay: f64,
    pub(crate) start: BuildStart,
    pub(crate) direction: Option<u32>,
    /// Raw `BuildAttributesTextDelivery` value for forward compatibility.
    pub(crate) text_delivery: Option<i32>,
    /// Raw `BuildAttributesDeliveryOption` value for forward compatibility.
    pub(crate) delivery_option: Option<i32>,
    pub(crate) event_trigger: Option<u32>,
    /// Present for Keynote's native `apple:action-rotation` action effect.
    pub(crate) rotation: Option<KeynoteRotationAction>,
    /// Present for Keynote's native `apple:action-scale` action effect.
    pub(crate) scale: Option<KeynoteScaleAction>,
    /// Present for Keynote's native `apple:action-opacity` action effect.
    pub(crate) opacity: Option<KeynoteOpacityAction>,
    /// Present for Keynote's native `apple:action-motion-path` action effect.
    pub(crate) move_action: Option<KeynoteMoveAction>,
    /// Present for Keynote's Blink/Bounce/Flip/Jiggle/Pop/Pulse actions.
    pub(crate) emphasis: Option<KeynoteEmphasisAction>,
    /// Present for Keynote's native `apple:keyboard` build-in/build-out effect.
    pub(crate) keyboard: Option<KeynoteKeyboardBuild>,
    /// Present for typed Dissolve, Shimmer, Skid, Swoosh, and Trace builds.
    pub(crate) object_effect: Option<KeynoteObjectBuildEffect>,
    /// Inline curve for a typed action whose acceleration is
    /// [`BuildAcceleration::Custom`].
    ///
    /// `None` preserves an opaque app-native custom curve while updating an
    /// existing build. New custom-curve actions require `Some`.
    pub(crate) timing_curve: Option<KeynoteBuildTimingCurve>,
    /// Raw parameters for native effects without a dedicated typed model.
    pub(crate) custom_parameters: KeynoteBuildCustomParameters,
}

impl KeynoteBuildSettings {
    /// Project the native adapter state into the bounded archive-free model.
    ///
    /// Native identifiers, presence bits, and opaque parameter fields stay in
    /// this crate; callers receive only the validated semantic build value.
    pub fn semantic(&self) -> Result<semantic_build::Settings> {
        builds::semantic_settings(self)
    }

    /// Return the semantic effect after validation.
    pub fn effect(&self) -> Result<semantic_build::Effect> {
        self.semantic().map(|settings| settings.effect().clone())
    }

    /// Replace the native build effect from a checked semantic value.
    pub fn set_effect(&mut self, effect: litchi_keynote::build::Effect) -> Result<()> {
        let mut candidate = self.clone();
        apply_semantic_build_effect(&mut candidate, effect)?;
        builds::validate_build_settings(&candidate)?;
        *self = candidate;
        Ok(())
    }

    /// Return the semantic start relationship.
    pub fn start(&self) -> Result<BuildStart> {
        self.semantic().map(|settings| settings.start())
    }

    /// Return the semantic duration.
    pub fn duration(&self) -> Result<litchi_keynote::Seconds> {
        self.semantic().map(|settings| settings.duration())
    }

    /// Return the semantic delay.
    pub fn delay(&self) -> Result<litchi_keynote::Seconds> {
        self.semantic().map(|settings| settings.delay())
    }

    /// Replace the start relationship transactionally.
    pub fn set_start(&mut self, start: BuildStart) -> Result<()> {
        let mut candidate = self.semantic()?;
        candidate.set_start(start).map_err(|error| {
            Error::ParseError(format!("invalid Keynote build start relationship: {error}"))
        })?;
        self.start = start;
        Ok(())
    }

    /// Replace the delay transactionally.
    pub fn set_delay(&mut self, delay: litchi_keynote::Seconds) -> Result<()> {
        let mut candidate = self.semantic()?;
        candidate
            .set_delay(delay)
            .map_err(|error| Error::ParseError(format!("invalid Keynote build delay: {error}")))?;
        self.delay = delay.as_f64();
        Ok(())
    }

    /// Replace the duration with a validated semantic value.
    pub fn set_duration(&mut self, duration: litchi_keynote::Seconds) -> Result<()> {
        let mut candidate = self.semantic()?;
        candidate.set_duration(duration);
        candidate.validate().map_err(|error| {
            Error::ParseError(format!("invalid Keynote build duration: {error}"))
        })?;
        self.duration = duration.as_f64();
        Ok(())
    }

    /// Replace the timing acceleration of a typed Rotate, Scale, Opacity, or
    /// Move action.
    pub fn set_action_acceleration(&mut self, acceleration: BuildAcceleration) -> Result<()> {
        if acceleration.kind().is_none() {
            return Err(Error::ParseError(
                "Keynote action acceleration is not a recognized value".to_owned(),
            ));
        }
        let mut candidate = self.clone();
        let mut found = false;
        if let Some(action) = candidate.rotation.as_mut() {
            action.acceleration = acceleration;
            found = true;
        } else if let Some(action) = candidate.scale.as_mut() {
            action.acceleration = acceleration;
            found = true;
        } else if let Some(action) = candidate.opacity.as_mut() {
            action.acceleration = acceleration;
            found = true;
        } else if let Some(action) = candidate.move_action.as_mut() {
            action.acceleration = acceleration;
            found = true;
        }
        if !found {
            return Err(Error::ParseError(
                "Keynote action acceleration requires a typed action".to_owned(),
            ));
        }
        if acceleration != BuildAcceleration::Custom {
            candidate.timing_curve = None;
        }
        builds::validate_build_settings(&candidate)?;
        *self = candidate;
        Ok(())
    }

    /// Replace the editable path of a typed Move action.
    pub fn set_move_path(&mut self, path: KeynoteMotionPath) -> Result<()> {
        if self.move_action.is_none() {
            return Err(Error::ParseError(
                "Keynote Move path requires a typed Move action".to_owned(),
            ));
        }
        builds::validate_motion_path(&path)?;
        let mut candidate = self.clone();
        let Some(action) = candidate.move_action.as_mut() else {
            return Err(Error::ParseError(
                "Keynote Move path requires a typed Move action".to_owned(),
            ));
        };
        action.path = path;
        builds::validate_build_settings(&candidate)?;
        *self = candidate;
        Ok(())
    }

    /// Replace whether a typed Move action follows its path tangent.
    pub fn set_move_alignment(&mut self, align_to_path: bool) -> Result<()> {
        let mut candidate = self.clone();
        let Some(action) = candidate.move_action.as_mut() else {
            return Err(Error::ParseError(
                "Keynote Move alignment requires a typed Move action".to_owned(),
            ));
        };
        action.align_to_path = align_to_path;
        builds::validate_build_settings(&candidate)?;
        *self = candidate;
        Ok(())
    }

    /// Replace the start relationship in a consuming builder step.
    pub fn with_start(mut self, start: BuildStart) -> Result<Self> {
        self.set_start(start)?;
        Ok(self)
    }

    /// Replace the delay in a consuming builder step.
    pub fn with_delay(mut self, delay: litchi_keynote::Seconds) -> Result<Self> {
        self.set_delay(delay)?;
        Ok(self)
    }

    /// Replace the duration in a consuming builder step.
    pub fn with_duration(mut self, duration: litchi_keynote::Seconds) -> Result<Self> {
        self.set_duration(duration)?;
        Ok(self)
    }

    /// Native playback trigger attached to a newly inserted audio clip.
    pub(crate) fn audio_start() -> Self {
        Self {
            effect: "apple:audio-start".to_owned(),
            duration: 0.5,
            text_delivery: None,
            delivery_option: None,
            ..Self::appear_in()
        }
    }

    /// Native playback trigger attached to a newly inserted movie.
    pub(crate) fn movie_start() -> Self {
        Self {
            effect: "apple:movie-start".to_owned(),
            duration: 0.5,
            text_delivery: None,
            delivery_option: None,
            ..Self::appear_in()
        }
    }

    /// Native-compatible object-level Appear build-in settings.
    pub fn appear_in() -> Self {
        Self {
            delivery: "All at Once".to_owned(),
            animation_type: "In".to_owned(),
            effect: "apple:bc-appear".to_owned(),
            duration: 1.0,
            delay: 0.0,
            start: BuildStart::OnClick,
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
            timing_curve: None,
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

    fn text_object_build(
        animation_type: &str,
        effect: KeynoteObjectBuildEffect,
        duration: f64,
    ) -> Self {
        Self {
            animation_type: animation_type.to_owned(),
            effect: object_build_effect_identifier(effect).to_owned(),
            duration,
            direction: native_object_build_direction(effect),
            object_effect: Some(effect),
            ..Self::appear_in()
        }
    }

    fn object_build(animation_type: &str, effect: KeynoteObjectBuildEffect, duration: f64) -> Self {
        Self {
            text_delivery: None,
            delivery_option: None,
            ..Self::text_object_build(animation_type, effect, duration)
        }
    }

    /// Native-compatible Dissolve Build In.
    pub fn dissolve_in() -> Self {
        Self::text_object_build("In", KeynoteObjectBuildEffect::Dissolve, 1.0)
    }

    /// Native-compatible Dissolve Build Out.
    pub fn dissolve_out() -> Self {
        Self::text_object_build("Out", KeynoteObjectBuildEffect::Dissolve, 1.0)
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
                acceleration: BuildAcceleration::EaseInOut,
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
                acceleration: BuildAcceleration::EaseInOut,
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
                acceleration: BuildAcceleration::EaseInOut,
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
                acceleration: BuildAcceleration::EaseInOut,
            }),
            ..Self::appear_in()
        }
    }

    /// Attach a custom timing curve to a typed action.
    ///
    /// This changes the action's acceleration to
    /// [`BuildAcceleration::Custom`]. It returns an error when this is
    /// not a Rotate, Scale, Opacity, or Move action.
    pub fn with_custom_timing_curve(
        mut self,
        timing_curve: KeynoteBuildTimingCurve,
    ) -> Result<Self> {
        self.set_custom_timing_curve(timing_curve)?;
        Ok(self)
    }

    /// Replace the custom timing curve of a typed action.
    ///
    /// This changes the action's acceleration to
    /// [`BuildAcceleration::Custom`]. It returns an error when this is
    /// not a Rotate, Scale, Opacity, or Move action.
    pub fn set_custom_timing_curve(&mut self, timing_curve: KeynoteBuildTimingCurve) -> Result<()> {
        validate_timing_curve(&timing_curve)?;
        let acceleration = if let Some(action) = self.rotation.as_mut() {
            &mut action.acceleration
        } else if let Some(action) = self.scale.as_mut() {
            &mut action.acceleration
        } else if let Some(action) = self.opacity.as_mut() {
            &mut action.acceleration
        } else if let Some(action) = self.move_action.as_mut() {
            &mut action.acceleration
        } else {
            return Err(Error::ParseError(
                "Keynote custom timing curves require a typed action".to_owned(),
            ));
        };
        *acceleration = BuildAcceleration::Custom;
        self.timing_curve = Some(timing_curve);
        Ok(())
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

fn apply_semantic_build_effect(
    settings: &mut KeynoteBuildSettings,
    effect: litchi_keynote::build::Effect,
) -> Result<()> {
    effect
        .validate()
        .map_err(|error| Error::ParseError(format!("invalid Keynote build effect: {error}")))?;
    let previous_timing_curve = settings.timing_curve.clone();
    settings.direction = None;
    settings.rotation = None;
    settings.scale = None;
    settings.opacity = None;
    settings.move_action = None;
    settings.emphasis = None;
    settings.keyboard = None;
    settings.object_effect = None;
    settings.timing_curve = None;
    settings.custom_parameters = KeynoteBuildCustomParameters::default();

    match effect {
        litchi_keynote::build::Effect::Appear => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = "apple:bc-appear".to_owned();
        },
        litchi_keynote::build::Effect::Dissolve => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = "apple:bc-dissolve".to_owned();
        },
        litchi_keynote::build::Effect::MoveIn => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = "apple:bc-move".to_owned();
        },
        litchi_keynote::build::Effect::Scale => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = "apple:bc-scale".to_owned();
        },
        litchi_keynote::build::Effect::FadeAndScale => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = "apple:bc-fade-scale".to_owned();
        },
        litchi_keynote::build::Effect::Action(action) => {
            settings.animation_type = "Action".to_owned();
            settings.effect = action.identifier().to_owned();
            match action {
                litchi_keynote::build::Action::Rotate(action) => {
                    settings.rotation = Some(KeynoteRotationAction {
                        total_degrees: action.total_degrees(),
                        direction: match action.direction() {
                            litchi_keynote::build::RotationDirection::Clockwise => {
                                KeynoteRotationDirection::Clockwise
                            },
                            litchi_keynote::build::RotationDirection::Counterclockwise => {
                                KeynoteRotationDirection::Counterclockwise
                            },
                        },
                        acceleration: action.acceleration(),
                    });
                },
                litchi_keynote::build::Action::Scale(action) => {
                    settings.scale = Some(KeynoteScaleAction {
                        scale_factor: action.factor(),
                        acceleration: action.acceleration(),
                    });
                },
                litchi_keynote::build::Action::Opacity(action) => {
                    settings.opacity = Some(KeynoteOpacityAction {
                        opacity_percent: action.percent(),
                        acceleration: action.acceleration(),
                    });
                },
                litchi_keynote::build::Action::Move(action) => {
                    settings.move_action = Some(KeynoteMoveAction {
                        path: native_motion_path_from_semantic(action.path())?,
                        align_to_path: action.align_to_path(),
                        acceleration: action.acceleration(),
                    });
                },
                litchi_keynote::build::Action::Unknown(_) => {
                    return Err(Error::ParseError(
                        "Keynote unknown actions cannot be written without native parameters"
                            .to_owned(),
                    ));
                },
                _ => {
                    return Err(Error::ParseError(
                        "Keynote action variant is not supported by this adapter".to_owned(),
                    ));
                },
            }
        },
        litchi_keynote::build::Effect::Emphasis(emphasis) => {
            settings.animation_type = "Action".to_owned();
            settings.effect = emphasis.identifier().to_owned();
            let (action, direction) = native_emphasis_action(emphasis)?;
            settings.emphasis = Some(action);
            settings.direction = direction;
        },
        litchi_keynote::build::Effect::Keyboard(keyboard) => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = KEYBOARD_BUILD_EFFECT.to_owned();
            settings.direction = Some(builds::native_keyboard_direction(
                match keyboard.direction() {
                    litchi_keynote::build::KeyboardDirection::Forward => {
                        KeynoteKeyboardDirection::Forward
                    },
                    litchi_keynote::build::KeyboardDirection::Backward => {
                        KeynoteKeyboardDirection::Backward
                    },
                },
            ));
            settings.keyboard = Some(KeynoteKeyboardBuild {
                direction: match keyboard.direction() {
                    litchi_keynote::build::KeyboardDirection::Forward => {
                        KeynoteKeyboardDirection::Forward
                    },
                    litchi_keynote::build::KeyboardDirection::Backward => {
                        KeynoteKeyboardDirection::Backward
                    },
                },
                show_cursor: keyboard.show_cursor(),
            });
        },
        litchi_keynote::build::Effect::Object(object) => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            let object = native_object_build_effect(object)?;
            settings.effect = builds::object_build_effect_identifier(object).to_owned();
            settings.direction = builds::native_object_build_direction(object);
            settings.object_effect = Some(object);
        },
        litchi_keynote::build::Effect::Unknown(value) => {
            settings.animation_type = build_in_or_out(settings.animation_type.as_str()).to_owned();
            settings.effect = value.as_str().to_owned();
        },
        _ => {
            return Err(Error::ParseError(
                "Keynote effect variant is not supported by this adapter".to_owned(),
            ));
        },
    }

    if builds::typed_action_acceleration(settings) == Some(BuildAcceleration::Custom) {
        settings.timing_curve = previous_timing_curve;
    }
    Ok(())
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "Reject future semantic build variants until their native parameters are modeled."
)]
fn native_emphasis_action(
    emphasis: litchi_keynote::build::Emphasis,
) -> Result<(KeynoteEmphasisAction, Option<u32>)> {
    use litchi_keynote::build::{FlipDirection, JiggleIntensity};

    match emphasis {
        litchi_keynote::build::Emphasis::Blink(action) => Ok((
            KeynoteEmphasisAction::Blink {
                repeat_count: action.repeat_count(),
                fade: action.fade(),
            },
            None,
        )),
        litchi_keynote::build::Emphasis::Bounce(action) => Ok((
            KeynoteEmphasisAction::Bounce {
                repeat_count: action.repeat_count(),
                decay: action.decay(),
            },
            None,
        )),
        litchi_keynote::build::Emphasis::Flip(action) => Ok((
            KeynoteEmphasisAction::Flip {
                repeat_count: action.repeat_count(),
                direction: match action.direction() {
                    FlipDirection::LeftToRight => KeynoteFlipDirection::LeftToRight,
                    FlipDirection::RightToLeft => KeynoteFlipDirection::RightToLeft,
                },
            },
            Some(builds::native_flip_direction(match action.direction() {
                FlipDirection::LeftToRight => KeynoteFlipDirection::LeftToRight,
                FlipDirection::RightToLeft => KeynoteFlipDirection::RightToLeft,
            })),
        )),
        litchi_keynote::build::Emphasis::Jiggle(action) => Ok((
            KeynoteEmphasisAction::Jiggle {
                intensity: match action.intensity() {
                    JiggleIntensity::Small => KeynoteJiggleIntensity::Small,
                    JiggleIntensity::Medium => KeynoteJiggleIntensity::Medium,
                    JiggleIntensity::Large => KeynoteJiggleIntensity::Large,
                },
            },
            None,
        )),
        litchi_keynote::build::Emphasis::Pop(action) => Ok((
            KeynoteEmphasisAction::Pop {
                scale_percent: action.scale_percent(),
            },
            None,
        )),
        litchi_keynote::build::Emphasis::Pulse(action) => Ok((
            KeynoteEmphasisAction::Pulse {
                repeat_count: action.repeat_count(),
                scale_percent: action.scale_percent(),
            },
            None,
        )),
        litchi_keynote::build::Emphasis::Unknown(_) => Err(Error::ParseError(
            "Keynote unknown emphasis actions cannot be written without native parameters"
                .to_owned(),
        )),
        _ => Err(Error::ParseError(
            "Keynote emphasis variant is not supported by this adapter".to_owned(),
        )),
    }
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "Reject future semantic build variants until their native parameters are modeled."
)]
fn native_object_build_effect(
    effect: litchi_keynote::build::ObjectEffect,
) -> Result<KeynoteObjectBuildEffect> {
    use litchi_keynote::build::{HorizontalDirection, SwooshDirection};

    match effect {
        litchi_keynote::build::ObjectEffect::Dissolve => Ok(KeynoteObjectBuildEffect::Dissolve),
        litchi_keynote::build::ObjectEffect::Shimmer => Ok(KeynoteObjectBuildEffect::Shimmer),
        litchi_keynote::build::ObjectEffect::Skid(direction) => {
            Ok(KeynoteObjectBuildEffect::Skid {
                direction: match direction {
                    HorizontalDirection::LeftToRight => {
                        KeynoteHorizontalBuildDirection::LeftToRight
                    },
                    HorizontalDirection::RightToLeft => {
                        KeynoteHorizontalBuildDirection::RightToLeft
                    },
                },
            })
        },
        litchi_keynote::build::ObjectEffect::Swoosh(direction) => {
            Ok(KeynoteObjectBuildEffect::Swoosh {
                direction: match direction {
                    SwooshDirection::Center => KeynoteSwooshDirection::Center,
                    SwooshDirection::FromLeft => KeynoteSwooshDirection::FromLeft,
                    SwooshDirection::FromRight => KeynoteSwooshDirection::FromRight,
                },
            })
        },
        litchi_keynote::build::ObjectEffect::Trace(direction) => {
            Ok(KeynoteObjectBuildEffect::Trace {
                direction: match direction {
                    HorizontalDirection::LeftToRight => {
                        KeynoteHorizontalBuildDirection::LeftToRight
                    },
                    HorizontalDirection::RightToLeft => {
                        KeynoteHorizontalBuildDirection::RightToLeft
                    },
                },
            })
        },
        _ => Err(Error::ParseError(
            "Keynote object effect variant is not supported by this adapter".to_owned(),
        )),
    }
}

fn native_motion_path_from_semantic(
    path: &litchi_keynote::build::Path,
) -> Result<KeynoteMotionPath> {
    let finite_f32 = |value: f64| -> Result<f32> {
        if !value.is_finite() || value.abs() > f64::from(f32::MAX) {
            return Err(Error::ParseError(
                "Keynote Move path values must fit in finite native f32 coordinates".to_owned(),
            ));
        }
        Ok(value as f32)
    };
    let point = |point: litchi_keynote::build::Point| -> Result<KeynoteMotionPathPoint> {
        Ok(KeynoteMotionPathPoint::new(
            finite_f32(point.x())?,
            finite_f32(point.y())?,
        ))
    };
    let subpaths = path
        .subpaths()
        .iter()
        .map(|subpath| {
            let nodes = subpath
                .nodes()
                .iter()
                .map(|node| {
                    Ok(KeynoteMotionPathNode {
                        in_control_point: point(node.in_control_point())?,
                        point: point(node.point())?,
                        out_control_point: point(node.out_control_point())?,
                        node_type: match node.kind() {
                            litchi_keynote::build::NodeKind::Sharp => {
                                KeynoteMotionPathNodeType::Sharp
                            },
                            litchi_keynote::build::NodeKind::Bezier => {
                                KeynoteMotionPathNodeType::Bezier
                            },
                            litchi_keynote::build::NodeKind::Smooth => {
                                KeynoteMotionPathNodeType::Smooth
                            },
                        },
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            Ok(KeynoteMotionSubpath {
                nodes,
                closed: subpath.closed(),
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let native = KeynoteMotionPath {
        subpaths,
        natural_width: finite_f32(path.natural_width())?,
        natural_height: finite_f32(path.natural_height())?,
        horizontal_flip: path.horizontal_flip(),
        vertical_flip: path.vertical_flip(),
    };
    builds::validate_motion_path(&native)?;
    Ok(native)
}

fn build_in_or_out(animation_type: &str) -> &'static str {
    if animation_type == "Out" { "Out" } else { "In" }
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

#[derive(Debug)]
struct CachedSlideGraph {
    title_placeholder: Option<u64>,
    body_placeholder: Option<u64>,
    owned_drawables: Box<[u64]>,
}

impl CachedSlideGraph {
    fn from_archive(slide: &kn::SlideArchive) -> Self {
        Self {
            title_placeholder: slide
                .title_placeholder
                .as_ref()
                .map(|reference| reference.identifier),
            body_placeholder: slide
                .body_placeholder
                .as_ref()
                .map(|reference| reference.identifier),
            owned_drawables: slide
                .owned_drawables
                .iter()
                .map(|reference| reference.identifier)
                .collect(),
        }
    }

    fn fits_cache_budget(&self) -> bool {
        self.owned_drawables.len() <= MAX_OPERATION_CACHED_DRAWABLES_PER_SLIDE
    }
}

/// Decode state shared by the reads that make up one editor operation.
///
/// The cache owns no mutable package state and is never stored on
/// [`KeynoteEditor`]. Dropping this value at the end of the operation is the
/// invalidation boundary for all entries, including after a staged mutation.
struct KeynoteOperation {
    graph: ObjectGraph,
    slide_cache: HashMap<u64, Rc<CachedSlideGraph>>,
    drawable_storage_cache: HashMap<u64, Option<u64>>,
}

impl KeynoteOperation {
    fn new(package: &IWorkPackage) -> Result<Self> {
        Ok(Self {
            graph: ObjectGraph::read(package)?,
            slide_cache: HashMap::with_capacity(MAX_OPERATION_CACHED_SLIDES),
            drawable_storage_cache: HashMap::with_capacity(MAX_OPERATION_CACHED_DRAWABLE_STORAGES),
        })
    }

    fn decode_slide(&self, identifier: u64) -> Result<kn::SlideArchive> {
        self.graph.decode(identifier, "KN.SlideArchive")
    }

    fn remember_slide(&mut self, identifier: u64, slide: &kn::SlideArchive) {
        if self.slide_cache.len() >= MAX_OPERATION_CACHED_SLIDES
            || slide.owned_drawables.len() > MAX_OPERATION_CACHED_DRAWABLES_PER_SLIDE
        {
            return;
        }
        self.slide_cache
            .insert(identifier, Rc::new(CachedSlideGraph::from_archive(slide)));
    }

    fn slide(&mut self, identifier: u64) -> Result<Rc<CachedSlideGraph>> {
        if let Some(slide) = self.slide_cache.get(&identifier) {
            return Ok(Rc::clone(slide));
        }
        let slide = self.decode_slide(identifier)?;
        let cached = Rc::new(CachedSlideGraph::from_archive(&slide));
        if self.slide_cache.len() < MAX_OPERATION_CACHED_SLIDES && cached.fits_cache_budget() {
            self.slide_cache.insert(identifier, Rc::clone(&cached));
        }
        Ok(cached)
    }

    fn drawable_storage(&mut self, identifier: u64) -> Result<Option<u64>> {
        if let Some(storage) = self.drawable_storage_cache.get(&identifier) {
            return Ok(*storage);
        }
        let storage = self.graph.drawable_storage(identifier)?;
        if self.drawable_storage_cache.len() < MAX_OPERATION_CACHED_DRAWABLE_STORAGES {
            self.drawable_storage_cache.insert(identifier, storage);
        }
        Ok(storage)
    }
}

/// Transactional editor for a Keynote package.
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

    pub(crate) fn from_package(package: IWorkPackage) -> Result<Self> {
        let editor = Self {
            text: IWorkTextEditor::from_package(package),
        };
        editor.slides()?;
        Ok(editor)
    }

    pub fn slides(&self) -> Result<Vec<KeynoteSlideInfo>> {
        let mut operation = KeynoteOperation::new(self.package())?;
        self.slides_with_operation(&mut operation)
    }

    fn slides_with_operation(
        &self,
        operation: &mut KeynoteOperation,
    ) -> Result<Vec<KeynoteSlideInfo>> {
        let document: kn::DocumentArchive = operation.graph.decode(1, "KN.DocumentArchive")?;
        let show: kn::ShowArchive = operation
            .graph
            .decode(document.show.identifier, "KN.ShowArchive")?;

        let mut slides = Vec::with_capacity(show.slide_tree.slides.len());
        let mut layout_catalog = None;
        for (index, node_reference) in show.slide_tree.slides.into_iter().enumerate() {
            let node: kn::SlideNodeArchive = operation
                .graph
                .decode(node_reference.identifier, "KN.SlideNodeArchive")?;
            let slide_reference = node.slide.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide node {} has no slide reference",
                    node_reference.identifier
                ))
            })?;
            let slide = operation.decode_slide(slide_reference.identifier)?;
            operation.remember_slide(slide_reference.identifier, &slide);
            let layout = match slide.template_slide.as_ref() {
                Some(template_slide) => {
                    let catalog = match layout_catalog.as_ref() {
                        Some(catalog) => catalog,
                        None => {
                            let theme = operation
                                .graph
                                .decode(show.theme.identifier, "KN.ThemeArchive")?;
                            layout_catalog.insert(slide_create::layout::LayoutCatalog::read(
                                &operation.graph,
                                &theme,
                            )?)
                        },
                    };
                    Some(catalog.current(template_slide.identifier)?)
                },
                None => None,
            };
            let is_title_visible = slide
                .title_placeholder
                .as_ref()
                .map(|reference| {
                    placeholder_ownership::validate(index, &slide, reference.identifier, "title")
                })
                .transpose()?;
            let is_body_visible = slide
                .body_placeholder
                .as_ref()
                .map(|reference| {
                    placeholder_ownership::validate(index, &slide, reference.identifier, "body")
                })
                .transpose()?;
            let title_storage_id = slide
                .title_placeholder
                .map(|reference| operation.drawable_storage(reference.identifier))
                .transpose()?
                .flatten();
            let title_storage_id = title_storage_id
                .map(crate::text::native_storage_id)
                .transpose()?;
            let body_storage_id = slide
                .body_placeholder
                .map(|reference| operation.drawable_storage(reference.identifier))
                .transpose()?
                .flatten();
            let body_storage_id = body_storage_id
                .map(crate::text::native_storage_id)
                .transpose()?;
            let title = title_storage_id
                .map(|identifier| operation.graph.storage_text(identifier.get()))
                .transpose()?;
            let body = body_storage_id
                .map(|identifier| operation.graph.storage_text(identifier.get()))
                .transpose()?;
            let notes_storage_id = slide
                .note
                .map(|reference| {
                    operation
                        .graph
                        .decode::<kn::NoteArchive>(reference.identifier, "KN.NoteArchive")
                        .map(|note| note.contained_storage.identifier)
                })
                .transpose()?;
            let notes_storage_id = notes_storage_id
                .map(crate::text::native_storage_id)
                .transpose()?;
            let notes = notes_storage_id
                .map(|identifier| operation.graph.storage_text(identifier.get()))
                .transpose()?;
            let transition = if slide.transition.attributes.animation_attributes.is_some() {
                let original = operation.graph.message_data_type(
                    slide_reference.identifier,
                    5,
                    "KN.SlideArchive",
                )?;
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
                layout,
                is_skipped: node.is_skipped,
                is_slide_number_visible: node.is_slide_number_visible,
                is_title_visible,
                is_body_visible,
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
    /// This includes title and body placeholders, ordinary text boxes, and
    /// text-bearing shapes in slide ownership order. Speaker notes are exposed separately on
    /// [`KeynoteSlideInfo`] because they are owned by `KN.NoteArchive` rather
    /// than a drawable.
    pub fn slide_text_storages(&self, slide_index: usize) -> Result<Vec<KeynoteSlideTextInfo>> {
        let mut operation = KeynoteOperation::new(self.package())?;
        let slides = self.slides_with_operation(&mut operation)?;
        slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let mut storage_owners = HashMap::<u64, (usize, u64)>::new();
        let mut result = Vec::new();
        for (owner_slide_index, owner) in slides.iter().enumerate() {
            let slide = operation.slide(owner.slide_id)?;
            let title = slide.title_placeholder;
            let body = slide.body_placeholder;
            let mut seen_drawables = HashSet::with_capacity(slide.owned_drawables.len());
            for &drawable_id in &slide.owned_drawables {
                if !seen_drawables.insert(drawable_id) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {owner_slide_index} repeats owned drawable {drawable_id}"
                    )));
                }
                let Some(storage_id) = operation.drawable_storage(drawable_id)? else {
                    continue;
                };
                let storage_id = crate::text::native_storage_id(storage_id)?;
                if let Some((previous_slide, previous_drawable)) =
                    storage_owners.insert(storage_id.get(), (owner_slide_index, drawable_id))
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
                    let shape: tswp::ShapeInfoArchive = operation.graph.decode_type(
                        drawable_id,
                        SHAPE_INFO_MESSAGE_TYPE,
                        "TSWP.ShapeInfoArchive",
                    )?;
                    match shape.is_text_box {
                        Some(true) => KeynoteSlideTextRole::TextBox,
                        Some(false) => KeynoteSlideTextRole::Shape,
                        None => {
                            return Err(Error::InvalidFormat(format!(
                                "Keynote shape {drawable_id} has no text-box classification"
                            )));
                        },
                    }
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

    /// Replace a UTF-16 range in a slide-owned text box, shape, or placeholder.
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

    /// Replace all text in a slide-owned text box, shape, or placeholder.
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

    /// Clear slide-owned drawable text without deleting its shape or placeholder.
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

    /// Read vertical alignment, edge insets, and autosizing for a text box.
    pub fn slide_text_box_text_layout(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Layout> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        shape_text_layout(self.package(), &graph.archive_name, drawable_object_id)
    }

    /// Replace text-frame layout while preserving text, columns, and drawing style.
    pub fn set_slide_text_box_text_layout(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        layout: Layout,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package().clone(),
            &graph.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_text_box_text_layout(slide_index, drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Keynote text-box layout update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove crate-authored text-frame layout overrides.
    pub fn reset_slide_text_box_text_layout(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_layout(
            self.package().clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the uniform column layout of an ordinary slide text box.
    pub fn slide_text_box_columns(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Columns> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        shape_text_columns(self.package(), &graph.archive_name, drawable_object_id)
    }

    /// Replace the uniform column layout of an ordinary slide text box.
    pub fn set_slide_text_box_columns(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        columns: &Columns,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let staged = set_shape_text_columns(
            self.package().clone(),
            &graph.archive_name,
            drawable_object_id,
            columns,
        )?;
        let verified = Self::from_package(staged)?;
        if &verified.slide_text_box_columns(slide_index, drawable_object_id)? != columns {
            return Err(Error::InvalidFormat(
                "Keynote text-box column update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited column layout after a crate-authored override.
    pub fn reset_slide_text_box_columns(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_columns(
            self.package().clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn slide_text_box_text_style(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_style(graph.storage_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    pub fn set_slide_text_box_text_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        style: TextStyle,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_style(graph.storage_id, style)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_style(slide_index, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Keynote text-box character formatting update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited character formatting while preserving paragraph overrides.
    pub fn reset_slide_text_box_text_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_style(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of an ordinary slide text box.
    pub fn slide_text_box_text_font(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextFont> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_font(graph.storage_id)
    }

    /// Atomically set a typed font identity across an ordinary slide text box.
    pub fn set_slide_text_box_text_font(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        font: TextFont,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_font(graph.storage_id, font)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the inherited font while preserving sibling overrides.
    pub fn reset_slide_text_box_text_font(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_font(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every explicit language boundary in an ordinary slide text box.
    pub fn slide_text_box_text_languages(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<TextLanguageRun>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_languages(graph.storage_id)
    }

    /// Read the effective language at one UTF-16 text boundary.
    pub fn slide_text_box_text_language(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<TextLanguage> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_language(graph.storage_id, position)
    }

    /// Atomically create or update one text-language boundary.
    pub fn set_slide_text_box_text_language(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_language(graph.storage_id, position, language)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Delete one nonzero language boundary so it inherits the preceding run.
    pub fn remove_slide_text_box_text_language_boundary(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.remove_text_language_boundary(graph.storage_id, position)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Restore automatic language selection across an ordinary slide text box.
    pub fn reset_slide_text_box_text_languages(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_languages(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every hyperlink in an ordinary slide text box.
    pub fn slide_text_box_hyperlinks(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHyperlink>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_hyperlinks(graph.storage_id)
    }

    /// Create a hyperlink over a nonempty, unoccupied UTF-16 text range.
    pub fn add_slide_text_box_hyperlink(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.add_text_hyperlink(graph.storage_id, range, target)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Update a text-box hyperlink's range and target without changing its ID.
    pub fn update_slide_text_box_hyperlink(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.update_text_hyperlink(graph.storage_id, id, range, target)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Delete a text-box hyperlink and its owned smart-field object.
    pub fn remove_slide_text_box_hyperlink(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.remove_text_hyperlink(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Read every plain highlight in an ordinary slide text box.
    pub fn slide_text_box_highlights(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHighlight>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_highlights(graph.storage_id)
    }

    /// Create a plain highlight over a nonempty UTF-16 text range.
    pub fn add_slide_text_box_highlight(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.add_text_highlight(graph.storage_id, range)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Move a plain text-box highlight without changing its ID.
    pub fn update_slide_text_box_highlight(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.update_text_highlight(graph.storage_id, id, range)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Delete a plain text-box highlight and its empty annotation graph.
    pub fn remove_slide_text_box_highlight(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.remove_text_highlight(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Read every ranged comment in an ordinary slide text box.
    pub fn slide_text_box_comments(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<TextComment>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_comments(graph.storage_id)
    }

    /// Create a ranged comment in an ordinary slide text box.
    pub fn add_slide_text_box_comment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.add_text_comment(graph.storage_id, range, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Update a text-box comment's range and body without changing its ID.
    pub fn update_slide_text_box_comment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.update_text_comment(graph.storage_id, id, range, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Delete a ranged text-box comment and its owned annotation graph.
    pub fn remove_slide_text_box_comment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextCommentId,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.remove_text_comment(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Read every direct reply to a slide text-box comment in stored order.
    pub fn slide_text_box_comment_replies(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_comment_replies(graph.storage_id, comment_id)
    }

    /// Append a direct reply to a slide text-box comment.
    pub fn add_slide_text_box_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply = staged.add_text_comment_reply(graph.storage_id, comment_id, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Update a direct slide text-box comment reply without changing its ID.
    pub fn update_slide_text_box_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply =
            staged.update_text_comment_reply(graph.storage_id, comment_id, reply_id, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Delete one direct slide text-box comment reply and its storage.
    pub fn remove_slide_text_box_comment_reply(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply = staged.remove_text_comment_reply(graph.storage_id, comment_id, reply_id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Read the canonical list preset of an ordinary slide text box.
    pub fn slide_text_box_paragraph_list(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphList> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_list(graph.storage_id)
    }

    /// Atomically apply a canonical list preset to an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_list(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        list: ParagraphList,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list(graph.storage_id, list)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Remove list formatting from an ordinary slide text box.
    pub fn reset_slide_text_box_paragraph_list(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every list-level boundary in an ordinary slide text box.
    pub fn slide_text_box_paragraph_list_levels(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_list_levels(graph.storage_id)
    }

    /// Read one paragraph's effective list nesting level.
    pub fn slide_text_box_paragraph_list_level(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_list_level(graph.storage_id, paragraph)
    }

    /// Atomically set one paragraph's list nesting level.
    pub fn set_slide_text_box_paragraph_list_level(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_level(graph.storage_id, paragraph, level)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_slide_text_box_paragraph_list_level(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_level(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one slide text-box paragraph continues or restarts list numbering.
    pub fn slide_text_box_paragraph_list_numbering(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_numbering(graph.storage_id, paragraph)
    }

    /// Continue or restart numbered-list sequencing at one slide text-box paragraph.
    pub fn set_slide_text_box_paragraph_list_numbering(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_numbering(graph.storage_id, paragraph, numbering)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Read one numbered slide text-box paragraph's effective label format.
    pub fn slide_text_box_paragraph_list_number_format(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberFormat> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_number_format(graph.storage_id, paragraph)
    }

    /// Set one numbered slide text-box paragraph's locale-aware label format.
    pub fn set_slide_text_box_paragraph_list_number_format(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        format: ParagraphListNumberFormat,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_format(graph.storage_id, paragraph, format)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_slide_text_box_paragraph_list_number_format(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_format(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one numbered slide text-box paragraph displays hierarchical numbering.
    pub fn slide_text_box_paragraph_list_number_tiering(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberTiering> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_number_tiering(graph.storage_id, paragraph)
    }

    /// Choose flat or hierarchical numbering for one slide text-box list level.
    pub fn set_slide_text_box_paragraph_list_number_tiering(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        tiering: ParagraphListNumberTiering,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_tiering(graph.storage_id, paragraph, tiering)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore flat numbering for one slide text-box list level.
    pub fn reset_slide_text_box_paragraph_list_number_tiering(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_tiering(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one numbered slide text-box paragraph's number-label size.
    pub fn slide_text_box_paragraph_list_number_scale(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberScale> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_number_scale(graph.storage_id, paragraph)
    }

    /// Set one numbered slide text-box paragraph's number-label size.
    pub fn set_slide_text_box_paragraph_list_number_scale(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        scale: ParagraphListNumberScale,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_scale(graph.storage_id, paragraph, scale)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_slide_text_box_paragraph_list_number_scale(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_scale(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one slide text-box paragraph's effective text-bullet marker.
    pub fn slide_text_box_paragraph_list_bullet(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBullet> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_list_bullet(graph.storage_id, paragraph)
    }

    /// Set one slide text-box paragraph's text-bullet marker.
    pub fn set_slide_text_box_paragraph_list_bullet(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        bullet: &ParagraphListBullet,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_bullet(graph.storage_id, paragraph, bullet)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one slide text-box paragraph.
    pub fn reset_slide_text_box_paragraph_list_bullet(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_bullet(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one slide text-box paragraph's effective bullet size and baseline.
    pub fn slide_text_box_paragraph_list_bullet_geometry(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBulletGeometry> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_bullet_geometry(graph.storage_id, paragraph)
    }

    /// Set one slide text-box paragraph's bullet size and baseline.
    pub fn set_slide_text_box_paragraph_list_bullet_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        geometry: ParagraphListBulletGeometry,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_bullet_geometry(graph.storage_id, paragraph, geometry)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_slide_text_box_paragraph_list_bullet_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_bullet_geometry(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one slide text-box list paragraph's label and text-gap indentation.
    pub fn slide_text_box_paragraph_list_indentation(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListIndentation> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_indentation(graph.storage_id, paragraph)
    }

    /// Set one slide text-box list paragraph's label and text-gap indentation.
    pub fn set_slide_text_box_paragraph_list_indentation(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        indentation: ParagraphListIndentation,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_indentation(graph.storage_id, paragraph, indentation)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_slide_text_box_paragraph_list_indentation(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_indentation(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one slide text-box list paragraph's effective label color.
    pub fn slide_text_box_paragraph_list_label_color(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLabelColor> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text
            .paragraph_list_label_color(graph.storage_id, paragraph)
    }

    /// Set one slide text-box list paragraph's bullet or number color.
    pub fn set_slide_text_box_paragraph_list_label_color(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        color: ParagraphListLabelColor,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_label_color(graph.storage_id, paragraph, color)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_slide_text_box_paragraph_list_label_color(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_label_color(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform underline and strikethrough formatting.
    pub fn slide_text_box_text_decorations(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextDecorations> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_decorations(graph.storage_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    pub fn set_slide_text_box_text_decorations(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        decorations: TextDecorations,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_decorations(graph.storage_id, decorations)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_decorations(slide_index, drawable_object_id)? != decorations
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box decoration update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited decorations while preserving sibling overrides.
    pub fn reset_slide_text_box_text_decorations(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_decorations(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective uniform text color of an ordinary slide text box.
    pub fn slide_text_box_text_color(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RgbaColor> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_color(graph.storage_id)
    }

    /// Atomically set one text color across an ordinary slide text box.
    pub fn set_slide_text_box_text_color(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        color: RgbaColor,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_color(graph.storage_id, color)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_color(slide_index, drawable_object_id)? != color {
            return Err(Error::InvalidFormat(
                "Keynote text-box color update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited text color while preserving sibling overrides.
    pub fn reset_slide_text_box_text_color(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_color(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform capitalization from an ordinary slide text box.
    pub fn slide_text_box_text_capitalization(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextCapitalization> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_capitalization(graph.storage_id)
    }

    /// Atomically set one capitalization mode across an ordinary slide text box.
    pub fn set_slide_text_box_text_capitalization(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        capitalization: TextCapitalization,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_capitalization(graph.storage_id, capitalization)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_capitalization(slide_index, drawable_object_id)?
            != capitalization
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box capitalization update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited capitalization while preserving sibling overrides.
    pub fn reset_slide_text_box_text_capitalization(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_capitalization(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform baseline script from an ordinary slide text box.
    pub fn slide_text_box_text_script(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextScript> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_script(graph.storage_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    pub fn set_slide_text_box_text_script(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        script: TextScript,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_script(graph.storage_id, script)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_script(slide_index, drawable_object_id)? != script {
            return Err(Error::InvalidFormat(
                "Keynote text-box script update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited baseline script while preserving sibling overrides.
    pub fn reset_slide_text_box_text_script(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_script(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of an ordinary slide text box.
    pub fn slide_text_box_text_baseline_shift(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextBaselineShift> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_baseline_shift(graph.storage_id)
    }

    /// Atomically set a signed custom baseline displacement.
    pub fn set_slide_text_box_text_baseline_shift(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shift: TextBaselineShift,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_baseline_shift(graph.storage_id, shift)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_baseline_shift(slide_index, drawable_object_id)? != shift {
            return Err(Error::InvalidFormat(
                "Keynote text-box baseline-shift update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited baseline displacement while preserving sibling overrides.
    pub fn reset_slide_text_box_text_baseline_shift(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_baseline_shift(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of an ordinary slide text box.
    pub fn slide_text_box_text_character_spacing(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextCharacterSpacing> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_character_spacing(graph.storage_id)
    }

    /// Atomically set character spacing across an ordinary slide text box.
    pub fn set_slide_text_box_text_character_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_character_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_character_spacing(slide_index, drawable_object_id)?
            != spacing
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box character-spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited character spacing while preserving sibling overrides.
    pub fn reset_slide_text_box_text_character_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_character_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of an ordinary slide text box.
    pub fn slide_text_box_text_ligatures(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextLigatures> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_ligatures(graph.storage_id)
    }

    /// Atomically set the ligature policy across an ordinary slide text box.
    pub fn set_slide_text_box_text_ligatures(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        ligatures: TextLigatures,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_ligatures(graph.storage_id, ligatures)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_ligatures(slide_index, drawable_object_id)? != ligatures {
            return Err(Error::InvalidFormat(
                "Keynote text-box ligature update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited ligatures while preserving sibling overrides.
    pub fn reset_slide_text_box_text_ligatures(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_ligatures(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective outline of an ordinary slide text box.
    pub fn slide_text_box_text_outline(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Outline> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_outline(graph.storage_id)
    }

    /// Atomically set a typed outline across an ordinary slide text box.
    pub fn set_slide_text_box_text_outline(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        outline: Outline,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_outline(graph.storage_id, outline)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_outline(slide_index, drawable_object_id)? != outline {
            return Err(Error::InvalidFormat(
                "Keynote text-box outline update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited outline while preserving sibling overrides.
    pub fn reset_slide_text_box_text_outline(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_outline(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective shadow of an ordinary slide text box.
    pub fn slide_text_box_text_shadow(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Shadow> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_shadow(graph.storage_id)
    }

    /// Atomically set a typed drop shadow across an ordinary slide text box.
    pub fn set_slide_text_box_text_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shadow: Shadow,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_shadow(graph.storage_id, shadow)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_shadow(slide_index, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Keynote text-box shadow update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited shadow while preserving sibling overrides.
    pub fn reset_slide_text_box_text_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_shadow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective solid background of an ordinary slide text box.
    pub fn slide_text_box_text_background(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TextAppearanceBackground> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_background(graph.storage_id)
    }

    /// Atomically set a solid background across an ordinary slide text box.
    pub fn set_slide_text_box_text_background(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        background: TextAppearanceBackground,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_background(graph.storage_id, background)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_text_background(slide_index, drawable_object_id)? != background {
            return Err(Error::InvalidFormat(
                "Keynote text-box background update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited text background while preserving sibling overrides.
    pub fn reset_slide_text_box_text_background(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph background.
    pub fn slide_text_box_paragraph_background(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphBackground> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_background(graph.storage_id)
    }

    /// Atomically set the paragraph background across a slide text box.
    pub fn set_slide_text_box_paragraph_background(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        background: ParagraphBackground,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_background(graph.storage_id, background)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_background(slide_index, drawable_object_id)?
            != background
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph background update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph background.
    pub fn reset_slide_text_box_paragraph_background(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph borders.
    pub fn slide_text_box_paragraph_borders(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Borders> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_borders(graph.storage_id)
    }

    /// Atomically set paragraph borders across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_borders(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        borders: Borders,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_borders(graph.storage_id, borders)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_borders(slide_index, drawable_object_id)? != borders {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph border update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph borders.
    pub fn reset_slide_text_box_paragraph_borders(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_borders(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph pagination and hyphenation controls.
    pub fn slide_text_box_paragraph_flow(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphFlow> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_flow(graph.storage_id)
    }

    /// Atomically set paragraph pagination and hyphenation controls.
    pub fn set_slide_text_box_paragraph_flow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        flow: ParagraphFlow,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_flow(graph.storage_id, flow)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_flow(slide_index, drawable_object_id)? != flow {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph flow update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph pagination and hyphenation controls.
    pub fn reset_slide_text_box_paragraph_flow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_flow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective base-writing direction of an ordinary slide text box.
    pub fn slide_text_box_paragraph_writing_direction(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphWritingDirection> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_writing_direction(graph.storage_id)
    }

    /// Set one base-writing direction across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_writing_direction(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        direction: ParagraphWritingDirection,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_writing_direction(graph.storage_id, direction)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_writing_direction(slide_index, drawable_object_id)?
            != direction
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph writing-direction update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited base-writing direction.
    pub fn reset_slide_text_box_paragraph_writing_direction(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_writing_direction(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph alignment of an ordinary slide text box.
    pub fn slide_text_box_paragraph_alignment(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Alignment> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_alignment(graph.storage_id)
    }

    /// Set one paragraph alignment across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_alignment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        alignment: Alignment,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_alignment(graph.storage_id, alignment)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_alignment(slide_index, drawable_object_id)?
            != alignment
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph-alignment update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited paragraph alignment after a private minimal override.
    pub fn reset_slide_text_box_paragraph_alignment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_alignment(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective line spacing of an ordinary slide text box.
    pub fn slide_text_box_paragraph_line_spacing(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<LineSpacing> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_line_spacing(graph.storage_id)
    }

    /// Set one typed line-spacing mode across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_line_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        spacing: LineSpacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_line_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_line_spacing(slide_index, drawable_object_id)?
            != spacing
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph line-spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited line spacing while preserving sibling paragraph overrides.
    pub fn reset_slide_text_box_paragraph_line_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_line_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing of an ordinary slide text box.
    pub fn slide_text_box_paragraph_spacing(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Spacing> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_spacing(graph.storage_id)
    }

    /// Atomically set before/after paragraph spacing across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        spacing: Spacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_spacing(slide_index, drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited paragraph spacing while preserving sibling overrides.
    pub fn reset_slide_text_box_paragraph_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right indentation of a slide text box.
    pub fn slide_text_box_paragraph_indents(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Indents> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_indents(graph.storage_id)
    }

    /// Atomically set paragraph indentation across an ordinary slide text box.
    pub fn set_slide_text_box_paragraph_indents(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        indents: Indents,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_indents(graph.storage_id, indents)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_indents(slide_index, drawable_object_id)? != indents {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph indentation update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited indentation while preserving sibling paragraph overrides.
    pub fn reset_slide_text_box_paragraph_indents(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_indents(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the decimal-tab alignment character of a slide text box.
    pub fn slide_text_box_paragraph_decimal_tab_character(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphDecimalTabCharacter> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_decimal_tab_character(graph.storage_id)
    }

    /// Atomically set the decimal-tab alignment character of a slide text box.
    pub fn set_slide_text_box_paragraph_decimal_tab_character(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        character: ParagraphDecimalTabCharacter,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_decimal_tab_character(graph.storage_id, character)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified
            .slide_text_box_paragraph_decimal_tab_character(slide_index, drawable_object_id)?
            != character
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box decimal-tab character update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited decimal-tab alignment character.
    pub fn reset_slide_text_box_paragraph_decimal_tab_character(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_decimal_tab_character(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the distance between implicit tab stops in a slide text box.
    pub fn slide_text_box_paragraph_default_tab_interval(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphDefaultTabInterval> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_default_tab_interval(graph.storage_id)
    }

    /// Atomically set the distance between implicit tab stops in a slide text box.
    pub fn set_slide_text_box_paragraph_default_tab_interval(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        interval: ParagraphDefaultTabInterval,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_default_tab_interval(graph.storage_id, interval)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified
            .slide_text_box_paragraph_default_tab_interval(slide_index, drawable_object_id)?
            != interval
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box default-tab interval update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited default-tab interval.
    pub fn reset_slide_text_box_paragraph_default_tab_interval(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_default_tab_interval(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of a slide text box.
    pub fn slide_text_box_paragraph_tab_stops(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ParagraphTabStops> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_tab_stops(graph.storage_id)
    }

    /// Atomically replace every explicit ruler tab stop of a slide text box.
    pub fn set_slide_text_box_paragraph_tab_stops(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        stops: ParagraphTabStops,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_tab_stops(graph.storage_id, stops)?;
        let expected = staged.paragraph_tab_stops(graph.storage_id)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_tab_stops(slide_index, drawable_object_id)? != expected
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box paragraph tab-stop update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited tab stops while preserving sibling paragraph overrides.
    pub fn reset_slide_text_box_paragraph_tab_stops(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_tab_stops(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// List every Drop Cap in a slide-owned text box.
    pub fn slide_text_box_paragraph_drop_caps(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<Placement>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_drop_caps(graph.storage_id)
    }

    /// Read the Drop Cap attached to one text-box paragraph.
    pub fn slide_text_box_paragraph_drop_cap(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<Option<DropCap>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.paragraph_drop_cap(graph.storage_id, paragraph)
    }

    /// Atomically create or replace a text-box Drop Cap.
    pub fn set_slide_text_box_paragraph_drop_cap(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
        drop_cap: DropCap,
    ) -> Result<()> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_drop_cap(graph.storage_id, paragraph, drop_cap)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.slide_text_box_paragraph_drop_cap(slide_index, drawable_object_id, paragraph)?
            != Some(drop_cap)
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box Drop Cap update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Atomically remove a text-box Drop Cap.
    pub fn remove_slide_text_box_paragraph_drop_cap(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.remove_paragraph_drop_cap(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
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
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = remap[&source.drawable_id];
        let new_storage_id = crate::text::native_storage_id(remap[&source.storage_id.get()])?;
        offset_keynote_drawable_clone(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            DRAWABLE_DUPLICATE_OFFSET,
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
            || created.storage.id != new_storage_id
            || created.storage.storage.text() != text
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
        comments.clear_comment(DrawableId::from_raw(drawable_object_id)?)?;
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
    pub fn slide_drawables(&self, slide_index: usize) -> Result<Vec<DrawableInfo>> {
        let owned = self.slide_owned_drawable_ids(slide_index)?;
        let mut drawables = IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .drawables()?
            .into_iter()
            .filter(|drawable| owned.contains(&drawable.id.get()))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.id.get());
        Ok(drawables)
    }

    /// Read a comment attached directly to a drawable owned by one slide.
    pub fn slide_drawable_comment(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<DrawableComment>> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .comment(DrawableId::from_raw(drawable_object_id)?)
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
        comments.set_comment(DrawableId::from_raw(drawable_object_id)?, text)?;
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
        comments.clear_comment(DrawableId::from_raw(drawable_object_id)?)?;
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
    ) -> Result<Vec<DrawableReply>> {
        self.require_slide_drawable(slide_index, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .replies(DrawableId::from_raw(drawable_object_id)?)
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
        let reply_id = comments.add_reply(DrawableId::from_raw(drawable_object_id)?, text)?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(reply_id.get())
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
        let reply_id = comments.set_reply(
            DrawableId::from_raw(drawable_object_id)?,
            StorageId::from_raw(reply_storage_object_id)?,
            text,
        )?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(reply_id.get())
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
        comments.remove_reply(
            DrawableId::from_raw(drawable_object_id)?,
            StorageId::from_raw(reply_storage_object_id)?,
        )?;
        let staged = comments.into_package();
        Self::from_package(staged.clone())?;
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
        if typed_action_acceleration(&settings) == Some(BuildAcceleration::Custom)
            && settings.timing_curve.is_none()
        {
            return Err(Error::ParseError(
                "Creating a custom-curve Keynote action requires a timing curve".to_owned(),
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
        if typed_action_acceleration(&settings) == Some(BuildAcceleration::Custom)
            && typed_action_acceleration(&build.settings) != Some(BuildAcceleration::Custom)
            && settings.timing_curve.is_none()
        {
            return Err(Error::ParseError(
                "Changing a Keynote action to a custom curve requires a timing curve".to_owned(),
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
        let mut next_identifier = next_object_identifier(self.package())?;
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

        if let Some(source_component) =
            component_identifier_for_entry(&staged, &source_archive_name)?
        {
            clone_component_registration(
                &mut staged,
                source_component,
                new_slide_id,
                &format!("Slide-{new_slide_id}"),
                &remap,
            )?;
            if let Some(document_component) =
                component_identifier_for_entry(&staged, &node_archive_name)?
            {
                add_component_object_uuids(&mut staged, document_component, &[new_node_id])?;
                add_component_link(&mut staged, document_component, new_slide_id)?;
            }
            set_package_last_object_identifier(&mut staged, next_identifier - 1)?;
        }

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

    pub(crate) fn package(&self) -> &IWorkPackage {
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
        let data_identifier = MediaAssetId::try_from(data_identifier)?;
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
        let data_identifier = MediaAssetId::try_from(data_identifier)?;
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

    // Used by crate-internal mutation fixtures; ordinary callers publish via `to_bytes`/`save`.
    #[allow(dead_code)]
    pub(crate) fn into_package(self) -> IWorkPackage {
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
    ) -> Result<TextStorageId> {
        self.slide_text_storages(slide_index)?
            .into_iter()
            .find(|text| text.drawable_object_id == drawable_object_id)
            .map(|text| text.storage.id)
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
            || storage_id != text.storage.id.get()
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
            storage_id: crate::text::native_storage_id(storage_id)?,
            object_ids,
            uuid_object_ids,
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
            .any(|drawable| drawable.id.get() == drawable_object_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide drawable {drawable_object_id} has no supported direct drawable payload"
            )));
        }
        Ok(())
    }
}

mod builds;
mod date_time_fields;
mod drawable_order;
mod named_paragraph_styles;
mod placeholder_ownership;
mod slide_audio;
mod slide_background;
mod slide_background_color;
mod slide_background_gradient_wire;
mod slide_background_reset;
mod slide_background_wire;
mod slide_charts;
mod slide_create;
mod slide_delete;
mod slide_graph;
mod slide_images;
mod slide_layout_media;
mod slide_layout_update;
mod slide_movies;
mod slide_number;
mod slide_preview;
mod slide_shapes;
mod slide_style_graph;
mod slide_style_metadata;
mod slide_style_registry;
mod slide_tables;
mod soundtrack;
mod soundtrack_items;
mod soundtrack_wire;
mod text_box_create;
mod transition;
mod transition_wire;

use builds::*;
pub use litchi_iwa_common::color::{RgbColorSpace, Rgba};
pub use litchi_keynote::Seconds;
pub use litchi_keynote::background::{Angle, Background, Gradient, Kind, Opaque, Stop};
pub use litchi_keynote::show::{Mode, Settings, Size};
pub use litchi_keynote::slide::media::MovieKind;
pub use litchi_keynote::transition::Effect;
pub use litchi_keynote::transition::{
    Acceleration, AccelerationKind, AnimationParameters, CustomParameters, Direction, MosaicType,
    TextDelivery, TextDeliveryKind,
};
pub use slide_audio::{KeynoteSlideAudioInfo, RemovedKeynoteSlideAudio};
pub use slide_charts::{KeynoteSlideChartInfo, RemovedKeynoteSlideChart};
use slide_graph::*;
pub use slide_images::{KeynoteSlideImageInfo, KeynoteSlideImageKind, RemovedKeynoteSlideImage};
pub use slide_movies::{KeynoteSlideMovieInfo, RemovedKeynoteSlideMovie};
pub use slide_shapes::{KeynoteSlideShapeInfo, RemovedKeynoteSlideShape};
pub use slide_tables::{
    KeynoteSlideTable, KeynoteSlideTableInfo, KeynoteTableCellConditionalHighlightInfo,
    KeynoteTableCellInset, KeynoteTableCellInsets, KeynoteTableCellLayout,
    KeynoteTableCellParagraphList, KeynoteTableCellParagraphListBullet,
    KeynoteTableCellParagraphListBulletGeometry, KeynoteTableCellParagraphListIndentation,
    KeynoteTableCellParagraphListLabelColor, KeynoteTableCellParagraphListLevel,
    KeynoteTableCellParagraphListLevelPlacement, KeynoteTableCellParagraphListNumberFormat,
    KeynoteTableCellParagraphListNumberScale, KeynoteTableCellParagraphListNumberTiering,
    KeynoteTableCellParagraphListNumbering, KeynoteTableCellParagraphListPlacement,
    KeynoteTableCellParagraphTabStops, KeynoteTableCellTextBackground,
    KeynoteTableCellTextBaselineShift, KeynoteTableCellTextCapitalization,
    KeynoteTableCellTextCharacterSpacing, KeynoteTableCellTextColor,
    KeynoteTableCellTextDecorations, KeynoteTableCellTextFont, KeynoteTableCellTextLigatures,
    KeynoteTableCellTextOutline, KeynoteTableCellTextScript, KeynoteTableCellTextShadow,
    KeynoteTableCellTextStyle, KeynoteTableCellTextWrap, KeynoteTableCellUpdate,
    KeynoteTableCellValue, KeynoteTableCellVerticalAlignment, KeynoteTableDimension,
    KeynoteTableDimensionSize, KeynoteTablePoints, KeynoteTableTitleSettings,
    RemovedKeynoteSlideTable,
};
pub use soundtrack_items::KeynoteSoundtrackItemInfo;
use transition_wire::{transition_settings_from_wire, validate_transition_wire};
#[cfg(test)]
mod operation_cache_tests;
#[cfg(test)]
mod tests;
