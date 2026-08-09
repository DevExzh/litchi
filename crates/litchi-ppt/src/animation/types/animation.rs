//! High-level animation information and `PowerPoint` 97 animation atoms.

use super::build::BuildInfo;
use super::effects::{AfterEffect, TimeNodeContainer};
use crate::animation::sound::AnimationSound;
use crate::animation::triggers::{InteractiveTrigger, IterationType, RepeatBehavior};
use crate::records::Record;

/// Animation information for a slide or shape.
#[derive(Debug, Clone)]
pub struct AnimationInfo {
    /// `PowerPoint` 97 shape animation atom, when present.
    pub legacy_atom: Option<LegacyAnimationAtom>,
    /// Build list (order of appearance animations)
    pub build_list: Option<BuildInfo>,
    /// Time node containers for advanced animations
    pub time_nodes: Vec<TimeNodeContainer>,
    /// Sound associated with animation
    pub sound: Option<AnimationSound>,
    /// Interactive trigger
    pub trigger: Option<InteractiveTrigger>,
    /// Iteration type (for text animations)
    pub iteration: IterationType,
    /// Repeat behavior
    pub repeat: RepeatBehavior,
    /// After-effect color (for dim effects)
    pub after_effect_color: Option<u32>,
    /// Raw animation records for advanced parsing
    pub raw_records: Vec<Record>,
}

impl Default for AnimationInfo {
    fn default() -> Self {
        Self::new()
    }
}

impl AnimationInfo {
    /// Create a new empty animation info.
    #[must_use]
    pub fn new() -> Self {
        Self {
            legacy_atom: None,
            build_list: None,
            time_nodes: Vec::new(),
            sound: None,
            trigger: None,
            iteration: IterationType::default(),
            repeat: RepeatBehavior::default(),
            after_effect_color: None,
            raw_records: Vec::new(),
        }
    }

    /// Check if this slide has any animations.
    #[must_use]
    pub fn has_animations(&self) -> bool {
        self.legacy_atom
            .as_ref()
            .is_some_and(|atom| atom.build_type != LegacyAnimationBuild::NoBuild)
            || self.build_list.is_some()
            || !self.time_nodes.is_empty()
    }

    /// Get the number of animated objects.
    #[must_use]
    pub fn animation_count(&self) -> usize {
        let legacy_count = usize::from(
            self.legacy_atom
                .as_ref()
                .is_some_and(|atom| atom.build_type != LegacyAnimationBuild::NoBuild),
        );
        let build_count = self.build_list.as_ref().map_or(0, |b| b.builds.len());
        legacy_count + build_count + self.time_nodes.len()
    }
}

/// Animation metadata associated with a legacy `PowerPoint` shape.
#[derive(Debug, Clone)]
pub struct ShapeAnimation {
    /// `OfficeArt` shape identifier.
    pub shape_id: u32,
    /// Parsed, inert animation metadata.
    pub animation: AnimationInfo,
}

/// `PowerPoint` 97 paragraph/chart build behavior stored in `AnimationInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyAnimationBuild {
    FollowMaster,
    #[default]
    NoBuild,
    OneBuild,
    Level1,
    Level2,
    Level3,
    Level4,
    Level5,
    GraphBySeries,
    GraphByCategory,
    GraphByElementInSeries,
    GraphByElementInCategory,
}

impl LegacyAnimationBuild {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0xFE => Some(Self::FollowMaster),
            0x00 => Some(Self::NoBuild),
            0x01 => Some(Self::OneBuild),
            0x02 => Some(Self::Level1),
            0x03 => Some(Self::Level2),
            0x04 => Some(Self::Level3),
            0x05 => Some(Self::Level4),
            0x06 => Some(Self::Level5),
            0x07 => Some(Self::GraphBySeries),
            0x08 => Some(Self::GraphByCategory),
            0x09 => Some(Self::GraphByElementInSeries),
            0x0A => Some(Self::GraphByElementInCategory),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::FollowMaster => 0xFE,
            Self::NoBuild => 0x00,
            Self::OneBuild => 0x01,
            Self::Level1 => 0x02,
            Self::Level2 => 0x03,
            Self::Level3 => 0x04,
            Self::Level4 => 0x05,
            Self::Level5 => 0x06,
            Self::GraphBySeries => 0x07,
            Self::GraphByCategory => 0x08,
            Self::GraphByElementInSeries => 0x09,
            Self::GraphByElementInCategory => 0x0A,
        }
    }
}

/// `PowerPoint` 97 animation effect code.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyAnimationEffect {
    #[default]
    Cut,
    Random,
    Blinds,
    Checker,
    Cover,
    Dissolve,
    Fade,
    Pull,
    RandomBars,
    Strips,
    Wipe,
    Zoom,
    Fly,
    Split,
    Flash,
    Diamond,
    Plus,
    Wedge,
    Wheel,
    Circle,
}

impl LegacyAnimationEffect {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0x00 => Some(Self::Cut),
            0x01 => Some(Self::Random),
            0x02 => Some(Self::Blinds),
            0x03 => Some(Self::Checker),
            0x04 => Some(Self::Cover),
            0x05 => Some(Self::Dissolve),
            0x06 => Some(Self::Fade),
            0x07 => Some(Self::Pull),
            0x08 => Some(Self::RandomBars),
            0x09 => Some(Self::Strips),
            0x0A => Some(Self::Wipe),
            0x0B => Some(Self::Zoom),
            0x0C => Some(Self::Fly),
            0x0D => Some(Self::Split),
            0x0E => Some(Self::Flash),
            0x11 => Some(Self::Diamond),
            0x12 => Some(Self::Plus),
            0x13 => Some(Self::Wedge),
            0x1A => Some(Self::Wheel),
            0x1B => Some(Self::Circle),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::Cut => 0x00,
            Self::Random => 0x01,
            Self::Blinds => 0x02,
            Self::Checker => 0x03,
            Self::Cover => 0x04,
            Self::Dissolve => 0x05,
            Self::Fade => 0x06,
            Self::Pull => 0x07,
            Self::RandomBars => 0x08,
            Self::Strips => 0x09,
            Self::Wipe => 0x0A,
            Self::Zoom => 0x0B,
            Self::Fly => 0x0C,
            Self::Split => 0x0D,
            Self::Flash => 0x0E,
            Self::Diamond => 0x11,
            Self::Plus => 0x12,
            Self::Wedge => 0x13,
            Self::Wheel => 0x1A,
            Self::Circle => 0x1B,
        }
    }

    pub(crate) const fn accepts_direction(self, direction: u8) -> bool {
        match self {
            Self::Cut | Self::Flash => direction <= 2,
            Self::Random => true,
            Self::Blinds | Self::Checker | Self::RandomBars | Self::Zoom => direction <= 1,
            Self::Cover | Self::Pull => direction <= 7,
            Self::Dissolve
            | Self::Fade
            | Self::Diamond
            | Self::Plus
            | Self::Wedge
            | Self::Circle => direction == 0,
            Self::Strips => direction >= 4 && direction <= 7,
            Self::Wipe | Self::Split => direction <= 3,
            Self::Fly => direction <= 0x1C,
            Self::Wheel => matches!(direction, 1 | 2 | 3 | 4 | 8),
        }
    }
}

/// Text subdivision behavior in a `PowerPoint` 97 animation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyTextBuildSubEffect {
    #[default]
    AllAtOnce,
    ByWord,
    ByCharacter,
}

impl LegacyTextBuildSubEffect {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::AllAtOnce),
            1 => Some(Self::ByWord),
            2 => Some(Self::ByCharacter),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::AllAtOnce => 0,
            Self::ByWord => 1,
            Self::ByCharacter => 2,
        }
    }
}

/// Exact 28-byte payload of an `[MS-PPT]` `AnimationInfoAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the bool fields mirror the independent flag bits of the fixed MS-PPT `AnimationInfoAtom` layout, so they cannot be merged into enums without losing the bit-level mapping"
)]
pub struct LegacyAnimationAtom {
    pub dim_color: u32,
    pub reverse: bool,
    pub automatic: bool,
    pub has_sound: bool,
    pub stop_sound: bool,
    pub play: bool,
    pub synchronous: bool,
    pub hide_while_not_playing: bool,
    pub animate_background: bool,
    pub sound_id_ref: u32,
    pub delay_time_ms: i32,
    pub order_id: i16,
    pub slide_count: u16,
    pub build_type: LegacyAnimationBuild,
    pub effect: LegacyAnimationEffect,
    pub effect_direction: u8,
    pub after_effect: AfterEffect,
    pub text_build_sub_effect: LegacyTextBuildSubEffect,
    pub ole_verb: u8,
}

impl Default for LegacyAnimationAtom {
    fn default() -> Self {
        Self {
            dim_color: 0,
            reverse: false,
            automatic: false,
            has_sound: false,
            stop_sound: false,
            play: false,
            synchronous: false,
            hide_while_not_playing: false,
            animate_background: false,
            sound_id_ref: 0,
            delay_time_ms: 0,
            order_id: 0,
            slide_count: 1,
            build_type: LegacyAnimationBuild::NoBuild,
            effect: LegacyAnimationEffect::Cut,
            effect_direction: 0,
            after_effect: AfterEffect::None,
            text_build_sub_effect: LegacyTextBuildSubEffect::AllAtOnce,
            ole_verb: 0,
        }
    }
}
