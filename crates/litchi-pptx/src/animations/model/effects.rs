//! Public animation effects, triggers, and sequence context.

use super::{Duration, Fill, MotionFraction, Repeat, Restart, Speed, SyncBehavior, TimeFilter};

/// `EffectInstance` effect type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Effect {
    /// Appear effect
    Appear,
    /// Fade effect
    Fade,
    /// Fly in effect
    FlyIn,
    /// Float in effect
    FloatIn,
    /// Split effect
    Split,
    /// Wipe effect
    Wipe,
    /// Zoom effect
    Zoom,
    /// Bounce effect
    Bounce,
    /// Spin effect
    Spin,
    /// Grow/Shrink effect
    GrowShrink,
    /// Custom/Unknown effect
    Custom(String),
}

impl Effect {
    /// Parse from preset ID.
    #[must_use]
    pub fn from_preset_id(id: u32) -> Self {
        match id {
            1 => Effect::Appear,
            2 => Effect::FlyIn,
            6 => Effect::GrowShrink,
            8 => Effect::Spin,
            10 => Effect::Fade,
            16 => Effect::Split,
            22 => Effect::Wipe,
            23 => Effect::Zoom,
            24 => Effect::Bounce,
            42 => Effect::FloatIn,
            _ => Effect::Custom(format!("preset_{id}")),
        }
    }

    /// Parse from preset class string (for backwards compatibility).
    #[must_use]
    pub fn from_preset(preset: &str) -> Self {
        match preset.to_lowercase().as_str() {
            "entr" | "appear" => Effect::Appear,
            "fade" => Effect::Fade,
            "fly" | "flyin" => Effect::FlyIn,
            "float" | "floatin" => Effect::FloatIn,
            "split" => Effect::Split,
            "wipe" => Effect::Wipe,
            "zoom" => Effect::Zoom,
            "bounce" => Effect::Bounce,
            "spin" => Effect::Spin,
            "grow" | "growshrink" => Effect::GrowShrink,
            other => Effect::Custom(other.to_string()),
        }
    }

    /// Get the preset ID for this effect.
    /// These are defined in ECMA-376 Part 1.
    #[must_use]
    pub fn preset_id(&self) -> u32 {
        match self {
            Effect::Appear => 1,
            Effect::FlyIn => 2,
            Effect::FloatIn => 42,
            Effect::Split => 16,
            Effect::Fade => 10,
            Effect::Wipe => 22,
            Effect::Zoom => 23,
            Effect::Bounce => 24,
            Effect::Spin => 8,       // Spin is emphasis, but using ID 8
            Effect::GrowShrink => 6, // GrowShrink is emphasis
            Effect::Custom(value) => value
                .split_once(':')
                .and_then(|(_, id)| id.parse().ok())
                .unwrap_or(1),
        }
    }

    /// Get the preset class for this effect.
    /// Valid values: "entr" (entrance), "exit", "emph" (emphasis), "path", "verb", "mediacall"
    #[must_use]
    pub fn preset_class(&self) -> &str {
        match self {
            // Entrance effects
            Effect::Appear => "entr",
            Effect::FlyIn => "entr",
            Effect::FloatIn => "entr",
            Effect::Split => "entr",
            Effect::Fade => "entr",
            Effect::Wipe => "entr",
            Effect::Zoom => "entr",
            Effect::Bounce => "entr",
            // Emphasis effects
            Effect::Spin => "emph",
            Effect::GrowShrink => "emph",
            // Default to entrance
            Effect::Custom(value) => value
                .split_once(':')
                .map(|(class, _)| class)
                .filter(|class| {
                    matches!(
                        *class,
                        "entr" | "exit" | "emph" | "path" | "verb" | "mediacall"
                    )
                })
                .unwrap_or("entr"),
        }
    }

    /// Get the preset class string (deprecated, use `preset_class` instead).
    #[deprecated(note = "Use preset_class() and preset_id() instead")]
    #[must_use]
    pub fn to_preset(&self) -> &str {
        self.preset_class()
    }
}

/// `EffectInstance` trigger type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Trigger {
    /// Start on click
    #[default]
    OnClick,
    /// Start with previous animation
    WithPrevious,
    /// Start after previous animation
    AfterPrevious,
}

/// Unsigned identifier linking a timing node to an entry in `p:bldLst`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct GroupId(u32);

impl GroupId {
    /// Construct an OOXML timing group identifier.
    #[must_use]
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the encoded unsigned group identifier.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl From<u32> for GroupId {
    fn from(value: u32) -> Self {
        Self::new(value)
    }
}

/// Event filtering supported by `PowerPoint` for a triggered sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EventFilter {
    /// Prevent the trigger event from bubbling beyond the interactive sequence.
    CancelBubble,
}

/// Structural sequence containing an animation effect.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum SequenceContext {
    /// The slide's ordinary click sequence.
    #[default]
    Main,
    /// A sequence activated by clicking a shape on the slide.
    Interactive {
        /// Shape whose click activates or advances the sequence.
        trigger_shape_id: u32,
        /// Optional `PowerPoint` event-bubbling filter on the `interactiveSeq` cTn.
        event_filter: Option<EventFilter>,
    },
}

/// `EffectInstance` direction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Direction {
    Up,
    Down,
    Left,
    Right,
    UpLeft,
    UpRight,
    DownLeft,
    DownRight,
    /// Toward the center, used by Zoom.
    In,
    /// Away from the center, used by Zoom.
    Out,
    /// Horizontal closing split.
    HorizontalIn,
    /// Horizontal opening split.
    HorizontalOut,
    /// Vertical closing split.
    VerticalIn,
    /// Vertical opening split.
    VerticalOut,
    /// A subtle zoom toward the center.
    InSlightly,
    /// A subtle zoom away from the center.
    OutSlightly,
    /// Zoom toward the center beginning at the screen center.
    InFromScreenCenter,
    /// Zoom away from the center ending at the screen center.
    OutFromScreenCenter,
}

/// An animation applied to a shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EffectInstance {
    /// Target shape ID
    pub shape_id: u32,
    /// `EffectInstance` effect
    pub effect: Effect,
    /// Trigger type
    pub trigger: Trigger,
    /// Duration in milliseconds
    pub duration: Duration,
    /// Delay before starting (ms)
    pub delay: u32,
    /// Direction (for directional effects)
    pub direction: Option<Direction>,
    /// Property state retained after the animation becomes inactive.
    pub fill: Option<Fill>,
    /// Policy for restarting this time node.
    pub restart: Option<Restart>,
    /// Whether the animation runs backward after reaching its end.
    pub auto_reverse: bool,
    /// Optional repeat count.
    pub repeat: Option<Repeat>,
    /// Optional nonzero playback speed.
    pub speed: Option<Speed>,
    /// Optional acceleration fraction.
    pub acceleration: Option<MotionFraction>,
    /// Optional deceleration fraction.
    pub deceleration: Option<MotionFraction>,
    /// Whether this time node is visible in the animation user interface.
    pub display: Option<bool>,
    /// Optional total duration for repeated playback.
    pub repeat_duration: Option<Duration>,
    /// Optional synchronization policy with the containing time group.
    pub sync_behavior: Option<SyncBehavior>,
    /// Whether this node is an after-effect.
    pub after_effect: Option<bool>,
    /// Optional normalized-time warp filter.
    pub time_filter: Option<TimeFilter>,
    /// Main-sequence or shape-triggered interactive-sequence context.
    pub sequence_context: SequenceContext,
    /// Optional build-list group containing this effect time node.
    pub group_id: Option<GroupId>,
    /// Sequence order (1-based)
    pub order: u32,
}

impl EffectInstance {
    /// Create a new animation.
    #[must_use]
    pub fn new(shape_id: u32, effect: Effect) -> Self {
        Self {
            shape_id,
            effect,
            trigger: Trigger::OnClick,
            duration: Duration::Finite(500),
            delay: 0,
            direction: None,
            fill: Some(Fill::Hold),
            restart: None,
            auto_reverse: false,
            repeat: None,
            speed: None,
            acceleration: None,
            deceleration: None,
            display: None,
            repeat_duration: None,
            sync_behavior: None,
            after_effect: None,
            time_filter: None,
            sequence_context: SequenceContext::Main,
            group_id: None,
            order: 1,
        }
    }

    /// Set the trigger type.
    #[must_use]
    pub fn with_trigger(mut self, trigger: Trigger) -> Self {
        self.trigger = trigger;
        self
    }

    /// Set the duration.
    #[must_use]
    pub fn with_duration(mut self, duration: impl Into<Duration>) -> Self {
        self.duration = duration.into();
        self
    }

    /// Set a finite duration in milliseconds.
    #[must_use]
    pub fn with_duration_ms(mut self, duration: u32) -> Self {
        self.duration = Duration::Finite(duration);
        self
    }

    /// Set the delay.
    #[must_use]
    pub fn with_delay(mut self, delay: u32) -> Self {
        self.delay = delay;
        self
    }

    /// Set the direction.
    #[must_use]
    pub fn with_direction(mut self, direction: Direction) -> Self {
        self.direction = Some(direction);
        self
    }

    /// Set the fill behavior for the animation time node.
    #[must_use]
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set the restart behavior for the animation time node.
    #[must_use]
    pub fn with_restart(mut self, restart: Restart) -> Self {
        self.restart = Some(restart);
        self
    }

    /// Enable or disable automatic reversal.
    #[must_use]
    pub fn with_auto_reverse(mut self, auto_reverse: bool) -> Self {
        self.auto_reverse = auto_reverse;
        self
    }

    /// Set the repeat behavior for the animation time node.
    #[must_use]
    pub fn with_repeat(mut self, repeat: Repeat) -> Self {
        self.repeat = Some(repeat);
        self
    }

    /// Set the nonzero playback speed.
    #[must_use]
    pub fn with_speed(mut self, speed: Speed) -> Self {
        self.speed = Some(speed);
        self
    }

    /// Set the acceleration fraction.
    #[must_use]
    pub fn with_acceleration(mut self, acceleration: MotionFraction) -> Self {
        self.acceleration = Some(acceleration);
        self
    }

    /// Set the deceleration fraction.
    #[must_use]
    pub fn with_deceleration(mut self, deceleration: MotionFraction) -> Self {
        self.deceleration = Some(deceleration);
        self
    }

    /// Set whether this node is displayed in animation user interfaces.
    #[must_use]
    pub fn with_display(mut self, display: bool) -> Self {
        self.display = Some(display);
        self
    }

    /// Set the total duration for repeated playback.
    #[must_use]
    pub fn with_repeat_duration(mut self, duration: impl Into<Duration>) -> Self {
        self.repeat_duration = Some(duration.into());
        self
    }

    /// Set synchronization with the containing time group.
    #[must_use]
    pub fn with_sync_behavior(mut self, behavior: SyncBehavior) -> Self {
        self.sync_behavior = Some(behavior);
        self
    }

    /// Mark or unmark this node as an after-effect.
    #[must_use]
    pub fn with_after_effect(mut self, after_effect: bool) -> Self {
        self.after_effect = Some(after_effect);
        self
    }

    /// Set a normalized-time warp filter.
    #[must_use]
    pub fn with_time_filter(mut self, filter: TimeFilter) -> Self {
        self.time_filter = Some(filter);
        self
    }

    /// Put this effect in a shape-triggered sequence using `PowerPoint`'s filter.
    #[must_use]
    pub fn with_interactive_trigger(mut self, trigger_shape_id: u32) -> Self {
        self.sequence_context = SequenceContext::Interactive {
            trigger_shape_id,
            event_filter: Some(EventFilter::CancelBubble),
        };
        self
    }

    /// Set the structural sequence context explicitly.
    #[must_use]
    pub fn with_sequence_context(mut self, context: SequenceContext) -> Self {
        self.sequence_context = context;
        self
    }

    /// Associate this effect cTn with a build-list timing group.
    #[must_use]
    pub fn with_group_id(mut self, group_id: impl Into<GroupId>) -> Self {
        self.group_id = Some(group_id.into());
        self
    }
}
