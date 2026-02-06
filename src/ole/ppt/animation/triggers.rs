//! Advanced animation triggers and interactive conditions.
//!
//! Provides support for complex trigger conditions beyond simple click/auto.

/// Interactive trigger type for animations.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum InteractiveTrigger {
    /// On click of slide
    #[default]
    OnSlideClick,
    /// On click of specific shape
    OnShapeClick { shape_id: u32 },
    /// On bookmark
    OnBookmark { bookmark: String },
    /// On media end
    OnMediaEnd { media_id: u32 },
    /// On media start
    OnMediaStart { media_id: u32 },
    /// On previous animation end
    OnPreviousEnd,
    /// On next animation start
    OnNextStart,
    /// With previous animation (parallel)
    WithPrevious,
    /// After previous animation (delay)
    AfterPrevious { delay_ms: u32 },
    /// Automatic after delay
    Automatic { delay_ms: u32 },
}

impl InteractiveTrigger {
    /// Check if this is a click-based trigger.
    pub fn is_click_based(&self) -> bool {
        matches!(self, Self::OnSlideClick | Self::OnShapeClick { .. })
    }

    /// Check if this is time-based trigger.
    pub fn is_time_based(&self) -> bool {
        matches!(self, Self::AfterPrevious { .. } | Self::Automatic { .. })
    }

    /// Check if this is media-based trigger.
    pub fn is_media_based(&self) -> bool {
        matches!(self, Self::OnMediaEnd { .. } | Self::OnMediaStart { .. })
    }

    /// Get delay in milliseconds (if applicable).
    pub fn delay_ms(&self) -> Option<u32> {
        match self {
            Self::AfterPrevious { delay_ms } | Self::Automatic { delay_ms } => Some(*delay_ms),
            _ => None,
        }
    }

    /// Get shape ID for shape click triggers.
    pub fn shape_id(&self) -> Option<u32> {
        match self {
            Self::OnShapeClick { shape_id } => Some(*shape_id),
            _ => None,
        }
    }
}

/// Animation condition for advanced timing control.
#[derive(Debug, Clone, PartialEq)]
pub enum AnimationCondition {
    /// Begin condition
    Begin(BeginCondition),
    /// End condition
    End(EndCondition),
    /// Next condition
    Next(NextCondition),
    /// Previous condition
    Previous(PreviousCondition),
}

/// Begin condition for animation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BeginCondition {
    /// Begin on click
    OnClick,
    /// Begin with previous
    WithPrevious,
    /// Begin after previous
    AfterPrevious,
    /// Begin on next click
    OnNextClick,
    /// Begin on media begin
    OnMediaBegin,
    /// Begin on media end
    OnMediaEnd,
}

/// End condition for animation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EndCondition {
    /// End with animation
    WithAnimation,
    /// End after animation
    AfterAnimation,
    /// End on click
    OnClick,
    /// End on next click
    OnNextClick,
}

/// Next condition for sequencing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NextCondition {
    /// Next on click
    OnClick,
    /// Next after current
    AfterCurrent,
    /// Next with current
    WithCurrent,
}

/// Previous condition for sequencing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviousCondition {
    /// Seek to previous
    Seek,
    /// Skip previous
    Skip,
}

/// Iteration type for animations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum IterationType {
    /// Iterate all at once
    #[default]
    All,
    /// Iterate by element
    ByElement,
    /// Iterate by word
    ByWord,
    /// Iterate by letter
    ByLetter,
}

/// Animation repeat behavior.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum RepeatBehavior {
    /// No repeat
    #[default]
    None,
    /// Repeat count times
    Count(u32),
    /// Repeat for duration in milliseconds
    Duration(u32),
    /// Repeat indefinitely
    Indefinite,
}

impl RepeatBehavior {
    /// Check if animation repeats.
    pub fn repeats(&self) -> bool {
        !matches!(self, Self::None)
    }

    /// Get repeat count (if applicable).
    pub fn count(&self) -> Option<u32> {
        match self {
            Self::Count(n) => Some(*n),
            _ => None,
        }
    }

    /// Check if repeats indefinitely.
    pub fn is_indefinite(&self) -> bool {
        matches!(self, Self::Indefinite)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_interactive_trigger_click_based() {
        let trigger = InteractiveTrigger::OnSlideClick;
        assert!(trigger.is_click_based());
        assert!(!trigger.is_time_based());
        assert!(!trigger.is_media_based());
    }

    #[test]
    fn test_interactive_trigger_shape_click() {
        let trigger = InteractiveTrigger::OnShapeClick { shape_id: 100 };
        assert!(trigger.is_click_based());
        assert_eq!(trigger.shape_id(), Some(100));
    }

    #[test]
    fn test_interactive_trigger_time_based() {
        let trigger = InteractiveTrigger::Automatic { delay_ms: 1000 };
        assert!(trigger.is_time_based());
        assert_eq!(trigger.delay_ms(), Some(1000));
    }

    #[test]
    fn test_interactive_trigger_after_previous() {
        let trigger = InteractiveTrigger::AfterPrevious { delay_ms: 500 };
        assert!(trigger.is_time_based());
        assert_eq!(trigger.delay_ms(), Some(500));
    }

    #[test]
    fn test_iteration_type_default() {
        assert_eq!(IterationType::default(), IterationType::All);
    }

    #[test]
    fn test_repeat_behavior_none() {
        let repeat = RepeatBehavior::None;
        assert!(!repeat.repeats());
        assert!(!repeat.is_indefinite());
        assert_eq!(repeat.count(), None);
    }

    #[test]
    fn test_repeat_behavior_count() {
        let repeat = RepeatBehavior::Count(5);
        assert!(repeat.repeats());
        assert!(!repeat.is_indefinite());
        assert_eq!(repeat.count(), Some(5));
    }

    #[test]
    fn test_repeat_behavior_indefinite() {
        let repeat = RepeatBehavior::Indefinite;
        assert!(repeat.repeats());
        assert!(repeat.is_indefinite());
        assert_eq!(repeat.count(), None);
    }
}
