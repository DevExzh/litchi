//! Package-independent, inert slide-show event values.

use crate::time::Offset;

/// A trigger type recorded by a PowerPoint slide show.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Trigger {
    None,
    OnBegin,
    OnEnd,
    Begin,
    End,
    OnClick,
    OnDoubleClick,
    OnMouseOver,
    OnMouseOut,
    OnNext,
    OnPrevious,
    OnStopAudio,
}

/// The recorded action represented by a PowerPoint slide-show event.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Kind {
    Trigger(Trigger),
    Play,
    Stop,
    Pause,
    Resume,
    /// Seek the targeted media object to an exact stream offset.
    Seek {
        at: Offset,
    },
    /// A reserved unknown event record for future PowerPoint extensions.
    Null,
}

/// A bounded, inert event record discovered on a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Event {
    pub(crate) slide_index: usize,
    pub(crate) event_index: usize,
    pub(crate) kind: Kind,
    pub(crate) time: Offset,
    pub(crate) object_id: u32,
}

impl Event {
    /// Return the zero-based index of the slide that owns this event.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this event on its slide.
    #[inline]
    pub fn event_index(&self) -> usize {
        self.event_index
    }

    /// Return the recorded event kind.
    #[inline]
    pub fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Return the exact normalized time offset in the slide timeline.
    #[inline]
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the DrawingML object identifier targeted by this event.
    #[inline]
    pub fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Return the exact normalized media-stream offset for a seek event.
    #[inline]
    pub fn seek_time(&self) -> Option<&Offset> {
        match &self.kind {
            Kind::Seek { at } => Some(at),
            _ => None,
        }
    }
}

/// A slide-show event ready for storage onto a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Draft {
    pub(crate) kind: Kind,
    pub(crate) time: Offset,
    pub(crate) object_id: u32,
}

impl Draft {
    /// Create an event from its complete, typed action state.
    #[must_use]
    pub fn new(kind: Kind, time: Offset, object_id: u32) -> Self {
        Self {
            kind,
            time,
            object_id,
        }
    }

    /// Create a trigger event.
    #[must_use]
    pub fn trigger(trigger: Trigger, time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Trigger(trigger), time, object_id)
    }

    /// Create a play event.
    #[must_use]
    pub fn play(time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Play, time, object_id)
    }

    /// Create a stop event.
    #[must_use]
    pub fn stop(time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Stop, time, object_id)
    }

    /// Create a pause event.
    #[must_use]
    pub fn pause(time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Pause, time, object_id)
    }

    /// Create a resume event.
    #[must_use]
    pub fn resume(time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Resume, time, object_id)
    }

    /// Create a seek event with a media-stream offset.
    #[must_use]
    pub fn seek(time: Offset, object_id: u32, seek_time: Offset) -> Self {
        Self::new(Kind::Seek { at: seek_time }, time, object_id)
    }

    /// Create a reserved null event.
    #[must_use]
    pub fn null(time: Offset, object_id: u32) -> Self {
        Self::new(Kind::Null, time, object_id)
    }

    /// Return the recorded event kind.
    pub fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Return the exact normalized universal time offset.
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the DrawingML object identifier targeted by this event.
    pub fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Return the exact normalized media-stream offset for a seek event.
    pub fn seek_time(&self) -> Option<&Offset> {
        match &self.kind {
            Kind::Seek { at } => Some(at),
            _ => None,
        }
    }
}
