//! Semantic `PresentationML` action-setting values.

use litchi_opc::PackURI;

/// The user interaction that activates an action setting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Trigger {
    /// The action is configured for a click.
    Click,
    /// The action is configured for a pointer hover.
    Hover,
}

/// A reserved `PowerPoint` slide-show jump target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Jump {
    /// End the slide show.
    EndShow,
    /// Jump to the first slide.
    FirstSlide,
    /// Jump to the last slide.
    LastSlide,
    /// Jump to the most recently viewed slide.
    LastSlideViewed,
    /// Jump to the next slide.
    NextSlide,
    /// Jump to the previous slide.
    PreviousSlide,
}

/// The recognized meaning of a stored action string.
///
/// The original action string remains available from [`Setting::action`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    /// A regular hyperlink relationship without a `PowerPoint` action string.
    Hyperlink,
    /// A relationship-targeted jump to a slide in this presentation.
    SlideJump,
    /// A reserved presentation-relative slide-show jump.
    SlideShowJump(Jump),
    /// A named custom show identified by its stored numeric ID.
    CustomShow { id: u32 },
    /// An external file action.
    File,
    /// An external presentation action with its stored starting slide index.
    Presentation { start_slide_index: u32 },
    /// A macro action. The stored macro name remains inert in [`Setting::action`].
    Macro,
    /// An external-program action.
    Program,
    /// A media-playback action.
    Media,
    /// A setting without an action string or relationship reference.
    None,
    /// An action string outside the bounded recognized `PowerPoint` vocabulary.
    Unknown,
}

/// A declared relationship target attached to an action setting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    /// An internal OPC target part.
    Internal {
        /// Absolute package part name.
        part_name: PackURI,
        /// Declared relationship type URI.
        relationship_type: String,
    },
    /// An external target retained as an inert string.
    External {
        /// Declared target URI or path.
        target: String,
        /// Declared relationship type URI.
        relationship_type: String,
    },
}

impl Target {
    /// Return the declared relationship type URI.
    #[inline]
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        match self {
            Self::Internal {
                relationship_type, ..
            }
            | Self::External {
                relationship_type, ..
            } => relationship_type,
        }
    }

    /// Return the target part name for an internal relationship.
    #[inline]
    #[must_use]
    pub fn part_name(&self) -> Option<&PackURI> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }

    /// Return the stored target string for an external relationship.
    #[inline]
    #[must_use]
    pub fn external_target(&self) -> Option<&str> {
        match self {
            Self::Internal { .. } => None,
            Self::External { target, .. } => Some(target),
        }
    }
}

/// An inert `PowerPoint` click or hover action setting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Setting {
    pub(super) slide_index: usize,
    pub(super) action_index: usize,
    pub(super) trigger: Trigger,
    pub(super) kind: Kind,
    pub(super) action: Option<String>,
    pub(super) relationship_id: Option<String>,
    pub(super) target: Option<Target>,
    pub(super) tooltip: Option<String>,
    pub(super) target_frame: Option<String>,
}

impl Setting {
    /// Return the zero-based index of the slide that owns this setting.
    #[inline]
    #[must_use]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this setting on the slide.
    #[inline]
    #[must_use]
    pub fn action_index(&self) -> usize {
        self.action_index
    }

    /// Return whether this setting is activated by click or hover.
    #[inline]
    #[must_use]
    pub fn trigger(&self) -> Trigger {
        self.trigger
    }

    /// Return the recognized action kind.
    #[inline]
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the original stored action string, when present.
    #[inline]
    #[must_use]
    pub fn action(&self) -> Option<&str> {
        self.action.as_deref()
    }

    /// Return the optional relationship ID from the owning slide.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Return the declared relationship target, when present.
    #[inline]
    #[must_use]
    pub fn target(&self) -> Option<&Target> {
        self.target.as_ref()
    }

    /// Return the optional stored tooltip.
    #[inline]
    #[must_use]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }

    /// Return the optional stored target frame.
    #[inline]
    #[must_use]
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }
}

pub(super) fn classify(action: Option<&str>, has_relationship: bool) -> Kind {
    let Some(action) = action else {
        return if has_relationship {
            Kind::Hyperlink
        } else {
            Kind::None
        };
    };
    match action {
        "ppaction://hlinkfile" => Kind::File,
        "ppaction://hlinksldjump" => Kind::SlideJump,
        "ppaction://hlinkshowjump?jump=endshow" => Kind::SlideShowJump(Jump::EndShow),
        "ppaction://hlinkshowjump?jump=firstslide" => Kind::SlideShowJump(Jump::FirstSlide),
        "ppaction://hlinkshowjump?jump=lastslide" => Kind::SlideShowJump(Jump::LastSlide),
        "ppaction://hlinkshowjump?jump=lastslideviewed" => {
            Kind::SlideShowJump(Jump::LastSlideViewed)
        },
        "ppaction://hlinkshowjump?jump=nextslide" => Kind::SlideShowJump(Jump::NextSlide),
        "ppaction://hlinkshowjump?jump=previousslide" => Kind::SlideShowJump(Jump::PreviousSlide),
        "ppaction://program" => Kind::Program,
        "ppaction://media" => Kind::Media,
        action if action.starts_with("ppaction://customshow?id=") => action
            .strip_prefix("ppaction://customshow?id=")
            .and_then(|value| value.parse().ok())
            .map_or(Kind::Unknown, |id| Kind::CustomShow { id }),
        action if action.starts_with("ppaction://hlinkpres?slideindex=") => action
            .strip_prefix("ppaction://hlinkpres?slideindex=")
            .and_then(|value| value.parse().ok())
            .map_or(Kind::Unknown, |start_slide_index| Kind::Presentation {
                start_slide_index,
            }),
        action
            if action.starts_with("ppaction://macro?name=")
                && action
                    .strip_prefix("ppaction://macro?name=")
                    .is_some_and(|name| !name.is_empty()) =>
        {
            Kind::Macro
        },
        _ => Kind::Unknown,
    }
}
