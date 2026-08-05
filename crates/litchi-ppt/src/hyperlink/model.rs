//! Typed semantic hyperlink and interaction values for legacy PowerPoint.

/// Mouse event that triggers an interactive action.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionTrigger {
    /// Mouse click.
    Click,
    /// Mouse pointer moved over the object.
    MouseOver,
}

/// Action stored in an `InteractiveInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionAction {
    NoAction,
    Macro,
    RunProgram,
    Jump,
    Hyperlink,
    Ole,
    Media,
    CustomShow,
}

/// Relative slide-show jump target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionJump {
    None,
    NextSlide,
    PreviousSlide,
    FirstSlide,
    LastSlide,
    LastSlideViewed,
    EndShow,
}

/// Interpretation of an interactive hyperlink reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionLinkTarget {
    NextSlide,
    PreviousSlide,
    FirstSlide,
    LastSlide,
    CustomShow,
    SlideNumber,
    Url,
    OtherPresentation,
    OtherFile,
    Nil,
}

/// Resource limits for an interactive-information container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InteractionLimits {
    /// Maximum complete container size, including its eight-byte header.
    pub max_record_bytes: usize,
    /// Maximum MacroNameAtom UTF-16 payload size.
    pub max_macro_name_bytes: usize,
}

impl Default for InteractionLimits {
    fn default() -> Self {
        Self {
            max_record_bytes: 1024 * 1024,
            max_macro_name_bytes: 64 * 1024,
        }
    }
}

/// Typed payload of an MS-PPT §2.6.10 InteractiveInfoAtom.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InteractiveInfoAtom {
    pub sound_id: u32,
    pub hyperlink_id: u32,
    pub action: InteractionAction,
    pub ole_verb: u8,
    pub jump: InteractionJump,
    pub animated: bool,
    pub stop_sound: bool,
    pub custom_show_return: bool,
    pub visited: bool,
    pub link_target: InteractionLinkTarget,
    /// Undefined bytes retained without interpretation.
    pub unused: [u8; 3],
}

/// Inert MS-PPT §2.6.11 MacroNameAtom data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroNameAtom {
    pub(super) text: String,
    pub(super) raw_utf16: Vec<u8>,
}

/// One click or mouse-over action attached to a shape or text range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Interaction {
    pub trigger: InteractionTrigger,
    pub sound_id: u32,
    pub hyperlink_id: u32,
    pub action: InteractionAction,
    pub ole_verb: u8,
    pub jump: InteractionJump,
    pub animated: bool,
    pub stop_sound: bool,
    pub custom_show_return: bool,
    pub visited: bool,
    pub link_target: InteractionLinkTarget,
    pub macro_name: Option<String>,
    /// Undefined atom bytes retained verbatim.
    pub unused: [u8; 3],
    /// Exact inert MacroNameAtom UTF-16 data, if present.
    pub macro_name_data: Option<Vec<u8>>,
}

/// Click and mouse-over actions attached to one slide shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeInteractionEntry {
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// At most one action for each [`InteractionTrigger`].
    pub interactions: Vec<Interaction>,
}

/// Additional hyperlink data introduced by PowerPoint 9.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkExtension {
    /// Optional text displayed as a hover screen tip.
    pub screen_tip: Option<String>,
    /// Whether the hyperlink was created in the Insert Hyperlink dialog.
    pub inserted_with_dialog: bool,
    /// Whether the base hyperlink location names a custom slide show.
    pub location_is_named_show: bool,
    /// Whether a named show returns to the originating slide.
    pub named_show_returns_to_slide: bool,
}

/// One base PowerPoint hyperlink definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    /// Positive identifier referenced by interactive information records.
    pub id: u32,
    /// Optional user-readable hyperlink name.
    pub friendly_name: Option<String>,
    /// Optional full destination-file path or URL.
    pub target: Option<String>,
    /// Optional location within the destination.
    pub location: Option<String>,
    /// Optional PowerPoint 9 metadata for this hyperlink.
    pub extension: Option<HyperlinkExtension>,
}

/// Hyperlink definitions resolved with their PowerPoint 9 extensions.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Hyperlinks {
    /// Seed used when allocating new external-object or hyperlink identifiers.
    pub id_seed: Option<i32>,
    /// Hyperlinks in base `ExObjListContainer` order.
    pub hyperlinks: Vec<Hyperlink>,
}
