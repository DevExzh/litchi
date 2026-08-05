use crate::package::Result;

use super::package::corrupted;

/// Placeholder metadata embedded in a PowerPoint shape's client data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Placeholder {
    /// Placeholder position from the PowerPoint `PlaceholderAtom`.
    pub position: Option<u16>,
    /// Exact PowerPoint placeholder kind.
    pub kind: crate::PlaceholderKind,
    /// Checked PowerPoint placeholder size.
    pub size: crate::AtomPlaceholderSize,
}

/// Host-specific meaning of an OfficeArt picture-frame shape.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FrameKind {
    /// An ordinary picture frame.
    #[default]
    Picture,
    /// A frame associated with an embedded or linked OLE object.
    Object,
    /// A frame associated with audio or video media.
    Media,
}

/// Checked PowerPoint shape bounds projected from either anchor encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Anchor {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
    width: i32,
    height: i32,
}

impl Anchor {
    pub(super) fn new(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        let width = right.checked_sub(left).ok_or_else(|| {
            corrupted("PowerPoint shape anchor width exceeds a signed coordinate")
        })?;
        let height = bottom.checked_sub(top).ok_or_else(|| {
            corrupted("PowerPoint shape anchor height exceeds a signed coordinate")
        })?;
        if width < 0 || height < 0 {
            return Err(corrupted("PowerPoint shape anchor has inverted bounds"));
        }
        Ok(Self {
            left,
            top,
            right,
            bottom,
            width,
            height,
        })
    }

    /// Minimum x-coordinate.
    pub const fn left(self) -> i32 {
        self.left
    }

    /// Minimum y-coordinate.
    pub const fn top(self) -> i32 {
        self.top
    }

    /// Maximum x-coordinate.
    pub const fn right(self) -> i32 {
        self.right
    }

    /// Maximum y-coordinate.
    pub const fn bottom(self) -> i32 {
        self.bottom
    }

    /// Width in PowerPoint master units.
    pub const fn width(self) -> i32 {
        self.width
    }

    /// Height in PowerPoint master units.
    pub const fn height(self) -> i32 {
        self.height
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct Frame {
    pub(super) kind: FrameKind,
    pub(super) object_id: Option<u32>,
}

/// PowerPoint-only behavior layered over a neutral OfficeArt shape.
///
/// Import this trait as `_` at PPT call sites.  The resulting method surface
/// stays short without exposing `ClientData` record IDs to API users.
pub trait ShapeExt {
    /// Decodes the shape's PPT textbox payload.
    fn text(&self) -> Result<Option<String>>;

    /// Projects legacy placeholder metadata into checked semantic values.
    fn placeholder(&self) -> Result<Option<Placeholder>>;

    /// Distinguishes pictures, OLE frames, and media frames.
    fn frame_kind(&self) -> Result<FrameKind>;

    /// Returns the PPT external-object reference for an OLE or media frame.
    fn external_object_id(&self) -> Result<Option<u32>>;

    /// Parses click and mouse-over actions with default limits.
    fn interactions(&self) -> Result<Vec<crate::Interaction>>;

    /// Parses click and mouse-over actions with explicit limits.
    fn interactions_with_limits(
        &self,
        limits: crate::InteractionLimits,
    ) -> Result<Vec<crate::Interaction>>;

    /// Parses range-anchored text actions with default limits.
    fn text_interactions(&self) -> Result<Vec<crate::TextInteraction>>;

    /// Parses range-anchored text actions with explicit limits.
    fn text_interactions_with_limits(
        &self,
        limits: crate::TextInteractionLimits,
    ) -> Result<Vec<crate::TextInteraction>>;

    /// Parses a context-validated placeholder atom with default limits.
    fn placeholder_atom(
        &self,
        context: crate::PlaceholderContext,
    ) -> Result<Option<crate::PlaceholderAtom>>;

    /// Parses a context-validated placeholder atom with explicit limits.
    fn placeholder_atom_with_limits(
        &self,
        context: crate::PlaceholderContext,
        limits: crate::PlaceholderLimits,
    ) -> Result<Option<crate::PlaceholderAtom>>;

    /// Parses PowerPoint 12 shape round-trip metadata.
    fn powerpoint12_shape_metadata(&self) -> Result<Option<crate::ShapeMetadata>>;

    /// Parses inert shape programmable tags with default limits.
    fn programmable_tags(&self) -> Result<Option<crate::ShapeProgrammableTags>>;

    /// Parses inert shape programmable tags with explicit limits.
    fn programmable_tags_with_limits(
        &self,
        limits: crate::ShapeProgrammableTagLimits,
    ) -> Result<Option<crate::ShapeProgrammableTags>>;

    /// Parses the PPT shape-flag projection with default limits.
    fn ppt_flags(&self) -> Result<Option<crate::ShapeFlagProjection>>;

    /// Parses the PPT shape-flag projection with explicit limits.
    fn ppt_flags_with(
        &self,
        limits: crate::ShapeFlagLimits,
    ) -> Result<Option<crate::ShapeFlagProjection>>;

    /// Parses inert legacy PowerPoint animation metadata.
    fn animation(&self) -> Result<Option<crate::animation::AnimationInfo>>;
}
