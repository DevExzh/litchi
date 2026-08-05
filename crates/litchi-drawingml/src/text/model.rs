//! Semantic DrawingML text domains.
//!
//! These compact, host-neutral values describe the closed schema domains
//! used by text bodies; lexical behavior is implemented by `codec`.

/// Vertical anchoring within a text body (`ST_TextAnchoringType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Anchor {
    /// Start at the top inset (`t`).
    #[default]
    Top,
    /// Center vertically (`ctr`).
    Center,
    /// End at the bottom inset (`b`).
    Bottom,
    /// Spread lines to fill the body (`just`).
    Justified,
    /// Spread words to fill the body (`dist`).
    Distributed,
}

/// Direction of text within a shape (`ST_TextVerticalType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Direction {
    /// Horizontal text (`horz`).
    #[default]
    Horizontal,
    /// Lines rotated 90 degrees (`vert`).
    Vertical,
    /// Lines rotated 270 degrees (`vert270`).
    Vertical270,
    /// Upright, stacked WordArt letters (`wordArtVert`).
    WordArtVertical,
    /// East Asian vertical text (`eaVert`).
    EastAsianVertical,
    /// Mongolian vertical text (`mongolianVert`).
    MongolianVertical,
    /// Right-to-left WordArt vertical text (`wordArtVertRtl`).
    WordArtVerticalRtl,
}

/// Whether text wraps inside the shape extents (`ST_TextWrappingType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Wrap {
    /// Wrap inside the bounding rectangle (`square`).
    #[default]
    Square,
    /// Do not wrap (`none`).
    None,
}

/// Autofit behavior selected by a text-body child element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Autofit {
    /// `a:noAutofit`.
    #[default]
    None,
    /// `a:spAutoFit`.
    Shape,
    /// `a:normAutofit`.
    Normal,
}

/// Lossless underline style shared by DrawingML and WordprocessingML.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Underline {
    #[default]
    None,
    Words,
    Single,
    Double,
    Heavy,
    Dotted,
    DottedHeavy,
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Wavy,
    WavyHeavy,
    WavyDouble,
}
