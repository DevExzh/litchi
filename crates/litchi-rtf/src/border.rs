//! RTF border and shading support.
//!
//! This module provides support for borders and shading in RTF documents.

use super::types::ColorRef;

/// Border style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderStyle {
    /// No border
    #[default]
    None,
    /// Single line border
    Single,
    /// Dotted border
    Dotted,
    /// Dashed border
    Dashed,
    /// Double line border
    Double,
    /// Triple line border
    Triple,
    /// Thick-thin small gap
    ThickThinSmall,
    /// Thin-thick small gap
    ThinThickSmall,
    /// Thin-thick-thin small gap
    ThinThickThinSmall,
    /// Thick-thin medium gap
    ThickThinMedium,
    /// Thin-thick medium gap
    ThinThickMedium,
    /// Thin-thick-thin medium gap
    ThinThickThinMedium,
    /// Thick-thin large gap
    ThickThinLarge,
    /// Thin-thick large gap
    ThinThickLarge,
    /// Thin-thick-thin large gap
    ThinThickThinLarge,
    /// Wavy border
    Wavy,
    /// Double wavy border
    WavyDouble,
    /// Striped border
    Striped,
    /// Embossed border
    Embossed,
    /// Engraved border
    Engraved,
    /// Outset border (3D)
    Outset,
    /// Inset border (3D)
    Inset,
}

/// Border definition
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Border {
    /// Border style
    pub style: BorderStyle,
    /// Border width (in twips)
    pub width: i32,
    /// Border color reference
    pub color_ref: ColorRef,
    /// Space between border and content (in twips)
    pub space: i32,
    /// Whether border has shadow
    pub shadow: bool,
    /// Whether border is frame (surrounds text)
    pub frame: bool,
}

impl Default for Border {
    fn default() -> Self {
        Self {
            style: BorderStyle::default(),
            width: 15, // 1pt
            color_ref: 0,
            space: 0,
            shadow: false,
            frame: false,
        }
    }
}

impl Border {
    /// Create a new border
    #[inline]
    pub fn new(style: BorderStyle) -> Self {
        Self {
            style,
            ..Default::default()
        }
    }

    /// Check if border is visible
    #[inline]
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None && self.width > 0
    }
}

/// Borders for a paragraph or table cell
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Borders {
    /// Top border
    pub top: Border,
    /// Bottom border
    pub bottom: Border,
    /// Left border
    pub left: Border,
    /// Right border
    pub right: Border,
}

impl Borders {
    /// Create new empty borders
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set all borders to the same style
    #[inline]
    pub fn all(border: Border) -> Self {
        Self {
            top: border,
            bottom: border,
            left: border,
            right: border,
        }
    }

    /// Check if any border is visible
    #[inline]
    pub fn has_any_border(&self) -> bool {
        self.top.is_visible()
            || self.bottom.is_visible()
            || self.left.is_visible()
            || self.right.is_visible()
    }
}

/// Shading pattern
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShadingPattern {
    /// Clear (no shading)
    #[default]
    Clear,
    /// Solid fill
    Solid,
    /// Horizontal stripes
    Horizontal,
    /// Vertical stripes
    Vertical,
    /// Forward diagonal stripes
    ForwardDiagonal,
    /// Backward diagonal stripes
    BackwardDiagonal,
    /// Crosshatch
    Cross,
    /// Diagonal crosshatch
    DiagonalCross,
    /// Dark horizontal
    DarkHorizontal,
    /// Dark vertical
    DarkVertical,
    /// Dark forward diagonal
    DarkForwardDiagonal,
    /// Dark backward diagonal
    DarkBackwardDiagonal,
    /// Dark crosshatch
    DarkCross,
    /// Dark diagonal crosshatch
    DarkDiagonalCross,
    /// 5% fill
    Percent5,
    /// 10% fill
    Percent10,
    /// 12.5% fill
    Percent12,
    /// 15% fill
    Percent15,
    /// 20% fill
    Percent20,
    /// 25% fill
    Percent25,
    /// 30% fill
    Percent30,
    /// 35% fill
    Percent35,
    /// 40% fill
    Percent40,
    /// 45% fill
    Percent45,
    /// 50% fill
    Percent50,
    /// 55% fill
    Percent55,
    /// 60% fill
    Percent60,
    /// 62.5% fill
    Percent62,
    /// 65% fill
    Percent65,
    /// 70% fill
    Percent70,
    /// 75% fill
    Percent75,
    /// 80% fill
    Percent80,
    /// 85% fill
    Percent85,
    /// 87.5% fill
    Percent87,
    /// 90% fill
    Percent90,
    /// 95% fill
    Percent95,
}

/// Shading/background fill
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Shading {
    /// Shading pattern
    pub pattern: ShadingPattern,
    /// Foreground color (pattern color)
    pub foreground_color: ColorRef,
    /// Background color (fill color)
    pub background_color: ColorRef,
}

impl Shading {
    /// Create new shading
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Create solid color shading
    #[inline]
    pub fn solid(color: ColorRef) -> Self {
        Self {
            pattern: ShadingPattern::Solid,
            foreground_color: color,
            background_color: color,
        }
    }

    /// Check if shading is visible
    #[inline]
    pub fn is_visible(&self) -> bool {
        self.pattern != ShadingPattern::Clear
    }
}

/// Tab stop alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TabAlignment {
    /// Left-aligned tab
    #[default]
    Left,
    /// Right-aligned tab
    Right,
    /// Centered tab
    Center,
    /// Decimal tab (align on decimal point)
    Decimal,
    /// Bar tab (vertical bar)
    Bar,
}

/// Tab stop leader character
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TabLeader {
    /// No leader
    #[default]
    None,
    /// Dot leader (........)
    Dot,
    /// Hyphen leader (--------)
    Hyphen,
    /// Underscore leader (________)
    Underscore,
    /// Thick line leader
    ThickLine,
    /// Equal sign leader (========)
    Equal,
}

/// Tab stop definition
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TabStop {
    /// Position (in twips from left margin)
    pub position: i32,
    /// Alignment
    pub alignment: TabAlignment,
    /// Leader character
    pub leader: TabLeader,
}

impl TabStop {
    /// Create a new tab stop
    #[inline]
    pub fn new(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::default(),
            leader: TabLeader::default(),
        }
    }

    /// Create a left-aligned tab stop
    #[inline]
    pub fn left(position: i32) -> Self {
        Self::new(position)
    }

    /// Create a right-aligned tab stop
    #[inline]
    pub fn right(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Right,
            leader: TabLeader::default(),
        }
    }

    /// Create a centered tab stop
    #[inline]
    pub fn center(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Center,
            leader: TabLeader::default(),
        }
    }

    /// Create a decimal tab stop
    #[inline]
    pub fn decimal(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Decimal,
            leader: TabLeader::default(),
        }
    }
}
