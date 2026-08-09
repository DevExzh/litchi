//! RTF border and shading support.
//!
//! This module provides support for borders and shading in RTF documents.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use super::types::ColorRef;

/// RTF 1.9.1 character-border line style selected after `chbrdr`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CharacterBorderStyle {
    #[default]
    None,
    Single,
    Thick,
    Dotted,
    Dashed,
    DashSmallGap,
    DotDash,
    DotDotDash,
    Double,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wavy,
    DoubleWavy,
    Striped,
    Embossed,
    Engraved,
    Outset,
    Inset,
    /// Hairline border (`\brdrhair`)
    Hairline,
}

impl CharacterBorderStyle {
    pub(crate) const fn control_word(self) -> &'static str {
        match self {
            Self::None => "brdrnone",
            Self::Single => "brdrs",
            Self::Thick => "brdrth",
            Self::Dotted => "brdrdot",
            Self::Dashed => "brdrdash",
            Self::DashSmallGap => "brdrdashsm",
            Self::DotDash => "brdrdashd",
            Self::DotDotDash => "brdrdashdd",
            Self::Double => "brdrdb",
            Self::Triple => "brdrtriple",
            Self::ThinThickSmallGap => "brdrtnthsg",
            Self::ThickThinSmallGap => "brdrthtnsg",
            Self::ThinThickThinSmallGap => "brdrtnthtnsg",
            Self::ThinThickMediumGap => "brdrtnthmg",
            Self::ThickThinMediumGap => "brdrthtnmg",
            Self::ThinThickThinMediumGap => "brdrtnthtnmg",
            Self::ThinThickLargeGap => "brdrtnthlg",
            Self::ThickThinLargeGap => "brdrthtnlg",
            Self::ThinThickThinLargeGap => "brdrtnthtnlg",
            Self::Wavy => "brdrwavy",
            Self::DoubleWavy => "brdrwavydb",
            Self::Striped => "brdrdashdotstr",
            Self::Embossed => "brdremboss",
            Self::Engraved => "brdrengrave",
            Self::Outset => "brdroutset",
            Self::Inset => "brdrinset",
            Self::Hairline => "brdrhair",
        }
    }
}

/// One character border. RTF applies the same border to every side of the run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterBorder {
    pub style: CharacterBorderStyle,
    /// Width in twips; RTF 1.9.1 limits `brdrw` to 75.
    pub width: u16,
    pub color_ref: ColorRef,
    /// Space between the border and text, in twips.
    pub space: u16,
    pub shadow: bool,
    pub frame: bool,
}

impl Default for CharacterBorder {
    fn default() -> Self {
        Self {
            style: CharacterBorderStyle::None,
            width: 10,
            color_ref: 0,
            space: 0,
            shadow: false,
            frame: false,
        }
    }
}

impl CharacterBorder {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.width > 75 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF character-border width must be in 0..=75 twips".to_string(),
            ));
        }
        Ok(())
    }
}

/// Exact character shading from the `chshdng`/`chbg*`, `chcfpat`, and
/// `chcbpat` families.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CharacterShading {
    /// Exact `\chshdngN` value in hundredths of a percent.
    pub amount: Option<u16>,
    /// Exact explicit `\chbg*` pattern.
    pub pattern: Option<ShadingPattern>,
    /// Exact `\chcfpatN` foreground color reference.
    pub foreground_color: Option<ColorRef>,
    /// Exact `\chcbpatN` background color reference.
    pub background_color: Option<ColorRef>,
}

impl CharacterShading {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.amount.is_some_and(|amount| amount > 10_000) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF character shading must be in 0..=10000".to_string(),
            ));
        }
        Ok(())
    }
}

/// Border style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderStyle {
    /// No border
    #[default]
    None,
    /// Single line border
    Single,
    /// Thick line border
    Thick,
    /// Dotted border
    Dotted,
    /// Dashed border
    Dashed,
    /// Dashed border with a small gap
    DashSmallGap,
    /// Alternating dot and dash
    DotDash,
    /// Alternating two dots and a dash
    DotDotDash,
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
    /// Hairline border (`\brdrhair`)
    Hairline,
}

impl BorderStyle {
    pub(crate) const fn control_word(self) -> &'static str {
        match self {
            Self::None => "brdrnone",
            Self::Single => "brdrs",
            Self::Thick => "brdrth",
            Self::Dotted => "brdrdot",
            Self::Dashed => "brdrdash",
            Self::DashSmallGap => "brdrdashsm",
            Self::DotDash => "brdrdashd",
            Self::DotDotDash => "brdrdashdd",
            Self::Double => "brdrdb",
            Self::Triple => "brdrtriple",
            Self::ThickThinSmall => "brdrthtnsg",
            Self::ThinThickSmall => "brdrtnthsg",
            Self::ThinThickThinSmall => "brdrtnthtnsg",
            Self::ThickThinMedium => "brdrthtnmg",
            Self::ThinThickMedium => "brdrtnthmg",
            Self::ThinThickThinMedium => "brdrtnthtnmg",
            Self::ThickThinLarge => "brdrthtnlg",
            Self::ThinThickLarge => "brdrtnthlg",
            Self::ThinThickThinLarge => "brdrtnthtnlg",
            Self::Wavy => "brdrwavy",
            Self::WavyDouble => "brdrwavydb",
            Self::Striped => "brdrdashdotstr",
            Self::Embossed => "brdremboss",
            Self::Engraved => "brdrengrave",
            Self::Outset => "brdroutset",
            Self::Inset => "brdrinset",
            Self::Hairline => "brdrhair",
        }
    }
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
    #[must_use]
    pub fn new(style: BorderStyle) -> Self {
        Self {
            style,
            ..Default::default()
        }
    }

    /// Check if border is visible
    #[inline]
    #[must_use]
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None && self.width > 0
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate_table(&self) -> crate::RtfResult<()> {
        if !(0..=75).contains(&self.width) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-border width must be in 0..=75 twips".to_string(),
            ));
        }
        if !(0..=crate::MAX_TABLE_DISTANCE_TWIPS).contains(&self.space) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-border spacing is out of range".to_string(),
            ));
        }
        Ok(())
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
    /// Bar border at the left edge of a bordered paragraph group (`\brdrbar`).
    pub bar: Border,
    /// Border between consecutive paragraphs of one border group (`\brdrbtw`).
    pub between: Border,
}

impl Borders {
    /// Create new empty borders
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set all borders to the same style
    #[inline]
    #[must_use]
    pub fn all(border: Border) -> Self {
        Self {
            top: border,
            bottom: border,
            left: border,
            right: border,
            bar: Border::default(),
            between: Border::default(),
        }
    }

    /// Check if any border is visible
    #[inline]
    #[must_use]
    pub fn has_any_border(&self) -> bool {
        self.top.is_visible()
            || self.bottom.is_visible()
            || self.left.is_visible()
            || self.right.is_visible()
            || self.bar.is_visible()
            || self.between.is_visible()
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
    /// Exact `\shadingN` value in hundredths of a percent.
    pub amount: Option<u16>,
    /// Legacy typed pattern used by mutation APIs when no exact amount exists.
    pub pattern: Option<ShadingPattern>,
    /// Exact `\cfpatN` foreground color reference.
    pub foreground_color: Option<ColorRef>,
    /// Exact `\cbpatN` background color reference.
    pub background_color: Option<ColorRef>,
}

impl Shading {
    /// Create new shading
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Create solid color shading
    #[inline]
    #[must_use]
    pub fn solid(color: ColorRef) -> Self {
        Self {
            amount: Some(10_000),
            pattern: Some(ShadingPattern::Solid),
            foreground_color: Some(color),
            background_color: Some(color),
        }
    }

    /// Whether any paragraph-shading control was explicitly retained.
    #[inline]
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.amount.is_some()
            || self.pattern.is_some()
            || self.foreground_color.is_some()
            || self.background_color.is_some()
    }

    /// Check if shading is visible
    #[inline]
    #[must_use]
    pub fn is_visible(&self) -> bool {
        self.amount.is_some_and(|amount| amount != 0)
            || self
                .pattern
                .is_some_and(|pattern| pattern != ShadingPattern::Clear)
            || self.foreground_color.is_some_and(|color| color != 0)
            || self.background_color.is_some_and(|color| color != 0)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.amount.is_some_and(|amount| amount > 10_000) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF paragraph shading must be in 0..=10000".to_string(),
            ));
        }
        Ok(())
    }

    /// Set or clear the exact `\shadingN` amount.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_amount(&mut self, amount: Option<u16>) -> crate::RtfResult<()> {
        if amount.is_some_and(|value| value > 10_000) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF paragraph shading must be in 0..=10000".to_string(),
            ));
        }
        self.amount = amount;
        Ok(())
    }

    /// Set or clear the paragraph shading foreground color.
    #[inline]
    pub fn set_foreground_color(&mut self, color: Option<ColorRef>) {
        self.foreground_color = color;
    }

    /// Set or clear the paragraph shading background color.
    #[inline]
    pub fn set_background_color(&mut self, color: Option<ColorRef>) {
        self.background_color = color;
    }

    /// Replace this metadata with the omitted/default state.
    #[inline]
    pub fn clear(&mut self) {
        *self = Self::default();
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
    /// Middle-dot leader
    MiddleDot,
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

impl Default for TabStop {
    fn default() -> Self {
        Self::new(0)
    }
}

/// Maximum number of custom tab stops retained for one paragraph.
///
/// The RTF grammar permits a repeated sequence without an encoded count. This
/// implementation bound keeps paragraph state fixed-size and prevents an
/// attacker-controlled allocation while accommodating normal word-processor
/// output.
pub const MAX_PARAGRAPH_TAB_STOPS: usize = 64;

/// Fixed-capacity custom tab stops for a paragraph or paragraph style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TabStops {
    entries: [TabStop; MAX_PARAGRAPH_TAB_STOPS],
    len: u8,
}

impl TabStops {
    /// Create an empty collection.
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            entries: [TabStop {
                position: 0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }; MAX_PARAGRAPH_TAB_STOPS],
            len: 0,
        }
    }

    /// Number of stored custom tab stops.
    #[inline]
    #[must_use]
    pub const fn len(&self) -> usize {
        self.len as usize
    }

    /// Whether no custom tab stops are stored.
    #[inline]
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Stored tab stops in declaration order.
    #[inline]
    #[must_use]
    pub fn as_slice(&self) -> &[TabStop] {
        self.entries.get(..self.len()).unwrap_or(&[])
    }

    /// Iterate over tab stops in declaration order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, TabStop> {
        self.as_slice().iter()
    }

    /// Append a tab stop, returning the value when the collection is full.
    ///
    /// # Errors
    /// Returns the tab stop as an error when the table is already full.
    pub fn push(&mut self, tab: TabStop) -> Result<(), TabStop> {
        let index = self.len();
        let Some(slot) = self.entries.get_mut(index) else {
            return Err(tab);
        };
        *slot = tab;
        self.len += 1;
        Ok(())
    }

    /// Remove all custom tab stops.
    #[inline]
    pub fn clear(&mut self) {
        self.len = 0;
    }
}

impl Default for TabStops {
    fn default() -> Self {
        Self::new()
    }
}

impl AsRef<[TabStop]> for TabStops {
    fn as_ref(&self) -> &[TabStop] {
        self.as_slice()
    }
}

impl<'a> IntoIterator for &'a TabStops {
    type Item = &'a TabStop;
    type IntoIter = std::slice::Iter<'a, TabStop>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

impl TabStop {
    /// Create a new tab stop
    #[inline]
    #[must_use]
    pub fn new(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::default(),
            leader: TabLeader::default(),
        }
    }

    /// Create a left-aligned tab stop
    #[inline]
    #[must_use]
    pub fn left(position: i32) -> Self {
        Self::new(position)
    }

    /// Create a right-aligned tab stop
    #[inline]
    #[must_use]
    pub fn right(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Right,
            leader: TabLeader::default(),
        }
    }

    /// Create a centered tab stop
    #[inline]
    #[must_use]
    pub fn center(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Center,
            leader: TabLeader::default(),
        }
    }

    /// Create a decimal tab stop
    #[inline]
    #[must_use]
    pub fn decimal(position: i32) -> Self {
        Self {
            position,
            alignment: TabAlignment::Decimal,
            leader: TabLeader::default(),
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    #![allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "test loop indices stay far below the i32 range"
    )]
    use super::*;

    #[test]
    fn tab_stops_reject_overflow_without_mutating_contents() {
        let mut tabs = TabStops::new();
        for position in 0..MAX_PARAGRAPH_TAB_STOPS {
            assert_eq!(tabs.push(TabStop::new(position as i32)), Ok(()));
        }
        let rejected = TabStop::new(10_000);
        assert_eq!(tabs.push(rejected), Err(rejected));
        assert_eq!(tabs.len(), MAX_PARAGRAPH_TAB_STOPS);
        assert_eq!(tabs.as_slice().last().unwrap().position, 63);
    }

    #[test]
    fn tab_stop_access_is_total_for_an_invalid_private_length() {
        let mut tabs = TabStops::new();
        tabs.len = u8::MAX;
        assert!(tabs.as_slice().is_empty());
        let rejected = TabStop::new(720);
        assert_eq!(tabs.push(rejected), Err(rejected));
    }
}
