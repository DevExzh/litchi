//! Typed RTF 1.9.1 section page borders.

use crate::{RtfError, RtfResult};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageBorderStyle {
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
}

impl PageBorderStyle {
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
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageBorder {
    pub style: PageBorderStyle,
    /// Decorative border-art identifier, 1 through 165, mutually exclusive with a line style.
    pub art: Option<u8>,
    /// Width in twips, limited by RTF 1.9.1 to 75.
    pub width: u8,
    pub color_ref: u16,
    /// Distance from the page/text reference edge in twips.
    pub space: u16,
    pub shadow: bool,
    pub frame: bool,
}

impl Default for PageBorder {
    fn default() -> Self {
        Self {
            style: PageBorderStyle::None,
            art: None,
            width: 10,
            color_ref: 0,
            space: 0,
            shadow: false,
            frame: false,
        }
    }
}

impl PageBorder {
    pub fn validate(&self) -> RtfResult<()> {
        if self.width > 75 {
            return invalid("RTF page-border width must be in 0..=75 twips");
        }
        if self.space > 1_440 {
            return invalid("RTF page-border spacing must be in 0..=1440 twips");
        }
        if self.art.is_some_and(|value| !(1..=165).contains(&value)) {
            return invalid("RTF page-border art must be in 1..=165");
        }
        if self.art.is_some() && self.style != PageBorderStyle::None {
            return invalid("RTF page-border art and line style are mutually exclusive");
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageBorderAppliesTo {
    #[default]
    AllSectionPages,
    FirstSectionPage,
    AllExceptFirstSectionPage,
    WholeDocument,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageBorderDepth {
    #[default]
    InFront,
    Behind,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageBorderOffset {
    #[default]
    Text,
    Page,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageBorderSide {
    Top,
    Left,
    Bottom,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PageBorders {
    pub top: Option<PageBorder>,
    pub left: Option<PageBorder>,
    pub bottom: Option<PageBorder>,
    pub right: Option<PageBorder>,
    pub applies_to: PageBorderAppliesTo,
    pub depth: PageBorderDepth,
    pub offset: PageBorderOffset,
    pub surround_header: bool,
    pub surround_footer: bool,
    pub snap_to_text_borders: bool,
}

impl PageBorders {
    pub fn is_empty(&self) -> bool {
        [self.top, self.left, self.bottom, self.right]
            .iter()
            .all(Option::is_none)
            && self.option_value() == 0
            && !self.surround_header
            && !self.surround_footer
            && !self.snap_to_text_borders
    }
    pub fn get(&self, side: PageBorderSide) -> Option<&PageBorder> {
        match side {
            PageBorderSide::Top => self.top.as_ref(),
            PageBorderSide::Left => self.left.as_ref(),
            PageBorderSide::Bottom => self.bottom.as_ref(),
            PageBorderSide::Right => self.right.as_ref(),
        }
    }
    pub fn set(&mut self, side: PageBorderSide, border: PageBorder) {
        *match side {
            PageBorderSide::Top => &mut self.top,
            PageBorderSide::Left => &mut self.left,
            PageBorderSide::Bottom => &mut self.bottom,
            PageBorderSide::Right => &mut self.right,
        } = Some(border);
    }
    pub fn option_value(&self) -> i32 {
        let applies = match self.applies_to {
            PageBorderAppliesTo::AllSectionPages => 0,
            PageBorderAppliesTo::FirstSectionPage => 1,
            PageBorderAppliesTo::AllExceptFirstSectionPage => 2,
            PageBorderAppliesTo::WholeDocument => 3,
        };
        applies
            | match self.depth {
                PageBorderDepth::InFront => 0,
                PageBorderDepth::Behind => 8,
            }
            | match self.offset {
                PageBorderOffset::Text => 0,
                PageBorderOffset::Page => 32,
            }
    }
    pub fn set_option_value(&mut self, value: i32) -> RtfResult<()> {
        if !(0..=255).contains(&value) || value & !0x2b != 0 {
            return invalid("RTF pgbrdropt contains reserved or out-of-range bits");
        }
        self.applies_to = match value & 7 {
            0 => PageBorderAppliesTo::AllSectionPages,
            1 => PageBorderAppliesTo::FirstSectionPage,
            2 => PageBorderAppliesTo::AllExceptFirstSectionPage,
            3 => PageBorderAppliesTo::WholeDocument,
            _ => return invalid("RTF pgbrdropt has a reserved page-applicability value"),
        };
        self.depth = if value & 8 == 0 {
            PageBorderDepth::InFront
        } else {
            PageBorderDepth::Behind
        };
        self.offset = if value & 32 == 0 {
            PageBorderOffset::Text
        } else {
            PageBorderOffset::Page
        };
        Ok(())
    }
    pub fn validate(&self) -> RtfResult<()> {
        for border in [self.top, self.left, self.bottom, self.right]
            .into_iter()
            .flatten()
        {
            border.validate()?;
        }
        Ok(())
    }
}

fn invalid<T>(message: impl Into<String>) -> RtfResult<T> {
    Err(RtfError::MalformedDocument(message.into()))
}
