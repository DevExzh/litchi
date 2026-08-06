//! Package-independent handout-master values and builders.

use std::str::FromStr;

/// Number of slides per handout page.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Layout {
    #[default]
    OneSlide,
    TwoSlides,
    ThreeSlides,
    FourSlides,
    SixSlides,
    NineSlides,
    Outline,
}

impl Layout {
    /// Get the layout type string for handout-master XML.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::OneSlide => "handout1",
            Self::TwoSlides => "handout2",
            Self::ThreeSlides => "handout3",
            Self::FourSlides => "handout4",
            Self::SixSlides => "handout6",
            Self::NineSlides => "handout9",
            Self::Outline => "handoutOutline",
        }
    }

    /// Get the `ST_PrintWhat` value for `presProps.xml`.
    pub fn print_what(&self) -> &'static str {
        match self {
            Self::OneSlide => "handouts1",
            Self::TwoSlides => "handouts2",
            Self::ThreeSlides => "handouts3",
            Self::FourSlides => "handouts4",
            Self::SixSlides => "handouts6",
            Self::NineSlides => "handouts9",
            Self::Outline => "outline",
        }
    }
}

impl FromStr for Layout {
    type Err = ();

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Ok(match value {
            "handout1" => Self::OneSlide,
            "handout2" => Self::TwoSlides,
            "handout3" => Self::ThreeSlides,
            "handout4" => Self::FourSlides,
            "handout6" => Self::SixSlides,
            "handout9" => Self::NineSlides,
            "handoutOutline" => Self::Outline,
            _ => Self::OneSlide,
        })
    }
}

/// Header/footer configuration for handouts.
#[derive(Debug, Clone, Default)]
pub struct HeaderFooter {
    pub show_header: bool,
    pub header_text: Option<String>,
    pub show_footer: bool,
    pub footer_text: Option<String>,
    pub show_slide_number: bool,
    pub show_date_time: bool,
    pub date_time_text: Option<String>,
    pub auto_date: bool,
}

/// Handout master for a presentation.
#[derive(Debug, Clone)]
pub struct Master {
    pub layout: Layout,
    pub header_footer: HeaderFooter,
    pub background_color: Option<String>,
    pub show_slide_images: bool,
}

impl Default for Master {
    fn default() -> Self {
        Self {
            layout: Layout::default(),
            header_footer: HeaderFooter::default(),
            background_color: None,
            show_slide_images: true,
        }
    }
}

impl Master {
    /// Create a new handout master with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the handout layout.
    pub fn with_layout(mut self, layout: Layout) -> Self {
        self.layout = layout;
        self
    }

    /// Set the header text.
    pub fn with_header(mut self, text: impl Into<String>) -> Self {
        self.header_footer.show_header = true;
        self.header_footer.header_text = Some(text.into());
        self
    }

    /// Set the footer text.
    pub fn with_footer(mut self, text: impl Into<String>) -> Self {
        self.header_footer.show_footer = true;
        self.header_footer.footer_text = Some(text.into());
        self
    }

    /// Enable slide numbers.
    pub fn with_slide_numbers(mut self) -> Self {
        self.header_footer.show_slide_number = true;
        self
    }

    /// Enable automatic date/time display.
    pub fn with_date_time(mut self) -> Self {
        self.header_footer.show_date_time = true;
        self.header_footer.auto_date = true;
        self
    }

    /// Set a fixed date text and disable automatic date.
    pub fn with_fixed_date(mut self, date_text: impl Into<String>) -> Self {
        self.header_footer.show_date_time = true;
        self.header_footer.auto_date = false;
        self.header_footer.date_time_text = Some(date_text.into());
        self
    }

    /// Set the background color as an RGB lexical value without `#`.
    pub fn with_background_color(mut self, color: impl Into<String>) -> Self {
        self.background_color = Some(color.into());
        self
    }
}
