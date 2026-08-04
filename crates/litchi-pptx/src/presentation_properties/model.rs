//! Typed PresentationML property values.

/// Relationship projection used by the HTML publishing property.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlTarget {
    pub relationship_id: String,
    pub target: Option<String>,
    pub relationship_type: Option<String>,
    pub external: Option<bool>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BrowserSupport {
    V3,
    V4,
    V3V4,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebScreenSize {
    S544x376,
    S640x480,
    S720x512,
    S800x600,
    S1024x768,
    S1152x882,
    S1152x900,
    S1280x1024,
    S1600x1200,
    S1800x1400,
    S1920x1200,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebColor {
    None,
    Browser,
    PresentationText,
    PresentationAccent,
    WhiteTextOnBlack,
    BlackTextOnWhite,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintOutput {
    Slides,
    Handouts1,
    Handouts2,
    Handouts3,
    Handouts4,
    Handouts6,
    Handouts9,
    Notes,
    Outline,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintColorMode {
    BlackWhite,
    Gray,
    Color,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SlideSelection {
    All,
    Range { start: u32, end: u32 },
    CustomShow(u32),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShowMode {
    Present,
    Browse { show_scrollbar: Option<bool> },
    Kiosk { restart: Option<u32> },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

/// DrawingML color plus its bounded source fragment.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Color {
    pub kind: ColorKind,
    pub attributes: Vec<(String, String)>,
    pub xml: Vec<u8>,
}

/// Extension payload preserved without interpreting unknown content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueExtension {
    pub uri: String,
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Extension {
    DiscardImageEditData(bool),
    DefaultImageDpi(u32),
    ChartTrackingReferenceBased(bool),
    Unknown(OpaqueExtension),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShowExtension {
    BrowseMode { show_status: Option<bool> },
    LaserColor(Color),
    ShowMediaControls(bool),
    Unknown(OpaqueExtension),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlPublish {
    pub show_speaker_notes: Option<bool>,
    pub browser: Option<BrowserSupport>,
    pub target: HtmlTarget,
    pub slides: SlideSelection,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Web {
    pub show_animation: Option<bool>,
    pub resize_graphics: Option<bool>,
    pub allow_png: Option<bool>,
    pub rely_on_vml: Option<bool>,
    pub organize_in_folders: Option<bool>,
    pub use_long_filenames: Option<bool>,
    pub image_size: Option<WebScreenSize>,
    pub encoding: Option<String>,
    pub color: Option<WebColor>,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Print {
    pub output: Option<PrintOutput>,
    pub color_mode: Option<PrintColorMode>,
    pub hidden_slides: Option<bool>,
    pub scale_to_fit_paper: Option<bool>,
    pub frame_slides: Option<bool>,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Show {
    pub loop_show: Option<bool>,
    pub show_narration: Option<bool>,
    pub show_animation: Option<bool>,
    pub use_timings: Option<bool>,
    pub mode: Option<ShowMode>,
    pub slides: Option<SlideSelection>,
    pub pen_color: Option<Color>,
    pub extensions: Vec<ShowExtension>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Properties {
    pub html_publish: Option<HtmlPublish>,
    pub web: Option<Web>,
    pub print: Option<Print>,
    pub show: Option<Show>,
    pub recent_colors: Vec<Color>,
    pub extensions: Vec<Extension>,
}

// Historical names remain source-compatible at the public facade.
pub type InertHtmlTarget = HtmlTarget;
pub type PresentationColor = Color;
pub type OpaquePresentationExtension = OpaqueExtension;
pub type PresentationPropertyExtension = Extension;
pub type SlideShowExtension = ShowExtension;
pub type HtmlPublishProperties = HtmlPublish;
pub type WebProperties = Web;
pub type PrintProperties = Print;
pub type ShowProperties = Show;
pub type PresentationProperties = Properties;
