//! Contextual values for DrawingML chart-style companion XML.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Document {
    pub(crate) info: Info,
    pub(crate) xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorDocument {
    pub(crate) info: ColorInfo,
    pub(crate) xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    pub id: Option<u32>,
    pub entries: Vec<Entry>,
    pub marker_layout: Option<MarkerLayout>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EntryKind {
    AxisTitle,
    CategoryAxis,
    ChartArea,
    DataLabel,
    DataLabelCallout,
    DataPoint,
    DataPoint3D,
    DataPointLine,
    DataPointMarker,
    DataPointWireframe,
    DataTable,
    DownBar,
    DropLine,
    ErrorBar,
    Floor,
    GridlineMajor,
    GridlineMinor,
    HiLoLine,
    LeaderLine,
    Legend,
    PlotArea,
    PlotArea3D,
    SeriesAxis,
    SeriesLine,
    Title,
    Trendline,
    TrendlineLabel,
    UpBar,
    ValueAxis,
    Wall,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    pub kind: EntryKind,
    pub modifiers: Vec<String>,
    pub line_reference: Reference,
    /// The validated XML Schema double lexical value; defaults to `1.0`.
    pub line_width_scale: String,
    pub fill_reference: Reference,
    pub effect_reference: Reference,
    pub font_reference: FontReference,
    pub shape_properties: Option<Payload>,
    pub default_run_properties: Option<Payload>,
    pub body_properties: Option<Payload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub index: u32,
    pub modifiers: Vec<String>,
    pub color: Option<Color>,
    pub style_color: Option<ColorValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontIndex {
    Major,
    Minor,
    None,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontReference {
    pub index: FontIndex,
    pub modifiers: Vec<String>,
    pub color: Option<Color>,
    pub style_color: Option<ColorValue>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorValue {
    pub raw: Option<String>,
    pub index: Option<u32>,
    pub automatic: bool,
    pub transforms: Vec<Transform>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerSymbol {
    Circle,
    Dash,
    Diamond,
    Dot,
    Plus,
    Square,
    Star,
    Triangle,
    X,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MarkerLayout {
    pub symbol: Option<MarkerSymbol>,
    pub size: Option<u8>,
}

/// A bounded summary of an inert DrawingML formatting subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Payload {
    pub child_elements: usize,
    pub attributes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Color {
    pub kind: ColorKind,
    /// Primary color value where the color model has one; component models use `components`.
    pub value: Option<String>,
    pub components: Vec<(String, String)>,
    pub transforms: Vec<Transform>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransformKind {
    Tint,
    Shade,
    Complement,
    Inverse,
    Grayscale,
    Alpha,
    AlphaOffset,
    AlphaModulation,
    Hue,
    HueOffset,
    HueModulation,
    Saturation,
    SaturationOffset,
    SaturationModulation,
    Luminance,
    LuminanceOffset,
    LuminanceModulation,
    Red,
    RedOffset,
    RedModulation,
    Green,
    GreenOffset,
    GreenModulation,
    Blue,
    BlueOffset,
    BlueModulation,
    Gamma,
    InverseGamma,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Transform {
    pub kind: TransformKind,
    /// Preserved integer lexical value for transforms that take `val`.
    pub value: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorMethod {
    Cycle,
    WithinLinear,
    AcrossLinear,
    WithinLinearReversed,
    AcrossLinearReversed,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorInfo {
    pub method: String,
    /// Unknown extension methods have the specified effective behavior `Cycle`.
    pub effective_method: ColorMethod,
    pub id: Option<u32>,
    pub colors: Vec<Color>,
    pub variations: Vec<Variation>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variation {
    pub transforms: Vec<Transform>,
}
