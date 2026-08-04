//! Closed DrawingML preset-geometry domains.
//!
//! The token sets in this module are the complete `ST_ShapeType` and
//! `ST_TextShapeType` enumerations from `dml-main.xsd` in both schema archives:
//! `3rdparty/ECMA-376-1_5th_edition_december_2016.zip` contains the nested
//! `OfficeOpenXML-XMLSchema-Strict.zip`, and
//! `3rdparty/ECMA-376-4_5th_edition_december_2016.zip` contains the nested
//! `OfficeOpenXML-XMLSchema-Transitional.zip`. Unknown values are rejected
//! instead of entering the semantic model as arbitrary strings.

use std::fmt;
use std::str::FromStr;

/// The closed DrawingML token domain that failed validation.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[repr(u8)]
enum Domain {
    Shape,
    TextShape,
}

impl Domain {
    const fn schema_name(self) -> &'static str {
        match self {
            Self::Shape => "ST_ShapeType",
            Self::TextShape => "ST_TextShapeType",
        }
    }
}

/// Error returned when a token is outside a closed DrawingML preset domain.
///
/// The error does not copy the rejected input. Callers that need to report it
/// can retain or borrow the token they supplied to [`FromStr`] or [`TryFrom`].
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct TokenError {
    domain: Domain,
}

impl TokenError {
    const fn new(domain: Domain) -> Self {
        Self { domain }
    }

    /// Return the ECMA-376 simple-type name whose validation failed.
    pub const fn domain(self) -> &'static str {
        self.domain.schema_name()
    }
}

impl fmt::Display for TokenError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "invalid {} token", self.domain.schema_name())
    }
}

impl std::error::Error for TokenError {}

macro_rules! token_enum {
    (
        $(#[$enum_meta:meta])*
        $name:ident, $domain:expr, $len:literal;
        $($(#[$variant_meta:meta])* $variant:ident => $token:literal),+ $(,)?
    ) => {
        $(#[$enum_meta])*
        #[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
        #[repr(u8)]
        pub enum $name {
            $(
                #[doc = concat!("`", $token, "`.")]
                $(#[$variant_meta])*
                $variant,
            )+
        }

        impl $name {
            /// Every value in schema order.
            pub const ALL: [Self; $len] = [$(Self::$variant),+];

            /// Return the exact ECMA-376 token.
            pub const fn token(self) -> &'static str {
                match self {
                    $(Self::$variant => $token,)+
                }
            }
        }

        impl FromStr for $name {
            type Err = TokenError;

            fn from_str(token: &str) -> Result<Self, Self::Err> {
                match token {
                    $($token => Ok(Self::$variant),)+
                    _ => Err(TokenError::new($domain)),
                }
            }
        }

        impl TryFrom<&str> for $name {
            type Error = TokenError;

            fn try_from(token: &str) -> Result<Self, Self::Error> {
                token.parse()
            }
        }

        impl fmt::Display for $name {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str(self.token())
            }
        }
    };
}

token_enum! {
    /// Preset shape geometry (`a:prstGeom@prst`, `ST_ShapeType`).
    Preset, Domain::Shape, 187;
    Line => "line",
    LineInv => "lineInv",
    Triangle => "triangle",
    RightTriangle => "rtTriangle",
    Rect => "rect",
    Diamond => "diamond",
    Parallelogram => "parallelogram",
    Trapezoid => "trapezoid",
    NonIsoscelesTrapezoid => "nonIsoscelesTrapezoid",
    Pentagon => "pentagon",
    Hexagon => "hexagon",
    Heptagon => "heptagon",
    Octagon => "octagon",
    Decagon => "decagon",
    Dodecagon => "dodecagon",
    Star4 => "star4",
    Star5 => "star5",
    Star6 => "star6",
    Star7 => "star7",
    Star8 => "star8",
    Star10 => "star10",
    Star12 => "star12",
    Star16 => "star16",
    Star24 => "star24",
    Star32 => "star32",
    RoundRect => "roundRect",
    Round1Rect => "round1Rect",
    Round2SameRect => "round2SameRect",
    Round2DiagRect => "round2DiagRect",
    SnipRoundRect => "snipRoundRect",
    Snip1Rect => "snip1Rect",
    Snip2SameRect => "snip2SameRect",
    Snip2DiagRect => "snip2DiagRect",
    Plaque => "plaque",
    Ellipse => "ellipse",
    Teardrop => "teardrop",
    HomePlate => "homePlate",
    Chevron => "chevron",
    PieWedge => "pieWedge",
    Pie => "pie",
    BlockArc => "blockArc",
    Donut => "donut",
    NoSmoking => "noSmoking",
    RightArrow => "rightArrow",
    LeftArrow => "leftArrow",
    UpArrow => "upArrow",
    DownArrow => "downArrow",
    StripedRightArrow => "stripedRightArrow",
    NotchedRightArrow => "notchedRightArrow",
    BentUpArrow => "bentUpArrow",
    LeftRightArrow => "leftRightArrow",
    UpDownArrow => "upDownArrow",
    LeftUpArrow => "leftUpArrow",
    LeftRightUpArrow => "leftRightUpArrow",
    QuadArrow => "quadArrow",
    LeftArrowCallout => "leftArrowCallout",
    RightArrowCallout => "rightArrowCallout",
    UpArrowCallout => "upArrowCallout",
    DownArrowCallout => "downArrowCallout",
    LeftRightArrowCallout => "leftRightArrowCallout",
    UpDownArrowCallout => "upDownArrowCallout",
    QuadArrowCallout => "quadArrowCallout",
    BentArrow => "bentArrow",
    UTurnArrow => "uturnArrow",
    CircularArrow => "circularArrow",
    LeftCircularArrow => "leftCircularArrow",
    LeftRightCircularArrow => "leftRightCircularArrow",
    CurvedRightArrow => "curvedRightArrow",
    CurvedLeftArrow => "curvedLeftArrow",
    CurvedUpArrow => "curvedUpArrow",
    CurvedDownArrow => "curvedDownArrow",
    SwooshArrow => "swooshArrow",
    Cube => "cube",
    Can => "can",
    LightningBolt => "lightningBolt",
    Heart => "heart",
    Sun => "sun",
    Moon => "moon",
    SmileyFace => "smileyFace",
    IrregularSeal1 => "irregularSeal1",
    IrregularSeal2 => "irregularSeal2",
    FoldedCorner => "foldedCorner",
    Bevel => "bevel",
    Frame => "frame",
    HalfFrame => "halfFrame",
    Corner => "corner",
    DiagStripe => "diagStripe",
    Chord => "chord",
    Arc => "arc",
    LeftBracket => "leftBracket",
    RightBracket => "rightBracket",
    LeftBrace => "leftBrace",
    RightBrace => "rightBrace",
    BracketPair => "bracketPair",
    BracePair => "bracePair",
    StraightConnector1 => "straightConnector1",
    BentConnector2 => "bentConnector2",
    BentConnector3 => "bentConnector3",
    BentConnector4 => "bentConnector4",
    BentConnector5 => "bentConnector5",
    CurvedConnector2 => "curvedConnector2",
    CurvedConnector3 => "curvedConnector3",
    CurvedConnector4 => "curvedConnector4",
    CurvedConnector5 => "curvedConnector5",
    Callout1 => "callout1",
    Callout2 => "callout2",
    Callout3 => "callout3",
    AccentCallout1 => "accentCallout1",
    AccentCallout2 => "accentCallout2",
    AccentCallout3 => "accentCallout3",
    BorderCallout1 => "borderCallout1",
    BorderCallout2 => "borderCallout2",
    BorderCallout3 => "borderCallout3",
    AccentBorderCallout1 => "accentBorderCallout1",
    AccentBorderCallout2 => "accentBorderCallout2",
    AccentBorderCallout3 => "accentBorderCallout3",
    WedgeRectCallout => "wedgeRectCallout",
    WedgeRoundRectCallout => "wedgeRoundRectCallout",
    WedgeEllipseCallout => "wedgeEllipseCallout",
    CloudCallout => "cloudCallout",
    Cloud => "cloud",
    Ribbon => "ribbon",
    Ribbon2 => "ribbon2",
    EllipseRibbon => "ellipseRibbon",
    EllipseRibbon2 => "ellipseRibbon2",
    LeftRightRibbon => "leftRightRibbon",
    VerticalScroll => "verticalScroll",
    HorizontalScroll => "horizontalScroll",
    Wave => "wave",
    DoubleWave => "doubleWave",
    Plus => "plus",
    FlowChartProcess => "flowChartProcess",
    FlowChartDecision => "flowChartDecision",
    FlowChartInputOutput => "flowChartInputOutput",
    FlowChartPredefinedProcess => "flowChartPredefinedProcess",
    FlowChartInternalStorage => "flowChartInternalStorage",
    FlowChartDocument => "flowChartDocument",
    FlowChartMultiDocument => "flowChartMultidocument",
    FlowChartTerminator => "flowChartTerminator",
    FlowChartPreparation => "flowChartPreparation",
    FlowChartManualInput => "flowChartManualInput",
    FlowChartManualOperation => "flowChartManualOperation",
    FlowChartConnector => "flowChartConnector",
    FlowChartPunchedCard => "flowChartPunchedCard",
    FlowChartPunchedTape => "flowChartPunchedTape",
    FlowChartSummingJunction => "flowChartSummingJunction",
    FlowChartOr => "flowChartOr",
    FlowChartCollate => "flowChartCollate",
    FlowChartSort => "flowChartSort",
    FlowChartExtract => "flowChartExtract",
    FlowChartMerge => "flowChartMerge",
    FlowChartOfflineStorage => "flowChartOfflineStorage",
    FlowChartOnlineStorage => "flowChartOnlineStorage",
    FlowChartMagneticTape => "flowChartMagneticTape",
    FlowChartMagneticDisk => "flowChartMagneticDisk",
    FlowChartMagneticDrum => "flowChartMagneticDrum",
    FlowChartDisplay => "flowChartDisplay",
    FlowChartDelay => "flowChartDelay",
    FlowChartAlternateProcess => "flowChartAlternateProcess",
    FlowChartOffPageConnector => "flowChartOffpageConnector",
    ActionButtonBlank => "actionButtonBlank",
    ActionButtonHome => "actionButtonHome",
    ActionButtonHelp => "actionButtonHelp",
    ActionButtonInformation => "actionButtonInformation",
    ActionButtonForwardNext => "actionButtonForwardNext",
    ActionButtonBackPrevious => "actionButtonBackPrevious",
    ActionButtonEnd => "actionButtonEnd",
    ActionButtonBeginning => "actionButtonBeginning",
    ActionButtonReturn => "actionButtonReturn",
    ActionButtonDocument => "actionButtonDocument",
    ActionButtonSound => "actionButtonSound",
    ActionButtonMovie => "actionButtonMovie",
    Gear6 => "gear6",
    Gear9 => "gear9",
    Funnel => "funnel",
    MathPlus => "mathPlus",
    MathMinus => "mathMinus",
    MathMultiply => "mathMultiply",
    MathDivide => "mathDivide",
    MathEqual => "mathEqual",
    MathNotEqual => "mathNotEqual",
    CornerTabs => "cornerTabs",
    SquareTabs => "squareTabs",
    PlaqueTabs => "plaqueTabs",
    ChartX => "chartX",
    ChartStar => "chartStar",
    ChartPlus => "chartPlus",
}

token_enum! {
    /// Preset text warp (`a:prstTxWarp@prst`, `ST_TextShapeType`).
    #[derive(Default)]
    TextPreset, Domain::TextShape, 41;
    #[default]
    NoShape => "textNoShape",
    Plain => "textPlain",
    Stop => "textStop",
    Triangle => "textTriangle",
    TriangleInverted => "textTriangleInverted",
    Chevron => "textChevron",
    ChevronInverted => "textChevronInverted",
    RingInside => "textRingInside",
    RingOutside => "textRingOutside",
    ArchUp => "textArchUp",
    ArchDown => "textArchDown",
    Circle => "textCircle",
    Button => "textButton",
    ArchUpPour => "textArchUpPour",
    ArchDownPour => "textArchDownPour",
    CirclePour => "textCirclePour",
    ButtonPour => "textButtonPour",
    CurveUp => "textCurveUp",
    CurveDown => "textCurveDown",
    CanUp => "textCanUp",
    CanDown => "textCanDown",
    Wave1 => "textWave1",
    Wave2 => "textWave2",
    DoubleWave1 => "textDoubleWave1",
    Wave4 => "textWave4",
    Inflate => "textInflate",
    Deflate => "textDeflate",
    InflateBottom => "textInflateBottom",
    DeflateBottom => "textDeflateBottom",
    InflateTop => "textInflateTop",
    DeflateTop => "textDeflateTop",
    DeflateInflate => "textDeflateInflate",
    DeflateInflateDeflate => "textDeflateInflateDeflate",
    FadeRight => "textFadeRight",
    FadeLeft => "textFadeLeft",
    FadeUp => "textFadeUp",
    FadeDown => "textFadeDown",
    SlantUp => "textSlantUp",
    SlantDown => "textSlantDown",
    CascadeUp => "textCascadeUp",
    CascadeDown => "textCascadeDown",
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;
    use std::mem::size_of;

    fn assert_closed_round_trip<T>(all: &[T], expected_len: usize, domain: &'static str)
    where
        T: Copy + Eq + std::hash::Hash + fmt::Debug + fmt::Display + FromStr<Err = TokenError>,
    {
        assert_eq!(all.len(), expected_len);
        let mut values = HashSet::with_capacity(all.len());
        let mut tokens = HashSet::with_capacity(all.len());
        for &value in all {
            let token = value.to_string();
            assert!(values.insert(value), "duplicate value for {token}");
            assert!(tokens.insert(token.clone()), "duplicate token {token}");
            assert_eq!(token.parse::<T>(), Ok(value));
        }
        let error = "notInTheSchema".parse::<T>().unwrap_err();
        assert_eq!(error.domain(), domain);
    }

    #[test]
    fn shape_tokens_are_complete_unique_and_round_trip() {
        assert_closed_round_trip(&Preset::ALL, 187, "ST_ShapeType");
        assert_eq!(Preset::try_from("rect"), Ok(Preset::Rect));
        assert!(Preset::try_from(" rect ").is_err());
        assert!(Preset::try_from("textBox").is_err());
        assert!(Preset::try_from("cross").is_err());
    }

    #[test]
    fn text_tokens_are_complete_unique_and_round_trip() {
        assert_closed_round_trip(&TextPreset::ALL, 41, "ST_TextShapeType");
        assert_eq!(TextPreset::try_from("textButton"), Ok(TextPreset::Button));
        assert!(TextPreset::try_from("\ttextButton\n").is_err());
        assert!(TextPreset::try_from("textButtonUp").is_err());
        assert!(TextPreset::try_from("textButtonDown").is_err());
    }

    #[test]
    fn domains_are_compact_and_allocation_free() {
        assert_eq!(size_of::<Preset>(), 1);
        assert_eq!(size_of::<TextPreset>(), 1);
        assert_eq!(size_of::<TokenError>(), 1);
    }
}
