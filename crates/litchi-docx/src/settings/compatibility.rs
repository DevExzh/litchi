#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::format_push_string,
    reason = "serialization preserves the established byte-emission path"
)]
use std::error::Error as StdError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

use super::support::escape_attribute;

/// `w:compatSetting` name identifying the targeted Word compatibility mode.
pub const COMPATIBILITY_MODE_SETTING_NAME: &str = "compatibilityMode";
/// `w:compatSetting` URI under which Word stores its compatibility settings.
pub const COMPATIBILITY_SETTING_URI: &str = "http://schemas.microsoft.com/office/word";

/// Error returned for a token outside the closed `WordprocessingML`
/// compatibility-flag domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseCompatFlagError;

impl Display for ParseCompatFlagError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML compatibility flag")
    }
}

impl StdError for ParseCompatFlagError {}

macro_rules! define_compat_flags {
    ($($variant:ident => $token:literal),+ $(,)?) => {
        /// A closed `CT_Compat` on/off flag from the Strict or Transitional
        /// `WordprocessingML` schema.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[repr(u8)]
        pub enum CompatFlag {
            $(
                #[doc = concat!("`", $token, "`.")]
                $variant,
            )+
        }

        impl CompatFlag {
            /// Every compatibility flag accepted by either checked-in schema.
            pub const ALL: &'static [Self] = &[$(Self::$variant),+];

            /// Return the exact `WordprocessingML` element-local-name token.
            #[must_use]
            pub const fn as_str(self) -> &'static str {
                match self {
                    $(Self::$variant => $token,)+
                }
            }
        }

        impl FromStr for CompatFlag {
            type Err = ParseCompatFlagError;

            fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
                match value {
                    $($token => Ok(Self::$variant),)+
                    _ => Err(ParseCompatFlagError),
                }
            }
        }

        impl Display for CompatFlag {
            fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
                formatter.write_str(self.as_str())
            }
        }
    };
}

define_compat_flags! {
    UseSingleBorderForContiguousCells => "useSingleBorderforContiguousCells",
    WpJustification => "wpJustification",
    NoTabHangIndent => "noTabHangInd",
    NoLeading => "noLeading",
    SpaceForUnderline => "spaceForUL",
    NoColumnBalance => "noColumnBalance",
    BalanceSingleByteDoubleByteWidth => "balanceSingleByteDoubleByteWidth",
    NoExtraLineSpacing => "noExtraLineSpacing",
    DoNotLeaveBackslashAlone => "doNotLeaveBackslashAlone",
    UnderlineTrailingSpace => "ulTrailSpace",
    DoNotExpandShiftReturn => "doNotExpandShiftReturn",
    SpacingInWholePoints => "spacingInWholePoints",
    LineWrapLikeWord6 => "lineWrapLikeWord6",
    PrintBodyTextBeforeHeader => "printBodyTextBeforeHeader",
    PrintColorBlack => "printColBlack",
    WpSpaceWidth => "wpSpaceWidth",
    ShowBreaksInFrames => "showBreaksInFrames",
    SubstituteFontBySize => "subFontBySize",
    SuppressBottomSpacing => "suppressBottomSpacing",
    SuppressTopSpacing => "suppressTopSpacing",
    SuppressSpacingAtTopOfPage => "suppressSpacingAtTopOfPage",
    SuppressTopSpacingWp => "suppressTopSpacingWP",
    SuppressSpaceBeforeAfterPageBreak => "suppressSpBfAfterPgBrk",
    SwapBordersFacingPages => "swapBordersFacingPages",
    ConvertMailMergeEscapes => "convMailMergeEsc",
    TruncateFontHeightsLikeWp6 => "truncateFontHeightsLikeWP6",
    MwSmallCaps => "mwSmallCaps",
    UsePrinterMetrics => "usePrinterMetrics",
    DoNotSuppressParagraphBorders => "doNotSuppressParagraphBorders",
    WrapTrailingSpaces => "wrapTrailSpaces",
    FootnoteLayoutLikeWord8 => "footnoteLayoutLikeWW8",
    ShapeLayoutLikeWord8 => "shapeLayoutLikeWW8",
    AlignTablesRowByRow => "alignTablesRowByRow",
    ForgetLastTabAlignment => "forgetLastTabAlignment",
    AdjustLineHeightInTable => "adjustLineHeightInTable",
    AutoSpaceLikeWord95 => "autoSpaceLikeWord95",
    NoSpaceRaiseLower => "noSpaceRaiseLower",
    DoNotUseHtmlParagraphAutoSpacing => "doNotUseHTMLParagraphAutoSpacing",
    LayoutRawTableWidth => "layoutRawTableWidth",
    LayoutTableRowsApart => "layoutTableRowsApart",
    UseWord97LineBreakRules => "useWord97LineBreakRules",
    DoNotBreakWrappedTables => "doNotBreakWrappedTables",
    DoNotSnapToGridInCell => "doNotSnapToGridInCell",
    SelectFieldWithFirstOrLastChar => "selectFldWithFirstOrLastChar",
    ApplyBreakingRules => "applyBreakingRules",
    DoNotWrapTextWithPunctuation => "doNotWrapTextWithPunct",
    DoNotUseEastAsianBreakRules => "doNotUseEastAsianBreakRules",
    UseWord2002TableStyleRules => "useWord2002TableStyleRules",
    GrowAutoFit => "growAutofit",
    UseFarEastLayout => "useFELayout",
    UseNormalStyleForList => "useNormalStyleForList",
    DoNotUseIndentAsNumberingTabStop => "doNotUseIndentAsNumberingTabStop",
    UseAlternateKinsokuLineBreakRules => "useAltKinsokuLineBreakRules",
    AllowSpaceOfSameStyleInTable => "allowSpaceOfSameStyleInTable",
    DoNotSuppressIndentation => "doNotSuppressIndentation",
    DoNotAutoFitConstrainedTables => "doNotAutofitConstrainedTables",
    AutoFitToFirstFixedWidthCell => "autofitToFirstFixedWidthCell",
    UnderlineTabInNumberedList => "underlineTabInNumList",
    DisplayHangulFixedWidth => "displayHangulFixedWidth",
    SplitPageBreakAndParagraphMark => "splitPgBreakAndParaMark",
    DoNotVerticallyAlignCellWithSpacing => "doNotVertAlignCellWithSp",
    DoNotBreakConstrainedForcedTable => "doNotBreakConstrainedForcedTable",
    DoNotVerticallyAlignInTextBox => "doNotVertAlignInTxbx",
    UseAnsiKerningPairs => "useAnsiKerningPairs",
    CachedColumnBalance => "cachedColBalance",
}

impl CompatFlag {
    /// Whether this flag is part of the Strict `WordprocessingML` domain.
    #[must_use]
    pub const fn is_strict(self) -> bool {
        matches!(
            self,
            Self::SpaceForUnderline
                | Self::BalanceSingleByteDoubleByteWidth
                | Self::DoNotLeaveBackslashAlone
                | Self::UnderlineTrailingSpace
                | Self::DoNotExpandShiftReturn
                | Self::AdjustLineHeightInTable
                | Self::ApplyBreakingRules
        )
    }
}

/// An on/off compatibility option from `w:compat`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CompatibilityOption {
    flag: CompatFlag,
    enabled: bool,
}

impl CompatibilityOption {
    /// Create a compatibility option.
    #[must_use]
    pub const fn new(flag: CompatFlag, enabled: bool) -> Self {
        Self { flag, enabled }
    }

    /// Return the schema-defined flag.
    #[inline]
    #[must_use]
    pub const fn flag(&self) -> CompatFlag {
        self.flag
    }

    /// Whether the option is enabled.
    #[inline]
    #[must_use]
    pub const fn is_enabled(&self) -> bool {
        self.enabled
    }

    /// Serialize the standalone compatibility flag element.
    #[must_use]
    pub fn to_xml(self, prefix: &str) -> String {
        if self.enabled {
            format!("<{prefix}:{}/>", self.flag.as_str())
        } else {
            format!("<{prefix}:{} {prefix}:val=\"off\"/>", self.flag.as_str())
        }
    }
}

/// A `w:compatSetting` name/URI/value triple from `w:compat`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompatibilitySetting {
    name: String,
    uri: String,
    value: String,
}

impl CompatibilitySetting {
    /// Create a compatibility-setting triple.
    pub fn new(name: impl Into<String>, uri: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            uri: uri.into(),
            value: value.into(),
        }
    }

    /// Return the setting name.
    #[inline]
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the URI scoping the setting name.
    #[inline]
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }

    /// Return the raw setting value.
    #[inline]
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Serialize a standalone `w:compatSetting` element.
    #[must_use]
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:compatSetting {prefix}:name=\"");
        escape_attribute(&mut xml, &self.name);
        xml.push_str(&format!("\" {prefix}:uri=\""));
        escape_attribute(&mut xml, &self.uri);
        xml.push_str(&format!("\" {prefix}:val=\""));
        escape_attribute(&mut xml, &self.value);
        xml.push_str("\"/>");
        xml
    }
}
