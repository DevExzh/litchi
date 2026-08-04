//! Typed WordprocessingML document-settings values and bounded XML codec.
//!
//! The package host owns relationship validation and settings-part orchestration;
//! this module owns the settings vocabulary that is entirely described by
//! `settings.xml`.  In particular, compatibility flags, note numbering,
//! protection, view, proofing, theme-font, and color-scheme values do not carry
//! package handles or relationship state.

use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::error::Error as StdError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

/// Transitional WordprocessingML namespace.
pub const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Strict WordprocessingML namespace.
pub const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

/// Maximum input accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum element count accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_NODES: usize = 250_000;
/// Maximum nesting depth accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_DEPTH: usize = 256;

/// Maximum number of Unicode scalar values accepted in a smart-tag namespace
/// URI, following the checked-in Office client limit.
pub const MAX_SMART_TAG_NAMESPACE_URI_CHARS: usize = 2083;
/// Maximum number of Unicode scalar values accepted in a smart-tag name,
/// following the checked-in Office client limit.
pub const MAX_SMART_TAG_NAME_CHARS: usize = 255;
/// Maximum number of Unicode scalar values accepted in a smart-tag download
/// URL, following the checked-in Office client limit.
pub const MAX_SMART_TAG_URL_CHARS: usize = 2083;

/// `w:compatSetting` name identifying the targeted Word compatibility mode.
pub const COMPATIBILITY_MODE_SETTING_NAME: &str = "compatibilityMode";
/// `w:compatSetting` URI under which Word stores its compatibility settings.
pub const COMPATIBILITY_SETTING_URI: &str = "http://schemas.microsoft.com/office/word";

/// A smart-tag vocabulary declaration from a WordprocessingML settings part.
///
/// The value is deliberately package-neutral: relationship resolution,
/// document matching, and settings-part orchestration remain in the host
/// package facade.  The three attributes are required by the host settings
/// vocabulary; an empty-but-present attribute is retained for compatibility
/// with the historical parser.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagType {
    namespace_uri: String,
    name: String,
    url: String,
}

impl SmartTagType {
    /// Construct a validated smart-tag vocabulary declaration.
    pub fn new(
        namespace_uri: impl Into<String>,
        name: impl Into<String>,
        url: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            namespace_uri: namespace_uri.into(),
            name: name.into(),
            url: url.into(),
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate the client length limits for this smart-tag declaration.
    pub fn validate(&self) -> Result<()> {
        validate_smart_tag_type(&self.namespace_uri, &self.name, &self.url)
    }

    /// Return the smart-tag vocabulary namespace URI.
    #[inline]
    pub fn namespace_uri(&self) -> &str {
        &self.namespace_uri
    }

    /// Return the smart-tag type name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the vocabulary download URL.
    #[inline]
    pub fn url(&self) -> &str {
        &self.url
    }
}

/// Validate the bounded client representation of a smart-tag declaration.
pub fn validate_smart_tag_type(namespace_uri: &str, name: &str, url: &str) -> Result<()> {
    validate_smart_tag_value(
        namespace_uri,
        "namespace URI",
        MAX_SMART_TAG_NAMESPACE_URI_CHARS,
    )?;
    validate_smart_tag_value(name, "name", MAX_SMART_TAG_NAME_CHARS)?;
    validate_smart_tag_value(url, "URL", MAX_SMART_TAG_URL_CHARS)
}

fn validate_smart_tag_value(value: &str, description: &str, maximum: usize) -> Result<()> {
    if value.chars().count() > maximum {
        return Err(invalid(format!(
            "Word smart-tag {description} exceeds {maximum} characters"
        )));
    }
    Ok(())
}

/// Error returned for a token outside the closed WordprocessingML
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
        /// WordprocessingML schema.
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

            /// Return the exact WordprocessingML element-local-name token.
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
    /// Whether this flag is part of the Strict WordprocessingML domain.
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
    pub const fn new(flag: CompatFlag, enabled: bool) -> Self {
        Self { flag, enabled }
    }

    /// Return the schema-defined flag.
    #[inline]
    pub const fn flag(&self) -> CompatFlag {
        self.flag
    }

    /// Whether the option is enabled.
    #[inline]
    pub const fn is_enabled(&self) -> bool {
        self.enabled
    }

    /// Serialize the standalone compatibility flag element.
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
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the URI scoping the setting name.
    #[inline]
    pub fn uri(&self) -> &str {
        &self.uri
    }

    /// Return the raw setting value.
    #[inline]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Serialize a standalone `w:compatSetting` element.
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

/// Placement of footnote or endnote text (`ST_FtnPos`/`ST_EdnPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum NotePosition {
    /// At the bottom of the page.
    PageBottom,
    /// Immediately beneath the page's text.
    BeneathText,
    /// At the end of the section.
    SectionEnd,
    /// At the end of the document.
    DocumentEnd,
}

impl NotePosition {
    /// Return the exact WordprocessingML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::PageBottom => "pageBottom",
            Self::BeneathText => "beneathText",
            Self::SectionEnd => "sectEnd",
            Self::DocumentEnd => "docEnd",
        }
    }

    /// Whether this placement is valid for an endnote.
    pub const fn valid_for_endnote(self) -> bool {
        matches!(self, Self::SectionEnd | Self::DocumentEnd)
    }
}

/// Error returned for a token outside `ST_FtnPos`/`ST_EdnPos`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseNotePositionError;

impl Display for ParseNotePositionError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML note position")
    }
}

impl StdError for ParseNotePositionError {}

impl FromStr for NotePosition {
    type Err = ParseNotePositionError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        match value {
            "pageBottom" => Ok(Self::PageBottom),
            "beneathText" => Ok(Self::BeneathText),
            "sectEnd" => Ok(Self::SectionEnd),
            "docEnd" => Ok(Self::DocumentEnd),
            _ => Err(ParseNotePositionError),
        }
    }
}

impl Display for NotePosition {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Numbering restart behavior for footnotes or endnotes (`w:numRestart`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NoteNumberingRestart {
    /// Numbering continues throughout the document.
    Continuous,
    /// Numbering restarts at each section.
    EachSection,
    /// Numbering restarts at each page.
    EachPage,
}

impl NoteNumberingRestart {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSection),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(invalid(format!(
                "invalid note numbering restart value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this restart behavior.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::EachSection => "eachSect",
            Self::EachPage => "eachPage",
        }
    }
}

macro_rules! define_note_formats {
    ($($variant:ident => $token:literal),+ $(,)?) => {
        /// Checked `ST_NumberFormat` value used by note numbering properties.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[repr(u8)]
        pub enum NoteNumberFormat {
            $($variant,)+
        }

        impl NoteNumberFormat {
            /// Return the exact WordprocessingML token.
            pub const fn as_str(self) -> &'static str {
                match self { $(Self::$variant => $token,)+ }
            }
        }

        impl FromStr for NoteNumberFormat {
            type Err = ParseNoteNumberFormatError;

            fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
                match value { $($token => Ok(Self::$variant),)+ _ => Err(ParseNoteNumberFormatError) }
            }
        }

        impl Display for NoteNumberFormat {
            fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
                formatter.write_str(self.as_str())
            }
        }
    };
}

/// Error returned for a token outside `ST_NumberFormat`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseNoteNumberFormatError;

impl Display for ParseNoteNumberFormatError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML note number format")
    }
}

impl StdError for ParseNoteNumberFormatError {}

define_note_formats! {
    Decimal => "decimal",
    UpperRoman => "upperRoman",
    LowerRoman => "lowerRoman",
    UpperLetter => "upperLetter",
    LowerLetter => "lowerLetter",
    Ordinal => "ordinal",
    CardinalText => "cardinalText",
    OrdinalText => "ordinalText",
    Hex => "hex",
    Chicago => "chicago",
    IdeographDigital => "ideographDigital",
    JapaneseCounting => "japaneseCounting",
    Aiueo => "aiueo",
    Iroha => "iroha",
    DecimalFullWidth => "decimalFullWidth",
    DecimalHalfWidth => "decimalHalfWidth",
    JapaneseLegal => "japaneseLegal",
    JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
    DecimalEnclosedCircle => "decimalEnclosedCircle",
    DecimalFullWidth2 => "decimalFullWidth2",
    AiueoFullWidth => "aiueoFullWidth",
    IrohaFullWidth => "irohaFullWidth",
    DecimalZero => "decimalZero",
    Bullet => "bullet",
    Ganada => "ganada",
    Chosung => "chosung",
    DecimalEnclosedFullStop => "decimalEnclosedFullstop",
    DecimalEnclosedParen => "decimalEnclosedParen",
    DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
    IdeographEnclosedCircle => "ideographEnclosedCircle",
    IdeographTraditional => "ideographTraditional",
    IdeographZodiac => "ideographZodiac",
    IdeographZodiacTraditional => "ideographZodiacTraditional",
    TaiwaneseCounting => "taiwaneseCounting",
    IdeographLegalTraditional => "ideographLegalTraditional",
    TaiwaneseCountingThousand => "taiwaneseCountingThousand",
    TaiwaneseDigital => "taiwaneseDigital",
    ChineseCounting => "chineseCounting",
    ChineseLegalSimplified => "chineseLegalSimplified",
    ChineseCountingThousand => "chineseCountingThousand",
    KoreanDigital => "koreanDigital",
    KoreanCounting => "koreanCounting",
    KoreanLegal => "koreanLegal",
    KoreanDigital2 => "koreanDigital2",
    VietnameseCounting => "vietnameseCounting",
    RussianLower => "russianLower",
    RussianUpper => "russianUpper",
    None => "none",
    NumberInDash => "numberInDash",
    Hebrew1 => "hebrew1",
    Hebrew2 => "hebrew2",
    ArabicAlpha => "arabicAlpha",
    ArabicAbjad => "arabicAbjad",
    HindiVowels => "hindiVowels",
    HindiConsonants => "hindiConsonants",
    HindiNumbers => "hindiNumbers",
    HindiCounting => "hindiCounting",
    ThaiLetters => "thaiLetters",
    ThaiNumbers => "thaiNumbers",
    ThaiCounting => "thaiCounting",
    BahtText => "bahtText",
    DollarText => "dollarText",
    Custom => "custom",
}

/// Document-level footnote or endnote properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NoteNumberingProperties<F = NoteNumberFormat> {
    position: Option<NotePosition>,
    format: Option<F>,
    start: Option<u32>,
    restart: Option<NoteNumberingRestart>,
}

impl<F> Default for NoteNumberingProperties<F> {
    fn default() -> Self {
        Self {
            position: None,
            format: None,
            start: None,
            restart: None,
        }
    }
}

impl<F> NoteNumberingProperties<F> {
    /// Construct note properties from validated component values.
    pub fn from_parts(
        position: Option<NotePosition>,
        format: Option<F>,
        start: Option<u32>,
        restart: Option<NoteNumberingRestart>,
    ) -> Self {
        Self {
            position,
            format,
            start,
            restart,
        }
    }

    /// Return the note placement, when specified.
    #[inline]
    pub const fn position(&self) -> Option<NotePosition> {
        self.position
    }

    /// Return the numbering format, when specified.
    #[inline]
    pub const fn format(&self) -> Option<F>
    where
        F: Copy,
    {
        self.format
    }

    /// Return the first note number, when specified.
    #[inline]
    pub const fn start(&self) -> Option<u32> {
        self.start
    }

    /// Return the numbering restart behavior, when specified.
    #[inline]
    pub const fn restart(&self) -> Option<NoteNumberingRestart> {
        self.restart
    }
}

impl NoteNumberingProperties<NoteNumberFormat> {
    fn try_map_format<G, E>(
        self,
        map: &mut impl FnMut(NoteNumberFormat) -> std::result::Result<G, E>,
    ) -> std::result::Result<NoteNumberingProperties<G>, E> {
        Ok(NoteNumberingProperties::from_parts(
            self.position,
            self.format.map(map).transpose()?,
            self.start,
            self.restart,
        ))
    }

    fn to_xml(&self, prefix: &str, name: &str) -> String {
        let mut xml = format!("<{prefix}:{name}>");
        if let Some(position) = self.position {
            xml.push_str(&format!(
                "<{prefix}:pos {prefix}:val=\"{}\"/>",
                position.as_str()
            ));
        }
        if let Some(format) = self.format {
            xml.push_str(&format!(
                "<{prefix}:numFmt {prefix}:val=\"{}\"/>",
                format.as_str()
            ));
        }
        if let Some(start) = self.start {
            xml.push_str(&format!("<{prefix}:numStart {prefix}:val=\"{start}\"/>"));
        }
        if let Some(restart) = self.restart {
            xml.push_str(&format!(
                "<{prefix}:numRestart {prefix}:val=\"{}\"/>",
                restart.as_str()
            ));
        }
        xml.push_str(&format!("</{prefix}:{name}>"));
        xml
    }
}

/// Type of document protection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProtectionType {
    /// No editing allowed.
    ReadOnly,
    /// Only comments allowed.
    Comments,
    /// Only tracked changes allowed.
    TrackedChanges,
    /// Only form fields allowed.
    Forms,
}

impl ProtectionType {
    /// Parse the optional `w:edit` token.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    /// Get the XML value for this protection type.
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

/// Document view mode from `w:view` (`ST_View`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentView {
    /// No explicit view is specified.
    None,
    /// Print layout view.
    Print,
    /// Outline view.
    Outline,
    /// Master pages view.
    MasterPages,
    /// Normal (draft) view.
    Normal,
    /// Web layout view.
    Web,
}

impl DocumentView {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "print" => Ok(Self::Print),
            "outline" => Ok(Self::Outline),
            "masterPages" => Ok(Self::MasterPages),
            "normal" => Ok(Self::Normal),
            "web" => Ok(Self::Web),
            _ => Err(invalid(format!("invalid document view value '{value}'"))),
        }
    }

    /// Get the XML value for this view mode.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Print => "print",
            Self::Outline => "outline",
            Self::MasterPages => "masterPages",
            Self::Normal => "normal",
            Self::Web => "web",
        }
    }
}

/// Proofing completion marker (`ST_ProofState`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProofState {
    /// Proofing for the region completed without errors.
    Clean,
    /// The region changed since proofing last ran.
    Dirty,
}

impl ProofState {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "clean" => Ok(Self::Clean),
            "dirty" => Ok(Self::Dirty),
            _ => Err(invalid(format!("invalid proof state value '{value}'"))),
        }
    }

    /// Get the XML value for this proof state.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Clean => "clean",
            Self::Dirty => "dirty",
        }
    }
}

/// Proofing completion markers from `w:proofState`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProofingState {
    spelling: Option<ProofState>,
    grammar: Option<ProofState>,
}

impl ProofingState {
    /// Create a proofing state with no markers.
    pub const fn new() -> Self {
        Self {
            spelling: None,
            grammar: None,
        }
    }

    /// Set the spelling proofing marker.
    pub fn set_spelling(&mut self, value: Option<ProofState>) -> &mut Self {
        self.spelling = value;
        self
    }

    /// Set the grammar proofing marker.
    pub fn set_grammar(&mut self, value: Option<ProofState>) -> &mut Self {
        self.grammar = value;
        self
    }

    /// Return the spelling proofing marker, when specified.
    #[inline]
    pub const fn spelling(&self) -> Option<ProofState> {
        self.spelling
    }

    /// Return the grammar proofing marker, when specified.
    #[inline]
    pub const fn grammar(&self) -> Option<ProofState> {
        self.grammar
    }

    /// Serialize a standalone `w:proofState` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:proofState");
        if let Some(spelling) = self.spelling {
            xml.push_str(&format!(" {prefix}:spelling=\"{}\"", spelling.as_str()));
        }
        if let Some(grammar) = self.grammar {
            xml.push_str(&format!(" {prefix}:grammar=\"{}\"", grammar.as_str()));
        }
        xml.push_str("/>");
        xml
    }
}

/// Maximum accepted length of a `w:themeFontLang` language tag.
pub const MAX_LANGUAGE_TAG_LENGTH: usize = 255;

fn validate_language_tag(value: &str, description: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_LANGUAGE_TAG_LENGTH
        || value.chars().any(char::is_control)
    {
        return Err(invalid(format!(
            "invalid Word {description} language tag '{value}'"
        )));
    }
    Ok(())
}

/// Theme font language defaults from `w:themeFontLang`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeFontLanguages {
    latin: Option<String>,
    east_asia: Option<String>,
    bidi: Option<String>,
}

impl ThemeFontLanguages {
    /// Create theme font language defaults with no languages set.
    pub const fn new() -> Self {
        Self {
            latin: None,
            east_asia: None,
            bidi: None,
        }
    }

    /// Set the Latin (`w:val`) theme language.
    pub fn set_latin(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "Latin theme font")?;
        }
        self.latin = value;
        Ok(self)
    }

    /// Set the East Asian (`w:eastAsia`) theme language.
    pub fn set_east_asia(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "East Asian theme font")?;
        }
        self.east_asia = value;
        Ok(self)
    }

    /// Set the complex-script (`w:bidi`) theme language.
    pub fn set_bidi(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "complex-script theme font")?;
        }
        self.bidi = value;
        Ok(self)
    }

    /// Return the Latin theme language, when specified.
    #[inline]
    pub fn latin(&self) -> Option<&str> {
        self.latin.as_deref()
    }

    /// Return the East Asian theme language, when specified.
    #[inline]
    pub fn east_asia(&self) -> Option<&str> {
        self.east_asia.as_deref()
    }

    /// Return the complex-script theme language, when specified.
    #[inline]
    pub fn bidi(&self) -> Option<&str> {
        self.bidi.as_deref()
    }

    /// Serialize a standalone `w:themeFontLang` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:themeFontLang");
        for (name, value) in [
            ("val", &self.latin),
            ("eastAsia", &self.east_asia),
            ("bidi", &self.bidi),
        ] {
            if let Some(tag) = value {
                xml.push_str(&format!(" {prefix}:{name}=\""));
                escape_attribute(&mut xml, tag);
                xml.push('"');
            }
        }
        xml.push_str("/>");
        xml
    }
}

/// Theme color slot produced by a `w:clrSchemeMapping` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ColorSchemeIndex {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ColorSchemeIndex {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "dark1" => Ok(Self::Dark1),
            "light1" => Ok(Self::Light1),
            "dark2" => Ok(Self::Dark2),
            "light2" => Ok(Self::Light2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hyperlink" => Ok(Self::Hyperlink),
            "followedHyperlink" => Ok(Self::FollowedHyperlink),
            _ => Err(invalid(format!(
                "invalid color scheme index value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this theme color slot.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// A remappable theme color slot on `w:clrSchemeMapping`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ColorSchemeSlot {
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ColorSchemeSlot {
    /// Number of remappable color slots.
    pub const COUNT: usize = 12;

    /// Every slot in schema attribute order.
    pub const ALL: [Self; Self::COUNT] = [
        Self::Background1,
        Self::Text1,
        Self::Background2,
        Self::Text2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    const fn index(self) -> usize {
        self as usize
    }

    /// Get the attribute name carrying this slot.
    pub const fn attribute_name(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "t1",
            Self::Background2 => "bg2",
            Self::Text2 => "t2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// Theme color slot remapping from `w:clrSchemeMapping`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ColorSchemeMapping {
    slots: [Option<ColorSchemeIndex>; ColorSchemeSlot::COUNT],
}

impl ColorSchemeMapping {
    /// Create a mapping with every slot left at its default.
    pub const fn new() -> Self {
        Self {
            slots: [None; ColorSchemeSlot::COUNT],
        }
    }

    /// Remap a slot to a theme color index.
    pub fn set(&mut self, slot: ColorSchemeSlot, index: ColorSchemeIndex) -> &mut Self {
        self.slots[slot.index()] = Some(index);
        self
    }

    /// Restore a slot to its default mapping.
    pub fn clear(&mut self, slot: ColorSchemeSlot) -> &mut Self {
        self.slots[slot.index()] = None;
        self
    }

    /// Return the theme color index a slot maps to, when remapped.
    #[inline]
    pub const fn get(&self, slot: ColorSchemeSlot) -> Option<ColorSchemeIndex> {
        self.slots[slot.index()]
    }

    /// Whether no slot is remapped.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.slots.iter().all(Option::is_none)
    }

    /// Iterate the remapped slots in schema attribute order.
    pub fn iter(&self) -> impl Iterator<Item = (ColorSchemeSlot, ColorSchemeIndex)> + '_ {
        ColorSchemeSlot::ALL
            .into_iter()
            .filter_map(|slot| self.get(slot).map(|index| (slot, index)))
    }

    /// Serialize a standalone `w:clrSchemeMapping` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:clrSchemeMapping");
        for (slot, index) in self.iter() {
            xml.push_str(&format!(
                " {prefix}:{}=\"{}\"",
                slot.attribute_name(),
                index.as_str()
            ));
        }
        xml.push_str("/>");
        xml
    }
}

/// Format-owned scalar settings extracted from a Word settings part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings<F = NoteNumberFormat> {
    protected: bool,
    protection_type: Option<ProtectionType>,
    track_revisions: bool,
    zoom_percent: Option<u32>,
    compatibility_options: Vec<CompatibilityOption>,
    compatibility_settings: Vec<CompatibilitySetting>,
    footnote_properties: Option<NoteNumberingProperties<F>>,
    endnote_properties: Option<NoteNumberingProperties<F>>,
    write_protection: bool,
    view: Option<DocumentView>,
    proofing_state: Option<ProofingState>,
    default_tab_stop_twips: Option<u32>,
    theme_font_languages: Option<ThemeFontLanguages>,
    color_scheme_mapping: Option<ColorSchemeMapping>,
}

impl<F> Default for Settings<F> {
    fn default() -> Self {
        Self {
            protected: false,
            protection_type: None,
            track_revisions: false,
            zoom_percent: None,
            compatibility_options: Vec::new(),
            compatibility_settings: Vec::new(),
            footnote_properties: None,
            endnote_properties: None,
            write_protection: false,
            view: None,
            proofing_state: None,
            default_tab_stop_twips: None,
            theme_font_languages: None,
            color_scheme_mapping: None,
        }
    }
}

impl<F> Settings<F> {
    /// Create empty settings values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the document-protection marker.
    pub fn set_protected(&mut self, value: bool) -> &mut Self {
        self.protected = value;
        self
    }

    /// Set the document-protection editing mode.
    pub fn set_protection_type(&mut self, value: Option<ProtectionType>) -> &mut Self {
        self.protection_type = value;
        self
    }

    /// Set the tracked-revisions marker.
    pub fn set_track_revisions(&mut self, value: bool) -> &mut Self {
        self.track_revisions = value;
        self
    }

    /// Set the document zoom percentage.
    pub fn set_zoom_percent(&mut self, value: Option<u32>) -> &mut Self {
        self.zoom_percent = value;
        self
    }

    /// Replace the compatibility option sequence.
    pub fn set_compatibility_options(&mut self, value: Vec<CompatibilityOption>) -> &mut Self {
        self.compatibility_options = value;
        self
    }

    /// Replace the compatibility-setting sequence.
    pub fn set_compatibility_settings(&mut self, value: Vec<CompatibilitySetting>) -> &mut Self {
        self.compatibility_settings = value;
        self
    }

    /// Replace the footnote properties.
    pub fn set_footnote_properties(
        &mut self,
        value: Option<NoteNumberingProperties<F>>,
    ) -> &mut Self {
        self.footnote_properties = value;
        self
    }

    /// Replace the endnote properties.
    pub fn set_endnote_properties(
        &mut self,
        value: Option<NoteNumberingProperties<F>>,
    ) -> &mut Self {
        self.endnote_properties = value;
        self
    }

    /// Set the write-protection marker.
    pub fn set_write_protected(&mut self, value: bool) -> &mut Self {
        self.write_protection = value;
        self
    }

    /// Set the document view mode.
    pub fn set_view(&mut self, value: Option<DocumentView>) -> &mut Self {
        self.view = value;
        self
    }

    /// Replace the proofing state.
    pub fn set_proofing_state(&mut self, value: Option<ProofingState>) -> &mut Self {
        self.proofing_state = value;
        self
    }

    /// Set the default tab stop interval.
    pub fn set_default_tab_stop_twips(&mut self, value: Option<u32>) -> &mut Self {
        self.default_tab_stop_twips = value;
        self
    }

    /// Replace the theme-font language defaults.
    pub fn set_theme_font_languages(&mut self, value: Option<ThemeFontLanguages>) -> &mut Self {
        self.theme_font_languages = value;
        self
    }

    /// Replace the theme color mapping.
    pub fn set_color_scheme_mapping(&mut self, value: Option<ColorSchemeMapping>) -> &mut Self {
        self.color_scheme_mapping = value;
        self
    }
}

impl<F: Copy> Settings<F> {
    /// Check if the document is protected.
    #[inline]
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Get the type of protection applied.
    #[inline]
    pub const fn protection_type(&self) -> Option<ProtectionType> {
        self.protection_type
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub const fn track_revisions(&self) -> bool {
        self.track_revisions
    }

    /// Get the zoom percentage.
    #[inline]
    pub const fn zoom_percent(&self) -> Option<u32> {
        self.zoom_percent
    }

    /// Return the on/off compatibility options in document order.
    #[inline]
    pub fn compatibility_options(&self) -> &[CompatibilityOption] {
        &self.compatibility_options
    }

    /// Return the `w:compatSetting` triples in document order.
    #[inline]
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        &self.compatibility_settings
    }

    /// Look up a `w:compatSetting` triple by name and URI.
    pub fn compatibility_setting(&self, name: &str, uri: &str) -> Option<&CompatibilitySetting> {
        self.compatibility_settings
            .iter()
            .find(|setting| setting.name == name && setting.uri == uri)
    }

    /// Return the Word compatibility mode, when declared.
    pub fn compatibility_mode(&self) -> Option<u32> {
        self.compatibility_setting(COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI)
            .and_then(|setting| setting.value.parse().ok())
    }

    /// Return the document-level footnote properties, if present.
    #[inline]
    pub fn footnote_properties(&self) -> Option<&NoteNumberingProperties<F>> {
        self.footnote_properties.as_ref()
    }

    /// Return the document-level endnote properties, if present.
    #[inline]
    pub fn endnote_properties(&self) -> Option<&NoteNumberingProperties<F>> {
        self.endnote_properties.as_ref()
    }

    /// Whether applications should recommend write protection.
    #[inline]
    pub const fn is_write_protected(&self) -> bool {
        self.write_protection
    }

    /// Return the document view mode, when specified.
    #[inline]
    pub const fn view(&self) -> Option<DocumentView> {
        self.view
    }

    /// Return the proofing completion markers, if present.
    #[inline]
    pub fn proofing_state(&self) -> Option<&ProofingState> {
        self.proofing_state.as_ref()
    }

    /// Return the default tab stop interval in twips, when specified.
    #[inline]
    pub const fn default_tab_stop_twips(&self) -> Option<u32> {
        self.default_tab_stop_twips
    }

    /// Return the theme font language defaults, if present.
    #[inline]
    pub fn theme_font_languages(&self) -> Option<&ThemeFontLanguages> {
        self.theme_font_languages.as_ref()
    }

    /// Return the theme color slot remapping, if present.
    #[inline]
    pub fn color_scheme_mapping(&self) -> Option<&ColorSchemeMapping> {
        self.color_scheme_mapping.as_ref()
    }

    /// Serialize the editing/view/theme settings in schema order.
    pub fn to_editing_settings_xml(&self, prefix: &str) -> String {
        let mut xml = String::new();
        if self.write_protection {
            xml.push_str(&format!("<{prefix}:writeProtection/>"));
        }
        if let Some(view) = self.view {
            xml.push_str(&format!(
                "<{prefix}:view {prefix}:val=\"{}\"/>",
                view.as_str()
            ));
        }
        if let Some(state) = &self.proofing_state {
            xml.push_str(&state.to_xml(prefix));
        }
        if let Some(twips) = self.default_tab_stop_twips {
            xml.push_str(&format!(
                "<{prefix}:defaultTabStop {prefix}:val=\"{twips}\"/>"
            ));
        }
        if let Some(languages) = &self.theme_font_languages {
            xml.push_str(&languages.to_xml(prefix));
        }
        if let Some(mapping) = &self.color_scheme_mapping {
            xml.push_str(&mapping.to_xml(prefix));
        }
        xml
    }
}

impl Settings<NoteNumberFormat> {
    /// Parse the bounded format-owned portion of a `settings.xml` document.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_SETTINGS_XML_BYTES {
            return Err(invalid(format!(
                "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes"
            )));
        }

        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut saw_root = false;
        let mut strict_wordprocessingml = false;
        let mut seen = SeenSettings::default();
        let mut saw_compat = false;
        let mut pending_group: Option<PendingGroup> = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML node counter overflow"))?;
                if nodes > MAX_SETTINGS_XML_NODES {
                    return Err(invalid(format!(
                        "settings XML exceeds {MAX_SETTINGS_XML_NODES} nodes"
                    )));
                }
            }

            match event {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("Word settings XML nesting is too deep"))?;
                    if depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri)) if uri == STRICT_WORD_NAMESPACE
                        );
                        saw_root = true;
                    } else if saw_root && is_wordprocessing_namespace(&namespace) {
                        if depth == 2 {
                            if let Some(group) =
                                begin_settings_group(&element, &settings, &mut saw_compat)?
                            {
                                pending_group = Some(group);
                            } else {
                                parse_setting(
                                    &element,
                                    decoder,
                                    &resolver,
                                    &mut settings,
                                    &mut seen,
                                )?;
                            }
                        } else if depth == 3
                            && let Some(group) = pending_group.as_mut()
                        {
                            parse_group_child(
                                group,
                                strict_wordprocessingml,
                                &element,
                                decoder,
                                &resolver,
                            )?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("Word settings XML nesting is too deep"))?;
                    if child_depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri)) if uri == STRICT_WORD_NAMESPACE
                        );
                        saw_root = true;
                    } else if saw_root && is_wordprocessing_namespace(&namespace) {
                        if child_depth == 2 {
                            if let Some(group) =
                                begin_settings_group(&element, &settings, &mut saw_compat)?
                            {
                                finish_settings_group(&mut settings, group)?;
                            } else {
                                parse_setting(
                                    &element,
                                    decoder,
                                    &resolver,
                                    &mut settings,
                                    &mut seen,
                                )?;
                            }
                        } else if child_depth == 3
                            && let Some(group) = pending_group.as_mut()
                        {
                            parse_group_child(
                                group,
                                strict_wordprocessingml,
                                &element,
                                decoder,
                                &resolver,
                            )?;
                        }
                    }
                },
                Event::End(_) => {
                    if depth == 2
                        && let Some(group) = pending_group.take()
                    {
                        finish_settings_group(&mut settings, group)?;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid Word settings XML nesting"))?;
                },
                Event::Eof if depth != 0 => {
                    return Err(invalid("unterminated Word settings XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(invalid("settings part has no settings root"));
        }
        Ok(settings)
    }

    /// Map the owner-local note format to a host format without reparsing XML.
    pub fn try_map_note_format<G, E>(
        self,
        mut map: impl FnMut(NoteNumberFormat) -> std::result::Result<G, E>,
    ) -> std::result::Result<Settings<G>, E> {
        let Settings {
            protected,
            protection_type,
            track_revisions,
            zoom_percent,
            compatibility_options,
            compatibility_settings,
            footnote_properties,
            endnote_properties,
            write_protection,
            view,
            proofing_state,
            default_tab_stop_twips,
            theme_font_languages,
            color_scheme_mapping,
        } = self;
        Ok(Settings {
            protected,
            protection_type,
            track_revisions,
            zoom_percent,
            compatibility_options,
            compatibility_settings,
            footnote_properties: footnote_properties
                .map(|value| value.try_map_format(&mut map))
                .transpose()?,
            endnote_properties: endnote_properties
                .map(|value| value.try_map_format(&mut map))
                .transpose()?,
            write_protection,
            view,
            proofing_state,
            default_tab_stop_twips,
            theme_font_languages,
            color_scheme_mapping,
        })
    }

    /// Serialize the complete modeled format-owned settings fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = String::new();
        if self.protected {
            xml.push_str(&format!("<{prefix}:documentProtection"));
            if let Some(protection_type) = self.protection_type {
                xml.push_str(&format!(" {prefix}:edit=\"{}\"", protection_type.to_xml()));
            }
            xml.push_str(&format!(" {prefix}:enforcement=\"on\"/>"));
        }
        if self.track_revisions {
            xml.push_str(&format!("<{prefix}:trackRevisions/>"));
        }
        if let Some(percent) = self.zoom_percent {
            xml.push_str(&format!("<{prefix}:zoom {prefix}:percent=\"{percent}\"/>"));
        }
        if !self.compatibility_options.is_empty() || !self.compatibility_settings.is_empty() {
            xml.push_str(&format!("<{prefix}:compat>"));
            for option in &self.compatibility_options {
                xml.push_str(&option.to_xml(prefix));
            }
            for setting in &self.compatibility_settings {
                xml.push_str(&setting.to_xml(prefix));
            }
            xml.push_str(&format!("</{prefix}:compat>"));
        }
        if let Some(properties) = &self.footnote_properties {
            xml.push_str(&properties.to_xml(prefix, "footnotePr"));
        }
        if let Some(properties) = &self.endnote_properties {
            xml.push_str(&properties.to_xml(prefix, "endnotePr"));
        }
        xml.push_str(&self.to_editing_settings_xml(prefix));
        xml
    }
}

#[derive(Debug, Default)]
struct SeenSettings {
    write_protection: bool,
    view: bool,
    proofing_state: bool,
    default_tab_stop: bool,
    theme_font_languages: bool,
    color_scheme_mapping: bool,
}

enum PendingGroup {
    Compatibility {
        options: Vec<CompatibilityOption>,
        settings: Vec<CompatibilitySetting>,
    },
    FootnoteProperties(NoteNumberingProperties),
    EndnoteProperties(NoteNumberingProperties),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoteKind {
    Footnote,
    Endnote,
}

fn begin_settings_group(
    element: &BytesStart<'_>,
    settings: &Settings,
    saw_compat: &mut bool,
) -> Result<Option<PendingGroup>> {
    match element.local_name().as_ref() {
        b"compat" => {
            if std::mem::replace(saw_compat, true) {
                return Err(invalid("duplicate compat settings group"));
            }
            Ok(Some(PendingGroup::Compatibility {
                options: Vec::new(),
                settings: Vec::new(),
            }))
        },
        b"footnotePr" => {
            if settings.footnote_properties.is_some() {
                return Err(invalid("duplicate footnotePr settings group"));
            }
            Ok(Some(PendingGroup::FootnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        b"endnotePr" => {
            if settings.endnote_properties.is_some() {
                return Err(invalid("duplicate endnotePr settings group"));
            }
            Ok(Some(PendingGroup::EndnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        _ => Ok(None),
    }
}

fn finish_settings_group(settings: &mut Settings, group: PendingGroup) -> Result<()> {
    match group {
        PendingGroup::Compatibility {
            options,
            settings: triples,
        } => {
            settings.compatibility_options = options;
            settings.compatibility_settings = triples;
        },
        PendingGroup::FootnoteProperties(properties) => {
            if settings.footnote_properties.replace(properties).is_some() {
                return Err(invalid("duplicate footnotePr settings group"));
            }
        },
        PendingGroup::EndnoteProperties(properties) => {
            if settings.endnote_properties.replace(properties).is_some() {
                return Err(invalid("duplicate endnotePr settings group"));
            }
        },
    }
    Ok(())
}

fn parse_group_child(
    group: &mut PendingGroup,
    strict_wordprocessingml: bool,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match group {
        PendingGroup::Compatibility { options, settings } => {
            if element.local_name().as_ref() == b"compatSetting" {
                reserve_one(settings, "Word compatibility settings")?;
                settings.push(CompatibilitySetting::new(
                    required_attribute(element, b"name", decoder, resolver, "compatSetting name")?,
                    required_attribute(element, b"uri", decoder, resolver, "compatSetting URI")?,
                    required_attribute(element, b"val", decoder, resolver, "compatSetting value")?,
                ));
            } else {
                let local_name = element.local_name();
                let raw = std::str::from_utf8(local_name.as_ref())
                    .map_err(|_| invalid("compatibility flag name is not valid UTF-8"))?;
                let flag = raw
                    .parse::<CompatFlag>()
                    .map_err(|_| invalid(format!("invalid compatibility flag '{raw}'")))?;
                if strict_wordprocessingml && !flag.is_strict() {
                    return Err(invalid(format!(
                        "compatibility flag '{raw}' is not valid in Strict WordprocessingML"
                    )));
                }
                if options.iter().any(|option| option.flag == flag) {
                    return Err(invalid(format!("duplicate compatibility flag '{raw}'")));
                }
                reserve_one(options, "Word compatibility options")?;
                options.push(CompatibilityOption::new(
                    flag,
                    parse_on_off(element, decoder, resolver)?,
                ));
            }
        },
        PendingGroup::FootnoteProperties(properties) => {
            parse_note_property_child(properties, NoteKind::Footnote, element, decoder, resolver)?;
        },
        PendingGroup::EndnoteProperties(properties) => {
            parse_note_property_child(properties, NoteKind::Endnote, element, decoder, resolver)?;
        },
    }
    Ok(())
}

fn parse_note_property_child(
    properties: &mut NoteNumberingProperties,
    kind: NoteKind,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"pos" => {
            if properties.position.is_some() {
                return Err(invalid("duplicate note position"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note position")?;
            let position = value
                .parse::<NotePosition>()
                .map_err(|_| invalid(format!("invalid note position '{value}'")))?;
            if kind == NoteKind::Endnote && !position.valid_for_endnote() {
                return Err(invalid(format!(
                    "position '{}' is not valid for an endnote",
                    position.as_str()
                )));
            }
            properties.position = Some(position);
        },
        b"numFmt" => {
            if properties.format.is_some() {
                return Err(invalid("duplicate note numbering format"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numFmt")?;
            properties.format = Some(
                value
                    .parse()
                    .map_err(|_| invalid(format!("invalid note numbering format '{value}'")))?,
            );
        },
        b"numStart" => {
            if properties.start.is_some() {
                return Err(invalid("duplicate note numbering start"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numStart")?;
            properties.start = Some(
                value
                    .parse()
                    .map_err(|_| invalid(format!("invalid note numbering start '{value}'")))?,
            );
        },
        b"numRestart" => {
            if properties.restart.is_some() {
                return Err(invalid("duplicate note numbering restart"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numRestart")?;
            properties.restart = Some(NoteNumberingRestart::from_xml(&value)?);
        },
        // `w:footnote`/`w:endnote` separator references carry no properties.
        _ => {},
    }
    Ok(())
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut Settings,
    seen: &mut SeenSettings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"documentProtection" => {
            settings.protected = true;
            if let Some(value) = word_attribute_value(element, b"edit", decoder, resolver)? {
                settings.protection_type = ProtectionType::from_xml(&value);
            }
            if let Some(value) = word_attribute_value(element, b"enforcement", decoder, resolver)? {
                settings.protected = parse_on_off_value(&value)?;
            }
        },
        b"trackRevisions" => {
            settings.track_revisions = parse_on_off(element, decoder, resolver)?;
        },
        b"zoom" => {
            if let Some(value) = word_attribute_value(element, b"percent", decoder, resolver)? {
                settings.zoom_percent = value.parse::<u32>().ok();
            }
        },
        b"writeProtection" => {
            if std::mem::replace(&mut seen.write_protection, true) {
                return Err(invalid("duplicate writeProtection setting"));
            }
            settings.write_protection = parse_on_off(element, decoder, resolver)?;
        },
        b"view" => {
            if std::mem::replace(&mut seen.view, true) {
                return Err(invalid("duplicate view setting"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "view mode")?;
            settings.view = Some(DocumentView::from_xml(&value)?);
        },
        b"proofState" => {
            if std::mem::replace(&mut seen.proofing_state, true) {
                return Err(invalid("duplicate proofState setting"));
            }
            let mut state = ProofingState::new();
            if let Some(value) = word_attribute_value(element, b"spelling", decoder, resolver)? {
                state.set_spelling(Some(ProofState::from_xml(&value)?));
            }
            if let Some(value) = word_attribute_value(element, b"grammar", decoder, resolver)? {
                state.set_grammar(Some(ProofState::from_xml(&value)?));
            }
            settings.proofing_state = Some(state);
        },
        b"defaultTabStop" => {
            if std::mem::replace(&mut seen.default_tab_stop, true) {
                return Err(invalid("duplicate defaultTabStop setting"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "default tab stop")?;
            settings.default_tab_stop_twips = Some(
                value
                    .parse()
                    .map_err(|_| invalid(format!("invalid default tab stop '{value}'")))?,
            );
        },
        b"themeFontLang" => {
            if std::mem::replace(&mut seen.theme_font_languages, true) {
                return Err(invalid("duplicate themeFontLang setting"));
            }
            let mut languages = ThemeFontLanguages::new();
            if let Some(value) = word_attribute_value(element, b"val", decoder, resolver)? {
                languages.set_latin(Some(value))?;
            }
            if let Some(value) = word_attribute_value(element, b"eastAsia", decoder, resolver)? {
                languages.set_east_asia(Some(value))?;
            }
            if let Some(value) = word_attribute_value(element, b"bidi", decoder, resolver)? {
                languages.set_bidi(Some(value))?;
            }
            settings.theme_font_languages = Some(languages);
        },
        b"clrSchemeMapping" => {
            if std::mem::replace(&mut seen.color_scheme_mapping, true) {
                return Err(invalid("duplicate clrSchemeMapping setting"));
            }
            let mut mapping = ColorSchemeMapping::new();
            for slot in ColorSchemeSlot::ALL {
                if let Some(value) = word_attribute_value(
                    element,
                    slot.attribute_name().as_bytes(),
                    decoder,
                    resolver,
                )? {
                    mapping.set(slot, ColorSchemeIndex::from_xml(&value)?);
                }
            }
            settings.color_scheme_mapping = Some(mapping);
        },
        _ => {},
    }
    Ok(())
}

fn validate_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(invalid(
            "settings part has an invalid or trailing root element",
        ));
    }
    Ok(())
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == TRANSITIONAL_WORD_NAMESPACE || *value == STRICT_WORD_NAMESPACE
    )
}

fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            return Err(invalid(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn required_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?
        .ok_or_else(|| invalid(format!("Word {description} attribute is required")))
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .as_deref()
        .map_or(Ok(true), parse_on_off_value)
}

fn parse_on_off_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" | "on" => Ok(true),
        "false" | "0" | "off" => Ok(false),
        _ => Err(invalid(format!("invalid Word on/off value '{value}'"))),
    }
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(message: impl Into<String>) -> Error {
    Error::Xml(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

    #[test]
    fn parses_requested_settings_values_in_both_dialects() {
        for namespace in [TRANSITIONAL, STRICT] {
            let compatibility_flag = if namespace == STRICT {
                "spaceForUL"
            } else {
                "useFELayout"
            };
            let xml = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:documentProtection w:edit="readOnly" w:enforcement="on"/><w:trackRevisions/><w:zoom w:percent="125"/><w:compat><w:{compatibility_flag}/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat><w:footnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachPage"/></w:footnotePr><w:view w:val="print"/><w:proofState w:spelling="clean"/><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1"/></w:settings>"#
            );
            let settings = Settings::parse(xml.as_bytes()).unwrap();
            assert!(settings.is_protected());
            assert_eq!(settings.protection_type(), Some(ProtectionType::ReadOnly));
            assert!(settings.track_revisions());
            assert_eq!(settings.zoom_percent(), Some(125));
            assert_eq!(settings.compatibility_options().len(), 1);
            assert_eq!(settings.compatibility_mode(), Some(14));
            assert_eq!(
                settings.footnote_properties().unwrap().format(),
                Some(NoteNumberFormat::LowerRoman)
            );
            assert_eq!(settings.view(), Some(DocumentView::Print));
            assert_eq!(
                settings.proofing_state().unwrap().spelling(),
                Some(ProofState::Clean)
            );
            assert_eq!(settings.default_tab_stop_twips(), Some(720));
            assert_eq!(
                settings.theme_font_languages().unwrap().latin(),
                Some("en-US")
            );
            assert_eq!(
                settings
                    .color_scheme_mapping()
                    .unwrap()
                    .get(ColorSchemeSlot::Background1),
                Some(ColorSchemeIndex::Light1)
            );
        }
    }

    #[test]
    fn rejects_strict_transitional_compatibility_flags_and_duplicates() {
        let strict = format!(
            r#"<w:settings xmlns:w="{STRICT}"><w:compat><w:useFELayout/></w:compat></w:settings>"#
        );
        assert!(Settings::parse(strict.as_bytes()).is_err());

        let duplicate = format!(
            r#"<w:settings xmlns:w="{TRANSITIONAL}"><w:view w:val="print"/><w:view w:val="web"/></w:settings>"#
        );
        assert!(Settings::parse(duplicate.as_bytes()).is_err());
    }

    #[test]
    fn writes_modeled_values_in_schema_order() {
        let xml = format!(
            r#"<w:settings xmlns:w="{TRANSITIONAL}"><w:documentProtection w:edit="comments" w:enforcement="on"/><w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:view w:val="outline"/><w:proofState w:grammar="dirty"/><w:themeFontLang w:val="en-US"/></w:settings>"#
        );
        let settings = Settings::parse(xml.as_bytes()).unwrap();
        let output = settings.to_xml("w");
        assert!(output.starts_with("<w:documentProtection"));
        assert!(output.contains("<w:compat><w:useFELayout/>"));
        assert!(output.contains("<w:footnotePr><w:numFmt w:val=\"decimal\"/></w:footnotePr>"));
        assert!(output.ends_with("<w:themeFontLang w:val=\"en-US\"/>"));
    }

    #[test]
    fn smart_tag_type_validates_client_lengths_without_rejecting_empty_present_values() {
        let value = SmartTagType::new("é".repeat(MAX_SMART_TAG_NAMESPACE_URI_CHARS), "name", "url")
            .unwrap();
        assert_eq!(
            value.namespace_uri().chars().count(),
            MAX_SMART_TAG_NAMESPACE_URI_CHARS
        );

        assert!(
            SmartTagType::new("namespace", "n".repeat(MAX_SMART_TAG_NAME_CHARS + 1), "url",)
                .is_err()
        );
        assert!(
            validate_smart_tag_type(
                "namespace",
                "name",
                &"u".repeat(MAX_SMART_TAG_URL_CHARS + 1),
            )
            .is_err()
        );

        let empty = SmartTagType::new("", "", "").unwrap();
        assert!(empty.namespace_uri().is_empty());
        assert!(empty.name().is_empty());
        assert!(empty.url().is_empty());
    }

    #[test]
    fn rejects_oversized_input_before_xml_allocation() {
        let input = vec![b' '; MAX_SETTINGS_XML_BYTES + 1];
        assert!(Settings::parse(&input).is_err());
    }
}
