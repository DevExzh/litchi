/// Character Properties (CHP) parser for DOC files.
///
/// CHP structures define character-level formatting such as:
/// - Font properties (bold, italic, underline, strikethrough)
/// - Font size and name
/// - Text color and highlighting
/// - Superscript/subscript
/// - Embedded objects and pictures
///
/// Based on Apache POI's CharacterSprmUncompressor and CharacterProperties.
use super::super::super::package::{Error as PackageError, Result};
use super::super::tap::TableStyleCondition;

/// Character Properties structure.
///
/// Contains formatting information for a run of text.
/// Based on Apache POI's CharacterProperties implementation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CharacterProperties {
    /// Bold text
    pub is_bold: Option<bool>,
    /// Italic text
    pub is_italic: Option<bool>,
    /// Underline style
    pub underline: UnderlineStyle,
    /// Strikethrough
    pub is_strikethrough: Option<bool>,
    /// Double strikethrough
    pub is_double_strikethrough: Option<bool>,
    /// Font size in half-points (e.g., 24 = 12pt)
    pub font_size: Option<u16>,
    /// Font index in font table (ASCII characters)
    pub font_index: Option<u16>,
    /// Font index for Far East characters
    pub font_index_fe: Option<u16>,
    /// Font index for other characters
    pub font_index_other: Option<u16>,
    /// Text color (RGB)
    pub color: Option<(u8, u8, u8)>,
    /// Color of the text underline.
    pub underline_color: Option<CharacterColor>,
    /// Border drawn on all four sides of this character run.
    pub border: Option<CharacterBorder>,
    /// Background shading for this character run.
    pub shading: Option<CharacterShading>,
    /// Highlight color
    pub highlight: Option<HighlightColor>,
    /// Superscript/subscript
    pub vertical_position: VerticalPosition,
    /// Vertical offset relative to the normal baseline, in signed half-points.
    pub position: CharacterPosition,
    /// Word-breaking behavior used when this run is hyphenated.
    pub hyphenation: HresiOperand,
    /// Animated text effect applied to this run.
    pub text_effect: TextEffect,
    /// Small caps
    pub is_small_caps: Option<bool>,
    /// All caps
    pub is_all_caps: Option<bool>,
    /// Hidden text
    pub is_hidden: Option<bool>,
    /// OLE2 object flag
    pub is_ole2: bool,
    /// Object flag (fObj)
    pub is_obj: bool,
    /// Special character flag (fSpec)
    pub is_spec: bool,
    /// Data flag (fData) - if true, pic_offset points to NilPICFAndBinData, not picture
    pub is_data: bool,
    /// Picture offset for embedded objects (fc in Data stream)
    pub pic_offset: Option<u32>,
    /// Object offset (fcObj)
    pub obj_offset: Option<u32>,
    /// Outline (hollow)
    pub is_outline: Option<bool>,
    /// Shadow
    pub is_shadow: Option<bool>,
    /// Embossed
    pub is_emboss: Option<bool>,
    /// Imprinted (engraved)
    pub is_imprint: Option<bool>,
    /// Character spacing in twips
    pub char_spacing: Option<i16>,
    /// Kerning in half-points
    pub kerning: Option<u16>,
    /// Character scale percentage
    pub char_scale: Option<u16>,
    /// Language ID
    pub language_id: Option<u16>,
    /// Far East language ID.
    pub language_id_fe: Option<u16>,
    /// Whether this run uses bidirectional/complex-script formatting.
    pub is_bidi: Option<bool>,
    /// Complex-script bold setting.
    pub is_bold_bidi: Option<bool>,
    /// Complex-script italic setting.
    pub is_italic_bidi: Option<bool>,
    /// Complex-script font index.
    pub font_index_bidi: Option<u16>,
    /// Complex-script language ID.
    pub language_id_bidi: Option<u16>,
    /// Complex-script indexed text color.
    pub color_index_bidi: Option<u16>,
    /// Complex-script font size in half-points.
    pub font_size_bidi: Option<u16>,
    /// Font/language bias for characters shared by multiple scripts.
    pub script_hint: Option<CharacterScriptHint>,
    /// Whether spelling and grammar proofing excludes this run.
    pub is_no_proof: Option<bool>,
    /// Whether complex-script formatting is forced for this run.
    pub is_complex_scripts: Option<bool>,
    /// Style index (istd)
    pub style_index: Option<u16>,
    /// Whether character properties before a tracked change are preserved.
    pub properties_preserved_for_revision: bool,
    /// Character state immediately before the active `sprmCWall` boundary.
    pub preserved_properties_for_revision: Option<Box<CharacterProperties>>,
    /// Conditional character formatting definitions carried by a table style.
    pub conditional_formats: Vec<CharacterConditionalFormatting>,
    /// Vanish (hidden)
    pub is_vanish: Option<bool>,
    /// Whether this run is marked as an inserted revision.
    pub is_revision_inserted: Option<bool>,
    /// Insertion revision author index in `SttbfRMark`.
    pub revision_author_index: Option<u16>,
    /// Packed insertion revision DTTM.
    pub revision_timestamp: Option<u32>,
    /// Insertion or modification reason code (`sprmCIdslRMark`).
    pub revision_id: Option<u16>,
    /// Character-formatting revision-save ID.
    pub formatting_revision_save_id: Option<u32>,
    /// Insertion revision-save ID.
    pub insertion_revision_save_id: Option<u32>,
    /// Whether this run is marked as a deleted revision.
    pub is_revision_deleted: Option<bool>,
    /// Deletion revision author index in `SttbfRMark`.
    pub deletion_author_index: Option<u16>,
    /// Packed deletion revision DTTM.
    pub deletion_timestamp: Option<u32>,
    /// Deletion reason code (`sprmCIdslRMarkDel`).
    pub deletion_revision_id: Option<u16>,
    /// Deletion revision-save ID.
    pub deletion_revision_save_id: Option<u32>,
    /// Whether this run has a tracked character-formatting change.
    pub has_formatting_revision: Option<bool>,
    /// Formatting revision author index in `SttbfRMark`.
    pub formatting_revision_author_index: Option<u16>,
    /// Packed formatting revision DTTM.
    pub formatting_revision_timestamp: Option<u32>,
    /// Revision metadata for a LISTNUM display-field result.
    pub display_field_revision: Option<DisplayFieldRevisionProperties>,
}

/// Vertical text position in signed half-points (`sprmCHpsPos`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CharacterPosition(i16);

impl CharacterPosition {
    /// Normal baseline position.
    pub const NORMAL: Self = Self(0);

    /// Construct a validated vertical position in half-points.
    pub fn new(half_points: i16) -> Result<Self> {
        if !(-3168..=3168).contains(&half_points) {
            return Err(PackageError::Corrupted(format!(
                "sprmCHpsPos value {half_points} is outside -3168..=3168"
            )));
        }
        Ok(Self(half_points))
    }

    /// Return the signed half-point offset.
    pub const fn half_points(self) -> i16 {
        self.0
    }
}

/// Word-breaking method stored in an [`HresiOperand`].
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum HyphenationMode {
    /// Insert a hyphen and continue on the next line.
    #[default]
    Normal,
    /// Add the replacement character before the hyphen.
    AddBefore,
    /// Change the character before the hyphen.
    ChangeBefore,
    /// Delete the character before the hyphen.
    DeleteBefore,
    /// Change the character after the hyphen.
    ChangeAfter,
    /// Delete two characters before the hyphen and replace them.
    DeleteAndChange,
}

/// Validated two-byte `HresiOperand` used by `sprmCHresi`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct HresiOperand {
    mode: HyphenationMode,
    replacement_character: Option<u8>,
}

impl HresiOperand {
    /// Normal word-breaking, whose dependent `ChHres` byte is zero.
    pub const fn normal() -> Self {
        Self {
            mode: HyphenationMode::Normal,
            replacement_character: None,
        }
    }

    /// Construct a non-normal word-breaking operand with a printable ASCII byte.
    pub fn with_character(mode: HyphenationMode, replacement_character: u8) -> Result<Self> {
        if mode == HyphenationMode::Normal {
            return Err(PackageError::Corrupted(
                "normal HresiOperand cannot have a replacement character".to_string(),
            ));
        }
        if !replacement_character.is_ascii_graphic() && replacement_character != b' ' {
            return Err(PackageError::Corrupted(format!(
                "sprmCHresi ChHres byte 0x{replacement_character:02X} is not printable ASCII"
            )));
        }
        Ok(Self {
            mode,
            replacement_character: Some(replacement_character),
        })
    }

    /// Return the word-breaking mode.
    pub const fn mode(self) -> HyphenationMode {
        self.mode
    }

    /// Return the dependent printable ASCII byte, or `None` for normal breaking.
    pub const fn replacement_character(self) -> Option<u8> {
        self.replacement_character
    }
}

impl Default for HresiOperand {
    fn default() -> Self {
        Self::normal()
    }
}

/// Animated text effect stored by `sprmCSfxText`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum TextEffect {
    /// No text effect.
    #[default]
    None = 0,
    /// Las Vegas Lights.
    LasVegasLights = 1,
    /// Blinking background.
    BlinkingBackground = 2,
    /// Sparkle Text.
    SparkleText = 3,
    /// Marching Black Ants.
    MarchingBlackAnts = 4,
    /// Marching Red Ants.
    MarchingRedAnts = 5,
    /// Shimmer.
    Shimmer = 6,
}

impl TryFrom<u8> for TextEffect {
    type Error = PackageError;

    fn try_from(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::LasVegasLights),
            2 => Ok(Self::BlinkingBackground),
            3 => Ok(Self::SparkleText),
            4 => Ok(Self::MarchingBlackAnts),
            5 => Ok(Self::MarchingRedAnts),
            6 => Ok(Self::Shimmer),
            _ => Err(PackageError::Corrupted(format!(
                "sprmCSfxText has invalid text effect {value}"
            ))),
        }
    }
}

impl From<TextEffect> for u8 {
    fn from(value: TextEffect) -> Self {
        value as u8
    }
}

/// Conditional character formatting carried by `sprmCCnf` in a table style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CharacterConditionalFormatting {
    /// Table location or band for which the nested properties apply.
    pub condition: TableStyleCondition,
    /// Typed character properties decoded from the nested grpprl.
    pub properties: Box<CharacterProperties>,
    /// Exact nested grpprl retained for lossless preservation.
    pub raw_grpprl: Vec<u8>,
}

/// Parsed `DispFldRmOperand` state for a LISTNUM display-field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DisplayFieldRevisionProperties {
    /// Whether the field result contains a revision.
    pub active: bool,
    /// Revision author index in `SttbfRMark`.
    pub author_index: u16,
    /// Packed revision DTTM.
    pub timestamp: u32,
    /// Previous LISTNUM field result.
    pub previous_result: String,
}

/// Underline styles supported in DOC format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UnderlineStyle {
    /// No underline
    #[default]
    None,
    /// Single underline
    Single,
    /// Double underline
    Double,
    /// Dotted underline
    Dotted,
    /// Dashed underline
    Dashed,
    /// Wavy underline
    Wavy,
    /// Thick underline
    Thick,
    /// Word-only underline (skip spaces)
    WordsOnly,
    /// Dash-dot underline
    DashDot,
    /// Dash-dot-dot underline
    DashDotDot,
}

/// Highlight colors available in DOC format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HighlightColor {
    None,
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    Yellow,
    White,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
}

/// Script bias stored by `sprmCIdctHint`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterScriptHint {
    /// Bias toward non-Far-East properties.
    Default,
    /// Bias toward Far East properties.
    FarEast,
    /// Bias toward complex-script properties.
    ComplexScript,
    /// A reserved value retained for forward compatibility.
    Reserved(u8),
}

/// A color stored in a legacy Word character property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterColor {
    /// The application chooses a context-appropriate color.
    Automatic,
    /// An explicit red, green, and blue color.
    Rgb(u8, u8, u8),
    /// An invalid or future palette index retained without interpretation.
    ReservedPaletteIndex(u8),
}

/// Border formatting applied around a character run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterBorder {
    /// Border color, including the context-dependent automatic color.
    pub color: CharacterColor,
    /// Border width in eighths of a point for ordinary character borders.
    pub width: u8,
    /// Border line style.
    pub style: CharacterBorderStyle,
    /// Distance between the border and text, in points.
    pub spacing: u8,
    /// Whether Word adds a shadow to the border.
    pub has_shadow: bool,
    /// Whether Word reverses the border to create a frame effect.
    pub has_frame: bool,
}

/// Line styles supported by MS-DOC character borders.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterBorderStyle {
    Single,
    Double,
    Thin,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
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
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
    /// An invalid or future style retained without coercion.
    Reserved(u8),
}

impl From<u8> for CharacterBorderStyle {
    fn from(value: u8) -> Self {
        match value {
            0x01 => Self::Single,
            0x03 => Self::Double,
            0x05 => Self::Thin,
            0x06 => Self::Dotted,
            0x07 => Self::Dashed,
            0x08 => Self::DotDash,
            0x09 => Self::DotDotDash,
            0x0A => Self::Triple,
            0x0B => Self::ThinThickSmallGap,
            0x0C => Self::ThickThinSmallGap,
            0x0D => Self::ThinThickThinSmallGap,
            0x0E => Self::ThinThickMediumGap,
            0x0F => Self::ThickThinMediumGap,
            0x10 => Self::ThinThickThinMediumGap,
            0x11 => Self::ThinThickLargeGap,
            0x12 => Self::ThickThinLargeGap,
            0x13 => Self::ThinThickThinLargeGap,
            0x14 => Self::Wave,
            0x15 => Self::DoubleWave,
            0x16 => Self::DashSmallGap,
            0x17 => Self::DashDotStroked,
            0x18 => Self::ThreeDEmboss,
            0x19 => Self::ThreeDEngrave,
            0x1A => Self::Outset,
            0x1B => Self::Inset,
            value => Self::Reserved(value),
        }
    }
}

/// Shading applied behind a character run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharacterShading {
    pub foreground_color: CharacterColor,
    pub background_color: CharacterColor,
    pub pattern: CharacterShadingPattern,
}

/// Shading patterns defined by the MS-DOC `Ipat` enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterShadingPattern {
    Clear,
    Solid,
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    DarkHorizontal,
    DarkVertical,
    DarkForwardDiagonal,
    DarkBackwardDiagonal,
    DarkCross,
    DarkDiagonalCross,
    Horizontal,
    Vertical,
    ForwardDiagonal,
    BackwardDiagonal,
    Cross,
    DiagonalCross,
    Percent2_5,
    Percent7_5,
    Percent12_5,
    Percent15,
    Percent17_5,
    Percent22_5,
    Percent27_5,
    Percent32_5,
    Percent35,
    Percent37_5,
    Percent42_5,
    Percent45,
    Percent47_5,
    Percent52_5,
    Percent55,
    Percent57_5,
    Percent62_5,
    Percent65,
    Percent67_5,
    Percent72_5,
    Percent77_5,
    Percent82_5,
    Percent85,
    Percent87_5,
    Percent92_5,
    Percent95,
    Percent97_5,
    Nil,
    /// An invalid or future pattern retained without coercion.
    Reserved(u16),
}

impl From<u16> for CharacterShadingPattern {
    fn from(value: u16) -> Self {
        match value {
            0x0000 => Self::Clear,
            0x0001 => Self::Solid,
            0x0002 => Self::Percent5,
            0x0003 => Self::Percent10,
            0x0004 => Self::Percent20,
            0x0005 => Self::Percent25,
            0x0006 => Self::Percent30,
            0x0007 => Self::Percent40,
            0x0008 => Self::Percent50,
            0x0009 => Self::Percent60,
            0x000A => Self::Percent70,
            0x000B => Self::Percent75,
            0x000C => Self::Percent80,
            0x000D => Self::Percent90,
            0x000E => Self::DarkHorizontal,
            0x000F => Self::DarkVertical,
            0x0010 => Self::DarkForwardDiagonal,
            0x0011 => Self::DarkBackwardDiagonal,
            0x0012 => Self::DarkCross,
            0x0013 => Self::DarkDiagonalCross,
            0x0014 => Self::Horizontal,
            0x0015 => Self::Vertical,
            0x0016 => Self::ForwardDiagonal,
            0x0017 => Self::BackwardDiagonal,
            0x0018 => Self::Cross,
            0x0019 => Self::DiagonalCross,
            0x0023 => Self::Percent2_5,
            0x0024 => Self::Percent7_5,
            0x0025 => Self::Percent12_5,
            0x0026 => Self::Percent15,
            0x0027 => Self::Percent17_5,
            0x0028 => Self::Percent22_5,
            0x0029 => Self::Percent27_5,
            0x002A => Self::Percent32_5,
            0x002B => Self::Percent35,
            0x002C => Self::Percent37_5,
            0x002D => Self::Percent42_5,
            0x002E => Self::Percent45,
            0x002F => Self::Percent47_5,
            0x0030 => Self::Percent52_5,
            0x0031 => Self::Percent55,
            0x0032 => Self::Percent57_5,
            0x0033 => Self::Percent62_5,
            0x0034 => Self::Percent65,
            0x0035 => Self::Percent67_5,
            0x0036 => Self::Percent72_5,
            0x0037 => Self::Percent77_5,
            0x0038 => Self::Percent82_5,
            0x0039 => Self::Percent85,
            0x003A => Self::Percent87_5,
            0x003B => Self::Percent92_5,
            0x003C => Self::Percent95,
            0x003D => Self::Percent97_5,
            0xFFFF => Self::Nil,
            value => Self::Reserved(value),
        }
    }
}

impl From<u8> for CharacterScriptHint {
    fn from(value: u8) -> Self {
        match value {
            0 => Self::Default,
            1 => Self::FarEast,
            2 => Self::ComplexScript,
            value => Self::Reserved(value),
        }
    }
}

// Re-export common VerticalPosition type
pub use litchi_core::VerticalPosition;

impl CharacterProperties {
    /// Create a new CharacterProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if any formatting is applied.
    pub fn has_formatting(&self) -> bool {
        self.is_bold.is_some()
            || self.is_italic.is_some()
            || self.underline != UnderlineStyle::None
            || self.is_strikethrough.is_some()
            || self.font_size.is_some()
            || self.color.is_some()
            || self.underline_color.is_some()
            || self.border.is_some()
            || self.shading.is_some()
            || self.highlight.is_some()
            || self.vertical_position != VerticalPosition::Normal
            || self.position != CharacterPosition::NORMAL
            || self.hyphenation != HresiOperand::normal()
            || self.text_effect != TextEffect::None
            || self.is_bidi.is_some()
            || self.is_bold_bidi.is_some()
            || self.is_italic_bidi.is_some()
            || self.font_index_bidi.is_some()
            || self.language_id_bidi.is_some()
            || self.color_index_bidi.is_some()
            || self.font_size_bidi.is_some()
            || self.language_id_fe.is_some()
            || self.script_hint.is_some()
            || self.is_no_proof.is_some()
            || self.is_complex_scripts.is_some()
            || self.properties_preserved_for_revision
            || !self.conditional_formats.is_empty()
    }
}
