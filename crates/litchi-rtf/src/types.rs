//! RTF document type definitions.

use super::border::{Borders, CharacterBorder, CharacterShading, Shading};
use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::num::NonZeroU16;

/// Font reference (index into font table).
pub type FontRef = u16;

/// Color reference (index into color table).
pub type ColorRef = u16;

/// RTF color representation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Color {
    /// Red component (0-255)
    pub red: u8,
    /// Green component (0-255)
    pub green: u8,
    /// Blue component (0-255)
    pub blue: u8,
}

impl Color {
    /// Create a new color.
    #[inline]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    /// Black color.
    #[inline]
    pub const fn black() -> Self {
        Self::new(0, 0, 0)
    }

    /// White color.
    #[inline]
    pub const fn white() -> Self {
        Self::new(255, 255, 255)
    }
}

/// Color table containing document colors.
#[derive(Debug, Clone)]
pub struct ColorTable {
    colors: Vec<Color>,
}

impl ColorTable {
    /// Create a new color table.
    #[inline]
    pub fn new() -> Self {
        Self { colors: Vec::new() }
    }

    /// Add a color to the table and return its index.
    #[inline]
    pub fn add(&mut self, color: Color) -> ColorRef {
        let index = self.colors.len() as ColorRef;
        self.colors.push(color);
        index
    }

    /// Get a color by reference.
    #[inline]
    pub fn get(&self, color_ref: ColorRef) -> Option<&Color> {
        self.colors.get(color_ref as usize)
    }

    /// Get all colors in the table.
    #[inline]
    pub fn colors(&self) -> &[Color] {
        &self.colors
    }
}

impl Default for ColorTable {
    fn default() -> Self {
        Self::new()
    }
}

/// Font family categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FontFamily {
    /// Nil (unknown or default)
    #[default]
    Nil,
    /// Roman (serif) fonts
    Roman,
    /// Swiss (sans-serif) fonts
    Swiss,
    /// Modern (monospace) fonts
    Modern,
    /// Script fonts
    Script,
    /// Decorative fonts
    Decor,
    /// Technical, symbol, and mathematical fonts
    Tech,
}

/// Font pitch preference from `fprq`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FontPitch {
    #[default]
    Default,
    Fixed,
    Variable,
}

/// Embedded font format from the `fontemb` destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum EmbeddedFontFormat {
    /// `ftnil` — unknown or unspecified font format.
    #[default]
    Nil,
    /// `fttruetype` — TrueType font data.
    TrueType,
}

/// Embedded font payload from the inert `fontemb` destination.
///
/// A font entry may embed the font bytes directly (hexadecimal or `bin`
/// payload) or reference an external font file through the `fontfile`
/// destination; both carriers are optional in the specification.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct EmbeddedFont<'a> {
    /// Declared embedded font format.
    pub format: EmbeddedFontFormat,
    /// File name from the nested `fontfile` destination, if present.
    pub file_name: Option<Cow<'a, str>>,
    /// Code page of the `fontfile` name from its `cpg` control word.
    pub file_code_page: Option<u16>,
    /// Decoded embedded font bytes, if the data is carried inline.
    pub data: Option<Vec<u8>>,
}

impl EmbeddedFont<'_> {
    /// Maximum accepted embedded font payload (32 MiB).
    pub const MAX_DATA_BYTES: usize = 32 * 1_048_576;
    /// Maximum accepted `fontfile` name length in bytes.
    pub const MAX_FILE_NAME_BYTES: usize = 4_096;

    pub fn validate(&self) -> RtfResult<()> {
        if self
            .file_name
            .as_ref()
            .is_some_and(|name| name.is_empty() || name.len() > Self::MAX_FILE_NAME_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "invalid or oversized RTF embedded font file name".to_string(),
            ));
        }
        if self
            .data
            .as_ref()
            .is_some_and(|data| data.is_empty() || data.len() > Self::MAX_DATA_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "invalid or oversized RTF embedded font data".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> EmbeddedFont<'static> {
        EmbeddedFont {
            format: self.format,
            file_name: self.file_name.map(|name| Cow::Owned(name.into_owned())),
            file_code_page: self.file_code_page,
            data: self.data,
        }
    }
}

/// Theme-font role of a font-table entry (`\flomajor` and friends).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontTheme {
    /// Major Latin-script theme font (`\flomajor`).
    MajorLatin,
    /// Major high-ANSI theme font (`\fhimajor`).
    MajorHighAnsi,
    /// Major double-byte (East Asian) theme font (`\fdbmajor`).
    MajorDoubleByte,
    /// Major bidi (complex scripts) theme font (`\fbimajor`).
    MajorBidi,
    /// Minor Latin-script theme font (`\flominor`).
    MinorLatin,
    /// Minor high-ANSI theme font (`\fhiminor`).
    MinorHighAnsi,
    /// Minor double-byte (East Asian) theme font (`\fdbminor`).
    MinorDoubleByte,
    /// Minor bidi (complex scripts) theme font (`\fbiminor`).
    MinorBidi,
}

impl FontTheme {
    /// The RTF control word selecting this role.
    pub fn control_word(self) -> &'static str {
        match self {
            Self::MajorLatin => "flomajor",
            Self::MajorHighAnsi => "fhimajor",
            Self::MajorDoubleByte => "fdbmajor",
            Self::MajorBidi => "fbimajor",
            Self::MinorLatin => "flominor",
            Self::MinorHighAnsi => "fhiminor",
            Self::MinorDoubleByte => "fdbminor",
            Self::MinorBidi => "fbiminor",
        }
    }

    /// The role for a control word, or `None` for an unknown selector.
    pub fn from_control_word(word: &str) -> Option<Self> {
        Some(match word {
            "flomajor" => Self::MajorLatin,
            "fhimajor" => Self::MajorHighAnsi,
            "fdbmajor" => Self::MajorDoubleByte,
            "fbimajor" => Self::MajorBidi,
            "flominor" => Self::MinorLatin,
            "fhiminor" => Self::MinorHighAnsi,
            "fdbminor" => Self::MinorDoubleByte,
            "fbiminor" => Self::MinorBidi,
            _ => return None,
        })
    }
}

/// Font definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font<'a> {
    /// Font name
    pub name: Cow<'a, str>,
    /// Font family category
    pub family: FontFamily,
    /// Character set (Windows codepage)
    pub charset: u8,
    /// Alternate font name from the inert `falt` destination.
    pub alternate_name: Option<Cow<'a, str>>,
    /// Non-tagged font name from the inert `fname` destination.
    pub non_tagged_name: Option<Cow<'a, str>>,
    /// Ten-byte PANOSE classification.
    pub panose: Option<[u8; 10]>,
    /// Pitch preference.
    pub pitch: FontPitch,
    /// Explicit font code page.
    pub code_page: Option<u16>,
    /// Embedded font payload from the inert `fontemb` destination.
    pub embedded: Option<EmbeddedFont<'a>>,
    /// Theme-font role from the major/minor theme selectors.
    pub theme: Option<FontTheme>,
}

impl<'a> Font<'a> {
    /// Create a new font.
    #[inline]
    pub fn new(name: Cow<'a, str>, family: FontFamily, charset: u8) -> Self {
        Self {
            name,
            family,
            charset,
            alternate_name: None,
            non_tagged_name: None,
            panose: None,
            pitch: FontPitch::Default,
            code_page: None,
            embedded: None,
            theme: None,
        }
    }

    pub fn validate(&self) -> RtfResult<()> {
        const MAX_FONT_NAME_BYTES: usize = 4_096;
        if self.name.is_empty()
            || self.name.len() > MAX_FONT_NAME_BYTES
            || self
                .alternate_name
                .as_ref()
                .is_some_and(|name| name.is_empty() || name.len() > MAX_FONT_NAME_BYTES)
            || self
                .non_tagged_name
                .as_ref()
                .is_some_and(|name| name.is_empty() || name.len() > MAX_FONT_NAME_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "invalid or oversized RTF font name".to_string(),
            ));
        }
        if let Some(embedded) = &self.embedded {
            embedded.validate()?;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> Font<'static> {
        Font {
            name: Cow::Owned(self.name.into_owned()),
            family: self.family,
            theme: self.theme,
            charset: self.charset,
            alternate_name: self
                .alternate_name
                .map(|name| Cow::Owned(name.into_owned())),
            non_tagged_name: self
                .non_tagged_name
                .map(|name| Cow::Owned(name.into_owned())),
            panose: self.panose,
            pitch: self.pitch,
            code_page: self.code_page,
            embedded: self.embedded.map(EmbeddedFont::into_owned),
        }
    }
}

/// Font table containing document fonts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontTable<'a> {
    pub(crate) fonts: Vec<Font<'a>>,
    pub(crate) defined: Vec<bool>,
}

impl<'a> FontTable<'a> {
    /// Create a new font table.
    #[inline]
    pub fn new() -> Self {
        Self {
            fonts: Vec::new(),
            defined: Vec::new(),
        }
    }

    /// Add a font to the table at a specific index.
    #[inline]
    pub fn insert(&mut self, index: FontRef, font: Font<'a>) {
        // Ensure the vector is large enough
        if index as usize >= self.fonts.len() {
            self.fonts.resize(
                (index as usize) + 1,
                Font::new(Cow::Borrowed(""), FontFamily::Nil, 0),
            );
            self.defined.resize((index as usize) + 1, false);
        }
        self.fonts[index as usize] = font;
        self.defined[index as usize] = true;
    }

    /// Get a font by reference.
    #[inline]
    pub fn get(&self, font_ref: FontRef) -> Option<&Font<'a>> {
        self.defined
            .get(font_ref as usize)
            .copied()
            .unwrap_or(false)
            .then(|| &self.fonts[font_ref as usize])
    }

    /// Get all fonts in the table.
    #[inline]
    pub fn fonts(&self) -> &[Font<'a>] {
        &self.fonts
    }

    pub fn is_defined(&self, font_ref: FontRef) -> bool {
        self.defined
            .get(font_ref as usize)
            .copied()
            .unwrap_or(false)
    }

    pub fn validate(&self) -> RtfResult<()> {
        if self.fonts.len() > 65_536 || self.defined.len() != self.fonts.len() {
            return Err(RtfError::MalformedDocument(
                "invalid RTF font-table size".to_string(),
            ));
        }
        const MAX_AGGREGATE_EMBEDDED_BYTES: usize = 256 * 1_048_576;
        let mut aggregate = 0usize;
        let mut embedded_aggregate = 0usize;
        for (font, defined) in self.fonts.iter().zip(&self.defined) {
            if !defined {
                continue;
            }
            font.validate()?;
            aggregate = aggregate
                .checked_add(font.name.len())
                .and_then(|total| {
                    total.checked_add(font.alternate_name.as_ref().map_or(0, |name| name.len()))
                })
                .and_then(|total| {
                    total.checked_add(font.non_tagged_name.as_ref().map_or(0, |name| name.len()))
                })
                .ok_or_else(|| {
                    RtfError::MalformedDocument("RTF font-table text size overflow".to_string())
                })?;
            embedded_aggregate = embedded_aggregate
                .checked_add(font.embedded.as_ref().map_or(0, |embedded| {
                    embedded.data.as_ref().map_or(0, |data| data.len())
                }))
                .ok_or_else(|| {
                    RtfError::MalformedDocument("RTF font-table embedded size overflow".to_string())
                })?;
        }
        if aggregate > 16 * 1_048_576 {
            return Err(RtfError::MalformedDocument(
                "RTF font-table text exceeds the safety limit".to_string(),
            ));
        }
        if embedded_aggregate > MAX_AGGREGATE_EMBEDDED_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF font-table embedded fonts exceed the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> FontTable<'static> {
        FontTable {
            fonts: self.fonts.into_iter().map(Font::into_owned).collect(),
            defined: self.defined,
        }
    }
}

impl<'a> Default for FontTable<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// Text alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Alignment {
    /// Left-aligned
    #[default]
    Left,
    /// Right-aligned
    Right,
    /// Centered
    Center,
    /// Justified
    Justify,
}

/// Spacing information for paragraphs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Spacing {
    /// Space before paragraph (in twips, 1/20th of a point)
    pub before: i32,
    /// Space after paragraph (in twips)
    pub after: i32,
    /// Line spacing (in twips)
    pub line: i32,
    /// Line spacing multiplier
    pub line_multiple: bool,
}

/// Paragraph spacing policy layered over raw twip spacing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphSpacingPolicy {
    /// `sbauto1` makes space-before automatic and ignores `sbN`.
    pub automatic_before: bool,
    /// `saauto1` makes space-after automatic and ignores `saN`.
    pub automatic_after: bool,
    /// `lisbN`, in hundredths of a character unit; overrides `sbN` when present.
    pub list_before: Option<u32>,
    /// `lisaN`, in hundredths of a character unit; overrides `saN` when present.
    pub list_after: Option<u32>,
    /// Whether lines snap to the document grid; disabled by `nosnaplinegrid`.
    pub snap_to_line_grid: bool,
    /// Suppress before/after spacing between adjacent paragraphs of the same style.
    pub contextual_spacing: bool,
}

impl Default for ParagraphSpacingPolicy {
    fn default() -> Self {
        Self {
            automatic_before: false,
            automatic_after: false,
            list_before: None,
            list_after: None,
            snap_to_line_grid: true,
            contextual_spacing: false,
        }
    }
}

/// Indentation information for paragraphs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Indentation {
    /// Left indent (in twips)
    pub left: i32,
    /// Right indent (in twips)
    pub right: i32,
    /// First line indent (in twips)
    pub first_line: i32,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ParagraphLogicalIndentation {
    pub start: Option<i32>,
    pub end: Option<i32>,
    pub first_line_character_units: Option<i32>,
    pub left_character_units: Option<i32>,
    pub right_character_units: Option<i32>,
    pub mirrored: bool,
}

/// Paragraph wrapping policy from `wrapdefault`, `nocwrap`, `nowwrap`, and `nooverflow`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ParagraphWrapping {
    #[default]
    Default,
    NoCharacterWrap,
    NoWordWrap,
    NoOverflow,
}

/// East Asian paragraph font alignment from the RTF 1.9.1 `fa*` selectors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ParagraphFontAlignment {
    #[default]
    Auto,
    Hanging,
    Center,
    Roman,
    Variable,
    Fixed,
}

/// Effective paragraph line-breaking and automatic-spacing policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ParagraphLineBreaking {
    /// Automatic paragraph hyphenation (`hyphpar`).
    pub automatic_hyphenation: bool,
    /// Automatic spacing between Asian and alphabetic text (`aspalpha`).
    pub auto_space_alphabetic: bool,
    /// Automatic spacing between Asian text and numbers (`aspnum`).
    pub auto_space_numbers: bool,
    /// Adjust the right indent for a document grid (`adjustright`).
    pub adjust_right_indent: bool,
    /// Wrapping/overflow selector.
    pub wrapping: ParagraphWrapping,
    /// East Asian font-alignment selector.
    pub font_alignment: ParagraphFontAlignment,
}

/// Maximum supported number of lines occupied by a paragraph drop cap.
pub const MAX_PARAGRAPH_DROP_CAP_LINES: u16 = 255;

/// Placement selected by the RTF `\\dropcaptN` paragraph property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphDropCapKind {
    /// Drop cap remains within the regular text margin (`\\dropcapt1`).
    InText,
    /// Drop cap is placed in the left margin (`\\dropcapt2`).
    Margin,
}

impl ParagraphDropCapKind {
    /// Return the canonical RTF numeric value.
    pub const fn as_rtf_value(self) -> i32 {
        match self {
            Self::InText => 1,
            Self::Margin => 2,
        }
    }
}

/// Complete, validated paragraph drop-cap settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphDropCap {
    kind: ParagraphDropCapKind,
    line_count: u8,
}

impl ParagraphDropCap {
    /// Construct a complete drop cap.
    pub fn new(kind: ParagraphDropCapKind, line_count: u16) -> crate::RtfResult<Self> {
        if !(1..=MAX_PARAGRAPH_DROP_CAP_LINES).contains(&line_count) {
            return Err(crate::RtfError::InvalidStructure(format!(
                "RTF drop-cap line count must be in 1..={MAX_PARAGRAPH_DROP_CAP_LINES}"
            )));
        }
        Ok(Self {
            kind,
            line_count: line_count as u8,
        })
    }

    /// Drop-cap placement.
    pub const fn kind(self) -> ParagraphDropCapKind {
        self.kind
    }

    /// Number of text lines occupied by the drop cap.
    pub const fn line_count(self) -> u8 {
        self.line_count
    }

    /// Validate the model before serialization.
    pub fn validate(self) -> crate::RtfResult<()> {
        if self.line_count == 0 {
            return Err(crate::RtfError::InvalidStructure(
                "RTF drop-cap line count must be nonzero".to_string(),
            ));
        }
        Ok(())
    }
}

/// Author and packed DTTM timestamp attached to a structural revision
/// marker such as `\prauthN`/`\prdateN` or `\srauthN`/`\srdateN`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RevisionMetadata {
    /// Index into the document's revision-author table (`revtbl`).
    pub author: Option<i32>,
    /// Packed signed RTF DTTM value, as used by `\revdttmN`.
    pub date: Option<i32>,
}

impl RevisionMetadata {
    /// Whether neither an author nor a date was authored.
    pub fn is_empty(&self) -> bool {
        self.author.is_none() && self.date.is_none()
    }

    /// Validate the author index against the RTF domain.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.author.is_some_and(|author| author < 0) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF revision author index cannot be negative".to_string(),
            ));
        }
        Ok(())
    }
}

/// Paragraph properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Paragraph {
    /// Applied paragraph-style handle from `\\sN`.
    ///
    /// The reference is retained as inert provenance; concrete paragraph and
    /// character properties remain independently represented.
    pub paragraph_style: Option<u16>,
    /// RSID attached to the paragraph formatting (`\pararsidN`).
    pub paragraph_rsid: Option<u32>,
    /// Paragraph outline level 0-9 for headings and TOC-by-outline
    /// (`\outlinelevelN`).
    pub outline_level: Option<u8>,
    /// Explicit paragraph direction; `None` uses left-to-right precedence.
    pub direction: Option<TextDirection>,
    /// Text alignment
    pub alignment: Alignment,
    /// Spacing
    pub spacing: Spacing,
    /// Automatic, list-unit, contextual, and grid spacing policy.
    pub spacing_policy: ParagraphSpacingPolicy,
    /// Indentation
    pub indentation: Indentation,
    pub logical_indentation: ParagraphLogicalIndentation,
    /// Custom tab stops in RTF declaration order
    pub tab_stops: super::border::TabStops,
    /// Borders
    pub borders: Borders,
    /// Shading/background
    pub shading: Shading,
    /// Keep paragraph on one page
    pub keep_together: bool,
    /// Keep with next paragraph
    pub keep_next: bool,
    /// Page break before
    pub page_break_before: bool,
    /// Widow/orphan control
    pub widow_control: bool,
    /// Complete drop-cap settings from `\\dropcapliN` and `\\dropcaptN`.
    pub drop_cap: Option<ParagraphDropCap>,
    /// Line-breaking and automatic-spacing policy.
    pub line_breaking: ParagraphLineBreaking,
    /// List override index (`\lsN`) applied to this paragraph
    pub list_override: Option<i32>,
    /// Zero-based list level (`\ilvlN`) applied to this paragraph
    pub list_level: Option<u8>,
    /// Index into the document's ordered inert legacy `pn` record table.
    pub legacy_numbering: Option<u32>,
    /// Author/date metadata for the revision that inserted this paragraph
    /// (`\prauthN`, `\prdateN`).
    pub revision: RevisionMetadata,
}

impl Paragraph {
    /// Set or clear the applied paragraph-style handle.
    #[inline]
    pub fn set_paragraph_style(&mut self, value: Option<u16>) {
        self.paragraph_style = value;
    }

    /// Replace the paragraph shading after validating its RTF domain.
    pub fn set_shading(&mut self, shading: Shading) -> crate::RtfResult<()> {
        shading.validate()?;
        self.shading = shading;
        Ok(())
    }

    /// Clear all explicit paragraph-shading controls.
    #[inline]
    pub fn clear_shading(&mut self) {
        self.shading.clear();
    }
}

/// Underline style
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
    /// Dash-dot underline
    DashDot,
    /// Dash-dot-dot underline
    DashDotDot,
    /// Word-only underline
    Words,
    /// Thick underline
    Thick,
    /// Wave underline
    Wave,
    /// Hairline underline
    Hairline,
    /// Thick dotted underline
    ThickDotted,
    /// Thick dashed underline
    ThickDashed,
    /// Thick dash-dot underline
    ThickDashDot,
    /// Thick dash-dot-dot underline
    ThickDashDotDot,
    /// Thick long-dash underline
    ThickLongDash,
    /// Long-dash underline
    LongDash,
    /// Heavy wave underline
    HeavyWave,
    /// Double wave underline
    DoubleWave,
}

/// Explicit bidirectional precedence for a character run or paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextDirection {
    LeftToRight,
    RightToLeft,
}

/// Character repertoire selected for a formatted run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterType {
    /// Low ANSI characters (`\\loch`, byte range 0x00 through 0x7f).
    LowAnsi,
    /// High ANSI characters (`\\hich`, byte range 0x80 through 0xff).
    HighAnsi,
    /// Double-byte characters (`\\dbch`).
    DoubleByte,
}

/// Exact form of the East Asian character-grid control.
///
/// Parameterless `\\cgrid` is retained separately from a numeric value because
/// both forms are emitted by common RTF producers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterGrid {
    /// A parameterless `\\cgrid` control.
    Parameterless,
    /// An explicit signed 16-bit `\\cgridN` value.
    Value(i16),
}

/// Associated-font baseline selected by `\\aupN` or `\\adnN`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AssociatedCharacterBaseline {
    /// Raise the associated font by the given number of half-points.
    RaisedHalfPoints(u16),
    /// Lower the associated font by the given number of half-points.
    LoweredHalfPoints(u16),
}

/// Underline styles defined for an associated font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AssociatedUnderlineStyle {
    None,
    Single,
    Dotted,
    Double,
    Words,
}

/// Associated character properties used for complex-script text.
///
/// Optional fields preserve the distinction between an omitted property and an
/// explicit off value such as `\ab0` or `\ai0`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AssociatedCharacterFormatting {
    /// Associated bold override from `\\ab` or `\\ab0`.
    pub bold: Option<bool>,
    /// Associated all-capitals override from `\\acaps` or `\\acaps0`.
    pub all_caps: Option<bool>,
    /// Associated foreground color from `\\acfN`.
    pub color_ref: Option<ColorRef>,
    /// Associated baseline from `\\adnN` or `\\aupN`.
    pub baseline: Option<AssociatedCharacterBaseline>,
    /// Associated character expansion in quarter-points from `\\aexpndN`.
    pub expansion_quarter_points: Option<i16>,
    /// Associated font reference from `\afN`.
    pub font_ref: Option<FontRef>,
    /// Associated font size in half-points from `\afsN`.
    pub font_size: Option<NonZeroU16>,
    /// Associated complex-script language from `\alangN`.
    pub language: Option<crate::LanguageId>,
    /// Associated italic override from `\ai` or `\ai0`.
    pub italic: Option<bool>,
    /// Associated outline override from `\\aoutl` or `\\aoutl0`.
    pub outline: Option<bool>,
    /// Associated small-capitals override from `\\ascaps` or `\\ascaps0`.
    pub small_caps: Option<bool>,
    /// Associated shadow override from `\\ashad` or `\\ashad0`.
    pub shadow: Option<bool>,
    /// Associated strikethrough override from `\\astrike` or `\\astrike0`.
    pub strike: Option<bool>,
    /// Associated underline selector, including an explicit no-underline value.
    pub underline: Option<AssociatedUnderlineStyle>,
}

impl AssociatedCharacterFormatting {
    /// Replace this metadata with the omitted/default state.
    #[inline]
    pub fn clear(&mut self) {
        *self = Self::default();
    }

    /// Set or clear the associated baseline after validating the RTF domain.
    pub fn set_baseline(
        &mut self,
        value: Option<AssociatedCharacterBaseline>,
    ) -> crate::RtfResult<()> {
        if matches!(
            value,
            Some(
                AssociatedCharacterBaseline::RaisedHalfPoints(value)
                    | AssociatedCharacterBaseline::LoweredHalfPoints(value)
            ) if i32::from(value) > crate::MAX_CHARACTER_BASELINE_HALF_POINTS
        ) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF associated character baseline is out of range".to_string(),
            ));
        }
        self.baseline = value;
        Ok(())
    }

    /// Set or clear associated quarter-point expansion.
    pub fn set_expansion_quarter_points(&mut self, value: Option<i32>) -> crate::RtfResult<()> {
        self.expansion_quarter_points = match value {
            None => None,
            Some(value)
                if (-crate::MAX_CHARACTER_EXPANSION..=crate::MAX_CHARACTER_EXPANSION)
                    .contains(&value) =>
            {
                Some(value as i16)
            },
            Some(_) => {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF associated character expansion is out of range".to_string(),
                ));
            },
        };
        Ok(())
    }

    /// Validate metadata before canonical serialization.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if let Some(baseline) = self.baseline {
            let value = match baseline {
                AssociatedCharacterBaseline::RaisedHalfPoints(value)
                | AssociatedCharacterBaseline::LoweredHalfPoints(value) => value,
            };
            if i32::from(value) > crate::MAX_CHARACTER_BASELINE_HALF_POINTS {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF associated character baseline is out of range".to_string(),
                ));
            }
        }
        if self.expansion_quarter_points.is_some_and(|value| {
            i32::from(value).unsigned_abs() > crate::MAX_CHARACTER_EXPANSION as u32
        }) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF associated character expansion is out of range".to_string(),
            ));
        }
        Ok(())
    }
}

/// Character formatting properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Formatting {
    /// Typed baseline, expansion, scaling, and kerning state.
    pub character_positioning: crate::CharacterPositioning,
    /// Applied character-style handle from `\\csN`.
    ///
    /// RTF requires the concrete style properties to accompany the reference,
    /// so this is retained as inert provenance rather than resolved into the
    /// surrounding formatting.
    pub character_style: Option<u16>,
    /// RSID of the revision that inserted this text (`\insrsidN`).
    pub insert_rsid: Option<u32>,
    /// RSID of the revision that deleted this text (`\delrsidN`).
    pub delete_rsid: Option<u32>,
    /// RSID attached to the character formatting (`\charrsidN`).
    pub char_style_rsid: Option<u32>,
    /// Underline color reference from `\ulcN`; `None` uses the text color.
    pub underline_color: Option<ColorRef>,
    /// Font reference
    pub font_ref: FontRef,
    /// Font size in half-points
    pub font_size: NonZeroU16,
    /// Color reference
    pub color_ref: ColorRef,
    /// Exact background color reference from `\cbN`; `None` preserves omission.
    pub background_color: Option<ColorRef>,
    /// Background/highlight color reference
    pub highlight_color: Option<ColorRef>,
    /// Character border introduced by `chbrdr`.
    pub character_border: Option<CharacterBorder>,
    /// Exact character shading controls.
    pub character_shading: Option<CharacterShading>,
    /// Bold
    pub bold: bool,
    /// Italic
    pub italic: bool,
    /// Underline style
    pub underline: UnderlineStyle,
    /// Strikethrough
    pub strike: bool,
    /// Double strikethrough
    pub double_strike: bool,
    /// Superscript
    pub superscript: bool,
    /// Subscript
    pub subscript: bool,
    /// Small caps
    pub smallcaps: bool,
    /// All caps
    pub all_caps: bool,
    /// Hidden text
    pub hidden: bool,
    /// Outline
    pub outline: bool,
    /// Shadow
    pub shadow: bool,
    /// Emboss
    pub emboss: bool,
    /// Engrave (imprint)
    pub imprint: bool,
    /// Character spacing (in twips)
    pub char_spacing: i32,
    /// Horizontal scaling (percentage)
    pub char_scale: i32,
    /// Kerning (in half-points)
    pub kerning: i32,
    /// Primary language applied to this run.
    pub language: Option<crate::LanguageId>,
    /// East Asian language applied to this run.
    pub east_asian_language: Option<crate::LanguageId>,
    /// Primary language retained while proofing is disabled.
    pub language_no_proof: Option<crate::LanguageId>,
    /// East Asian language retained while proofing is disabled.
    pub east_asian_language_no_proof: Option<crate::LanguageId>,
    /// Whether spelling and grammar proofing is disabled for this run.
    pub no_proof: bool,
    /// Explicit character-run direction; `None` uses left-to-right precedence.
    pub direction: Option<TextDirection>,
    /// Character repertoire selected by `\\loch`, `\\hich`, or `\\dbch`.
    pub character_type: Option<CharacterType>,
    /// Complex-script selector from `\\fcs0` or `\\fcs1`.
    ///
    /// `None` means that no selector was present; `Some(false)` preserves an
    /// explicit `\\fcs0`.
    pub complex_script: Option<bool>,
    /// East Asian character-grid metadata from `\\cgrid` or `\\cgridN`.
    pub character_grid: Option<CharacterGrid>,
    /// Associated complex-script character properties.
    pub associated: AssociatedCharacterFormatting,
}

impl Default for Formatting {
    fn default() -> Self {
        Self {
            character_positioning: crate::CharacterPositioning::default(),
            character_style: None,
            insert_rsid: None,
            delete_rsid: None,
            char_style_rsid: None,
            underline_color: None,
            font_ref: 0,
            font_size: NonZeroU16::new(24).expect("non-zero default font size"),
            color_ref: 0,
            background_color: None,
            highlight_color: None,
            character_border: None,
            character_shading: None,
            bold: false,
            italic: false,
            underline: UnderlineStyle::default(),
            strike: false,
            double_strike: false,
            superscript: false,
            subscript: false,
            smallcaps: false,
            all_caps: false,
            hidden: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            char_spacing: 0,
            char_scale: 100,
            kerning: 0,
            language: None,
            east_asian_language: None,
            language_no_proof: None,
            east_asian_language_no_proof: None,
            no_proof: false,
            direction: None,
            character_type: None,
            complex_script: None,
            character_grid: None,
            associated: AssociatedCharacterFormatting::default(),
        }
    }
}

impl Formatting {
    /// Set or clear the applied character-style handle.
    #[inline]
    pub fn set_character_style(&mut self, value: Option<u16>) {
        self.character_style = value;
    }

    /// Set or clear the exact `\cbN` character background color.
    #[inline]
    pub fn set_background_color(&mut self, value: Option<ColorRef>) {
        self.background_color = value;
    }

    /// Clear the exact character background color without changing highlighting.
    #[inline]
    pub fn clear_background_color(&mut self) {
        self.background_color = None;
    }

    /// Set or clear the character repertoire selector.
    #[inline]
    pub fn set_character_type(&mut self, value: Option<CharacterType>) {
        self.character_type = value;
    }

    /// Set or clear the complex-script selector.
    #[inline]
    pub fn set_complex_script(&mut self, value: Option<bool>) {
        self.complex_script = value;
    }

    /// Set or clear the East Asian character-grid metadata.
    #[inline]
    pub fn set_character_grid(&mut self, value: Option<CharacterGrid>) {
        self.character_grid = value;
    }
}

/// A text run with formatting.
#[derive(Debug, Clone)]
pub struct Run<'a> {
    /// Text content
    pub text: Cow<'a, str>,
    /// Character formatting
    pub formatting: Formatting,
}

impl<'a> Run<'a> {
    /// Create a new run.
    #[inline]
    pub fn new(text: Cow<'a, str>, formatting: Formatting) -> Self {
        Self { text, formatting }
    }

    /// Get the text content.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Check if this run is bold.
    #[inline]
    pub fn bold(&self) -> Option<bool> {
        Some(self.formatting.bold)
    }

    /// Check if this run is italic.
    #[inline]
    pub fn italic(&self) -> Option<bool> {
        Some(self.formatting.italic)
    }

    /// Check if this run has strikethrough.
    #[inline]
    pub fn strikethrough(&self) -> Option<bool> {
        Some(self.formatting.strike || self.formatting.double_strike)
    }

    /// Check if this run has underline.
    #[inline]
    pub fn underline(&self) -> bool {
        !matches!(self.formatting.underline, UnderlineStyle::None)
    }

    /// Get the vertical position of this run (superscript/subscript).
    #[inline]
    pub fn vertical_position(&self) -> Option<litchi_core::style::text::pos::VerticalPosition> {
        if self.formatting.superscript {
            Some(litchi_core::style::text::pos::VerticalPosition::Superscript)
        } else if self.formatting.subscript {
            Some(litchi_core::style::text::pos::VerticalPosition::Subscript)
        } else {
            None
        }
    }
}

/// A styled block of text with paragraph and character formatting.
#[derive(Debug, Clone)]
pub struct StyleBlock<'a> {
    /// Paragraph properties
    pub paragraph: Paragraph,
    /// Character formatting
    pub formatting: Formatting,
    /// Text content
    pub text: Cow<'a, str>,
}

impl<'a> StyleBlock<'a> {
    /// Create a new style block.
    #[inline]
    pub fn new(text: Cow<'a, str>, formatting: Formatting, paragraph: Paragraph) -> Self {
        Self {
            text,
            formatting,
            paragraph,
        }
    }

    /// Get the text content.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A paragraph with content (runs).
///
/// This represents a paragraph in the unified Document API, containing
/// both paragraph properties and the runs that make up the paragraph content.
#[derive(Debug, Clone)]
pub struct ParagraphContent<'a> {
    /// Paragraph properties (alignment, spacing, indentation)
    pub properties: Paragraph,
    /// Runs contained in this paragraph
    pub runs: Vec<Run<'a>>,
}

/// Document element - either a paragraph or a table.
///
/// This enum is used by the `elements()` method to represent
/// the mixed content of an RTF document in sequential order.
#[derive(Debug, Clone)]
pub enum DocumentElement<'a> {
    /// A paragraph with formatted runs
    Paragraph(ParagraphContent<'a>),
    /// A table with rows and cells
    Table(super::table::Table<'a>),
}

impl<'a> ParagraphContent<'a> {
    /// Create a new paragraph with content.
    #[inline]
    pub fn new(properties: Paragraph, runs: Vec<Run<'a>>) -> Self {
        Self { properties, runs }
    }

    /// Get the text content of the paragraph.
    #[inline]
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_ref()).collect()
    }

    /// Get the runs in this paragraph.
    #[inline]
    pub fn runs(&self) -> &[Run<'a>] {
        &self.runs
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_color_new() {
        let color = Color::new(255, 128, 0);
        assert_eq!(color.red, 255);
        assert_eq!(color.green, 128);
        assert_eq!(color.blue, 0);
    }

    #[test]
    fn test_color_black() {
        let color = Color::black();
        assert_eq!(color.red, 0);
        assert_eq!(color.green, 0);
        assert_eq!(color.blue, 0);
    }

    #[test]
    fn test_color_white() {
        let color = Color::white();
        assert_eq!(color.red, 255);
        assert_eq!(color.green, 255);
        assert_eq!(color.blue, 255);
    }

    #[test]
    fn test_color_clone() {
        let color = Color::new(100, 150, 200);
        let cloned = color;
        assert_eq!(cloned.red, color.red);
        assert_eq!(cloned.green, color.green);
        assert_eq!(cloned.blue, color.blue);
    }

    #[test]
    fn test_color_debug() {
        let color = Color::new(255, 0, 0);
        let debug = format!("{:?}", color);
        assert!(debug.contains("Color"));
        assert!(debug.contains("255"));
    }

    #[test]
    fn test_color_table_new() {
        let table = ColorTable::new();
        assert!(table.colors().is_empty());
    }

    #[test]
    fn test_color_table_default() {
        let table: ColorTable = Default::default();
        assert!(table.colors().is_empty());
    }

    #[test]
    fn test_color_table_add() {
        let mut table = ColorTable::new();
        let idx1 = table.add(Color::new(255, 0, 0));
        let idx2 = table.add(Color::new(0, 255, 0));
        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(table.colors().len(), 2);
    }

    #[test]
    fn test_color_table_get() {
        let mut table = ColorTable::new();
        table.add(Color::new(255, 0, 0));
        table.add(Color::new(0, 255, 0));
        assert!(table.get(0).is_some());
        assert!(table.get(1).is_some());
        assert!(table.get(2).is_none());
    }

    #[test]
    fn test_font_family_variants() {
        assert_eq!(FontFamily::default(), FontFamily::Nil);
        assert_ne!(FontFamily::Roman, FontFamily::Swiss);
    }

    #[test]
    fn test_font_new() {
        let font = Font::new(Cow::Borrowed("Arial"), FontFamily::Swiss, 0);
        assert_eq!(font.name, "Arial");
        assert_eq!(font.family, FontFamily::Swiss);
        assert_eq!(font.charset, 0);
    }

    #[test]
    fn test_font_table_new() {
        let table: FontTable = FontTable::new();
        assert!(table.fonts().is_empty());
    }

    #[test]
    fn test_font_table_insert() {
        let mut table: FontTable = FontTable::new();
        table.insert(0, Font::new(Cow::Borrowed("Arial"), FontFamily::Swiss, 0));
        table.insert(1, Font::new(Cow::Borrowed("Times"), FontFamily::Roman, 0));
        assert_eq!(table.fonts().len(), 2);
    }

    #[test]
    fn test_font_table_get() {
        let mut table: FontTable = FontTable::new();
        table.insert(0, Font::new(Cow::Borrowed("Arial"), FontFamily::Swiss, 0));
        assert!(table.get(0).is_some());
        assert!(table.get(1).is_none());
    }

    #[test]
    fn test_font_table_sparse() {
        let mut table: FontTable = FontTable::new();
        table.insert(5, Font::new(Cow::Borrowed("Arial"), FontFamily::Swiss, 0));
        assert_eq!(table.fonts().len(), 6);
    }

    #[test]
    fn test_alignment_variants() {
        assert_eq!(Alignment::default(), Alignment::Left);
        assert_ne!(Alignment::Left, Alignment::Right);
    }

    #[test]
    fn test_spacing_default() {
        let spacing = Spacing::default();
        assert_eq!(spacing.before, 0);
        assert_eq!(spacing.after, 0);
        assert_eq!(spacing.line, 0);
        assert!(!spacing.line_multiple);
    }

    #[test]
    fn test_indentation_default() {
        let indent = Indentation::default();
        assert_eq!(indent.left, 0);
        assert_eq!(indent.right, 0);
        assert_eq!(indent.first_line, 0);
    }

    #[test]
    fn test_paragraph_default() {
        let para = Paragraph::default();
        assert_eq!(para.alignment, Alignment::Left);
        assert!(!para.keep_together);
        assert!(!para.keep_next);
        assert!(!para.page_break_before);
        assert!(!para.widow_control);
    }

    #[test]
    fn test_underline_style_variants() {
        assert_eq!(UnderlineStyle::default(), UnderlineStyle::None);
        assert_ne!(UnderlineStyle::Single, UnderlineStyle::Double);
    }

    #[test]
    fn test_formatting_default() {
        let fmt = Formatting::default();
        assert_eq!(fmt.font_ref, 0);
        assert!(!fmt.bold);
        assert!(!fmt.italic);
        assert!(!fmt.strike);
        assert!(!fmt.superscript);
        assert!(!fmt.subscript);
    }

    #[test]
    fn test_run_new() {
        let fmt = Formatting::default();
        let run = Run::new(Cow::Borrowed("Hello"), fmt);
        assert_eq!(run.text(), "Hello");
        assert!(!run.bold().unwrap());
    }

    #[test]
    fn test_run_bold() {
        let fmt = Formatting {
            bold: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Bold"), fmt);
        assert!(run.bold().unwrap());
    }

    #[test]
    fn test_run_italic() {
        let fmt = Formatting {
            italic: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Italic"), fmt);
        assert!(run.italic().unwrap());
    }

    #[test]
    fn test_run_strikethrough() {
        let fmt = Formatting {
            strike: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Strike"), fmt);
        assert!(run.strikethrough().unwrap());
    }

    #[test]
    fn test_run_double_strikethrough() {
        let fmt = Formatting {
            double_strike: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("DStrike"), fmt);
        assert!(run.strikethrough().unwrap());
    }

    #[test]
    fn test_run_underline() {
        let fmt = Formatting {
            underline: UnderlineStyle::Single,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Underline"), fmt);
        assert!(run.underline());
    }

    #[test]
    fn test_run_no_underline() {
        let fmt = Formatting::default();
        let run = Run::new(Cow::Borrowed("No Underline"), fmt);
        assert!(!run.underline());
    }

    #[test]
    fn test_run_vertical_position_superscript() {
        let fmt = Formatting {
            superscript: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Super"), fmt);
        assert!(matches!(
            run.vertical_position(),
            Some(litchi_core::style::text::pos::VerticalPosition::Superscript)
        ));
    }

    #[test]
    fn test_run_vertical_position_subscript() {
        let fmt = Formatting {
            subscript: true,
            ..Formatting::default()
        };
        let run = Run::new(Cow::Borrowed("Sub"), fmt);
        assert!(matches!(
            run.vertical_position(),
            Some(litchi_core::style::text::pos::VerticalPosition::Subscript)
        ));
    }

    #[test]
    fn test_run_vertical_position_none() {
        let fmt = Formatting::default();
        let run = Run::new(Cow::Borrowed("Normal"), fmt);
        assert!(run.vertical_position().is_none());
    }

    #[test]
    fn test_style_block_new() {
        let fmt = Formatting::default();
        let para = Paragraph::default();
        let block = StyleBlock::new(Cow::Borrowed("Text"), fmt, para);
        assert_eq!(block.text(), "Text");
    }

    #[test]
    fn test_paragraph_content_new() {
        let para = Paragraph::default();
        let content = ParagraphContent::new(para, vec![]);
        assert!(content.runs().is_empty());
    }

    #[test]
    fn test_paragraph_content_text() {
        let fmt = Formatting::default();
        let runs = vec![
            Run::new(Cow::Borrowed("Hello "), fmt),
            Run::new(Cow::Borrowed("World"), fmt),
        ];
        let content = ParagraphContent::new(Paragraph::default(), runs);
        assert_eq!(content.text(), "Hello World");
    }
}
