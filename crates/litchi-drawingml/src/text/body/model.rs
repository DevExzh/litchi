//! Typed, host-neutral `DrawingML` text-body values.
//!
//! Every property value is *local*: it reports what the element itself
//! declares, so an absent attribute is `None` rather than an inherited or
//! schema-default value. Inheritance (list styles, placeholders, masters and
//! themes) is a separate, resolved view that this model does not compute.
//! Body properties ([`Properties`]) predate that rule and keep the schema
//! defaults their hosts have always applied.

use super::super::{
    Alignment, Anchor, AutoNumber, Autofit, Caps, Columns, Coordinate32, Direction, FontAlignment,
    Strike, TabAlignment, TextPoint, TextSize, Underline, Wrap,
};
use crate::color;

/// Default horizontal body inset in EMUs (0.1 inch).
const DEFAULT_HORIZONTAL_INSET_EMU: i32 = 91_440;
/// Default vertical body inset in EMUs (0.05 inch).
const DEFAULT_VERTICAL_INSET_EMU: i32 = 45_720;

/// Text insets from `a:bodyPr`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Insets {
    /// Left inset.
    pub left: Coordinate32,
    /// Top inset.
    pub top: Coordinate32,
    /// Right inset.
    pub right: Coordinate32,
    /// Bottom inset.
    pub bottom: Coordinate32,
}

impl Default for Insets {
    fn default() -> Self {
        Self {
            left: Coordinate32::from(DEFAULT_HORIZONTAL_INSET_EMU),
            top: Coordinate32::from(DEFAULT_VERTICAL_INSET_EMU),
            right: Coordinate32::from(DEFAULT_HORIZONTAL_INSET_EMU),
            bottom: Coordinate32::from(DEFAULT_VERTICAL_INSET_EMU),
        }
    }
}

/// Text-body properties from `a:bodyPr`, with schema defaults applied.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    /// The source body contains markup not represented by this local projection.
    /// Authoring refuses it rather than silently discarding that markup.
    pub has_unmodeled_content: bool,
    /// Text insets.
    pub insets: Insets,
    /// Vertical anchoring of the text.
    pub vertical_anchor: Anchor,
    /// Whether the anchor point is horizontally centered (`anchorCtr`).
    pub anchor_center: bool,
    /// Text direction.
    pub direction: Direction,
    /// Text wrap behavior.
    pub wrap: Wrap,
    /// Autofit behavior.
    pub autofit: Autofit,
    /// Number of text columns (`numCol`; one when absent).
    pub column_count: Columns,
    /// Whether paragraph spacing is ignored at the first and last paragraphs.
    pub space_first_last_paragraph: bool,
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            has_unmodeled_content: false,
            insets: Insets::default(),
            vertical_anchor: Anchor::default(),
            anchor_center: false,
            direction: Direction::default(),
            wrap: Wrap::default(),
            autofit: Autofit::default(),
            column_count: Columns::ONE,
            space_first_last_paragraph: false,
        }
    }
}

/// A font a run or bullet names (`a:latin`, `a:ea`, `a:cs`, `a:sym`,
/// `a:buFont`; `a:CT_TextFont`).
///
/// The typeface is kept exactly as written, so a theme font reference such as
/// `+mn-lt` stays a reference instead of becoming a resolved family name.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Font {
    /// Typeface name or theme font reference (`typeface`).
    pub typeface: String,
    /// PANOSE classification as written (`panose`).
    pub panose: Option<String>,
    /// Pitch and family (`pitchFamily`).
    pub pitch_family: Option<i8>,
    /// Character set (`charset`).
    pub charset: Option<i8>,
}

impl Font {
    /// The theme font this typeface refers to, when it is a theme reference
    /// (`+mj-lt`, `+mn-ea`, ...): the major (`mj`) or minor (`mn`) font and
    /// its script slot (`lt`, `ea`, `cs`).
    #[must_use]
    pub fn theme_reference(&self) -> Option<(&str, &str)> {
        let reference = self.typeface.strip_prefix('+')?;
        let (font, script) = reference.split_once('-')?;
        (matches!(font, "mj" | "mn") && matches!(script, "lt" | "ea" | "cs"))
            .then_some((font, script))
    }
}

/// The fill painting a run's glyphs (`EG_FillProperties` inside run
/// properties).
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextFill {
    /// No fill (`a:noFill`): the glyphs are not painted.
    None,
    /// One color (`a:solidFill`); `None` when the element names no color.
    Solid(Option<color::Value>),
    /// A gradient (`a:gradFill`); its stops are not modeled.
    Gradient,
    /// A pattern (`a:pattFill`); not modeled further.
    Pattern,
    /// A picture (`a:blipFill`); not modeled further.
    Picture,
    /// The containing group's fill (`a:grpFill`).
    Group,
}

impl TextFill {
    /// The color of a solid fill, when it names one.
    #[must_use]
    pub const fn solid_color(&self) -> Option<&color::Value> {
        match self {
            Self::Solid(Some(color)) => Some(color),
            _ => None,
        }
    }
}

/// A click action on a run (`a:hlinkClick`; `a:CT_Hyperlink`).
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Hyperlink {
    /// Relationship naming the target part or URL (`r:id`). It is empty or
    /// absent for an action that needs no target, such as a jump to the next
    /// slide.
    pub(crate) relationship_id: Option<String>,
    /// Action URI such as `ppaction://hlinksldjump` (`action`).
    pub action: Option<String>,
    /// Tooltip text (`tooltip`).
    pub tooltip: Option<String>,
}

/// Local character properties (`a:rPr`, `a:defRPr`, `a:endParaRPr`;
/// `a:CT_TextCharacterProperties`).
#[derive(Clone, Default, PartialEq, Eq)]
pub struct RunProperties {
    /// Bold toggle (`b`).
    pub bold: Option<bool>,
    /// Italic toggle (`i`).
    pub italic: Option<bool>,
    /// Underline style (`u`).
    pub underline: Option<Underline>,
    /// Strikethrough (`strike`).
    pub strike: Option<Strike>,
    /// Font size in hundredths of a point (`sz`).
    pub size: Option<TextSize>,
    /// Baseline offset in thousandths of a percent of the font size
    /// (`baseline`); positive values raise the text (superscript).
    pub baseline: Option<i32>,
    /// Character spacing in hundredths of a point (`spc`).
    pub spacing: Option<TextPoint>,
    /// Smallest font size, in hundredths of a point, that is kerned
    /// (`kern`).
    pub kerning: Option<TextPoint>,
    /// Capitalization (`cap`).
    pub caps: Option<Caps>,
    /// Language tag (`lang`).
    pub language: Option<String>,
    /// Alternative language tag (`altLang`).
    pub alternate_language: Option<String>,
    /// Latin font (`a:latin`).
    pub latin: Option<Font>,
    /// East Asian font (`a:ea`).
    pub east_asian: Option<Font>,
    /// Complex-script font (`a:cs`).
    pub complex_script: Option<Font>,
    /// Symbol font (`a:sym`).
    pub symbol: Option<Font>,
    /// Glyph fill: the run's color when it is a solid fill.
    pub fill: Option<TextFill>,
    /// Highlight color behind the glyphs (`a:highlight`).
    pub highlight: Option<color::Value>,
    /// Click action (`a:hlinkClick`).
    pub hyperlink: Option<Hyperlink>,
}

impl RunProperties {
    /// The color of the run's solid fill, when it names one.
    #[must_use]
    pub fn color(&self) -> Option<&color::Value> {
        self.fill.as_ref().and_then(TextFill::solid_color)
    }
}

/// A text field (`a:fld`), such as a slide number or a date.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Field {
    /// Field identifier, a GUID as written (`id`).
    pub(crate) id: String,
    /// Field type, for example `slidenum` or `datetime1` (`type`).
    pub kind: Option<String>,
}

/// The markup a [`Run`] was read from.
#[derive(Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum RunKind {
    /// A text run (`a:r`).
    #[default]
    Text,
    /// A line break (`a:br`); its text is `"\n"`.
    Break,
    /// A field (`a:fld`); its text is the cached field result.
    Field(Field),
}

/// One `DrawingML` run inside a text-body paragraph.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Run {
    /// The markup the run was read from.
    pub kind: RunKind,
    /// Run text: a line break is `"\n"` and a field is its cached result.
    pub text: String,
    /// Local character properties.
    pub properties: RunProperties,
}

impl Run {
    /// A text run with no local properties.
    #[must_use]
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }

    /// Whether the run is a text run (`a:r`).
    #[must_use]
    pub const fn is_text(&self) -> bool {
        matches!(self.kind, RunKind::Text)
    }
}

/// Paragraph spacing (`a:lnSpc`, `a:spcBef`, `a:spcAft`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Spacing {
    /// A proportion of the line height, in thousandths of a percent
    /// (`a:spcPct`); `100000` is single spacing.
    Percent(u32),
    /// An exact amount, in hundredths of a point (`a:spcPts`).
    Points(u32),
}

/// Bullet color (`a:buClrTx`, `a:buClr`).
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum BulletColor {
    /// The bullet takes the color of the first run (`a:buClrTx`).
    FollowText,
    /// An explicit color (`a:buClr`).
    Color(color::Value),
}

/// Bullet size (`a:buSzTx`, `a:buSzPct`, `a:buSzPts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BulletSize {
    /// The bullet takes the size of the first run (`a:buSzTx`).
    FollowText,
    /// A proportion of the text size in thousandths of a percent
    /// (`a:buSzPct`, 25000..=400000).
    Percent(u32),
    /// An exact size in hundredths of a point (`a:buSzPts`).
    Points(TextSize),
}

/// Bullet font (`a:buFontTx`, `a:buFont`).
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum BulletFont {
    /// The bullet takes the font of the first run (`a:buFontTx`).
    FollowText,
    /// An explicit font (`a:buFont`).
    Font(Font),
}

/// The bullet of a paragraph (`a:buNone`, `a:buAutoNum`, `a:buChar`,
/// `a:buBlip`).
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Bullet {
    /// Explicitly no bullet (`a:buNone`).
    None,
    /// An automatically numbered bullet (`a:buAutoNum`).
    AutoNumber {
        /// Numbering scheme (`type`).
        scheme: AutoNumber,
        /// First number, 1..=32767 (`startAt`; one when absent).
        start_at: Option<u16>,
    },
    /// A character bullet (`a:buChar`).
    Character(String),
    /// A picture bullet (`a:buBlip`).
    Picture(Picture),
}

/// A custom tab stop (`a:tab`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TabStop {
    /// Position from the left margin (`pos`).
    pub position: Option<Coordinate32>,
    /// Alignment at the stop (`algn`).
    pub alignment: Option<TabAlignment>,
}

/// Local paragraph properties (`a:pPr`; `a:CT_TextParagraphProperties`).
#[derive(Clone, Default, PartialEq, Eq)]
pub struct ParagraphProperties {
    /// Outline level, 0..=8 (`lvl`).
    pub level: Option<u8>,
    /// Horizontal alignment (`algn`).
    pub alignment: Option<Alignment>,
    /// Left margin in EMUs, 0..=51206400 (`marL`).
    pub margin_left: Option<i32>,
    /// Right margin in EMUs, 0..=51206400 (`marR`).
    pub margin_right: Option<i32>,
    /// First-line indent in EMUs, -51206400..=51206400 (`indent`).
    pub indent: Option<i32>,
    /// Default tab size (`defTabSz`).
    pub default_tab_size: Option<Coordinate32>,
    /// Right-to-left reading order (`rtl`).
    pub right_to_left: Option<bool>,
    /// East Asian line breaking (`eaLnBrk`).
    pub east_asian_line_break: Option<bool>,
    /// Latin line breaking inside words (`latinLnBrk`).
    pub latin_line_break: Option<bool>,
    /// Hanging punctuation (`hangingPunct`).
    pub hanging_punctuation: Option<bool>,
    /// Vertical font alignment (`fontAlgn`).
    pub font_alignment: Option<FontAlignment>,
    /// Line spacing (`a:lnSpc`).
    pub line_spacing: Option<Spacing>,
    /// Space before the paragraph (`a:spcBef`).
    pub space_before: Option<Spacing>,
    /// Space after the paragraph (`a:spcAft`).
    pub space_after: Option<Spacing>,
    /// Bullet color (`a:buClrTx`, `a:buClr`).
    pub bullet_color: Option<BulletColor>,
    /// Bullet size (`a:buSzTx`, `a:buSzPct`, `a:buSzPts`).
    pub bullet_size: Option<BulletSize>,
    /// Bullet font (`a:buFontTx`, `a:buFont`).
    pub bullet_font: Option<BulletFont>,
    /// Bullet (`a:buNone`, `a:buAutoNum`, `a:buChar`, `a:buBlip`).
    pub bullet: Option<Bullet>,
    /// Custom tab stops (`a:tabLst`); an empty list is an explicit
    /// `a:tabLst` with no stops.
    pub tab_stops: Option<Vec<TabStop>>,
    /// Character properties runs of the paragraph default to (`a:defRPr`).
    pub default_run: Option<Box<RunProperties>>,
}

/// One paragraph inside a `DrawingML` text body.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Paragraph {
    /// Local paragraph properties (`a:pPr`).
    pub properties: ParagraphProperties,
    /// Runs, line breaks and fields in document order.
    pub runs: Vec<Run>,
    /// Character properties of the paragraph mark (`a:endParaRPr`).
    pub end_properties: Option<Box<RunProperties>>,
}

impl Paragraph {
    /// Concatenate the paragraph's run text without evaluating fields.
    #[must_use]
    pub fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }
}

/// One host-neutral `DrawingML` text story.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Body {
    /// Body properties with `DrawingML` defaults applied.
    pub properties: Properties,
    /// Paragraphs in document order.
    pub paragraphs: Vec<Paragraph>,
}

impl Body {
    /// Return all paragraph text separated by newline characters.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(Paragraph::text)
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[cfg(test)]
mod tests {
    use super::{Body, Font, Insets, Paragraph, Properties, Run, RunKind};

    #[test]
    fn applies_drawingml_defaults_and_joins_text() {
        let body = Body {
            properties: Properties::default(),
            paragraphs: vec![
                Paragraph {
                    runs: vec![Run::plain("first")],
                    ..Paragraph::default()
                },
                Paragraph {
                    runs: vec![Run::plain("second")],
                    ..Paragraph::default()
                },
            ],
        };
        assert_eq!(body.text(), "first\nsecond");
        assert_eq!(body.properties.insets, Insets::default());
        assert!(body.paragraphs[0].runs[0].is_text());
        assert_eq!(body.paragraphs[0].runs[0].kind, RunKind::Text);
    }

    #[test]
    fn theme_font_references_are_recognized_without_resolution() {
        let font = |typeface: &str| Font {
            typeface: typeface.to_owned(),
            ..Font::default()
        };
        assert_eq!(font("+mn-lt").theme_reference(), Some(("mn", "lt")));
        assert_eq!(font("+mj-ea").theme_reference(), Some(("mj", "ea")));
        assert_eq!(font("Calibri").theme_reference(), None);
        assert_eq!(font("+mx-lt").theme_reference(), None);
        assert_eq!(font("+mn-zz").theme_reference(), None);
    }
}

/// A picture bullet; native relationship metadata is available through `body::raw`.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Picture {
    pub(crate) relationship_id: Option<String>,
}

macro_rules! redacted_debug {
    ($($ty:ty),* $(,)?) => {$(
        impl std::fmt::Debug for $ty {
            fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                f.debug_struct(stringify!($ty)).finish_non_exhaustive()
            }
        }
    )*};
}
redacted_debug!(
    Font,
    TextFill,
    Hyperlink,
    RunProperties,
    Field,
    RunKind,
    Run,
    BulletColor,
    BulletFont,
    Bullet,
    ParagraphProperties,
    Paragraph,
    Body,
    Picture
);
