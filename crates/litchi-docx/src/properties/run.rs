//! Direct run properties: the children of a run's own `w:rPr`.

use quick_xml::events::BytesStart;

use super::Invalid;
use super::scan::{self, Context, Site};
use super::value::{
    Border, Color, Emphasis, Fonts, HalfPoints, Highlight, Language, Shading, Sign, TextEffect,
    ThemeFont, Twips, Unit, VerticalAlign, hex_color, measure, on_off, text_scale,
};
use super::{collect, optional, parse_border, parse_color, parse_shading};
use crate::error::Result;
use crate::format::UnderlineStyle;
use crate::paragraph::{RunUnderline, RunUnderlineColor};

/// The run-property elements this reader types, in schema order. A
/// property's position here is its bit in [`Run::invalid_properties`].
pub(super) const ELEMENTS: [&str; 37] = [
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "specVanish",
    "oMath",
];

const _: () = assert!(ELEMENTS.len() <= 64, "an Invalid mask holds 64 properties");

/// Direct (local) run properties: what one run's own `w:rPr` states.
///
/// Read with [`crate::Run::direct_properties`]. Every field is exactly what
/// the run authors; `None` means the run does not state the property, which
/// is then inherited from the run's character style, the paragraph's style
/// and the document defaults (none of which is resolved here). An on/off
/// property written `<w:b/>` or `<w:b w:val="1"/>` is `Some(true)`, and an
/// explicit `<w:b w:val="0"/>` is `Some(false)`.
///
/// Revision snapshots (`w:rPrChange`), the runs of a text box or a phonetic
/// guide inside the run, and any element outside the WordprocessingML
/// namespace are never read. Word 2010 text effects stay with
/// [`crate::Run::effects`] and OpenType features with
/// [`crate::Run::open_type`].
///
/// ```
/// use litchi_docx::Run;
/// use litchi_docx::properties::{HalfPoints, HexColor, Highlight};
///
/// let run = Run::new(
///     br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:b/><w:i w:val="0"/><w:color w:val="1F4E79"/><w:sz w:val="28"/><w:highlight w:val="yellow"/></w:rPr><w:t>x</w:t></w:r>"#.to_vec(),
/// );
/// let properties = run.direct_properties()?;
/// assert_eq!(properties.bold, Some(true));
/// assert_eq!(properties.italic, Some(false));
/// assert_eq!(properties.strike, None);
/// assert_eq!(properties.size, Some(HalfPoints(28)));
/// assert_eq!(properties.highlight, Some(Highlight::Yellow));
/// assert_eq!(properties.color.map(|color| color.value), Some(HexColor::Rgb([0x1F, 0x4E, 0x79])));
/// assert_eq!(properties.invalid_properties().count(), 0);
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Run {
    /// `w:rStyle`: the character style identifier.
    pub style: Option<String>,
    /// `w:rFonts`
    pub fonts: Option<Fonts>,
    /// `w:b`
    pub bold: Option<bool>,
    /// `w:bCs`: bold for complex-script text.
    pub bold_complex_script: Option<bool>,
    /// `w:i`
    pub italic: Option<bool>,
    /// `w:iCs`: italic for complex-script text.
    pub italic_complex_script: Option<bool>,
    /// `w:caps`: display lowercase letters as capitals.
    pub caps: Option<bool>,
    /// `w:smallCaps`
    pub small_caps: Option<bool>,
    /// `w:strike`: single strikethrough.
    pub strike: Option<bool>,
    /// `w:dstrike`: double strikethrough.
    pub double_strike: Option<bool>,
    /// `w:outline`
    pub outline: Option<bool>,
    /// `w:shadow`
    pub shadow: Option<bool>,
    /// `w:emboss`
    pub emboss: Option<bool>,
    /// `w:imprint`
    pub imprint: Option<bool>,
    /// `w:noProof`: exclude from spelling and grammar checks.
    pub no_proof: Option<bool>,
    /// `w:snapToGrid`
    pub snap_to_grid: Option<bool>,
    /// `w:vanish`: hidden text.
    pub vanish: Option<bool>,
    /// `w:webHidden`
    pub web_hidden: Option<bool>,
    /// `w:color`
    pub color: Option<Color>,
    /// `w:spacing`: added space between characters.
    pub character_spacing: Option<Twips>,
    /// `w:w`: horizontal scale in percent (0 to 600); `<w:w/>` is 100.
    pub scale: Option<u16>,
    /// `w:kern`: the smallest font size that is kerned.
    pub kerning: Option<HalfPoints>,
    /// `w:position`: baseline raise (positive) or lowering (negative).
    pub position: Option<HalfPoints>,
    /// `w:sz`: the font size.
    pub size: Option<HalfPoints>,
    /// `w:szCs`: the font size of complex-script text.
    pub size_complex_script: Option<HalfPoints>,
    /// `w:highlight`
    pub highlight: Option<Highlight>,
    /// `w:u`. `<w:u/>` without `w:val` is a single underline, as
    /// [`crate::Run::underline_formatting`] reads it.
    pub underline: Option<RunUnderline>,
    /// `w:effect`
    pub effect: Option<TextEffect>,
    /// `w:bdr`: a border around the run.
    pub border: Option<Border>,
    /// `w:shd`
    pub shading: Option<Shading>,
    /// `w:vertAlign`: superscript, subscript or the baseline.
    pub vertical_align: Option<VerticalAlign>,
    /// `w:rtl`: right-to-left text.
    pub right_to_left: Option<bool>,
    /// `w:cs`: treat the run as complex-script text.
    pub complex_script: Option<bool>,
    /// `w:em`
    pub emphasis: Option<Emphasis>,
    /// `w:lang`
    pub language: Option<Language>,
    /// `w:specVanish`: always hidden, including in the paragraph mark.
    pub special_vanish: Option<bool>,
    /// `w:oMath`: the run is Office Math text.
    pub office_math: Option<bool>,
    pub(super) invalid: Invalid,
}

impl Run {
    /// The direct properties the run states with a value outside the
    /// property's schema domain (an unknown token, a number that does not
    /// parse or overflows, a missing required attribute), by element local
    /// name in schema order.
    ///
    /// Each such property is left unset rather than guessed. For `w:rFonts`
    /// and `w:lang`, whose attributes are read one by one, only the invalid
    /// attribute is left unset.
    pub fn invalid_properties(&self) -> impl Iterator<Item = &'static str> + '_ {
        self.invalid.names(&ELEMENTS)
    }
}

/// Collects one run's direct properties. A repeated property element is read
/// from its first occurrence, except `w:rFonts` and `w:lang`, whose
/// attributes are taken from the first occurrence that states each, as Word
/// merges them.
#[derive(Default)]
pub(super) struct Reader {
    run: Run,
    seen: u64,
}

impl Reader {
    pub(super) fn finish(self) -> Run {
        self.run
    }

    /// Read one direct child of a `w:rPr`.
    pub(super) fn property(
        &mut self,
        context: &Context<'_>,
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        let Some(index) = ELEMENTS.iter().position(|name| name.as_bytes() == local) else {
            return Ok(());
        };
        let bit = 1u64 << index;
        let repeated = self.seen & bit != 0;
        self.seen |= bit;
        let run = &mut self.run;
        let valid = match local {
            b"rFonts" => {
                let [
                    hint,
                    ascii,
                    h_ansi,
                    east_asia,
                    cs,
                    ascii_theme,
                    h_ansi_theme,
                    east_asia_theme,
                    cs_theme,
                ] = collect(
                    context,
                    element,
                    [
                        b"hint",
                        b"ascii",
                        b"hAnsi",
                        b"eastAsia",
                        b"cs",
                        b"asciiTheme",
                        b"hAnsiTheme",
                        b"eastAsiaTheme",
                        b"cstheme",
                    ],
                )?;
                let fonts = run.fonts.get_or_insert_with(Fonts::default);
                let mut valid = true;
                merge(
                    &mut fonts.hint,
                    hint.as_deref(),
                    super::value::FontHint::from_xml,
                    &mut valid,
                );
                merge_text(&mut fonts.ascii, ascii);
                merge_text(&mut fonts.h_ansi, h_ansi);
                merge_text(&mut fonts.east_asia, east_asia);
                merge_text(&mut fonts.complex_script, cs);
                merge(
                    &mut fonts.ascii_theme,
                    ascii_theme.as_deref(),
                    ThemeFont::from_xml,
                    &mut valid,
                );
                merge(
                    &mut fonts.h_ansi_theme,
                    h_ansi_theme.as_deref(),
                    ThemeFont::from_xml,
                    &mut valid,
                );
                merge(
                    &mut fonts.east_asia_theme,
                    east_asia_theme.as_deref(),
                    ThemeFont::from_xml,
                    &mut valid,
                );
                merge(
                    &mut fonts.complex_script_theme,
                    cs_theme.as_deref(),
                    ThemeFont::from_xml,
                    &mut valid,
                );
                valid
            },
            b"lang" => {
                let [latin, east_asia, bidi] =
                    collect(context, element, [b"val", b"eastAsia", b"bidi"])?;
                let language = run.language.get_or_insert_with(Language::default);
                merge_text(&mut language.latin, latin);
                merge_text(&mut language.east_asia, east_asia);
                merge_text(&mut language.bidi, bidi);
                true
            },
            _ if repeated => true,
            b"rStyle" => {
                let [value] = collect(context, element, [b"val"])?;
                value
                    .map(|value| run.style = Some(value.into_owned()))
                    .is_some()
            },
            b"color" => set(&mut run.color, parse_color(context, element)?),
            b"spacing" => set(
                &mut run.character_spacing,
                value(context, element, |raw| {
                    measure(raw, Unit::Twips, Sign::Signed)
                })?
                .map(Twips),
            ),
            b"w" => {
                let [value] = collect(context, element, [b"val"])?;
                set(
                    &mut run.scale,
                    value.map_or(Some(100), |raw| text_scale(&raw)),
                )
            },
            b"kern" => set(
                &mut run.kerning,
                half_points(context, element, Sign::Unsigned)?,
            ),
            b"position" => set(
                &mut run.position,
                half_points(context, element, Sign::Signed)?,
            ),
            b"sz" => set(
                &mut run.size,
                half_points(context, element, Sign::Unsigned)?,
            ),
            b"szCs" => set(
                &mut run.size_complex_script,
                half_points(context, element, Sign::Unsigned)?,
            ),
            b"highlight" => set(
                &mut run.highlight,
                value(context, element, Highlight::from_xml)?,
            ),
            b"u" => set(&mut run.underline, parse_underline(context, element)?),
            b"effect" => set(
                &mut run.effect,
                value(context, element, TextEffect::from_xml)?,
            ),
            b"bdr" => set(&mut run.border, parse_border(context, element)?),
            b"shd" => set(&mut run.shading, parse_shading(context, element)?),
            b"vertAlign" => set(
                &mut run.vertical_align,
                value(context, element, VerticalAlign::from_xml)?,
            ),
            b"em" => set(
                &mut run.emphasis,
                value(context, element, Emphasis::from_xml)?,
            ),
            _ => {
                let toggle = toggle(context, element)?;
                let Some(slot) = toggle_slot(run, local) else {
                    return Ok(());
                };
                set(slot, toggle)
            },
        };
        if !valid {
            self.run.invalid.insert(index);
        }
        Ok(())
    }
}

/// The field of an on/off run property.
fn toggle_slot<'r>(run: &'r mut Run, local: &[u8]) -> Option<&'r mut Option<bool>> {
    Some(match local {
        b"b" => &mut run.bold,
        b"bCs" => &mut run.bold_complex_script,
        b"i" => &mut run.italic,
        b"iCs" => &mut run.italic_complex_script,
        b"caps" => &mut run.caps,
        b"smallCaps" => &mut run.small_caps,
        b"strike" => &mut run.strike,
        b"dstrike" => &mut run.double_strike,
        b"outline" => &mut run.outline,
        b"shadow" => &mut run.shadow,
        b"emboss" => &mut run.emboss,
        b"imprint" => &mut run.imprint,
        b"noProof" => &mut run.no_proof,
        b"snapToGrid" => &mut run.snap_to_grid,
        b"vanish" => &mut run.vanish,
        b"webHidden" => &mut run.web_hidden,
        b"rtl" => &mut run.right_to_left,
        b"cs" => &mut run.complex_script,
        b"specVanish" => &mut run.special_vanish,
        b"oMath" => &mut run.office_math,
        _ => return None,
    })
}

/// Store a parsed value; `None` (invalid) leaves the field unset.
fn set<T>(slot: &mut Option<T>, value: Option<T>) -> bool {
    match value {
        Some(value) => {
            *slot = Some(value);
            true
        },
        None => false,
    }
}

/// Fill an unset attribute slot from a repeated element, as Word merges
/// `w:rFonts` and `w:lang`.
fn merge<T>(
    slot: &mut Option<T>,
    raw: Option<&str>,
    parse: impl Fn(&str) -> Option<T>,
    valid: &mut bool,
) {
    let Some(raw) = raw else {
        return;
    };
    match parse(raw) {
        Some(value) => {
            if slot.is_none() {
                *slot = Some(value);
            }
        },
        None => *valid = false,
    }
}

fn merge_text(slot: &mut Option<String>, raw: Option<std::borrow::Cow<'_, str>>) {
    if slot.is_none()
        && let Some(raw) = raw
    {
        *slot = Some(raw.into_owned());
    }
}

/// `CT_OnOff`: a missing `w:val` is on. Returns `None` for an invalid value.
pub(super) fn toggle(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<bool>> {
    let [value] = collect(context, element, [b"val"])?;
    Ok(value.map_or(Some(true), |raw| on_off(&raw)))
}

/// A required `w:val`, parsed; `None` when it is missing or invalid.
pub(super) fn value<T>(
    context: &Context<'_>,
    element: &BytesStart<'_>,
    parse: impl Fn(&str) -> Option<T>,
) -> Result<Option<T>> {
    let [value] = collect(context, element, [b"val"])?;
    Ok(value.and_then(|raw| parse(&raw)))
}

fn half_points(
    context: &Context<'_>,
    element: &BytesStart<'_>,
    sign: Sign,
) -> Result<Option<HalfPoints>> {
    Ok(value(context, element, |raw| measure(raw, Unit::HalfPoints, sign))?.map(HalfPoints))
}

/// `CT_Underline`, as [`crate::Run::underline_formatting`] types it.
fn parse_underline(
    context: &Context<'_>,
    element: &BytesStart<'_>,
) -> Result<Option<RunUnderline>> {
    let [style, color, theme, tint, shade] = collect(
        context,
        element,
        [b"val", b"color", b"themeColor", b"themeTint", b"themeShade"],
    )?;
    Ok((|| {
        Some(RunUnderline {
            style: match style.as_deref() {
                None => UnderlineStyle::Single,
                Some(raw) => UnderlineStyle::from_xml(raw)?,
            },
            color: optional(color.as_deref(), hex_color)?.map(|color| match color {
                super::value::HexColor::Auto => RunUnderlineColor::Auto,
                super::value::HexColor::Rgb(rgb) => RunUnderlineColor::Rgb(rgb),
            }),
            theme_color: optional(theme.as_deref(), super::value::theme_color)?,
            theme_tint: optional(tint.as_deref(), super::value::hex_byte)?,
            theme_shade: optional(shade.as_deref(), super::value::hex_byte)?,
        })
    })())
}

/// Read a run fragment's direct properties.
pub(crate) fn parse(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Run> {
    let mut reader = Reader::default();
    scan::walk(
        xml,
        inherited,
        b"r",
        &[b"rPr"],
        "run",
        |context, site, local, element| match site {
            Site::Property(_) => reader.property(context, local, element),
            Site::Nested(..) => Ok(()),
        },
    )?;
    Ok(reader.finish())
}

impl std::fmt::Debug for Run {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Run").finish_non_exhaustive()
    }
}
