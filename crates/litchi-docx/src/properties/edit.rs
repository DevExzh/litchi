//! Splicing typed direct properties into existing run and paragraph markup.
//!
//! The readers next to this module turn a run's `w:rPr` and a paragraph's
//! `w:pPr` into [`Run`] and [`Paragraph`]. This module goes the other way for
//! the main-document transaction: given the typed value a caller wants, it
//! rewrites only the property elements whose typed value differs from what
//! the fragment states, and leaves every other byte of the fragment alone.
//!
//! - An element whose typed value is unchanged is not touched, so its
//!   lexical form (`w:val="true"`, `w:sz w:val="12pt"`) survives.
//! - A changed element keeps its prefix, its children, every attribute the
//!   typed model does not own (other namespaces, unknown names), and every
//!   typed attribute whose value already reads as the target, byte for byte.
//!   Its other typed attributes are rewritten or removed so the element reads
//!   back exactly as the target; an invalid value in a changed element is
//!   replaced.
//! - A new element is written in canonical form and inserted in schema
//!   order: after the last known sibling that the schema puts before it,
//!   otherwise before the first known sibling it puts after it. The
//!   WordprocessingML properties stay ahead of trailing extension elements,
//!   change-tracking snapshots (`w:rPrChange`, `w:pPrChange`) stay last, and
//!   the paragraph mark's `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` stay first.
//! - Removing a property removes only its element. Removing a property
//!   container (`w:numPr`, `w:pBdr`, `w:tabs`, the paragraph mark's `w:rPr`)
//!   is refused when it holds change-tracking markup or markup the typed
//!   model does not own, because that markup would go with it.
//! - A property element written more than once is refused when it has to
//!   change, because which occurrence a consumer honours is not proven.
//!
//! Names are resolved exactly as the readers resolve them (see `scan`), so
//! the elements edited here are the elements the readers read, and every
//! splice is checked by reading the result back with those readers.

use std::borrow::Cow;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

use super::paragraph::{
    Indentation, Justification, Numbering, Paragraph, ParagraphBorders, TabKind, TabStop,
    TextDirection,
};
use super::run::Run;
use super::scan::{Resolution, Root, resolve, split};
use super::table::{Height, Row, Table, TableBorders};
use super::value::{
    Border, Color, Fonts, HexColor, Language, Shading, Sign, Unit, hex_byte, hex_color, integer,
    measure, on_off, text_scale,
};
use crate::document::Refusal;
use crate::error::Error;
use crate::paragraph::{ParagraphSpacing, RunUnderline, RunUnderlineColor};

/// Deepest element nesting accepted in one fragment, as the readers accept.
const MAX_DEPTH: usize = 128;

/// `EG_RPrBase` in schema order, then the run's change-tracking snapshot.
const RUN_ORDER: &[&str] = &[
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
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
];

/// `CT_ParaRPr`: the paragraph mark's tracked insertion, deletion and move
/// markers, then `EG_RPrBase`, then the change-tracking snapshot.
const MARK_ORDER: &[&str] = &[
    "ins",
    "del",
    "moveFrom",
    "moveTo",
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
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
];

/// `CT_PPrBase` in schema order, then the rest of `CT_PPr`.
const PARAGRAPH_ORDER: &[&str] = &[
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
];

/// `CT_NumPr` in schema order.
const NUMBERING_ORDER: &[&str] = &["ilvl", "numId", "numberingChange", "ins"];
/// The children of `w:numPr` the typed model owns.
const NUMBERING_TYPED: &[&str] = &["ilvl", "numId"];
/// `CT_PBdr` in schema order; every side is typed.
const BORDERS_ORDER: &[&str] = &["top", "left", "bottom", "right", "between", "bar"];
/// The only child of `w:tabs`.
const TABS_TYPED: &[&str] = &["tab"];
/// `CT_TblPr` in schema order.
const TABLE_ORDER: &[&str] = &[
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
];
/// `CT_TblBorders` in schema order, with the transitional `left`/`right`
/// beside the `start`/`end` they spell; every side is typed.
const TABLE_BORDERS_ORDER: &[&str] = &[
    "top", "start", "left", "bottom", "end", "right", "insideH", "insideV",
];
/// `CT_TrPr` children in the order Word writes them (the schema admits any
/// order), then the tracked-change markup.
const ROW_ORDER: &[&str] = &[
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
    "ins",
    "del",
    "trPrChange",
];
/// What may precede `w:trPr` in a row (`CT_Row`).
const ROW_LEADING: &[&str] = &["tblPrEx"];
/// Elements that record a tracked change and must never be dropped.
const TRACKED: &[&str] = &[
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "rPrChange",
    "pPrChange",
    "numberingChange",
];

/// Why a property splice failed.
#[derive(Debug)]
pub(crate) enum SpliceError {
    /// The requested change cannot be written without guessing.
    Refused(Refusal),
    /// The fragment is not well-formed WordprocessingML, or the splice did
    /// not read back as requested.
    Document(Error),
    Transaction(crate::document::Error),
}

impl From<Error> for SpliceError {
    fn from(error: Error) -> Self {
        Self::Document(error)
    }
}

impl From<Refusal> for SpliceError {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

// ---------------------------------------------------------------------------
// Plans: what one property element must look like after the splice.
// ---------------------------------------------------------------------------

/// How a raw attribute value is compared with a planned value.
#[derive(Clone, Copy)]
enum Kind {
    /// Text and closed vocabularies: the decoded value must be equal.
    Exact,
    /// `ST_OnOff`.
    OnOff,
    /// `ST_HexColor`.
    HexColor,
    /// `ST_UcharHexNumber`.
    HexByte,
    /// A measure stored in `Unit`.
    Measure(Unit, Sign),
    /// `xsd:integer`.
    Integer,
    /// `ST_TextScale`.
    Scale,
}

impl Kind {
    /// The canonical form of a decoded attribute value, or `None` when the
    /// value is outside the domain the readers accept.
    fn canonical(self, value: &str) -> Option<Cow<'_, str>> {
        match self {
            Self::Exact => Some(Cow::Borrowed(value)),
            Self::OnOff => on_off(value).map(|on| Cow::Borrowed(if on { "1" } else { "0" })),
            Self::HexColor => hex_color(value).map(|color| Cow::Owned(hex(color))),
            Self::HexByte => hex_byte(value).map(|byte| Cow::Owned(hex_pair(byte))),
            Self::Measure(unit, sign) => {
                measure(value, unit, sign).map(|number| Cow::Owned(number.to_string()))
            },
            Self::Integer => integer(value).map(|number| Cow::Owned(number.to_string())),
            Self::Scale => text_scale(value).map(|scale| Cow::Owned(scale.to_string())),
        }
    }
}

/// What one typed attribute must be after the splice.
enum Want {
    /// The attribute must not be written.
    Absent,
    /// The attribute must read as this canonical value.
    Value(String),
    /// The attribute may be absent or read as this canonical value; a new
    /// element omits it. `CT_OnOff`'s `w:val` is on when absent.
    Implied(&'static str),
}

/// One typed attribute of a planned element.
struct Attribute {
    local: &'static str,
    kind: Kind,
    want: Want,
}

/// The children a planned element must have.
enum Children {
    /// A leaf property: whatever children the element has are kept.
    Keep,
    /// A property container whose own property children change.
    Container {
        order: &'static [&'static str],
        changes: Vec<Change>,
    },
    /// The tab stops of `w:tabs`, replaced as one list.
    Tabs(Vec<Plan>),
}

/// The typed state one property element must have after the splice.
struct Plan {
    attributes: SmallVec<[Attribute; 4]>,
    children: Children,
}

/// One property element to write (`Some`) or remove (`None`).
struct Change {
    local: &'static str,
    plan: Option<Plan>,
    /// For a property container (`w:numPr`, `w:pBdr`, `w:tabs`, the
    /// paragraph mark's `w:rPr`), the children the typed model owns; a
    /// container holding anything else is not removed.
    holds: Option<&'static [&'static str]>,
}

fn leaf<const N: usize>(attributes: [Attribute; N]) -> Plan {
    Plan {
        attributes: attributes.into_iter().collect(),
        children: Children::Keep,
    }
}

fn value(local: &'static str, kind: Kind, value: String) -> Attribute {
    Attribute {
        local,
        kind,
        want: Want::Value(value),
    }
}

fn optional(local: &'static str, kind: Kind, value: Option<String>) -> Attribute {
    Attribute {
        local,
        kind,
        want: value.map_or(Want::Absent, Want::Value),
    }
}

/// Record a change when the typed value differs.
fn field<T: PartialEq>(
    changes: &mut Vec<Change>,
    local: &'static str,
    current: &Option<T>,
    target: &Option<T>,
    plan: impl FnOnce(&T) -> Result<Plan, Refusal>,
) -> Result<(), Refusal> {
    if current != target {
        changes.push(Change {
            local,
            plan: target.as_ref().map(plan).transpose()?,
            holds: None,
        });
    }
    Ok(())
}

/// Record a change of an on/off property.
fn toggle(
    changes: &mut Vec<Change>,
    local: &'static str,
    current: Option<bool>,
    target: Option<bool>,
) {
    if current != target {
        changes.push(Change {
            local,
            plan: target.map(|on| {
                leaf([Attribute {
                    local: "val",
                    kind: Kind::OnOff,
                    want: if on {
                        Want::Implied("1")
                    } else {
                        Want::Value("0".to_owned())
                    },
                }])
            }),
            holds: None,
        });
    }
}

/// Text written into an attribute: every character must be legal in XML.
fn text(value: &str) -> Result<String, Refusal> {
    if value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        Ok(value.to_owned())
    } else {
        Err(Refusal::InvalidPropertyValue)
    }
}

fn hex(color: HexColor) -> String {
    match color {
        HexColor::Auto => "auto".to_owned(),
        HexColor::Rgb([red, green, blue]) => format!("{red:02X}{green:02X}{blue:02X}"),
    }
}

fn hex_pair(byte: u8) -> String {
    format!("{byte:02X}")
}

/// A measure the readers accept: an unsigned measure must not be negative.
fn measured(number: i64, sign: Sign) -> Result<String, Refusal> {
    if sign == Sign::Unsigned && number < 0 {
        return Err(Refusal::InvalidPropertyValue);
    }
    Ok(number.to_string())
}

fn single_value(kind: Kind, token: String) -> Plan {
    leaf([value("val", kind, token)])
}

fn color_plan(color: &Color) -> Plan {
    leaf([
        value("val", Kind::HexColor, hex(color.value)),
        optional(
            "themeColor",
            Kind::Exact,
            color.theme.map(|theme| theme.as_str().to_owned()),
        ),
        optional("themeTint", Kind::HexByte, color.theme_tint.map(hex_pair)),
        optional("themeShade", Kind::HexByte, color.theme_shade.map(hex_pair)),
    ])
}

fn optional_text(local: &'static str, value: Option<&str>) -> Result<Attribute, Refusal> {
    Ok(optional(local, Kind::Exact, value.map(text).transpose()?))
}

fn fonts_plan(fonts: &Fonts) -> Result<Plan, Refusal> {
    let theme = |slot: Option<super::value::ThemeFont>| slot.map(|font| font.as_str().to_owned());
    Ok(leaf([
        optional(
            "hint",
            Kind::Exact,
            fonts.hint.map(|hint| hint.as_str().to_owned()),
        ),
        optional_text("ascii", fonts.ascii.as_deref())?,
        optional_text("hAnsi", fonts.h_ansi.as_deref())?,
        optional_text("eastAsia", fonts.east_asia.as_deref())?,
        optional_text("cs", fonts.complex_script.as_deref())?,
        optional("asciiTheme", Kind::Exact, theme(fonts.ascii_theme)),
        optional("hAnsiTheme", Kind::Exact, theme(fonts.h_ansi_theme)),
        optional("eastAsiaTheme", Kind::Exact, theme(fonts.east_asia_theme)),
        optional("cstheme", Kind::Exact, theme(fonts.complex_script_theme)),
    ]))
}

fn underline_plan(underline: &RunUnderline) -> Plan {
    leaf([
        value("val", Kind::Exact, underline.style.as_str().to_owned()),
        optional(
            "color",
            Kind::HexColor,
            underline.color.map(|color| match color {
                RunUnderlineColor::Auto => hex(HexColor::Auto),
                RunUnderlineColor::Rgb(rgb) => hex(HexColor::Rgb(rgb)),
            }),
        ),
        optional(
            "themeColor",
            Kind::Exact,
            underline.theme_color.map(|theme| theme.as_str().to_owned()),
        ),
        optional(
            "themeTint",
            Kind::HexByte,
            underline.theme_tint.map(hex_pair),
        ),
        optional(
            "themeShade",
            Kind::HexByte,
            underline.theme_shade.map(hex_pair),
        ),
    ])
}

fn border_plan(border: &Border) -> Result<Plan, Refusal> {
    Ok(leaf([
        value("val", Kind::Exact, border.style.as_str().to_owned()),
        optional("color", Kind::HexColor, border.color.map(hex)),
        optional(
            "themeColor",
            Kind::Exact,
            border.theme_color.map(|theme| theme.as_str().to_owned()),
        ),
        optional("themeTint", Kind::HexByte, border.theme_tint.map(hex_pair)),
        optional(
            "themeShade",
            Kind::HexByte,
            border.theme_shade.map(hex_pair),
        ),
        optional(
            "sz",
            Kind::Measure(Unit::EighthPoints, Sign::Unsigned),
            border
                .size
                .map(|size| measured(size.0, Sign::Unsigned))
                .transpose()?,
        ),
        optional(
            "space",
            Kind::Measure(Unit::Points, Sign::Unsigned),
            border
                .space
                .map(|space| measured(space.0, Sign::Unsigned))
                .transpose()?,
        ),
        optional("shadow", Kind::OnOff, border.shadow.map(on_off_text)),
        optional("frame", Kind::OnOff, border.frame.map(on_off_text)),
    ]))
}

fn on_off_text(on: bool) -> String {
    if on { "1" } else { "0" }.to_owned()
}

fn shading_plan(shading: &Shading) -> Plan {
    let theme = |theme: Option<crate::color::Theme>| theme.map(|theme| theme.as_str().to_owned());
    leaf([
        value("val", Kind::Exact, shading.pattern.as_str().to_owned()),
        optional("color", Kind::HexColor, shading.color.map(hex)),
        optional("themeColor", Kind::Exact, theme(shading.theme_color)),
        optional("themeTint", Kind::HexByte, shading.theme_tint.map(hex_pair)),
        optional(
            "themeShade",
            Kind::HexByte,
            shading.theme_shade.map(hex_pair),
        ),
        optional("fill", Kind::HexColor, shading.fill.map(hex)),
        optional("themeFill", Kind::Exact, theme(shading.theme_fill)),
        optional(
            "themeFillTint",
            Kind::HexByte,
            shading.theme_fill_tint.map(hex_pair),
        ),
        optional(
            "themeFillShade",
            Kind::HexByte,
            shading.theme_fill_shade.map(hex_pair),
        ),
    ])
}

fn language_plan(language: &Language) -> Result<Plan, Refusal> {
    Ok(leaf([
        optional_text("val", language.latin.as_deref())?,
        optional_text("eastAsia", language.east_asia.as_deref())?,
        optional_text("bidi", language.bidi.as_deref())?,
    ]))
}

fn half_points(number: i64, sign: Sign) -> Result<Plan, Refusal> {
    Ok(single_value(
        Kind::Measure(Unit::HalfPoints, sign),
        measured(number, sign)?,
    ))
}

/// The property changes that turn `current` into `target`, in schema order.
fn run_changes(current: &Run, target: &Run) -> Result<Vec<Change>, Refusal> {
    let mut changes = Vec::new();
    field(
        &mut changes,
        "rStyle",
        &current.style,
        &target.style,
        |style| Ok(single_value(Kind::Exact, text(style)?)),
    )?;
    field(
        &mut changes,
        "rFonts",
        &current.fonts,
        &target.fonts,
        fonts_plan,
    )?;
    toggle(&mut changes, "b", current.bold, target.bold);
    toggle(
        &mut changes,
        "bCs",
        current.bold_complex_script,
        target.bold_complex_script,
    );
    toggle(&mut changes, "i", current.italic, target.italic);
    toggle(
        &mut changes,
        "iCs",
        current.italic_complex_script,
        target.italic_complex_script,
    );
    toggle(&mut changes, "caps", current.caps, target.caps);
    toggle(
        &mut changes,
        "smallCaps",
        current.small_caps,
        target.small_caps,
    );
    toggle(&mut changes, "strike", current.strike, target.strike);
    toggle(
        &mut changes,
        "dstrike",
        current.double_strike,
        target.double_strike,
    );
    toggle(&mut changes, "outline", current.outline, target.outline);
    toggle(&mut changes, "shadow", current.shadow, target.shadow);
    toggle(&mut changes, "emboss", current.emboss, target.emboss);
    toggle(&mut changes, "imprint", current.imprint, target.imprint);
    toggle(&mut changes, "noProof", current.no_proof, target.no_proof);
    toggle(
        &mut changes,
        "snapToGrid",
        current.snap_to_grid,
        target.snap_to_grid,
    );
    toggle(&mut changes, "vanish", current.vanish, target.vanish);
    toggle(
        &mut changes,
        "webHidden",
        current.web_hidden,
        target.web_hidden,
    );
    field(
        &mut changes,
        "color",
        &current.color,
        &target.color,
        |color| Ok(color_plan(color)),
    )?;
    field(
        &mut changes,
        "spacing",
        &current.character_spacing,
        &target.character_spacing,
        |spacing| {
            Ok(single_value(
                Kind::Measure(Unit::Twips, Sign::Signed),
                spacing.0.to_string(),
            ))
        },
    )?;
    field(&mut changes, "w", &current.scale, &target.scale, |scale| {
        if *scale > 600 {
            return Err(Refusal::InvalidPropertyValue);
        }
        Ok(single_value(Kind::Scale, scale.to_string()))
    })?;
    field(
        &mut changes,
        "kern",
        &current.kerning,
        &target.kerning,
        |kerning| half_points(kerning.0, Sign::Unsigned),
    )?;
    field(
        &mut changes,
        "position",
        &current.position,
        &target.position,
        |position| half_points(position.0, Sign::Signed),
    )?;
    field(&mut changes, "sz", &current.size, &target.size, |size| {
        half_points(size.0, Sign::Unsigned)
    })?;
    field(
        &mut changes,
        "szCs",
        &current.size_complex_script,
        &target.size_complex_script,
        |size| half_points(size.0, Sign::Unsigned),
    )?;
    field(
        &mut changes,
        "highlight",
        &current.highlight,
        &target.highlight,
        |highlight| Ok(single_value(Kind::Exact, highlight.as_str().to_owned())),
    )?;
    field(
        &mut changes,
        "u",
        &current.underline,
        &target.underline,
        |underline| Ok(underline_plan(underline)),
    )?;
    field(
        &mut changes,
        "effect",
        &current.effect,
        &target.effect,
        |effect| Ok(single_value(Kind::Exact, effect.as_str().to_owned())),
    )?;
    field(
        &mut changes,
        "bdr",
        &current.border,
        &target.border,
        border_plan,
    )?;
    field(
        &mut changes,
        "shd",
        &current.shading,
        &target.shading,
        |shading| Ok(shading_plan(shading)),
    )?;
    field(
        &mut changes,
        "vertAlign",
        &current.vertical_align,
        &target.vertical_align,
        |align| Ok(single_value(Kind::Exact, align.as_str().to_owned())),
    )?;
    toggle(
        &mut changes,
        "rtl",
        current.right_to_left,
        target.right_to_left,
    );
    toggle(
        &mut changes,
        "cs",
        current.complex_script,
        target.complex_script,
    );
    field(
        &mut changes,
        "em",
        &current.emphasis,
        &target.emphasis,
        |emphasis| Ok(single_value(Kind::Exact, emphasis.as_str().to_owned())),
    )?;
    field(
        &mut changes,
        "lang",
        &current.language,
        &target.language,
        language_plan,
    )?;
    toggle(
        &mut changes,
        "specVanish",
        current.special_vanish,
        target.special_vanish,
    );
    toggle(
        &mut changes,
        "oMath",
        current.office_math,
        target.office_math,
    );
    Ok(changes)
}

fn numbering_plan(current: Option<&Numbering>, target: &Numbering) -> Plan {
    let empty = Numbering::default();
    let current = current.unwrap_or(&empty);
    let mut changes = Vec::new();
    if current.level != target.level {
        changes.push(Change {
            local: "ilvl",
            plan: target
                .level
                .map(|level| single_value(Kind::Integer, level.to_string())),
            holds: None,
        });
    }
    if current.id != target.id {
        changes.push(Change {
            local: "numId",
            plan: target
                .id
                .map(|id| single_value(Kind::Integer, id.to_string())),
            holds: None,
        });
    }
    Plan {
        attributes: SmallVec::new(),
        children: Children::Container {
            order: NUMBERING_ORDER,
            changes,
        },
    }
}

fn borders_plan(
    current: Option<&ParagraphBorders>,
    target: &ParagraphBorders,
) -> Result<Plan, Refusal> {
    let empty = ParagraphBorders::default();
    let current = current.unwrap_or(&empty);
    let mut changes = Vec::new();
    for (local, current, target) in [
        ("top", &current.top, &target.top),
        ("left", &current.left, &target.left),
        ("bottom", &current.bottom, &target.bottom),
        ("right", &current.right, &target.right),
        ("between", &current.between, &target.between),
        ("bar", &current.bar, &target.bar),
    ] {
        field(&mut changes, local, current, target, border_plan)?;
    }
    Ok(Plan {
        attributes: SmallVec::new(),
        children: Children::Container {
            order: BORDERS_ORDER,
            changes,
        },
    })
}

fn table_borders_plan(
    current: Option<&TableBorders>,
    target: &TableBorders,
) -> Result<Plan, Refusal> {
    let empty = TableBorders::default();
    let current = current.unwrap_or(&empty);
    let mut changes = Vec::new();
    for (local, current, target) in [
        ("top", &current.top, &target.top),
        ("start", &current.start, &target.start),
        ("left", &current.left, &target.left),
        ("bottom", &current.bottom, &target.bottom),
        ("end", &current.end, &target.end),
        ("right", &current.right, &target.right),
        (
            "insideH",
            &current.inside_horizontal,
            &target.inside_horizontal,
        ),
        ("insideV", &current.inside_vertical, &target.inside_vertical),
    ] {
        field(&mut changes, local, current, target, border_plan)?;
    }
    Ok(Plan {
        attributes: SmallVec::new(),
        children: Children::Container {
            order: TABLE_BORDERS_ORDER,
            changes,
        },
    })
}

/// The `w:tblPr` changes from `current` to `target`. The style and the
/// borders are written (record 0342); a change of any other property is
/// refused.
fn table_changes(current: &Table, target: &Table) -> Result<Vec<Change>, Refusal> {
    let mut rest = target.clone();
    rest.style.clone_from(&current.style);
    rest.borders = current.borders;
    if !same_table(&rest, current) {
        return Err(Refusal::UnmodelledProperty);
    }
    let mut changes = Vec::new();
    field(
        &mut changes,
        "tblStyle",
        &current.style,
        &target.style,
        |style| Ok(single_value(Kind::Exact, text(style)?)),
    )?;
    if current.borders != target.borders {
        changes.push(Change {
            local: "tblBorders",
            plan: target
                .borders
                .as_ref()
                .map(|borders| table_borders_plan(current.borders.as_ref(), borders))
                .transpose()?,
            holds: Some(TABLE_BORDERS_ORDER),
        });
    }
    Ok(changes)
}

fn height_plan(height: &Height) -> Result<Plan, Refusal> {
    Ok(leaf([
        value(
            "val",
            Kind::Measure(Unit::Twips, Sign::Unsigned),
            measured(height.value.0, Sign::Unsigned)?,
        ),
        optional(
            "hRule",
            Kind::Exact,
            height.rule.map(|rule| rule.as_str().to_owned()),
        ),
    ]))
}

/// The `w:trPr` changes from `current` to `target`. The flags `cantSplit`,
/// `tblHeader` and `hidden` and the height are written (record 0342); a
/// change of any other property is refused.
fn row_changes(current: &Row, target: &Row) -> Result<Vec<Change>, Refusal> {
    let mut rest = target.clone();
    rest.cant_split = current.cant_split;
    rest.height = current.height;
    rest.header = current.header;
    rest.hidden = current.hidden;
    if !same_row(&rest, current) {
        return Err(Refusal::UnmodelledProperty);
    }
    let mut changes = Vec::new();
    toggle(
        &mut changes,
        "cantSplit",
        current.cant_split,
        target.cant_split,
    );
    field(
        &mut changes,
        "trHeight",
        &current.height,
        &target.height,
        height_plan,
    )?;
    toggle(&mut changes, "tblHeader", current.header, target.header);
    toggle(&mut changes, "hidden", current.hidden, target.hidden);
    Ok(changes)
}

fn tabs_plan(tabs: &[TabStop], strict: bool) -> Result<Plan, Refusal> {
    let mut plans = Vec::new();
    for tab in tabs {
        if strict && matches!(tab.kind, TabKind::Left | TabKind::Right) {
            return Err(Refusal::TransitionalValue);
        }
        plans.push(leaf([
            value("val", Kind::Exact, tab.kind.as_str().to_owned()),
            optional(
                "leader",
                Kind::Exact,
                tab.leader.map(|leader| leader.as_str().to_owned()),
            ),
            value(
                "pos",
                Kind::Measure(Unit::Twips, Sign::Signed),
                tab.position.0.to_string(),
            ),
        ]));
    }
    Ok(Plan {
        attributes: SmallVec::new(),
        children: Children::Tabs(plans),
    })
}

fn spacing_plan(spacing: &ParagraphSpacing) -> Result<Plan, Refusal> {
    let twips = |value: Option<u64>| {
        value
            .map(|value| {
                i64::try_from(value)
                    .map(|value| value.to_string())
                    .map_err(|_| Refusal::InvalidPropertyValue)
            })
            .transpose()
    };
    let unsigned = Kind::Measure(Unit::Twips, Sign::Unsigned);
    Ok(leaf([
        optional("before", unsigned, twips(spacing.before)?),
        optional(
            "beforeLines",
            Kind::Integer,
            spacing.before_lines.map(|lines| lines.to_string()),
        ),
        optional(
            "beforeAutospacing",
            Kind::OnOff,
            spacing.before_auto_spacing.map(on_off_text),
        ),
        optional("after", unsigned, twips(spacing.after)?),
        optional(
            "afterLines",
            Kind::Integer,
            spacing.after_lines.map(|lines| lines.to_string()),
        ),
        optional(
            "afterAutospacing",
            Kind::OnOff,
            spacing.after_auto_spacing.map(on_off_text),
        ),
        optional(
            "line",
            Kind::Measure(Unit::Twips, Sign::Signed),
            spacing.line.map(|line| line.to_string()),
        ),
        optional(
            "lineRule",
            Kind::Exact,
            spacing.line_rule.map(|rule| rule.as_str().to_owned()),
        ),
    ]))
}

fn indentation_plan(indentation: &Indentation, strict: bool) -> Result<Plan, Refusal> {
    if strict
        && (indentation.left.is_some()
            || indentation.left_chars.is_some()
            || indentation.right.is_some()
            || indentation.right_chars.is_some())
    {
        return Err(Refusal::TransitionalValue);
    }
    let signed = Kind::Measure(Unit::Twips, Sign::Signed);
    let unsigned = Kind::Measure(Unit::Twips, Sign::Unsigned);
    let twips = |value: Option<super::value::Twips>, sign: Sign| {
        value.map(|value| measured(value.0, sign)).transpose()
    };
    let chars = |value: Option<i64>| value.map(|value| value.to_string());
    Ok(leaf([
        optional("start", signed, twips(indentation.start, Sign::Signed)?),
        optional("startChars", Kind::Integer, chars(indentation.start_chars)),
        optional("end", signed, twips(indentation.end, Sign::Signed)?),
        optional("endChars", Kind::Integer, chars(indentation.end_chars)),
        optional("left", signed, twips(indentation.left, Sign::Signed)?),
        optional("leftChars", Kind::Integer, chars(indentation.left_chars)),
        optional("right", signed, twips(indentation.right, Sign::Signed)?),
        optional("rightChars", Kind::Integer, chars(indentation.right_chars)),
        optional(
            "hanging",
            unsigned,
            twips(indentation.hanging, Sign::Unsigned)?,
        ),
        optional(
            "hangingChars",
            Kind::Integer,
            chars(indentation.hanging_chars),
        ),
        optional(
            "firstLine",
            unsigned,
            twips(indentation.first_line, Sign::Unsigned)?,
        ),
        optional(
            "firstLineChars",
            Kind::Integer,
            chars(indentation.first_line_chars),
        ),
    ]))
}

fn mark_plan(current: Option<&Run>, target: &Run) -> Result<Plan, Refusal> {
    let empty = Run::default();
    Ok(Plan {
        attributes: SmallVec::new(),
        children: Children::Container {
            order: MARK_ORDER,
            changes: run_changes(current.unwrap_or(&empty), target)?,
        },
    })
}

/// The property changes that turn `current` into `target`, in schema order.
fn paragraph_changes(
    current: &Paragraph,
    target: &Paragraph,
    strict: bool,
) -> Result<Vec<Change>, Refusal> {
    let mut changes = Vec::new();
    field(
        &mut changes,
        "pStyle",
        &current.style,
        &target.style,
        |style| Ok(single_value(Kind::Exact, text(style)?)),
    )?;
    toggle(
        &mut changes,
        "keepNext",
        current.keep_next,
        target.keep_next,
    );
    toggle(
        &mut changes,
        "keepLines",
        current.keep_lines,
        target.keep_lines,
    );
    toggle(
        &mut changes,
        "pageBreakBefore",
        current.page_break_before,
        target.page_break_before,
    );
    toggle(
        &mut changes,
        "widowControl",
        current.widow_control,
        target.widow_control,
    );
    if current.numbering != target.numbering {
        changes.push(Change {
            local: "numPr",
            plan: target
                .numbering
                .as_ref()
                .map(|numbering| numbering_plan(current.numbering.as_ref(), numbering)),
            holds: Some(NUMBERING_TYPED),
        });
    }
    toggle(
        &mut changes,
        "suppressLineNumbers",
        current.suppress_line_numbers,
        target.suppress_line_numbers,
    );
    if current.borders != target.borders {
        changes.push(Change {
            local: "pBdr",
            plan: target
                .borders
                .as_ref()
                .map(|borders| borders_plan(current.borders.as_ref(), borders))
                .transpose()?,
            holds: Some(BORDERS_ORDER),
        });
    }
    field(
        &mut changes,
        "shd",
        &current.shading,
        &target.shading,
        |shading| Ok(shading_plan(shading)),
    )?;
    if current.tabs != target.tabs {
        changes.push(Change {
            local: "tabs",
            plan: target
                .tabs
                .as_deref()
                .map(|tabs| tabs_plan(tabs, strict))
                .transpose()?,
            holds: Some(TABS_TYPED),
        });
    }
    toggle(
        &mut changes,
        "suppressAutoHyphens",
        current.suppress_auto_hyphens,
        target.suppress_auto_hyphens,
    );
    toggle(&mut changes, "kinsoku", current.kinsoku, target.kinsoku);
    toggle(
        &mut changes,
        "wordWrap",
        current.word_wrap,
        target.word_wrap,
    );
    toggle(
        &mut changes,
        "overflowPunct",
        current.overflow_punctuation,
        target.overflow_punctuation,
    );
    toggle(
        &mut changes,
        "topLinePunct",
        current.top_line_punctuation,
        target.top_line_punctuation,
    );
    toggle(
        &mut changes,
        "autoSpaceDE",
        current.auto_space_east_asian_latin,
        target.auto_space_east_asian_latin,
    );
    toggle(
        &mut changes,
        "autoSpaceDN",
        current.auto_space_east_asian_numbers,
        target.auto_space_east_asian_numbers,
    );
    toggle(&mut changes, "bidi", current.bidi, target.bidi);
    toggle(
        &mut changes,
        "adjustRightInd",
        current.adjust_right_indent,
        target.adjust_right_indent,
    );
    toggle(
        &mut changes,
        "snapToGrid",
        current.snap_to_grid,
        target.snap_to_grid,
    );
    field(
        &mut changes,
        "spacing",
        &current.spacing,
        &target.spacing,
        spacing_plan,
    )?;
    field(
        &mut changes,
        "ind",
        &current.indentation,
        &target.indentation,
        |indentation| indentation_plan(indentation, strict),
    )?;
    toggle(
        &mut changes,
        "contextualSpacing",
        current.contextual_spacing,
        target.contextual_spacing,
    );
    toggle(
        &mut changes,
        "mirrorIndents",
        current.mirror_indents,
        target.mirror_indents,
    );
    toggle(
        &mut changes,
        "suppressOverlap",
        current.suppress_overlap,
        target.suppress_overlap,
    );
    field(
        &mut changes,
        "jc",
        &current.justification,
        &target.justification,
        |justification| {
            if strict && matches!(justification, Justification::Left | Justification::Right) {
                return Err(Refusal::TransitionalValue);
            }
            Ok(single_value(Kind::Exact, justification.as_str().to_owned()))
        },
    )?;
    field(
        &mut changes,
        "textDirection",
        &current.text_direction,
        &target.text_direction,
        |direction| {
            if strict
                && matches!(
                    direction,
                    TextDirection::BtLr
                        | TextDirection::LrTb
                        | TextDirection::LrTbV
                        | TextDirection::TbLrV
                        | TextDirection::TbRl
                        | TextDirection::TbRlV
                )
            {
                return Err(Refusal::TransitionalValue);
            }
            Ok(single_value(Kind::Exact, direction.as_str().to_owned()))
        },
    )?;
    field(
        &mut changes,
        "textAlignment",
        &current.text_alignment,
        &target.text_alignment,
        |alignment| Ok(single_value(Kind::Exact, alignment.as_str().to_owned())),
    )?;
    field(
        &mut changes,
        "textboxTightWrap",
        &current.textbox_tight_wrap,
        &target.textbox_tight_wrap,
        |wrap| Ok(single_value(Kind::Exact, wrap.as_str().to_owned())),
    )?;
    field(
        &mut changes,
        "outlineLvl",
        &current.outline_level,
        &target.outline_level,
        |level| {
            if *level > 9 {
                return Err(Refusal::InvalidPropertyValue);
            }
            Ok(single_value(Kind::Integer, level.to_string()))
        },
    )?;
    if current.mark != target.mark {
        changes.push(Change {
            local: "rPr",
            plan: target
                .mark
                .as_deref()
                .map(|mark| mark_plan(current.mark.as_deref(), mark))
                .transpose()?,
            holds: Some(&super::run::ELEMENTS),
        });
    }
    Ok(changes)
}

// ---------------------------------------------------------------------------
// The fragment outline.
// ---------------------------------------------------------------------------

/// One element of a scanned fragment.
struct Node {
    /// Offset of the element's `<`.
    start: usize,
    /// Offset just past the start tag.
    tag_end: usize,
    /// Offset just past the element.
    end: usize,
    empty: bool,
    word: bool,
    local: SmallVec<[u8; 16]>,
    prefix: SmallVec<[u8; 8]>,
    /// Children, recorded only for the root, the property container and the
    /// nested property containers.
    children: Vec<Node>,
    opaque_content: bool,
}

impl Node {
    fn is(&self, local: &str) -> bool {
        self.word && self.local.as_slice() == local.as_bytes()
    }
}

/// The outline of one run or paragraph fragment.
struct Fragment {
    root: Root,
    tree: Node,
}

/// Outline `xml`, whose root must be a WordprocessingML `root_local`.
///
/// The root's children are recorded; so are the children of the root's
/// WordprocessingML `container` child, and the children of that container's
/// WordprocessingML children named in `nested`.
fn scan(
    xml: &[u8],
    root_local: &[u8],
    container: &[u8],
    nested: &[&[u8]],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Fragment, SpliceError> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited)?;
    let origin = ReaderOrigin::of(xml);
    let mut written = None::<Root>;
    let mut stack: Vec<(Node, bool)> = Vec::new();
    let mut finished = None::<Node>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::InvalidFormat("Word fragment offset overflow".into()))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Xml("invalid Word property XML".into()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::InvalidFormat("Word fragment offset overflow".into()))?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                let (mut node, _) = stack
                    .pop()
                    .ok_or_else(|| Error::InvalidFormat("invalid Word fragment nesting".into()))?;
                node.end = end;
                match stack.last_mut() {
                    Some((parent, true)) => parent.children.push(node),
                    Some(_) => {},
                    None => finished = Some(node),
                }
                continue;
            },
            Event::Eof => break,
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "Word fragment holds a declaration or processing instruction".into(),
                )
                .into());
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => continue,
            Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                for (node, _) in &mut stack {
                    node.opaque_content = true;
                }
                continue;
            },
            Event::Decl(_) => continue,
        };
        super::scan::audit_attributes(&element, reader.resolver(), reader.decoder())?;
        let level = stack.len() + 1;
        if level > MAX_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "Word fragment nesting exceeds the {MAX_DEPTH} depth limit"
            ))
            .into());
        }
        let (local, prefix) = split(element.name());
        let resolution = resolve(reader.resolver(), prefix, true);
        if level == 1 {
            if written.is_some() {
                return Err(
                    Error::InvalidFormat("Word fragment has more than one root".into()).into(),
                );
            }
            let root = root_of(prefix, resolution);
            if local != root_local || !root.is_word(prefix, resolution) {
                return Err(
                    Error::InvalidFormat("Word fragment has an invalid root".into()).into(),
                );
            }
            written = Some(root);
        }
        let word = written
            .as_ref()
            .is_some_and(|root| root.is_word(prefix, resolution));
        let parent_records = stack.last().is_some_and(|(_, record)| *record);
        let record = match level {
            1 => true,
            2 => word && local == container,
            3 => parent_records && word && nested.contains(&local),
            _ => false,
        };
        let node = Node {
            start,
            tag_end: end,
            end,
            empty,
            word,
            local: SmallVec::from_slice(local),
            prefix: SmallVec::from_slice(prefix.unwrap_or_default()),
            children: Vec::new(),
            opaque_content: false,
        };
        if empty {
            match stack.last_mut() {
                Some((parent, true)) => parent.children.push(node),
                Some(_) => {},
                None => finished = Some(node),
            }
        } else {
            stack.push((node, record));
        }
    }
    match (written, finished) {
        (Some(root), Some(tree)) if stack.is_empty() => Ok(Fragment { root, tree }),
        _ => Err(Error::InvalidFormat("unterminated Word fragment".into()).into()),
    }
}

/// How a fragment's root is written, as the readers classify it.
fn root_of(prefix: Option<&[u8]>, resolution: Resolution<'_>) -> Root {
    match resolution {
        Resolution::Bound(_) => Root::Bound,
        Resolution::Undeclared => Root::Prefix(SmallVec::from_slice(prefix.unwrap_or_default())),
        Resolution::Unbound => Root::Unprefixed,
    }
}

/// The root's single WordprocessingML `local` child, if any.
fn child<'n>(parent: &'n Node, local: &str) -> Result<Option<&'n Node>, SpliceError> {
    let mut found = parent.children.iter().filter(|node| node.is(local));
    let first = found.next();
    if found.next().is_some() {
        return Err(Refusal::DuplicateProperty.into());
    }
    Ok(first)
}

// ---------------------------------------------------------------------------
// Applying plans.
// ---------------------------------------------------------------------------

/// What every edit of one fragment shares.
struct Context<'a> {
    xml: &'a [u8],
    inherited: &'a [(Option<Vec<u8>>, Vec<u8>)],
    root: &'a Root,
    /// The prefix written on new WordprocessingML attributes.
    attribute_prefix: &'a [u8],
}

/// One byte-range replacement in the fragment.
struct Edit {
    start: usize,
    end: usize,
    bytes: Vec<u8>,
}

fn rank(order: &[&str], local: &[u8]) -> Option<usize> {
    order.iter().position(|name| name.as_bytes() == local)
}

/// Where a new element of `new_rank` goes among the container's children:
/// immediately after the last known WordprocessingML sibling the schema puts
/// before it, otherwise immediately before the first known one it puts after
/// it, otherwise first.
///
/// Preferring the lower sibling keeps the WordprocessingML properties ahead
/// of trailing extension elements (Word 2010 `w14:` effects follow them), and
/// change-tracking snapshots, which rank last, stay last.
fn anchor(container: &Node, order: &[&str], new_rank: usize) -> usize {
    let mut after_lower = None;
    let mut before_higher = None;
    for node in container.children.iter().filter(|node| node.word) {
        let Some(node_rank) = rank(order, &node.local) else {
            continue;
        };
        if node_rank < new_rank {
            after_lower = Some(node.end);
        } else if before_higher.is_none() {
            before_higher = Some(node.start);
        }
    }
    after_lower.or(before_higher).unwrap_or(container.tag_end)
}

fn qualified(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) -> Result<(), SpliceError> {
    reserve_output(
        output,
        bounded_add(bounded_add(prefix.len(), local.len())?, 8)?,
    )?;
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
    Ok(())
}

/// Escape an attribute value so that it reads back, after attribute-value
/// normalization, as exactly `value`.
fn escape_attribute(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#9;"),
            '\n' => output.extend_from_slice(b"&#10;"),
            '\r' => output.extend_from_slice(b"&#13;"),
            _ => {
                let mut buffer = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
}

fn write_attribute(
    output: &mut Vec<u8>,
    prefix: &[u8],
    local: &str,
    value: &str,
) -> Result<(), SpliceError> {
    reserve_output(
        output,
        bounded_add(
            bounded_add(prefix.len(), local.len())?,
            bounded_add(value.len().checked_mul(6).unwrap_or(usize::MAX), 8)?,
        )?,
    )?;
    output.push(b' ');
    qualified(output, prefix, local.as_bytes())?;
    output.extend_from_slice(b"=\"");
    escape_attribute(output, value);
    output.push(b'"');
    Ok(())
}

/// Whether the plan writes any attribute, which needs a bound prefix.
fn writes_attributes(plan: &Plan) -> bool {
    plan.attributes
        .iter()
        .any(|attribute| matches!(attribute.want, Want::Value(_)))
        || match &plan.children {
            Children::Keep => false,
            Children::Container { changes, .. } => changes
                .iter()
                .any(|change| change.plan.as_ref().is_some_and(writes_attributes)),
            Children::Tabs(tabs) => tabs.iter().any(writes_attributes),
        }
}

/// Write a new element in canonical form.
fn serialize(
    output: &mut Vec<u8>,
    prefix: &[u8],
    attribute_prefix: &[u8],
    local: &str,
    plan: &Plan,
) -> Result<(), SpliceError> {
    let bytes = serialized_bound(prefix, attribute_prefix, local, plan)?;
    reserve_output(output, bytes)?;
    output.push(b'<');
    qualified(output, prefix, local.as_bytes())?;
    for attribute in &plan.attributes {
        if let Want::Value(value) = &attribute.want {
            write_attribute(output, attribute_prefix, attribute.local, value)?;
        }
    }
    let mut inner = Vec::new();
    serialize_children(&mut inner, prefix, attribute_prefix, &plan.children)?;
    if inner.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        output.extend_from_slice(&inner);
        output.extend_from_slice(b"</");
        qualified(output, prefix, local.as_bytes())?;
        output.push(b'>');
    }
    Ok(())
}

fn serialize_children(
    output: &mut Vec<u8>,
    prefix: &[u8],
    attribute_prefix: &[u8],
    children: &Children,
) -> Result<(), SpliceError> {
    match children {
        Children::Keep => {},
        Children::Container { changes, .. } => {
            for change in changes {
                if let Some(plan) = &change.plan {
                    serialize(output, prefix, attribute_prefix, change.local, plan)?;
                }
            }
        },
        Children::Tabs(tabs) => {
            for tab in tabs {
                serialize(output, prefix, attribute_prefix, "tab", tab)?;
            }
        },
    }
    Ok(())
}

/// The qualified name of an element, as its start tag writes it.
fn tag_name(tag: &[u8]) -> Result<&[u8], SpliceError> {
    let name = tag
        .get(1..)
        .and_then(|rest| {
            rest.iter()
                .position(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
                .map(|length| &rest[..length])
        })
        .filter(|name| !name.is_empty())
        .ok_or_else(|| Error::InvalidFormat("Word fragment has an unnamed tag".into()))?;
    Ok(name)
}

/// Turn an empty-element tag (`<x .../>`) into the start tag `<x ...>`.
fn opened(tag: &[u8]) -> Result<Vec<u8>, SpliceError> {
    let slash = tag
        .len()
        .checked_sub(2)
        .filter(|slash| tag.get(*slash..) == Some(b"/>".as_slice()))
        .ok_or_else(|| Error::InvalidFormat("Word fragment has a malformed empty tag".into()))?;
    let mut output = tag[..slash].to_vec();
    output.push(b'>');
    Ok(output)
}

fn closing(name: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(name.len() + 3);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    output
}

/// Apply `changes` to the children of `container`.
fn edit_container(
    context: &Context<'_>,
    container: &Node,
    order: &'static [&'static str],
    changes: &[Change],
    edits: &mut Vec<Edit>,
) -> Result<(), SpliceError> {
    if container.empty {
        let mut inner = Vec::new();
        for change in changes {
            if let Some(plan) = &change.plan {
                serialize(
                    &mut inner,
                    &container.prefix,
                    context.attribute_prefix,
                    change.local,
                    plan,
                )?;
            }
        }
        if !inner.is_empty() {
            let tag = &context.xml[container.start..container.tag_end];
            let mut bytes = opened(tag)?;
            bytes.extend_from_slice(&inner);
            bytes.extend_from_slice(&closing(tag_name(tag)?));
            edits.push(Edit {
                start: container.start,
                end: container.end,
                bytes,
            });
        }
        return Ok(());
    }
    for change in changes {
        let existing = child(container, change.local)?;
        match (existing, &change.plan) {
            (None, None) => {},
            (None, Some(plan)) => {
                let new_rank = rank(order, change.local.as_bytes()).ok_or_else(|| {
                    Error::InvalidFormat("planned Word property has no schema rank".into())
                })?;
                let at = anchor(container, order, new_rank);
                let mut bytes = Vec::new();
                serialize(
                    &mut bytes,
                    &container.prefix,
                    context.attribute_prefix,
                    change.local,
                    plan,
                )?;
                edits.push(Edit {
                    start: at,
                    end: at,
                    bytes,
                });
            },
            (Some(node), None) => {
                removable(node, change.holds)?;
                edits.push(Edit {
                    start: node.start,
                    end: node.end,
                    bytes: Vec::new(),
                });
            },
            (Some(node), Some(plan)) => rewrite(context, node, plan, edits)?,
        }
    }
    Ok(())
}

/// Refuse to remove a property container that holds change-tracking markup
/// or markup the typed model does not own.
fn removable(node: &Node, typed: Option<&[&str]>) -> Result<(), SpliceError> {
    let Some(typed) = typed else {
        return Ok(());
    };
    if node.opaque_content {
        return Err(Refusal::ComplexContent.into());
    }
    for child in &node.children {
        let local = std::str::from_utf8(&child.local).unwrap_or_default();
        if child.word && TRACKED.contains(&local) {
            return Err(Refusal::TrackedChange.into());
        }
        if !child.word || !typed.contains(&local) {
            return Err(Refusal::ComplexContent.into());
        }
    }
    Ok(())
}

/// Rewrite an existing property element to `plan`.
fn rewrite(
    context: &Context<'_>,
    node: &Node,
    plan: &Plan,
    edits: &mut Vec<Edit>,
) -> Result<(), SpliceError> {
    let tag = rewrite_tag(context, node, &plan.attributes)?;
    let original = &context.xml[node.start..node.tag_end];
    match &plan.children {
        Children::Keep => {
            if let Some(tag) = tag {
                edits.push(Edit {
                    start: node.start,
                    end: node.tag_end,
                    bytes: tag,
                });
            }
        },
        Children::Container { order, changes } => {
            if node.empty {
                let mut inner = Vec::new();
                for change in changes {
                    if let Some(plan) = &change.plan {
                        serialize(
                            &mut inner,
                            &node.prefix,
                            context.attribute_prefix,
                            change.local,
                            plan,
                        )?;
                    }
                }
                if inner.is_empty() {
                    if let Some(tag) = tag {
                        edits.push(Edit {
                            start: node.start,
                            end: node.tag_end,
                            bytes: tag,
                        });
                    }
                } else {
                    let mut bytes = opened(tag.as_deref().unwrap_or(original))?;
                    bytes.extend_from_slice(&inner);
                    bytes.extend_from_slice(&closing(tag_name(original)?));
                    edits.push(Edit {
                        start: node.start,
                        end: node.end,
                        bytes,
                    });
                }
            } else {
                if let Some(tag) = tag {
                    edits.push(Edit {
                        start: node.start,
                        end: node.tag_end,
                        bytes: tag,
                    });
                }
                edit_container(context, node, order, changes, edits)?;
            }
        },
        Children::Tabs(tabs) => {
            let mut bytes = Vec::new();
            for tab in tabs {
                serialize(
                    &mut bytes,
                    &node.prefix,
                    context.attribute_prefix,
                    "tab",
                    tab,
                )?;
            }
            if node.empty {
                let tag = tag.as_deref().unwrap_or(original);
                if bytes.is_empty() {
                    if tag != original {
                        edits.push(Edit {
                            start: node.start,
                            end: node.tag_end,
                            bytes: tag.to_vec(),
                        });
                    }
                } else {
                    let mut replacement = opened(tag)?;
                    replacement.extend_from_slice(&bytes);
                    replacement.extend_from_slice(&closing(tag_name(original)?));
                    edits.push(Edit {
                        start: node.start,
                        end: node.end,
                        bytes: replacement,
                    });
                }
            } else {
                if let Some(tag) = tag {
                    edits.push(Edit {
                        start: node.start,
                        end: node.tag_end,
                        bytes: tag,
                    });
                }
                let existing: SmallVec<[&Node; 8]> =
                    node.children.iter().filter(|child| child.word).collect();
                if existing.iter().any(|child| !child.is("tab")) {
                    return Err(Refusal::ComplexContent.into());
                }
                let at = existing.first().map_or(node.tag_end, |first| first.start);
                edits.push(Edit {
                    start: at,
                    end: at,
                    bytes,
                });
                for child in existing {
                    edits.push(Edit {
                        start: child.start,
                        end: child.end,
                        bytes: Vec::new(),
                    });
                }
            }
        },
    }
    Ok(())
}

/// One attribute of a start tag, as written.
struct RawAttribute {
    /// Offset of the whitespace before the attribute name.
    start: usize,
    /// Offset just past the attribute name.
    name_end: usize,
    /// Offset just past the closing quote.
    end: usize,
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// The attributes of a well-formed start tag, with the offset where the tag's
/// tail (`>` or `/>`, after optional whitespace) starts.
fn raw_attributes(tag: &[u8]) -> Option<(SmallVec<[RawAttribute; 8]>, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !is_xml_space(tag[cursor]) && !matches!(tag[cursor], b'/' | b'>') {
        cursor += 1;
    }
    let mut attributes = SmallVec::new();
    loop {
        let start = cursor;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        match tag.get(cursor)? {
            b'/' | b'>' => return Some((attributes, start)),
            _ => {},
        }
        while cursor < tag.len()
            && !is_xml_space(tag[cursor])
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if !matches!(quote, b'"' | b'\'') {
            return None;
        }
        cursor += 1;
        cursor += tag.get(cursor..)?.iter().position(|byte| *byte == quote)?;
        cursor += 1;
        attributes.push(RawAttribute {
            start,
            name_end,
            end: cursor,
        });
    }
}

/// The attributes of the element starting at `at`: the WordprocessingML
/// local name of each (`None` for another namespace) and its decoded,
/// normalized value, in document order.
fn element_attributes(
    context: &Context<'_>,
    at: usize,
) -> Result<SmallVec<[(Option<SmallVec<[u8; 16]>>, String); 8]>, SpliceError> {
    let mut reader = NsReader::from_reader(context.xml);
    seed_namespaces(&mut reader, context.inherited)?;
    let origin = ReaderOrigin::of(context.xml);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::InvalidFormat("Word fragment offset overflow".into()))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Xml("invalid Word property XML".into()))?;
        let element = match event {
            Event::Start(element) | Event::Empty(element) if start == at => element,
            Event::Eof => {
                return Err(
                    Error::InvalidFormat("Word property element disappeared".into()).into(),
                );
            },
            _ => continue,
        };
        let mut attributes = SmallVec::new();
        for attribute in element.checked_attributes() {
            let attribute =
                attribute.map_err(|_| Error::Xml("invalid Word property XML".into()))?;
            let (local, prefix) = split(attribute.key);
            let word = context
                .root
                .is_word(prefix, resolve(reader.resolver(), prefix, false));
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|_| Error::Xml("invalid Word property XML".into()))?;
            attributes.push((
                word.then(|| SmallVec::from_slice(local)),
                decoded.into_owned(),
            ));
        }
        return Ok(attributes);
    }
}

/// The start tag of `node` rewritten so its typed attributes are `planned`,
/// or `None` when they already are.
fn rewrite_tag(
    context: &Context<'_>,
    node: &Node,
    planned: &[Attribute],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if planned.is_empty() {
        return Ok(None);
    }
    let tag = &context.xml[node.start..node.tag_end];
    let (raw, tail) = raw_attributes(tag)
        .ok_or_else(|| Error::InvalidFormat("Word property tag is malformed".into()))?;
    let decoded = element_attributes(context, node.start)?;
    if decoded.len() != raw.len() {
        return Err(Error::InvalidFormat("Word property tag attributes diverged".into()).into());
    }
    let mut handled: SmallVec<[bool; 12]> = SmallVec::from_elem(false, planned.len());
    let name_end = raw.first().map_or(tail, |attribute| attribute.start);
    let mut output = tag[..name_end].to_vec();
    let mut changed = false;
    for (attribute, (word, value)) in raw.iter().zip(&decoded) {
        let index = word.as_ref().and_then(|local| {
            planned
                .iter()
                .position(|plan| plan.local.as_bytes() == local.as_slice())
        });
        let Some(index) = index else {
            output.extend_from_slice(&tag[attribute.start..attribute.end]);
            continue;
        };
        if handled[index] {
            return Err(Refusal::DuplicateProperty.into());
        }
        handled[index] = true;
        let plan = &planned[index];
        let keep = match &plan.want {
            Want::Absent => false,
            Want::Value(wanted) => plan.kind.canonical(value).as_deref() == Some(wanted.as_str()),
            Want::Implied(wanted) => plan.kind.canonical(value).as_deref() == Some(*wanted),
        };
        if keep {
            output.extend_from_slice(&tag[attribute.start..attribute.end]);
            continue;
        }
        changed = true;
        if let Want::Value(wanted) = &plan.want {
            output.extend_from_slice(&tag[attribute.start..attribute.name_end]);
            output.extend_from_slice(b"=\"");
            escape_attribute(&mut output, wanted);
            output.push(b'"');
        }
    }
    for (plan, handled) in planned.iter().zip(&handled) {
        if *handled {
            continue;
        }
        if let Want::Value(wanted) = &plan.want {
            changed = true;
            write_attribute(&mut output, context.attribute_prefix, plan.local, wanted)?;
        }
    }
    output.extend_from_slice(&tag[tail..]);
    Ok(changed.then_some(output))
}

/// Apply non-overlapping edits to `xml`; insertions at one offset keep their
/// order and precede a replacement that starts there.
fn apply(xml: &[u8], mut edits: Vec<Edit>) -> Result<Vec<u8>, SpliceError> {
    edits.sort_by_key(|edit| (edit.start, edit.end));
    let capacity = edits.iter().try_fold(xml.len(), |total, edit| {
        bounded_add(total, edit.bytes.len())
    })?;
    let mut output = Vec::new();
    reserve_output(&mut output, capacity)?;
    let mut cursor = 0usize;
    for edit in edits {
        if edit.start < cursor || edit.end < edit.start || edit.end > xml.len() {
            return Err(Error::InvalidFormat("Word property edits overlap".into()).into());
        }
        output.extend_from_slice(&xml[cursor..edit.start]);
        output.extend_from_slice(&edit.bytes);
        cursor = edit.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

/// The prefix new attributes are written with: the root's, which the readers
/// and the part both resolve to WordprocessingML.
fn attribute_prefix(fragment: &Fragment, needed: bool) -> Result<&[u8], SpliceError> {
    if needed && fragment.tree.prefix.is_empty() {
        return Err(Refusal::NamespaceBinding.into());
    }
    Ok(&fragment.tree.prefix)
}

/// Splice `changes` into the `container` child of a fragment's root,
/// creating the container when it is absent: after the root's last
/// WordprocessingML child named in `leading`, otherwise as its first child.
fn splice(
    xml: &[u8],
    root_local: &[u8],
    container: &'static str,
    order: &'static [&'static str],
    nested: &[&[u8]],
    leading: &[&str],
    changes: &[Change],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Vec<u8>, SpliceError> {
    let fragment = scan(xml, root_local, container.as_bytes(), nested, inherited)?;
    let needed = changes
        .iter()
        .any(|change| change.plan.as_ref().is_some_and(writes_attributes));
    let context = Context {
        xml,
        inherited,
        root: &fragment.root,
        attribute_prefix: attribute_prefix(&fragment, needed)?,
    };
    let mut edits = Vec::new();
    if let Some(node) = child(&fragment.tree, container)? {
        edit_container(&context, node, order, changes, &mut edits)?;
    } else {
        let mut bytes = Vec::new();
        let prefix = fragment.tree.prefix.as_slice();
        bytes.push(b'<');
        qualified(&mut bytes, prefix, container.as_bytes())?;
        bytes.push(b'>');
        for change in changes {
            if let Some(plan) = &change.plan {
                serialize(
                    &mut bytes,
                    prefix,
                    context.attribute_prefix,
                    change.local,
                    plan,
                )?;
            }
        }
        bytes.extend_from_slice(b"</");
        qualified(&mut bytes, prefix, container.as_bytes())?;
        bytes.push(b'>');
        let root = &fragment.tree;
        if root.empty {
            let tag = &xml[root.start..root.tag_end];
            let mut replacement = opened(tag)?;
            replacement.extend_from_slice(&bytes);
            replacement.extend_from_slice(&closing(tag_name(tag)?));
            edits.push(Edit {
                start: root.start,
                end: root.end,
                bytes: replacement,
            });
        } else {
            let at = root
                .children
                .iter()
                .filter(|node| leading.iter().any(|local| node.is(local)))
                .map(|node| node.end)
                .max()
                .unwrap_or(root.tag_end);
            edits.push(Edit {
                start: at,
                end: at,
                bytes,
            });
        }
    }
    apply(xml, edits)
}

// ---------------------------------------------------------------------------
// The crate-facing operations.
// ---------------------------------------------------------------------------

/// Whether two run property values state the same typed properties, ignoring
/// which properties each reports as invalid.
pub(crate) fn same_run(left: &Run, right: &Run) -> bool {
    let mut left = left.clone();
    left.invalid = right.invalid;
    left == *right
}

/// Whether two paragraph property values state the same typed properties,
/// ignoring which properties each reports as invalid.
pub(crate) fn same_paragraph(left: &Paragraph, right: &Paragraph) -> bool {
    let mut left = left.clone();
    left.invalid = right.invalid;
    match (left.mark.as_deref_mut(), right.mark.as_deref()) {
        (Some(left), Some(right)) => left.invalid = right.invalid,
        (None, None) => {},
        _ => return false,
    }
    left == *right
}

/// Rewrite one run fragment so that its direct properties read as `target`.
///
/// `current` must be what [`super::run::parse`] reads from `xml`. Returns
/// `None` when nothing changes. The result is read back and compared with
/// `target` before it is returned.
pub(crate) fn splice_run(
    xml: &[u8],
    current: &Run,
    target: &Run,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if current == target {
        return Ok(None);
    }
    admit_run(target)?;
    let changes = run_changes(current, target)?;
    if changes.is_empty() {
        return Ok(None);
    }
    let output = splice(xml, b"r", "rPr", RUN_ORDER, &[], &[], &changes, inherited)?;
    if !same_run(&super::run::parse(&output, inherited)?, target) {
        return Err(
            Error::InvalidFormat("run property splice failed typed readback".into()).into(),
        );
    }
    Ok(Some(output))
}

/// Rewrite one paragraph fragment so that its direct properties read as
/// `target`; see [`splice_run`]. `strict` refuses values that only the
/// Transitional vocabulary spells.
pub(crate) fn splice_paragraph(
    xml: &[u8],
    current: &Paragraph,
    target: &Paragraph,
    strict: bool,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if current == target {
        return Ok(None);
    }
    admit_paragraph(target)?;
    let changes = paragraph_changes(current, target, strict)?;
    if changes.is_empty() {
        return Ok(None);
    }
    let output = splice(
        xml,
        b"p",
        "pPr",
        PARAGRAPH_ORDER,
        &[b"numPr", b"pBdr", b"tabs", b"rPr"],
        &[],
        &changes,
        inherited,
    )?;
    if !same_paragraph(&super::paragraph::parse(&output, inherited)?, target) {
        return Err(
            Error::InvalidFormat("paragraph property splice failed typed readback".into()).into(),
        );
    }
    Ok(Some(output))
}

/// Whether two table property values state the same typed properties,
/// ignoring which properties each reports as invalid.
pub(crate) fn same_table(left: &Table, right: &Table) -> bool {
    let mut left = left.clone();
    left.invalid = right.invalid;
    left == *right
}

/// Whether two row property values state the same typed properties,
/// ignoring which properties each reports as invalid.
pub(crate) fn same_row(left: &Row, right: &Row) -> bool {
    let mut left = left.clone();
    left.invalid = right.invalid;
    left == *right
}

/// Rewrite one table fragment so that its direct properties read as
/// `target`; see [`splice_run`]. Only the style (`w:tblStyle`) and the
/// borders (`w:tblBorders`) are written (record 0342).
pub(crate) fn splice_table(
    xml: &[u8],
    current: &Table,
    target: &Table,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if same_table(current, target) {
        return Ok(None);
    }
    let changes = table_changes(current, target)?;
    if changes.is_empty() {
        return Ok(None);
    }
    let output = splice(
        xml,
        b"tbl",
        "tblPr",
        TABLE_ORDER,
        &[b"tblBorders"],
        &[],
        &changes,
        inherited,
    )?;
    if !same_table(&super::table::parse_table(&output, inherited)?, target) {
        return Err(
            Error::InvalidFormat("table property splice failed typed readback".into()).into(),
        );
    }
    Ok(Some(output))
}

/// Rewrite one row fragment so that its direct properties read as
/// `target`; see [`splice_run`]. The height and the `cantSplit`,
/// `tblHeader` and `hidden` flags are written (record 0342); a new
/// `w:trPr` follows the row's `w:tblPrEx`.
pub(crate) fn splice_row(
    xml: &[u8],
    current: &Row,
    target: &Row,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if same_row(current, target) {
        return Ok(None);
    }
    let changes = row_changes(current, target)?;
    if changes.is_empty() {
        return Ok(None);
    }
    let output = splice(
        xml,
        b"tr",
        "trPr",
        ROW_ORDER,
        &[],
        ROW_LEADING,
        &changes,
        inherited,
    )?;
    if !same_row(&super::table::parse_row(&output, inherited)?, target) {
        return Err(
            Error::InvalidFormat("row property splice failed typed readback".into()).into(),
        );
    }
    Ok(Some(output))
}

/// A complete `w:rPr` for a new run, written with `prefix`, or `None` when
/// `target` states no property.
pub(crate) fn run_properties_element(
    prefix: &[u8],
    target: &Run,
) -> Result<Option<Vec<u8>>, SpliceError> {
    admit_run(target)?;
    let changes = run_changes(&Run::default(), target)?;
    container_element(prefix, "rPr", &changes)
}

/// A complete `w:pPr` for a new paragraph, written with `prefix`, or `None`
/// when `target` states no property; see [`run_properties_element`].
pub(crate) fn paragraph_properties_element(
    prefix: &[u8],
    target: &Paragraph,
    strict: bool,
) -> Result<Option<Vec<u8>>, SpliceError> {
    admit_paragraph(target)?;
    let changes = paragraph_changes(&Paragraph::default(), target, strict)?;
    container_element(prefix, "pPr", &changes)
}

fn container_element(
    prefix: &[u8],
    local: &str,
    changes: &[Change],
) -> Result<Option<Vec<u8>>, SpliceError> {
    if changes.is_empty() {
        return Ok(None);
    }
    let mut output = Vec::new();
    output.push(b'<');
    qualified(&mut output, prefix, local.as_bytes())?;
    output.push(b'>');
    for change in changes {
        if let Some(plan) = &change.plan {
            serialize(&mut output, prefix, prefix, change.local, plan)?;
        }
    }
    output.extend_from_slice(b"</");
    qualified(&mut output, prefix, local.as_bytes())?;
    output.push(b'>');
    Ok(Some(output))
}

/// The schema rank of a `w:pPr` child, for tests.
#[cfg(test)]
pub(crate) fn paragraph_rank_for_tests(local: &str) -> Option<usize> {
    rank(PARAGRAPH_ORDER, local.as_bytes())
}

fn seed_namespaces(
    reader: &mut NsReader<&[u8]>,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<(), Error> {
    for (prefix, namespace) in inherited {
        let prefix = prefix.as_deref().map_or(
            quick_xml::name::PrefixDeclaration::Default,
            quick_xml::name::PrefixDeclaration::Named,
        );
        reader
            .resolver_mut()
            .add(prefix, quick_xml::name::Namespace(namespace))
            .map_err(|_| Error::Xml("invalid inherited property namespace".into()))?;
    }
    Ok(())
}

const MAX_STAGING_BYTES: usize = 64 * 1024 * 1024;
fn bounded_add(left: usize, right: usize) -> Result<usize, SpliceError> {
    let observed = left.checked_add(right).unwrap_or(usize::MAX);
    if observed > MAX_STAGING_BYTES {
        return Err(SpliceError::Transaction(crate::document::Error::Limit {
            kind: crate::document::LimitKind::FormattingStagingBytes,
            observed,
            maximum: MAX_STAGING_BYTES,
        }));
    }
    Ok(observed)
}
fn reserve_output(output: &mut Vec<u8>, bytes: usize) -> Result<(), SpliceError> {
    bounded_add(output.len(), bytes)?;
    output
        .try_reserve_exact(bytes)
        .map_err(|source| Error::Allocation {
            resource: "Word property output",
            source,
        })?;
    Ok(())
}
fn serialized_bound(
    prefix: &[u8],
    attribute_prefix: &[u8],
    local: &str,
    plan: &Plan,
) -> Result<usize, SpliceError> {
    let name = bounded_add(prefix.len(), local.len())?;
    let mut size = bounded_add(bounded_add(name, name)?, 8)?;
    for attribute in &plan.attributes {
        if let Want::Value(value) = &attribute.want {
            let escaped = value.len().checked_mul(6).unwrap_or(usize::MAX);
            size = bounded_add(
                size,
                bounded_add(
                    bounded_add(attribute_prefix.len(), attribute.local.len())?,
                    bounded_add(escaped, 5)?,
                )?,
            )?;
        }
    }
    match &plan.children {
        Children::Keep => {},
        Children::Container { changes, .. } => {
            for change in changes {
                if let Some(plan) = &change.plan {
                    size = bounded_add(
                        size,
                        serialized_bound(prefix, attribute_prefix, change.local, plan)?,
                    )?;
                }
            }
        },
        Children::Tabs(tabs) => {
            for plan in tabs {
                size = bounded_add(
                    size,
                    serialized_bound(prefix, attribute_prefix, "tab", plan)?,
                )?;
            }
        },
    }
    Ok(size)
}

pub(crate) fn admit_run(run: &Run) -> Result<usize, SpliceError> {
    let mut size = 4096;
    let mut text = |value: &Option<String>| -> Result<(), SpliceError> {
        if let Some(value) = value {
            size = bounded_add(size, value.len().checked_mul(6).unwrap_or(usize::MAX))?;
        }
        Ok(())
    };
    text(&run.style)?;
    if let Some(fonts) = &run.fonts {
        text(&fonts.ascii)?;
        text(&fonts.h_ansi)?;
        text(&fonts.east_asia)?;
        text(&fonts.complex_script)?;
    }
    if let Some(language) = &run.language {
        text(&language.latin)?;
        text(&language.east_asia)?;
        text(&language.bidi)?;
    }
    Ok(size)
}
pub(crate) fn admit_paragraph(paragraph: &Paragraph) -> Result<usize, SpliceError> {
    let mut size = 4096;
    if let Some(style) = &paragraph.style {
        size = bounded_add(size, style.len().checked_mul(6).unwrap_or(usize::MAX))?;
    }
    if let Some(tabs) = &paragraph.tabs {
        size = bounded_add(size, tabs.len().checked_mul(512).unwrap_or(usize::MAX))?;
    }
    if let Some(mark) = &paragraph.mark {
        size = bounded_add(size, admit_run(mark)?)?;
    }
    Ok(size)
}
